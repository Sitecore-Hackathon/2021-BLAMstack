<#
.SYNOPSIS
    Driven by the Microsoft Azure Cognitive Services Text Analyzer API, this Sitecore PowerShell Extensions utility 
    can be used by Content Authors and Marketers to analyze field content in Sitecore to extract keywords. 

.DESCRIPTION
    Azure Cognitive Services Text Analyzer is a cloud-based API service.
    https://docs.microsoft.com/en-us/azure/cognitive-services/text-analytics/overview

    The script allows users to right-click on a Sitecore item, select 'Keyword Analyzer' from the list of scripts, and choose a field to analyze.  
    The analysis returns list of extracted keywords which can then be used to manually populate a meta keywords field, for example. 
    These results are displayed in a Show-Result modal.

.NOTES
    This script was developed to work as a `Context Menu` and a `Ribbon Button` PowerShell Script item.

.AUTHOR
    Gabe Streza
    Team `BLAMstack` - Sitecore Hackathon 2021 
#>

function Invoke-ItemAnalysis {
    param (
        [Parameter(Mandatory = $true)]
        [Item]$TargetItem,

        [Parameter(Mandatory = $true)]
        [string]$FieldName,

        [Parameter(Mandatory = $true)]
        [string]$Language,

        [Parameter(Mandatory = $true)]
        [string]$SubscriptionKey
    )

    $targetItem = Get-Item $TargetItem.ID -Language $Language

    $fieldValue = $targetItem.Fields[$FieldName].Value

    if ($fieldValue -eq "") {
        Write-Host "The '$FieldName' field on the '$Language' language version does not have any content to analyze." -ForegroundColor Yellow
        Show-Result -Text -Width 450 -Height 600 
        Exit
    }

    # Sanitze string - strip HTML
    $sanitizedFieldValue = $fieldValue -replace '<[^>]+>', '';
    $sanitizedFieldValue = $sanitizedFieldValue -replace '\"', '';
    $sanitizedFieldValue = $sanitizedFieldValue -replace '\?', '';
    $sanitizedFieldValue = $sanitizedFieldValue -replace '&nbsp;', '';
    $sanitizedFieldValue = $sanitizedFieldValue.replace("`n", " ").replace("`r", " ")

    # Call the API service
    Invoke-KeywordAnalysis -FieldValue $sanitizedFieldValue -Language $Language -SubscriptionKey $SubscriptionKey
}

function Invoke-KeywordAnalysis {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FieldValue,

        [Parameter(Mandatory = $true)]
        [string]$Language,

        [Parameter(Mandatory = $true)]
        [string]$SubscriptionKey

    )

    try {
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add("Ocp-Apim-Subscription-Key", $SubscriptionKey)
        $headers.Add("Content-Type", "application/json")
            
        $body = "{
            `n    `"documents`": [
            `n        {
            `n            `"language`": `"$Language`",
            `n            `"id`": `"1`",
            `n            `"text`": `"$FieldValue`"
            `n        }
            `n    ]
            `n}"
            
        $response = Invoke-RestMethod $KeywordAnalyticsEndpoint -Method 'POST' -Headers $headers -Body $body
        $keyPhrases = $response.documents.keyPhrases

        Clear-Host
        $Host.UI.RawUI.BackgroundColor = 'Black'
        $Host.UI.RawUI.ForegroundColor = 'Cyan'

        Write-Host "Key Phrases: " -ForegroundColor DarkGray
        foreach ($phrase in $keyPhrases.Split(' ')) {
            Write-Host $phrase
        }
       
        Show-Result -Text -Width 450 -Height 600 
    } 
    catch [System.Net.WebException] {
        # An error occured calling the API
        Write-Host 'Error calling API' -ForegroundColor Red
        Write-Host $Error[0] -ForegroundColor Red
        return $null
    } 
}

# Host used for API endpoint
$KeywordAnalyticsEndpointHost = "https://centralus.api.cognitive.microsoft.com"

# Input a key from the your own Cognitive Service API in Azure Portal (Resource Management -> Keys)
$SubscriptionKey = "00000000000000000000000000000000"

# Path used for API endpoint
$KeywordAnalyticsEndpointPath = "/text/analytics/v2.1/keyPhrases"

# Settings item is located at `/sitecore/system/Modules/PowerShell/Script Library/CognitiveTextAnalysis/API Settings`
$settingsItem = Get-Item "{E5006EF7-DD97-442A-B2F4-D8D6CF2D2FC4}"
if ($null -eq $settingsItem) {
    Show-Alert "API Settings item is missing.  Please reinstall the module."
    Exit
}

if ($settingsItem.Fields["Endpoint"].Value -ne "") {
    $KeywordAnalyticsEndpointHost = $settingsItem.Fields["Endpoint"].Value
}
else {
    Show-Alert "Custom endpoint host must be present on the 'API Settings' item.  `n`nPlease check the value on '/sitecore/system/Modules/PowerShell/Script Library/CognitiveTextAnalysis/API Settings'. `n`n ID: '{E5006EF7-DD97-442A-B2F4-D8D6CF2D2FC4}'"
    Exit 
}

if ($settingsItem.Fields["API Key"].Value -ne "") {
    if ($settingsItem.Fields["API Key"].Value.Length -ne "32") {
        Show-Alert "API key must be 32 characters in length.  `n`nPlease check the value on '/sitecore/system/Modules/PowerShell/Script Library/CognitiveTextAnalysis/API Settings'. `n`n ID: '{E5006EF7-DD97-442A-B2F4-D8D6CF2D2FC4}'"
        Exit
    }
    $SubscriptionKey = $settingsItem.Fields["API Key"].Value
}
else {
    Show-Alert "API key must be present on the 'API Settings'  `n`nPlease check the value on '/sitecore/system/Modules/PowerShell/Script Library/CognitiveTextAnalysis/API Settings'. `n`n ID: '{E5006EF7-DD97-442A-B2F4-D8D6CF2D2FC4}'"
    Exit 
}

$KeywordAnalyticsEndpoint = "$($KeywordAnalyticsEndpointHost)$($KeywordAnalyticsEndpointPath)"

# Get the current context item
$item = Get-Item "."

# Obtain context item's language versions
$siteLangOptions = New-Object System.Collections.Specialized.OrderedDictionary
foreach ($lang in $item.Languages) {
    $tempitem = Get-Item -Path $item.Paths.Path -Language $lang
    if ($tempitem.Versions.Count -gt 0) {
        $siteLangOptions.Add($lang.Name, $lang.Name)
    }
}

# Obtain context item's Single-Line, Multi-Line Text, and Rich Text fields (ignore System fields)
$fieldOptions = New-Object System.Collections.Specialized.OrderedDictionary
$item.Fields | Where-Object { ($_.TypeKey -eq "single-line text" -or "multi-line text" -or "rich text") -and ($_.Name -notlike "__*") } | ForEach-Object {
    if (-not [string]::IsNullOrEmpty($item.Fields[$_.Name].Value)) {
        $fieldOptions.Add($_.Name, $_.Name)
    }
}

# Window with options to select language and field to analyze
$dialogProps = @{
    Parameters       = @(
        @{ Name = "fieldSelection"; Title = "Field to analyze"; options = $fieldOptions; editor = "radio" },
        @{ Name = "languageSelection"; Title = "Language"; options = $siteLangOptions; editor = "radio" }
    )
    Description      = "Select a field to invoke a `Keyword Analysis` of the field's text using Microsoft Cognitive Services Text Analytics." 
    Title            = "Text Keyword Analyzer" 
    OkButtonName     = "Continue" 
    CancelButtonName = "Cancel"
    Width            = 400 
    Height           = 275 
    Icon             = "apps/32x32/Analysis (1).png"
}

# Wait for user input from options menu
$dialogResult = Read-Variable @dialogProps
if ($dialogResult -ne "ok") {
    # Exit if cancelled
    Exit
}

if (($null -eq $fieldSelection) -or ($null -eq $languageSelection)) {
    Show-Alert "You must select at a field and a language to begin the analysis."
    Exit
}

# Call Analyze-Field and pass in params
Invoke-ItemAnalysis -TargetItem $item -FieldName $fieldSelection -Language $languageSelection -SubscriptionKey $SubscriptionKey