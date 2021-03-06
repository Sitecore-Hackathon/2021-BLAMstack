<#
.SYNOPSIS
Driven by the Microsoft Azure Cognitive Services Text Analyzer API, this Sitecore PowerShell Extensions utility 
can be used by Content Authors and Marketers to analyze field content in Sitecore to review the content's sentiment rating. 

.DESCRIPTION
    Azure Cognitive Services Text Analyzer is a cloud-based API service. 
    https://docs.microsoft.com/en-us/azure/cognitive-services/text-analytics/overview

    The script allows users to right-click on a Sitecore item, select 'Sentiment Analyzer'
    from the list of scripts, and choose a field to analyze.  The analysis returns a overall sentiment ruling, and a sentence-by-sentence breakdown of each sentence's sentiment ruling and confidence scores.  These results are displayed in a Show-Result modal and formatted for ease of content mitigation.

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
    
    # Call the service
    Invoke-SentimentAnalysis -FieldValue $sanitizedFieldValue -Language $Language -SubscriptionKey $SubscriptionKey
}

function Invoke-SentimentAnalysis {
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
            
        $response = Invoke-RestMethod $SentimentAnalyticsEndpoint -Method 'POST' -Headers $headers -Body $body
        $documents = $response.documents

        Clear-Host
        $Host.UI.RawUI.BackgroundColor = 'Black'
        $Host.UI.RawUI.ForegroundColor = 'Cyan'

        Write-Host "Overall Field Content Sentiment: " -ForegroundColor DarkGray -NoNewline
        $overallSentimentScore = $documents.sentiment.ToUpper()
        switch ($overallSentimentScore) {
            "POSITIVE" { 
                Write-Host "😄👍 $($overallSentimentScore)`n`n" -ForegroundColor Green
            }
            "NEUTRAL" {
                Write-Host "✋😐🤚 $($overallSentimentScore)`n`n" -ForegroundColor White
            }
            "MIXED" {
                Write-Host "🎭 $($overallSentimentScore)`n`n" -ForegroundColor Yellow
            }
            "NEGATIVE" {
                Write-Host "👎😒 $($overallSentimentScore)`n`n" -ForegroundColor Red
            }
            Default { break }
        }  

        $response.documents.sentences | ForEach-Object { 
            $sentimentScore = $_.sentiment.ToUpper();
            $sentenceText = ($_.text -replace '&nbsp;', '').Trim();
            $confidenceScoresArray = $_.confidenceScores;
            Write-Host "`n---------"
            Write-Host "`n[ Sentence Analyzed ]" -ForegroundColor DarkGray
            Write-Host $sentenceText -ForegroundColor White

            switch ($sentimentScore) {
                "POSITIVE" { 
                    Write-Host "`n[ Sentiment ]" -ForegroundColor DarkGray
                    Write-Host "😄👍 $sentimentScore" -ForegroundColor Green
                    foreach ($confidenceScore in $confidenceScoresArray) {

                        Write-Host "`n[ Confidence Scores ]" -ForegroundColor DarkGray

                        Write-Host " - Positive: " -ForegroundColor White -NoNewline
                        if ([double]$confidenceScore.positive -eq 0.0) {
                            Write-Host $confidenceScore.positive -ForegroundColor DarkGray
                        }
                        else {
                            Write-Host $confidenceScore.positive -ForegroundColor Green
                        }

                        Write-Host " - Neutral: " -ForegroundColor White -NoNewline
                        if ([double]$confidenceScore.neutral -eq 0.0) {
                            Write-Host $confidenceScore.neutral -ForegroundColor DarkGray
                        }
                        else {
                            Write-Host $confidenceScore.neutral
                        }

                        Write-Host " - Negative: " -ForegroundColor White -NoNewline
                        if ([double]$confidenceScore.negative -eq 0.0) {
                            Write-Host $confidenceScore.negative -ForegroundColor DarkGray
                        }
                        else {
                            Write-Host $confidenceScore.negative -ForegroundColor Red
                        }
                    }

                }
                "NEUTRAL" {
                    Write-Host "`n[ Sentiment ]" -ForegroundColor DarkGray
                    Write-Host "✋😐🤚 $sentimentScore" -ForegroundColor White
                    foreach ($confidenceScore in $confidenceScoresArray) {

                        Write-Host "`n[ Confidence Scores ]" -ForegroundColor DarkGray

                        Write-Host " - Positive: " -ForegroundColor White -NoNewline
                        if ([double]$confidenceScore.positive -eq 0.0) {
                            Write-Host $confidenceScore.positive -ForegroundColor DarkGray
                        }
                        else {
                            Write-Host $confidenceScore.positive -ForegroundColor Green
                        }

                        Write-Host " - Neutral: " -ForegroundColor White -NoNewline
                        if ([double]$confidenceScore.neutral -eq 0.0) {
                            Write-Host $confidenceScore.neutral -ForegroundColor DarkGray
                        }
                        else {
                            Write-Host $confidenceScore.neutral
                        }

                        Write-Host " - Negative: " -ForegroundColor White -NoNewline
                        if ([double]$confidenceScore.negative -eq 0.0) {
                            Write-Host $confidenceScore.negative -ForegroundColor DarkGray
                        }
                        else {
                            Write-Host $confidenceScore.negative -ForegroundColor Red
                        }
                    }
                }
                "MIXED" {
                    Write-Host "`n[ Sentiment ]" -ForegroundColor DarkGray
                    Write-Host "🎭 $sentimentScore" -ForegroundColor Yellow
                    foreach ($confidenceScore in $confidenceScoresArray) {

                        Write-Host "`n[ Confidence Scores ]" -ForegroundColor DarkGray    

                        Write-Host " - Positive: " -ForegroundColor White -NoNewline
                        if ([double]$confidenceScore.positive -eq 0.0) {
                            Write-Host $confidenceScore.positive -ForegroundColor DarkGray
                        }
                        else {
                            Write-Host $confidenceScore.positive -ForegroundColor Green
                        }

                        Write-Host " - Neutral: " -ForegroundColor White -NoNewline
                        if ([double]$confidenceScore.neutral -eq 0.0) {
                            Write-Host $confidenceScore.neutral -ForegroundColor DarkGray
                        }
                        else {
                            Write-Host $confidenceScore.neutral
                        }

                        Write-Host " - Negative: " -ForegroundColor White -NoNewline
                        if ([double]$confidenceScore.negative -eq 0.0) {
                            Write-Host $confidenceScore.negative -ForegroundColor DarkGray
                        }
                        else {
                            Write-Host $confidenceScore.negative -ForegroundColor Red
                        }
                    }
                }
                "NEGATIVE" {
                    Write-Host "`n[ Sentiment ]" -ForegroundColor DarkGray
                    Write-Host "👎😒 $sentimentScore" -ForegroundColor Red

                    Write-Host "`n[ Confidence Scores ]" -ForegroundColor DarkGray

                    foreach ($confidenceScore in $confidenceScoresArray) {

                        Write-Host " - Positive: " -ForegroundColor White -NoNewline
                        if ([double]$confidenceScore.positive -eq 0.0) {
                            Write-Host $confidenceScore.positive -ForegroundColor DarkGray
                        }
                        else {
                            Write-Host $confidenceScore.positive -ForegroundColor Green
                        }

                        Write-Host " - Neutral: " -ForegroundColor White -NoNewline
                        if ([double]$confidenceScore.neutral -eq 0.0) {
                            Write-Host $confidenceScore.neutral -ForegroundColor DarkGray
                        }
                        else {
                            Write-Host $confidenceScore.neutral
                        }

                        Write-Host " - Negative: " -ForegroundColor White -NoNewline
                        if ([double]$confidenceScore.negative -eq 0.0) {
                            Write-Host $confidenceScore.negative -ForegroundColor DarkGray
                        }
                        else {
                            Write-Host $confidenceScore.negative -ForegroundColor Red
                        }
                    }
                }
                Default {}
            }            
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

# Input a key from the your own Cognitive Service API in Azure Portal (Resource Management -> Keys) using the `API Settings` item
$SubscriptionKey = "00000000000000000000000000000000"

# Host used for API endpoint
$SentimentAnalyticsEndpointHost = "https://centralus.api.cognitive.microsoft.com"

# Path used for API endpoint
$SentimentAnalyticsEndpointPath = "/text/analytics/v3.1-preview.1/sentiment"

# Settings item is located at `/sitecore/system/Modules/PowerShell/Script Library/CognitiveTextAnalysis/API Settings`
$settingsItem = Get-Item "{E5006EF7-DD97-442A-B2F4-D8D6CF2D2FC4}"
if ($null -eq $settingsItem) {
    Show-Alert "API Settings item is missing.  Please reinstall the module."
    Exit
}

if ($settingsItem.Fields["Endpoint"].Value -ne "") {
    $SentimentAnalyticsEndpointHost = $settingsItem.Fields["Endpoint"].Value
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

$SentimentAnalyticsEndpoint = "$($SentimentAnalyticsEndpointHost)$($SentimentAnalyticsEndpointPath)"

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
    Description      = "Select a field to invoke a `Sentiment Analysis` of the field's text using Microsoft Cognitive Services Text Analytics." 
    Title            = "Text Sentiment Analyzer" 
    OkButtonName     = "Continue" 
    CancelButtonName = "Cancel"
    Width            = 400 
    Height           = 275 
    Icon             = "office/32x32/speech_balloon.png"
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