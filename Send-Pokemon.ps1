<#
.NOTES
    NAME: Send-Pokemon.ps1
    Type: PowerShell

        AUTHOR:  David Schulte
        DATE:    2022-06-20
        EMAIL:   celerium@celerium.org
        Updated:
        Date:

    VERSION HISTORY:
    0.1 - 2022-06-20 - Initial Release

    TODO:
    N\A

.SYNOPSIS
    Sends a Pokemon image & stats to a Teams channel.

.DESCRIPTION
    The Send-Pokemon script sends a Pokemon image & stats to a Teams channel using a Teams webhook connector URI.

    This script will only send to teams if the $DeployPokemon variable is set to true which is randomized.

    Pokemon images & facts are pulled from the pokeapi.glitch.me API

    Unless the -Verbose parameter is used, no output is displayed.

.PARAMETER TeamsURI
    A string that defines where the Microsoft Teams connector URI sends information to.

.EXAMPLE
    .\Send-Pokemon.ps1 -TeamsURI 'https://outlook.office.com/webhook/123/123/123/.....'

    Using the defined webhooks connector URI a random Pokemon image & stats are sent to the webhooks Teams channel.

    No output is displayed to the console.

.EXAMPLE
    .\Send-Pokemon.ps1 -TeamsURI 'https://outlook.office.com/webhook/123/123/123/.....' -Verbose

    Using the defined webhooks connector URI a random Pokemon image & stats are sent to the webhooks Teams channel.

    Output is displayed to the console.

.INPUTS
    TeamsURI

.OUTPUTS
    Console, TXT

.LINK
    Celerium - https://www.celerium.org/
    Pokemon Data - pokeapi.glitch.me

#>

<############################################################################################
                                        Code
############################################################################################>
#Requires -Version 5.0

#Region  [ Parameters ]

[CmdletBinding()]
param(
        [Parameter(ValueFromPipeline = $true, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$TeamsURI
    )

#EndRegion  [ Parameters ]

Write-Verbose ''
Write-Verbose "START - $(Get-Date -Format yyyy-MM-dd-HH:mm)"
Write-Verbose ''
Write-Verbose " - (1/3) - $(Get-Date -Format MM-dd-HH:mm) - Gathering Pokemon Data"

#Region     [ Prerequisites ]

    $Log = "C:\Celerium\Logs\Send-Pokemon-Report"
    $TXTReport = "$Log\Send-PokemonLog.txt"

    #Min & Max not compatible with -Count in PS5
    $CurrentDate = (Get-Date).DayOfWeek.value__
    $RandomDate = 1..5 | Get-Random -Count 2

    $DeployPokemon = foreach ($Number in $RandomDate){
        $Match = if ( $CurrentDate -eq $Number ){$true}else{$false}
        $Match
    }

        if ($DeployPokemon -contains $true){
            Write-Verbose " -       - $(Get-Date -Format MM-dd-HH:mm) - [ $CurrentDate |  $RandomDate ]"
        }
        else{
            Write-Verbose " -       - $(Get-Date -Format MM-dd-HH:mm) - [ $CurrentDate |  $RandomDate ]"
            Write-Verbose " -       - $(Get-Date -Format MM-dd-HH:mm) - Sorry, No wild Pokemon were found today."
            exit
        }

#EndRegion  [ Prerequisites ]

#Region  [ Main Code ]

try {

    $PokemonCounts = (Invoke-RestMethod 'https://pokeapi.glitch.me/v1/pokemon/counts' -ErrorAction Stop).total + 1
        $PokemonNumber = Get-Random -Minimum 1 -Maximum $PokemonCounts

    $PokemonData = Invoke-RestMethod -Uri "https://pokeapi.glitch.me/v1/pokemon/$PokemonNumber" -ErrorAction Stop
        if ($PokemonData.Count -gt 1){
            $PokemonName = ($PokemonData | Select-Object -First 1).name
            Write-Verbose " -       - $(Get-Date -Format MM-dd-HH:mm) - $PokemonName has [ $($PokemonData.Count) ] values"

            $PokemonData = $PokemonData | Get-Random -Count 1
        }

}
catch {
    Write-Error $_

    if ( (Test-Path -Path $Log -PathType Container) -eq $false ){
        New-Item -Path $Log -ItemType Directory > $null
    }

    (Get-Date -Format yyyy-MM-dd-HH:mm) + " - " + "[ Step (1/3) ]" + " - " + $_.Exception.Message | Out-File $TXTReport -Append -Encoding utf8

    exit
}

#EndRegion  [ Main Code ]

Write-Verbose " - (2/3) - $(Get-Date -Format MM-dd-HH:mm) - Formatting Pokemon Data"

#Region     [ Adjust for meta types ]

    switch ($PokemonData){
        {$_.starter -eq $true} {
            $TitleText = "A starter Pokemon has appeared!"
            $TitleColor = "light"
        }
        {$_.Legendary -eq $true} {
            $TitleText = "A Legendary Pokemon has appeared!"
            $TitleColor = "good"
        }
        {$_.Mythical -eq $true} {
            $TitleText = "A Mythical Pokemon has appeared!"
            $TitleColor = "warning"
        }
        {$_.UltraBeast -eq $true} {
            $TitleText = "A UltraBeast Pokemon has appeared!"
            $TitleColor = "warning"
        }
        {$_.Mega -eq $true} {
            $TitleText = "A Mega Pokemon has appeared!"
            $TitleColor = "attention"
        }
        default {
            $TitleText = "A wild Pokemon has appeared!"
            $TitleColor = "default"
        }
    }


#EndRegion  [ Adjust for meta types ]

Write-Verbose " - (3/3) - $(Get-Date -Format MM-dd-HH:mm) - Sending Pokemon Data"

#Region     [ Teams Code ]

$JSONBody = @"
{
    "type": "message",
    "attachments": [
        {
            "contentType": "application/vnd.microsoft.card.adaptive",
            "contentUrl": null,
            "content": {
                "$('$schema')": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.4",
                "body": [
                    {
                        "type": "TextBlock",
                        "size": "Large",
                        "weight": "Bolder",
                        "color": "$TitleColor",
                        "text": " $TitleText"
                    },
                    {
                        "type": "ColumnSet",
                        "style": "emphasis",
                        "columns": [
                            {
                                "type": "Column",
                                "style": "default",
                                "items": [
                                    {
                                        "type": "Image",
                                        "url": "$($PokemonData.sprite)",
                                        "altText": "$($PokemonData.name)",
                                        "msTeams": {
                                            "allowExpand": true
                                        }
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "Gen$($PokemonData.gen): $($PokemonData.name) - #$($PokemonData.number)",
                                        "wrap": true
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "items": [
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Type: $( $PokemonData.types -join ', ' )"
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Height: $( (($PokemonData.height) -replace ("'",'ft ')) -replace ('"','in'))"
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Weight: $($PokemonData.weight)"
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": ""
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Evolution State: $( $PokemonData.family.evolutionStage )"
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Evolution Line: $( $PokemonData.family.evolutionline  -join ', ' )",
                                                "wrap": true
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Description: $( $PokemonData.description )",
                                                "wrap": true
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "items": [
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Abilities:"
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Normal:  $( $PokemonData.abilities.normal  -join ', ' ) "
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Hidden:  $( $PokemonData.abilities.hidden  -join ', ' ) "
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": ""
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Meta:"
                                            }
                                        ]
                                    },
                                    {
                                        "type": "Container",
                                        "style": "default",
                                        "id": "meta",
                                        "isVisible": false,
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Starter: $( $PokemonData.starter )"
                                            },
                                            {
                                                "type": "TextBlock",
                                                "text": "Legendary: $( $PokemonData.legendary )"
                                            },
                                            {
                                                "type": "TextBlock",
                                                "text": "Mythical: $( $PokemonData.mythical )"
                                            },
                                            {
                                                "type": "TextBlock",
                                                "text": "UltraBeast: $( $PokemonData.ultraBeast )"
                                            },
                                            {
                                                "type": "TextBlock",
                                                "text": "Mega: $( $PokemonData.mega )"
                                            }
                                        ]
                                    }
                                ]
                            }
                        ]
                    }
                ],
                "actions": [
                    {
                        "type": "Action.OpenUrl",
                        "title": "Research more Pokemon",
                        "url": "https://regalion.surge.sh/"
                    },
                    {
                        "type": "Action.OpenUrl",
                        "title": "Source",
                        "url": "$($PokemonData.sprite)"
                    },
                    {
                        "type": "Action.ToggleVisibility",
                        "title": "Show Meta",
                        "targetElements": [
                            {
                                "elementId": "meta",
                                "isVisible": true
                            }
                        ]
                    },
                    {
                        "type": "Action.ToggleVisibility",
                        "title": "Hide Meta!",
                        "targetElements": [
                            {
                                "elementId": "meta",
                                "isVisible": false
                            }
                        ]
                    }
                ],
                "msTeams": {
                    "width": "Full"
                }
            }
        }
    ]
}
"@

try {

    Invoke-RestMethod -Uri $TeamsURI -Method Post -ContentType 'application/json' -Body $JsonBody -ErrorAction Stop > $null

}
catch {
    Write-Error $_

    if ( (Test-Path -Path $Log -PathType Container) -eq $false ){
        New-Item -Path $Log -ItemType Directory > $null
    }

    (Get-Date -Format yyyy-MM-dd-HH:mm) + " - " + "[ Step (3/3) ]" + " - " + $_.Exception.Message | Out-File $TXTReport -Append -Encoding utf8

    exit
}

#EndRegion  [ Teams Code ]

Write-Verbose ''
Write-Verbose "End - $(Get-Date -Format yyyy-MM-dd-HH:mm)"
Write-Verbose ''