<#
  Copyright 2026 Christoph Becker

  Licensed under the Apache License, Version 2.0 (the "License");
  you may not use this file except in compliance with the License.
  You may obtain a copy of the License at

  http://www.apache.org/licenses/LICENSE-2.0

  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS,
  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  See the License for the specific language governing permissions and
  limitations under the License.
#>


function ExitScript {
  param (
    [int]$ExitCode
  )

  Write-Host ""
  Write-Host "Press Enter to exit..."
  Read-Host
  exit $ExitCode
}


function DoPost {
  param (
    [Microsoft.PowerShell.Commands.WebRequestSession]$Session,
    [string]$Page,
    [hashtable]$Body
  )

  Invoke-WebRequest `
    -UseBasicParsing `
    -Uri "$BaseUrl/$Page" `
    -Method POST `
    -Body $Body `
    -WebSession $Session
}


function DoGet {
  param (
    [Microsoft.PowerShell.Commands.WebRequestSession]$Session,
    [string]$Page,
    [string]$OutFile
  )

  Invoke-WebRequest `
    -UseBasicParsing `
    -Uri "$BaseUrl/$Page" `
    -WebSession $Session `
    -OutFile $OutFile
}


Set-Variable BaseUrl -Option Constant -Value "https://langknecht.alco-web.de"

$Credential = Import-Clixml "$PSScriptRoot\credentials.xml"

$Session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

$LoginPage = DoPost `
  -Session $Session `
  -Page "index.php" `
  -Body @{
    "username"    = $Credential.UserName
    "password"    = $Credential.GetNetworkCredential().Password
    "remember-me" = "on"
  }

if ($LoginPage.StatusCode -ne 200 -or
  $LoginPage.Content -like '*Die Anmeldung ist fehlgeschlagen!*') {
  Write-Host "Login fehlgeschlagen." -ForegroundColor Red
  ExitScript 1
}

Write-Host "Login erfolgreich." -ForegroundColor Green

$abrechnungPage = DoPost `
  -Session $Session `
  -Page "obj-abrechnung.php" `
  -Body @{
    "abrechnungszeit" = "0"
  }

if ($abrechnungPage.StatusCode -ne 200) {
  Write-Host "Abrechnung konnte nicht abgerufen werden: $($abrechnungPage.StatusCode)" -ForegroundColor Red
  ExitScript 1
}

$abrechnungPage.Content -match '<option\s+value="[^"]+"\s+selected>([^<]+)</option>' | Out-Null
$abrechnungszeitraum = $matches[1].Trim()

Write-Host "Lade Unterlagen von Abrechnung $abrechnungszeitraum herunter..."

$abrechnungsFolder = Join-Path (Get-Location) $abrechnungszeitraum
New-Item -ItemType Directory -Path $abrechnungsFolder -Force | Out-Null

$abrechnungPage.Links |
  Where-Object { $_.href -like "*kontoauszug.php?ID*" } |
  ForEach-Object {

  $name = $_.outerHTML -replace '^.*>([^<]*)</a>.*$', '$1'
  $name = $name.Trim() -replace '/', ' '

  Write-Host "Verarbeite: $name"

  $abrechnungsDetailFolder = Join-Path $abrechnungsFolder $name
  $csvPath = Join-Path $abrechnungsDetailFolder "kontoauszug.csv"

  $kontoauszugPage = DoGet `
    -Session $Session `
    -Page $_.href

  if ($kontoauszugPage.StatusCode -ne 200) {
    Write-Host "  Fehler beim Abrufen des Kontoauszugs: $($kontoauszugPage.StatusCode)" -ForegroundColor Red
    ExitScript 1
  }

  $rows = [regex]::Matches(
    $kontoauszugPage.Content,
    '(?s)<tr>\s*(.*?)\s*</tr>'
  )

  $countBuchungen = 0
  $downloadedIds = @{}

  $buchungen = foreach ($row in $rows) {
    $rowHtml = $row.Groups[1].Value

    if ($rowHtml -notmatch '<td[^>]*>\s*(\d{2}\.\d{2}\.\d{4})\s*</td>') {
      continue
    }

    $datum = $matches[1]
    $cells = [regex]::Matches(
      $rowHtml,
      '(?s)<td[^>]*>(.*?)</td>'
    )

    $brutto       = ($cells[1].Groups[1].Value -replace '<[^>]+>', '').Trim()
    $saldo1       = ($cells[2].Groups[1].Value -replace '<[^>]+>', '').Trim()
    $buchungstext = ($cells[3].Groups[1].Value -replace '<[^>]+>', '').Trim()
    $netto        = ($cells[4].Groups[1].Value -replace '<[^>]+>', '').Trim()
    $saldo2       = ($cells[5].Groups[1].Value -replace '<[^>]+>', '').Trim()
    $countBuchungen = $countBuchungen + 1

    if ($countBuchungen -eq "1"){
      New-Item -ItemType Directory `
        -Path $abrechnungsDetailFolder `
        -Force | Out-Null
    }

    $pdfID = ""
    if ($cells[3].Groups[1].Value -match 'href="([^"]*showpdf\.php\?ID=(\d+)[^"]*)"') {
      $pdfLink = $matches[1]
      $pdfID = $matches[2]
    }

    [PSCustomObject]@{
      'Datum'        = $datum
      'Brutto'       = $brutto
      'Saldo 1'      = $saldo1
      'Buchungstext' = $buchungstext
      'PDF-ID'       = $pdfID
      'Netto'        = $netto
      'Saldo 2'      = $saldo2
    }

    if ([string]::IsNullOrEmpty($pdfID)) {
      continue
    }

    if ($downloadedIds.ContainsKey($pdfID)) {
      Write-Host "  Überspringe PDF-ID $pdfID (bereits heruntergeladen)"
      continue
    }

    $datumObj = [datetime]::ParseExact(
      $datum,
      'dd.MM.yyyy',
      $null
    )

    $beschreibung = $buchungstext -replace '[\\/:*?"<>|]', ' '
    $beschreibung = $beschreibung -replace '\s+', ' '
    $beschreibung = $beschreibung.Trim()

    $dateiname = "$($datumObj.ToString('yyyyMMdd')) - $beschreibung ($pdfId).pdf"
    $dateipfad = Join-Path $abrechnungsDetailFolder $dateiname

    Write-Host "  Download: $dateiname"

    DoGet `
      -Session $Session `
      -Page $pdfLink `
      -OutFile $dateipfad

    $downloadedIds[$pdfId] = $true
  }

  if ($countBuchungen -gt "0"){
    $buchungen | Export-Csv `
      -Path $csvPath `
      -Delimiter ';' `
      -NoTypeInformation `
      -Encoding UTF8
  } else {
    Write-Host "  Keine Buchungen gefunden!"
  }
}

ExitScript 0
