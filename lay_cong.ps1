param([switch]$KiemTra)

# Dien thong tin API mot lan tai day, giu dau nhay don.
$deviceUrl = 'http://DIA_CHI_MAY/ISAPI/AccessControl/AcsEvent?format=json'
$username = 'Dien ten dang nhap'
$password = 'Dien mat khau'

$ErrorActionPreference = 'Stop'
$utf8Bom = New-Object System.Text.UTF8Encoding $true
$utf8 = New-Object System.Text.UTF8Encoding $false
$culture = [Globalization.CultureInfo]::InvariantCulture
$localOffset = [TimeSpan]::FromHours(7)
$configPath = Join-Path $PSScriptRoot 'dieu_kien.json'

function Read-EventTime($value) {
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) {
        throw 'Thieu startTime hoac endTime trong dieu_kien.json.'
    }
    if ($value -is [datetime] -or $value -is [datetimeoffset]) {
        return ([datetimeoffset]$value).ToOffset($localOffset)
    }
    return ([datetimeoffset]::Parse([string]$value, $culture)).ToOffset($localOffset)
}

$config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
$condition = $config.AcsEventCond
if ($null -eq $condition) {
    throw 'dieu_kien.json can co doi tuong AcsEventCond.'
}
$start = Read-EventTime $condition.startTime
$end = Read-EventTime $condition.endTime
if ($start -gt $end) { throw 'startTime phai truoc hoac bang endTime.' }
$pageSize = 1000
if ($null -ne $condition.maxResults) { $pageSize = [int]$condition.maxResults }
if ($pageSize -lt 1 -or $pageSize -gt 1000) {
    throw 'maxResults can tu 1 den 1000.'
}
$startText = $start.ToString('yyyy-MM-ddTHH:mm:sszzz', $culture)
$endText = $end.ToString('yyyy-MM-ddTHH:mm:sszzz', $culture)
$filename = 'RAW_MayChamCong_{0}_{1}.csv' -f $start.ToString('yyyy-MM-dd'), $end.ToString('yyyy-MM-dd')
$outputPath = Join-Path $PSScriptRoot $filename
Write-Host "Ky lay RAW: $startText den $endText"
Write-Host "File ket qua: $outputPath"
if ($KiemTra) {
    Write-Host 'Cau hinh hop le. Moi lan tai se bat dau tu vi tri 0.'
    return
}

if ($deviceUrl -match 'DIA_CHI_MAY' -or $username -eq 'Dien ten dang nhap' -or $password -eq 'Dien mat khau' -or
    [string]::IsNullOrWhiteSpace($username) -or [string]::IsNullOrWhiteSpace($password)) {
    throw 'Hay dien deviceUrl, username va password o dau file lay_cong.ps1.'
}
$endpoint = $null
if (-not [uri]::TryCreate($deviceUrl, [UriKind]::Absolute, [ref]$endpoint) -or $endpoint.Scheme -notin @('http', 'https') -or $endpoint.UserInfo) {
    throw 'URL may khong hop le. Khong dat tai khoan/mat khau trong URL.'
}
$curlLogin = '{0}:{1}' -f $username, $password

$runId = [guid]::NewGuid().ToString('N')
$requestPath = Join-Path ([IO.Path]::GetTempPath()) "chamcong_request_$runId.json"
$responsePath = Join-Path ([IO.Path]::GetTempPath()) "chamcong_response_$runId.json"
$csvPath = Join-Path $PSScriptRoot ".chamcong_$runId.csv.tmp"
$events = New-Object 'System.Collections.Generic.List[object]'
$position = 0
$expectedTotal = $null

try {
    while ($true) {
        $pageCondition = @{}
        foreach ($property in $condition.PSObject.Properties) {
            $pageCondition[$property.Name] = $property.Value
        }
        $pageCondition.startTime = $startText
        $pageCondition.endTime = $endText
        $pageCondition.searchID = $runId
        $pageCondition.searchResultPosition = $position
        $pageCondition.maxResults = $pageSize
        if (-not $pageCondition.ContainsKey('major')) { $pageCondition.major = 0 }
        if (-not $pageCondition.ContainsKey('minor')) { $pageCondition.minor = 0 }
        $request = @{ AcsEventCond = $pageCondition } | ConvertTo-Json -Depth 15
        [IO.File]::WriteAllText($requestPath, $request, $utf8)

        & curl.exe --silent --show-error --fail --connect-timeout 15 --max-time 120 --digest --user $curlLogin -H "Content-Type: application/json" --data-binary "@$requestPath" --output $responsePath $deviceUrl
        if ($LASTEXITCODE -ne 0) {
            throw 'Khong tai duoc du lieu tu may cham cong. CSV chua duoc cap nhat.'
        }
        $result = Get-Content -LiteralPath $responsePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $page = $result.AcsEvent
        if ($null -eq $page -or $page.responseStatusStrg -notin @('OK', 'MORE', 'NO MATCH')) {
            throw 'May cham cong tra ve trang thai khong hop le.'
        }
        if ($null -eq $page.totalMatches -or $null -eq $page.numOfMatches) {
            throw 'Phan hoi thieu tong so su kien.'
        }
        if ($null -eq $expectedTotal) { $expectedTotal = [int]$page.totalMatches }
        if ([int]$page.totalMatches -ne $expectedTotal) {
            throw 'Tong so su kien thay doi trong luc tai. Hay chay lai de lay du lieu nhat quan.'
        }
        $items = @()
        if ($null -ne $page.InfoList) { $items = @($page.InfoList) }
        if ($items.Count -ne [int]$page.numOfMatches) {
            throw 'So dong cua trang khong khop numOfMatches.'
        }
        foreach ($event in $items) {
            $stamp = Read-EventTime $event.time
            if ($stamp -lt $start -or $stamp -gt $end) {
                throw 'May tra ve su kien nam ngoai khoang ngay da chon.'
            }
            $event.time = $stamp.ToString('yyyy-MM-ddTHH:mm:sszzz', $culture)
            $events.Add($event)
        }
        $position += $items.Count
        Write-Host "Da tai $position / $expectedTotal su kien."
        if ($position -gt $expectedTotal) { throw 'So dong vuot qua tong may tra ve.' }
        if ($page.responseStatusStrg -in @('OK', 'NO MATCH')) { break }
        if ($items.Count -eq 0) { throw 'Trang rong khi may van bao con du lieu.' }
    }
    if ($events.Count -ne $expectedTotal) {
        throw 'Du lieu chua du so luong, CSV chua duoc cap nhat.'
    }
    $lines = @($events | Select-Object time, employeeNoString, name, cardReaderNo, doorNo, currentVerifyMode | ConvertTo-Csv -NoTypeInformation)
    if ($events.Count -eq 0) {
        $lines = @('"time","employeeNoString","name","cardReaderNo","doorNo","currentVerifyMode"')
    }
    [IO.File]::WriteAllLines($csvPath, [string[]]$lines, $utf8Bom)
    if ([IO.File]::Exists($outputPath)) {
        [IO.File]::Replace($csvPath, $outputPath, [System.Management.Automation.Language.NullString]::Value)
    } else {
        [IO.File]::Move($csvPath, $outputPath)
    }
    Write-Host "Hoan thanh: $filename. So su kien: $($events.Count)."
} finally {
    foreach ($temporary in @($requestPath, $responsePath, $csvPath)) {
        if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
    }
}
