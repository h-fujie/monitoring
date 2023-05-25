Param (
    [string] $IniFile = $Env:USERPROFILE + '\develop\monitoring\settings.ini',
    [string] $ResultFile = $Env:USERPROFILE + '\develop\monitoring\results.properties'
)

function LoadIniFile([string] $IniFilePath) {
    # 改行込みのパラメータ読み込みは非対応
    if ($null -eq $IniFilePath -or $IniFilePath -eq '') {
        throw 'iniファイルのパスが指定されていません。'
    }
    if (-not (Test-Path $IniFilePath)) {
        throw ('iniファイルが存在しません。 IniFilePath: ' + $IniFilePath)
    }
    $Ini = [ordered] @{}
    $Ini['__NoSection__'] = @{}
    switch -Regex -File $IniFilePath {
        '^\s*\[([^\]]+)\]\s*$' {
            $Section = $Matches[1]
            if ($Ini.Contains($Section)) {
                throw ('セクションが重複しています。 Section: ' + $Section)
            }
            $Ini[$Section] = @{}
            continue
        }
        '^\s*;.*$' {
            # NOP
            continue
        }
        '^\s*$' {
            # NOP
            continue
        }
        '^(.+)=(.*)$' {
            $Key = $Matches[1].Trim()
            $Value = $Matches[2].Trim()
            if ($null -eq $Section) {
                $Section = '__NoSection__'
            }
            if ($Ini[$Section].Contains($Key)) {
                throw ('キーが重複しています。 Section: ' + $Section + ', Key: ' + $Key)
            }
            $Ini[$Section][$Key] = $Value
            continue
        }
        default {
            throw ('解析不能です。')
        }
    }
    return $Ini
}

function IsCpuProcessUsageHigh([string] $ProcName, [int32] $Threshold) {
    $Processes = ((Get-Counter -Counter '\Process(_total)\% Processor Time', ('\Process(' + $ProcName + '*)\% Processor Time')).CounterSamples | `
        Select-Object @{ label = "Process"; expression = { $_.Path -replace ('^.*\((_total|' + $ProcName + ')(#\d+)?\).*$'), '$1' }}, CookedValue)
    $Total = ($Processes | Where-Object { $_.Process -eq '_total' })[0].CookedValue
    $Usage = ($Processes | Where-Object { $_.Process -eq $ProcName } | Measure-Object -Property CookedValue -Sum).Sum
    return ($Usage * 100) / $Total -gt $Threshold
}

function IsCpuUsageHigh([int32] $Threshold) {
    return ((Get-Counter -Counter '\Processor(_total)\% Processor Time').CounterSamples | `
        Select-Object CookedValue)[0].CookedValue -gt $Threshold
}

function ReadResultData([string] $ResultFilePath) {
    $Results = @{}
    if (Test-Path $ResultFilePath) {
        switch -Regex -File $ResultFilePath {
            '^(.+)=(.*)$' {
                $Key = $Matches[1].Trim()
                $Value = $Matches[2].Trim()
                if ($Results.Contains($Key)) {
                    throw ('キーが重複しています。 Key: ' + $Key)
                }
                $Results[$Key] = $Value
                continue
            }
            default {
                throw ('解析不能です。')
            }
        }
    }
    return $Results
}

function WriteResultData([string] $ResultFilePath, [hashtable] $Results) {
    if (Test-Path $ResultFilePath) {
        Clear-Content $ResultFilePath
    }
    foreach ($Key in $Results.Keys) {
        ($Key + '=' + $Results[$Key]) >> $ResultFilePath
    }
}

function SendMail([hashtable] $MailSettings, [string] $Subject, [string] $Body) {
    Send-MailMessage `
        -From $MailSettings['from'] `
        -To $MailSettings['to'] `
        -SmtpServer $MailSettings['smtpServer'] `
        -Port $MailSettings['port'] `
        -UseSsl `
        -Credential (New-Object System.Management.Automation.PSCredential('[email]', (ConvertTo-SecureString '[token]' -AsPlainText -Force))) `
        -Subject $Subject `
        -Body $Body
}

function Main() {
    [hashtable] $Results = ReadResultData $ResultFile
    [hashtable] $Ini = LoadIniFile $IniFile
    foreach ($Section in $Ini.Keys) {
        if ($Section -eq '__NoSection__' -or $Section -eq 'MailSettings') {
            continue
        }
        [boolean] $IsHigh = $false
        if ($Ini[$Section].Contains('process')) {
            $IsHigh = IsCpuProcessUsageHigh $Ini[$Section]['process'] $Ini[$Section]['threshold']
        } else {
            $IsHigh = IsCpuUsageHigh $Ini[$Section]['threshold']
        }
        if ($IsHigh -and $Results.Contains($Section) -and [System.Convert]::ToBoolean($Results[$Section])) {
            SendMail $Ini['MailSettings'] $Ini[$Section]['subject'] $Ini[$Section]['body']
        }
        $Results[$Section] = $IsHigh
    }
    WriteResultData $ResultFile $Results
}

Main
