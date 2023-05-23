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

function IsCpuUsageHigh([string] $ProcName, [int32] $Threthold) {
    return ((Get-Counter ('\Process(' + $ProcName + '*)\% Processor Time')).CounterSamples | `
        Where-Object { $_.CookedValue -gt $Threthold } | `
        Measure-Object).Count -gt 0
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
        ($Key + '=' + $Results[$Key]) > $ResultFilePath
    }
}

function SendMail([hashtable] $Ini, [string] $subject, [string] $body) {
    Send-MailMessage `
        -From $Ini['MailSettings']['from'] `
        -To $Ini['MailSettings']['to'] `
        -SmtpServer $Ini['MailSettings']['smtpServer'] `
        -Port $Ini['MailSettings']['port'] `
        -UseSsl `
        -Credential (New-Object System.Management.Automation.PSCredential('[email]', (ConvertTo-SecureString '[token]' -AsPlainText -Force))) `
        -Subject $subject `
        -Body $body
}

function Main() {
    [hashtable] $Results = ReadResultData ($Env:USERPROFILE + '\develop\monitoring\results.properties')
    [hashtable] $Ini = LoadIniFile ($Env:USERPROFILE + '\develop\monitoring\settings.ini')
    foreach ($Section in $Ini.Keys) {
        if ($Section -eq '__NoSection__' -or $Section -eq 'MailSettings') {
            continue
        }
        [boolean] $IsHigh = IsCpuUsageHigh $Ini[$Section]['process'] $Ini[$Section]['threshold']
        if ($IsHigh -and [System.Convert]::ToBoolean($Results[$Section])) {
            SendMail $Ini $Ini[$Section]['subject'] $Ini[$Section]['body']
        }
        $Results[$Section] = $IsHigh
    }
    WriteResultData ($Env:USERPROFILE + '\develop\monitoring\results.properties') $Results
}

Main
