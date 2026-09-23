param(
    [ValidateSet('Install', 'Setup', 'Desktop', 'Daily', 'Diagnose')]
    [string]$Mode = 'Install'
)

$ErrorActionPreference = 'Stop'

function Resolve-CertificatePython {
    param([bool]$RequirePackages = $true, [bool]$RequireOutlook = $true)
    $candidates = [System.Collections.Generic.List[object]]::new()
    $script:CertificatePythonProbeFailures = @()
    $localPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
    if (Test-Path -LiteralPath $localPython -PathType Leaf) {
        [void]$candidates.Add([pscustomobject]@{ Executable = $localPython; Prefix = @() })
    }
    # Python Install Manager puts its py.exe launcher in WindowsApps. Do not
    # run that launcher as though it were python.exe; ask it for installed
    # runtimes and probe the real python.exe paths it reports instead.
    $managerSources = @{}
    foreach ($name in @('py.exe', 'pymanager.exe')) {
        foreach ($command in @(Get-Command $name -CommandType Application -All -ErrorAction SilentlyContinue)) {
            $managerPath = [string]$command.Source
            if (-not $managerPath -or $managerSources.ContainsKey($managerPath)) { continue }
            $managerSources[$managerPath] = $true
            $isWindowsApps = $managerPath -match '\\Microsoft\\WindowsApps\\'
            $listedRuntimes = $false
            try {
                $runtimeLines = @(& $managerPath list --format=exe 2>$null)
                $listExitCode = $LASTEXITCODE
            } catch {
                $runtimeLines = @()
                $listExitCode = 1
            }
            if ($listExitCode -eq 0) {
                foreach ($runtimeLine in $runtimeLines) {
                    $runtimePath = ([string]$runtimeLine).Trim().Trim('"')
                    if ($runtimePath -match '(?i)^[A-Z]:\\.*python(?:w)?\.exe$' -and (Test-Path -LiteralPath $runtimePath -PathType Leaf)) {
                        [void]$candidates.Add([pscustomobject]@{ Executable = $runtimePath; Prefix = @() })
                        $listedRuntimes = $true
                    }
                }
            }
            if ($name -eq 'py.exe' -and -not $isWindowsApps) {
                # Compatibility with the classic Python Launcher. Never pass
                # -3 to the WindowsApps manager shim if it lists no runtimes.
                [void]$candidates.Add([pscustomobject]@{ Executable = $managerPath; Prefix = @('-3') })
            } elseif ($isWindowsApps -and -not $listedRuntimes) {
                $script:CertificatePythonProbeFailures += "$managerPath is present, but Python Install Manager did not report an installed runtime."
            }
        }
    }

    foreach ($command in @(Get-Command 'python.exe' -CommandType Application -All -ErrorAction SilentlyContinue)) {
        # The WindowsApps python.exe alias can open the Store instead of Python.
        if ($command.Source -match '\\Microsoft\\WindowsApps\\') { continue }
        [void]$candidates.Add([pscustomobject]@{ Executable = $command.Source; Prefix = @() })
    }
    # Keep path discovery separate from dependency imports. On Windows
    # PowerShell 5.1, pywin32 GUI imports can interfere with captured output.
    # Avoid a quoted assert message: Windows PowerShell 5.1 can misquote that
    # embedded string when forwarding the -c argument to python.exe.
    $pathProbe = 'import sys; assert sys.version_info >= (3, 11); print(sys.executable)'
    $packageProbe = ''
    if ($RequirePackages) {
        $packageProbe = 'import tkinter, openpyxl, reportlab; '
        if ($RequireOutlook) { $packageProbe += 'import pythoncom, win32com.client, win32api, win32gui; ' }
    }
    foreach ($candidate in $candidates) {
        $arguments = @($candidate.Prefix) + @('-c', $pathProbe)
        try {
            $pathOutput = @(& $candidate.Executable @arguments 2>$null)
            $pathExitCode = $LASTEXITCODE
        } catch {
            $pathOutput = @($_.Exception.Message)
            $pathExitCode = 1
        }
        $path = $null
        if ($pathExitCode -eq 0) {
            foreach ($line in @($pathOutput | Select-Object -Last 3)) {
                $candidatePath = ([string]$line).Trim().Trim('"')
                if ($candidatePath -and (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
                    $path = $candidatePath
                    break
                }
            }
        }
        $label = $candidate.Executable
        if ($candidate.Prefix.Count) { $label += ' ' + ($candidate.Prefix -join ' ') }
        if (-not $path) {
            $detail = (($pathOutput | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ }) -join ' ')
            if ($detail.Length -gt 280) { $detail = $detail.Substring(0, 280) + '...' }
            if (-not $detail) { $detail = "Python path check exited with code $pathExitCode without returning a usable interpreter path" }
            $script:CertificatePythonProbeFailures += "$label - $detail"
            continue
        }
        if ($RequirePackages) {
            try {
                $packageOutput = @(& $path -c $packageProbe 2>&1)
                $packageExitCode = $LASTEXITCODE
            } catch {
                $packageOutput = @($_.Exception.Message)
                $packageExitCode = 1
            }
            if ($packageExitCode -ne 0) {
                $detail = (($packageOutput | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ }) -join ' ')
                if ($detail.Length -gt 280) { $detail = $detail.Substring(0, 280) + '...' }
                if (-not $detail) { $detail = "required-package check exited with code $packageExitCode" }
                $script:CertificatePythonProbeFailures += "$path - $detail"
                continue
            }
        }
        return $path
    }
    return $null
}

try {
    Set-Location -LiteralPath $PSScriptRoot
    $certificatePython = Resolve-CertificatePython -RequirePackages ($Mode -ne 'Setup') -RequireOutlook ($Mode -ne 'Desktop')
    if (-not $certificatePython -and $Mode -eq 'Daily') {
        # Let the guarded runner record missing-package failures without a popup.
        $certificatePython = Resolve-CertificatePython -RequirePackages $false
    }
    if (-not $certificatePython) {
        $diagnostics = @($script:CertificatePythonProbeFailures | Select-Object -Unique)
        if ($diagnostics.Count) {
            throw ('No usable Python 3.11+ installation was found. Checks tried:' + [Environment]::NewLine + ' - ' + ($diagnostics -join ([Environment]::NewLine + ' - ')) + [Environment]::NewLine + 'Run Setup Once.bat from this Windows account, or ask IT to repair Python/packages for this account.')
        }
        throw 'No Python installation or Python Install Manager runtime was found. Check Python Manager for an installed 3.11+ runtime, then run Setup Once.bat from this Windows account.'
    }
    if ($Mode -eq 'Diagnose') {
        Write-Host "Python check passed. Interpreter: $certificatePython"
        Write-Host 'Python 3.11+ and the required packages are available.'
        exit 0
    }
    if ($Mode -eq 'Setup') {
        $pipArguments = @('-m', 'pip', 'install')
        $virtual = & $certificatePython -c 'import sys; print(int(sys.prefix != sys.base_prefix))'
        if ($LASTEXITCODE -ne 0) { throw 'Could not check the Python installation.' }
        if ([string]$virtual -eq '0') { $pipArguments += '--user' }
        $pipArguments += @('-r', (Join-Path $PSScriptRoot 'requirements.txt'))
        & $certificatePython @pipArguments
        if ($LASTEXITCODE -ne 0) { throw 'Package installation failed. Check internet access or ask IT for help.' }
        & $certificatePython -c 'import tkinter, openpyxl, reportlab, pythoncom, win32com.client, win32api, win32gui'
        if ($LASTEXITCODE -ne 0) { throw 'Setup is incomplete. Ask IT to repair this Python installation.' }
        Write-Host 'Setup complete. Double-click Prepare Certificates.bat.'
        exit 0
    }
    if ($Mode -eq 'Desktop') {
        & $certificatePython (Join-Path $PSScriptRoot 'desktop_app.py')
        exit $LASTEXITCODE
    }
    $certificatePythonw = Join-Path (Split-Path $certificatePython) 'pythonw.exe'
    if (-not (Test-Path -LiteralPath $certificatePythonw -PathType Leaf)) {
        throw 'pythonw.exe was not found beside Python. Ask IT to repair the Python installation.'
    }
    $scriptArgument = '"' + (Join-Path $PSScriptRoot 'scheduled_run.py') + '"'
    if ($Mode -eq 'Daily') {
        $process = Start-Process -FilePath $certificatePythonw -ArgumentList $scriptArgument -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -Wait -PassThru
        exit $process.ExitCode
    }

    # Validate the saved setup without creating drafts or contacting Outlook.
    & $certificatePython (Join-Path $PSScriptRoot 'validate_daily_setup.py')
    if ($LASTEXITCODE -ne 0) { throw 'The daily task was not changed. Correct the setup shown above, then try again.' }
    $certificateUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $action = New-ScheduledTaskAction -Execute $certificatePythonw -Argument $scriptArgument -WorkingDirectory $PSScriptRoot
    $triggers = @((New-ScheduledTaskTrigger -Daily -At '09:00'), (New-ScheduledTaskTrigger -AtLogOn -User $certificateUser))
    # Give network drives and Classic Outlook time to become available at sign-in.
    $triggers[1].Delay = 'PT2M'
    # Do not forcibly terminate a batch halfway through saving Outlook/Excel.
    $taskSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    $principal = New-ScheduledTaskPrincipal -UserId $certificateUser -LogonType Interactive -RunLevel Limited
    Register-ScheduledTask -TaskName 'Certificate Recycling Drafts' -Action $action -Trigger $triggers -Settings $taskSettings -Principal $principal -Description 'Check both Ecoreco plant workbooks at 09:00 and shortly after sign-in; prepare Ready certificate drafts and show at most one quiet summary per day. Never send mail.' -Force | Out-Null
    Write-Host 'Daily check installed for 09:00 and two minutes after Windows sign-in. Missed starts catch up when this user is signed in. Overlapping runs are skipped.'
    exit 0
} catch {
    Write-Host ('Certificate Automation: ' + $_.Exception.Message)
    exit 1
}
