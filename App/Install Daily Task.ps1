param(
    [ValidateSet('Install', 'Setup', 'Desktop', 'Recovery', 'Daily', 'Diagnose', 'Reschedule', 'Disable', 'Snooze', 'CancelSnooze')]
    [string]$Mode = 'Install',
    [string]$ReminderTime = '',
    [string]$SnoozeAt = ''
)

$ErrorActionPreference = 'Stop'
$certificateRoot = if ((Split-Path $PSScriptRoot -Leaf) -eq 'App') { Split-Path $PSScriptRoot -Parent } else { $PSScriptRoot }
$certificateBootstrapPythonDir = Join-Path $env:LOCALAPPDATA 'Programs\Python\CertificateAutomation-3.13'

function Resolve-CertificatePython {
    param([bool]$RequirePackages = $true)
    $candidates = [System.Collections.Generic.List[object]]::new()
    $script:CertificatePythonProbeFailures = @()
    $localPython = Join-Path $certificateRoot '.venv\Scripts\python.exe'
    if (Test-Path -LiteralPath $localPython -PathType Leaf) {
        [void]$candidates.Add([pscustomobject]@{ Executable = $localPython; Prefix = @() })
    }
    $bootstrapPython = Join-Path $certificateBootstrapPythonDir 'python.exe'
    if (Test-Path -LiteralPath $bootstrapPython -PathType Leaf) {
        [void]$candidates.Add([pscustomobject]@{ Executable = $bootstrapPython; Prefix = @() })
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
    $pathProbe = 'import sys; assert sys.version_info >= (3, 11) and sys.maxsize > 2**32; print(sys.executable)'
    $packageProbe = ''
    if ($RequirePackages) {
        $packageProbe = 'import tkinter, openpyxl, reportlab, pypdf, fontTools, pypdfium2, PIL; '
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

function Install-CertificatePythonRuntime {
    if (-not [Environment]::Is64BitOperatingSystem) {
        throw 'Automatic Python setup requires 64-bit Windows. Ask IT to install 64-bit Python 3.11 or newer with Tkinter, then run Setup Once.bat again.'
    }
    # Python 3.13 remains on the official standalone Windows installer path.
    # Update the pinned version and digest from python.org when refreshing setup.
    $pythonVersion = '3.13.17'
    $installerName = "python-$pythonVersion-amd64.exe"
    $installerUri = "https://www.python.org/ftp/python/$pythonVersion/$installerName"
    $expectedSha256 = '5DA012101BA2E17113A9A1CD901E9CE1441086F2B7ECB5DE06EC94163899A190'
    $installerPath = Join-Path $env:TEMP ("certificate-automation-python-$pythonVersion-" + [Guid]::NewGuid().ToString('N') + '.exe')

    try {
        Write-Host "No supported Python installation was found. Downloading the official Python $pythonVersion installer..."
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $installerUri -OutFile $installerPath -UseBasicParsing
        $actualSha256 = (Get-FileHash -LiteralPath $installerPath -Algorithm SHA256).Hash.ToUpperInvariant()
        if ($actualSha256 -ne $expectedSha256) {
            throw 'The downloaded Python installer did not match its published SHA-256 checksum. The file was not run.'
        }

        Write-Host 'Installing Python for this Windows account...'
        $installerArguments = '/quiet InstallAllUsers=0 PrependPath=0 Include_launcher=0 Include_tcltk=1 Include_pip=1 Include_test=0 Include_doc=0 Shortcuts=0 TargetDir="' + $certificateBootstrapPythonDir + '"'
        $process = Start-Process -FilePath $installerPath -ArgumentList $installerArguments -Wait -PassThru
        if ($process.ExitCode -ne 0) {
            throw "The Python installer exited with code $($process.ExitCode)."
        }
        if (-not (Test-Path -LiteralPath (Join-Path $certificateBootstrapPythonDir 'python.exe') -PathType Leaf)) {
            throw 'Python setup finished without creating the expected interpreter. Ask IT to install Python 3.13 for this Windows account.'
        }
    } catch {
        throw ('Could not install Python automatically. Check internet access to python.org and Windows app-install permissions, then retry Setup Once.bat. Details: ' + $_.Exception.Message)
    } finally {
        if (Test-Path -LiteralPath $installerPath -PathType Leaf) {
            Remove-Item -LiteralPath $installerPath -Force -ErrorAction SilentlyContinue
        }
    }
}

try {
    Set-Location -LiteralPath $certificateRoot
    if ($Mode -eq 'CancelSnooze') {
        $snoozeTaskName = 'Ecoreco Certificate Reminder Snooze'
        $snoozeTask = Get-ScheduledTask -TaskName $snoozeTaskName -ErrorAction SilentlyContinue
        if ($snoozeTask) {
            Unregister-ScheduledTask -TaskName $snoozeTaskName -Confirm:$false
        }
        Write-Host 'Any pending one-hour reminder was cleared.'
        exit 0
    }
    if ($Mode -eq 'Disable') {
        foreach ($taskName in @('Ecoreco Certificate Reminders', 'Ecoreco Certificate Reminder Snooze', 'Certificate Recycling Drafts', 'Certificate Recycling Daily Reminder')) {
            $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            if ($task) { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false }
        }
        Write-Host 'Daily certificate reminders are turned off.'
        exit 0
    }
    if ($Mode -eq 'Snooze') {
        if ($SnoozeAt -notmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$') {
            throw 'The one-hour reminder time is invalid. Close the reminder and try again.'
        }
        $runAt = [DateTime]::ParseExact($SnoozeAt, 'yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
        if ($runAt -le (Get-Date)) { throw 'The reminder time has already passed. Choose Remind me in 1 hour again.' }
        $certificatePython = Resolve-CertificatePython -RequirePackages $false
        if (-not $certificatePython) { throw 'Python could not be found, so Windows could not schedule the reminder. Keep the app open and ask IT to repair Python.' }
        $certificatePythonw = Join-Path (Split-Path $certificatePython) 'pythonw.exe'
        if (-not (Test-Path -LiteralPath $certificatePythonw -PathType Leaf)) {
            throw 'pythonw.exe was not found beside Python. Windows could not schedule the reminder.'
        }
        $runner = Join-Path $PSScriptRoot 'scheduled_run.py'
        $arguments = '"' + $runner + '" --snoozed-reminder'
        $action = New-ScheduledTaskAction -Execute $certificatePythonw -Argument $arguments -WorkingDirectory $PSScriptRoot
        $trigger = New-ScheduledTaskTrigger -Once -At $runAt
        $taskSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval ([TimeSpan]::FromMinutes(5))
        $certificateUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $principal = New-ScheduledTaskPrincipal -UserId $certificateUser -LogonType Interactive -RunLevel Limited
        Register-ScheduledTask -TaskName 'Ecoreco Certificate Reminder Snooze' -Action $action -Trigger $trigger -Settings $taskSettings -Principal $principal -Description 'Show a one-time Ecoreco certificate review reminder in the signed-in user session. Does not wake the computer, scan files, or send mail.' -Force | Out-Null
        Write-Host "One-hour reminder scheduled for $($runAt.ToString('HH:mm')). It will wait until this user is signed in and the computer is awake; it will not wake the computer."
        exit 0
    }
    if ($Mode -ne 'Recovery' -and -not $ReminderTime) {
        $settingsPath = Join-Path $certificateRoot 'settings.local.json'
        if (Test-Path -LiteralPath $settingsPath -PathType Leaf) {
            try {
                $savedSettings = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json
                if ($savedSettings.reminder_time) { $ReminderTime = [string]$savedSettings.reminder_time }
            } catch {
                throw 'Cannot read the saved reminder time. Open the app, correct Settings, and try again.'
            }
        }
        if (-not $ReminderTime) { $ReminderTime = '11:00' }
    }
    if ($Mode -ne 'Recovery' -and $ReminderTime -notmatch '^(?:[01][0-9]|2[0-3]):[0-5][0-9]$') {
        throw 'Reminder time must use 24-hour format, for example 11:00 or 17:30.'
    }
    if ($Mode -eq 'Reschedule') {
        $taskName = 'Ecoreco Certificate Reminders'
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if (-not $existingTask) {
            throw 'The daily reminder is not installed yet. Use Save & update reminder in the app.'
        }
        $certificateUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $triggers = @((New-ScheduledTaskTrigger -Daily -At $ReminderTime), (New-ScheduledTaskTrigger -AtLogOn -User $certificateUser))
        # Sign-in remains a catch-up opportunity, but reminders themselves wait until the chosen time.
        $triggers[1].Delay = 'PT2M'
        $existingTask.Triggers = $triggers
        $existingTask.Description = "Check both Ecoreco plant workbooks daily at $ReminderTime and shortly after sign-in; remind about certificates from inward day 25 until Completed or N/A. Never creates drafts or sends mail."
        # Set-ScheduledTask has no -Description parameter on Windows PowerShell.
        Set-ScheduledTask -InputObject $existingTask | Out-Null
        Write-Host "Daily reminder time updated to $ReminderTime. The sign-in catch-up remains enabled."
        exit 0
    }
    $certificatePython = Resolve-CertificatePython -RequirePackages ($Mode -ne 'Setup' -and $Mode -ne 'Snooze')
    if (-not $certificatePython -and $Mode -eq 'Setup') {
        Install-CertificatePythonRuntime
        $certificatePython = Resolve-CertificatePython -RequirePackages $false
    }
    if (-not $certificatePython -and $Mode -eq 'Daily') {
        # Let the guarded runner record missing-package failures without a popup.
        $certificatePython = Resolve-CertificatePython -RequirePackages $false
    }
    if (-not $certificatePython) {
        $diagnostics = @($script:CertificatePythonProbeFailures | Select-Object -Unique)
        if ($diagnostics.Count) {
            throw ('No usable 64-bit Python 3.11+ installation was found. Checks tried:' + [Environment]::NewLine + ' - ' + ($diagnostics -join ([Environment]::NewLine + ' - ')) + [Environment]::NewLine + 'Run Setup Once.bat from this Windows account to install Python and packages, or ask IT to repair this account.')
        }
        throw 'No 64-bit Python 3.11+ installation or Python Install Manager runtime was found. Run Setup Once.bat from this Windows account to install Python and the required packages automatically.'
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
        $lockedRequirements = Join-Path $PSScriptRoot 'requirements.lock.txt'
        $requirementsPath = if (Test-Path -LiteralPath $lockedRequirements -PathType Leaf) { $lockedRequirements } else { Join-Path $PSScriptRoot 'requirements.txt' }
        $pipArguments += @('-r', $requirementsPath)
        & $certificatePython @pipArguments
        if ($LASTEXITCODE -ne 0) { throw 'Package installation failed. Check internet access or ask IT for help.' }
        & $certificatePython -c 'import tkinter, openpyxl, reportlab, pypdf, fontTools, pypdfium2, PIL'
        if ($LASTEXITCODE -ne 0) { throw 'Setup is incomplete. Ask IT to repair this Python installation.' }
        Write-Host 'Setup complete. Double-click Prepare Certificates.bat.'
        exit 0
    }
    if ($Mode -eq 'Recovery') {
        & $certificatePython (Join-Path $PSScriptRoot 'recover_workspace.py')
        exit $LASTEXITCODE
    }
    if ($Mode -eq 'Desktop') {
        $certificatePythonw = Join-Path (Split-Path $certificatePython) 'pythonw.exe'
        $dashboard = Join-Path $PSScriptRoot 'web_app.py'
        if (Test-Path -LiteralPath $certificatePythonw -PathType Leaf) {
            Start-Process -FilePath $certificatePythonw -ArgumentList ('"' + $dashboard + '"') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden
            Write-Host 'The local certificate dashboard is starting in your browser.'
            exit 0
        }
        & $certificatePython $dashboard
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

    # Validate connected plant files; local-only workspaces also support reminders.
    & $certificatePython (Join-Path $PSScriptRoot 'validate_daily_setup.py')
    if ($LASTEXITCODE -ne 0) { throw 'The daily task was not changed. Correct the setup shown above, then try again.' }
    $certificateUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $action = New-ScheduledTaskAction -Execute $certificatePythonw -Argument $scriptArgument -WorkingDirectory $PSScriptRoot
    $triggers = @((New-ScheduledTaskTrigger -Daily -At $ReminderTime), (New-ScheduledTaskTrigger -AtLogOn -User $certificateUser))
    # Give network drives time to become available at sign-in.
    $triggers[1].Delay = 'PT2M'
    # Do not forcibly terminate a batch halfway through updating PDFs or Excel.
    $taskSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval ([TimeSpan]::FromMinutes(5))
    $principal = New-ScheduledTaskPrincipal -UserId $certificateUser -LogonType Interactive -RunLevel Limited
    Register-ScheduledTask -TaskName 'Ecoreco Certificate Reminders' -Action $action -Trigger $triggers -Settings $taskSettings -Principal $principal -Description "Check both Ecoreco plant workbooks at $ReminderTime and shortly after sign-in; remind about certificates from inward day 25 until Completed or N/A. Never creates drafts or sends mail." -Force | Out-Null
    foreach ($oldTaskName in @('Certificate Recycling Drafts', 'Certificate Recycling Daily Reminder')) {
        $oldTask = Get-ScheduledTask -TaskName $oldTaskName -ErrorAction SilentlyContinue
        if ($oldTask) { Unregister-ScheduledTask -TaskName $oldTaskName -Confirm:$false }
    }
    Write-Host "Daily check installed for $ReminderTime and two minutes after Windows sign-in. Missed starts catch up when this user is signed in. Overlapping runs are skipped."
    exit 0
} catch {
    Write-Host ('Certificate Automation: ' + $_.Exception.Message)
    exit 1
}
