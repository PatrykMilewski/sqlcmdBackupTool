<#

.SYNOPSIS
This script is executer of sqlcmd.exe program with custom arguements. It's mainly designed for creating database copies
back to 1 week + monthly copy, that is untouchable. It also compress backup with 7zip program (it may be changed, by sending
different arguments

.PARAMETER baseName
Name of base, that script will make backup of, needed for backup name.

.PARAMETER backupDestination
Destination of backup, simple UNC path.

.PARAMETER finalDestination
Full path to backup of sql base.

.PARAMETER iniFileName
Name of .ini file with sql commands for execution.

.PARAMETER iniFileLocation
Localisation of .ini file with sql commands for execution.

.PARAMETER logFileName
Name of log file.

.PARAMETER logFileLocation
Localisation of log file.

.PARAMETER lastMonthlyBackupIniName
Name of .ini file with last monthly backup date.

.PARAMETER lastMonthlyBackupIniLocation
Localisation of .ini file with last monthly backup date.
   
.PARAMETER backupDeviceName
Name of MSSQL backup device

.PARAMETER backupDeviceLocation
Localisation of backup device mapped for MSSQL server.

.PARAMETER exeFileName
Name of sqlcmd.exe program.

.PARAMETER exeFileLocation
Localisation of sqlcmd.exe

.PARAMETER compressionLevel
Level of compression.

.PARAMETER emailNotifyFrom
Name of email address from which notification will be send.

.PARAMETER emailNotifyTo
Email of notification target.

.PARAMETER emailNotifyPassword
Password of email address from emailNotifyFrom.

.PARAMETER emailSmtpServer
Address of smtp server, for sending notifications.

.PARAMETER emailSmtpPort
Port of smtp server, for sending notifications.	

.PARAMETER dontSendNotifications
If you add this argument, email notifications will be not send.

.OUTPUTS
This script only returns error codes.
Error codes are organised by binary number, if there is error code number 1, then the youngest bit of return code will
be "1", for example: 00000001. If there is error code number 4, then 4th bit from right will be "1", example: 0001000.
Error codes:
1 - localisation of .ini file is unreachable, it means, that there are no sqlcommands to execute.
2 - localisation of backup destination is unreachable.
3 - localisation of .ini file, that contains last monthly backup date, is unreachable.
4 - localistaion of final destination is unreachable.
5 - localisation of sqlcmd exe file is unreachable.
6 - .ini file is empty, there are no sqlcommands to execute.
7 - removing uncompressed base throwed exception.
8 - copying monthly backup to it's destination, throwed exception.
9 - making of monthly backup failed, because there is already one with this date.
10 - compression of database backup throwed exception.
11 - copying backup from backup device throwed exception.
12 - invoking sqlcmd throwed exception.
13 - problem with sending email notification.

#>

[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True, Position=1)]
	 [String]$baseName,
	[Parameter(Mandatory=$True, Position=2)]
	 [String]$backupDestination,
    [Parameter(Mandatory=$True, Position=3)]
     [String]$backupDeviceName,
	[Parameter(Mandatory=$False)]
	 [String]$finalDestination = ([IO.Path]::Combine($backupDestination, $baseName)),
	 [String]$iniFileName = "sqlcmdBackup.ini",
	 [String]$iniFileLocation = ([IO.Path]::Combine($PSScriptRoot, $baseName, $iniFileName)),
	 [String]$logFileName = "sqlcmdBackup.txt",
	 [String]$logFileLocation = ([IO.Path]::Combine($PSScriptRoot, $baseName, $logFileName)),
	 [String]$lastMonthlyBackupIniName = "lastMonthlyBackupDate.ini",
	 [String]$lastMonthlyBackupIniLocation = ([IO.Path]::Combine($PSScriptRoot, $baseName, $lastMonthlyBackupIniName)),
     [String]$backupDeviceLocation = ([IO.Path]::Combine("C:\Program Files\Microsoft SQL Server\MSSQL12.MSSQL\MSSQL\Backup", $backupDeviceName)),
     [String]$exeFileName = "SQLCMD.exe",
	 [String]$exeFileLocation = ([IO.Path]::Combine("C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\110\Tools\Binn", $exeFileName)),
     [String]$emailNotifyFrom = "powiadomienia.madler@gmail.com",
	 [String]$emailNotifyTo = "patryk.milewski@gmail.com",
	 [String]$emailNotifyPassword,
	 [String]$emailSmtpServer = "smtp.gmail.com",
	 [String]$emailSmtpPort = "587",
	 [Switch]$dontSendNotifications = $True,
	[ValidateSet('None', 'Fast', 'Low', 'Normal', 'High', 'Ultra')]
     [String]$compresssionLevel = "Normal"
)

Set-Variable -Name exitCode -Value 0 -Scope script

function addLog($newLog) {
	"[" + (Get-Date) + "] " + $newLog | Add-Content $logFileLocation
}

function sendEmailNotification($errorCode) {
	try {
		$password = ConvertTo-SecureString $emailNotifyPassword -AsPlainText -Force
		$emailCredentials = New-Object pscredential ($emailNotifyFrom, $password)
		$subject = "$env:USERDOMAIN: MSSQL błąd przy tworzeniu kopii bazy $baseName"
		$body = "Podczas tworzenia się kopii na komputerze: $env:COMPUTERNAME wystąpił błąd. Numer błędu: $errorCode. Nazwa bazy: $baseName."
		Send-MailMessage -From $emailNotifyFrom -To -$emailNotifyTo -Subject $subject -Body $body -SmtpServer $emailSmtpServer -Port $emailSmtpPort -Credential $emailCredentials -Encoding UTF8 -UseSsl -ErrorAction Stop
	}
	catch {
		addLog $error[0]
		$Script:exitCode += 4096
	}
}

function testPaths() {
	if (!(Test-Path $iniFileLocation)) {
		$script:exitCode += 1
	}
	if (!(Test-Path $backupDestination)) {
		$script:exitCode += 2
	}
	if (!(Test-Path $lastMonthlyBackupIniLocation)) {
		$script:exitCode += 4
	}
	if (!(Test-Path -Path $finalDestination)) {
		$script:exitCode += 8
	}
    if (!(Test-Path $exeFileLocation)) {
		$script:exitCode += 16
	}
}

function readLastMonthlyBackupDate() {
	if (Test-Path $lastMonthlyBackupIniLocation) {
		$iniFile = [System.IO.File]::OpenText($lastMonthlyBackupIniLocation)
		return $iniFile.ReadLine()
	}
	else {
		return "00000000"
	}
}

function isMonthlyBackupNeeded([String]$todayDate, [String]$savedDate) {
	if ($savedDate -eq "00000000") {
		return $True
	}
	$firstDate = [datetime]($todayDate.Substring(0,4) + "-" + $todayDate.Substring(4,2) + "-" + $todayDate.Substring(6,2))
	$secondDate = [datetime]($savedDate.Substring(0,4) + "-" + $savedDate.Substring(4,2) + "-" + $savedDate.Substring(6,2))
	$dayDifference = $firstDate - $secondDate
	if ($dayDifference.Days -gt 30) {
		return $True
	}
	else {
		return $False
	}
}

if (!(Test-Path $logFileLocation)) {
    addLog "Log file created."
}

$date = Get-Date
$dateMonth = ""
$dateDay = ""
if ($date.Month -lt 10) {
	$dateMonth = "0" + [String]$date.Month
}
else {
	$dateMonth = [String]$date.Month
}
if ($date.Day -lt 10) {
	$dateDay = "0" + [String]$date.Day
}
else {
	$dateDay = [String]$date.Day
}
$todayDate = [String]$date.Year + $dateMonth + $dateDay
$weekDay = (Get-Date).DayOfWeek
$lastMonthlyBackup = readLastMonthlyBackupDate
$makeMonthlyBackup = isMonthlyBackupNeeded $todayDate $lastMonthlyBackup

$baseUncompressedName = $baseName + "-" + $todayDate + ".bak"

$sqlCmdlineQuery = Get-Content $iniFileLocation
if ($sqlCmdlineQuery -eq $null) {
	#iniFile is empty
	$script:exitCode += 32
}
else {
    $sqlCmdlineQuery = $sqlCmdlineQuery -replace '&BACKUP_DEVICE_LOC&', $backupDeviceLocation
    $sqlCmdlineQuery = $sqlCmdlineQuery -replace '&BACKUP_NAME&', $baseUncompressedName
    $sqlCmdlineQuery = $sqlCmdlineQuery -replace '&BASE_NAME&', $baseName
    $sqlCmdlineQuery = '"' + $sqlCmdlineQuery + '"'
}

$arguments = "-E -Q $sqlCmdlineQuery"

try {
    $stderr = [IO.Path]::Combine($PSScriptRoot, $baseName, "sqlcmdStdErr.log")
    $stdout = [IO.Path]::Combine($PSScriptRoot, $baseName, "sqlcmdStdOut.log")
	$returnValue = Start-Process -NoNewWindow -FilePath $exeFileLocation -ArgumentList $arguments -RedirectStandardOutput $stdout -RedirectStandardError $stderr -ErrorAction Stop -Wait
    $stderrOutput = get-content $stderr
    $stdoutOutput = get-content $stdout
    if ($stderrOutput -ne $null) {
        addLog "Sqlcmd standard error output: $stderrOutput"
    }
    if ($stdoutOutput -ne $null -and !($stdoutOutput -like "*The backup set on file * is valid.*")) {
        addLog "Sqlcmd standard output: $stdoutOutput"
    }
    Remove-item $stderr
    Remove-item $stdout

	$finalDestWithWeekDay = [IO.Path]::Combine($finalDestination, $weekDay)
	if (Test-Path -Path $finalDestWithWeekDay) {
		Remove-Item $finalDestWithWeekDay -Recurse
	}
	New-Item -Type Directory $finalDestWithWeekDay

	try {
		Copy-Item $backupDeviceLocation $finalDestWithWeekDay -ErrorAction Stop
		$destOfCompArchieve = [IO.Path]::Combine($finalDestWithWeekDay, ($baseName + "-" + $todayDate + ".zip"))
		try {
            Compress-7Zip -Path $finalDestWithWeekDay\$backupDeviceName -ArchiveFileName $destOfCompArchieve -CompressionLevel $compresssionLevel -ErrorAction Stop

            try {
			    Remove-Item $finalDestWithWeekDay\$backupDeviceName -ErrorAction Stop
		    }
		    catch {
			    #problem with removing uncompressed backup
			    addLog $error[0]
				$script:exitCode += 64
		    }
			if ($makeMonthlyBackup -eq $True) {
				$monthlyBackupDest = [IO.Path]::Combine($finalDestination, $todayDate)
				#check if folder doesn't exist, or exist and is empty (hidden files doesn't count)
				if (!(Test-Path -Path $monthlyBackupDest) -or ((Test-Path -Path $monthlyBackupDest) -and (!(Test-Path -Path $monthlyBackupDest\*)))) {
					if (!(Test-Path -Path $monthlyBackupDest)) {
						New-Item -ItemType Directory $monthlyBackupDest
					}
					try {
						Copy-Item $destOfCompArchieve $monthlyBackupDest -ErrorAction Stop
						New-Item -ItemType File $lastMonthlyBackupIniLocation -Value $todayDate -Force
					}
					catch {
						addLog $error[0]
						#problem with copying archieve to monthly backup destination
						$script:exitCode += 128
					}
				}
				else {
					addLog "Couldn't create monthly backup, because there is already one with today's date."
                    New-Item -ItemType File $lastMonthlyBackupIniLocation -Value $todayDate -Force
					$script:exitCode += 256
				}
			}
		}
		catch {
			addLog $error[0]
			#problem with compression of database backup
			$script:exitCode += 512
		}
	}
	catch {
		addLog $error[0]
		#problem with copying backup from backup device to backup destination
		$script:exitCode += 1024
	}
}
catch {
	addLog $error[0]
	#problem with invoking sqlcmd
	$script:exitCode += 2048
}

$script:exitCode += testPaths
if ($script:exitCode -ne 0) {
	addLog "Script exited with code: $script:exitCode"
	if ($dontSendNotifications -eq $False) {
		sendEmailNotification $script:exitCode
	}
}

return $script:exitCode