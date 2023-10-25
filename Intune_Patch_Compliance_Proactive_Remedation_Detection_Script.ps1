cls
<# 
.SYNOPSIS
  <Intune_Patch_Compliance_Proactive_Remedation_Detection_Script>
.DESCRIPTION
  <Intune_Patch_Compliance_Proactive_Remedation_Detection_Script>
.Demo
<YouTube video link--> https://www.youtube.com/watch?v=hAVgNvEAdKc
.INPUTS
  <Provide all required inforamtion in User Input Section>
.OUTPUTS
  <It will update the patching status under Proactive Remedation Detection Output>
.NOTES
  Version:        1.0
  Author:         Chander Mani Pandey
  Creation Date:  12 Jan 2023
  Find Author on 
  Youtube:-        https://www.youtube.com/@chandermanipandey8763
  Twitter:-        https://twitter.com/Mani_CMPandey
  Facebook:-       https://www.facebook.com/profile.php?id=100087275409143&mibextid=ZbWKwL
  LinkedIn:-       https://www.linkedin.com/in/chandermanipandey
  Reddit:-         https://www.reddit.com/u/ChanderManiPandey 
 #>

cls

#-------------------------------------------------- Last BootUp Time ---------------------------------------------------------------------------------------------

# Check if fast start up is enabled
$fastStartUp = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -Name 'HiberbootEnabled').HiberbootEnabled

# Get the last boot time from the Win32_OperatingSystem class
$lastBootTime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime

# If fast start up is enabled, get the time of the last full shutdown
if ($fastStartUp) {
    $lastFullShutdownTime = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Windows' -Name 'ShutdownTime').ShutdownTime
    # Compare the last boot time and the last full shutdown time
    # and use the more recent time as the last boot time
    if ($lastFullShutdownTime -gt $lastBootTime) {
        $lastBootTime = $lastFullShutdownTime
    }
}

# Convert the last boot time to a DateTime object and display it
$LastBootUpdate = $lastBootTime.ToString("MM/dd/yy")

#----------------------------------------- LastSuccessScanDate,LastScan Daysfrom Current Date ---------------------------------------------------------------------------

$lastscandate = (New-Object -ComObject "microsoft.update.autoupdate").results
$sd = (get-date).Date - ($lastscandate.LastSearchSuccessDate).Date 
$LastScanInDaysfromCurrentDate = $sd.Days 
$LastSuccessScanDate = ($lastscandate.LastSearchSuccessDate).ToString("MM/dd/yy")


#--------------------------------------------------free space on the C: drive in GB-----------------------------------------------------------------------------
# Get the free space on the C: drive in bytes
$freeSpace = (Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'" | Select-Object -ExpandProperty FreeSpace)

# Convert the free space to GB
$freeSpaceGB = [Math]::Round(($freeSpace / 1GB), 2)

#------------------------------------------Required Services Status---------------------------------------------------------------------------------------------------
#----------------------------------------Windows Update Agent--------------------------------
$serviceName = 'wuauserv'
$service = Get-Service -Name $serviceName

if ($service.Status -eq 'Running') {
    $WindowsUpdateService =  "Running"+ " (" + $($service.StartType) + ")"
}
else {
    #Write-Output "The Windows Update service is not running. Its startup type is $($service.StartType)."
    $WindowsUpdateService =  "NotRunning"+ " (" + $($service.StartType) + ")"
}

#$WindowsUpdateService

#----------------------------------------Update Orchestrator Service--------------------------------
$serviceName = 'UsoSvc'
$service = Get-Service -Name $serviceName

if ($service.Status -eq 'Running') {
    $UsoService =  "Running"+ " (" + $($service.StartType) + ")"
}
else {
   # Write-Output "The Windows Update service is not running. Its startup type is $($service.StartType)."
    $UsoService =  "NotRunning"+ " (" + $($service.StartType) + ")"
}

#$UsoService

#------------------------------------------Microsoft Update Health Service---------------------------------

$serviceName = 'Microsoft Update Health Service'
$service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
if ($service) {


    if ($service.Status -eq 'Running') {
    $uhService =  "Running"+ " (" + $($service.StartType) + ")"
    }
    else {
   # Write-Output "The Windows Update service is not running. Its startup type is $($service.StartType)."
    $uhService =  "NotRunning"+ " (" + $($service.StartType) + ")"
    }
    } 

else {
   $uhService = "$serviceName is not installed."
}

#$uhService

#---------------------------------------------Microsoft Account Sign-in Assistant--------------------------------

$serviceName = 'wlidsvc'
$service = Get-Service -Name $serviceName

if ($service.Status -eq 'Running') {
    $wlidService =  "Running"+ " (" + $($service.StartType) + ")"
}
else {
   # Write-Output "The Windows Update service is not running. Its startup type is $($service.StartType)."
    $wlidService =  "NotRunning"+ " (" + $($service.StartType) + ")"
}
#$wlidService
#----------------------------------------Windows SKU--------------------------------
$WindowsSKU= (Get-WmiObject -Class Win32_OperatingSystem).Caption
#$WindowsSKU
#-----------------------------------------Windows SKU--------------------------------
#-----------------------------------------MajorVersion------------------------------
$version = (Get-WmiObject -Class Win32_OperatingSystem).Version
$Ver = $version.Split(".")[2]
#-----------------------------------------MiverVersion------------------------------
$registryPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
$registryValue = "UBR"
$minorBuildNumber = (Get-ItemProperty -Path $registryPath -Name $registryValue).$registryValue
$OS_Version = If($Ver -eq "10249")
            {'1507'} 
          elseif($Ver -eq 10586)
           {"1511"}
          elseif($Ver -eq 14393) 
           {"1607"} 
         elseif($Ver -eq 15063)
           {"1703"} 
         elseif($Ver -eq 16299)
           {"1709"} 
         elseIf($Ver -eq 17134)
           {"1803"} 
         elseIf($Ver -eq 17763)
           {'1809'} 
         elseIf($Ver -eq 18362)
           {"1903"} 
         elseIf($Ver -eq 18363)
           {"1909"} 
         elseIf($Ver -eq 19041)
           {"2004"} 
         elseIf($Ver -eq 19042)
           {"20H2"} 
         elseIf($Ver -eq 19043)
           {"21H1"} 
         elseIf($Ver -eq 19044)
           {"21H2"} 
         elseIf($Ver -eq 19045)
           {"22H2"}
         elseIf($Ver -ge 20000)
           {"Win11"} 
          elseIf($Ver -eq 0)
           {"Need To Check"}
         else 
           {$Ver }
 #------------------------------------------------End of Support OS-------------------------------------------------------------
 
$ver = $OS_Version
$EOL = (1507,1511,1607,1703,1709,1803,1809,1903,1909,2004)
$InSupport = ( "20H2","21H1","21H2","22H2","Win11")
$NoVersionInfo =  ("0")
$EOLStatus = $NULL

if ($EOL -contains $ver) {
    $EOLStatus = "EOL"
} elseif ($InSupport -contains $ver) {
     $EOLStatus = "Supported OS"
} elseif ($NoversionInfor -contains $ver) {
    $EOLStatus = "NoOsVersionInfo"
} else {
    $EOLStatus = $OS_Version
}

#-------------------------------------------RebootPending--------------------------------------------------------------------
$RebootPending = $null
If ( Test-path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" ) 
{ 
$a = "Reboot Pending"
$b = "$((Get-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -Name RebootRequiredSince).RebootRequiredSince)"
$RebootPending =$a +'('+$b+')'
} 

else { 
$RebootPending = "No Reboot Pending"
}
#$RebootPending

#-------------------------------------------Make and Model--------------------
$computer = Get-WmiObject -Class Win32_ComputerSystem -Namespace "root\cimv2"
$make = $computer.Manufacturer
$model = $computer.Model

#---------------------------------------------Serial number---------------------------------------------------
$bios = Get-WmiObject -Class Win32_BIOS
$serialNumber = $bios.SerialNumber


#-----------------------------------------Detecting Machine is Compliant against Latest Patch Tuesday against -----------------------------------------------------------

$CollectedData = $PatchDetails = $LatestPatches = @()

$InstlPatch = $InstlPatchRD = $OSBuild = $String = ""

$PatchReleaseDays = 0
#=====================Checking OS Version Missing===============--------------------------------------------------------------------------------------------------------
$OSBuild = ([System.Environment]::OSVersion.Version).Build
IF (!($OSBuild)) {
    $String = 'Failed to Find Build Info'
    Write-Host $String
    exit 1
}

#===========Detecting latest Intalled KB==========================
[string]$InstlPatch = (Get-HotFix | Where-Object {$_.Description -match 'security'} | Sort-Object HotFixID -Descending | Select-Object -First 1).HotFixID
IF (!($InstlPatch)) {
    $String = 'Failed To Find Installed Patch'
    Write-Host $String
    exit 1
}
#Windows 11 Update HistorY URL
$URI = 'https://aka.ms/Windows11UpdateHistory'
$CollectedData += (Invoke-WebRequest -Uri $URI -UseBasicParsing -ErrorAction Continue).Links
#Windows 10 Update HistorY URLWindows 11 URL
#$URI = 'https://aka.ms/WindowsUpdateHistory'
$URI = "https://support.microsoft.com/en-us/help/4043454";
$CollectedData += (Invoke-WebRequest -Uri $URI -UseBasicParsing -ErrorAction Continue).Links

#$CollectedData | Export-Csv -Path "c:\report.csv" -NoTypeInformation

#============Checking if able to downlaod MS patch from MS sites(Internet)===========================================================================================

IF (!($CollectedData)) {
    $String = 'Failed To Download MSPatchList'
    Write-Host $String
    exit 1
}

#============Checking if able to Find MS patch =====================================================================================================================

$CollectedData = ($CollectedData | Where-Object {$_.class -eq 'supLeftNavLink' -and $_.outerHTML -match 'KB' -and $_.outerHTML -notmatch 'out-of-band' -and $_.outerHTML -notmatch 'preview' -and $_.outerHTML -notmatch 'mobile' -and $_.outerHTML -match $OSBuild}).outerHTML;
$CollectedData = $CollectedData | Select-Object -Unique
IF (!($CollectedData)) {
    $String = 'Failed To Find MSPatch'
    Write-Host $String
    exit 1
}

Foreach ($Line in $CollectedData) {
	$ReleaseDate = $PatchID = ""
    $ReleaseDate = (($Line.Split('>')[1]).Split('&')[0]).trim()
    IF ($ReleaseDate -match 'build') {
        $ReleaseDate = ($ReleaseDate.split('-')[0]).trim()
    }
	$PatchID = ($Line.Split(' ;-') | Where-Object {$_ -match 'KB'}).trim()
    $PatchDetails += [PSCustomObject] @{MajorBuild = $OSBuild; PatchID = $PatchID; ReleaseDate = $ReleaseDate;}
}
$PatchDetails = $PatchDetails | Select-Object MajorBuild,PatchID,ReleaseDate -Unique | Sort-Object PatchID -Descending;
IF (!($PatchDetails)) {
    $String = 'Failed To Find Patch List'
    Write-Host $String
    exit 1
}
$Today = Get-Date; $LatestDate = ($PatchDetails | Select-Object -First 1).ReleaseDate
$DiffDays = ([datetime]$Today - [datetime]$LatestDate).Days
[Int]$DateVar = $PatchReleaseDays/28
#IF ([int]$DiffDays -gt [int]$PatchReleaseDays) 
If ($DateVar -eq 0)
{
    $LatestPatches += $PatchDetails | Select-Object -First 1
}
ELSE {
    $LatestPatches += $PatchDetails | Select-Object -Skip $DateVar -First 1
    
}
IF (!($LatestPatches)) {
    $String = 'Failed To Find Latest Patch'
    Write-Host $String
    exit 1
}
Foreach ($BLD in $PatchDetails) {IF ($InstlPatch -eq $BLD.PatchID) {$InstlPatchRD = $BLD.ReleaseDate}}
IF (!($InstlPatchRD)) {
    $String = $InstlPatch + ';Failed To Find RlsDt'
    Write-Host $String
    exit 1
}
#=================================== Converting dates in MM/dd/yy format================================================================================================

$LPRD = Get-Date -Date $LatestPatches.releasedate
$LatestPatchesreleasedate = $LPRD.ToString("MM/dd/yy")

$IPRD = Get-Date -Date $InstlPatchRD
$FinalInstlPatchRD = $IPRD.ToString("MM/dd/yy")


$KBN1 = $KBN2 = "";
[int]$KBN1 = ($InstlPatch).Replace('KB','')
[int]$KBN2 = ($LatestPatches.PatchID).Replace('KB','')
IF ([int]$KBN1 -ge [int]$KBN2) {
    # Compliant against which month Patch Tuesday,latest Intalled Patch on device,latest Intalled Patch realeased date,LastSuccessScanDate,LastScanInDaysfromCurrentDate,LastBootUp time,$WindowsUpdateService,  $UsoService,$UhService,$wlidService,$WindowsSKU,$OS_Version , $EOLStatus , $RebootPending ,$make ,$model ,$serialNumber
    $String = 'Compliant'+';'+ $LatestPatchesreleasedate +';'+$InstlPatch+';'+ $FinalInstlPatchRD +';'+$LastSuccessScanDate +';'+ $LastScanInDaysfromCurrentDate +';'+ $LastBootUpdate +';'+ $freeSpaceGB +';'+  $WindowsUpdateService +';'+  $UsoService +';'+  $UhService +';'+  $wlidService +';'+  $WindowsSKU+';'+ $OS_Version +';'+ $EOLStatus +';'+ $RebootPending +';'+ $make +';'+ $model +';'+ $serialNumber
    Write-Host $String
    exit 0
}
ELSE {
    $String = 'Non Compliant'+';'+ $LatestPatchesreleasedate +';'+$InstlPatch+';'+ $FinalInstlPatchRD +';'+$LastSuccessScanDate +';'+ $LastScanInDaysfromCurrentDate +';'+ $LastBootUpdate +';'+ $freeSpaceGB +';'+  $WindowsUpdateService +';'+  $UsoService +';'+  $UhService +';'+  $wlidService +';'+  $WindowsSKU+';'+ $OS_Version +';'+ $EOLStatus +';'+ $RebootPending +';'+ $make +';'+ $model +';'+ $serialNumber
    Write-Host $String
    exit 1
}

#-------------------------------End-------------------------------------------------------------------------------------------------------------------------------
