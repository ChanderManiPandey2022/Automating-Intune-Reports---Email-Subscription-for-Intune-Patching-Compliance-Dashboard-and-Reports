<#
.SYNOPSIS
  <Intune_Patch_Compliance_Calculation_&_Email_Subscription Using PowerShell>
.DESCRIPTION
  <Intune_Patch_Compliance_Calculation_&_Email_Subscription Using PowerShell>
.Demo
<YouTube video link--> https://www.youtube.com/watch?v=hAVgNvEAdKc
.INPUTS
  <Provide all required inforamtion in User Input Section-line No 96-105 & 142-145>
.OUTPUTS
  <You will get Intune_Patch_Compliance_Calculation_&_Email_Subscription + report in CSV>
.NOTES
  Version:        1.0
  Author:         Chander Mani Pandey
  Creation Date:  12 Jan 2023
  Find Author on 
  Youtube:-        https://www.youtube.com/@chandermanipandey8763
  Twitter:-        https://twitter.com/Mani_CMPandey
  Facebook:-       https://www.facebook.com/profile.php?id=100087275409143&mibextid=ZbWKwL
  LinkedIn:-       https://www.linkedin.com/in/chandermanipandey
  
 #>
 #---------------------------------------------------- User Input Section line No 96-105 & 142-145 ---------------------------------------------------------------------

 Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' 
$error.clear() ## this is the clear error history 
cls
$ErrorActionPreference = 'SilentlyContinue';
#######################################################################################################################################################
 #---------------------------------Patching Against which Patch Tuesday-------------------------------------------------------------------------
 # Create an empty array of PSObject objects
$buildInfoArray = @()
#====================================================================================
#Creating working Folder
New-Item -ItemType Directory -Path $WorkingFolder -Force

# Add each Build and Operating System to the array
"22631,Windows 11 23H2","22623,Windows 11 22H2","22621,Windows 11 22H2 B1","22471,Windows 11 21H2","22468,Windows 11 21H2 B6","22463,Windows 11 21H2 B5",
"22458,Windows 11 21H2 B4","22454,Windows 11 21H2 B3","22449,Windows 11 21H2 B2","22000,Windows 11 21H2 B1","21996,Windows 11 Dev",
"19045,Windows 10 22H2","19044,Windows 10 21H2","19043,Windows 10 21H1","19042,Windows 10 20H2","19041,Windows 10 2004","19008,Windows 10 20H1",
"18363,Windows 10 1909","18362,Windows 10 1903","17763,Windows 10 1809","17134,Windows 10 1803","16299,Windows 10 1709 FC","15254,Windows 10 1709",
"15063,Windows 10 1703","14393,Windows 10 1607","10586,Windows 10 1511","10240,Windows 10 1507","9600,Windows 8.1",
"7601,Windows 7" | ForEach-Object {
    # Create a new PSObject object
    $buildInfo = New-Object -TypeName PSObject

    # Add the Build and Operating System properties to the object
    $buildInfo | Add-Member -MemberType NoteProperty -Name "Build" -Value ($_ -split ",")[0]
    $buildInfo | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value ($_ -split ",")[1]

    # Add the object to the array
    $buildInfoArray += $buildInfo
}

# Print the array of PSObject objects to the screen
#$buildInfoArray

#===============================================================================================================
$CollectedData = $BuildDetails = $PatchDetails = $MajorBuilds = $LatestPatches = @();
$BuildDetails = $buildInfoArray
#Download Windows Master Patch List
Write-Host "Downoading Patch List from Microsoft"-ForegroundColor yellow
$URI = "https://aka.ms/Windows11UpdateHistory";
$CollectedData += (Invoke-WebRequest -Uri $URI -UseBasicParsing -ErrorAction Continue).Links;
$URI = "https://aka.ms/WindowsUpdateHistory";
$CollectedData += (Invoke-WebRequest -Uri $URI -UseBasicParsing -ErrorAction Continue).Links;
#Filter Windows Master Patch List
Write-Host "Filtering Patch List"-ForegroundColor yellow
IF ($CollectedData) {
	$CollectedData = ($CollectedData | Where-Object {$_.class -eq "supLeftNavLink" -and $_.outerHTML -match "KB" -and $_.outerHTML -notmatch "out-of-band" -and $_.outerHTML -notmatch "preview" -and $_.outerHTML -notmatch "mobile"}).outerHTML
}

#Consolidate the Master Patch and Format the output
Write-Host "Consolidating Patch List"-ForegroundColor yellow
Foreach ($Line in $CollectedData) {
	$ReleaseDate = $PatchID = ""; $Builds = @();	
    $ReleaseDate = (($Line.Split(">")[1]).Split("&”)[0]).trim();
        IF ($ReleaseDate -match "build") {$ReleaseDate = ($ReleaseDate.split("-")[0]).trim();}
	$PatchID = ($Line.Split(" ;-") | Where-Object {$_ -match "KB"}).trim();
    $Builds = ($Line.Split(",) ") | Where-Object {$_ -like "*.*"}).trim();
	Foreach ($BLD in $Builds) {
		$MjBld = $MnBld = ""; $MjBld = $BLD.Split(".")[0]; $MnBld = $BLD.Split(".")[1];
            Foreach ($Line1 in $BuildDetails) {
                $BldNo = $OS = ""; $BldNo = $Line1.Build; $OS = $Line1.OperatingSystem; $MajorBuilds += $BldNo;
                IF ($MjBld -eq $BldNo) {Break;}
                ELSE {$OS = "Unknown";}
            }
            $PatchDetails += [PSCustomObject] @{OperatingSystem = $OS; Build = $BLD; MajorBuild = $MjBld; MinorBuild = $MnBld; PatchID = $PatchID; ReleaseDate = $ReleaseDate;}
       }
}
$MajorBuilds = $MajorBuilds | Select-Object -Unique | Sort-Object -Descending;
$PatchDetails = $PatchDetails | Select-Object OperatingSystem, Build, MajorBuild, MinorBuild, PatchID, ReleaseDate -Unique | Sort-Object MajorBuild,PatchID -Descending | Select-Object -First 1
$PatchTuedsay = $PatchDetails.ReleaseDate

#######################  User Input Section ##################################################################################
$WorkingFolder = "C:\TEMP\IntunePatchingReport" 
$ReportPath    = "C:\TEMP\IntunePatchingReport\IntunePatchingReport.csv"
$From          = "xyz@abc.com"
$To            = "123@abc.com"
$CC            = "456@abc.com"
$Subject       = "Intune Patching- Windows Patching Compliance Report against $PatchTuedsay (Patch Tuesday)"
$SmtpServer    = "smtpserver"
$Port          = '25'
$Priority      = "Normal"
$ProactiveDetectionPolicyGUID = "31b42771-771f-4753-8dd5-534343a8ab6"
    
#######################################################################################################################
$Date = Get-Date -Format "MMMMMMMM dd, yyyy";
$PatchingMonth = "";
$PatchReleaseDays = 0;
# Create an empty array of PSObject objects
$buildInfoArray = @()
#====================================================================================
#Creating working Folder
New-Item -ItemType Directory -Path $WorkingFolder -Force  -InformationAction SilentlyContinue

#----------------------------------------------------------------------------------------------------------------------------------------
$ReportName = "Patching Report"
$Path ="$WorkingFolder\PD_Dump\"


$MGIModule = Get-module -Name "Microsoft.Graph.Intune" -ListAvailable -InformationAction SilentlyContinue
Write-Host "Checking Microsoft.Graph.Intune is Installed or Not"

    If ($MGIModule -eq $null) 
    {
        Write-Host "Microsoft.Graph.Intune module is not Installed"
        Write-Host "Installing Microsoft.Graph.Intune module"
        Install-Module -Name Microsoft.Graph.Intune -Force
        Write-Host "Importing Microsoft.Graph.Intune module"
        Import-Module Microsoft.Graph.Intune -Force -InformationAction SilentlyContinue
    }

    ELSE 
    {   Write-Host "Microsoft.Graph.Intune is Installed"
        Write-Host "Importing Microsoft.Graph.Intune module"
        Import-Module Microsoft.Graph.Intune -Force -InformationAction SilentlyContinue
    }
$tenant = “Lab.onmicrosoft.com”
$authority = “https://login.windows.net/$tenant”
$clientId = “b71d5aa6-6c38-4f965585858e-04516bb58180”
$clientSecret = "oMG8Q~WOW8AIXMjSzQBFG67676xI.P1sUQ6cRcME"
Update-MSGraphEnvironment -AppId $clientId -Quiet -InformationAction SilentlyContinue
Update-MSGraphEnvironment -AuthUrl $authority -Quiet -InformationAction SilentlyContinue
Connect-MSGraph -ClientSecret $ClientSecret -InformationAction SilentlyContinue
Update-MSGraphEnvironment -SchemaVersion "Beta" -Quiet -InformationAction SilentlyContinue



#============Create Request Body==========================================================================================================================================================
$postBody = @{

 'reportName' = "DeviceRunStatesByProactiveRemediation"
 'filter' = "PolicyId eq '$ProactiveDetectionPolicyGUID'"
 
 }
#=========== MakeRequest ==================================================================================================================================================================


$exportJob = Invoke-MSGraphRequest -HttpMethod POST -Url "DeviceManagement/reports/exportJobs" -Content $postBody

Write-Host "Export Job initiated for $ReportName Report "


#====================================Checking Report Ready status==========================================================================================================================
do{ 
$exportJob = Invoke-MSGraphRequest -HttpMethod Get -Url "DeviceManagement/reports/exportJobs('$($exportJob.id)')" -InformationAction SilentlyContinue
    Start-sleep -second 2
    Write-Host -NoNewline '...........'

  } while ($exportJob.status -eq 'inprogress')

  Write-Host 'Report is in Ready(Completed) status for Downloading' -ForegroundColor Yellow

  If ($exportJob.status -eq 'completed') 
  { $fileName = (Split-path -Path $exportJob.url -Leaf).split('?')[0]
  Write-host "Export Job completed.......  Writing File $fileName to Disk........" -ForegroundColor Yellow
  Invoke-WebRequest -Uri $exportJob.url -Method Get -OutFile $fileName
  Remove-Item –path $path* -include *.csv
  Expand-Archive -Path $fileName -DestinationPath $Path 
  $FileName = Get-ChildItem -Path $Path* -Include *.csv | Where {! $_.PSIsContainer } | Select Name,FullName
  $FileName.fullName
  $ReportDump = Import-csv -path $FileName.fullName
  #Out-Host -InputObject $ReportDump

  $ReportD = $ReportDump 
  $ReportD.count
     
  } 
  ======================================
  $Compliant = 0
  $Non_Compliant = 0
  $NeedToCheck = 0
  $PatchingReportInfo = @()
foreach($Device in $ReportD )
{ 
  $PatchingReportSProps = [ordered] @{
        
        Device_Name = $Device.DeviceName
        User_Name = $Device.UserName
        UPN = $Device.UPN
        OS_Version = $Device.OSVersion
        Model =$Device.Model
        Join_Type =$Device.JoinType
        PreRemediationDetectionScriptOutput = $Device.PreRemediationDetectionScriptOutput
        SplitPreRemediationDetectionScriptOutput = $parts = $Device.PreRemediationDetectionScriptOutput.Split(";")
        Patching_Status  =  $parts[0]
        Latest_Patches_RD = $parts[1]
        Installed_Patch= $parts[2]
        Installed_Patch_RD = $parts[3] 
        Last_Successfull_ScanDate= $parts[4] 
        LastScan_from_CurrentDate_InDays= $parts[5] 
        Last_BootUp_date = $parts[6]
        C_Drive_freeSpace_GB =     $parts[7] 
        Windows_Update_Service= $parts[8] 
        Update_Orchestrator_Service= $parts[9] 
        Microsoft_Update_Health_Service = $parts[10] 
        Microsoft_Account_Signin_Assistant_Service= $parts[11] 
        Windows_SKU= $parts[12] 
        OSVersion = $parts[13] 
        OS_EOL_Status = $parts[14] 
        Reboot_Pending = $parts[15] 
        Make = $parts[16] 
        Serial_Number = $parts[17] 
        }   
    
    if ($parts[0] -eq "Compliant")
    {$Compliant++}


elseif ($parts[0] -eq "Non Compliant")
    {$Non_Compliant++}

    elseif ($parts[0] -like "Failed*")
    {$NeedToCheck++}

     else{}
         
  $PatchingReportobject = New-Object -Type PSObject -Property $PatchingReportSProps
  $PatchingReportInfo +=$PatchingReportobject
 }
 $FinalReport = $PatchingReportInfo | Select-Object -Property Device_Name,User_Name,UPN,OS_Version,Serial_Number,Windows_SKU,OS_Version, Make,Model,Join_type,Patching_Status,Latest_Patches_RD,Installed_Patch,Installed_Patch_RD,Last_Successfull_ScanDate,LastScan_from_CurrentDate_InDays,Last_BootUp_date,C_Drive_freeSpace_GB,Windows_Update_Service,Update_Orchestrator_Service,Microsoft_Update_Health_Service,Microsoft_Account_Signin_Assistant_Service,OS_EOL_Status,Reboot_Pending
 $FinalReport | Export-Csv -Path $ReportPath  -NoTypeInformation
 $Total = $FinalReport.Count
 $Compliance_Round = ($Compliant / $Total ) * 100 
 $Compliance = [math]::Round($Compliance_Round,2)  # Round to Specific Decimal Place
 

#-----------------------------------------------------------------------------------------------------------------------------------------------------
$EmailBody1 = @" 

<p>Hello All</p>
 <p></p>
 <p>Please Find Windows 10/11 Patching compliance Report against $PatchTuedsay (Patch Tuesday).</p>

<head>
	<style> 	table, th,
                td {border: 3px solid black;}
	</style>
</head>

<table style="width: 68%" style="border-collapse: collapse; border: 1px solid #008080;">

 <tr>
    <td colspan="2" bgcolor="#71B2EE" style="background-color:Tan; font-size: large; height: 35px;">
        <b>Windows Patching Compliance Dashboard</b>   
    </td>
 </tr>


 <!----For Total Devices-------------------------------------------------------------------------------------->
<tr <tr style="background-color:lightgrey">
    <td style="width: 201px; height: 35px">&nbsp;Total Devices</td>
    <th style="height: 35px; width: 233px;">
    <b>VarTotal</b></td>
 </tr>
 
  <! --- For Compliance Devcies -------------------------------------------------------------------------------->
 <tr <tr style="background-color:MediumAquaMarine">
    <td style="width: 201px; height: 35px">&nbsp;Compliant Devices</td>
    <th style="height: 35px; width: 233px;">
    <b>Varsuccess</b></td>
 </tr>

<!----For Non_Compliant Devcies-------------------------------------------------------------------------------------->
<tr <tr style="background-color:Salmon">
    <td style="width: 201px; height: 35px">&nbsp;Non Compliant Devices</td>
    <th style="height: 35px; width: 233px;">
    <b>Varfailure</b></td>
 </tr>

<!----For NeedToCheck Devcies-------------------------------------------------------------------------------------->
<tr <tr style="background-color:LightGoldenrodYellow">
    <td style="width: 201px; height: 35px">&nbsp;Need To Check Device</td>
    <th style="height: 35px; width: 233px;">
    <b>VarNeedToCheck</b></td>
 </tr>


  <!----For Over All Compliance---------------------------------------------------------------------------------------->
<tr <tr style="background-color:lightgreen">
    <td style="width: 201px; height: 35px">&nbsp;<b>Compliance(%)</b></td>
    <th style="height: 35px; width: 233px;"> 
    <b>VarCompliance%</b></th>
 </tr>


</table>
<p>Regards</p>
<p>Patch Management Team</p>
 
"@
$EmailBody1= $EmailBody1.Replace("VarTotal",$Total)
$EmailBody1= $EmailBody1.Replace("Varsuccess",$Compliant)
$EmailBody1= $EmailBody1.Replace("Varfailure",$Non_Compliant)
$EmailBody1= $EmailBody1.Replace("VarCompliance",$Compliance)
$EmailBody1= $EmailBody1.Replace("VarNeedToCheck",$NeedToCheck)

  #___________________________________________________________________________________________________________________________________________________________
  #___________________________________________________________________________________________________________________________________________________________
### SENDING EMAIL ##########################################################################################################################

 Write-Host "Sending Mail to $to,$cc" -ForegroundColor yellow
#Email Params
$Parameters = @{
    From        = $from
    To          = $To
    Subject     = $Subject 
    Body        = $EmailBody1
    BodyAsHTML  = $True
    CC          = $CC
    Port        = $Port
    Priority    = $Priority 
    SmtpServer  = $SmtpServer
    Attachments = $ReportPath
}

#Sending email
Send-MailMessage @Parameters 
Write-Host "Mail successfully sent to $to,$cc" -ForegroundColor Green





