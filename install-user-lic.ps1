<#PSScriptInfo
.Version 1.3
.AUTHOR Deklan van de Laarschot
.COMPANYNAME Next2IT
.COPYRIGHT 2021 Next2IT
.TAGS O365
.LICENSEURI
.ICONURI
.EXTERNALMODULEDEPENDENCIES 
.REQUIREDSCRIPTS
.EXTERNALSCRIPTDEPENDENCIES
.RELEASENOTES
.PRIVATEDATA
#>

<# 
.DESCRIPTION 
 Installs the Office 365 suite for Windows using the Office Deployment Tool 
 Added email notifications on failures.
#> 

[CmdletBinding(DefaultParameterSetName = 'XMLFile')]
Param(
  [Parameter(ParameterSetName = "XMLFile")][ValidateNotNullOrEmpty()][String]$ConfiguratonXMLFile,
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("TRUE","FALSE")]$AcceptEULA = "TRUE",
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("Broad","Targeted","Monthly")]$Channel = "Broad",
  [Parameter(ParameterSetName = "NoXML")][Switch]$DisplayInstall = $False,
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("Groove","Outlook","OneNote","Access","OneDrive","Publisher","Word","Excel","PowerPoint","Teams","Lync")][Array]$ExcludeApps,
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("64","32")]$OfficeArch = "64",
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("O365ProPlusRetail","O365BusinessRetail")]$OfficeEdition = "O365ProPlusRetail",
  [Parameter(ParameterSetName = "NoXML")][ValidateSet(0,1)]$SharedComputerLicensing = "0",
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("TRUE","FALSE")]$EnableUpdates = "TRUE",
  [Parameter(ParameterSetName = "NoXML")][String]$LoggingPath,
  [Parameter(ParameterSetName = "NoXML")][String]$SourcePath,
  [Parameter(ParameterSetName = "NoXML")][ValidateSet("TRUE","FALSE")]$PinItemsToTaskbar = "TRUE",
  [Parameter(ParameterSetName = "NoXML")][Switch]$KeepMSI = $False,
  [String]$OfficeInstallDownloadPath = "C:\Office\Office365Install"
)

# Set external breakout IP address

$externalip = "8.8.8.8"

Function Send-Mail ($message) {
  #$emailSmtpServer = "mailserver"
  #$emailFrom = "office365@next2it.co.uk"
  #$emailTo = "support@next2it.co.uk"
  #$emailSubject = "Office 365 Installation Failed"
  #Send-MailMessage -To $emailTo -From $emailFrom -Subject $emailSubject -Body "Install failed on $env:computername, the reason for failure was $message" -SmtpServer $emailSmtpServer
  Write-Warning "Email disabled."
}

Function Generate-XMLFile{

  If($ExcludeApps){
    $ExcludeApps | ForEach-Object{
      $ExcludeAppsString += "<ExcludeApp ID =`"$_`" />"
    }
  }

  If($OfficeArch){
    $OfficeArchString = "`"$OfficeArch`""
  }

  If($KeepMSI){
    $RemoveMSIString = $Null
  }Else{
    $RemoveMSIString =  "<RemoveMSI />"
  }

  If($Channel){
    $ChannelString = "Channel=`"$Channel`""
  }Else{
    $ChannelString = $Null
  }

  If((Resolve-DnsName -Name myip.opendns.com -Server 208.67.222.220).IPAddress -eq $externalip){
    $SourcePathString = "SourcePath=`"\\server\files`"" 
    Write-Warning "MPLS Detected......"
  }Else{
    $SourcePathString = $Null
    Write-Warning "Direct Internet...."
  }


  If($DisplayInstall){
    $SilentInstallString = "Full"
  }Else{
    $SilentInstallString = "None"
  }

  If($LoggingPath){
    $LoggingString = "<Logging Level=`"Standard`" Path=`"$LoggingPath`" />"
  }Else{
    $LoggingString = $Null
  }
  #XML data that will be used for the download/install

  $OfficeXML = [XML]@"
      <Configuration>
        <Add OfficeClientEdition=$OfficeArchString $ChannelString $SourcePathString  >
          <Product ID="$OfficeEdition">
            <Language ID="MatchOS" />
            $ExcludeAppsString
          </Product>
        </Add>  
        <Property Name="PinIconsToTaskbar" Value="$PinItemsToTaskbar" />
        <Property Name="SharedComputerLicensing" Value="$SharedComputerlicensing" />
        <Display Level="$SilentInstallString" AcceptEULA="$AcceptEULA" />
        <Updates Enabled="$EnableUpdates" />
        $RemoveMSIString
        $LoggingString
      </Configuration>
"@
  
  #Save the XML file
  $OfficeXML.Save("$OfficeInstallDownloadPath\OfficeInstall.xml")
  Return "$OfficeInstallDownloadPath\OfficeInstall.xml"
}

Function Test-URL{
  Param(
	$CurrentURL
  )

  Try{
    $HTTPRequest = [System.Net.WebRequest]::Create($CurrentURL)
    $HTTPResponse = $HTTPRequest.GetResponse()
    $HTTPStatus = [Int]$HTTPResponse.StatusCode

    If($HTTPStatus -ne 200) {
      Return $False
    }

    $HTTPResponse.Close()

  }Catch{
	  Return $False
  }	
  Return $True
}


$VerbosePreference = "Continue"
$ErrorActionPreference = "Continue"

try {
  $CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
  If(!($CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))){
      Write-Warning "Script is not running as Administrator"
      Write-Warning "Please rerun this script as Administrator."
      Send-Mail "Not rans as admin" 
      Exit
  }
  
  
  If(!($ConfiguratonXMLFile)){ #If the user didn't specify with -ConfigurationXMLFile param, we make one!
    $ConfiguratonXMLFile = Generate-XMLFile
  }Else{
    If(!(Test-Path $ConfiguratonXMLFile)){
      Write-Warning "The configuration XML file is not a valid file"
      Write-Warning "Please check the path and try again"
      Send-Mail "Invalied XML File"
      Exit
    }
  }

  #Run the O365 install
  Try{
    Write-Verbose "Downloading and installing Office 365"
    $OfficeInstall = Start-Process "$OfficeInstallDownloadPath\Setup.exe" -ArgumentList "/configure $ConfiguratonXMLFile" -Wait -PassThru
  }Catch{
    Write-Warning "Error running the Office install. The error is below:"
    Write-Warning $_
    Send-Mail "Error running Office install"
  } 
}
catch {
  Send-Mail "Script failed!" 
}

#Check if Office 365 suite was installed correctly.

$RegLocations = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
                 )

$OfficeInstalled = $False
Foreach ($Key in (Get-ChildItem $RegLocations) ) {
  If($Key.GetValue("DisplayName") -like "*Microsoft 365 Apps*") {
    $OfficeVersionInstalled = $Key.GetValue("DisplayName")
    $OfficeInstalled = $True
  }
}

If($OfficeInstalled){
  Write-Verbose "$($OfficeVersionInstalled) installed successfully!"
}Else{
  Write-Warning "Office 365 was not detected after the install ran"
  Send-Mail "Error Office not detected after installation!"
}

#Create Desktop Icons

Copy-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk" -Destination $env:Public\Desktop\ -Force
Copy-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" -Destination $env:Public\Desktop\ -Force
Copy-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk" -Destination $env:Public\Desktop\ -Force
Copy-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk" -Destination $env:Public\Desktop\ -Force
