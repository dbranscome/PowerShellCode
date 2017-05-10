#requires -Version 3.0 -Modules MSOnline, SkypeOnlineConnector

<#
    .SYNOPSIS
    Enable Skype for Business online users for any voice service(s)
    .DESCRIPTION
    Enable Skype for Business online users and/or migrate users to Skype for Business online.  Configure any voice services including (Cloud PSTN Conferencing, Cloud PBX, or Cloud PSTN Calling) using either E5 or specific service add-ons.
    .LINK
    http://www.skypeoperationsframework.com
    .EXAMPLE
    Enable-SkypeForBusinessVoiceUsers.ps1 -FileName .\UsersToEnable.csv -o365AdminUserName user@domain.onmicrosoft.com

    This is an example for running the script to enable online users
    .EXAMPLE
    Enable-SkypeForBusinessVoiceUsers.ps1 -FileName .\UsersToEnable.csv -o365AdminUserName user@domain.onmicrosoft.com -OnPremisesAdminUserName user@domain.com -OnPremisesFqdn PoolWebFQDN.domain.com 

    This is an example for running the script when users will be migrated from on-premises servers
    .PARAMETER FileName
    The path to the CSV input file name for users to process. This file must end in a .csv file extension.
    .PARAMETER o365AdminUserName
    Administrative user principal name that will be used to connect to Office 365, enable users, and perform configuration steps.
    .PARAMETER OnPremisesAdminUserName
    Administrative user principal name that will be used to connect to the on-premises pool to migrate users to Skype for Business online.
    .PARAMETER OnPremisesFQDN
    The internal web services fully quallified domain name for an on-premises Lync 2013 or Skype for Business 2015 pool.  This will be used to connect to PowerShell remoting facilitating user migrations.
    .PARAMETER PauseForProvisioning
    TRUE/FALSE value to indicate if the script should pause and wait for users to be provisioned before attempting to migrate and/or configure services. If users are not completely provisioned, errors are expected during service configuration.   If you choose to not pause for provisioning, you can re-run the script to complete configuration after provisioing has occured.  
    .PARAMETER ResultsFileName
    The file path and name for capturing script results.  This file name must end in a .csv file extension.
    .PARAMETER PauseForProvisioning
    The file path and name for capturing script verbose logging. 
    .OUTPUTS
    The script will return a table listing the users and script execution results.  These results will also be saved to a results.csv file in the local directory unless another filename and path is specified.
    .NOTES
    © 2016 Microsoft Corporation.  All rights reserved.  This document is provided 
    "as-is." Information and views expressed in this document, including URL and 
    other Internet Web site references, may change without notice.  

    This document does not provide you with any legal rights to any intellectual 
    property in any Microsoft product. Skype for Business customers and partners 
    may copy, use and share these materials for planning, deployment and operation 
    of Skype for Business.  

    Skype Operations Framework  
    The shift to the cloud requires rethinking how enterprises and partners Plan, 
    Deliver and Operate Skype for Business. The Skype Operations Framework (SOF) 
    provides a multi-faceted approach to the successful deployment of Skype for 
    Business, providing: 
    -Practical guidance, recommended practices, tools, and assets to enable 
    enterprises to plan, deliver and operate a reliable and cost effective 
    Skype for Business Service in the cloud   
    -A common understanding of the Skype for Business online lifecycle for 
    customers and partners to effectively engage and drive Skype usage and 
    customer success  
    -Training for customers and partners via Skype Academy  
    -Feedback mechanisms to capture and incorporate updates from the field 

    To find out more please visit www.skypeoperationsframework.com   
    We want to hear from you about how you are using the tools and assets, 
    what works and what does not.  If you feel there is anything missing, or any 
    other feedback that you would like to provide, please go to 
    www.skypefeedback.com to provide your feedback. 
#>    
param(
  #Input CSV file name and path
  [Parameter(Mandatory=$true, HelpMessage='Enter the path and filename for the CSV file to be processed')]
  [ValidatePattern('.\.csv$')]  
  [ValidateScript({Test-Path -Path $_})]
  [string] $FileName,
  
  #Office 365 administrative user principal name that will be used for assigning licenses and configuring users
  [Parameter(Mandatory=$true, HelpMessage='Enter the Office 365 admin user principal name to be used for assigning licenses and configuring users')]
  [ValidatePattern('.@\w+\.\w+')]
  [string] $o365AdminUserName,
  
  #On-premises administrative user name that will be used for migrating users  
  #[ValidatePattern('.@\w+\.\w+')]
  [AllowNull()]
  [string] $OnPremisesAdminUserName,
  
  #On-premises internal web servcies FQDN with will be used to connect PowerShell remoting
  #[ValidatePattern('\w+\.\w+')]
  [AllowNull()]
  [string] $OnPremisesFqdn,
  
  #Input to tell script if it should pause to wait for license provisioning
  [bool] $PauseForProvisioning = $true,
  
  #File name for verbose output
  [string] $VerboseLogFileName = 'VerboseOutput.Log',
  
  #File name for results file
  [ValidatePattern('.csv$')]
  [string] $ResultsFileName = 'Results.csv'  
)

#Define script vesion for reporting and troubleshooting
$EnableSkypeForBusinessVoiceUsersScriptVersion = 'V4.00'


#region Script Variable Definitions

#Set the error action preference to 'stop'.  This will allow the script to catch errors from the remote PSSession.
$script:ErrorActionPreference = 'Stop' 

#Set the warning preference to not write to the script output
$script:WarningPreference = 'SilentlyContinue' 

#Define contstant Variables
New-Variable -Name LicName_SkypeforBusiness -Value 'MCOSTANDARD' -Option ReadOnly
New-Variable -Name LicName_Exchange -Value 'EXCHANGE_S_ENTERPRISE' -Option ReadOnly

New-Variable -Name LicName_PSTNConferencing -Value 'MCOMEETADV' -Option ReadOnly
New-Variable -Name LicName_CloudPBX -Value 'MCOEV' -Option ReadOnly
New-Variable -Name LicName_PSTNLocal -Value 'MCOPSTN1' -Option ReadOnly
New-Variable -Name LicName_PSTNLocalAndInternational -Value 'MCOPSTN2' -Option ReadOnly
New-Variable -Name LicName_E3 -Value 'ENTERPRISEPACK' -Option ReadOnly
New-Variable -Name LicName_E5 -Value 'ENTERPRISEPREMIUM' -Option ReadOnly


New-Variable -Name LicenseEnablement_Success -Value 'Enabled' -Option ReadOnly
New-Variable -Name LicenseEnablement_Failure -Value 'License enablement failed' -Option ReadOnly
$SuccessfulResultCollection_LicenseEnablement = @()
$SuccessfulResultCollection_LicenseEnablement += $LicenseEnablement_Success


New-Variable -Name Migration_Success -Value 'Migrated' -Option ReadOnly
New-Variable -Name Migration_NotAttempted_AlreadyOnline -Value 'Not attempted - User already homed online' -Option ReadOnly
New-Variable -Name Migration_NotAttempted_NotEnabled -Value 'Not attempted - User not enabled on-premises' -Option ReadOnly
New-Variable -Name Migration_NotAttempted -Value 'Not attempted.' -Option ReadOnly
New-Variable -Name Migration_Failed -Value 'Migration failed' -Option ReadOnly
New-Variable -Name Migration_Skipped -Value 'Skipped - User not designated for migration' -Option ReadOnly
$SuccessfulResultCollection_Migration = @()
$SuccessfulResultCollection_Migration += $Migration_Success
$SuccessfulResultCollection_Migration += $Migration_NotAttempted_AlreadyOnline
$SuccessfulResultCollection_Migration += $Migration_Skipped

New-Variable -Name Configuration_Success -Value 'Configured' -Option ReadOnly
New-Variable -Name Configuration_NotRequired -Value 'Not required' -Option ReadOnly
New-Variable -Name Configuration_NotAttempted -Value 'Not attempted' -Option ReadOnly
New-Variable -Name Configuration_Failure -Value 'Configuration failed' -Option ReadOnly
$SuccessfulResultCollection_Configuration = @()
$SuccessfulResultCollection_Configuration += $Configuration_Success

#endregion

#region Function Definitions

function Limit-UserReportingObject
{

  param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
    [object[]] $Input,
  
    [Parameter(Mandatory=$true,HelpMessage='User Principal Name')]
    [string] $UserPrincipalName
  )
  process
  {
    if ($Input.UPN -eq $UserPrincipalName)
    {
      $_
    }
  }
}
function Limit-PSTNCallingUsersWithNoPhoneAssignment
{
  process
  {
    if (($_.EnablePSTNCallingDomestic -eq 'TRUE' -or $_.EnablePSTNCallingDomesticAndInternational -eq 'TRUE') -and [string]::IsNullOrEmpty($_.PhoneNumberToAssign) -and [string]::IsNullOrEmpty($_.PhoneNumberCityCode))
    {
      $_
    }
  }
}

function Limit-PSTNCallingUsersWithNoEmergencyDescription
{
  process
  {
    if (($_.EnablePSTNCallingDomestic -eq 'TRUE' -or $_.EnablePSTNCallingDomesticAndInternational -eq 'TRUE') -and [string]::IsNullOrEmpty($_.EmergencyAddressDescription))
    {
      $_
    }
  }
}

function Limit-PSTNCallingMutuallyExclusiveLicenses
{
  process
  {
    if ($_.EnablePSTNCallingDomestic -eq 'TRUE' -and $_.EnablePSTNCallingDomesticAndInternational -eq 'TRUE')
    {
      $_
    }
  }
}

function Limit-EmptyUsageLocation
{
  process
  {
    if ([string]::IsNullOrEmpty($_.UsageLocation))
    {
      $_
    }
  }
}

function Limit-UsersToMigrate
{
  process
  {
    if ($_.MigrateToSkypeForBusinessOnline -eq 'TRUE')
    {
      $_
    }
  }
}

function Write-VerboseLog{
  <#
      .SYNOPSIS
      Writes messages to Verbose output and Verbose log file
      .DESCRIPTION
      This fuction will direct verbose output to the console if the -Verbose 
      paramerter is specified.  It will also write output to the Verboselog.
  #>
  param(
    [String]$Message  
  )

  $VerboseMessage = ('{0} Line:{1} {2}' -f (Get-Date), $MyInvocation.ScriptLineNumber, $Message)
  #OLD $VerboseMessage = "$(Get-Date):$($Invocation.MyCommand) $($Message)"
  Write-Verbose -Message $VerboseMessage
  Add-Content -Path $VerboseLogFileName -Value $VerboseMessage
}

function Write-ScriptError{
  param(
    [Parameter(Mandatory=$true,HelpMessage='Provide message to include in the log')]
    [string] $Message,
    
    [Parameter(Mandatory=$true,HelpMessage='Error object to report')]
    [Object]$ErrorObject,

    [bool] $Terminating = $false
  )
  
  Write-VerboseLog -Message $Message
  Write-VerboseLog -Message ('Error occurred: {0}' -f $ErrorObject.Exception.Message)
  Write-VerboseLog -Message ('STACKTRACE:{0}' -f $ErrorObject.ScriptStackTrace)
  if ($Terminating){Write-Error -Message $Message -ErrorAction Stop}
}

function Get-LicenseSku
{
  param(
    [Parameter(Mandatory=$true,HelpMessage='Account name for the tenant')]
    [String] $AccountName,

    [Parameter(Mandatory=$true,HelpMessage='Name of license to assign')]
    [String] $LicenseName
  )

  return ('{0}:{1}' -f $AccountName, $LicenseName)
}

function Limit-AccountSku
{
  <#
      .SYNOPSIS
      Return the user license that matches the supplied account SKU.

      .DESCRIPTION
      Accepts a list of licenses assigned from a user and returns the license for the defined SKU 

      .PARAMETER LicensePackSku
      Name of license SKU to be returned.

      .EXAMPLE
      (Get-MSOLUser -UserPrincipalName user@domain.com).Licenses | Limit-AccountSku -LicensePackSku AccountName:ENTERPRISEPACK
      Returns the ENTERPRISEPACK license object assigned to that user.

      .NOTES
      This is useful when trying to determine the servcie status of licenses.


      .INPUTS
      An array of licenses assigned to a user object.  This is most likely the Licenses property returned from the Get-MSOLUser cmdlet.

      .OUTPUTS
      The license object matching the supplied account SKU.
  #>


  param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
    [object[]] $Input,
  
    [Parameter(Mandatory=$true,HelpMessage='License SKU')]
    [string] $LicensePackSku
  )
  process
  {
    if ($Input.AccountSkuId -eq $LicensePackSku)
    {
      $_
    }
  }
}

function Get-UpdatedLicenseOptions
{
  param
  (
    [Parameter(Mandatory=$true, Position=0, HelpMessage='Object if services and corresponding status')]
    [Object]$ServiceStatusList,

    [Parameter(Mandatory=$true,HelpMessage='Name of the license(s) to assign to the user (i.e. MCOSTANDARD, MCOEV, etc.)')]
    [String[]]$LicensesToAssign,
    
    [Parameter(Mandatory=$true,HelpMessage='License SKU')]
    [string] $LicensePackSku
            
  )
  
  #Define a new list of disabled services
  $DisabledServices = @()
  $ServiceStatusList.GetEnumerator() | ForEach-Object {
    if ($_.ProvisioningStatus -eq 'Disabled' -and !($LicensesToAssign -contains $_.ServicePlan.ServiceName)) {
      #Service disabled - adding to list
      $DisabledServices += ($_.ServicePlan.ServiceName.ToString())
    }
  }
  Write-VerboseLog -Message ('Updated service status list to apply: {0}' -f ($DisabledServices | Out-String))
  
  #Create license options to carry forward existing disabled services
  $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $LicensePackSku -DisabledPlans $DisabledServices
  
  return $LicenseOptions
}



function Enable-ServicesInLicensePack{
  <#
      .SYNOPSIS
      Enable Skype for Business services within a license pack (i.e. E3 or E5)
      .DESCRIPTION
      License packs contain multple services that can be enabled or disabled.  
      This function will ensure that the desired services are enabled in that 
      license pack.
  #>
  param(
    [Parameter(Mandatory=$true,HelpMessage='Get-MSOLUser output object continaing user Skype for Business configuration information')]
    [Object]$UserConfig,
     
    [Parameter(Mandatory=$true,HelpMessage='Name of the license pack to assign (i.e. ENTERPRISEPACK or ENTERPRISEPREMIUM')]
    [string]$LicensePackName,
    
    [Parameter(Mandatory=$true,HelpMessage='Name of the license(s) to assign to the user (i.e. MCOSTANDARD, MCOEV, etc.)')]
    [String[]]$LicensesToAssign,
     
    [Parameter(Mandatory=$true,HelpMessage='Tenant account name to be used when identifying licenses')]
    [string]$AccountName,
    
    [Parameter(Mandatory=$true,HelpMessage='License update type (New/Update)')]
    [string]$UpdateType         
  )
  

  Write-VerboseLog -Message ('Function Invocation: {0}' -f ($MyInvocation.BoundParameters | out-string)) 
  try{
    $LicensePackSkuName= Get-LicenseSku -AccountName $AccountName -LicenseName $LicensePackName
    $MSOLSku = Get-MsolAccountSku | Limit-AccountSku -LicensePackSku $LicensePackSkuName
		
    #Get services status for the desired license 
    switch ($UpdateType){
      'New' {$ServiceStatusList = (Get-MSOLAccountSku | Limit-AccountSku -LicensePackSku $LicensePackSkuName).ServiceStatus; break}
      'Update' {$ServiceStatusList = @(($UserConfig.Licenses | Limit-AccountSku -LicensePackSku $LicensePackSkuName).ServiceStatus); break}
    }
    
    Write-VerboseLog -Message ("Current license pack service status: `n{0}" -f ($ServiceStatusList | Out-String))
    Write-VerboseLog -Message ('Setting user license to enable service: {0}' -f $LicenseToAssign)
    
    $UpdatedLicenseOptions = Get-UpdatedLicenseOptions -ServiceStatusList $ServiceStatusList -LicensesToAssign $LicensesToAssign -LicensePackSku $LicensePackSkuName
    
    switch($UpdateType){
      'New' {
        if ($MSOLSku.ActiveUnits -gt $MSOLSku.ConsumedUnits) {
          Set-MsolUserLicense -UserPrincipalName $UserConfig.UserPrincipalName -AddLicenses $LicensePackSkuName -LicenseOptions $UpdatedLicenseOptions -ErrorAction Stop
        } else {
          Write-Error -Message 'Not enough licenses available to assign' -Exception 'The number of license active units is not greater than the number of consumed units' -ErrorAction stop
        }
        break
      }
      'Update' {Set-MsolUserLicense -UserPrincipalName $UserConfig.UserPrincipalName -LicenseOptions $UpdatedLicenseOptions -ErrorAction Stop; break}
    }
    #Did not catch any errors - operation successful
    return $LicenseEnablement_Success
  } catch {
    Write-ScriptError -Message ('{0}: {1}' -f $LicenseEnablement_Failure, $UserConfig.UserPrincipalName) -ErrorObject $_
    return ('{0}: {1}' -f $LicenseEnablement_Failure, $_.Exception.Message)
  }
}


function Wait-LicenseEnablement
{
  <#
      .SYNOPSIS
      Wait until licenses have been enabled
      .DESCRIPTION
      Check every 60 seconds for licenses to be in an acceptable state for further provisioning or configuration
      .EXAMPLE
      Wait-LicenseEnablement -UserPrincipalName user@domain.com -ServicesToCheck @('MCOSTANDARD','MCOEV')
  #>
  param
  (
    [Parameter(Mandatory=$true,HelpMessage='User principal name', Position=0)]
    [string]
    $UserPrincipalName,

    [Parameter(Position=1)]
    [string[]]
    $ServicesToCheck = @('MCOSTANDARD','MCOEV','MCOMEETADV','MCOPSTN1','MCOPSTN2','EXCHANGE_S_ENTERPRISE')
  )
  
  #Check Licensing status - define allowable values
  $UserReady = $false
  $LoopCount = 1
  $AllowableServiceStatus = @('Disabled', 'Success')
  
  #wait until the licensing provisioing is complete
  while (-not $UserReady){
    
    #Get User MSOL Config
    Write-VerboseLog -Message ('Getting MSOL user config for {0}' -f ($UserPrincipalName))
    $UserLicenseReadiness = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction Stop
    
    #Reset user readiness value to true as long as we know we have licesnses to evaluate.
    #If it doesn't pass the checks below, the value will be set back to FALSE
    if ($UserLicenseReadiness.Licenses.Count -gt 0) {$UserReady = $true} else {$UserReady = $false}
    
    #Check if licenses are in the allowable configuration
    foreach ($License in $UserLicenseReadiness.Licenses){

      Write-VerboseLog -Message ('Checking license service status for:{0}' -f ($license.ServiceStatus | Out-String))
      foreach ($Service in $License.ServiceStatus){
        #Check if the service is a value we want to check AND that the status for that service is an acceptable state
        if ($ServicesToCheck -contains $Service.ServicePlan.ServiceName -and $AllowableServiceStatus -notcontains $Service.ProvisioningStatus){ 
          #User not yet in acceptable service state
          Write-VerboseLog -Message ('{0} is not in a provisioned state. Will wait and loop again to check for provisioning completion.' -f $Service.ServicePlan.ServiceName) 
          $UserReady = $false 
        }
      }
    }
    
    if ($UserReady) {
      Write-VerboseLog -Message ('User provisioning completed for {0}' -f $user.UserPrincipalName)
    } else { 
      Write-VerboseLog -Message ('Licenses have been assigned to {0}, but are not yet fully provisioned.  Licenses must be fully provisioned before any further configuration can be completed. Pausing for 60 seconds before checking again. Attempt number: {1}' -f $user.UserPrincipalName, $LoopCount)
      Write-Host -ForegroundColor Yellow ('Licenses have been assigned to {0}, but are not yet fully provisioned.  Licenses must be fully provisioned before any further configuration can be completed. Pausing for 60 seconds before checking again. Attempt number: {1}' -f $user.UserPrincipalName, $LoopCount)
      $LoopCount++
      Start-Sleep -Seconds 60 
    }
  }
}



function Enable-AddOnLicense{
  <#
      .SYNOPSIS
      Enable the defined user for the PSTN Conferencing Add-on (MCOMEETADV)
      .DESCRIPTION
      Enable the defined user for the PSTN Conferencing Add-on (MCOMEETADV).  
      This function should be used for users that are not assigned an E5 license.
  #>
  param(
    [Parameter(Mandatory=$True,HelpMessage='MSOL user object')]
    [Object]$UserConfig,

    [Parameter(Mandatory=$true,HelpMessage='Account name for tenant')]
    [string]$AccountName,

    [Parameter(Mandatory=$true,HelpMessage='License name to assign')]
    [string]$LicenseName
  )
  Write-VerboseLog -Message ('Function Invocation: {0}' -f ($MyInvocation.BoundParameters | out-string))
  
  try{
    $LicenseSkuName = Get-LicenseSku -AccountName $AccountName -LicenseName $LicenseName
    $MSOLSku = Get-MsolAccountSku | Limit-AccountSku -LicensePackSku $LicenseSkuName
    
    #check if user is already licensed
    if ($UserConfig.Licenses.AccountSkuId -notcontains $LicenseSkuName){
      Write-VerboseLog -Message ('{0} does not have the {1} license assigned.  Assigning the license now.' -f $UserConfig.UserPrincipalName, $LicenseName)
      if ($MSOLSku.ActiveUnits -gt $MSOLSku.ConsumedUnits){
        Set-MsolUserLicense -UserPrincipalName $UserConfig.UserPrincipalName -AddLicenses $LicenseSkuName -ErrorAction Stop
      } else {
        Write-Error -Message 'Not enough licenses available to assign' -Exception 'The number of license active units is not greater than the number of consumed units' -ErrorAction stop
      }
    } else {
      Write-VerboseLog -Message ('{0} already has the {1} license assigned.' -f $UserConfig.UserPrincipalName, $LicenseName)
    }
    return $LicenseEnablement_Success
  } catch {
    Write-ScriptError -Message ('{0}: {1}' -f $LicenseEnablement_Failure, $UserConfig.UserPrincipalName) -ErrorObject $_
    return ('{0}: {1}' -f $LicenseEnablement_Failure, $_.Exception.Message)
  }
}


function Set-PSTNCallingPhoneNumber{
  param(
    [Parameter(Mandatory=$True,HelpMessage='MSOL user config')]
    [Object]$UserConfig,

    [Parameter(Mandatory=$true,HelpMessage='Phone number to assign')]
    [string]$PhoneNumber,

    [Parameter(Mandatory=$true,HelpMessage='Location ID of the emergency address location')]
    [guid] $LocationId
  ) 

  Write-VerboseLog -Message ('Function Invocation: {0}' -f ($MyInvocation.BoundParameters | out-string))
  
  try{
    Write-VerboseLog -Message ('Configuring {0} with PSTN calling telephone number {1} and LocationId {2}.' -f $UserConfig.UserPrincipalName, $PhoneNumber, $LocationId)
    Set-CsOnlineVoiceUser -Identity $UserConfig.UserPrincipalName -TelephoneNumber $PhoneNumber -LocationID $LocationId -ErrorAction Stop

    #Removing enablement of CsUMMailbox since it is now built-in to the enablement process
    #Write-VerboseLog -Message ('Enabling CsOnlineUmMailbox for {0}.' -f $UserConfig.UserPrincipalName)
    #$null = Enable-CsOnlineUMMailBox -Identity $UserConfig.UserPrincipalName -LineUri $PhoneNumber -ErrorAction Stop
    return $Configuration_Success
  } catch {
    Write-ScriptError -Message ('{0}: {1}' -f $Configuration_Failure, $UserConfig.UserPrincipalName) -ErrorObject $_
    return ('{0}: {1}' -f $Configuration_Failure, $_.Exception.Message)
  }
}

function Get-AccountName{
  Write-VerboseLog -Message ('Function Invocation: {0}' -f ($MyInvocation.BoundParameters | out-string))
  return (Get-MSOLAccountSKU)[0].AccountName
}

function Grant-TenantDialplan{
  param(
    [Parameter(Mandatory=$True,HelpMessage='MSOL user config')]
    [Object]$UserConfig,

    [Parameter(Mandatory=$true,HelpMessage='Tenant Dial Plan to assign')]
    [string]$DialPlan
  )

  Write-VerboseLog -Message ('Function Invocation: {0}' -f ($MyInvocation.BoundParameters | out-string))
  try{
    # Tenant dial plan exists, assign to user
    Write-VerboseLog -Message ('Configuring {0} with Tenant dial plan {1}.' -f $UserConfig.UserPrincipalName, $DialPlan)
    Grant-CsTenantDialPlan -Identity $UserConfig.UserPrincipalName -PolicyName $DialPlan -ErrorAction Stop
    return $Configuration_Success

  } catch {
    Write-ScriptError -Message ('{0}: {1}' -f $Configuration_Failure, $UserConfig.UserPrincipalName) -ErrorObject $_
    return ('{0}: {1}' -f $Configuration_Failure, $_.Exception.Message)
  }

}

#endregion

#region Script Initial Setup

#create new file for verbose logging
$null = New-Item -Path $VerboseLogFileName -ItemType File -Force
Write-VerboseLog -Message 'Script started'

#Document script version to be used during troubleshooting
Write-VerboseLog -Message ('Script version: {0}' -f ($EnableSkypeForBusinessVoiceUsersScriptVersion))
Write-Host ('Script version: {0}' -f ($EnableSkypeForBusinessVoiceUsersScriptVersion))

#Setup reporting
$ResultsReporting = @()

#endregion

#region CSV Input and Validations

#Get CSV values
try{
  $UsersToProcess = @(Import-Csv -Path $FileName)
}
catch
{
  Write-ScriptError -Message ('Failed to import CSV file {0}. Exiting Script...' -f ($FileName)) -ErrorObject $_ -Terminating $true
}


#Documenat input CSV file to log
Write-VerboseLog -Message ($UsersToProcess | Out-String)


#Validate input file assumptions
$UsersToMigrate = @($UsersToProcess | Limit-UsersToMigrate)
$UsersWithNoLocation = @($UsersToProcess | Limit-EmptyUsageLocation)
$UsersWithPSTNCallingDomesticAndInternational = @($UsersToProcess | Limit-PSTNCallingMutuallyExclusiveLicenses)
$UsersWithPSTNCallingAndNoEmergencyLocation = @($UsersToProcess | Limit-PSTNCallingUsersWithNoEmergencyDescription)
$UsersWthPSTNCallingNoPhoneNumberNoCityCode = @($UsersToProcess | Limit-PSTNCallingUsersWithNoPhoneAssignment)
$DistinctUPNValues = @($UsersToProcess.UserPrincipalName | Sort-Object | Get-Unique)

#The following statements validate the CSV input file for terminating errors

#Check number of user to migrate
if ($UsersToMigrate.count -gt 0) {
  Write-VerboseLog -Message ('Found {0} users disignated for migration.' -f $UsersToMigrate.count)
  $ConnectToOnPremisesServices = $true	
} else { $ConnectToOnPremisesServices = $false }

#Check users that don't have any location
if ($UsersWithNoLocation.count -ne 0) {
  Write-VerboseLog -Message ("Found user(s) that don't have a location defined.  {0}" -f ($UsersWithNoLocation | out-string))
  Write-Host 'Input file validation error: Some user(s) do not have a usage location defined.'
  exit
}

#Check for users that are specifiec to enable both PSTN calling domestic and PSTN calling international
if ($UsersWithPSTNCallingDomesticAndInternational.count -ne 0) {
  Write-VerboseLog -Message ('Found user(s) that have PSTN Calling Domestic and International defined.  {0}' -f ($UsersWithPSTNCallingDomesticAndInternational | out-string))
  Write-Host 'Input file validation error: Some user(s) have both PSTN Calling Domestic and International set to TRUE.'
  exit
}

#Check for users that are specificed for PSTN calling but do not have an emergency location specified
if ($UsersWithPSTNCallingAndNoEmergencyLocation.count -ne 0) {
  Write-VerboseLog -Message ('Found user(s) that have PSTN Calling defined with no Emergency locaiton.  {0}' -f ($UsersWithPSTNCallingAndNoEmergencyLocation | out-string))
  Write-Host 'Input file validation error: Some PSTN Calling user(s) do not have a EmergencyAddressDescription defined.'
  exit
}

#Check for users that will be assigned PSTN calling services but don't have a phone number nor city code specified
if ($UsersWthPSTNCallingNoPhoneNumberNoCityCode.count -ne 0) {
  Write-VerboseLog -Message ('Found user(s) that have PSTN Calling defined with no phone number to assign and no phone number city code.  {0}' -f ($UsersWthPSTNCallingNoPhoneNumberNoCityCode | out-string))
  Write-Host 'Input file validation error: Some PSTN Calling user(s) do not have a PhoneNumberToAssign nor a PhoneNumberCityCode defined.'
  exit
}

#Check that there are no users listed twice in the input file
if ($UsersToProcess.count -ne $DistinctUPNValues.count) {
  Write-VerboseLog -Message 'The number of distinct UserPrincipalNames does not match the number of lines in the CSV file.'
  Write-Host 'Input file validation error: UserPrinicpalNames in the CSV file are not distinct.  UserPrincipalNames cannot be repeated in the same input file.'
  exit
}


#endregion

#region Credentials prompts and required environment setup
try{
  $O365Credentials = Get-Credential -UserName $O365AdminUserName -Message 'Provide admin credentials for Office 365 user management' -ErrorAction Stop

  if ($ConnectToOnPremisesServices){
    $OnPremisesCredentials = Get-Credential -UserName $OnPremisesAdminUserName -Message 'Provide admin credentials for on-premesis user migration' -ErrorAction Stop
  }

  #Validate on-premises pool fqdn
  if ($ConnectToOnPremisesServices -and [string]::IsNullOrEmpty($OnPremisesFqdn)){
    $OnPremisesFqdn = Read-Host -Prompt 'Please provice the internal web services fully quallified domain name for an on-premises Lync 2013 or Skype for Business 2015 pool migrate users'
  }

  #Connect to MSOL
  Write-VerboseLog -Message 'Connecting to MsolService'
  Connect-MsolService -Credential $O365Credentials -ErrorAction Stop

  Write-VerboseLog -Message 'Connecting to CsOnlineSession'
  $OnlineSession = New-CsOnlineSession -Credential $O365Credentials

  $OnlineSessionComputerName = $OnlineSession.ComputerName
  $null = Import-PSSession -Session $OnlineSession -AllowClobber -ErrorAction Stop

  if ($ConnectToOnPremisesServices){
    Write-VerboseLog -Message 'Connecting to On-Premises pool'
    $OnPremisesSession = New-PSSession -ConnectionURI ('https://{0}/OcsPowershell' -f ($OnPremisesFqdn)) -Credential $OnPremisesCredentials
    $null = Import-PSSession -Session $OnPremisesSession -AllowClobber -ErrorAction Stop
  }

  Write-VerboseLog -Message 'Getting MSOL account name'
  $AccountName = Get-AccountName
  Write-VerboseLog -Message ('MSOL account name: {0}' -f $AccountName)


} catch {
  Write-ScriptError -Message 'Unable to connect to required services' -ErrorObject $_ -Terminating $true
}
#endregion

#region User License Enablement
foreach ($user in $UsersToProcess){
  Write-Host -ForegroundColor Green ('Enabling licenses for user: {0}' -f $user.UserPrincipalName)
	
  #Define user reporting object
  $UserResultReport = New-Object -TypeName PSObject -Property @{
    UPN = ('{0}' -f $user.UserPrincipalName)
    LicenseEnablementResult = ''
    MigrationResult = ''
    ConfigurationResult = ''
    PhoneNumberAssigned = ''
  }
	
  Write-VerboseLog -Message ('Processing license enablement for: {0}' -f ($user | out-string))
  try{
    #Create variable to store the step by step enablement results
    $UserLicenseEnablementResults = @()
	
    #Get User MSOL config
    Write-VerboseLog -Message ('Getting MSOL user config for {0}' -f $user.UserPrincipalName)
    $UserConfig = Get-MsolUser -UserPrincipalName $user.UserPrincipalName  -ErrorAction Stop
    Write-VerboseLog -Message ('{0}' -f ($UserConfig | Format-List | out-string))
    
    #Set User usage location
    Write-VerboseLog -Message ('Setting {0} usage location to {1}' -f $UserConfig.UserPrincipalName, $user.UsageLocation)
    Set-MsolUser -UserPrincipalName $UserConfig.UserPrincipalName -UsageLocation $user.UsageLocation  -ErrorAction Stop

    #Define the base set of services that must be enabled.
    $ServicesToEnable = @()
    $ServicesToEnable += $LicName_SkypeforBusiness
    $ServicesToEnable += $LicName_Exchange

    #Variable to update if a license was already assigned
    $LicensePreviouslyAssigned = $false

    #Check existing user licenses for enterprise license pack
    switch ($UserConfig.Licenses.AccountSkuId){
      (Get-LicenseSku -AccountName $AccountName -LicenseName $LicName_E5){
        $LicensePreviouslyAssigned = $true
        #E5 license assigned - update services
        $UserLicenseEnablementResults += Enable-ServicesInLicensePack -UserConfig $UserConfig -AccountName $AccountName -LicensePackName $LicName_E5 -LicensesToAssign $ServicesToEnable -UpdateType 'Update' -ErrorAction Stop                                                                                                       
      }
      (Get-LicenseSku -AccountName $AccountName -LicenseName $LicName_E3){
        $LicensePreviouslyAssigned = $true
        #E3 license assigned - update services
        $UserLicenseEnablementResults += Enable-ServicesInLicensePack -UserConfig $UserConfig -AccountName $AccountName -LicensePackName $LicName_E3 -LicensesToAssign $ServicesToEnable -UpdateType 'Update' -ErrorAction Stop
        
        #Enable Cloud PSTN Conferencing and Cloud PBX
        if ($user.EnableCloudPSTNConferencing -eq 'TRUE') { $UserLicenseEnablementResults += Enable-AddOnLicense -UserConfig $UserConfig -AccountName $AccountName -LicenseName $LicName_PSTNConferencing  -ErrorAction Stop}
        if ($user.EnableCloudPBX -eq 'TRUE') { $UserLicenseEnablementResults += Enable-AddOnLicense -UserConfig $UserConfig -AccountName $AccountName -LicenseName $LicName_CloudPBX -ErrorAction Stop}                                                                                                                                                                                                                                
      }
      #default{ $LicensePreviouslyAssigned = $true}
    }

    if (-not $LicensePreviouslyAssigned){
      #Get account license Skus
      $LicenseSkus = Get-MsolAccountSku
      
      #No enterprise license assigned - need to assign one
      if ($LicenseSkus.AccountSkuId -contains (Get-LicenseSku -AccountName $AccountName -LicenseName $LicName_E5)) {
        #E5 available
        $UserLicenseEnablementResults += Enable-ServicesInLicensePack -UserConfig $UserConfig -AccountName $AccountName -LicensePackName $LicName_E5 -LicensesToAssign $ServicesToEnable -UpdateType 'New' -ErrorAction Stop
      } elseif ($LicenseSkus.AccountSkuId -contains (Get-LicenseSku -AccountName $AccountName -LicenseName $LicName_E3)) {
        #E3 available
        $UserLicenseEnablementResults += Enable-ServicesInLicensePack -UserConfig $UserConfig -AccountName $AccountName -LicensePackName $LicName_E3 -LicensesToAssign $ServicesToEnable -UpdateType 'New' -ErrorAction Stop
          
        #Enable Cloud PSTN Conferencing and Cloud PBX
        if ($user.EnableCloudPSTNConferencing -eq 'TRUE') { $UserLicenseEnablementResults += Enable-AddOnLicense -UserConfig $UserConfig -AccountName $AccountName -LicenseName $LicName_PSTNConferencing  -ErrorAction Stop}
        if ($user.EnableCloudPBX -eq 'TRUE') { $UserLicenseEnablementResults += Enable-AddOnLicense -UserConfig $UserConfig -AccountName $AccountName -LicenseName $LicName_CloudPBX -ErrorAction Stop}                                                                          
      } else {
        #No enterprise pack licenses available
        Write-Error -Message 'No enterprise pack licenses to assign' -Exception 'There were no enterprise pack licenses found to assign to the user' -ErrorAction stop
      } 
    }

    #Assign licenses for PSTN calling
    if ($user.EnablePSTNCallingDomestic -eq 'TRUE') { $UserLicenseEnablementResults += Enable-AddOnLicense -UserConfig $UserConfig -AccountName $AccountName -LicenseName $LicName_PSTNLocal  -ErrorAction Stop }
    if ($user.EnablePSTNCallingDomesticAndInternational -eq 'TRUE') { $UserLicenseEnablementResults += Enable-AddOnLicense -UserConfig $UserConfig -AccountName $AccountName -LicenseName $LicName_PSTNLocalAndInternational  -ErrorAction Stop}

  } catch {
    Write-ScriptError -Message ('{0}: {1}' -f $LicenseEnablement_Failure, $UserConfig.UserPrincipalName) -ErrorObject $_
    $UserLicenseEnablementResults += ('{0}: {1}' -f $LicenseEnablement_Failure, $UserConfig.UserPrincipalName)
  } 

  if (@($UserLicenseEnablementResults | Group-Object).count -eq 1 -and $UserLicenseEnablementResults[0] -eq $LicenseEnablement_Success) { $UserResultReport.LicenseEnablementResult = $LicenseEnablement_Success }
  else { $UserResultReport.LicenseEnablementResult = [string]::Join(' ', ($UserLicenseEnablementResults | Where-Object {$_ -ne $LicenseEnablement_Success})) }

  #Report user result
  $ResultsReporting += $UserResultReport
  $ResultsReporting | Export-Csv -Force -Path $ResultsFileName -NoTypeInformation
	
  Write-Host ('Result: {0}' -f $UserResultReport.LicenseEnablementResult)
}
#endregion

#region Migrate Users
foreach ($user in $UsersToProcess){
  Write-VerboseLog -Message ('Processing user migration of user: {0}' -f $user.UserPrincipalName)

  #Get the user enablement result from licensing assignment
  $UserEnablementResult = $ResultsReporting | Limit-UserReportingObject -UserPrincipalName $user.UserPrincipalName

  #Create variable to store the step by step enablement results
  $UserMigrationResults = @()
  Write-Host -ForegroundColor Green ('Migrating user: {0}' -f $user.UserPrincipalName)
  
  if ($user.MigrateToSkypeForBusinessOnline -eq 'TRUE'){
    try{

      #Check that user enablement completed successfully
      if ($UserEnablementResult.LicenseEnablementResult -eq $LicenseEnablement_Success){
        
        #Pause to allow license provisioning to complete
        if ($PauseForProvisioning){ Wait-LicenseEnablement -UserPrincipalName $user.UserPrincipalName -ServicesToCheck  @($LicName_SkypeforBusiness) }
        
        #Get CS configuration for the user
        $CsOnlineUser = Get-CsOnlineUser -Identity $user.UserPrincipalName
        Write-VerboseLog -Message ("CsOnlineUser Values:`n {0}" -f ($CsOnlineUser | out-string))
				
        #Check user current hosting location status
        if ($CsOnlineUser.InterpretedUserType -eq 'HybridOnPrem'){
          #User in proper state - try to move user
          Write-VerboseLog -Message 'Attempting to migrate user'
          Move-CsUser -Identity $user.UserPrincipalName -Target Sipfed.online.lync.com -Credential $O365Credentials -HostedMigrationOverrideUrl ('https://{0}/HostedMigration/HostedMigrationService.svc' -f ($OnlineSessionComputerName)) -Confirm:$false -ErrorAction:Stop
          
          #Didn't catch any errors - Migration success
          $UserMigrationResults += $Migration_Success          
        } elseif ($CsOnlineUser.InterpretedUserType -eq 'HybridOnline' -or $CsOnlineUser.InterpretedUserType -eq 'DirSyncedPureOnline'){
          Write-VerboseLog -Message ('{0}: {1} is already homed online.' -f $Migration_NotAttempted_AlreadyOnline, $user.UserPrincipalName)
          $UserMigrationResults += $Migration_NotAttempted_AlreadyOnline      
        } else {
          Write-VerboseLog -Message ('{0}: {1} ' -f $Migration_NotAttempted_NotEnabled, $user.UserPrincipalName)
          $UserMigrationResults += $Migration_NotAttempted_NotEnabled
        }
			
      } else {
        Write-VerboseLog -Message ('{0} was not successful in license enablement.  Not attempting to migrate user.' -f $user.UserPrincipalName)
        $UserMigrationResults += $Migration_NotAttempted
      }

    } catch {
      Write-ScriptError -Message ('Failed to migrate {0}' -f $UserConfig.UserPrincipalName) -ErrorObject $_
      $UserMigrationResults += ('Error occured: {0}.' -f $_.Exception.Message)		
    }	
  } else {
    Write-VerboseLog -Message ('User not designated for migration: {0}' -f $user.UserPrincipalName)
    $UserMigrationResults += $Migration_Skipped
  }

<#
  if (@($UserMigrationResults | Group-Object).count -eq 1 -and $UserMigrationResults[0] -eq 'Migrated') { ($ResultsReporting | Where-Object {$_.UPN -eq  $user.UserPrincipalName}).MigrationResult = 'Migrated' }
  else { ($ResultsReporting | Where-Object {$_.UPN -eq $user.UserPrincipalName}).MigrationResult = [string]::Join(' ', ($UserMigrationResults | Where-Object {$_ -ne 'Migrated'})) }

#>  
  ($ResultsReporting | Limit-UserReportingObject -UserPrincipalName $user.UserPrincipalName).MigrationResult = [string]::Join(' ',$UserMigrationResults)
  
  #Update CSV file with results
  $ResultsReporting | Export-Csv -Force -Path $ResultsFileName -NoTypeInformation	
  Write-Host  ('Result: {0}' -f $($ResultsReporting | Limit-UserReportingObject -UserPrincipalName $user.UserPrincipalName).MigrationResult)

}
#endregion

#region Configure User Features
foreach ($user in $UsersToProcess){
  Write-VerboseLog -Message ('Processing feature configuration of user: {0}' -f $user.UserPrincipalName)

  Write-Host -ForegroundColor Green ('Configuring services for user: {0}' -f $user.UserPrincipalName)

  try{
    #Get the user enablement result from licensing assignment
    $UserEnablementResult = $ResultsReporting | Limit-UserReportingObject -UserPrincipalName $user.UserPrincipalName

    #Create variable to store the step by step enablement results
    $UserFeatureEnablementResults = @()

    if($UserEnablementResult.LicenseEnableResult -notin $SuccessfulResultCollection_LicenseEnablement -and $UserEnablementResult.MigrationResult -notin $SuccessfulResultCollection_Migration){
      #User not successful in enablment and migration operations
      Write-VerboseLog -Message ('{0}: {1} was not successufl in license enablement or migration. {1}' -f $Configuration_NotAttempted,$user.UserPrincipalName)
      $UserFeatureEnablementResults += $Configuration_NotAttempted
    } elseif ($user.EnableCloudPBX -eq 'False' -or [string]::IsNullOrEmpty($user.EnableCloudPBX)){
      Write-VerboseLog -Message ('{0} does not require additional configuartion.' -f $user.UserPrincipalName)
      $UserFeatureEnablementResults += $Configuration_Success
    } elseif ($UserEnablementResult.LicenseEnablementResult -in $SuccessfulResultCollection_LicenseEnablement){
      #User is ready to configure CloudPBX options
      
      #Pause to ensure licenses are fully provisioned
      if ($PauseForProvisioning){ Wait-LicenseEnablement -UserPrincipalName $user.UserPrincipalName -ServicesToCheck  @('MCOSTANDARD','MCOEV','MCOMEETADV','MCOPSTN1','MCOPSTN2','EXCHANGE_S_ENTERPRISE')}
      
      #Get necessary user configuration information
      Write-VerboseLog -Message 'Getting MsolUser configuration'
      $UserConfig = Get-MsolUser -UserPrincipalName $user.UserPrincipalName -ErrorAction Stop
      Write-VerboseLog -Message ('{0}' -f ($UserConfig | Format-List | Out-String))
      Write-VerboseLog -Message 'Getting CsOnlineUser configuration'
      $UserCsConfig = Get-CsOnlineUser -Identity $user.UserPrincipalName -ErrorAction Stop
      Write-VerboseLog -Message ('{0}' -f ($UserCsConfig | Format-List | Out-String))
      Write-VerboseLog -Message 'Getting CsOnlineVoiceConfig configuration'
      $UserCsOnlineVoiceConfig = Get-CsOnlineVoiceUser -Identity $user.UserPrincipalName -ErrorAction stop -WarningAction SilentlyContinue
      Write-VerboseLog -Message ('{0}' -f ($UserCsOnlineVoiceConfig | Out-String))
      
      if (-not [string]::IsNullOrEmpty($user.PhoneNumberToAssign) -and $user.PhoneNumberToAssign -eq $UserCsOnlineVoiceConfig.Number.Id){
        #User has the correct phone number assigned
        Write-VerboseLog -Message 'User already enabled with a phone number.  No need to assign a new phone number'
        ($ResultsReporting | Where-Object {$_.UPN -eq $user.UserPrincipalName}).PhoneNumberAssigned = $UserCsOnlineVoiceConfig.Number.id
        
        
        if(-not [string]::IsNullOrEmpty($user.TenantDialPlan) -and $user.TenantDialPlan -ne $UserCsConfig.TenantDialPlan){
        #DialPlan assignment needs to be updated
            Write-VerboseLog -Message 'DialPlan needs to be assigned'
            $UserFeatureEnablementResults += Grant-TenantDialplan -UserConfig $user -DialPlan $user.TenantDialPlan -ErrorAction Stop
        }


        $UserFeatureEnablementResults += $Configuration_Success                
      } elseif ($user.EnablePSTNCallingDomestic -eq 'true' -or $user.EnablePSTNCallingDomesticAndInternational -eq 'true'){
        #Process PSTNCalling Logic  
        Write-VerboseLog -Message ('Validating CsOnlineEnhancedEmergencyServiceDisclaimer for country or region {0}' -f $user.UsageLocation) 
        $EnhancedEmergencyDisclaimer = Get-CsOnlineEnhancedEmergencyServiceDisclaimer -CountryOrRegion $user.UsageLocation 
        if ($EnhancedEmergencyDisclaimer.Response.Value -ne 'Accepted') {$null = Set-CsOnlineEnhancedEmergencyServiceDisclaimer -CountryOrRegion $user.UsageLocation -ErrorAction stop}

        #Get LocationID
        Write-VerboseLog -Message ('Getting CsOnlineLisCivicAddress with description: {0}' -f $user.EmergencyAddressDescription)
        $LocationId = Get-CsOnlineLisCivicAddress -Description $user.EmergencyAddressDescription -ErrorAction Stop
				
        #Determine phone number to assign 
        if ([string]::IsNullOrEmpty($user.PhoneNumberToAssign)) {
          if ($UserCsOnlineVoiceConfig.Number -eq $null){
            #Get a new phone number to assign
            $PhoneNumberToAssign = (Get-CsOnlineTelephoneNumber -IsNotAssigned -InventoryType Subscriber -CityCode $user.PhoneNumberCityCode -ResultSize 1 -ErrorAction Stop).Id
          } else {
            #Keep the same number
            $PhoneNumberToAssign = $UserCsOnlineVoiceConfig.Number.Id
          }
        } else {
          $PhoneNumberToAssign = $user.PhoneNumberToAssign
        }
				
        $UserFeatureEnablementResults += Set-PSTNCallingPhoneNumber -UserConfig $UserConfig -PhoneNumber $PhoneNumberToAssign -LocationID $LocationId.DefaultLocationId -ErrorAction Stop
        ($ResultsReporting | Where-Object {$_.UPN -eq $user.UserPrincipalName}).PhoneNumberAssigned = $PhoneNumberToAssign

        if(-not [string]::IsNullOrEmpty($user.TenantDialPlan) -and $user.TenantDialPlan -ne $UserCsConfig.TenantDialPlan){
        #DialPlan assignment needs to be updated
            Write-VerboseLog -Message 'DialPlan needs to be assigned'
            $UserFeatureEnablementResults += Grant-TenantDialplan -UserConfig $user -DialPlan $user.TenantDialPlan -ErrorAction Stop
        }

      } elseif ($user.EnableCloudPBX -eq 'true' -and -not [string]::IsNullOrEmpty($user.PhoneNumberToAssign) ){
        #Process Hybrid Logic
        Write-VerboseLog -Message 'Setting user for on-premises pstn connectivity'
        Set-CsUser -Identity $UserConfig.UserPrincipalName -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -OnPremLineURI ('tel:{0}' -f $user.PhoneNumberToAssign) -ErrorAction Stop
        ($ResultsReporting | Where-Object {$_.UPN -eq $user.UserPrincipalName}).PhoneNumberAssigned = $user.PhoneNumberToAssign
        
        if(-not [string]::IsNullOrEmpty($user.TenantDialPlan) -and $user.TenantDialPlan -ne $UserCsConfig.TenantDialPlan){
        #DialPlan assignment needs to be updated
            Write-VerboseLog -Message 'DialPlan needs to be assigned'
            $UserFeatureEnablementResults += Grant-TenantDialplan -UserConfig $user -DialPlan $user.TenantDialPlan -ErrorAction Stop
        }

        #Write-VerboseLog -Message 'Enabling CsOnlineUmMailbox'
        #$null = Enable-CsOnlineUMMailBox -Identity $UserConfig.UserPrincipalName -LineUri $user.PhoneNumberToAssign -ErrorAction Stop
        $UserFeatureEnablementResults += $Configuration_Success                
      } else {
        Write-VerboseLog -Message 'No voice features to configure'
        $UserFeatureEnablementResults += $Configuration_NotRequired                             
      }				
		
    } else {
      Write-VerboseLog -Message ('{0} was not successful in license enablement.  Not attempting to configure user.' -f $user.UserPrincipalName)
      $UserFeatureEnablementResults += $Configuration_NotAttempted
    }

  } catch {
    Write-ScriptError -Message ('Failed to enable {0}' -f $UserConfig.UserPrincipalName) -ErrorObject $_
    $UserFeatureEnablementResults += ('{0}: {1}.' -f $Configuration_Failure, $_.Exception.Message)
  }
	
  if (@($UserFeatureEnablementResults | Group-Object).count -eq 1 -and $UserFeatureEnablementResults[0] -eq $Configuration_Success) { ($ResultsReporting | Where-Object {$_.UPN -eq  $user.UserPrincipalName}).ConfigurationResult = $Configuration_Success }
  else { ($ResultsReporting | Where-Object {$_.UPN -eq $user.UserPrincipalName}).ConfigurationResult = [string]::Join(' ', ($UserFeatureEnablementResults | Where-Object {$_ -ne $Configuration_Success})) }

  $ResultsReporting | Export-Csv -Force -Path $ResultsFileName -NoTypeInformation	
  Write-Host  ('Result: {0}' -f $($ResultsReporting | Where-Object {$_.UPN -eq $user.UserPrincipalName}).ConfigurationResult)

}
#endregion

#region Report results and cleanup

#Report Results
Write-VerboseLog -Message ($ResultsReporting | Format-Table -Property UPN,LicenseEnablementResult,MigrationResult,ConfigurationResult,PhoneNumberAssigned | out-string)
$ResultsReporting | Format-Table -Property UPN,LicenseEnablementResult,MigrationResult,ConfigurationResult,PhoneNumberAssigned

#Cleanup PSSessions
Remove-PSSession -Session $OnlineSession
if ($ConnectToOnPremisesServices) { Remove-PSSession -Session $OnPremisesSession }

#endregion

# SIG # Begin signature block
# MIIdxQYJKoZIhvcNAQcCoIIdtjCCHbICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUjH95roDRFd10s5z0NbzKhbFa
# kqOgghhlMIIEwzCCA6ugAwIBAgITMwAAAMhHIp2jDcrAWAAAAAAAyDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwOTA3MTc1ODU0
# WhcNMTgwOTA3MTc1ODU0WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# Ojk4RkQtQzYxRS1FNjQxMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAoUNNyknhIcQy
# V4oQO4+cu9wdeLc624e9W0bwCDnHpdxJqtEGkv7f+0kYpyYk8rpfCe+H2aCuA5F0
# XoFWLSkOsajE1n/MRVAH24slLYPyZ/XO7WgMGvbSROL97ewSRZIEkFm2dCB1DRDO
# ef7ZVw6DMhrl5h8s299eDxEyhxrY4i0vQZKKwDD38xlMXdhc2UJGA0GZ16ByJMGQ
# zBqsuRyvxAGrLNS5mjCpogEtJK5CCm7C6O84ZWSVN8Oe+w6/igKbq9vEJ8i8Q4Vo
# hAcQP0VpW+Yg3qmoGMCvb4DVRSQMeJsrezoY7bNJjpicVeo962vQyf09b3STF+cq
# pj6AXzGVVwIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFA/hZf3YjcOWpijw0t+ejT2q
# fV7MMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAJqUDyiyB97jA9U9vp7HOq8LzCIfYVtQfJi5PUzJrpwzv6B7
# aoTC+iCr8QdiMG7Gayd8eWrC0BxmKylTO/lSrPZ0/3EZf4bzVEaUfAtChk4Ojv7i
# KCPrI0RBgZ0+tQPYGTjiqduQo2u4xm0GbN9RKRiNNb1ICadJ1hkf2uzBPj7IVLth
# V5Fqfq9KmtjWDeqey2QBCAG9MxAqMo6Epe0IDbwVUbSG2PzM+rLSJ7s8p+/rxCbP
# GLixWlAtuY2qFn01/2fXtSaxhS4vNzpFhO/z/+m5fHm/j/88yzRvQfWptlQlSRdv
# wO72Vc+Nbvr29nNNw662GxDbHDuGN3S65rjPsAkwggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCBhEwggP5
# oAMCAQICEzMAAACOh5GkVxpfyj4AAAAAAI4wDQYJKoZIhvcNAQELBQAwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xNjExMTcyMjA5MjFaFw0xODAy
# MTcyMjA5MjFaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDQh9RCK36d2cZ61KLD4xWS
# 0lOdlRfJUjb6VL+rEK/pyefMJlPDwnO/bdYA5QDc6WpnNDD2Fhe0AaWVfIu5pCzm
# izt59iMMeY/zUt9AARzCxgOd61nPc+nYcTmb8M4lWS3SyVsK737WMg5ddBIE7J4E
# U6ZrAmf4TVmLd+ArIeDvwKRFEs8DewPGOcPUItxVXHdC/5yy5VVnaLotdmp/ZlNH
# 1UcKzDjejXuXGX2C0Cb4pY7lofBeZBDk+esnxvLgCNAN8mfA2PIv+4naFfmuDz4A
# lwfRCz5w1HercnhBmAe4F8yisV/svfNQZ6PXlPDSi1WPU6aVk+ayZs/JN2jkY8fP
# AgMBAAGjggGAMIIBfDAfBgNVHSUEGDAWBgorBgEEAYI3TAgBBggrBgEFBQcDAzAd
# BgNVHQ4EFgQUq8jW7bIV0qqO8cztbDj3RUrQirswUgYDVR0RBEswSaRHMEUxDTAL
# BgNVBAsTBE1PUFIxNDAyBgNVBAUTKzIzMDAxMitiMDUwYzZlNy03NjQxLTQ0MWYt
# YmM0YS00MzQ4MWU0MTVkMDgwHwYDVR0jBBgwFoAUSG5k5VAF04KqFzc3IrVtqMp1
# ApUwVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9jcmwvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNybDBhBggrBgEF
# BQcBAQRVMFMwUQYIKwYBBQUHMAKGRWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9w
# a2lvcHMvY2VydHMvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNydDAMBgNV
# HRMBAf8EAjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBEiQKsaVPzxLa71IxgU+fKbKhJ
# aWa+pZpBmTrYndJXAlFq+r+bltumJn0JVujc7SV1eqVHUqgeSxZT8+4PmsMElSnB
# goSkVjH8oIqRlbW/Ws6pAR9kRqHmyvHXdHu/kghRXnwzAl5RO5vl2C5fAkwJnBpD
# 2nHt5Nnnotp0LBet5Qy1GPVUCdS+HHPNIHuk+sjb2Ns6rvqQxaO9lWWuRi1XKVjW
# kvBs2mPxjzOifjh2Xt3zNe2smjtigdBOGXxIfLALjzjMLbzVOWWplcED4pLJuavS
# Vwqq3FILLlYno+KYl1eOvKlZbiSSjoLiCXOC2TWDzJ9/0QSOiLjimoNYsNSa5jH6
# lEeOfabiTnnz2NNqMxZQcPFCu5gJ6f/MlVVbCL+SUqgIxPHo8f9A1/maNp39upCF
# 0lU+UK1GH+8lDLieOkgEY+94mKJdAw0C2Nwgq+ZWtd7vFmbD11WCHk+CeMmeVBoQ
# YLcXq0ATka6wGcGaM53uMnLNZcxPRpgtD1FgHnz7/tvoB3kH96EzOP4JmtuPe7Y6
# vYWGuMy8fQEwt3sdqV0bvcxNF/duRzPVQN9qyi5RuLW5z8ME0zvl4+kQjOunut6k
# LjNqKS8USuoewSI4NQWF78IEAA1rwdiWFEgVr35SsLhgxFK1SoK3hSoASSomgyda
# Qd691WZJvAuceHAJvDCCB3owggVioAMCAQICCmEOkNIAAAAAAAMwDQYJKoZIhvcN
# AQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAw
# BgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEx
# MB4XDTExMDcwODIwNTkwOVoXDTI2MDcwODIxMDkwOVowfjELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUg
# U2lnbmluZyBQQ0EgMjAxMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# AKvw+nIQHC6t2G6qghBNNLrytlghn0IbKmvpWlCquAY4GgRJun/DDB7dN2vGEtgL
# 8DjCmQawyDnVARQxQtOJDXlkh36UYCRsr55JnOloXtLfm1OyCizDr9mpK656Ca/X
# llnKYBoF6WZ26DJSJhIv56sIUM+zRLdd2MQuA3WraPPLbfM6XKEW9Ea64DhkrG5k
# NXimoGMPLdNAk/jj3gcN1Vx5pUkp5w2+oBN3vpQ97/vjK1oQH01WKKJ6cuASOrdJ
# Xtjt7UORg9l7snuGG9k+sYxd6IlPhBryoS9Z5JA7La4zWMW3Pv4y07MDPbGyr5I4
# ftKdgCz1TlaRITUlwzluZH9TupwPrRkjhMv0ugOGjfdf8NBSv4yUh7zAIXQlXxgo
# tswnKDglmDlKNs98sZKuHCOnqWbsYR9q4ShJnV+I4iVd0yFLPlLEtVc/JAPw0Xpb
# L9Uj43BdD1FGd7P4AOG8rAKCX9vAFbO9G9RVS+c5oQ/pI0m8GLhEfEXkwcNyeuBy
# 5yTfv0aZxe/CHFfbg43sTUkwp6uO3+xbn6/83bBm4sGXgXvt1u1L50kppxMopqd9
# Z4DmimJ4X7IvhNdXnFy/dygo8e1twyiPLI9AN0/B4YVEicQJTMXUpUMvdJX3bvh4
# IFgsE11glZo+TzOE2rCIF96eTvSWsLxGoGyY0uDWiIwLAgMBAAGjggHtMIIB6TAQ
# BgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQUSG5k5VAF04KqFzc3IrVtqMp1ApUw
# GQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB
# /wQFMAMBAf8wHwYDVR0jBBgwFoAUci06AjGQQ7kUBU7h6qfHMdEjiTQwWgYDVR0f
# BFMwUTBPoE2gS4ZJaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJv
# ZHVjdHMvTWljUm9vQ2VyQXV0MjAxMV8yMDExXzAzXzIyLmNybDBeBggrBgEFBQcB
# AQRSMFAwTgYIKwYBBQUHMAKGQmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kv
# Y2VydHMvTWljUm9vQ2VyQXV0MjAxMV8yMDExXzAzXzIyLmNydDCBnwYDVR0gBIGX
# MIGUMIGRBgkrBgEEAYI3LgMwgYMwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2lvcHMvZG9jcy9wcmltYXJ5Y3BzLmh0bTBABggrBgEFBQcC
# AjA0HjIgHQBMAGUAZwBhAGwAXwBwAG8AbABpAGMAeQBfAHMAdABhAHQAZQBtAGUA
# bgB0AC4gHTANBgkqhkiG9w0BAQsFAAOCAgEAZ/KGpZjgVHkaLtPYdGcimwuWEeFj
# kplCln3SeQyQwWVfLiw++MNy0W2D/r4/6ArKO79HqaPzadtjvyI1pZddZYSQfYtG
# UFXYDJJ80hpLHPM8QotS0LD9a+M+By4pm+Y9G6XUtR13lDni6WTJRD14eiPzE32m
# kHSDjfTLJgJGKsKKELukqQUMm+1o+mgulaAqPyprWEljHwlpblqYluSD9MCP80Yr
# 3vw70L01724lruWvJ+3Q3fMOr5kol5hNDj0L8giJ1h/DMhji8MUtzluetEk5CsYK
# wsatruWy2dsViFFFWDgycScaf7H0J/jeLDogaZiyWYlobm+nt3TDQAUGpgEqKD6C
# PxNNZgvAs0314Y9/HG8VfUWnduVAKmWjw11SYobDHWM2l4bf2vP48hahmifhzaWX
# 0O5dY0HjWwechz4GdwbRBrF1HxS+YWG18NzGGwS+30HHDiju3mUv7Jf2oVyW2ADW
# oUa9WfOXpQlLSBCZgB/QACnFsZulP0V3HjXG0qKin3p6IvpIlR+r+0cjgPWe+L9r
# t0uX4ut1eBrs6jeZeRhL/9azI2h15q/6/IvrC4DqaTuv/DDtBEyO3991bWORPdGd
# Vk5Pv4BXIqF4ETIheu9BCrE/+6jMpF3BoYibV3FWTkhFwELJm3ZbCoBIa/15n8G9
# bW1qyVJzEw16UM0xggTKMIIExgIBATCBlTB+MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5n
# IFBDQSAyMDExAhMzAAAAjoeRpFcaX8o+AAAAAACOMAkGBSsOAwIaBQCggd4wGQYJ
# KoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQB
# gjcCARUwIwYJKoZIhvcNAQkEMRYEFOZlS90Vn+OSJ1NZixGdcXQ5elOtMH4GCisG
# AQQBgjcCAQwxcDBuoEyASgBFAG4AYQBiAGwAZQAtAFMAawB5AHAAZQBGAG8AcgBC
# AHUAcwBpAG4AZQBzAHMAVgBvAGkAYwBlAFUAcwBlAHIAcwAuAHAAcwAxoR6AHGh0
# dHA6Ly93d3cubWljcm9zb2Z0LmNvbS91YyAwDQYJKoZIhvcNAQEBBQAEggEAMWwi
# yoPn78dT/v8KQ+CuJIMFvzEdfiz/ArhULLRaYZHVb1VsFEBOmLQYdyFI8M69qCDk
# Y6AO7Rhk3DR43/TduJXm0jyZSoSYuAD3/byeQAdr/ohJW0HnMcrXYdckLfBh1TGN
# iDJ0YRTvotgKh1dx88DWx5lgWjBbLxodYff0Tep/mEda7iFLYs/BQgPELL9vF2Am
# BJGkI/N0h8HQTQKjqp0bEtG0cnbuk9wCBtqz5bpTMMC/01+feNUY5S5CB17BhnXX
# +r7L3YiVDdkyBXzwRzf+IIbEx2qbe170MIgukMXmkZAWven5oW+9hLCRU4M+JcjK
# 8SmvR4LtKY2qCxHB7KGCAigwggIkBgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhN
# aWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAADIRyKdow3KwFgAAAAAAMgwCQYF
# Kw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkF
# MQ8XDTE3MDMyMzIwMjQyM1owIwYJKoZIhvcNAQkEMRYEFJ3qCF6doxSkzNUttLs2
# l8AYqZshMA0GCSqGSIb3DQEBBQUABIIBAJB3lAbk4+h4USNfg/juVZBp8cHgH0VA
# bZtSfyERYsBiw2DUo0xSS7IWnlSO3DHS03c6/ZP1W6ST2DoireeYkiNksUKyUCQK
# tgC0f5Qdw3x/tSc4/KZ8lnqZm8qhAURyjG3IzgeAWao4BShVSMvlgr+RfPe4ukIi
# LJOtnxDjEuMYbfZJ71Gf05ktwTlN/95zNrO0sKyVXWDQrqwKDGrEUtf35vEI9QnQ
# Hlhbd/yxvc1m1+rzBnCweex4wDWFbPjVkYfoSb1MxqfcbQEx+AmAIRS9zdQ/KUJb
# dkP4en7ufPvx1q8Eik32xrW6bZwLTsqSLoh+C8aIIepgZhIq/gR7r+4=
# SIG # End signature block
