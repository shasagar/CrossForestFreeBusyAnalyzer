

[CmdletBinding()]

Param
(
	[string]$LogFile=".\$FolderName\FreeBusyAnalyzer.log",
	[Switch]$NonInteractive
   
)
$Time = Get-Date -Format MMddyyyyhhmmss
$FolderName = "LogFiles" + $Time
New-Item -itemtype Directory -Path .\$FolderName
# Function to write activities in the log file
Function WriteLog 
{
	Param ([string]$string, [String]$Color)
	
# Get the current date
	[string]$date = Get-Date -Format G
		
# Write everything to the log file
    
	( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
	
# If NonInteractive true then supress host output
	if (!($NonInteractive)){
		( "[" + $date + "] - " + $string) | Write-Host -f $Color
	}
}

# Setup a new O365 Powershell Session

Function New-LocalExchangeSession 
{
    $Error.Clear()
    $ADExchangeURL = "http://$ServerName/PowerShell"
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ADExchangeURL -Authentication kerberos -Credential $ADCreds -Name "Local.Exchange"
    Import-PSSession $session -AllowClobber -ErrorAction SilentlyContinue
    if($Error)
    {
        if (!(get-pssession | Where-Object { $_.Name -match "Local.Exchange"}))
        {
            WriteLog -color Red "Error: Local Exchange Server session was unsuccessful, plese make sure provided creds or server name are accurate"
            EXIT
        }

    }
    Write-verbose "Testing for Active Directory PowerShell module"

        if (!(get-module ActiveDirectory)) 
        { 
        import-module ActiveDirectory 
        }          
    
        $ADView = Get-ADServerSettings
        If (-Not $ADView.ViewEntireForest)
        {
            Set-ADServerSettings -ViewEntireForest:$True
        }
    
}
Function New-O365Session 
{
    Get-PSSession | Where-object {$_.state -eq 'broken'}| Remove-PSSession -Confirm:$false
    [System.GC]::Collect()
    $Error.Clear()
# Create the session
    WriteLog "Creating new PS Session" -Color yellow
    Install-Module ExchangeOnlineManagement -Force
    Install-Module MSOnline
    Install-Module Azure
    Connect-ExchangeOnline
    Import-Module MSOnline	
    Connect-MsolService	
}
#Collect Inputs
Write-host -f cyan "Note: Make sure you have opened Windows PowerShell with Elevated Privileges"
Write-host -f cyan "Please provide exchange on-premise admin creds"
$ADCreds = Get-Credential
$ServerName = Read-Host "Please supply one of the local exchange servers' name"
$OnpremRcpt = Read-host "One of the on-premise recipients' email"
$EXORcpt = Read-host "One of the EXO recipients' email address"
$RemoteRoutingDomain = Read-host "Tenent Remote Routing domain (ex: contoso.mail.onmicrosoft.com)"
$DirectionOftheIssue = Read-host "In which direction free/busy is NOT working? (OnpremToEXO / EXOToOnprem)"

Writelog -color cyan "Collecting Exchange on-premise Free/Busy configuration"
Disconnect-ExchangeOnline
New-LocalExchangeSession
$Error.clear()
Get-Recipient $OnpremRcpt -ErrorAction SilentlyContinue | Export-Clixml .\$FolderName\OnpremMBX.xml
if ($Error)
{
    WriteLog -Color Red "Error: Provided on-prem email $OnpremRcpt does not exist in exchange on-prem environment"
    EXIT
}
$Error.clear()
Get-Recipient $EXORcpt -ErrorAction SilentlyContinue | Export-Clixml .\$FolderName\EXORcpt.xml
if ($Error)
{
    WriteLog -Color Red "Error: Provided EXO email $EXORcpt does not exist in exchange on-prem environment"
    EXIT
}
Get-ExchangeServer | Export-Clixml C:\temp\ExInfo.xml
Get-IntraOrganizationConnector | Export-Clixml .\$FolderName\OnPremIOC.xml
Get-OrganizationRelationship | Export-Clixml .\$FolderName\OnPremOR.xml
Get-AuthServer | Export-Clixml .\$FolderName\AuthServer.xml
Get-AuthConfig | Export-Clixml .\$FolderName\AuthConfig.xml
Get-WebServicesVirtualDirectory | Export-Clixml .\$FolderName\WebVD.xml
Get-ExchangeCertificate | Export-Clixml .\$FolderName\ExchangeCert.xml
Get-FederationTrust | Export-Clixml .\$FolderName\FederationTrust.xml

Writelog -color cyan "Collecting Exchange Online Free/Busy Configuration"
New-O365Session
$Error.clear()
Get-EXORecipient $OnpremRcpt -ErrorAction SilentlyContinue | Export-Clixml .\$FolderName\OnpremRcpt.xml
if ($Error)
{
    WriteLog -Color Red "Error: Provided on-prem email $OnpremRcpt does not exist in exchange Online environment"
}
$Error.clear()
Get-EXORecipient $EXORcpt -ErrorAction SilentlyContinue | Export-Clixml .\$FolderName\EXOMBX.xml
if ($Error)
{
    WriteLog -Color Red "Error: Provided EXO email $EXORcpt does not exist in exchange online environment"
}
Get-IntraOrganizationConnector | Export-Clixml .\$FolderName\EXOIOC.xml
Get-OrganizationRelationship | Export-Clixml .\$FolderName\EXOOR.xml
Get-MsolServicePrincipalCredential -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000" -ReturnKeyValues $true | Export-Clixml .\$FolderName\MSOLPrincipalCreds.xml

If($DirectionOftheIssue -eq 'OnpremToEXO')
{
#EXO Mailbox healthcheck
    $EXOMbx = Get-EXOMailbox $EXORcpt
    $emails = $exoMbx.EmailAddresses
    foreach($email in $emails)
    {
        $emailMatch = "No"
        if($email -like "*$RemoteRoutingDomain*")
        {
            $emailMatch = "Yes"
            Break
        }
    }
        if($emailMatch -like 'No')
        {
            WriteLog -color Red "Error: EXO Mailbox does not have $RemoteRoutingDomain email address stamped"
        }
    

#EXO Mailbox calendar permission check
    $EXOAlias = $EXOMbx.Alias
    $CalFolder = $EXOAlias + ":\Calendar"
    $EXOMbxCalPerms = Get-MailboxFolderPermission $CalFolder
    foreach($perm in $exoMbxCalPerms)
    {
        if($perm.user -like 'Default' -and $perm.accessrights -Notlike 'AvailabilityOnly')
        {
            WriteLog -Color Yellow "Warning: Please check and confirm that Default Calendar folder permission is at least set to AvailabilityOnly"
        }
    }

    WriteLog -color cyan "Disconnecting existing exchange online session"
    Disconnect-ExchangeOnline

    WriteLog -color cyan "Reconnecting Local Exchange on-prem session"
    New-LocalExchangeSession

    WriteLog -color cyan "Checking on-premise free/busy configurations"
    if((Get-RemoteMailbox $EXORcpt).RemoteRoutingAddress -NotLike '*mail.onmicrosoft.com*')
    {
        WriteLog -color Red "Error: EXO Recipient does not have correct Remote Routing Address set"
    }
    
#IntraOrganization Connector [IOC] configuration check
    $OnpremIOC = Get-IntraOrganizationConnector | Where-Object {$_.TargetAddressDomains -eq $RemoteRoutingDomain}
    if($OnpremIOC.Enabled -eq $TRUE)
    {
        WriteLog -color cyan "IOC is enabled, performing OAUTH checks.."
        $OAuthResult = Test-OAUTHConnectivity -Service AutoD -TargetUri $OnPremIOC.DiscoveryEndpoint -Mailbox $OnpremRcpt
        $OAuthResult | Export-Clixml .\$FolderName\OAuthResult.xml
        if($OAUTHResult.ResultType -Like '*Fail*')
        {    
            WriteLog -color Red "Error: OAUTH Test failed, please review OAUTHResult.xml for more information"
        }
    }

#Checking OAUTH on Web Services Virtual Directory
            $WebServices = Get-WebServicesVirtualDirectory
            foreach($WebService in $WebServices)
            {
                $OAUTHChk = $WebService.OAuthAuthentication
                if($OAUTHChk -eq $False)
                {
                WriteLog -color Red "Error: OAUTH Authentication is NOT enabled on Server $WebService"
                }
            }
#Checking AuthServer Configuration
            $AuthServer = Get-AuthServer
            If(!$AuthServer)
            {
                WriteLog -color Red "Error: No AuthServer was found, please follow URL https://learn.microsoft.com/en-us/exchange/configure-oauth-authentication-between-exchange-and-exchange-online-organizations-exchange-2013-help"            
            }
            else 
            {
                $AuthServerName = $AuthServer | where-object {$_.Name -like '*ACS*'}
                if($AuthServerName.count -ge 2)
                {
                    WriteLog -color yellow "Warning: OAUTH is configured with more than 1 tenant, need to make sure DomainName in ACS OAUTH server is configured correctly to get OAUTH token; contact MS Support"
                }
            }    
            
#Checking certificate on AuthConfig
            $AuthConfig = Get-AuthConfig
            $ConfigCert = $AuthConfig.CurrentCertificateThumbprint
            $ExCerts = Get-ExchangeCertificate $ConfigCert
            if($ExCerts.NotAfter -lt (Get-Date))
            {
                Writelog -color Red "Error: Auth Certificate is expired, please recreate Auth Cert"
            }
            $MSOLCertInfo = Get-MsolServicePrincipalCredential -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000" -ReturnKeyValues $true
            if($ExCerts.NotAfter.date -ne $MSOLCertInfo.EndDate.date)
            {
                WriteLog -color Red "Error: MSOL certificate is NOT matching with on-premise certificate, reconfigure OAUTH manually or re-run HCW"
            }
    if($OnpremIOC.Enabled -eq $False)
    {
        WriteLog -color cyan "Checking OrganizationRelationship [OR] configuration"
        $OnpremOR = Get-OrganizationRelationship | Where-Object {$_.DomainNames -eq $RemoteRoutingDomain}
        if($OnpremOR.Enabled -eq $False)
        {
            WriteLog -color Red "Error: OrganizationRelationship is disabled. As IOC is also disabled for the RemoteDomain, please make sure to enable OrganizationRelationship to make Free/Busy working"
        }
        if($OnpremOR.Enabled -eq $TRUE)
        {
            if($OnpremOR.FreeBusyAccessEnabled -eq $False -or $NULL -eq $OnpremOR.FreeBusyAccessLevel -or $OnPremOR.TargetApplicationUri -NotLike '*Outlook.com*' -or $NULL -eq $OnPremOR.TargetAutodiscoverEpr)
            {
                WriteLog -color Red "Error: OrganizationRelationship configuration is Not configured correctly"
                $OnPremOR | Format-list FreeBusyAccessEnabled,FreeBusyAccessLevel,TargetApplicationUri,TargetAutodiscoverEpr
                
                write-host -f cyan "Please make sure above Output is matching with the Sample output below:
                FreeBusyAccessEnabled : True
                FreeBusyAccessLevel   : LimitedDetails
                TargetApplicationUri  : outlook.com
                TargetAutodiscoverEpr : https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"
                
            }
        }
    }
}

If($DirectionOftheIssue -eq 'EXOToOnprem')
{
    $OnpremAppURI = (Get-FederationTrust).ApplicationURI
#Onprem Mailbox calendar permission check
    $OnpremMBX = Get-Mailbox $OnpremRcpt
    $OnpremAlias = $OnpremMbx.Alias
    $CalFolder = $OnpremAlias + ":\Calendar"
    $OnpremMbxCalPerms = Get-MailboxFolderPermission $CalFolder
    foreach($perm in $OnpremMbxCalPerms)
    {
        if($perm.user -like 'Default' -and $perm.accessrights -Notlike 'AvailabilityOnly')
        {
            WriteLog -Color Yellow "Warning: Please check and confirm that Default Calendar folder permission is at least set to AvailabilityOnly"
        }
    }
    New-O365Session
#Checking Onprem MBX existence in EXO
    if(!(Get-Recipient $OnpremRcpt))
    {
        Writelog -color -Red "Error: Onprem email address does not exist in EXO, please make sure the object is in sync"
        EXIT
    }
#IntraOrganization Connector [IOC] configuration check
    $PrimaryDomain = $OnpremRcpt.split('@')[1]
    $EXOIOC = Get-IntraOrganizationConnector | Where-Object {$_.TargetAddressDomains -eq $PrimaryDomain}
    if($EXOIOC.DiscoveryEndpoint -like '*msappproxy*')
    {
        Writelog -color cyan "Hybrid Modern is configured with Hybrid Agent."
        Writelog -color cyan "Make sure External URL set on the Hybrid Agent configuration, is accessiable. Review article https://learn.microsoft.com/en-us/exchange/hybrid-deployment/hybrid-agent for more info on Hybrid Agent"
    }
    if($EXOIOC.Enabled -eq $TRUE)
    {
        WriteLog -color cyan "IOC is enabled, performing OAUTH checks.."
        $OAuthResult = Test-OAUTHConnectivity -Service AutoD -TargetUri $EXOIOC.DiscoveryEndpoint -Mailbox $EXORcpt
        $OAuthResult | Export-Clixml .\$FolderName\EXO_OAuthResult.xml
        if($OAUTHResult.ResultType -Like '*Fail*')
        {    
            WriteLog -color Red "Error: EXO OAUTH Test failed, please review EXO_OAUTHResult.xml for more information"
        }
    }        
    if($EXOIOC.Enabled -eq $False)
    {
        WriteLog -color cyan "Checking OrganizationRelationship [OR] configuration"
        $EXOOR = Get-OrganizationRelationship | Where-Object {$_.DomainNames -eq $PrimaryDomain}
        if($EXOOR.Enabled -eq $False)
        {
            WriteLog -color Red "Error: OrganizationRelationship is disabled. As IOC is also disabled for the RemoteDomain, please make sure to enable OrganizationRelationship to make Free/Busy working"
        }
        if($EXOOR.Enabled -eq $TRUE)
        {
            if($EXOOR.TargetAutodiscoverEpr -like '*MSAppProxy*')
            {
                Writelog -color cyan "Hybrid Modern is configured with Hybrid Agent."
                Writelog -color cyan "Make sure External URL set on the Hybrid Agent configuration, is accessiable. Review article https://learn.microsoft.com/en-us/exchange/hybrid-deployment/hybrid-agent for more info on Hybrid Agent"
    
            }
            if($EXOOR.FreeBusyAccessEnabled -eq $False -or $NULL -eq $EXOOR.FreeBusyAccessLevel -or $NULL -eq $EXOOR.TargetApplicationUri -or $NULL -eq $OnPremOR.TargetAutodiscoverEpr)
            {
                WriteLog -color Red "Error: OrganizationRelationship configuration is Not configured correctly"
                $EXOOR | Format-list FreeBusyAccessEnabled,FreeBusyAccessLevel,TargetApplicationUri,TargetAutodiscoverEpr
                
                write-host -f cyan "Please make sure above Output is matching with the Sample output below:
                FreeBusyAccessEnabled : True
                FreeBusyAccessLevel   : LimitedDetails
                TargetApplicationUri  : $OnpremAppURI
                TargetAutodiscoverEpr : https://autodiscover.contoso.com/autodiscover/autodiscover.svc"
                
            }
        }
    }
}

Writelog -color Yellow "Script has reviewed basic configuration and collected required configuration logs for MS to review. If the issue still unidentified, please perform following steps:                                       
1) Enable Outlook Diagnostic loggin and restart outlook                                                                                                                                                                          
2) Install fiddler classic [https://www.telerik.com/download/fiddler] and enable HTTPS decryption                                                                                                                                
3) Open fiddler and let it running                                                                                                                                                                                              
4) Reproduce the issue by accessing cross forest free/busy                                                                                                                                                                      
5) Save fiddler logs / Outlook Diagnostic Logs / IIS Logs / HTTPProxy logs of Autodiscover & EWS"                                                                                                                                


Writelog -color Yellow "A few other basic checks:                                                                                                                                                                                                                                                                                                                                                                 
1) Make sure Autodiscover and/or EWS URL is published. In the case of Hybrid Agent, make sure the Hybrid Agent is accessiable.                                                                                                   
2) Make sure all the required EXO IPs are allowed. Ref. article [https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide]                                                      

Other useful article: https://techcommunity.microsoft.com/t5/exchange-team-blog/demystifying-hybrid-free-busy-what-are-the-moving-parts/ba-p/607704"
