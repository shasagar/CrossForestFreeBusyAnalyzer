[CmdletBinding()]
Param
(
	[string]$logFile=".\$folderName\FreeBusyAnalyzer.log",
	[Switch]$nonInteractive   
)
$time = Get-Date -Format MMddyyyyhhmmss
$folderName = "LogFiles" + $time
New-Item -itemtype Directory -Path .\$folderName
# Function to write activities in the log file
Function WriteLog 
{
	Param ([string]$string, [String]$color)
# Get the current date
	[string]$date = Get-Date -Format G
# Write everything to the log file
	( "[" + $date + "] - " + $string) | Out-File -FilePath $logFile -Append
# If NonInteractive true then supress host output
	if (!($nonInteractive)){
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
#Collect required Inputs
Write-host -f cyan "Note: Make sure you have opened Windows PowerShell with Elevated Privileges"
Write-host -f cyan "Provide exchange on-premise admin creds"
$ADCreds = Get-Credential
$serverName = Read-Host "Supply one of the local exchange servers' name"
$onpremRcpt = Read-host "One of the on-premise recipients' email address"
$exoRcpt = Read-host "One of the Exchange online recipients' email address"
$remoteRoutingDomain = Read-host "Tenent Remote Routing domain (ex: contoso.mail.onmicrosoft.com)"
$directionOftheIssue = Read-host "In which direction free/busy is NOT working? (OnpremToEXO / EXOToOnprem)"
Writelog -color cyan "Collecting Exchange on-premise Free/Busy configuration"
Disconnect-ExchangeOnline -confirm:$false
New-LocalExchangeSession
$Error.clear()
Get-Recipient $onpremRcpt -ErrorAction SilentlyContinue | Export-Clixml .\$FolderName\OnpremMBX.xml
if ($Error)
{
    WriteLog -Color Red "Error: Provided on-prem email $onpremRcpt does not exist in exchange on-prem environment"
    EXIT
}
$Error.clear()
Get-Recipient $exoRcpt -ErrorAction SilentlyContinue | Export-Clixml .\$FolderName\EXORcpt.xml
if ($Error)
{
    WriteLog -Color Red "Error: Provided EXO email $exoRcpt does not exist in exchange on-prem environment"
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
    WriteLog -Color Red "Error: Provided on-prem email $onpremRcpt does not exist in exchange Online environment"
}
$Error.clear()
Get-EXORecipient $exoRcpt -ErrorAction SilentlyContinue | Export-Clixml .\$FolderName\EXOMBX.xml
if ($Error)
{
    WriteLog -Color Red "Error: Provided EXO email $exoRcpt does not exist in exchange online environment"
}
Get-IntraOrganizationConnector | Export-Clixml .\$FolderName\EXOIOC.xml
Get-OrganizationRelationship | Export-Clixml .\$FolderName\EXOOR.xml
Get-MsolServicePrincipalCredential -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000" -ReturnKeyValues $true | Export-Clixml .\$FolderName\MSOLPrincipalCreds.xml

If($DirectionOftheIssue -eq 'OnpremToEXO')
{
#EXO Mailbox healthcheck
    $exoMbx = Get-EXOMailbox $exoRcpt
    $emails = $exoMbx.EmailAddresses
    foreach($email in $emails)
    {
        $emailMatch = "No"
        if($email -like "*$remoteRoutingDomain*")
        {
            $emailMatch = "Yes"
            Break
        }
    }
        if($emailMatch -like 'No')
        {
            WriteLog -color Red "Error: EXO Mailbox does not have $remoteRoutingDomain email address stamped"
        }
#EXO Mailbox calendar permission check
    $exoAlias = $exoMbx.Alias
    $calFolder = $exoAlias + ":\Calendar"
    $exoMbxCalPerms = Get-MailboxFolderPermission $calFolder
    foreach($perm in $exoMbxCalPerms)
    {
        if($perm.user -like 'Default' -and $perm.accessrights -Notlike 'AvailabilityOnly')
        {
            WriteLog -Color Yellow "Warning: make sure that Default Calendar folder permission is at least set to AvailabilityOnly"
        }
    }
    WriteLog -color cyan "Disconnecting existing exchange online session"
    Disconnect-ExchangeOnline -confirm:$false

    WriteLog -color cyan "Reconnecting Local Exchange on-prem session"
    New-LocalExchangeSession

    WriteLog -color cyan "Checking on-premise free/busy configurations"
    if((Get-RemoteMailbox $exoRcpt).RemoteRoutingAddress -NotLike '*mail.onmicrosoft.com*')
    {
        WriteLog -color Red "Error: EXO Recipient does not have correct Remote Routing Address set"
    }
    
#IntraOrganization Connector [IOC] configuration check
    $OnpremIOC = Get-IntraOrganizationConnector | Where-Object {$_.TargetAddressDomains -eq $remoteRoutingDomain}
    if($onpremIOC.Enabled -eq $TRUE)
    {
        WriteLog -color cyan "IOC is enabled, performing OAUTH checks.."
        $oauthResult = Test-OAUTHConnectivity -Service AutoD -TargetUri $onPremIOC.DiscoveryEndpoint -Mailbox $onpremRcpt
        $oauthResult | Export-Clixml .\$FolderName\oauthResult.xml
        if($oauthResult.ResultType -Like '*Fail*')
        {    
            WriteLog -color Red "Error: OAUTH Test failed, please review OAUTHResult.xml for more information"
        }
    }

#Checking OAUTH on Web Services Virtual Directory
            $webServices = Get-WebServicesVirtualDirectory
            foreach($webService in $webServices)
            {
                $oauthChk = $WebService.OAuthAuthentication
                if($OAUTHChk -eq $False)
                {
                WriteLog -color Red "Error: OAUTH Authentication is NOT enabled on Server $webService"
                }
            }
#Checking AuthServer Configuration
            $authServer = Get-AuthServer
            If(!$authServer)
            {
                WriteLog -color Red "Error: No AuthServer was found, please follow URL https://learn.microsoft.com/en-us/exchange/configure-oauth-authentication-between-exchange-and-exchange-online-organizations-exchange-2013-help"            
            }
            else 
            {
                $authServerName = $authServer | where-object {$_.Name -like '*ACS*'}
                if($authServerName.count -ge 2)
                {
                    WriteLog -color yellow "Warning: OAUTH is configured with more than 1 tenant, need to make sure DomainName in ACS OAUTH server is configured correctly to get OAUTH token; contact MS Support"
                }
            }    
            
#Checking certificate on AuthConfig
            $authConfig = Get-AuthConfig
            $configCert = $authConfig.CurrentCertificateThumbprint
            $exCerts = Get-ExchangeCertificate $configCert
            if($exCerts.NotAfter -lt (Get-Date))
            {
                Writelog -color Red "Error: Auth Certificate is expired, please recreate Auth Cert"
            }
            $msolCertInfo = Get-MsolServicePrincipalCredential -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000" -ReturnKeyValues $true
            if($exCerts.NotAfter.date -ne $msolCertInfo.EndDate.date)
            {
                WriteLog -color Red "Error: MSOL certificate is NOT matching with on-premise certificate, reconfigure OAUTH manually or re-run HCW"
            }
    if($onpremIOC.Enabled -eq $False)
    {
        WriteLog -color cyan "Checking OrganizationRelationship [OR] configuration"
        $onpremOR = Get-OrganizationRelationship | Where-Object {$_.DomainNames -eq $remoteRoutingDomain}
        if($onpremOR.Enabled -eq $False)
        {
            WriteLog -color Red "Error: OrganizationRelationship is disabled. As IOC is also disabled for the RemoteDomain, please make sure to enable OrganizationRelationship to make Free/Busy working"
        }
        if($onpremOR.Enabled -eq $TRUE)
        {
            if($onpremOR.FreeBusyAccessEnabled -eq $False -or $NULL -eq $onpremOR.FreeBusyAccessLevel -or $onPremOR.TargetApplicationUri -NotLike '*Outlook.com*' -or $NULL -eq $onPremOR.TargetAutodiscoverEpr)
            {
                WriteLog -color Red "Error: OrganizationRelationship configuration is Not configured correctly"
                $onPremOR | Format-list FreeBusyAccessEnabled,FreeBusyAccessLevel,TargetApplicationUri,TargetAutodiscoverEpr
                
                write-host -f cyan "Please make sure above Output is matching with the Sample output below:
                FreeBusyAccessEnabled : True
                FreeBusyAccessLevel   : LimitedDetails
                TargetApplicationUri  : outlook.com
                TargetAutodiscoverEpr : https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"
                
            }
        }
    }
}

If($directionOftheIssue -eq 'EXOToOnprem')
{
    Disconnect-ExchangeOnline -confirm:$false
    New-LocalExchangeSession
    $onpremAppURI = (Get-FederationTrust).ApplicationURI
#Onprem Mailbox calendar permission check
    $onpremMBX = Get-Mailbox $onpremRcpt
    $onpremAlias = $onpremMbx.Alias
    $calFolder = $onpremAlias + ":\Calendar"
    $onpremMbxCalPerms = Get-MailboxFolderPermission $calFolder
    foreach($perm in $onpremMbxCalPerms)
    {
        if($perm.user -like 'Default' -and $perm.accessrights -Notlike 'AvailabilityOnly')
        {
            WriteLog -Color Yellow "Warning: Please check and confirm that Default Calendar folder permission is at least set to AvailabilityOnly"
        }
    }
    New-O365Session
#Checking Onprem MBX existence in EXO
    if(!(Get-Recipient $onpremRcpt))
    {
        Writelog -color -Red "Error: Onprem email address does not exist in EXO, please make sure the object is in sync"
        EXIT
    }
#IntraOrganization Connector [IOC] configuration check
    $primaryDomain = $onpremRcpt.split('@')[1]
    $exoIOC = Get-IntraOrganizationConnector | Where-Object {$_.TargetAddressDomains -eq $PrimaryDomain}
    
    if($exoIOC.Enabled -eq $TRUE)
    {
        WriteLog -color cyan "IOC is enabled, performing OAUTH checks.."
        if($exoIOC.DiscoveryEndpoint -like '*msappproxy*')
    {
        Writelog -color cyan "Hybrid Modern is configured with Hybrid Agent."
        Writelog -color cyan "Make sure External URL set on the Hybrid Agent configuration, is accessiable. Review article https://learn.microsoft.com/en-us/exchange/hybrid-deployment/hybrid-agent for more info on Hybrid Agent"
    }
        
        $oauthResult = Test-OAUTHConnectivity -Service AutoD -TargetUri $exoIOC.DiscoveryEndpoint -Mailbox $exoRcpt
        $oauthResult | Export-Clixml .\$FolderName\EXO_OAuthResult.xml
        if($oauthResult.ResultType -Like '*Fail*')
        {    
            WriteLog -color Red "Error: EXO OAUTH Test failed, please review EXO_OAUTHResult.xml for more information"
        }
    }        
    if($exoIOC.Enabled -eq $False)
    {
        WriteLog -color cyan "Checking OrganizationRelationship [OR] configuration"
        $exoOR = Get-OrganizationRelationship | Where-Object {$_.DomainNames -eq $primaryDomain}
        if($exoOR.Enabled -eq $False)
        {
            WriteLog -color Red "Error: OrganizationRelationship is disabled. As IOC is also disabled for the RemoteDomain, please make sure to enable OrganizationRelationship to make Free/Busy working"
        }
        if($exoOR.Enabled -eq $TRUE)
        {
            if($exoOR.TargetAutodiscoverEpr -like '*MSAppProxy*')
            {
                Writelog -color cyan "Hybrid Modern is configured with Hybrid Agent."
                Writelog -color cyan "Make sure External URL set on the Hybrid Agent configuration, is accessiable. Review article https://learn.microsoft.com/en-us/exchange/hybrid-deployment/hybrid-agent for more info on Hybrid Agent"
    
            }
            if($exoOR.FreeBusyAccessEnabled -eq $False -or $NULL -eq $exoOR.FreeBusyAccessLevel -or $exoOR.TargetApplicationUri -notlike $onpremAppURI -or $NULL -eq $exoOR.TargetAutodiscoverEpr)
            {
                WriteLog -color Red "Error: OrganizationRelationship configuration is Not configured correctly"
                $exoOR | Format-list FreeBusyAccessEnabled,FreeBusyAccessLevel,TargetApplicationUri,TargetAutodiscoverEpr
                
                write-host -f cyan "Please make sure above Output is matching with the Sample output below:
                FreeBusyAccessEnabled : True
                FreeBusyAccessLevel   : LimitedDetails
                TargetApplicationUri  : $onpremAppURI
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
