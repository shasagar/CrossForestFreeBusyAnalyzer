# CrossForestFreeBusyAnalyzer
## The purpuse of this code is to perform basic free busy configuration checks in hybrid environment:
  * Intra Organization Connector configuration 
  * User configuration at both the environment 
  * Organization Relationship configuration 
  * Hybrid Agent configuration 
  * OAUTH configuration
## Execution Requirements:
  * Windows PowerShell with Elevated Privileges
  * Exchange Onpremise Admin account
  * Exchange Online Global Admin account
## Required Inputs:
  * One of the exchange onpremise users' email address
  * One of the exchange online users' email address
  * Exchange Online Routing Domain
    ```powershell
    Supply one of the local exchange servers' name: 
    One of the on-premise recipients' email address: 
    One of the Exchange online recipients' email address: 
    Tenent Remote Routing domain (ex: contoso.mail.onmicrosoft.com): 
    In which direction free/busy is NOT working? (OnpremToEXO / EXOToOnprem):
    ``` 
## Outputs:
  * Code log file
  * A folder with all the required configurations in XML
## Process/Examples:
 * Download the script and save it locally on exchange server
 * Open Windows PowerShell as an administrator
 * Locate the directory where you have script saved
 * Execute the script.
   ```powershell
   cd .\Scripts\FreeBusyAnalyzer\
   .\FreeBusyAnalyzerv1.0.ps1
   ```
   
