<#  
.SYNOPSIS  
	This script searches through and removes users that are no longer valid or no longer have Enterprise Voice enabled from Skype for Business Response Groups


.DESCRIPTION  
	Created by James Arber. www.skype4badmin.com
	Although every effort has been made to ensure this script is free of errors, dates change and sometimes I goof. 
	Please use at your own risk.
		    
	
.NOTES  
    Version      	   	: 0.01 Devel
	Date			    : 21/02/2018
	Lync Version		: Tested against Skype4B Server 2015 and Lync Server 2013
    Author    			: James Arber
	Header stolen from  : Greig Sheridan who stole it from Pat Richard's amazing "Get-CsConnections.ps1"

	Revision History	: v0.01: Internal build
						
	Disclaimer   		: Whilst I take considerable effort to ensure this script is error free and wont harm your enviroment.
								I have no way to test every possible senario it may be used in. I provide these scripts free
								to the Lync and Skype4B community AS IS without any warranty on its appropriateness for use in
								your enviroment. I disclaim all implied warranties including,
  								without limitation, any implied warranties of merchantability or of fitness for a particular
  								purpose. The entire risk arising out of the use or performance of the sample scripts and
  								documentation remains with you. In no event shall I be liable for any damages whatsoever
  								(including, without limitation, damages for loss of business profits, business interruption,
  								loss of business information, or other pecuniary loss) arising out of the use of or inability
  								to use the script or documentation.

	Acknowledgements 	: Testing and Advice
  								Greig Sheriden https://greiginsydney.com/about/ @greiginsydney

						: Auto Update Code
								Pat Richard http://www.ehloworld.com @patrichard

						: Proxy Detection
								Michel de Rooij	http://eightwone.com

  								
.INPUTS 
    None. Invoke-CsRgsUserCleanUp.ps1 does not accept pipelined input.

.OUTPUTS
    None. Invoke-CsRgsUserCleanUp.ps1 only provides user feedback and cannot be piped.

.PARAMETER -FrontEndPool <FrontEnd FQDN> 
    Frontend Pool to perform the cleanup on 
    If you dont specify a ServiceID or FrontEndPool, the script will try and guess the frontend to put the holidays on.
    Specifiying this instead of ServiceID will cause the script to confirm the pool unless -Unattended is specified

.PARAMETER -DisableScriptUpdate
    Stops the script from checking online for an update and prompting the user to download. Ideal for scheduled tasks

.PARAMETER -RemoveExistingRules
    Deprecated. Script now updates existing rulesets rather than removing them. Kept for backwards compatability

.PARAMETER -Unattended
    Assumes yes for pool selection critera when multiple pools are present and Poolfqdn is specified.
	Also assumes any matches will be removed automatically! Make sure your backup script is running and use with caution!
    

.LINK  
    http://www.skype4badmin.com/##TODO##


.EXAMPLE

	PS C:\> Invoke-CsRgsUserCleanUp.ps1
	Attemptes to automatically locate pool and to confirm if the users are okay to be removed.

    PS C:\> Invoke-CsRgsUserCleanUp.ps1 -FrontEndPool AUMELSFBFE.Skype4badmin.local
	Finds and prompts user to remove invalid users from the AUMELSFBFE pool.

	PS C:\> Invoke-CsRgsUserCleanUp.ps1 -FrontEndPool AUMELSFBFE.Skype4badmin.local -RemoveUsers
	Finds and removes invalid users from the SFBFE01 pool.

	PS C:\> Invoke-CsRgsUserCleanUp.ps1 -FrontEndPool AUMELSFBFE.Skype4badmin.local -Unattended  
	Finds and removes all instances of invalid users in response groups hosted on AUMELSFBFE.Skype4badmin.local without prompting for anything 
	Also disables script updates

#>
# Script Config
[CmdletBinding(DefaultParametersetName="Common")]
param(
	[Parameter(Mandatory=$false, Position=1)] $FrontEndPool,
	[Parameter(Mandatory=$false, Position=2)] [switch]$DisableScriptUpdate,
    [Parameter(Mandatory=$false, Position=3)] [switch]$Unattended,
	[Parameter(Mandatory=$false, Position=4)] [switch]$RemoveUsers,
	[Parameter(Mandatory=$false, Position=5)] [switch]$AllPools,
	[Parameter(Mandatory=$false, Position=6)] [string]$LogFilePath

	)
#region config
	If (!$LogFileLocation) {$LogFileLocation = $PSCommandPath -replace ".ps1",".log"}
	[single]$Version = "0.01"
	#Todo, add a backup file to check.
#endregion config


#region Fucntions
Function Get-IEProxy {
	Write-Host "Info: Checking for proxy settings" -ForegroundColor Green
        If ( (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').ProxyEnable -ne 0) {
            $proxies = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer
            if ($proxies) {
                if ($proxies -ilike "*=*") {
                    return $proxies -replace "=", "://" -split (';') | Select-Object -First 1
                }
                Else {
                    return ('http://{0}' -f $proxies)
                }
            }
            Else {
                return $null
            }
        }
        Else {
            return $null
        }
    }

Function Write-Log {
    PARAM(
         [String]$Message,
         [String]$Path = $LogFileLocation,
         [int]$severity = 1,
         [string]$component = "Default"
         )

         $TimeZoneBias = Get-WmiObject -Query "Select Bias from Win32_TimeZone"
         $Date= Get-Date -Format "HH:mm:ss"
         $Date2= Get-Date -Format "MM-dd-yyyy"

         $MaxLogFileSizeMB = 10
         If(Test-Path $Path)
         {
            if(((gci $Path).length/1MB) -gt $MaxLogFileSizeMB) # Check the size of the log file and archive if over the limit.
            {
                $ArchLogfile = $Path.replace(".log", "_$(Get-Date -Format dd-MM-yyy_hh-mm-ss).lo_")
                ren $Path $ArchLogfile
            }
         }
         
		 "$env:ComputerName date=$([char]34)$date2$([char]34) time=$([char]34)$date$([char]34) component=$([char]34)$component$([char]34) type=$([char]34)$severity$([char]34) Message=$([char]34)$Message$([char]34)"| Out-File -FilePath $Path -Append -NoClobber -Encoding default
         #If the log entry is just informational (less than 2), output it to write verbose
		 if ($severity -le 2) {"Info: $Message"| Write-Host -ForegroundColor Green}
		 #If the log entry has a severity of 3 assume its a warning and write it to write-warning
		 if ($severity -eq 3) {"$date $Message"| Write-Warning}
		 #If the log entry has a severity of 4 or higher, assume its an error and display an error message (Note, critical errors are caught by throw statements so may not appear here)
		 if ($severity -ge 4) {"$date $Message"| Write-Error}
} 

#endregion Functions




#Define Listnames
Write-Log -component "Bootstrap" -Message "Invoke-CsRgsUserCleanUp.ps1 Version $version" -severity 1

if ($Unattended) {$DisableScriptUpdate = $true}

#Get Proxy Details
	    $ProxyURL = Get-IEProxy
    If ( $ProxyURL) {
		Write-Log -component "Bootstrap" -Message "Using proxy address $ProxyURL" -severity 1
            }
    Else {
		Write-Log -component "Bootstrap" -Message "No proxy setting detected, using direct connection" -severity 1
        }

if ($DisableScriptUpdate -eq $false) {
	Write-Host "Info: Checking for Script Update (10 seconds max)" -ForegroundColor Green #todo
    $GitHubScriptVersion = Invoke-WebRequest https://raw.githubusercontent.com/atreidae/Invoke-CsRgsUserCleanUp/devel/version -TimeoutSec 10 -Proxy $ProxyURL ##TODO branch
        If ($GitHubScriptVersion.Content.length -eq 0) {

            Write-Warning "Error checking for new version. You can check manualy here"
            Write-Warning "http://www.skype4badmin.com/australian-holiday-rulesets-for-response-group-service/" ##TODO URL
            Write-Host "Info: Pausing for 5 seconds" -ForegroundColor Green
            start-sleep 5
            }
        else { 
                if ([single]$GitHubScriptVersion.Content -gt [single]$version) {
                 Write-Host "Info: New Version Available" -ForegroundColor Green
                    #New Version available

                    #Prompt user to download
				$title = "Update Available"
				$message = "an update to this script is available, did you want to download it?"

				$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
					"Launches a browser window with the update"

				$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
					"No thanks."

				$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

				$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

				switch ($result)
					{
						0 {Write-Host "Info: User opted to download update" -ForegroundColor Green
							start "http://www.skype4badmin.com/australian-holiday-rulesets-for-response-group-service/" ##todo URL
							Write-Warning "Exiting script"
							Exit
						}
						1 {Write-Host "Info: User opted to skip update" -ForegroundColor Green
							
							}
							
					}
                 }   
                 Else{
                 Write-Host "Info: Script is up to date" -ForegroundColor Green
                 }
        
	       }

	}

Write-Log -component "Bootstrap" -Message "Importing modules" -severity 1
Import-Module Lync
Import-module SkypeForBusiness



Write-Log -component "Bootstrap" -Message "Gathering Front End Pool Data" -severity 1
$Pools = (Get-CsService -Registrar)

Write-Log -component "Bootstrap" -Message "Parsing command line parameters" -severity 1
If ($AllPools) {Write-Log -component "Bootstrap" -Message "Allpools True, Skipping pool check" -severity 1}
if (!$allpools) { 
	# Detect and deal with null service ID
	If ($FrontEndPool -eq $null) {
		Write-Log -component "Bootstrap" -Message "No Frontend Pool entered, Searching for valid Pool" -severity 3
		Write-Log -component "Bootstrap" -Message "Looking for Front End Pools" -severity 1
			$PoolNumber = ($Pools).count
			if ($PoolNumber -eq 1) { 
				Write-Log -component "Bootstrap" -Message "Only found 1 Front End Pool, $Pools.poolfqdn, Selecting it" -severity 1
				$RGSIDs = (Get-CsRgsConfiguration -Identity $pools.PoolFqdn)
				$Poolfqdn = $Pools.poolfqdn
				#Prompt user to confirm
				Write-Log -component "Bootstrap" -Message "Found RGS Service ID $RGSIDs" -severity 1
					$title = "Use this Front End Pool?"
					$message = "Use the Response Group Server on $poolfqdn ?"

					$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
						"Continues using the selected Front End Pool."

					$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
						"Aborts the script."

					$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

					$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

					switch ($result)
						{
							0 {Write-Log -component "Bootstrap" -Message "Updating ServiceID parameter" -severity 1
								$ServiceID = $RGSIDs.Identity.tostring()}
							1 {Write-Log -component "Bootstrap" -Message "Couldn't Autolocate RGS pool. Aborting script" -severity 3
								Throw "Couldn't Autolocate RGS pool. Abort script"}
							
						}

					}
			

		Else {
		#More than 1 Pool Detected and the user didnt specify anything
			Write-Log -component "Bootstrap" -Message "Found $PoolNumber Front End Pools" -severity 1
	
			If ($FrontEndPool -eq $null) {
				Write-Log -component "Bootstrap" -Message "Prompting user to select Front End Pool" -severity 1
				Write-Log -component "Bootstrap" -Message "Couldn't Locate ServiceID or PoolFQDN on the command line and more than one Front End Pool was detected" -severity 3
				#Menu code thanks to Grieg.
				#First figure out the maximum width of the pools name (for the tabular menu):
				$width=0
				foreach ($Pool in ($Pools)) {
					if ($Pool.Poolfqdn.Length -gt $width) {
						$width = $Pool.Poolfqdn.Length
					}
				}

				#Provide an on-screen menu of Front End Pools for the user to choose from:
				$index = 0
				write-host ("Index  "), ("Pool FQDN".Padright($width + 1)," "), "Site ID"
				foreach ($Pool in ($Pools)) {
					write-host ($index.ToString()).PadRight(7," "), ($Pool.Poolfqdn.Padright($width + 1)," "), $pool.siteid.ToString()
					$index++
					}
				$index--	#Undo that last increment
				Write-Host
				Write-Host "Choose the Front End Pool you wish to use"
				$chosen = read-host "Or any other value to quit"

				if ($chosen -notmatch '^\d$') {Exit}
				if ([int]$chosen -lt 0) {Exit}
				if ([int]$chosen -gt $index) {Exit}
				$FrontEndPool = $pools[$chosen].PoolFqdn
				$Poolfqdn = $FrontEndPool
				$RGSIDs = (Get-CsRgsConfiguration -Identity $FrontEndPool)
			}


		#User specified the pool at the commandline or we collected it earlier
		
		Write-Log -component "Bootstrap" -Message "Using Front End Pool $FrontendPool" -severity 1
		$RGSIDs = (Get-CsRgsConfiguration -Identity $FrontEndPool)
		$Poolfqdn = $FrontEndPool



	if (!$Unattended) {
		#Prompt user to confirm
			$title = "Use this Pool?"
			$message = "Use the Response Group Server on $poolfqdn ?"

			$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
				"Continues using the selected Front End Pool."

			$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
				"Aborts the script."

			$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

			$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

			switch ($result)
				{
					0 {Write-Log -component "Bootstrap" -Message "Updating ServiceID" -severity 1
						$ServiceID = $RGSIDs.Identity.tostring()}
					1 {Write-Log -component "Bootstrap" -Message "Couldnt Autolocate RGS pool. Abort script" -severity 3
						Throw "Couldnt Autolocate RGS pool. Abort script"}
				}
			}

		} 

	}
} #end AllPools If statement

#We should have a valid Pool FQDN by now, enumerate RGS objects.
	Write-Log -component "Process" -Message "Gathering RGS Agent Groups" -severity 1
	switch ($AllPools)
				{
					$false {Write-Log -component "Process" -Message "Pulling data from $poolFqdn" -severity 1
						$RgsGroups = (Get-CsRgsAgentGroup | where {$_.OwnerPool -eq $poolFqdn})}
					$true {Write-Log -component "Bootstrap" -Message "Pulling data from all Frontend Pools" -severity 3
						$RgsGroups = (Get-CsRgsAgentGroup)}
				}
	
	Write-Log -component "Process" -Message "Found $($RgsGroups.count) groups" -severity 1
	$totalInvalidUsers = 0
	ForEach($RgsGroup in $RgsGroups){
		$InvalidGroupUsers = 0
		Write-Log -component "Process" -Message "Checking users in group $($RgsGroup.Name)" -severity 1
		$RgsGroupUsers = $RgsGroup.AgentsByUri
		ForEach($RgsUser in $RgsGroupUsers){
			#reset the status of the query
			$UserOK = $true
			#Pull the user from Skype4B, and return the value of EnterpriseVoice into UserOK
			$UserOk = (get-csuser ($RgsUser.absoluteuri)).EnterpriseVoiceEnabled

			#If the user is invalid, Throw a warning and increment the invalid user counter
			if(!$userOk -eq $true) {
				Write-Log -component "Process" -Message "User $($RgsUser.AbsolutePath) in Group $($RgsGroup.Name) Invalid " -severity 3
				$InvalidGroupUsers++ 
				}
		
			} #End of User Loop
		Write-Log -component "Process" -Message "Group $($RgsGroup.Name)" -severity 1
		Write-Log -component "Process" -Message "Group Users: $($RgsGroup.AgentsByURI.Count)" -severity 1
		Write-Log -component "Process" -Message "Group Invalid Users: $InvalidGroupUsers" -severity 1
		$totalInvalidUsers = $InvalidGroupUsers + $InvalidGroupUsers
		
		} #End of Group Loop
         Write-Log -component "Process" -Message "Found $totalInvalidUsers invalid users total." -severity 1   
		 If (!$RemoveUsers) {
			 if ($totalInvalidUsers -ge 1) { Write-Log -component "Process" -Message "Found invalid users, run script again with -RemoveUsers switch to remove them " -severity 3 } 
			 }
