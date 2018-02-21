# Invoke-CsRgsUserCleanUp
This script searches through and removes users that are no longer valid or no longer have Enterprise Voice enabled from Skype for Business Response Groups
I've built this mainly for unattended use, [Greig Sheridan](https://greiginsydney.com/) has a much better interactive version [here](https://greiginsydney.com/get-invalidrgsagents-ps1/)



## DESCRIPTION  
Created by James Arber. [www.skype4badmin.com](http://www.skype4badmin.com)
    
	
## NOTES 

Version			: 0.01

Date			: 21/02/2018

Lync Version		: Tested against Skype4B 2015

Author    		: James Arber

Header stolen from  	: Greig Sheridan who stole it from Pat Richard's amazing "Get-CsConnections.ps1"

## Update History

**:v0.01: Internal Build**

	
## LINK  

## KNOWN ISSUES
   None at this stage, this is however in development code and bugs are expected

## Script Specifics

**EXAMPLE** Attemptes to automatically locate pool and to confirm if the users are okay to be removed.  
`PS C:\> .\Invoke-CsRgsUserCleanUp.ps1`

**EXAMPLE** Finds and prompts user to remove invalid users from the AUMELSFBFE pool.  
`PS C:\> Invoke-CsRgsUserCleanUp.ps1 -FrontEndPool AUMELSFBFE.Skype4badmin.local`

**EXAMPLE** Finds and removes invalid users from the AUMELSFBFE pool.  
`PS C:\> Invoke-CsRgsUserCleanUp.ps1 -FrontEndPool AUMELSFBFE.Skype4badmin.local -RemoveUsers`

**EXAMPLE** Finds and removes all instances of invalid users in response groups hosted on AUMELSFBFE.Skype4badmin.local without prompting for anything. Also disables script updates  
`PS C:\> Invoke-CsRgsUserCleanUp.ps1 -FrontEndPool AUMELSFBFE.Skype4badmin.local -Unattended`

**PARAMETER -FrontEndPool <FrontEnd FQDN>**  
Frontend Pool to perform the cleanup on. Use -AllPools to run on all FrontEnds

**PARAMETER -DisableScriptUpdate**  
Stops the script from checking online for an update and prompting the user to download. Ideal for scheduled tasks

**PARAMETER -Unattended**  
Assumes yes for pool selection critera when multiple pools are present and Poolfqdn is specified.
Also assumes any matches will be removed automatically! Make sure your backup script is running and use with caution!

**PARAMETER -AllPools**  
Performes the clean up on all FrontEnd Pools visible in the Topology

**INPUT**  
None. Invoke-CsRgsUserCleanUp.ps1 does not accept pipelined input.

**OUTPUT**  
None. Invoke-CsRgsUserCleanUp.ps1 only provides user feedback and cannot be piped.
