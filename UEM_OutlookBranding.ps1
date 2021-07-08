<#
	.DESCRIPTION
		This script sets UEM settings with PowerShell.
		When executed under SYSTEM authority a scheduled task is created to ensure recurring or once script execution on each user logon.
	.NOTES
        BASE Author: Nicola Suter, nicolonsky tech: https://tech.nicolonsky.ch
        Modified to UEM by Insign.it
#>

[CmdletBinding()]
Param()

###########################################################################################
# Start transcript for logging															  #
###########################################################################################

Start-Transcript -Path $(Join-Path $env:temp "UEM_OutlookBranding.log")

###########################################################################################
# Helper function to determine a users group membership									  #
###########################################################################################

function Get-ADGroupMembership {
	param(
		[parameter(Mandatory=$true)]
		[string]$UserPrincipalName
	)
	process{

		try{

			$Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
			$Searcher.Filter = "(&(userprincipalname=$UserPrincipalName))"
			$Searcher.SearchRoot = "LDAP://$env:USERDNSDOMAIN"
			$DistinguishedName = $Searcher.FindOne().Properties.distinguishedname
			$Searcher.Filter = "(member:1.2.840.113556.1.4.1941:=$DistinguishedName)"
			
			[void]$Searcher.PropertiesToLoad.Add("name")
			
			$List = [System.Collections.Generic.List[String]]@()

			$Results = $Searcher.FindAll()
			
			foreach ($Result in $Results) {
				$ResultItem = $Result.Properties
				[void]$List.add($ResultItem.name)
			}
		
			$List

		}catch{
			#Nothing we can do
			Write-Warning $_.Exception.Message
		}
	}
}

###########################################################################################
# Get current group membership for the group filter capabilities			            			  #
###########################################################################################

if ($driveMappingConfig.GroupFilter){
	try{
		#check if running as user and not system
		if (-not ($(whoami -user) -match "S-1-5-18")){

			$groupMemberships = Get-ADGroupMembership -UserPrincipalName $(whoami -upn)
		}
	}catch{
		#nothing we can do
	}	 
}
###########################################################################################
# UEM CODE														                                                    #
###########################################################################################

$ValueSimple = "3c,00,00,00,1f,00,00,f8,00,00,00,40,c8,00,00,00,00,00,00,00,00,00,00,ff,00,22,41,72,69,61,6c,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00"
$ValueComplex = "3c,68,74,6d,6c,3e,0d,0a,0d,0a,3c,68,65,61,64,3e,0d,0a,3c,73,74,79,6c,65,3e,0d,0a,0d,0a,20,2f,2a,20,53,74,79,6c,65,20,44,65,66,69,6e,69,74,69,6f,6e,73,20,2a,2f,0d,0a,20,73,70,61,6e,2e,50,65,72,73,6f,6e,61,6c,43,6f,6d,70,6f,73,65,53,74,79,6c,65,0d,0a,09,7b,6d,73,6f,2d,73,74,79,6c,65,2d,6e,61,6d,65,3a,22,50,65,72,73,6f,6e,61,6c,20,43,6f,6d,70,6f,73,65,20,53,74,79,6c,65,22,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,74,79,70,65,3a,70,65,72,73,6f,6e,61,6c,2d,63,6f,6d,70,6f,73,65,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,6e,6f,73,68,6f,77,3a,79,65,73,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,75,6e,68,69,64,65,3a,6e,6f,3b,0d,0a,09,6d,73,6f,2d,61,6e,73,69,2d,66,6f,6e,74,2d,73,69,7a,65,3a,31,30,2e,30,70,74,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,66,6f,6e,74,2d,73,69,7a,65,3a,31,31,2e,30,70,74,3b,0d,0a,09,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22,41,72,69,61,6c,22,2c,73,61,6e,73,2d,73,65,72,69,66,3b,0d,0a,09,6d,73,6f,2d,61,73,63,69,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,41,72,69,61,6c,3b,0d,0a,09,6d,73,6f,2d,68,61,6e,73,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,41,72,69,61,6c,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22,54,69,6d,65,73,20,4e,65,77,20,52,6f,6d,61,6e,22,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,74,68,65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,62,69,64,69,3b,0d,0a,09,63,6f,6c,6f,72,3a,77,69,6e,64,6f,77,74,65,78,74,3b,7d,0d,0a,2d,2d,3e,0d,0a,3c,2f,73,74,79,6c,65,3e,0d,0a,3c,2f,68,65,61,64,3e,0d,0a,0d,0a,3c,2f,68,74,6d,6c,3e,0d,0a"
$registryPath = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\MailSettings'
$Name1Simple = "ComposeFontSimple"
$Name1Complex = "ComposeFontComplex"
$Name2Simple = "ReplyFontSimple"
$Name2Complex = "ReplyFontComplex"
$Name3Simple = "TextFontSimple"
$Name3Complex = "TextFontComplex"

$hexSimple = $ValueSimple.Split(',') | % { "0x$_"}
$hexComplex = $ValueComplex.Split(',') | % { "0x$_"}

IF(!(Test-Path $registryPath))
  {
    New-Item -Path $registryPath -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name $name1simple -Value ([byte[]]$hexsimple) -PropertyType Binary -Force
    New-ItemProperty -Path $registryPath -Name $name2simple -Value ([byte[]]$hexsimple) -PropertyType Binary -Force
    New-ItemProperty -Path $registryPath -Name $name3simple -Value ([byte[]]$hexsimple) -PropertyType Binary -Force
    New-ItemProperty -Path $registryPath -Name $name1complex -Value ([byte[]]$hexcomplex) -PropertyType Binary -Force
    New-ItemProperty -Path $registryPath -Name $name2complex -Value ([byte[]]$hexcomplex) -PropertyType Binary -Force
    New-ItemProperty -Path $registryPath -Name $name3complex -Value ([byte[]]$hexcomplex) -PropertyType Binary -Force
    }

ELSE {
    Set-ItemProperty -Path $registryPath -Name $name1simple -Value ([byte[]]$hexsimple) -Force
    Set-ItemProperty -Path $registryPath -Name $name2simple -Value ([byte[]]$hexsimple) -Force
    Set-ItemProperty -Path $registryPath -Name $name3simple -Value ([byte[]]$hexsimple) -Force
    Set-ItemProperty -Path $registryPath -Name $name1complex -Value ([byte[]]$hexcomplex) -Force
    Set-ItemProperty -Path $registryPath -Name $name2complex -Value ([byte[]]$hexcomplex) -Force
    Set-ItemProperty -Path $registryPath -Name $name3complex -Value ([byte[]]$hexcomplex) -Force
    }



###########################################################################################
# End & finish transcript														                                		  #
###########################################################################################

Stop-transcript

###########################################################################################
# Done																				                                        	  #
###########################################################################################

#!SCHTASKCOMESHERE!#

###########################################################################################
# If this script is running under system (IME) scheduled task is created  (recurring)	    #
###########################################################################################

Start-Transcript -Path $(Join-Path -Path $env:temp -ChildPath "UEM_OutlookBranding.log")

if ($(whoami -user) -match "S-1-5-18"){

	Write-Output "Running as System --> creating scheduled task which will run on user logon"

	###########################################################################################
	# Get the current script path and content and save it to the client					          	  #
	###########################################################################################

	$currentScript= Get-Content -Path $($PSCommandPath)
	
	$schtaskScript=$currentScript[(0) .. ($currentScript.IndexOf("#!SCHTASKCOMESHERE!#") -1)]

	$scriptSavePath=$(Join-Path -Path $env:ProgramData -ChildPath "UEM_OutlookBranding")

	if (-not (Test-Path $scriptSavePath)){

		New-Item -ItemType Directory -Path $scriptSavePath -Force
	}

	$scriptSavePathName="UEM_OutlookBrandingReg.ps1"

	$scriptPath= $(Join-Path -Path $scriptSavePath -ChildPath $scriptSavePathName)

	$schtaskScript | Out-File -FilePath $scriptPath -Force

	###########################################################################################
	# Create dummy vbscript to hide PowerShell Window popping up at logon				          	  #
	###########################################################################################

	$vbsDummyScript = "
	Dim shell,fso,file
	Set shell=CreateObject(`"WScript.Shell`")
	Set fso=CreateObject(`"Scripting.FileSystemObject`")
	strPath=WScript.Arguments.Item(0)
	If fso.FileExists(strPath) Then
		set file=fso.GetFile(strPath)
		strCMD=`"powershell -nologo -executionpolicy ByPass -command `" & Chr(34) & `"&{`" &_ 
		file.ShortPath & `"}`" & Chr(34) 
		shell.Run strCMD,0
	End If
	"

	$scriptSavePathName="UEM-OutlookBrandingVBSHelper.vbs"

	$dummyScriptPath= $(Join-Path -Path $scriptSavePath -ChildPath $scriptSavePathName)
	
	$vbsDummyScript | Out-File -FilePath $dummyScriptPath -Force

	$wscriptPath = Join-Path $env:SystemRoot -ChildPath "System32\wscript.exe"

	###########################################################################################
	# Register a scheduled task to run for all users and execute the script on logon	    	  #
	###########################################################################################

	$schtaskName= "UEM-OutlookBrandingTasks"
	$schtaskDescription="UEM task envoker"

	$trigger = New-ScheduledTaskTrigger -AtLogOn
	#Execute task in users context
	$principal= New-ScheduledTaskPrincipal -GroupId "S-1-5-32-545" -Id "Author"
	#call the vbscript helper and pass the PosH script as argument
	$action = New-ScheduledTaskAction -Execute $wscriptPath -Argument "`"$dummyScriptPath`" `"$scriptPath`""
	$settings= New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
	
	$null=Register-ScheduledTask -TaskName $schtaskName -Trigger $trigger -Action $action  -Principal $principal -Settings $settings -Description $schtaskDescription -Force

	Start-ScheduledTask -TaskName $schtaskName
}

Stop-Transcript

###########################################################################################
# Done																					                                          #
###########################################################################################
