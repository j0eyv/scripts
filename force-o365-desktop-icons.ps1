if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk")){  

if(Test-Path -Path "C:\Program Files\Microsoft Office 15\ClientX64\OfficeClickToRun.exe"){

#Restore Shortcuts to Public desktop
#    $ComObj = New-Object -ComObject WScript.Shell
#    $ShortCut = $ComObj.CreateShortcut("C:\Users\Public\desktop\Excel.lnk")
#    $ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
#    $ShortCut.Description = "Excel"
#    $ShortCut.FullName 
#    $ShortCut.WindowStyle = 1
#     $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\Excel.exe, 0";
#    $ShortCut.Save()#

#    $ComObj = New-Object -ComObject WScript.Shell
#    $ShortCut = $ComObj.CreateShortcut("C:\Users\Public\desktop\Outlook.lnk")
#    $ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\Outlook.exe"
#    $ShortCut.Description = "Outlook"
#    $ShortCut.FullName 
#    $ShortCut.WindowStyle = 1
#    $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\Outlook.exe, 0";
#    $ShortCut.Save()

#    $ComObj = New-Object -ComObject WScript.Shell
#    $ShortCut = $ComObj.CreateShortcut("C:\Users\Public\desktop\Word.lnk")
#    $ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\Winword.EXE"
#    $ShortCut.Description = "Word"
#    $ShortCut.FullName 
#    $ShortCut.WindowStyle = 1
#    $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\Winword.exe, 0";
#    $ShortCut.Save()

#    $ComObj = New-Object -ComObject WScript.Shell
#    $ShortCut = $ComObj.CreateShortcut("C:\Users\Public\desktop\Powerpoint.lnk")
#    $ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\PowerPNT.exe"
#    $ShortCut.Description = "PowerPoint"
#    $ShortCut.FullName 
#    $ShortCut.WindowStyle = 1
#    $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\PowerPNT.exe, 0";
#    $ShortCut.Save()

		if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Powerpoint.lnk")
		$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\PowerPNT.exe"
		$ShortCut.Description = "PowerPoint"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
        $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\PowerPNT.exe, 0";
		$ShortCut.Save()
	}

	if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Word.lnk")
		$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\Winword.EXE"
		$ShortCut.Description = "Word"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
        $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\Winword.exe, 0";
		$ShortCut.Save()
	}

	if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk")
		$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\Outlook.exe"
		$ShortCut.Description = "Outlook"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
        $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\Outlook.exe, 0";
		$ShortCut.Save()
	}

	if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")
		$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
		$ShortCut.Description = "Excel"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
        $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\Excel.exe, 0";
		$ShortCut.Save()
	}



} elseif(Test-Path -Path "C:\Program Files\Microsoft Office 15\ClientX32\OfficeClickToRun.exe"){

#Restore Shortcuts to Public desktop
#    $ComObj = New-Object -ComObject WScript.Shell
#    $ShortCut = $ComObj.CreateShortcut("C:\Users\Public\desktop\Excel.lnk")
#    $ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
#    $ShortCut.Description = "Excel"
#    $ShortCut.FullName 
#    $ShortCut.WindowStyle = 1
#    $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\Excel.exe, 0";
#    $ShortCut.Save()
#    $ComObj = New-Object -ComObject WScript.Shell
#    $ShortCut = $ComObj.CreateShortcut("C:\Users\Public\desktop\Outlook.lnk")
#    $ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\Outlook.exe"
#    $ShortCut.Description = "Outlook"
#    $ShortCut.FullName 
#    $ShortCut.WindowStyle = 1
#    $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\Outlook.exe, 0";
#    $ShortCut.Save()
#    $ComObj = New-Object -ComObject WScript.Shell
#    $ShortCut = $ComObj.CreateShortcut("C:\Users\Public\desktop\Word.lnk")
#    $ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\Winword.EXE"
#    $ShortCut.Description = "Word"
#    $ShortCut.FullName 
#    $ShortCut.WindowStyle = 1
#    $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\Winword.exe, 0";
#    $ShortCut.Save()
#    $ComObj = New-Object -ComObject WScript.Shell
#    $ShortCut = $ComObj.CreateShortcut("C:\Users\Public\desktop\Powerpoint.lnk")
#    $ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\PowerPNT.exe"
#    $ShortCut.Description = "PowerPoint"
#    $ShortCut.FullName 
#    $ShortCut.WindowStyle = 1
#    $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\PowerPNT.exe, 0";
#    $ShortCut.Save()
	
		if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk")
		$ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\PowerPNT.exe"
		$ShortCut.Description = "PowerPoint"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
        $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\PowerPNT.exe, 0";
		$ShortCut.Save()
	}

	if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk")
		$ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\Winword.EXE"
		$ShortCut.Description = "Word"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
        $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\Winword.exe, 0";
		$ShortCut.Save()
	}

	if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk")
		$ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\Outlook.exe"
		$ShortCut.Description = "Outlook"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
        $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\Outlook.exe, 0";
		$ShortCut.Save()
	}

	if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")
		$ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
		$ShortCut.Description = "Excel"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
        $ShortCut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\Excel.exe, 0";
		$ShortCut.Save()
	}



}
}else{ 
write-host "nothing to repair"}

#Restore Other Shortcuts to Public desktop

if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk")){  
 $ComObj = New-Object -ComObject WScript.Shell
    $ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk")
    $ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    $ShortCut.Description = "Edge"
    $ShortCut.FullName 
    $ShortCut.WindowStyle = 1
    $ShortCut.IconLocation = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe, 0";
    $ShortCut.Save()

    $ComObj = New-Object -ComObject WScript.Shell
    $ShortCut = $ComObj.CreateShortcut("C:\Users\Public\desktop\Microsoft Edge.lnk")
    $ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    $ShortCut.Description = "Edge"
    $ShortCut.FullName 
    $ShortCut.WindowStyle = 1
    $ShortCut.IconLocation = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe, 0";
    $ShortCut.Save()
}