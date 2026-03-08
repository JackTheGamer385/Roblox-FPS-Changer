Option Explicit

Dim ScriptVersion
ScriptVersion = "1"

Dim Shell, Fso, Http, GitText, GitVersion, LocalPath, File, XmlText, RegEx, Match, CurrentFps, FpsValue, InputValue, HelpText, RobloxOpen

Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")
Set Http = CreateObject("MSXML2.XMLHTTP")

HelpText = "Help:" & vbCrLf & vbCrLf & _
"How to Use:" & vbCrLf & _
"1. Close Roblox before running the FPS Changer." & vbCrLf & _
"2. If Roblox is open, choose Yes to close it or Cancel to stop." & vbCrLf & _
"3. Enter the FPS you want." & vbCrLf & _
"4. Setting the FPS to 0 will make it uncapped." & vbCrLf & _
"5. After the FPS is changed, reopen Roblox." & vbCrLf & vbCrLf & _
"How It Works:" & vbCrLf & _
"The tool updates Roblox’s client settings file and changes the FramerateCap value." & vbCrLf & _
"This is a normal client-side setting and does not modify gameplay, memory, or anything unsafe." & vbCrLf & _
"It is allowed by Roblox TOS because it only adjusts a configuration value Roblox already uses." & vbCrLf & vbCrLf & _
"Support:" & vbCrLf & _
"If you need help, contact @jackthegamer385 on Discord." & vbCrLf & vbCrLf & _
"GitHub Page:" & vbCrLf & _
"https://github.com/JackTheGamer385/Roblox-FPS-Changer"

LocalPath = WScript.ScriptFullName

Http.Open "GET", "https://raw.githubusercontent.com/JackTheGamer385/Roblox-FPS-Changer/refs/heads/main/FPSChanger.vbs", False
Http.Send

If Http.Status = 200 Then
    GitText = Http.ResponseText
    Dim RegVer
    Set RegVer = New RegExp
    RegVer.Pattern = "ScriptVersion\s*=\s*""(.*?)"""
    RegVer.Global = False
    Set Match = RegVer.Execute(GitText)
    If Match.Count > 0 Then
        GitVersion = Match(0).SubMatches(0)
        If GitVersion <> ScriptVersion Then
            InputValue = InputBox("A newer version of the FPS Changer is available." & vbCrLf & _
            "Current version: " & ScriptVersion & vbCrLf & _
            "Latest version: " & GitVersion & vbCrLf & vbCrLf & _
            "Type 'yes' to update, 'cancel' to skip, or 'help' for instructions.", "Update Available")
            
            If InputValue = "" Or LCase(InputValue) = "cancel" Then
            ElseIf LCase(InputValue) = "help" Then
                MsgBox HelpText, vbInformation, "Help"
            ElseIf LCase(InputValue) = "yes" Then
                Set File = Fso.OpenTextFile(LocalPath, 2)
                File.Write GitText
                File.Close
                MsgBox "Updated successfully. Please reopen the script.", vbInformation, "Updated"
                WScript.Quit
            End If
        End If
    End If
End If

RobloxOpen = Shell.AppActivate("Roblox")

If RobloxOpen Then
    InputValue = InputBox("Roblox is currently open." & vbCrLf & _
    "Type 'yes' to close it, 'cancel' to stop, or 'help' for instructions.", "Roblox Open")
    
    If InputValue = "" Or LCase(InputValue) = "cancel" Then
        WScript.Quit
    End If
    
    If LCase(InputValue) = "help" Then
        MsgBox HelpText, vbInformation, "Help"
        WScript.Quit
    End If
    
    If LCase(InputValue) = "yes" Then
        Shell.Run "taskkill /IM RobloxPlayerBeta.exe /F", 0, True
    Else
        WScript.Quit
    End If
End If

Dim FilePath
FilePath = "C:\Users\" & CreateObject("WScript.Network").UserName & "\AppData\Local\Roblox\GlobalBasicSettings_13.xml"

If Not Fso.FileExists(FilePath) Then
    MsgBox "Settings file not found:" & vbCrLf & FilePath, 16, "Error"
    WScript.Quit
End If

Set File = Fso.OpenTextFile(FilePath, 1)
XmlText = File.ReadAll
File.Close

Set RegEx = New RegExp
RegEx.Pattern = "<int name=""FramerateCap"">(.*?)</int>"
RegEx.Global = False

Set Match = RegEx.Execute(XmlText)

If Match.Count > 0 Then
    CurrentFps = Match(0).SubMatches(0)
Else
    CurrentFps = "Unknown"
End If

FpsValue = InputBox("Your current FPS cap is: " & CurrentFps & vbCrLf & vbCrLf & _
"Enter the FPS you want to set." & vbCrLf & _
"Type 'help' for instructions or 'cancel' to stop.", "FPS Changer")

If FpsValue = "" Or LCase(FpsValue) = "cancel" Then
    WScript.Quit
End If

If LCase(FpsValue) = "help" Then
    MsgBox HelpText, vbInformation, "Help"
    WScript.Quit
End If

If Not IsNumeric(FpsValue) Then
    MsgBox "FPS must be a number.", 16, "Invalid Input"
    WScript.Quit
End If

Dim NewText
NewText = "<int name=""FramerateCap"">" & FpsValue & "</int>"

XmlText = RegEx.Replace(XmlText, NewText)

Set File = Fso.OpenTextFile(FilePath, 2)
File.Write XmlText
File.Close

InputValue = InputBox("FPS changed to " & FpsValue & "." & vbCrLf & _
"Type 'help' for instructions or press OK to finish.", "Done")

If LCase(InputValue) = "help" Then
    MsgBox HelpText, vbInformation, "Help"
End If
