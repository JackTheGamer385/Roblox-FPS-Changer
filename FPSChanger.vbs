Option Explicit

Dim Shell, Fso, FilePath, XmlText, FpsValue, NewText, RegEx, File, RobloxOpen, HelpText, InputValue, CurrentFps, Match

Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

HelpText = "How To Use:" & vbCrLf & _
"1. Make sure Roblox is closed before running this tool." & vbCrLf & _
"2. If Roblox is open, type 'yes' to close it or 'cancel' to stop." & vbCrLf & _
"3. Enter the FPS you want. Set 0 for uncapped." & vbCrLf & _
"4. Open Roblox again after the FPS is changed."

RobloxOpen = Shell.AppActivate("Roblox")

If RobloxOpen Then
    InputValue = InputBox("Roblox is currently open." & vbCrLf & "Type 'yes' to close it, 'cancel' to stop, or 'help' for instructions.", "Roblox Open")
    
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

FpsValue = InputBox("Your current FPS cap is: " & CurrentFps & vbCrLf & vbCrLf & "Enter the FPS you want to set." & vbCrLf & "Type 'help' for instructions or 'cancel' to stop.", "FPS Changer")

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

NewText = "<int name=""FramerateCap"">" & FpsValue & "</int>"

XmlText = RegEx.Replace(XmlText, NewText)

Set File = Fso.OpenTextFile(FilePath, 2)
File.Write XmlText
File.Close

InputValue = InputBox("FPS changed to " & FpsValue & "." & vbCrLf & "Type 'help' for instructions or press OK to finish.", "Done")

If LCase(InputValue) = "help" Then
    MsgBox HelpText, vbInformation, "Help"
End If
