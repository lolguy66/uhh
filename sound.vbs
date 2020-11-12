CheckTask

LaunchTask "Nagasaki"
LaunchTask "Hiroshima"

Sub Nagasaki()
    Dim ProgramPath, WshShell, ProgramArgs, WaitOnReturn,intWindowStyle
    Set WshShell=CreateObject ("WScript.Shell")
    Do
    Message = MsgBox("Error Could not close tab.", 5+4096+48, "Fatal Error.")
    ProgramPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
    ProgramPath = ProgramPath+"\sound.vbs"
    ProgramArgs="/hello /world"
    intWindowStyle=1
    WaitOnReturn=False
    WshShell.Run Chr (34) & ProgramPath & Chr (34) & Space (1) & ProgramArgs,intWindowStyle, WaitOnReturn
    Loop
End Sub
Sub Hiroshima()
      Set Sound = CreateObject("WMPlayer.OCX.7")
      SoundPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
    SoundPath = SoundPath+"\mmmm.mp3"
    Sound.URL = SoundPath
      Sound.Controls.play
      do while Sound.currentmedia.duration = 0
      wscript.sleep 100
      loop
      wscript.sleep (int(Sound.currentmedia.duration)+1)*1000
End Sub

Sub CheckTask()
    If WScript.Arguments.Named.Exists("task") Then
        On Error Resume Next
        Execute WScript.Arguments.Named.Item("task")
        WScript.Quit
    End If
End Sub

Sub LaunchTask(sTaskName)
    CreateObject("WScript.Shell").Exec """" & WScript.FullName & """ """ & WScript.ScriptFullName & """ ""/task:" & sTaskName & """"
End Sub
