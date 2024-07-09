Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'This procedure is executed when this program is started.
Public Sub Main()
On Error GoTo ErrorTrap

   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   With Settings
      .NoMessageNotification = False
      .UsePassword = False
   End With
   
   Settings = LoadSettings()
      
   MainWindow.Show

   Do While DoEvents() > 0
      If MainWindow.MessageBox.Text = vbNullString Then Unload MainWindow
   Loop
   
EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub



