Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

Private Const MAX_PASSWORD_LENGTH As Long = 255   'Defines the maximum password length.

'This procedure is executed when this program is started.
Private Sub Main()
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
   Loop
   
   SaveSettings Settings
   
EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub
'This procedure gives the command to save a message and requests a password for it if necessary.
Public Sub PutMessage(MessageText As String, Settings As SettingsStr)
On Error GoTo ErrorTrap
Dim Password As String

   If Settings.UsePassword Then
      Password = Crypt(InputBox$("Enter password:"))
      
      If Password = vbNullString Then
         MsgBox "Specify a password.", vbExclamation
      Else
         If Len(Password) > MAX_PASSWORD_LENGTH Then
            MsgBox "Password is too long.", vbExclamation
         Else
            SaveMessage Crypt(MessageText), Password
         End If
      End If
   Else
      SaveMessage MessageText, Password:=vbNullString
   End If
   
EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub




'This procedure saves the specified message and optional password.
Private Sub SaveMessage(MessageText As String, Password As String)
On Error GoTo ErrorTrap
Dim FileH As Integer

   Screen.MousePointer = vbHourglass
            
   FileH = FreeFile()
   Open "Agenda.dat" For Output Lock Read Write As FileH
      Print #FileH, Chr$(Len(Password));
      Print #FileH, Password;
      Print #FileH, MessageText;
   Close FileH
   
   Password = vbNullString
   
EndRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub

'This procedure saves this program's settings.
Private Sub SaveSettings(Settings As SettingsStr)
On Error GoTo ErrorTrap
Dim FileH As Integer

   FileH = FreeFile()
   Open "Agenda.set" For Output Lock Read Write As FileH
      With Settings
         Print #FileH, Chr$(Abs(.NoMessageNotification));
         Print #FileH, Chr$(Abs(.UsePassword));
      End With
   Close FileH
   
EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub
