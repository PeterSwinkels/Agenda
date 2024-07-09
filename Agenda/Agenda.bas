Attribute VB_Name = "AgendaModule"
'This module contains procedures shared between the Agenda and Reader programs.
Option Explicit

Private Type MessageStr
   Password As String
   Message As String
End Type

Public Type SettingsStr
   UsePassword As Boolean
   NoMessageNotification As Boolean
End Type

Public Settings As SettingsStr

'This procedure encrypts/decrypts the specified text and returns the result.
Public Function Crypt(Text As String) As String
On Error GoTo ErrorTrap
Dim NewText As String
Dim Position As Long

   NewText = Text
   For Position = 1 To Len(NewText)
      Mid$(NewText, Position, 1) = Chr$(&HFF Xor Asc(Mid$(NewText, Position, 1)))
   Next Position
   
EndRoutine:
   Crypt = NewText
   Exit Function

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Function



'This procedure gives the command to load a message, requests a password if necessary, and then returns the result.
Public Function GetMessage() As String
On Error GoTo ErrorTrap
Dim Message As MessageStr
Dim Password As String

   With Message
      Message = LoadMessage(Settings)
   
      If Len(.Password) > 0 Then
         Password = Crypt(InputBox$("Enter password:"))
            
         If Password = vbNullString Then
            .Message = vbNullString
         Else
            If Password = .Password Then
               .Message = Crypt(.Message)
            Else
               .Message = vbNullString
               MsgBox "Wrong password.", vbExclamation
            End If
         End If
         
         .Password = vbNullString
      End If
   End With
   
EndRoutine:
   Screen.MousePointer = vbDefault
   GetMessage = Message.Message
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Function
'This procedure handles any errors that occur.
Public Function HandleError(Optional AskForAction As Boolean = True) As Long
Dim Description As String
Dim ErrorCode As Long
Static Choice As Long

   If AskForAction Then
      Description = Err.Description
      ErrorCode = Err.Number
      
      On Error Resume Next
      Choice = MsgBox("Error: " & ErrorCode & " - " & Description & ".", vbAbortRetryIgnore Or vbDefaultButton2 Or vbExclamation)

      If Choice = vbAbort Then End
   End If
   
   HandleError = Choice
End Function

'This procedure loads a message and returns it.
Private Function LoadMessage(Settings As SettingsStr) As MessageStr
On Error GoTo ErrorTrap
Dim FileH As Integer
Dim Length As Long
Dim Message As MessageStr

   Screen.MousePointer = vbHourglass
   
   With Message
      If Dir$("Agenda.dat", vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = vbNullString Then
         If Settings.NoMessageNotification Then MsgBox "There is no message.", vbInformation
         
         .Password = vbNullString
         .Message = vbNullString
      Else
         FileH = FreeFile()
         Open "Agenda.dat" For Binary Lock Read Write As FileH
            Length = Asc(Input$(1, FileH))
            .Password = Input$(Length, FileH)
            .Message = Input$(LOF(FileH) - Loc(FileH), FileH)
         Close FileH
      End If
   End With
   
EndRoutine:
   Screen.MousePointer = vbDefault
   LoadMessage = Message
   Exit Function
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Function
'This procedure loads the program's settings and returns them.
Public Function LoadSettings() As SettingsStr
On Error GoTo ErrorTrap
Dim FileH As Integer
Dim NewSettings As SettingsStr

   If Not Dir$("Agenda.set", vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = vbNullString Then
      FileH = FreeFile()
      Open "Agenda.set" For Binary Lock Read Write As FileH
         With NewSettings
            .NoMessageNotification = CBool(Asc(Input$(1, FileH)))
            .UsePassword = CBool(Asc(Input$(1, FileH)))
         End With
      Close FileH
   End If
   
EndRoutine:
   LoadSettings = NewSettings
   Exit Function

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Function








'This program returns this program's information.
Public Function ProgramInformation() As String
On Error GoTo ErrorTrap
Dim Information As String

   With App
      Information = .Title & " v" & .Major & "." & .Minor & .Revision & " - by: " & .CompanyName
   End With

EndRoutine:
   ProgramInformation = Information
   Exit Function

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Function


