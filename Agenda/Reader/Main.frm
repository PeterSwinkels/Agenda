VERSION 5.00
Begin VB.Form MainWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Agenda for  Windows"
   ClientHeight    =   4815
   ClientLeft      =   1380
   ClientTop       =   1740
   ClientWidth     =   7335
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   20.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   61.125
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox MessageBox 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Menu QuitMainMenu 
      Caption         =   "&Quit"
   End
   Begin VB.Menu DeleteMainMenu 
      Caption         =   "&Delete"
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains the main interface window.
Option Explicit

'This procedure deletes the message after confirmation from the user.
Private Sub DeleteMainMenu_Click()
On Error GoTo ErrorTrap
Dim Choice As Long

   Choice = MsgBox("Do you want to delete the message?", vbYesNo Or vbDefaultButton2 Or vbQuestion)
   If Choice = vbYes Then
      MessageBox.Text = vbNullString
      Kill "Agenda.dat"
   End If

EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap

   Me.Width = Screen.Width / 2
   Me.Height = Screen.Height / 2
   Me.Caption = ProgramInformation()

   MessageBox.Text = GetMessage()
   
EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub


'This procedure adjusts this window's controls to its new size.
Private Sub Form_Resize()
On Error Resume Next

   MessageBox.Width = Me.ScaleWidth
   MessageBox.Height = Me.ScaleHeight
End Sub


'This procedure closes this program.
Private Sub QuitMainMenu_Click()
On Error GoTo ErrorTrap

   Unload Me

EndRoutine:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub

