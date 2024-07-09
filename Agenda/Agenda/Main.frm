VERSION 5.00
Begin VB.Form MainWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Agenda for  Windows"
   ClientHeight    =   4815
   ClientLeft      =   900
   ClientTop       =   1710
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
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Menu FileMainMenu 
      Caption         =   "&File"
      Begin VB.Menu NewMenu 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu LoadMenu 
         Caption         =   "&Load"
         Shortcut        =   ^L
      End
      Begin VB.Menu SaveMenu 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu OptionsMenu 
         Caption         =   "&Options"
         Begin VB.Menu UsePasswordMenu 
            Caption         =   "Use &password."
            Shortcut        =   ^P
         End
         Begin VB.Menu NoMessageNotificationMenu 
            Caption         =   "No message notification."
            Shortcut        =   ^M
         End
      End
      Begin VB.Menu QuitMenu 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu InformationMainMenu 
      Caption         =   "&Information"
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains the main interface window.
Option Explicit

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap

   Me.Width = Screen.Width / 2
   Me.Height = Screen.Height / 2

   With Settings
      Me.NoMessageNotificationMenu.Checked = .NoMessageNotification
      Me.UsePasswordMenu.Checked = .UsePassword
   End With

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

'This procedure displays this program's information.
Private Sub InformationMainMenu_Click()
On Error GoTo ErrorTrap

   MsgBox App.Comments, vbInformation, ProgramInformation()

EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub

'This procedure gives the command to load a saved message if present.
Private Sub LoadMenu_Click()
On Error GoTo ErrorTrap

   MessageBox.Text = GetMessage()
   
EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub

'This procedure starts a new message after confirmation from the user.
Private Sub NewMenu_Click()
On Error GoTo ErrorTrap
Dim Choice As Long

   Choice = MsgBox("Do you want to start a new message?", vbYesNo Or vbDefaultButton2 Or vbQuestion)
   If Choice = vbYes Then MessageBox.Text = vbNullString

EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub

'This procedure enables/disables the "no message" notification.
Private Sub NoMessageNotificationMenu_Click()
On Error GoTo ErrorTrap

   NoMessageNotificationMenu.Checked = Not NoMessageNotificationMenu.Checked
   Settings.NoMessageNotification = NoMessageNotificationMenu.Checked
   
EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub

'This procedure disables/enables the usage of a password to protect a message.
Private Sub UsePasswordMenu_Click()
On Error GoTo ErrorTrap

   UsePasswordMenu.Checked = Not UsePasswordMenu.Checked
   Settings.UsePassword = UsePasswordMenu.Checked
   
EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub

'This procedure closes this program.
Private Sub QuitMenu_Click()
On Error GoTo ErrorTrap

   Unload Me

EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub

'This procedure gives the command to save a message.
Private Sub SaveMenu_Click()
On Error GoTo ErrorTrap

   PutMessage MessageBox.Text, Settings

EndRoutine:
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(AskForAction:=False) = vbIgnore Then Resume EndRoutine
End Sub


