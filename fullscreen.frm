VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FullScreen 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "SuperConsole"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.TextBox RTF1 
      Height          =   465
      HideSelection   =   0   'False
      Left            =   810
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1260
      Width           =   870
   End
   Begin VB.DirListBox DL 
      Height          =   315
      Left            =   1215
      TabIndex        =   4
      Top             =   4500
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.FileListBox FL 
      Height          =   4575
      Left            =   7200
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer END_TIMER 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5490
      Top             =   1980
   End
   Begin VB.Timer TCUR 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5670
      Top             =   2835
   End
   Begin VB.PictureBox PM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6795
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSWinsockLib.Winsock WSP 
      Left            =   6885
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WSA 
      Left            =   6570
      Top             =   2115
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox TX1 
      Height          =   330
      Left            =   2385
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Dummy TextBox (Invisible)"
      Top             =   765
      Width           =   2085
   End
   Begin VB.Label lblPRESS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press any key to continue..."
      Height          =   195
      Left            =   450
      TabIndex        =   3
      Top             =   0
      Width           =   2040
   End
End
Attribute VB_Name = "FullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=====================================
'SUPER CONSOLE 1.2 BY SUSHANT PANDURANGI
'=====================================
'This is a DOS-like console made in VB. It has some
'of the standard functions of DOS that can be done
'by using VB. It also has additional and more useful
'tools, like the whois & isport commands, etc.
'=====================================
'Please vote if you consider this worth it. Thanks!
'=====================================
'Sushant Pandurangi (sushant@phreaker.net)
'=====================================
'Have a look at http://sushantshome.tripod.com/vb/
'for more files, tutorials, source code, etc.(Please do
'be sure to include the ending / in the web address.)
'=====================================
Option Explicit

Private Sub END_TIMER_Timer()
'This is done because you cant unload directly after
'processing from the RTF box, it causes a crash in
'riched32.dll.
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
RTF1_KeyPress KeyAscii
End Sub

Private Sub Form_Load()
GetOptions  'get the user options or load default values
End Sub

Sub Form_Resize()
On Error Resume Next
'Arrange the RTF box so it fits snuggily
RTF1.Move 0, 0, Width, Height
PM.Move (ScaleWidth - PM.Width) / 2, (ScaleHeight - PM.Height) / 2
TX1.Move Width + 1500, Height + 1500 'out of the picture
lblPRESS.Move 0, PM.Height + PM.Top + 45
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetOptions
End
End Sub

Private Sub PM_Click()
RTF1_KeyPress 0
End Sub

Private Sub PM_KeyPress(KeyAscii As Integer)
PM.Visible = False
RTF1_KeyPress KeyAscii
End Sub

Private Sub RTF1_Change()
'Keep the SelStart at end
RTF1.SelStart = Len(RTF1.Text)
End Sub

Private Sub RTF1_GotFocus()
On Error Resume Next
TX1.SetFocus
End Sub

Private Sub RTF1_KeyPress(KeyAscii As Integer)
Dim TEMP As String
If Right(RTF1.Text, 1) = Cursor Then RTF1.Text = Left(RTF1.Text, Len(RTF1.Text) - 1)
If PM.Visible Then
PM.Visible = False
RTF1.Visible = True
Exit Sub
End If
If KeyAscii = 13 Then
'Enter key was pressed, do what is required with
'the data that has been entered by the user(s)
'Firstly put a vbNewline in the RTF then process
'the given command which is collected in 'Data'
RTF1.Text = RTF1.Text & vbNewLine
Parse Data
Data = ""
'That's all, we deal with backspace now. Observe:
ElseIf KeyAscii = 8 Then
RTF1.Text = RTF1.Text & Cursor
If Data = "" Then Exit Sub
'Exit Sub because you dont want to alter lines that
'have already been parsed. Else if it is at beginning
'of a line it will strip out the previous vbNewline and
'take you to the last line, which isn't so professional!
RTF1.Text = Left(RTF1.Text, Len(RTF1.Text) - Len(Data) - Len(Cursor)) & Left(Data, Len(Data) - 1)
'similarly remove the rightmost char from data as well
Data = Left(Data, Len(Data) - 1)
Else
'Append to data, this is not BkSp nor Enter, so its a
'number, char, or symbol etc. (ABC123<>?":|+_\=-)
Data = Data & Chr$(KeyAscii)
RTF1.Text = RTF1.Text & Chr(KeyAscii) & Cursor
End If
'put the Selstart at end of text
RTF1.SelStart = Len(RTF1.Text)
End Sub

Private Sub RTF1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
TX1.SetFocus
End Sub

Private Sub TCUR_Timer()
If RTF1.SelLength > 0 Then Exit Sub
If Right(RTF1.Text, 1) = Cursor Then
RTF1.Text = Left(RTF1.Text, Len(RTF1.Text) - 1)
Else
RTF1.Text = RTF1.Text & Cursor
RTF1.SelStart = Len(RTF1.Text) - 1
End If
End Sub

Private Sub TX1_KeyPress(KeyAscii As Integer)
RTF1_KeyPress KeyAscii
End Sub

Sub SetOptions()
If RTF1.BackColor = vbBlack Then ColourScheme = 0 Else ColourScheme = 1
With Me
SaveValue "General", "Cursor", Cursor, INIFile
SaveValue "General", "Prompt", Prompt, INIFile
SaveValue "General", "Colour", CStr(ColourScheme), INIFile
SaveValue "General", "VWFont", VWFont, INIFile
SaveValue "General", "FWFont", FWFont, INIFile
End With
End Sub

Sub GetOptions()
On Error Resume Next
Cursor = ReadValue("General", "Cursor", INIFile, "_")
Prompt = ReadValue("General", "Prompt", INIFile, ">>> ")
VWFont = ReadValue("General", "VWFont", INIFile, "Verdana")
FWFont = ReadValue("General", "FWFont", INIFile, "Fixedsys")
TCUR.Enabled = CBool(ReadValue("General", "CBlink", INIFile, 0))
RTF1.Font.Name = VWFont
If Mid$(Prompt, 2, 1) = ":" Then
'prompt contains path
ChDir Left(Prompt, Len(Prompt) - 1)
End If
ColourScheme = ReadValue("General", "Colour", INIFile, 1)
RTF1.SelStart = 0
RTF1.SelLength = Len(RTF1.Text)
If ColourScheme = 0 Then
RTF1.BackColor = vbBlack
RTF1.ForeColor = vbWhite
Else
RTF1.BackColor = vbWhite
RTF1.ForeColor = vbBlack
End If
RTF1.SelLength = 0
RTF1.Font.Bold = False
RTF1.Font.Italic = False
RTF1.Font.Underline = False
SetOptions
End Sub

Private Sub WSA_Connect()
WSA.SendData "whois " & WhoIsHost & vbCrLf
End Sub

Private Sub WSA_DataArrival(ByVal bytesTotal As Long)
WSA.GetData WhoIsData
Notify WhoIsData
Notify vbNewLine & Prompt & Cursor
End Sub

Private Sub WSA_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Notify Description & "." & vbNewLine & Prompt & Cursor
WSA.Close
End Sub

Private Sub WSP_Connect()
Notify WSP.RemotePort & " is ACTIVE." & vbNewLine & Prompt & Cursor
WSP.Close
End Sub

Private Sub WSP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Notify WSP.RemotePort & " is not active." & vbNewLine & Prompt & Cursor
WSP.Close
End Sub

