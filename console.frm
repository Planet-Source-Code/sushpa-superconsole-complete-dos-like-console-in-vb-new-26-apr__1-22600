VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Console 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   "SuperConsole 1.2 By Sushant Pandurangi"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "console.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox RTF1 
      Height          =   1410
      HideSelection   =   0   'False
      Left            =   90
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   765
      Width           =   2175
   End
   Begin VB.DirListBox DL 
      Appearance      =   0  'Flat
      Height          =   990
      Left            =   6795
      TabIndex        =   6
      Top             =   945
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox TX1 
      Height          =   330
      Left            =   1935
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Dummy TextBox (Invisible)"
      Top             =   765
      Width           =   2085
   End
   Begin VB.PictureBox PM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6345
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSWinsockLib.Winsock WSP 
      Left            =   6435
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WSA 
      Left            =   6120
      Top             =   2115
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList IML1 
      Left            =   6705
      Top             =   2835
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "console.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "console.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "console.frx":05C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "console.frx":08DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "console.frx":0A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "console.frx":0D56
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "console.frx":0EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "console.frx":11CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   635
      ButtonWidth     =   2037
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "IML1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear Text"
            Object.ToolTipText     =   "Clear the window"
            Object.Tag             =   "cls"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Quick Help"
            Object.ToolTipText     =   "Get quick help"
            Object.Tag             =   "?"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Customize"
            Object.Tag             =   "!"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.Toolbar TB2 
         Height          =   330
         Left            =   3825
         TabIndex        =   3
         Top             =   0
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "IML1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Copy selected text"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copy all text"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Select all text"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "About this program"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Full screen"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Timer TCUR 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5220
      Top             =   2835
   End
   Begin VB.Timer END_TIMER 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   1980
   End
   Begin VB.FileListBox FL 
      Height          =   2430
      Left            =   6750
      System          =   -1  'True
      TabIndex        =   1
      Top             =   2070
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPRESS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press any key to continue..."
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2040
   End
End
Attribute VB_Name = "Console"
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
PauseMode = False
Caption = App.Title & " " & App.Major & "." & App.Minor & " By " & App.CompanyName
INIFile = App.Path & "\settings.ini" 'configure INI
GetOptions  'get the user options or load default values
ClearWindow 'Clear the command window, show copyright
End Sub

Sub Form_Resize()
On Error Resume Next
'Arrange the RTF box so it fits snuggily
RTF1.Move 0, TB.Height, ScaleWidth, ScaleHeight - TB.Height
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
'RTF1.SelStart = Len(RTF1.Text)
End Sub

Private Sub RTF1_GotFocus()
'dont allow RTF to get focus, else it will show caret
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
If PauseMode = True Then
Notify vbNewLine & Prompt & Cursor
PauseMode = False
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
RTF1.SelStart = Len(RTF1.Text)
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
'dont allow focus on RTF
On Error Resume Next
TX1.SetFocus
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Tag
Case "cls"
ClearWindow
Case "?"
RTF1.SelStart = Len(RTF1.Text) - 1
RTF1.SelText = vbNewLine & HelpLine & Prompt 'cursor already there
Case "!"
SetOptions
Options.Show vbModal
End Select
End Sub

Private Sub TCUR_Timer()
'blinking cursor
If RTF1.SelLength > 0 Then Exit Sub
If Right(RTF1.Text, 1) = "|" Or Right(RTF1.Text, 1) = "_" Then
RTF1.Text = Left(RTF1.Text, Len(RTF1.Text) - 1)
Else
RTF1.Text = RTF1.Text & Cursor
RTF1.SelStart = Len(RTF1.Text) - 1
End If
End Sub

Private Sub TB2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
Clipboard.Clear
Clipboard.SetText RTF1.SelText
Case 2
Clipboard.Clear
Clipboard.SetText RTF1.Text
Case 4
RTF1.SelStart = 0
RTF1.SelLength = Len(RTF1.Text) - 1
Case 6
'about box
frmAbout.Show vbModal
Case 8
'full screen
    Dim TEMPTEXT As String
        Me.Hide
        If Right(RTF1.Text, 1) = "_" Or Right(RTF1.Text, 1) = "|" Then
        TEMPTEXT = Left(RTF1.Text, Len(RTF1.Text) - 1) & vbNewLine & Prompt & Cursor
        Else
        TEMPTEXT = RTF1.Text & vbNewLine & Prompt & Cursor
        End If
        Set fMAIN = New FullScreen
        fMAIN.RTF1.Text = TEMPTEXT
        fMAIN.Show
End Select
End Sub

Private Sub TX1_KeyPress(KeyAscii As Integer)
'pass it on
RTF1_KeyPress KeyAscii
End Sub

Sub SetOptions()
If RTF1.BackColor = vbBlack Then ColourScheme = 0 Else ColourScheme = 1
With Me
SaveValue .Name, "Width", .Width, INIFile
SaveValue .Name, "Height", .Height, INIFile
SaveValue .Name, "Left", .Left, INIFile
SaveValue .Name, "Top", .Top, INIFile
SaveValue "General", "Cursor", Cursor, INIFile
SaveValue "General", "Prompt", Prompt, INIFile
SaveValue "General", "Colour", CStr(ColourScheme), INIFile
SaveValue "General", "VWFont", VWFont, INIFile
SaveValue "General", "FWFont", FWFont, INIFile
End With
End Sub

Sub GetOptions()
On Error Resume Next
Me.Width = ReadValue(Me.Name, "Width", INIFile)
Me.Height = ReadValue(Me.Name, "Height", INIFile)
Me.Left = ReadValue(Me.Name, "Left", INIFile)
Me.Top = ReadValue(Me.Name, "Top", INIFile)
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
