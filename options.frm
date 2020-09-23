VERSION 5.00
Begin VB.Form Options 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.OptionButton opWHITE 
      BackColor       =   &H80000005&
      Caption         =   "Colour 2"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2745
      TabIndex        =   11
      Top             =   1935
      Value           =   -1  'True
      Width           =   960
   End
   Begin VB.OptionButton opBLACK 
      BackColor       =   &H80000008&
      Caption         =   "Colour 1"
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   2745
      TabIndex        =   10
      Top             =   1620
      Width           =   960
   End
   Begin VB.CommandButton cmNO 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   420
      Left            =   2700
      TabIndex        =   9
      Top             =   630
      Width           =   1050
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   420
      Left            =   2700
      TabIndex        =   8
      Top             =   180
      Width           =   1050
   End
   Begin VB.Frame Characters 
      Caption         =   "Characters"
      Height          =   645
      Left            =   90
      TabIndex        =   5
      Top             =   1530
      Width           =   2535
      Begin VB.CheckBox chBlink 
         Caption         =   "&Blink"
         Height          =   195
         Left            =   1440
         TabIndex        =   13
         Top             =   315
         Width           =   780
      End
      Begin VB.ComboBox cbCursor 
         Height          =   315
         ItemData        =   "options.frx":014A
         Left            =   720
         List            =   "options.frx":0154
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Curso&r:"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   270
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fonts"
      Height          =   1410
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2535
      Begin VB.ComboBox cbFWFont 
         Height          =   315
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1000
         Width           =   2355
      End
      Begin VB.ComboBox cbVWFont 
         Height          =   315
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   450
         Width           =   2355
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fix&ed-width:"
         Height          =   195
         Left            =   125
         TabIndex        =   2
         Top             =   810
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Defa&ult font:"
         Height          =   195
         Left            =   125
         TabIndex        =   1
         Top             =   250
         Width           =   930
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Colo&urs:"
      Height          =   195
      Left            =   2745
      TabIndex        =   12
      Top             =   1395
      Width           =   825
   End
End
Attribute VB_Name = "Options"
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
Dim pos As Integer

Private Sub cmNO_Click()
Unload Me
End Sub

Private Sub cmOK_Click()
SetOptions
Unload Me
End Sub

Private Sub Form_Load()
For pos = 0 To Screen.FontCount - 1
If IsFixed(Screen.Fonts(pos)) Then
cbFWFont.AddItem Screen.Fonts(pos)
End If
cbVWFont.AddItem Screen.Fonts(pos)
Next pos
cbVWFont.ListIndex = 0
cbFWFont.ListIndex = 0
cbCursor.ListIndex = 0
GetOptions
End Sub

Function IsFixed(Fontname As String) As Boolean
Font.Name = Fontname
IsFixed = (TextWidth("W") = TextWidth("."))
End Function

Sub SetOptions()
SaveValue "General", "Cursor", cbCursor.List(cbCursor.ListIndex), INIFile
SaveValue "General", "Colour", CBin(opWHITE.Value), INIFile
SaveValue "General", "FWFont", cbFWFont.List(cbFWFont.ListIndex), INIFile
SaveValue "General", "VWFont", cbVWFont.List(cbVWFont.ListIndex), INIFile
SaveValue "General", "CBlink", chBlink.Value, INIFile
End Sub

Sub GetOptions()
On Error Resume Next
cbVWFont.Text = ReadValue("General", "VWFont", INIFile, "Verdana")
cbFWFont.Text = ReadValue("General", "FWFont", INIFile, "Lucida Console")
cbCursor.Text = ReadValue("General", "Cursor", INIFile, "_")
ColourScheme = ReadValue("General", "Colour", INIFile, 1)
chBlink.Value = ReadValue("General", "CBlink", INIFile, 0)
opBLACK.Value = (ColourScheme = 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
fMAIN.GetOptions
End Sub
