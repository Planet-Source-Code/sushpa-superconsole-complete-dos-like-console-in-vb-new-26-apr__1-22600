VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SuperConsole"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "aboutbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmClose 
      Caption         =   "Clo&se"
      Default         =   -1  'True
      Height          =   420
      Left            =   3465
      TabIndex        =   3
      Top             =   1710
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   225
      Picture         =   "aboutbox.frx":000C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   135
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "http://sushantshome.tripod.com"
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   1845
      Width           =   2340
   End
   Begin VB.Label Label2 
      Caption         =   $"aboutbox.frx":0316
      Height          =   870
      Left            =   180
      TabIndex        =   1
      Top             =   720
      Width           =   4290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SuperConsole"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   945
      TabIndex        =   0
      Top             =   135
      Width           =   2790
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub CmClose_Click()
Unload Me
End Sub
