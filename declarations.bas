Attribute VB_Name = "Declarations"
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
'=====================================
Public fMAIN As Form
Public Arguments() As String
Public Const HelpLine = vbNewLine & "Welcome to SuperConsole. The following is a list of standard commands:" & vbNewLine & vbNewLine & "CLS     DEL    COPY     MOVE     LIST     CD     MD     RD     DIR     NAME     !     ?     PING     WHOIS     IP     HOST     PORT     SCRN     LOAD     HELP     PAUSE     ECHO" & vbNewLine & vbNewLine & "Also, you can enter the filename or pathname of an executable file and run it from within the console. You can specify other (non-executable) files with their extensions and open them in their associated program." & vbNewLine & vbNewLine
Public ColourScheme As Integer
Public WhoIsData As String, WhoIsHost As String
Public PingStat As Long
Dim FileDate As String
Public Const NbSp = " "
Public Const Comma = ","
Public INIFile   As String
Public FWFont As String, VWFont As String
Public Cursor As String
Public PauseMode As Boolean
Public Data As String, Prompt As String
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Dim pos As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PASTE = &H302
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301

Sub Parse(CommandLine As String)
On Error GoTo handler 'see what happens
Dim InStrPos As Long, Result As Long, CMD As String, ARGS As String
InStrPos = InStr(1, CommandLine, NbSp) 'find where first space is
If InStrPos = 0 Then CMD = CommandLine Else CMD = Left(CommandLine, InStrPos)
'in case there is no space, CMD that is the command part without the arguments is
'the same as the entire command else we separate the part after the first space.
CMD = Trim(CMD) 'remove spaces
If CMD = "" Then 'nothing is given
fMAIN.RTF1.Text = fMAIN.RTF1.Text & Prompt & Cursor
Exit Sub 'nothing to do now
End If
'the below line gets the arguments part.
ARGS = Right(CommandLine, Len(CommandLine) - Len(CMD))
'Based on the command, this will process the input.
'On error it is set up to go to either label invalid or
'noargs which is appropriate according to the error.
SplitStr ARGS, Arguments(), Comma
'ABOVE:SplitStr is an alternative to the VB6 split function.
'It will split the arguments given for specific commands.
Select Case LCase(CMD)
Case "cls"
    'clear window
    ClearWindow
Case "echo"
    'echo
    Notify ARGS & vbNewLine & Prompt & Cursor
Case "pause"
PauseMode = True
Notify "Press any key to continue..."
Case "port"
    'find if port is active
    IsPort CLng(Arguments(1)), Arguments(0)
Case "ip"
    'find ip address
    Dim IPS
    IPS = GetIPAddress(Arguments(0))
        If IPS = "" Then
            Notify "Could not get IP address." & vbNewLine & Prompt & Cursor
        Else
            Notify IPS & vbNewLine & Prompt & Cursor
        End If
Case "host"
    'find hostname
    Dim HS
    HS = GetHostFromIP(Arguments(0))
        If HS = "" Then
            Notify "Could not get Hostname." & vbNewLine & Prompt & Cursor
        Else
            Notify HS & vbNewLine & Prompt & Cursor
        End If
Case "ping"
    'ping
    If Arguments(0) = "" Then Arguments(0) = GetIPAddress()
    If DoPing(Arguments(0), PingStat) = True Then
        Notify "Pinged " & Arguments(0) & " in " & PingStat & " MS." & vbNewLine & Prompt & Cursor
    Else
        Notify "Could not ping " & Arguments(0) & "." & vbNewLine & Prompt & Cursor
    End If
Case "load"
    If Right(Arguments(0), 3) = "jpg" Or Right(Arguments(0), 3) = "gif" Or Right(Arguments(0), 3) = "bmp" Then
    fMAIN.PM.Picture = LoadPicture(Arguments(0))
    fMAIN.PM.Visible = True
    fMAIN.RTF1.Visible = False
    fMAIN.RTF1.Text = fMAIN.RTF1.Text & Prompt & Cursor
    fMAIN.Form_Resize
    Else
    Open Arguments(0) For Input As #1
    fMAIN.RTF1.Text = fMAIN.RTF1.Text & vbNewLine & Input(LOF(1), 1) & vbNewLine & Prompt & Cursor
    Close #1
    End If
Case "scrn"
'fullscreen
Dim TEMPTEXT As String
    If TypeOf fMAIN Is Console Then
        fMAIN.Hide
        If Right(fMAIN.RTF1.Text, 1) = "_" Or Right(fMAIN.RTF1.Text, 1) = "|" Then
        TEMPTEXT = Left(fMAIN.RTF1.Text, Len(fMAIN.RTF1.Text) - 1) & Prompt & Cursor
        Else
        TEMPTEXT = fMAIN.RTF1.Text & Prompt & Cursor
        End If
        Set fMAIN = New FullScreen
        fMAIN.RTF1.Text = TEMPTEXT
        fMAIN.Show
    Else
        fMAIN.Hide
        If Right(fMAIN.RTF1.Text, 1) = "_" Or Right(fMAIN.RTF1.Text, 1) = "|" Then
        TEMPTEXT = Left(fMAIN.RTF1.Text, Len(fMAIN.RTF1.Text) - 1) & Prompt & Cursor
        Else
        TEMPTEXT = fMAIN.RTF1.Text & Prompt & Cursor
        End If
        Set fMAIN = New Console
        fMAIN.RTF1.Text = TEMPTEXT
        fMAIN.Show
    End If
Case "!"
    'set options
    fMAIN.SetOptions
    Options.Show vbModal
    fMAIN.RTF1.Text = fMAIN.RTF1.Text & Prompt & Cursor
Case "whois"
    WhoIsInfo Arguments(0)
Case "?"
    Notify HelpLine & Prompt & Cursor
Case "help"
'open helpfile
    Open App.Path & "\help.txt" For Input As #1
    fMAIN.RTF1.Text = fMAIN.RTF1.Text & vbNewLine & Input(LOF(1), 1) & vbNewLine & Prompt & Cursor
    Close #1
Case "move"
    'command to move file
    fMAIN.FL.Pattern = Arguments(0)
    If Right(Arguments(1), 1) <> "\" Then Arguments(1) = Arguments(1) & "\"
        For pos = 0 To fMAIN.FL.ListCount - 1
            Result = MoveFile(fMAIN.FL.List(pos), Arguments(1) & fMAIN.FL.List(pos))
            If Result = 1 Then Notify fMAIN.FL.List(pos) & " moved." & vbNewLine Else Notify "Error moving " & fMAIN.FL.List(pos) & "." & vbNewLine
        Next pos
    Notify Prompt & Cursor
    fMAIN.FL.Pattern = "*.*"
Case "copy"
    'command to copy file
    Dim sEXIsts As Long
    fMAIN.FL.Pattern = Arguments(0)
    If Right(Arguments(1), 1) <> "\" Then Arguments(1) = Arguments(1) & "\"
        For pos = 0 To fMAIN.FL.ListCount - 1
            Result = CopyFile(fMAIN.FL.List(pos), Arguments(1) & fMAIN.FL.List(pos), sEXIsts)
            MsgBox sEXIsts
            If Result = 1 Then Notify fMAIN.FL.List(pos) & " copied." & vbNewLine Else Notify "Error copying " & fMAIN.FL.List(pos) & "." & vbNewLine
        Next pos
    Notify Prompt & Cursor
    fMAIN.FL.Pattern = "*.*"
Case "del"
    'command to delete file
    fMAIN.FL.Pattern = Arguments(0)
        For pos = 0 To fMAIN.FL.ListCount - 1
            Result = DeleteFile(fMAIN.FL.List(pos))
            If Result = 1 Then Notify fMAIN.FL.List(pos) & " deleted." & vbNewLine Else Notify "Error deleting " & fMAIN.FL.List(pos) & "." & vbNewLine
        Next pos
    fMAIN.FL.Pattern = "*.*"
    Notify Prompt & Cursor
Case "cd"
    'command to change dir
        If Arguments(0) = "$" Then Arguments(0) = App.Path
        ChDir Arguments(0)
        Prompt = CurDir() & ">"
        fMAIN.RTF1.Text = fMAIN.RTF1.Text & Prompt & Cursor
        fMAIN.FL.Path = CurDir()
Case "rd"
    'command to remove dir
        If Arguments(0) = "$" Then Arguments(0) = App.Path
        RmDir Arguments(0)
        Notify "Folder [" & Arguments(0) & "] removed."
Case "dir"
    'current dir
        Notify "Current directory is " & CurDir & vbNewLine & Prompt & Cursor
Case "md"
    'command to create dir
        If Arguments(0) = "$" Then Arguments(0) = App.Path
        MkDir Arguments(0)
        ChDir CurDir() & "\" & Arguments(0)
        Prompt = CurDir() & ">"
        fMAIN.RTF1.Text = fMAIN.RTF1.Text & Prompt & Cursor
        fMAIN.FL.Path = CurDir()
Case "name"
    'command to rename
    fMAIN.FL.Pattern = Arguments(0)
        For pos = 0 To fMAIN.FL.ListCount - 1
            Result = MoveFile(fMAIN.FL.List(pos), Arguments(1) & fMAIN.FL.List(pos))
            If Result = 1 Then Notify fMAIN.FL.List(pos) & " renamed." & vbNewLine Else Notify "Error renaming " & fMAIN.FL.List(pos) & "." & vbNewLine
        Next pos
    Notify Prompt & Cursor
    fMAIN.FL.Pattern = "*.*"
Case "quit"
    'command to say bye bye,confusing?
    fMAIN.END_TIMER.Enabled = True
Case "list"
    'command to list files
    fMAIN.DL.Path = CurDir()
    If Arguments(0) <> "" Then fMAIN.DL.Path = Arguments(0)
    For pos = 0 To fMAIN.DL.ListCount - 1
        FileDate = Space(5) & FileDateTime(fMAIN.DL.List(pos))
        fMAIN.RTF1.Text = fMAIN.RTF1.Text & "[" & GetFile(fMAIN.DL.List(pos)) & "]" & Space(5) & FileLen(fMAIN.DL.List(pos)) & " bytes " & FileDate & vbNewLine
    Next pos
    fMAIN.FL.Path = CurDir()
    If Arguments(0) <> "" Then fMAIN.FL.Pattern = Arguments(0)
    For pos = 0 To fMAIN.FL.ListCount - 1
        FileDate = Space(5) & FileDateTime(fMAIN.FL.List(pos))
        fMAIN.RTF1.Text = fMAIN.RTF1.Text & fMAIN.FL.List(pos) & Space(5) & FileLen(fMAIN.FL.List(pos)) & " bytes " & FileDate & vbNewLine
    Next pos
    fMAIN.RTF1.Text = fMAIN.RTF1.Text & "Total number of items in this folder: " & fMAIN.FL.ListCount & " files, " & fMAIN.DL.ListCount & " folders." & vbNewLine
    fMAIN.RTF1.Text = fMAIN.RTF1.Text & Prompt & Cursor
Case Else
    'unknown command, but can be file name
    Result = ShellExecute(0, "open", CMD, ARGS, CurDir(), 10)
    If Result < 32 Then
        Notify "Unable to execute command '" & CMD & "'." & vbNewLine & Prompt & Cursor
    Else
        fMAIN.RTF1.Text = fMAIN.RTF1.Text & Prompt & Cursor
    End If
End Select
fMAIN.FL.Refresh
Data = "" 'remove data
Exit Sub 'dont go to invalid and make me look like a dork
invalid: 'LABEL invalid
Notify CMD & " is unrecognized." 'invalid
NewLine 'append newline
Data = "" 'remove data
Exit Sub
noargs: 'LABEL noargs
Notify "Insufficient or no arguments for " & CMD & "." & vbNewLine & Prompt & Cursor
NewLine 'append newline
Data = "" 'remove data
Exit Sub 'skip whats next
handler: 'ERROR HANDLER
If Err.Number = 9 Then GoTo noargs Else Notify (Error & "."): Close #1
NewLine 'append newline
'number 9 is subscript out of range eg. Arguments(x) doesnt exist
End Sub

Sub Notify(WhatText As String)
fMAIN.RTF1.Text = fMAIN.RTF1.Text & WhatText
End Sub

Sub ClearWindow()
fMAIN.RTF1.Text = "SuperConsole Utility Version " & App.Major & "." & App.Minor & App.Revision & " (C) " & App.CompanyName & ", 2001." & vbNewLine & "Enter ? for quick or 'help' for detailed information on using SuperConsole." & vbNewLine & Prompt & Cursor
End Sub

Private Sub SplitStr(strMessage As String, VariableHere() As String, Char As String)
Dim intAccs As Long
Dim i As Long
Dim lngSpacePos As Long, lngStart As Long
    strMessage = Trim$(strMessage)
    lngSpacePos = 1
    lngSpacePos = InStr(lngSpacePos, strMessage, Char)
    Do While lngSpacePos
        intAccs = intAccs + 1
        lngSpacePos = InStr(lngSpacePos + 1, strMessage, Char)
    Loop
    ReDim VariableHere(intAccs)
    lngStart = 1
    For i = 0 To intAccs
        lngSpacePos = InStr(lngStart, strMessage, Char)
        If lngSpacePos Then
            VariableHere(i) = Mid(strMessage, lngStart, lngSpacePos - lngStart)
            lngStart = lngSpacePos + Len(Char)
        Else
            VariableHere(i) = Right(strMessage, Len(strMessage) - lngStart + 1)
        End If
    Next
End Sub

Sub Main()
'startup
On Error Resume Next
Dim Commands() As String, Contents As String
Set fMAIN = New Console
fMAIN.Show
If Command$() <> "" Then
ChDir App.Path
Open Command$ For Input As #1
Contents = Input(LOF(1), 1)
Close #1
SplitStr Contents, Commands(), vbNewLine
For pos = 0 To UBound(Commands) - 1
fMAIN.RTF1.SetFocus
SendKeys Commands(pos) & Chr(13)
Next pos
End If
End Sub

Function ParseLng(Expression As String) As Long
'return only the numeric portion
Dim pos As Long, TEMP As String
For pos = 1 To Len(Expression)
If IsNumeric(Mid$(Expression, pos, 1)) = True Then
TEMP = TEMP & CStr(Mid$(Expression, pos, 1))
End If
Next pos
ParseLng = CLng(TEMP)
End Function

Sub NewLine()
'add a new line, but only if not already there
If (Right(fMAIN.RTF1.Text, Len(vbNewLine & Prompt & Cursor)) <> vbNewLine & Prompt & Cursor) Then
fMAIN.RTF1.Text = fMAIN.RTF1.Text & vbNewLine & Prompt & Cursor
End If
End Sub

Function DoPing(lpHost As String, RTT As Long, Optional DataSize As Long = 32, Optional Timeout As Long = 500) As Boolean
Dim StrRTT As String, fDM As Boolean
Ping lpHost, StrRTT, fDM, DataSize, Timeout 'do the ping
DoPing = fDM 'did it match
If fDM Then RTT = ParseLng(StrRTT) 'only numeric part
End Function

Function WhoIsInfo(lpHost As String)
WhoIsHost = lpHost
fMAIN.WSA.Connect "whois.geektools.com", 43
End Function

Public Function ReadValue(Section As String, Key As String, FileName As String, Optional Default As String)
    ' Read from INI file
    Dim sReturn As String
    sReturn = String(255, Chr(0))
    ReadValue = Left(sReturn, GetPrivateProfileString(Section, Key, Default, sReturn, Len(sReturn), FileName))
End Function

Public Sub SaveValue(Section As String, Key As String, Value As String, FileName As String)
    ' Write to INI file
    WritePrivateProfileString Section, Key, Value, FileName
End Sub

Function CBin(Expression) As Integer
'convert BOOL to 1 or 0
If Expression = True Then CBin = 1 Else CBin = 0
End Function

Sub IsPort(PortNumber As Long, Hostname As String)
fMAIN.WSP.Connect Hostname, PortNumber
End Sub

Public Function Reverse(sString As String) As String
Attribute Reverse.VB_Description = "Reverse a given string."
'VB6 has this as an in-built function called
'StrReverse(String) but I am not sure of VB5.
Dim i As Integer, S As String
For i = 1 To Len(sString)
S = S & Mid(sString, Len(sString) + 1 - i, 1)
Next i
Reverse = S
End Function

Function GetFile(sPath As String) As String
Attribute GetFile.VB_Description = "Get the filename portion from the string"
    'Returns only file title
    Dim i, j As Integer
    i = InStr(1, Reverse(sPath), "\")
    GetFile = Right(sPath, i - 1)
End Function
