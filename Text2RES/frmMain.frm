VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Text2RES"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Add Text File(s)"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSave 
      Height          =   405
      Left            =   1200
      TabIndex        =   0
      Text            =   "MyResFile"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtLog 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1320
      Width           =   5415
   End
   Begin VB.TextBox txtStart 
      Height          =   405
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "1"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdMakeRES 
      Caption         =   "Make RES"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Caption         =   "(0 - 65535)"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Caption         =   "Save As:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblInfo 
      Caption         =   "Start ID:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Text2RES
'Paul Bahlawan
'April 29, 2011

'Make a string table resource file (.res) from text files (.txt)
'Requires: rc.exe and rcdll.dll (ships with VB6 program disc)

Option Explicit

Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

'Add contents of text file(s) to a temporary .rc file
Private Sub cmdFile_Click()
Dim x As Long
Dim ID As Long
Dim temp As String
Dim fileList() As String

    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Select text file(s)"
    CommonDialog1.Filter = "Text (*.txt)|*.txt|All (*.*)|*.*"
    
    CommonDialog1.ShowOpen
    CommonDialog1.InitDir = ""
    
    If CommonDialog1.FileName = "" Then 'No file(s) selected
        Exit Sub
    End If
    
    'Prepare list of files
    fileList = Split(CommonDialog1.FileName, vbNullChar)
    If UBound(fileList) > 0 Then 'More than 1 file selected
        For x = 1 To UBound(fileList)
            fileList(x) = fileList(0) & "\" & fileList(x)
        Next x
    Else 'Only 1 file selected
        ReDim Preserve fileList(1)
        fileList(1) = fileList(0)
    End If

    'Remove old RC file
    If txtSave.Enabled = True Then
        txtSave.Enabled = False
        If CBool(PathFileExists(App.Path & "\" & txtSave.Text & ".RC")) Then
            Kill App.Path & "\" & txtSave.Text & ".RC"
        End If
    End If
    
    On Error GoTo errout:
    
    ID = CLng(txtStart.Text)
    
    'Make/add to .RC file from .TXT file(s)
    Open App.Path & "\" & txtSave.Text & ".RC" For Append As #2
    
    For x = 1 To UBound(fileList)
        Open fileList(x) For Input As #1
    
        Print #2, "STRINGTABLE"
        Print #2, "BEGIN"
        Do While Not EOF(1)
            Line Input #1, temp
            Print #2, CStr(ID) & " """ & temp & """"
            ID = ID + 1
        Loop
        Print #2, "END"
    
        Close #1
        
        addLog "Added " & txtStart.Text & "-" & CStr(ID - 1) & " : " & Right$(fileList(x), Len(fileList(x)) - InStrRev(fileList(x), "\"))
        txtStart.Text = CStr(ID)
    Next x
    
    Close #2
    
    Exit Sub 'Normal exit
    
errout:
    Close
    addLog "ERROR adding text: " & Err.Description
End Sub

'Make the .res file from the temporary .rc file
Private Sub cmdMakeRES_Click()
Dim RCfile As String
Dim strMake As String

    RCfile = txtSave.Text & ".RC"

    'Make sure .RC file exsists
    If Not CBool(PathFileExists(App.Path & "\" & RCfile)) Then
        addLog "ERROR making RES -- " & RCfile & " Not found"
        Exit Sub
    End If

    'Set working path for rc.exe
    SetCurrentDirectory App.Path
    
    On Error GoTo errout:
    
    'Make .RES file from .RC file
    strMake = "c:\Program Files\Microsoft Visual Studio\VB98\Wizards\rc.exe " & RCfile
    Shell (strMake)
    
    addLog "Finished making " & txtSave.Text & ".RES"
    
    Exit Sub 'Normal exit
    
errout:
    addLog "ERROR making RES -- RC.EXE or RCDLL.DLL Not found"
End Sub

Private Sub cmdNew_Click()
    txtLog.Text = ""
    txtSave.Enabled = True
    txtStart.Text = "1"
End Sub

Private Sub addLog(sText As String)
    txtLog.Text = txtLog.Text & sText & vbCrLf
    txtLog.SetFocus
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub Form_Load()
    CommonDialog1.Flags = CommonDialog1.Flags Or cdlOFNAllowMultiselect Or cdlOFNExplorer
    CommonDialog1.InitDir = App.Path
End Sub
