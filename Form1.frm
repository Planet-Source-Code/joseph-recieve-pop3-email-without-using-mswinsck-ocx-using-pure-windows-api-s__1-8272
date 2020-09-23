VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recieve POP email without mswinsck.ocx - by Joseph Ninan <josephninan@crosswinds.net>"
   ClientHeight    =   6405
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Form1.frx":0442
   ScaleHeight     =   6405
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDelete 
      Caption         =   "Delete Mail after recieving"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton cmdCLS 
      Caption         =   "Clear Log"
      Height          =   495
      Left            =   7800
      TabIndex        =   19
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load todays Log file"
      Height          =   495
      Left            =   6360
      TabIndex        =   17
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveLog 
      Caption         =   "Save Log"
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtm_pass 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Text            =   "jofu"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtm_user 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Text            =   "jofu"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "Recieve Mail"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtpopserver 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtport 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "110"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtStatus 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4920
      Width           =   8775
   End
   Begin VB.Label Label6 
      Caption         =   "Mail will be saved as c:\popmail{time}.eml"
      Height          =   735
      Left            =   3120
      TabIndex        =   21
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   $"Form1.frx":074C
      Height          =   1215
      Left            =   240
      TabIndex        =   20
      Top             =   2280
      Width           =   8655
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   4320
      Width           =   7095
   End
   Begin VB.Label Label12 
      Caption         =   "Rate this program. Search for Joseph Ninan at Planetsourcecode.Other codes like Permutations - Get all possible passwords"
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   3960
      TabIndex        =   16
      ToolTipText     =   "http://www.planetsourcecode.com"
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label Label11 
      Caption         =   "Visit my site"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7080
      MouseIcon       =   "Form1.frx":096E
      TabIndex        =   15
      ToolTipText     =   "http://www.jofu.8m.com"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Rate my code"
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Will be saved under file c:\poplog*.*"
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label8 
      Caption         =   $"Form1.frx":0C78
      Height          =   855
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   8535
   End
   Begin VB.Label Label7 
      Caption         =   "Debug Info......"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Rcpt"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "&Mail Server:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu menFile 
      Caption         =   "&File"
      Begin VB.Menu miExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This was programmed by  Joseph Ninan
' email - josephninan@crosswinds.net
' S4-Computer Science and engineering
' SCT College of engineering
' Trivandrum, Kerala, India
' Phone - 0091-471-594477
' www.jofu.8m.com

Private Sub cmdCLS_Click()
    Me.txtStatus.Text = ""
End Sub

Private Sub cmdLoad_Click()
    Dim afile, aLog, bLog As Variant
    Call cmdSaveLog_Click
    On Error Resume Next
    afile = "c:\poplog" & Format(Date, "dddd_mmm_d_yyyy") & ".txt"
    Open afile For Input As #5
    While Not EOF(5)
        Line Input #5, aLog
        bLog = bLog & vbCrLf & aLog
    Wend
    Me.txtStatus.Text = bLog
    Close

End Sub

Private Sub cmdSaveLog_Click()
    Dim afile As Variant
    On Error Resume Next
    afile = "c:\poplog" & Format(Date, "dddd_mmm_d_yyyy") & ".txt"
    Open afile For Append As #5
        Print #5, txtStatus.Text
        txtStatus.Text = ""
    Close
End Sub

Private Sub cmdSendMail_Click()
Dim aRes, i, subPos As Integer
Dim message As Variant
    'popconnect=-5  Misc
    'popconnect--4   Timeout
    'popconnect =-3  Invalid password
    'popconnect =-2  Invalid user
    'popconnect =-1  POP Mail - Sever connect met with some error
    'Else popconnect returns the no of email messages in your box
    aRes = popconnect(Me.txtpopserver, Me.txtport, Me.txtm_user, Me.txtm_pass)
    MsgBox ("The result of pop is " & aRes)
    If aRes < 1 Then
        MsgBox "popconnect returned a code of " & aRes
        Exit Sub
    End If
    'Now call mail=getmail(mailno,delteflag to recieve mail)
    
    For i = 1 To aRes
        If Me.chkDelete.Value = 1 Then
            message = getmail((i), True)
        Else
            message = getmail((i), False)
        End If
        message = "Message " & i & vbCrLf & message
        subPos = InStr(1, message, "Subject")
        subject = Mid(message, subPos, (InStr(subPos, message, vbCrLf) - subPos))
        MsgBox message, , subject
        afile = "C:\" & "popmail" & Format(Time, "hh_mm_ss_AMPM") & ".eml"
        Open afile For Output As #1
            Write #1, message
        Close #1
            
        'Log message
    Next i
    PopQuit

End Sub
Private Sub Form_Load()
    'Dim MyString, MyHost As String, MyReg As New cReadWriteEasyReg, i As Integer
    'ok, we have to start winsock, DUH!
    Call StartWinsock("")
    'lets subclassing the handle
    'for the connection we are going to make
    Call Hook(Form1.hWnd)
    Me.Label13.Caption = "c:\poplog" & Format(Date, "dddd_mmm_d_yyyy") & ".txt"

    '
    'This function will return a specific value from the registry

    'If Not MyReg.OpenRegistry(HKEY_CURRENT_USER, "Software\Microsoft\Internet Account Manager\Accounts\00000001") Then
    '    MsgBox "Couldn't open the registry.....Proceeding with default values.... Change them as you wish...."
    '    Exit Sub
    'End If
    'MyString = MyReg.GetValue("SMTP Email Address")
    'MyHost = MyReg.GetValue("SMTP Server")
    'Me.txtm_from.Text = MyString
    'Me.txtname_from = Me.txtname_from & "-" & MyString
    'Me.txtm_ReplyTo = MyString
    'Me.txthost.Text = MyHost
    'MyReg.CloseRegistry
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MousePointer = 1
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'lets close the connection
    Call closesocket(mysock)
    'lets unhook the hwnd so we dont
    'get an error
    Call UnHook(Form1.hWnd)
End Sub
Private Sub Label11_Click()
    Dim a
    On Error Resume Next
    a = Shell("explorer http://www.jofu.8m.com", vbNormalFocus)
End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MousePointer = 99
End Sub
Private Sub Label12_Click()
    Dim a
    On Error Resume Next
    a = Shell("explorer http://www.planetsourcecode.com", vbNormalFocus)
End Sub
Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MousePointer = 99
End Sub
Private Sub miExit_Click()
    'lets close the connection
    Call closesocket(mysock)
    'lets unhook the hwnd so we dont
    'get an error
    Call UnHook(Form1.hWnd)
    End
End Sub

Private Sub txtm_data_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MousePointer = 1

End Sub

Private Sub txtStatus_Change()
    'keep the txtbox at the very bottom at all times
    txtStatus.SelStart = Len(txtStatus)
End Sub
