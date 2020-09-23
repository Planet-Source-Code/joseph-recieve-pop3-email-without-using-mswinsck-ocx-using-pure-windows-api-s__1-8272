Attribute VB_Name = "popmail"
'This was programmed by  Joseph Ninan
' email - josephninan@crosswinds.net
' S4-Computer Science and engineering
' SCT College of engineering
' Trivandrum, Kerala, India
' Phone - 0091-471-594477
'Home address
'Liju Bhavan
'Muttampuram Lane
'Sreekariyam PO
'Trivandrum
'Kerala State
'India
'PIN 695017
' www.jofu.8m.com

'Add this to your form_load event()
    'ok, we have to start winsock, DUH!
    'Call StartWinsock("")
    'lets subclassing the handle
    'for the connection we are going to make
    'Call Hook(Form1.hWnd)
    
'Also this to your form terminate event
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'lets close the connection
    'Call closesocket(mysock)
    'lets unhook the hwnd so we dont
    'get an error
    'Call UnHook(Form1.hWnd)
'End Sub

Option Explicit

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = -4
Public lpPrevWndProc As Long
Public mysock As Long
Public Progress, mProgress As Integer
Public MailContent, rtncode As Variant
Public pop_error As Boolean
Public e_err, e_errstr, timeout As Variant
Public RecvBuffer, MessageDetail, MsgListDetail As String
Public ReadFlag As Boolean



Public Function Hook(ByVal hWnd As Long)
    'ok, we are going to catch ALL msg's sent
    'to the handle we are subclassing (form1)
    lpPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Function

Public Sub UnHook(ByVal hWnd As Long)
    'if we dont un-subclass before we shutdown
    'the program, we get an illigal procedure error.
    'fun.
    Call SetWindowLong(hWnd, GWL_WNDPROC, lpPrevWndProc)
End Sub
Public Function FindField(s_temp As Variant, s_field As Integer) As Variant
    Dim l, firstpos, lastpos, i, fieldcount
    s_temp = s_temp & " "
    'Removing extra spaces
    'Finding fields
    l = Len(s_temp)
    firstpos = 1
    For i = 1 To l
        If Mid(s_temp, i, 1) = " " Then
            lastpos = i
            fieldcount = fieldcount + 1
        End If
        If fieldcount = s_field Then
            FindField = Mid(s_temp, firstpos, lastpos - firstpos + 1)
            Exit Function
        Else
            firstpos = lastpos
        End If
    Next i
End Function


Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim x As Long
Dim wp As Integer
Dim temp As Variant
Dim ReadBuffer(1000) As Byte
'Debug.Print uMsg, wParam, lParam
    Select Case uMsg
        Case 1025:
            Debug.Print uMsg, wParam, lParam
            'Log uMsg & "  " & wParam & "  " & lParam
            e_err = WSAGetAsyncError(lParam)
            e_errstr = GetWSAErrorString(e_err)
            
            If e_err <> 0 Then
                Log "Error String returned -> " & e_err & " - " & e_errstr
                Log "Terminating...."
                pop_error = True
                'Exit Function
            End If
            Select Case lParam
            Case FD_READ: 'lets check for data
                    x = recv(mysock, ReadBuffer(0), 1000, 0) 'try to get some
                    If x > 0 Then 'was there any?
                        ReadFlag = False
                        RecvBuffer = StrConv(ReadBuffer, vbUnicode) 'yep, lets change it to stuff we can understand
                        Log RecvBuffer
                        rtncode = Mid(RecvBuffer, 1, 3)
                        'Log "Analysing code " & rtncode & "..."
                        Select Case rtncode
                        Case "+OK"
                            Progress = Progress + 1
                            If Progress = 5 Then
                                MsgListDetail = RecvBuffer
                            End If
                            Log ">>Progress becomes " & Progress

                        Case "-ERR"
                            pop_error = True
                        Case Else
                            If Progress = 5 Then
                                MessageDetail = RecvBuffer
                                Progress = Progress + 1
                            End If
                            If Progress = 11 Then
                                MailContent = RecvBuffer
                                Progress = Progress + 1
                            End If


                        End Select
                    End If
            Case FD_CONNECT: 'did we connect?
                    mysock = wParam 'yep, we did! yayay
                    'Log WSAGetAsyncError(lParam) & "error code"
                    'Log mysock & " - Mysocket Value"

            Case FD_CLOSE: 'uh oh. they closed the connection
                    Call closesocket(wp)   'so we need to close
            End Select
    End Select
    'let the msg get through to the form
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Public Sub Log(ByVal sText As String)
    ' this way it doesnt refresh the whole thing every time, no blinking...
    With Form1.txtStatus
        .SelStart = Len(.Text)
        .SelText = sText & Chr$(13) & Chr$(10)
        .SelLength = 0
    End With

End Sub


Public Function popconnect(m_host, m_port, m_user, m_pass As String) As Integer
'popconnect=-5  Misc
'popconnect--4   Timeout
'popconnect =-3  Invalid password
'popconnect =-2  Invalid user
'popconnect =-1  POP Mail - Sever connect met with some error
'Else popconnect returns the no of email messages in your box
Dim temp, timeout As Variant
    Progress = 0
    pop_error = False
    timeout = Timer + 60
    Log "Will timeout in 60 seconds"
    'make sure the port is closed!
    If mysock <> 0 Then Call closesocket(mysock)

    'let's connect!!!       host            port       handle
    temp = ConnectSock(m_host, m_port, 0, Form1.hWnd, True)
    Log "Connect socket return value" & temp
    Log "Connected to " & m_host & " at port " & m_port
    If temp = INVALID_SOCKET Then
        Log "Error -Invalid Socket"
        popconnect = -1
        Exit Function
    End If
    While mysock = 0  'make sure we are connected
        DoEvents
        If pop_error = True Then
            Log "Error .. No connection"
            popconnect = -1
            Exit Function
        End If
    Wend
    timeout = Timer + 60
    Log "Connection Established..."
    While Progress < 1
        DoEvents
        If pop_error = True Then
            Log "Error trying to connect to POP server"
            popconnect = -1
            Call closesocket(mysock)
            mysock = 0
            Exit Function
        End If
        If Timer > timeout Then
            Call closesocket(mysock)
            popconnect = -4
            mysock = 0
            Log "Timeout while trying to connect to server"
            Exit Function
        End If
    Wend


    Log ">USER " & m_user
    Call SendData(mysock, "USER " & m_user & vbCrLf)
    
    
    While Progress < 2
        DoEvents
        If pop_error = True Then
            Log "Invalid Username"
            popconnect = -2
            Call closesocket(mysock)
            mysock = 0
            Exit Function
        End If
        
        If Timer > timeout Then
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout after sending user info"
            popconnect = -4
            Exit Function
        End If
    Wend
    
    Log ">PASS " & m_pass
    Call SendData(mysock, "PASS " & m_pass & vbCrLf)
    While Progress < 3
        DoEvents
        If pop_error = True Then
            
            Call closesocket(mysock)
            mysock = 0
            Log "Invalid Password"
            popconnect = -3
            Exit Function
        End If
        If Timer > timeout Then
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout at progress step " & Progress
            popconnect = -4
            Exit Function
        End If
    Wend
    
    
    Log ">STAT"
    Call SendData(mysock, "STAT" & vbCrLf)
    While Progress < 4
        DoEvents
        If pop_error = True Then
            Log "Error in between getting pop details"
            popconnect = -5
            Call closesocket(mysock)
            mysock = 0
            Exit Function
        End If
        If Timer > timeout Then
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout after sending STAT"
            popconnect = -4
            Exit Function
        End If
    Wend
    Log ">LIST"
    Call SendData(mysock, "LIST" & vbCrLf)
    While Progress < 5
        DoEvents
        If pop_error = True Then
            Log "Error in between pop after LIST"
            popconnect = -5
            Exit Function
        End If
        
        If Timer > timeout Then
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout after sending" & Progress
            popconnect = -4
            Exit Function
        End If
    Wend
    'Log MsgListDetail
    While Progress < 6
        DoEvents
        If pop_error = True Then
            Log "Error in between pop after LIST"
            popconnect = -5
            Exit Function
        End If
        
        If Timer > timeout Then
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout after sending" & Progress
            popconnect = -4
            Exit Function
        End If
    Wend
    'Log MessageDetail
    popconnect = Val(FindField(MsgListDetail, 2))

End Function
Public Function getmail(mail_no As Integer, DeleteFlag As Boolean) As Variant
    Dim WholeMail, atemp
    Progress = 10
    Log mysock
    WholeMail = ""
    atemp = 100
    Progress = 10
    Log ">RETR " & mail_no
    timeout = Timer + 20
    Call SendData(mysock, "RETR " & mail_no & vbCrLf)
    While Progress < 11
        DoEvents
        If pop_error = True Then
            Log "Error in between pop "
            Exit Function
        End If
        
        If Timer > timeout Then
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout after sending" & Progress
            Exit Function
        End If
    Wend
    'While atemp = 0
    'Progress = 11
    While Progress < 12
        DoEvents
        If pop_error = True Then
            Log "Error in between pop "
            Exit Function
        End If
        
        If Timer > timeout Then
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout after sending" & Progress
            Exit Function
        End If
    Wend

    WholeMail = WholeMail & MailContent
    atemp = InStr(1, MailContent, vbCrLf & "." & vbCrLf, vbTextCompare)
    'Wend
    If DeleteFlag = True Then
        Call SendData(mysock, "DELE " & mail_no & vbCrLf)
        Progress = 20
        While Progress < 21
            DoEvents
            If pop_error = True Then
                Log "Error in between pop after DELE"
                Exit Function
            End If
        
            If Timer > timeout Then
                Call closesocket(mysock)
                mysock = 0
                Log "Timeout after sending" & Progress
                Exit Function
            End If
        Wend
    End If

    getmail = WholeMail

End Function
Public Sub PopQuit()
    Call SendData(mysock, "QUIT" & vbCrLf)
    Progress = 30
    While Progress < 31
        DoEvents
        If pop_error = True Then
            Log "Error in between pop after QUIT"
            Exit Sub
        End If
        
        If Timer > timeout Then
            Call closesocket(mysock)
            mysock = 0
            Log "Timeout after sending" & Progress
            Exit Sub
        End If
        
    Wend
    
    Call closesocket(mysock)
End Sub
