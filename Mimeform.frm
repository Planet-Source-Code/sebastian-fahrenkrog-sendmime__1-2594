VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Sends E-Mail with Attachement!"
   ClientHeight    =   5664
   ClientLeft      =   1656
   ClientTop       =   2208
   ClientWidth     =   8184
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5664
   ScaleWidth      =   8184
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton delattach 
      Caption         =   "Del Attachement"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
   Begin VB.ListBox AttachementList 
      Height          =   432
      Left            =   4440
      TabIndex        =   14
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton Exit 
      BackColor       =   &H00808080&
      Caption         =   "Exit"
      Height          =   375
      Left            =   4200
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   5280
      Width           =   3855
   End
   Begin VB.CommandButton SendMimeConnect 
      Appearance      =   0  '2D
      BackColor       =   &H00808080&
      Caption         =   "Send"
      Height          =   375
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   8
      Top             =   5280
      Width           =   3975
   End
   Begin VB.ComboBox MailServer 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Text            =   "mail.kdt.de"
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Attachement 
      BackColor       =   &H00000000&
      Caption         =   "Add Attachement"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Tobox 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "galgen@wtal.de"
      Top             =   720
      Width           =   2175
   End
   Begin VB.ComboBox Frombox 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Text            =   "me@host.com"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Subjekt 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      MaxLength       =   78
      TabIndex        =   4
      Top             =   1560
      Width           =   7335
   End
   Begin VB.TextBox DataArrival 
      Appearance      =   0  '2D
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3960
      Width           =   7935
   End
   Begin VB.TextBox Mailtxt 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   5
      Top             =   1920
      Width           =   7935
   End
   Begin VB.Label Process 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   7935
   End
   Begin VB.Label ggg 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Server:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   105
      TabIndex        =   13
      Top             =   360
      Width           =   525
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "To:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "From:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Subject:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bTrans As Boolean
Dim m_iStage As Integer
Dim Sock As Integer
Dim RC As Integer
Dim Bytes As Integer
Dim ResponseCode As Integer
Dim path As Variant

'*****************************************
'For the Mime File Field!
'*****************************************

Private Type OPENFILENAME
       lStructSize As Long
       hwndOwner As Long
       hInstance As Long
       lpstrFilter As String
       lpstrCustomFilter As String
       nMaxCustFilter As Long
       nFilterIndex As Long
       lpstrFile As String
       nMaxFile As Long
       lpstrFileTitle As String
       nMaxFileTitle As Long
       lpstrInitialDir As String
       lpstrTitle As String
       flags As Long
       nFileOffset As Integer
       nFileExtension As Integer
       lpstrDefExt As String
       lCustData As Long
       lpfnHook As Long
       lpTemplateName As String
End Type

Const OFN_READONLY = &H1
Const OFN_OVERWRITEPROMPT = &H2
Const OFN_HIDEREADONLY = &H4
Const OFN_NOCHANGEDIR = &H8
Const OFN_SHOWHELP = &H10
Const OFN_ENABLEHOOK = &H20
Const OFN_ENABLETEMPLATE = &H40
Const OFN_ENABLETEMPLATEHANDLE = &H80
Const OFN_NOVALIDATE = &H100
Const OFN_ALLOWMULTISELECT = &H200
Const OFN_EXTENSIONDIFFERENT = &H400
Const OFN_PATHMUSTEXIST = &H800
Const OFN_FILEMUSTEXIST = &H1000
Const OFN_CREATEPROMPT = &H2000
Const OFN_SHAREAWARE = &H4000
Const OFN_NOREADONLYRETURN = &H8000
Const OFN_NOTESTFILECREATE = &H10000
Const OFN_NONETWORKBUTTON = &H20000
Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Const OFN_EXPLORER = &H80000 ' new look commdlg
Const OFN_NODEREFERENCELINKS = &H100000
Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Const OFN_SHAREFALLTHROUGH = 2
Const OFN_SHARENOWARN = 1
Const OFN_SHAREWARN = 0

Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'This is for the WaitforResponse Routine
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'Dec's for the X disabling

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Const MF_BYPOSITION = &H400&
Const MF_REMOVE = &H1000&

'For MIME processing
Dim Mime As Boolean

'For Filehandling
Dim Mimefilename As String
Dim Mimefiles As Integer


Sub DisableX(frm As Form)
     Dim hMenu As Long
     Dim nCount As Long
     hMenu = GetSystemMenu(frm.hWnd, 0)
     nCount = GetMenuItemCount(hMenu)

     'Get rid of the Close menu and its separator
     Call RemoveMenu(hMenu, nCount - 1, MF_REMOVE Or MF_BYPOSITION)
     Call RemoveMenu(hMenu, nCount - 2, MF_REMOVE Or MF_BYPOSITION)

     'Make sure the screen updates
     'our change
     DrawMenuBar frm.hWnd
End Sub

'***************************************************************
'Thanks to Luis Cantero for this Routines

Sub Startrek(frm As Form)
GotoVal = frm.Height / 2
For Gointo = 1 To GotoVal
DoEvents
frm.Height = frm.Height - 100
frm.Top = (Screen.Height - frm.Height) \ 2
If frm.Height <= 500 Then Exit For
Next Gointo
horiz:
frm.Height = 30
GotoVal = frm.Width / 2
For Gointo = 1 To GotoVal
DoEvents
frm.Width = frm.Width - 100
frm.Left = (Screen.Width - frm.Width) \ 2
If frm.Width <= 2000 Then Exit For
Next Gointo
End Sub

Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
Dim ofn As OPENFILENAME
Dim A As Long
ofn.lStructSize = Len(ofn)
ofn.hwndOwner = Form1.hWnd
ofn.hInstance = App.hInstance
If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
For A = 1 To Len(Filter)
If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
Next
ofn.lpstrFilter = Filter
ofn.lpstrFile = Space$(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = InitDir
ofn.lpstrTitle = Title
ofn.flags = OFN_HIDEREADONLY Or OFN_CREATEPROMPT
A = GetSaveFileName(ofn)
If (A) Then
SaveDialog = Left$(Trim$(ofn.lpstrFile), Len(Trim$(ofn.lpstrFile)) - 1)
Mimefilename = Left$(Trim$(ofn.lpstrFileTitle), Len(Trim$(ofn.lpstrFileTitle)) - 1)
Else
SaveDialog = ""
End If
End Function

'***************************************************************

Private Sub Attachement_Click()

Mime = True

Mimefiles = Mimefiles + 1

path = SaveDialog(Me, "*.*", "Attache file as", App.path)

Form1.AttachementList.List(Mimefiles - 1) = path

End Sub

Private Sub delattach_Click()
If Form1.AttachementList.List(AttachementList.ListIndex) <> "" Then
path = ""
Form1.AttachementList.List(AttachementList.ListIndex) = ""
Mimefiles = Mimefiles - 1
End If
End Sub

'***************************************************************
'Routine for connecting to the server
'***************************************************************

Private Sub SendMimeConnect_Click()

' Little Error check
If Tobox.Text = "" Or InStr(Tobox.Text, "@") = 0 Then
MsgBox "To: Is not correct!"
Exit Sub
End If

Dim StartupData As WSADataType
Dim SocketBuffer As sockaddr
Dim IpAddr As Long
    
'Ini the Winsocket
RC = WSAStartup(&H101, StartupData)
RC = WSAStartup(&H101, StartupData)
    

    
'Open a free Socket (with this source code you can also
'open several connections! Very useful for E-Mail Applications...)
Sock = socket(AF_INET, SOCK_STREAM, 0)
If Sock = SOCKET_ERROR Then
    Process.Caption = "Cannot Create Socket."
    Exit Sub
End If

'Checks if the Hostname exists
If RC = SOCKET_ERROR Then Exit Sub
IpAddr = GetHostByNameAlias(MailServer)
If IpAddr = -1 Then
    Process.Caption = "Unknown Host: " + MailServer
    Exit Sub
End If


'This part is responsible for the connection
SocketBuffer.sin_family = AF_INET
SocketBuffer.sin_port = htons(25)
SocketBuffer.sin_addr = IpAddr
SocketBuffer.sin_zero = String$(8, 0)
    
RC = connect(Sock, SocketBuffer, Len(SocketBuffer))

'If an error occured close the connection and
'send an error message to the text window
If RC = SOCKET_ERROR Then
        Process.Caption = "Cannot Connect to " + MailServer + _
                            Chr$(13) + Chr$(10) + _
                            GetWSAErrorString(WSAGetLastError())
        closesocket Sock
        RC = WSACleanup()
        Exit Sub
Else
Process.Caption = "Connected to " & MailServer.Text
End If

'Select Receive Window
RC = WSAAsyncSelect(Sock, DataArrival.hWnd, _
                        ByVal &H202, ByVal FD_READ Or FD_CLOSE)
    If RC = SOCKET_ERROR Then
        Process.Caption = "Cannot Process Asynchronously."
        closesocket Sock
        RC = WSACleanup()
        Exit Sub
    End If

bTrans = True
m_iStage = 0
DataArrival = ""

ResponseCode = 220
Call WaitForResponse

End Sub

Private Sub Exit_Click()
On Error Resume Next
Call Startrek(Me)

closesocket Sock
RC = WSACleanup()
End
End Sub

Private Sub Form_Load()
Call DisableX(Me)
End Sub

'***************************************************************
'Routine for arraving Data
'***************************************************************

Private Sub DataArrival_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MsgBuffer As String * 2048


    
On Error Resume Next

 

    If Sock > 0 Then
        'Receive up to 2048 chars
        Bytes = recv(Sock, ByVal MsgBuffer, 2048, 0)
        
        If Bytes > 0 Then
            
        DataArrival = DataArrival + _
                            MsgBuffer + _
                            Chr$(13) + Chr$(10)
         'Scrolls down the Textbox
         DataArrival.SelStart = Len(DataArrival)
         
        If bTrans Then
            'Checks if the Response code is correct
            If ResponseCode = Left(MsgBuffer, 3) Then
            MsgBuffer = vbNullString
            m_iStage = m_iStage + 1
            Transmit m_iStage
            Else
            'If the Response Code is not right reset the connection
                closesocket (Sock)
                RC = WSACleanup()
                Sock = 0
                Process.Caption = "The Server responds with an unexpected Response Code!"
                Exit Sub
            End If
        End If

        ElseIf WSAGetLastError() <> WSAEWOULDBLOCK Then
            closesocket (Sock)
            RC = WSACleanup()
            Sock = 0
        End If
    End If

Refresh


End Sub

'***************************************************************
'Sends the E-Mail
'***************************************************************

Private Sub Transmit(iStage As Integer)
Dim Helo As String
Dim pos As Integer

Select Case m_iStage

Case 1:
Helo = Frombox.Text
pos = Len(Helo) - InStr(Helo, "@")
Helo = Right$(Helo, pos)

ResponseCode = 250
WinsockSendData ("HELO " & Helo & vbCrLf)

Call WaitForResponse

Case 2:
ResponseCode = 250
WinsockSendData ("MAIL FROM: <" & Trim(Frombox.Text) & ">" & vbCrLf)

Call WaitForResponse

Case 3:
ResponseCode = 250
WinsockSendData ("RCPT TO: <" & Trim(Tobox.Text) & ">" & vbCrLf)

Call WaitForResponse

Case 4:
ResponseCode = 354
WinsockSendData ("DATA" & vbCrLf)

Call WaitForResponse

Case 5:
' Calls the routine to send the Header
ResponseCode = 250
Call SendMimetxt(Frombox.Text, Tobox.Text, Subjekt.Text, Mailtxt.Text, Form1.AttachementList.List(0))

Call WaitForResponse



'Finish the E-Mail sending process
Case 6:
ResponseCode = 221
WinsockSendData ("QUIT" & vbCrLf)
Process.Caption = "E-Mail was sended!"

m_iStage = 0
bTrans = False
Call WaitForResponse

End Select
End Sub

'***************************************************************
'Routine for sending a MIME txt
'***************************************************************

Sub SendMimetxt(txtFrom, txtTo, txtSubjekt, txtMail, txtMimePath)
Dim temp As Variant

If txtMimePath <> "" Then
'Prepare the MIME Mail Header

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If you want additional Headers like Date,Message-Id,...etc. !
'simply add them below                                      !
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
temp = temp & "From: " & txtFrom & vbNewLine
temp = temp & "To: " & txtTo & vbNewLine
temp = temp & "Subject: " & txtSubjekt & vbNewLine

'Do not change this Headers
temp = temp & "Mime-Version: 1.0" & vbNewLine
temp = temp & vbCrLf & "Content-Type: multipart/mixed; boundary=" + Chr(34) + "NextMimePart" + Chr(34) + vbNewLine
temp = temp & "This is a multi-part message in MIME format." + vbNewLine
temp = temp & "--NextMimePart" + vbNewLine

'Header plus Message
temp = temp + vbCrLf + Mailtxt.Text

'Send the Mime Header and the Message
WinsockSendData (temp & vbCrLf)

'Call Attachement Routine
SendMimeAttachement (txtMimePath)

Else
'Send the E-Mail without Attachement

temp = temp & "From: " & txtFrom & vbNewLine
temp = temp & "To: " & txtTo & vbNewLine
temp = temp & "Subject: " & txtSubjekt & vbNewLine
temp = temp & vbCrLf & txtMail

'Send Data and finish it!
WinsockSendData (temp)
WinsockSendData (vbCrLf & "." & vbCrLf)
End If

End Sub

'**************************************************************
'NEW! Waits until time out, while waiting for response
'**************************************************************

Private Sub WaitForResponse()
Dim Start As Long
Dim Tmr As Long

'Works with an Api Declaration because it's more precious

Start = timeGetTime
While Bytes > 0
    Tmr = timeGetTime - Start
    DoEvents ' Let System keep checking for incoming response
        
    'Wait 50 (50000 Miliseconds) seconds for response
    If Tmr > 50000 Then
        Process.Caption = "SMTP service error, timed out while waiting for response"
        End
    End If
Wend
End Sub

'***************************************************************
'Routine for sending a MIME Attachement
'***************************************************************

Private Sub SendMimeAttachement(path As Variant)
'Dim Global
Dim l As Long, i As Long, FileIn As Long
Dim temp As Variant
'For Encoding BASE64
Dim b As Integer
Dim Base64Tab As Variant
Dim bin(3) As Byte
Dim s As Variant


'Base64Tab holds the encode tab
Base64Tab = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "/")

'Gets the next free filenumber
FileIn = FreeFile

'Open Base64 Input File
Open path For Binary As FileIn

'Preparing the Mime Header
temp = vbCrLf + "--NextMimePart" + vbNewLine
temp = temp + "Content-Type: application/octet-stream; name=" + Chr(34) + Mimefilename + Chr(34) + vbNewLine
temp = temp + "Content-Transfer-Encoding: base64" + vbNewLine
temp = temp + "Content-Disposition: attachment; filename=" + Chr(34) + Mimefilename + Chr(34) + vbNewLine

WinsockSendData (temp & vbCrLf)


l = LOF(FileIn) - (LOF(FileIn) Mod 3)

For i = 1 To l Step 3

'Read three bytes
Get FileIn, , bin(0)
Get FileIn, , bin(1)
Get FileIn, , bin(2)


'Always wait until there're more then 64 characters
If Len(s) > 64 Then
    
    
    Process.Caption = "Send Attachement..." & i & " Bytes from " & l
    DoEvents
    s = s + vbCrLf
    WinsockSendData (s)
    s = ""
    
End If

'Calc Base64-encoded char

    b = (bin(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
    
    'the character s holds the encoded chars
    s = s + Base64Tab(b)

    b = ((bin(0) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
    s = s + Base64Tab(b)
    b = ((bin(1) And &HF) * 4) Or ((bin(2) \ 64) And &H3)
    s = s + Base64Tab(b)
    b = bin(2) And &H3F
    s = s + Base64Tab(b)


 Next i

'Now, you need to check if there is something left
If Not (LOF(FileIn) Mod 3 = 0) Then

'Reads the number of bytes left
For i = 1 To (LOF(FileIn) Mod 3)
    Get FileIn, , bin(i - 1)
Next i



'If there are only 2 chars left
If (LOF(FileIn) Mod 3) = 2 Then

    b = (bin(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
    s = s + Base64Tab(b)
    b = ((bin(0) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
    s = s + Base64Tab(b)
    b = ((bin(1) And &HF) * 4) Or ((bin(2) \ 64) And &H3)
    s = s + Base64Tab(b)
    s = s + "="

'If there is only one char left
Else
    b = (bin(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
    s = s + Base64Tab(b)
    b = ((bin(1) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
    s = s + Base64Tab(b)
    s = s + "=="
End If
End If

'Send the characters left
If s <> "" Then
    s = s & vbCrLf
    WinsockSendData (s)
End If

'Send the last part of the MIME Body
WinsockSendData (vbCrLf & "--NextMimePart--" & vbCrLf)
WinsockSendData (vbCrLf & "." & vbCrLf)

Close FileIn
End Sub


Private Sub WinsockSendData(DatatoSend As String)
Dim RC As Integer
Dim MsgBuffer As String * 2048

MsgBuffer = DatatoSend

'You can open more than one connection!
RC = send(Sock, ByVal MsgBuffer, Len(DatatoSend), 0)
    
'If an error occurs send an error message and
'reset the winsock
If RC = SOCKET_ERROR Then
    Process.Caption = "Cannot Send Request." + _
                            Chr$(13) + Chr$(10) + _
                            Str$(WSAGetLastError()) + _
                            GetWSAErrorString(WSAGetLastError())
    closesocket Sock
    RC = WSACleanup()
    Exit Sub
End If


End Sub

