VERSION 5.00
Begin VB.Form frmSendMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMTP w\ Attachments - no WInsock Control"
   ClientHeight    =   7080
   ClientLeft      =   7335
   ClientTop       =   2010
   ClientWidth     =   6615
   Icon            =   "frmSendMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   6615
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   3960
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Text            =   "frmSendMail.frx":000C
      Top             =   3030
      Width           =   6375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send message"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Text            =   "frmSendMail.frx":0016
      Top             =   1560
      Width           =   6375
   End
   Begin VB.TextBox txtSubject 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Text            =   "SMTP with attachment TEST"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtRecipient 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Text            =   "foo@cox.net"
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtSender 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "foo@cox.net"
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtHost 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "smtp.east.cox.net"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1245
      TabIndex        =   3
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Recipient e-mail address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Your e-mail address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   405
      TabIndex        =   1
      Top             =   480
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SMTP Host:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private m_State As SMTP_State
Private AttachmentFiles As String
Dim WithEvents Winsock1 As CSocketMaster
Attribute Winsock1.VB_VarHelpID = -1

Private Enum SMTP_State
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum


Private Function MXQuery()


    On Error GoTo Err_MXQuery

    MX_Query ("")

    If MX.Count Then
        MXQuery = MX.Best

    End If

Exit Function

Err_MXQuery:

End Function

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdSend_Click()

    Winsock1.Connect Trim$(txtHost), 25
        
    'see in DataArrival
    m_State = MAIL_CONNECT
    
End Sub

Private Sub Form_Load()
    'geez is this a beautiful .cls!
    Set Winsock1 = New CSocketMaster
    
    'try to get a SMTP server on this machine
    txtHost = MXQuery
    
    'cox cable likes smtp rather than the returned mx for my accnt, may not need this for your server
    txtHost = Replace(txtHost, "mx", "smtp")
    
    'This is 2 attachments located in app.path - a .txt file and a .gif
    'they must be encoded first
    AttachmentFiles = UUEncodeFile(App.Path & "\2.gif") & vbCrLf
    AttachmentFiles = AttachmentFiles & UUEncodeFile(App.Path & "\Attach.txt") & vbCrLf
   
    'To attach just 1 file:
    'AttachmentFiles = UUEncodeFile(App.Path & "\5.jpg") & vbCrLf
    
    
     txtMessage = txtMessage & vbNewLine & vbnwline & "Attachments are encoded and ready to go..." & _
     vbNewLine & "Add any text you would like to send w\ the attachments in this textbox.(txtMessage)"
                        
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Winsock1 = Nothing
End Sub


Private Sub txtMessage_Click()
txtMessage = vbNullString
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    '
    'Retrive data from winsock buffer
    '
    Winsock1.GetData strServerResponse
    '
    frmSendMail.txtStatus = frmSendMail.txtStatus & vbNewLine & strServerResponse
    frmSendMail.txtStatus.SelStart = Len(frmSendMail.txtStatus.Text)
    '
    'Get server response code (first three symbols)
    '
    strResponseCode = Left(strServerResponse, 3)
    '
    'Only these three codes tell us that previous
    'command accepted successfully and we can go on
    '
    If strResponseCode = "250" Or _
       strResponseCode = "220" Or _
       strResponseCode = "354" Then
       
        Select Case m_State
            Case MAIL_CONNECT
                'Change current state of the session
                m_State = MAIL_HELO
                '
                'Remove blank spaces
                strDataToSend = Trim$(txtSender)
                '
                'Retrieve mailbox name from e-mail address
                strDataToSend = Left$(strDataToSend, _
                                InStr(1, strDataToSend, "@") - 1)
                'Send HELO command to the server
                Winsock1.SendData "HELO " & strDataToSend & vbCrLf
                '
                frmSendMail.txtStatus = frmSendMail.txtStatus & vbNewLine & "HELO " & strDataToSend
               frmSendMail.txtStatus.SelStart = Len(frmSendMail.txtStatus.Text)
                '
            Case MAIL_HELO
                '
                'Change current state of the session
                m_State = MAIL_FROM
                '
                'Send MAIL FROM command to the server
                Winsock1.SendData "MAIL FROM:" & Trim$(txtSender) & vbCrLf
                '
                frmSendMail.txtStatus = frmSendMail.txtStatus & vbNewLine & "MAIL FROM:" & Trim$(txtSender)
               frmSendMail.txtStatus.SelStart = Len(frmSendMail.txtStatus.Text)
                '
            Case MAIL_FROM
                '
                'Change current state of the session
                m_State = MAIL_RCPTTO
                '
                'Send RCPT TO command to the server
                Winsock1.SendData "RCPT TO:" & Trim$(txtRecipient) & vbCrLf
                '
                frmSendMail.txtStatus = frmSendMail.txtStatus & vbNewLine & "RCPT TO:" & Trim$(txtRecipient)
               frmSendMail.txtStatus.SelStart = Len(frmSendMail.txtStatus.Text)
            Case MAIL_RCPTTO
                '
                'Change current state of the session
                m_State = MAIL_DATA
                '
                'Send DATA command to the server
                Winsock1.SendData "DATA" & vbCrLf
                '
                frmSendMail.txtStatus = frmSendMail.txtStatus & vbNewLine & "DATA"
               frmSendMail.txtStatus.SelStart = Len(frmSendMail.txtStatus.Text)
            Case MAIL_DATA
                '
                'Change current state of the session
                m_State = MAIL_DOT
                '
                'So now we are sending a message body
                'Each line of text must be completed with
                'linefeed symbol (Chr$(10) or vbLf) not with vbCrLf - This is wrong, it should be vbCrLf
                'see   http://cr.yp.to/docs/smtplf.html       for details
                '
                'Send Subject line
                Winsock1.SendData "From:" & txtSenderName & " <" & txtSender & ">" & vbCrLf
                Winsock1.SendData "To:" & txtRecipientName & " <" & txtRecipient & ">" & vbCrLf
                
                '
                
                '
                If Len(txtReplyTo) > 0 Then
                    Winsock1.SendData "Subject:" & txtSubject & vbCrLf
                    Winsock1.SendData "Reply-To:" & txtReplyToName & " <" & txtReplyTo & ">" & vbCrLf & vbCrLf
                Else
                    Winsock1.SendData "Subject:" & txtSubject & vbCrLf & vbCrLf
                End If
                'Dim varLines() As String
                'Dim varLine As String
                Dim strMessage As String
                'Dim i
                '
                'Add atacchments
                strMessage = txtMessage & vbCrLf & vbCrLf & AttachmentFiles
                'clear memory
               AttachmentFiles = ""
                'label1.caption = Len(strMessage)
                'These lines aren't needed, see
                '
                'http://cr.yp.to/docs/smtplf.html for details
                '
                '*****************************************
                'Parse message to get lines (for VB6 only)
                'varLines() = Split(strMessage, vbNewLine)
                'Parse message to get lines (for VB5 and lower)
                'SplitMessage strMessage, varLines()
                'clear memory
                'strMessage = ""
                '
                'Send each line of the message
                'For i = LBound(varLines()) To UBound(varLines())
                '    Winsock1.SendData varLines(i) & vbCrLf
                '    '
                '    label1.caption = varLines(i)
                'Next
                '
                '******************************************
                Winsock1.SendData strMessage & vbCrLf
                strMessage = ""
                '
                'Send a dot symbol to inform server
                'that sending of message comleted
                Winsock1.SendData "." & vbCrLf
                '
               ' Label1.Caption = "."
                '
            Case MAIL_DOT
                'Change current state of the session
                m_State = MAIL_QUIT
                '
                'Send QUIT command to the server
                Winsock1.SendData "QUIT" & vbCrLf
                '
                frmSendMail.txtStatus = frmSendMail.txtStatus & vbNewLine & "QUIT"
               frmSendMail.txtStatus.SelStart = Len(frmSendMail.txtStatus.Text)
            Case MAIL_QUIT
                '
                'Close connection
                Winsock1.CloseSck
                '
        End Select
       
    Else
        '
        'If we are here server replied with
        'unacceptable respose code therefore we need
        'close connection and inform user about problem
        '
        Winsock1.CloseSck
        '
        If Not m_State = MAIL_QUIT Then
            frmSendMail.txtStatus = frmSendMail.txtStatus & vbNewLine & vbNewLine & "SMTP Error: " & strServerResponse
            frmSendMail.txtStatus.SelStart = Len(frmSendMail.txtStatus.Text)
        Else
           frmSendMail.txtStatus = frmSendMail.txtStatus & vbNewLine & vbNewLine & "Message sent successfuly." & _
           vbNewLine & "STATUS:" & vbNewLine & "Ready..."
           frmSendMail.txtStatus.SelStart = Len(frmSendMail.txtStatus.Text)
        End If
        '
    End If
  
End Sub


Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 txtStatus = txtStatus & vbNewLine & "Winsock Error number " & Number & vbCrLf & _
            Description
            Winsock1.CloseSck
            frmSendMail.txtStatus.SelStart = Len(frmSendMail.txtStatus.Text)

End Sub


