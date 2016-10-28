VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "hx19access"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13320
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "hx19wifiAccess.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   11760
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "10.10.100.254"
      RemotePort      =   8899
   End
   Begin VB.CheckBox Check3 
      Caption         =   "X"
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   6012
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Broadcast hx19setup.txt"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Log"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      Text            =   "Send String>"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   390
      Left            =   7440
      TabIndex        =   5
      Text            =   "String Size>"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sync Mode"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8640
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TX"
      Height          =   375
      Left            =   10560
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8292
      Left            =   8520
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   3132
   End
   Begin VB.TextBox Text6 
      Height          =   600
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim linebuffer(100) As String            'only used for the scrolling routine, not linked to hx19 operations
Dim nn%                                  'used only for scrolling
Dim haltDisplay As Boolean

Private Sub Command2_Click()
Dim conf$
    Open "hx19setup.txt" For Input As 1
    Do
        Input #1, conf
        If Len(conf) < 4 Then GoTo sdone
        Text7 = Text7 + conf + vbCrLf
        checkOut conf
        Do: DoEvents: cc = Form1.MSComm1.Input: Loop Until cc = "k" Or cc = "#"
        Do: DoEvents: cc = Form1.MSComm1.Input: Loop Until cc = chr13
    Loop Until EOF(1)
sdone:
    Text7 = Text7 + "DONE" + vbCrLf
    Close
End Sub
 
Private Sub Check1_Click()
    'Any hx19ms receiving the command $ will initiate syncronized strobing
    'the command % stops the sync sequence
    If Check1.Value = 1 Then checkOut "M&$" Else checkOut "M&%"
End Sub

Private Sub Command3_Click()
    checkOut Text6                          'here the content of Text6 is transmitted via usb com port to the hx19ms
End Sub

Private Sub checkOut(temst As String)
    Dim xsum As Integer, xx As String
    'this routine sums up all Ascii characters entered, and creates an hx19 accepted checksum.
    xsum = 0
    For i = 1 To Len(temst)                'compute the checksum of the string
       xx = Mid(temst, i, 1)
       xsum = xsum + Asc(xx)               'accumulate ASCII codes
    Next
    temst = temst + "/" + Hex(xsum)        'append the checksum in hexadecimal format
    'Print temst
    Winsock1.SendData (temst + Chr(13))
 End Sub
Private Sub Form_Load()
    Winsock1.RemoteHost = "10.10.100.254"
    Winsock1.RemotePort = 8899
    Winsock1.Protocol = sckTCPProtocol
    Winsock1.Connect
End Sub

Private Sub Text6_Change()
    'This routine keeps track of the total characters entered, should not exeed 116 characters.
    Text5 = Format(Len(Text6), "#")
End Sub

'Remaining routines are unimportant and secondary to the understanding of the hx19

Private Sub Check2_Click()
    'data coming from the hx19ms is text format and can be viewed using windows notepad
     If Check2.Value = 1 Then
        Open "hx19access.txt" For Output As 1             'save incoming hx19ms text data on file called hx19.log
     Else
       Close 1
     End If
End Sub

Private Sub tScroll(nline)                      'scrolls 16 lines of text through text window
Dim jj%
'routine scrolls down one line for display purposes only, it has otherwise nothing to do with hx19 system
    linebuffer(nn) = nline
    nn = (nn + 1) And 7
    jj = nn + 1
    Text3 = ""
    Do
     Text3 = Text3 + linebuffer(jj)
     jj = (jj + 1) And 7
    Loop Until jj = nn
End Sub

Private Sub Form_Terminate()
    Winsock1.Close
    Close                                           'make sure no files are left open when the program ends
    End                                             'if end isn't executed before termination, com port may be left open
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Winsock1.GetData strData, vbString, bytesTotal 'note that winsock updates 10 times per sec but hx19 16 times
    'winsock seems to suppress carriage return (delimiter) occasionally so we must use X as reference, or find another programming language
    Print strData;
    tScroll strData + Chr(10)
    strData = ""
    
    'Print "end"
    ii = 1
'    For i = 1 To Len(strData) + 1
 '    If "X" = Mid(strData, i, i + 1) Then Print "x";: ii = i + 1
  '   If "X" = Mid(strData, i, i + 1) Then tScroll Mid(strData, ii, i): ii = i + 1
  '  Next
   
    
    
    'Print strData
End Sub

