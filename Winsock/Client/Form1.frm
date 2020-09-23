VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Echo Client"
   ClientHeight    =   795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   795
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   3960
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Coded by: Jordan Weber"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Winsock.SendData Text1.Text 'send the text to the server

End Sub

Private Sub Form_Load()

  With Winsock
    .RemoteHost = "127.0.0.1" 'set winsock up to yourself
    .RemotePort = "12345" 'default server port
    .Connect 'connect with server(using specs above)
  End With
  
End Sub

Private Sub Winsock_Close()

End 'close program if server is closed

End Sub

Private Sub Winsock_Connect()

  Me.Caption = "Connected to: " & Winsock.RemoteHostIP 'show user were connected

End Sub


Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Dim rdata As String

Winsock.GetData rdata 'get data

MsgBox rdata 'show user that data has arrived
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

MsgBox "Error: " & Description, , "Echo" 'inform user

End 'close
End Sub
