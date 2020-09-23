VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Echo Server"
   ClientHeight    =   465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   3240
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Coded by: Jordan Weber"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'set up winsocks port than have it wait for connections
With Winsock(0)
    .LocalPort = "12345"
    .Listen
End With

End Sub

Private Sub Winsock_Close(Index As Integer)
'close it up
Winsock(Index).Close
'unload the control
Unload Winsock(Index)
'clean up our variable
Socket(Index).Used = 0

End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

Dim loopc As Integer

For loopc = 1 To MaxUsers 'check for a free socket
    If Socket(loopc).Used = 0 Then 'if socket is free
        Load Winsock(loopc) 'load winsock
        Winsock(loopc).Accept requestID 'accept user
        Socket(loopc).Used = 1 'make socket used
        
        Exit Sub
    End If
Next

End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)

Dim rdata As String

Winsock(Index).GetData rdata 'recieve the data

Winsock(Index).SendData rdata 'send the data back to client
End Sub

Private Sub Winsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

MsgBox "Error: " & Description 'inform the user

End 'close program(if error keeps occuring its annoying)

End Sub
