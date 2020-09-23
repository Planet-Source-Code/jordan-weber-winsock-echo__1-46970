Attribute VB_Name = "Module1"
Public Const MaxUsers = 999 'used for loop

Type Sockets
    Used As Byte 'this will tell us if that index is in use
End Type

Public Socket(1 To MaxUsers) As Sockets

