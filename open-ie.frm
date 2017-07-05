VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "IE调用中间件"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   5265
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtMain 
      Height          =   630
      Left            =   240
      TabIndex        =   0
      Top             =   1185
      Width           =   4575
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   840
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private clientCount As Integer
Private Sub Form_Load()
    wskServer(0).LocalPort = 8989
    wskServer(0).Listen
End Sub

Private Sub wskServer_Close(Index As Integer)
    'MsgBox (1)
    'wskServer.Listen
    'wskServer.LocalPort = 8989
    'Unload Me
    'Load Me
End Sub

Private Sub wskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    If Index = 0 Then
          clientCount = clientCount + 1       '客户请求多一个
          Load wskServer(clientCount)         '载入一个服务端为新增的客户服务
          wskServer(clientCount).LocalPort = 0    '侦听端口为随机，不能设为2000，因为有sckServer(0)在使用了。
          wskServer(clientCount).Accept requestID        '接受请求
     End If

    'If wskServer.State <> sckClosed Then
        'wskServer.Close
    'End If
    'wskServer.Accept requestID
    
End Sub

Private Sub wskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
     Dim tempS As String
     wskServer(Index).GetData tempS
     txtMain.Text = tempS
     Dim a
     a = Split(tempS, " ")
    
     If a(1) <> "/favicon.ico" Then
         Dim url
     url = Split(a(1), "url=")(1)
        txtMain.Text = url
     Dim RetVal
        RetVal = Shell("C:/Program Files (x86)/Internet Explorer/iexplore.exe " & url, 1)
        'Shell "C:/Program Files (x86)/Internet Explorer/iexplore.exe " & "http://www.baidu.com", 1
        txtMain.Text = txtMain.Text & "------RetVal=" & RetVal
        If RetVal <> 0 Then
            wskServer(Index).SendData "ok"
        Else
            wskServer(Index).SendData "fail"
        End If
     End If
     
     
End Sub


Private Sub wskServer_SendComplete(Index As Integer)
    wskServer(Index).Close
    'Unload Me
    'Load Me
    'Dim v
    'v = Shell("C:/Program Files (x86)/Internet Explorer/iexplore.exe ", 1)
End Sub

