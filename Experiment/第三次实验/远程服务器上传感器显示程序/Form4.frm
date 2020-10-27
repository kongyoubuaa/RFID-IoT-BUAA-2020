VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{B6C10482-FB89-11D4-93C9-006008A7EED4}#1.0#0"; "TeeChart5.ocx"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFC0&
   Caption         =   "BC28_UDP服务器显示"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13185
   BeginProperty Font 
      Name            =   "黑体"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":08CA
   ScaleHeight     =   8205
   ScaleWidth      =   13185
   StartUpPosition =   3  '窗口缺省
   Begin TeeChart.TChart TChart1 
      Height          =   3855
      Left            =   2400
      TabIndex        =   12
      Top             =   3960
      Width           =   7815
      Base64          =   $"Form4.frx":1194
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   9600
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   960
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton close 
      BackColor       =   &H00E0E0E0&
      Caption         =   "关闭"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton TCP_start 
      BackColor       =   &H00E0E0E0&
      Caption         =   "启动服务器"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   240
      TabIndex        =   5
      Text            =   "8888"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   11040
      Top             =   4440
   End
   Begin Project1.MorphDisplay temp 
      Height          =   1155
      Left            =   3360
      TabIndex        =   9
      Top             =   2040
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   2037
      BurnInColor     =   4210688
      InterDigitGap   =   20
      InterSegmentGap =   2
      InterSegmentGapExp=   1
      NumDigits       =   2
      NumDigitsExp    =   2
      SegmentHeight   =   20
      SegmentHeightExp=   20
      SegmentStyle    =   0
      SegmentStyleExp =   0
      SegmentWidth    =   8
      SegmentWidthExp =   6
      Value           =   "0"
      XOffset         =   10
      XOffsetExp      =   305
      YOffset         =   8
      YOffsetExp      =   58
   End
   Begin Project1.MorphDisplay humi 
      Height          =   1155
      Left            =   7200
      TabIndex        =   10
      Top             =   2040
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   2037
      BurnInColor     =   4210688
      InterDigitGap   =   20
      InterSegmentGap =   2
      InterSegmentGapExp=   1
      NumDigits       =   2
      NumDigitsExp    =   2
      SegmentHeight   =   20
      SegmentHeightExp=   20
      SegmentStyle    =   0
      SegmentStyleExp =   0
      SegmentWidth    =   8
      SegmentWidthExp =   6
      Value           =   "0"
      XOffset         =   10
      XOffsetExp      =   305
      YOffset         =   8
      YOffsetExp      =   58
   End
   Begin VB.Label nowtime 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   9600
      TabIndex        =   13
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "本地端口"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      TabIndex        =   4
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   3
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "湿度："
      Height          =   735
      Left            =   6240
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "温度："
      Height          =   735
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "BC28温湿度采集显示--墨子号科技"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3720
      TabIndex        =   0
      Top             =   600
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   8415
      Left            =   -240
      Picture         =   "Form4.frx":15A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long


Private Sub close_Click()
End
End Sub

Private Sub Form_Load() '启动默认加载
On Error Resume Next '得到存在串口的函数
Winsock1.close
temp.Value = 0
humi.Value = 0

End Sub





Private Sub TCP_start_Click()
If TCP_start.Caption = "启动服务器" Then
Winsock1.LocalPort = Val(Text1.Text) '端口号
Winsock1.Bind
TCP_start.Caption = "关闭监听"
Else
Winsock1.close
End If
End Sub

Private Sub Timer1_Timer()
nowtime = Now
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    'If Winsock1.State <> sckClosed Then Winsock1.Close
    Do While (Winsock1.State <> sckClosed)
        Winsock1.close
    Loop
    Winsock1.Accept requestID '注释：请求到达时，接受连接
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim buffer() As Byte, strdata As String
Text2.Text = ""
Winsock1.GetData buffer, vbString '从网络中获取数据
For i = 0 To UBound(buffer)
strdata = strdata & Hex(buffer(i))
If buffer(i) = &HFF Then '表明读取到了数据起始位
temp.Value = Hex(buffer(i + 1))
humi.Value = Hex(buffer(i + 2))
TChart1.Series(0).Add Int(temp.Value), a, 255 '图表添加
TChart1.Series(1).Add Int(humi.Value), b, &HFF00&
End If
Next

Text2.Text = strdata

End Sub

