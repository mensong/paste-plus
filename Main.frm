VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Paste ++"
   ClientHeight    =   1275
   ClientLeft      =   8355
   ClientTop       =   4155
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   6.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
SetHotKey
Me.Move 0, 0
Me.Height = 1050
Me.Width = 5520

SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3 '窗口置顶

Dim i As Integer
For i = 0 To 4
    Text(i).MousePointer = 1
    Text(i).ToolTipText = "CTRL +" & Str(i + 1)
Next i

'Me.Width = Screen.Width
'Frame1.Move Me.Width - Frame1.Width
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Height = 1050
Me.Width = 5520
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim ret As Long
'取消Message的截取，而使之又只送往原来的Window Procedure
ret = SetWindowLong(Me.hwnd, GWL_WNDPROC, preWinProc)
Call UnregisterHotKey(Me.hwnd, uVirtKey)
End Sub


Function Paste(ByVal Index As Integer)
If Text(Index).Text <> "" Then
Clipboard.Clear
Clipboard.SetText Me.Text(Index).Text
End If
End Function

Private Sub Text_DblClick(Index As Integer)
Paste (Index)
End Sub
