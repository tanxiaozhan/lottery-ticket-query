VERSION 5.00
Begin VB.Form Main 
   Caption         =   "双色球随机号生成程序"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11415
   StartUpPosition =   2  '屏幕中心
   Begin 工程1.XPButton XPButton1 
      Height          =   855
      Left            =   4800
      TabIndex        =   23
      Top             =   6840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
      Caption         =   "退  出"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin 工程1.FTextBox txtBox 
      Height          =   780
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      MaxLength       =   3
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox 
      Height          =   780
      Index           =   1
      Left            =   3312
      TabIndex        =   3
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      MaxLength       =   3
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox 
      Height          =   780
      Index           =   2
      Left            =   4584
      TabIndex        =   4
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      MaxLength       =   3
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox 
      Height          =   780
      Index           =   3
      Left            =   5856
      TabIndex        =   5
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      MaxLength       =   3
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox 
      Height          =   780
      Index           =   4
      Left            =   7128
      TabIndex        =   6
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      MaxLength       =   3
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox 
      Height          =   780
      Index           =   5
      Left            =   8400
      TabIndex        =   7
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      MaxLength       =   3
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox 
      Height          =   780
      Index           =   6
      Left            =   10080
      TabIndex        =   8
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      MaxLength       =   3
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox1 
      Height          =   780
      Index           =   0
      Left            =   2040
      TabIndex        =   9
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox2 
      Height          =   780
      Index           =   0
      Left            =   2040
      TabIndex        =   10
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox1 
      Height          =   780
      Index           =   1
      Left            =   3312
      TabIndex        =   11
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox1 
      Height          =   780
      Index           =   2
      Left            =   4584
      TabIndex        =   12
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox1 
      Height          =   780
      Index           =   3
      Left            =   5856
      TabIndex        =   13
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox1 
      Height          =   780
      Index           =   4
      Left            =   7128
      TabIndex        =   14
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox1 
      Height          =   780
      Index           =   5
      Left            =   8400
      TabIndex        =   15
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox1 
      Height          =   780
      Index           =   6
      Left            =   10080
      TabIndex        =   16
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox2 
      Height          =   780
      Index           =   1
      Left            =   3312
      TabIndex        =   17
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox2 
      Height          =   780
      Index           =   2
      Left            =   4584
      TabIndex        =   18
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox2 
      Height          =   780
      Index           =   3
      Left            =   5856
      TabIndex        =   19
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox2 
      Height          =   780
      Index           =   4
      Left            =   7128
      TabIndex        =   20
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox2 
      Height          =   780
      Index           =   5
      Left            =   8400
      TabIndex        =   21
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin 工程1.FTextBox txtBox2 
      Height          =   780
      Index           =   6
      Left            =   10080
      TabIndex        =   22
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1376
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Times New Roman"
      FontSize        =   30
      ForeColor       =   16777215
      Locked          =   -1  'True
      Enabled         =   0   'False
      AutoSelAll      =   -1  'True
      Alignment       =   2
      isNumber        =   -1  'True
      afterdecimal    =   0
   End
   Begin VB.Label Label5 
      Caption         =   "-3 — 3"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "输入号码"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   25
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "-7 — 8"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   24
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "双色球随机选号程序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "双色球随机选号程序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   7455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.Left = Label2.Left + 50
    Label1.Top = Label2.Top + 50
End Sub

Private Sub Form_Resize()
    Shape1.Left = 0
    Shape1.Top = 0
    Shape1.Width = Me.Width
End Sub

Private Sub txtBox_Change(Index As Integer)
    Randomize
    txtBox1(Index).Text = Int(Rnd(1) * 16) - 7
    txtBox2(Index).Text = Int(Rnd(1) * 7) - 3
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Index < 6 Then txtBox(Index + 1).SetFocus
    
End Sub

Private Sub XPButton1_Click()
    Unload Me
End Sub
