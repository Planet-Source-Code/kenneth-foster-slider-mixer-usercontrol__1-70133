VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "New Fader Control"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin Project1.Fader Fader4 
      Height          =   330
      Left            =   180
      TabIndex        =   11
      Top             =   3945
      Width           =   3135
      _ExtentX        =   582
      _ExtentY        =   582
      Style           =   1
      HalfMark        =   -1  'True
      TickMarkCnt     =   10
   End
   Begin Project1.Fader Fader3 
      Height          =   330
      Left            =   195
      TabIndex        =   9
      Top             =   3555
      Width           =   3120
      _ExtentX        =   582
      _ExtentY        =   582
      Style           =   1
      HalfMark        =   -1  'True
      BackColor       =   14737632
      ButSz           =   1
   End
   Begin Project1.Fader Fader2 
      Height          =   2940
      Left            =   2730
      TabIndex        =   7
      Top             =   270
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   5186
      Max             =   255
      HalfMarkColor   =   16777215
      BackColor       =   12648447
      TickMarkCnt     =   20
   End
   Begin Project1.Fader Fader1 
      Height          =   3045
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   5371
      HalfMark        =   -1  'True
      TickMarkCnt     =   10
   End
   Begin Project1.Fader Fader1 
      Height          =   1725
      Index           =   1
      Left            =   2115
      TabIndex        =   1
      Top             =   240
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   3043
      Max             =   50
      HalfMark        =   -1  'True
      HalfMarkColor   =   16711680
      BackColor       =   255
      TickMarks       =   0   'False
   End
   Begin Project1.Fader Fader1 
      Height          =   3045
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   5371
      HalfMark        =   -1  'True
      ButSz           =   1
   End
   Begin Project1.Fader Fader1 
      Height          =   1725
      Index           =   3
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   3043
      Max             =   25
      HalfMark        =   -1  'True
      HalfMarkColor   =   16777215
      BackColor       =   16711680
   End
   Begin Project1.Fader Fader1 
      Height          =   1260
      Index           =   4
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   2223
   End
   Begin Project1.Fader Fader1 
      Height          =   3045
      Index           =   5
      Left            =   1185
      TabIndex        =   5
      Top             =   240
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   5371
      TickMarks       =   0   'False
   End
   Begin Project1.Fader Fader1 
      Height          =   1275
      Index           =   6
      Left            =   2130
      TabIndex        =   6
      Top             =   2025
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   2249
      Max             =   75
      BackColor       =   49152
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   705
      TabIndex        =   12
      Top             =   3315
      Width           =   345
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2730
      TabIndex        =   10
      Top             =   3255
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1545
      TabIndex        =   8
      Top             =   4320
      Width           =   435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Fader1_Scrolling(Index As Integer)
   Label1.Caption = Fader1(Index).Value
End Sub

Private Sub Fader2_Scrolling()
Label3.Caption = Fader2.Value
End Sub

Private Sub Fader3_Scrolling()
Label2.Caption = Fader3.Value
End Sub

Private Sub Fader4_Scrolling()
   Label2.Caption = Fader4.Value
End Sub

