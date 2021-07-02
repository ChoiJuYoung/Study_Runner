VERSION 5.00
Begin VB.Form Frm_HowTo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "다시는 이 창을 띄우지 않습니다."
      Height          =   180
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   3015
   End
   Begin Study_Runner.jcbutton jcbutton2 
      Height          =   225
      Left            =   3120
      TabIndex        =   0
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   397
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "End"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '투명
      Caption         =   "Y키를 누르면 지우개가 사용됩니다."
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Caption         =   "도중 종료는 Esc키를 누르시면 됩니다."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Caption         =   "다가오는 LOL로고를 피하셔야 합니다."
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "목숨이 모두 고갈되면 게임 오버가 됩니다."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "3단 점프 : Shift + Enter"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "2단 점프 : Ctrl + Space Bar"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "1단 점프 : Space Bar"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Frm_HowTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Checking As String '레지스트리 변수값 확인 변수

Private Sub Form_Load()
'다시 보지 않음 확인부
Checking = GetSetting("STRE", "STRE", "STRE")
DoEvents
If Checking = "Yes" Then
    jcbutton2_Click
End If
End Sub
Private Sub jcbutton2_Click()
'게임 실행부
If Check1.Value = 1 Then
    SaveSetting "STRE", "STRE", "STRE", "Yes"
End If
Frm_Play.Show
Unload Me
End Sub
