VERSION 5.00
Begin VB.Form Frm_Main 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Main"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4650
   StartUpPosition =   2  '화면 가운데
   Begin Study_Runner.jcbutton CmdCon 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "공적 보기"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Study_Runner.jcbutton CmdRank 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "랭킹 보기"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin Study_Runner.jcbutton Cmd_Help 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "입시 News"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin Study_Runner.jcbutton Cmd_CR 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "CopyRight"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin Study_Runner.jcbutton Cmd_Load 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "Load Runner"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin Study_Runner.jcbutton Cmd_New 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "New Runner"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   120
      Picture         =   "Frm_Main.frx":08CA
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_CR_Click()
'CopyRight 폼 오픈
FrmCopy.Show
End Sub

Private Sub Cmd_Help_Click()
'EBSi 입시정보 폼 오픈
Frm_Help.Show
End Sub

Private Sub Cmd_Load_Click()
'불러오기 폼 오픈
Call load(Frm_Load, Frm_Main)
DoEvents
End Sub

Private Sub Cmd_New_Click()
'새 게임 시작 ^_^ ;
On Error GoTo howto:
Frm_HowTo.Show
Unload Me

Exit Sub

howto:
Unload Me
End Sub

Private Sub CmdCon_Click()
FrmContri.Show
End Sub

Private Sub CmdRank_Click()
'랭킹보기 폼 오픈
FrmShowRank.Show
End Sub

Private Sub Form_Load()
'공적 가져오기용
Call load(Frm_Loading, Frm_Loading)
DoEvents
'변수값 초기화
Dim j As Integer
SPoint = 0
LifPoint = 3
GPoint = 0
Randomize Q
Stage = 1
For i = 1 To 12
    Stg_Clear(i) = False
    For j = 1 To 3
        QuestionB(j, i) = 0
        Randomize Qq(j)
    Next
Next

'정답 입력
''넌센스
Question(1, 2) = 30
Question(1, 3) = 0
Question(1, 4) = "참새"
Question(1, 5) = "야금야금"
Question(1, 6) = "이쑤시개"
Question(1, 7) = "홍길동전"
Question(1, 8) = "배우"
Question(1, 9) = "끼리끼리"
Question(1, 10) = "허수"
Question(1, 11) = "경마장"
Question(1, 12) = "슈퍼마리오"

''과학
Question(2, 1) = "상대성"
Question(2, 2) = "중화반응"
Question(2, 3) = "리보솜"
Question(2, 4) = "HR도"
Question(2, 5) = "렌츠"
Question(2, 6) = "산화수"
Question(2, 7) = "용불용설"
Question(2, 8) = "모스굳기계"

''사회
Question(3, 1) = "자연현상"
Question(3, 2) = "당위법칙"
Question(3, 3) = "존재법칙"
Question(3, 4) = "객관적태도"
Question(3, 5) = "도피형"
Question(3, 6) = "사회조직"
Question(3, 7) = "연공서열"
Question(3, 8) = "공유성"
Question(3, 9) = "사회문화현상"

End Sub
