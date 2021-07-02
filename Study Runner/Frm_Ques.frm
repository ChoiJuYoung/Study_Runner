VERSION 5.00
Begin VB.Form Frm_Ques 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Question"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   Icon            =   "Frm_Ques.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   8145
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame FmeSel 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   3255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8055
      Begin Study_Runner.jcbutton CmdNon 
         Height          =   615
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1085
         ButtonStyle     =   9
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14800597
         Caption         =   "기타 문제"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Study_Runner.jcbutton CmdSoc 
         Height          =   615
         Left            =   0
         TabIndex        =   7
         Top             =   2520
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1085
         ButtonStyle     =   9
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14800597
         Caption         =   "사회 문제"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Study_Runner.jcbutton CmdSci 
         Height          =   615
         Left            =   0
         TabIndex        =   6
         Top             =   1320
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1085
         ButtonStyle     =   9
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14800597
         Caption         =   "과학 문제"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   2
      End
   End
   Begin Study_Runner.jcbutton Cmd_Ans 
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "정답"
      CaptionEffects  =   0
   End
   Begin VB.TextBox txtAns 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   270
      Left            =   1440
      TabIndex        =   2
      Text            =   "Answer"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label lblItem 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "보유 개수 : 1111"
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblItem 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "보유개수 : 1111"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Image ImgItem 
      Height          =   480
      Index           =   1
      Left            =   4800
      Top             =   2520
      Width           =   900
   End
   Begin VB.Image ImgItem 
      Height          =   465
      Index           =   0
      Left            =   1440
      Top             =   2520
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "             수치 문제의 정답은 단위를 빼고 숫자만 적어주세요.           모든 답안은 띄어쓰기 없이 해주세요."
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Label lblQus 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label1"
      Height          =   975
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label lblLev 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Frm_Ques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m As Integer, L As Integer, v As Integer '문제에 사용될 변수들
Private k As Integer, non As Integer, 정답 As Integer

Private Sub Cmd_Ans_Click()
'정답 판별
If k <= 3 Then
    If txtAns = Question(k, Q) Then
        MsgBox "정답!"
        QuestionB(k, Q) = 1
        GoTo ABC:
    Else
        non = Wrong
        If non = 1 Then
            lblQus_Click
            Exit Sub
        ElseIf non = 2 Then
            GoTo ABC:
        ElseIf non = 3 Then
            Exit Sub
        End If
    End If
Else
    If txtAns = 정답 Then
        MsgBox "정답!"
        GoTo ABC:
    Else
        non = Wrong
        If non = 1 Then
            lblQus_Click
            Exit Sub
        ElseIf non = 2 Then
            GoTo ABC:
        ElseIf non = 3 Then
            Exit Sub
        End If
    End If
End If
Exit Sub

ABC:
Qq(1) = 0
Qq(2) = 0
Qq(3) = 0
Unload Me
Frm_Play.Timer1.Enabled = True
Frm_Play.Timer2.Enabled = True
Frm_Play.Timer3.Enabled = True
End Sub

Private Sub CmdNon_Click()
k = 1
Call MakeQuestion(k, lblQus)
FmeSel.Visible = False
End Sub

Private Sub CmdSci_Click()
k = 2
Call MakeQuestion(k, lblQus)
FmeSel.Visible = False
End Sub

Private Sub CmdSoc_Click()
k = 3
Call MakeQuestion(k, lblQus)
FmeSel.Visible = False
End Sub

Private Sub Form_Load()
'초기 세팅
lblQus_Click
For i = 0 To 1
    ImgItem(i) = LoadResPicture(i + 108, vbResBitmap)
Next

If LifPoint = 0 Then
    End
End If

'커맨드버튼 관리 부분
For p = 2 To 12
    If QuestionB(1, p) = 1 Then
        Qq(1) = Qq(1) + 1
    End If
Next

For p = 1 To 8
    If QuestionB(2, p) = 1 Then
        Qq(2) = Qq(2) + 1
    End If
Next

For p = 1 To 9
    If QuestionB(3, p) = 1 Then
        Qq(3) = Qq(3) + 1
    End If
Next




If Qq(1) = 11 Then
    CmdNon.Enabled = False
    Contri(7) = True
Else
    CmdNon.Caption = "기타 문제 (남은수 : " & (11 - Qq(1)) & ")"
End If

If Qq(2) = 8 Then
    CmdSci.Enabled = False
    Contri(6) = True
Else
    CmdSci.Caption = "과학 문제 (남은수 : " & (8 - Qq(2)) & ")"
End If

If Qq(3) = 9 Then
    CmdSoc.Enabled = False
    Contri(8) = True
Else
    CmdSoc.Caption = "사회 문제 (남은수 : " & (9 - Qq(3)) & ")"
End If



If CmdNon.Enabled = False And CmdSci.Enabled = False And CmdSoc.Enabled = False Then
    k = 4
    Call MakeQuestion(k, lblQus)
    FmeSel.Visible = False
End If




DoEvents
'스테이지 표시 부분
If Stage = 1 Then
    lblLev = "Level : 고등학교 1학년 봄"
ElseIf Stage = 2 Then
    lblLev = "Level : 고등학교 1학년 여름"
ElseIf Stage = 3 Then
    lblLev = "Level : 고등학교 1학년 가을"
ElseIf Stage = 4 Then
    lblLev = "Level : 고등학교 1학년 겨울"
ElseIf Stage = 5 Then
    lblLev = "Level : 고등학교 2학년 봄"
ElseIf Stage = 6 Then
    lblLev = "Level : 고등학교 2학년 여름"
ElseIf Stage = 7 Then
    lblLev = "Level : 고등학교 2학년 가을"
ElseIf Stage = 8 Then
    lblLev = "Level : 고등학교 2학년 겨울"
ElseIf Stage = 9 Then
    lblLev = "Level : 고등학교 3학년 봄"
ElseIf Stage = 10 Then
    lblLev = "Level : 고등학교 3학년 여름"
ElseIf Stage = 11 Then
    lblLev = "Level : 고등학교 3학년 가을"
Else
    lblLev = "Level : 고등학교 3학년 겨울"
End If
End Sub

Private Sub ImgItem_Click(Index As Integer)
'아이템(Open Book) 사용부
If Index = 0 Then
    If (item(Index + 2) >= 1) And k <= 3 Then
        MsgBox Left(Question(k, Q), 1)
        item(Index + 2) = item(Index + 2) - 1
        lblQus_Click
    ElseIf (item(Index + 2) >= 1) And k >= 4 Then
        MsgBox Left(정답, 1)
        item(Index + 2) = item(Index + 2) - 1
        lblQus_Click
    Else
        MsgBox "아이템이 없습니다."
    End If
End If
End Sub

Private Sub lblQus_Click()
'아이템개수 재표시
For i = 0 To 1
    lblItem(i) = "보유 개수 : " & item(i + 2)
Next
End Sub

Private Function MakeQuestion(QuestionCode As Integer, QuestionLabel As Label)
If QuestionCode = 1 Then
'넌센스 문제
    Q = Int((11 * Rnd) + 2)
    Do Until QuestionB(k, Q) = 0
        Q = Int((11 * Rnd) + 2)
    Loop
    
    Select Case Q
    Case 2
        QuestionLabel = "<사고 넌센스>한 사람이 양을 90마리 가지고 있다. 강을 건너려고 하니 뱃사공이 건너편으로 건너가는 양의 반을 뱃삯으로 요규했다. 한 마리도 남기지 않고 건너간다면, 뱃사공에게 몇 마리의 양을 주어야 하는가?"
    Case 3
        QuestionLabel = "<기호 넌센스>1 × 2 × 3 × 4 × 5 × 6 × 7 × 8 × 9 × 10 × 11 × 12 × 13 × 14 × 15 × 16 × 17 × 18 × 19 × 20 · 0  × 21 × 22 × 23 × 24 × 25 × 26 × 27 × 28 × 29 × 30 × 31 × 32 × ... × 100의 값은 얼마인가?"
    Case 4
        QuestionLabel = "<그냥 넌센스>진짜 새의 이름은 무엇인가?"
    Case 5
        QuestionLabel = "<그냥 넌센스>금은 금인데 도둑 고양이에게 가장 어울리는 금은?"
    Case 6
        QuestionLabel = "<그냥 넌센스>고기를 먹을 때 마다 따라오는 개는?"
    Case 7
        QuestionLabel = "<그냥 넌센스>붉은색 길 위에 동전 하나가 떨어져있다. 그 동전의 이름은?"
    Case 8
        QuestionLabel = "<그냥 넌센스>배울 것 다 배워도 계속해서 배우라는 소리를 듣는 사람은?"
    Case 9
        QuestionLabel = "<그냥 넌센스>싸우다가 코가 빠져버린 두 코끼리의 이름은?"
    Case 10
        QuestionLabel = "<그냥 넌센스>허수아비의 아들 이름은?"
    Case 11
        QuestionLabel = "<그냥 넌센스>언제나 말다툼이 있는 곳은?"
    Case 12
        QuestionLabel = "<그냥 넌센스>슈퍼맨과 함께 하늘을 날고 있는 말의 이름은 무엇일까?"
    End Select
ElseIf QuestionCode = 2 Then
'과학문제
    Q = Int((8 * Rnd) + 1)
    Do Until QuestionB(k, Q) = 0
        Q = Int((8 * Rnd) + 1)
    Loop
    
    Select Case Q
    Case 1
        QuestionLabel = "시간 팽창은 아인슈타인의 특수 OOO이론의 내용중 하나이다." '상대성
    Case 2
        QuestionLabel = "산과 염기가 반응하여 물과 염을 생성하는 반응은 무엇인가?" '중화반응
    Case 3
        QuestionLabel = "RNA와 단백질로 이루어진 복합체로서 세포질 속에서 단백질을 합성하는 역할을 하는것은 무엇인가?" '리보솜
    Case 4
        QuestionLabel = "별의 절대등급을 세로축, 온도 또는 분광형, 색지수를 가로축에 잡고 그 관계를 나타낸 도표를 말하며 에이치아르도라고도부르는 이것은 무엇인가?(-는 포함하지 마시오.)"
    Case 5
        QuestionLabel = "유도 기전력의 방향은 코일 면을 통과하는 자속의 변화를 방해하는 방향으로 나타난다는 법칙을 OO의 법칙이라고 한다."
    Case 6
        QuestionLabel = "화합물을 구성하는 각 원자에 전체 전자를 일정한 방법으로 배분하였을 때, 그 원자가 가진 전하의 수를 무엇이라고 하는가?"
    Case 7
        QuestionLabel = "생물에는 환경에 대한 적응력이 있어, 자주 사용하는 기관은 발달하고 사용하지 않는 기관은 퇴화하여 없어지게 된다는 학설로 라마르크가 제창한 진화설은 무엇인가?"
    Case 8
        QuestionLabel = "가장 무른 것을 1로 하고 가장 단단한 것을 10으로 하여 10개의 광물에 굳기 순서대로 번호를 붙여 놓은 것을 무엇이라 하는가?"
    End Select
ElseIf QuestionCode = 3 Then
'사회문제
    Q = Int((9 * Rnd) + 1)
    Do Until QuestionB(k, Q) = 0
        Q = Int((9 * Rnd) + 1)
    Loop
    
    Select Case Q
    Case 1
        QuestionLabel = "자연의 세계에서 인간의 의지와 상관없이 발생하는 다양한 현상을 일컫는 말"
    Case 2
        QuestionLabel = "사회문화현상의 특징에는 가치함축성, OOOO, 개연성과 확률의 원리, 보편성과 특수성이 있다."
    Case 3
        QuestionLabel = "자연현상의 특징에는 몰가치성, OOOO, 필연성과 인과 법칙, 보편성이 있다."
    Case 4
        QuestionLabel = "사회 문화현상의 탐구에 필요한 태도중에서 탐구 과정에서 연구자가 자신의 주관적 가치나 편견, 이해관계등을 배제하고 사회문화 현상이 가진 사실로서의 특성만을 파악하는 태도는 무엇인가?"
    Case 5
        QuestionLabel = "머튼의 사회적응 방식의 유형중에서 문화적 목표와 수단을 모두 포기하는 유형은?"
    Case 6
        QuestionLabel = "사회집단중에서도 그 목표와 경계가 뚜렷하고, 구성원의 지위와 역할이 명확하며 목적달성을 위한 공식적인 규범과 절차가 체계적으로 구성되어 있는 집단은?"
    Case 7
        QuestionLabel = "조직 내의 근무기간에 따라 지위가 서열화 되는것은?"
    Case 8
        QuestionLabel = "문화의 속성중에서 의미하는 바가 문화는 한 사회의 구성원들이 공통으로 가지는 행동 및 사고방식이다. 인 속성을 문화의 OOO라고 한다."
    Case 9
        QuestionLabel = "인간이 공동체를 이루고 살아가면서 인위적으로 만들어내는 현상은?"
    End Select
'수치만 바꾸는 문제들
Else
    Q = Int((5 * Rnd) + 1)
    Select Case Q
    Case 1
        m = 10 * Int((3 * Rnd) + 1)
        L = Int((100 * Rnd) + 1)
        v = 100 * Int((5 * Rnd) + 1)
        QuestionLabel = "총신의 길이가" & L & "cm인 총에서 질량" & 4 * m & "g의 총알이 발사되 모래주머니에" & m & "cm만큼 들어가 박혔다. 이 때 평균 힘의 크기를 구하시오. (단, 처음 속도는" & v & "m/s이며 힘의 단위는 N이다.)"
        정답 = v ^ 2 / 5
    Case 2
        m = Int((10 * Rnd) + 1)
        QuestionLabel = "길이의 비가 1 : " & m & "이고, 단면적의 비가 " & m & " : 1인 같은 물질의 저항체 A,B를 직렬로 연결하였다. 저항체 A에 걸리는 전압이 10V이면 저항체 B에 걸리는 전압은 몇 V인가?"
        정답 = 10 * m ^ 2
    Case 3
        m = Int((20 * Rnd) + 1)
        QuestionLabel = "같은 재질,크기의 용수철 5개가 병렬로 연결되어있다. 이 용수철 하나의 용수철 상수가" & m & "N/m이라 할 때, 전체의 용수철 상수는 얼마인가?"
        정답 = 5 * m
    Case 4
        m = Int((10 * Rnd) + 1)
        v = Int((10 * Rnd) + 1)
        QuestionLabel = "물체의 질량이 " & m & "kg, 속도가 " & v & "m/s일 때, 이 물체의 운동량을 구하시오."
        정답 = v * m
    Case 5
        m = Int((10 * Rnd) + 1)
        QuestionLabel = "P의 전력을 전압 V로 보내고 있다. 전압을 (1 / " & m & ")배로 줄였을 때, 손실되는 전력은 몇배가 되는가?"
        정답 = m ^ 2
    End Select
End If
End Function
