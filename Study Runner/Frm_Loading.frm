VERSION 5.00
Begin VB.Form Frm_Loading 
   BackColor       =   &H00000000&
   BorderStyle     =   0  '없음
   ClientHeight    =   10065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   8880
      Top             =   9240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   9240
   End
   Begin VB.Label lblVer 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "버전을 확인 중입니다 ..."
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9000
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   480
      Picture         =   "Frm_Loading.frx":0000
      Top             =   1080
      Width           =   10500
   End
End
Attribute VB_Name = "Frm_Loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'변수드을~
Private k As Integer, VersionP As String
Const Version = "1.0"
Private winhttp As New WinHttpRequest
Private Sub Form_Load()
'랭킹 값 가져오기
For i = 1 To 10
    RankName(i) = GetSetting("StudyRun", "RankName", i)
    RankValue(i) = GetSetting("StudyRun", "RankValue", i)
Next

'Timer1 Open
i = 1
k = 0
Timer1.Enabled = True

'강제 오픈
Me.Show

On Error GoTo ABC:
'버전 파싱 부분
winhttp.Open "GET", "http://jinie.woobi.co.kr/mcbldr/page/folder_company/page_company/"
winhttp.Send vbNullString
VersionP = Split(StrConv(winhttp.ResponseBody, vbUnicode), "Study")(1)

'인터넷 접속 관련 오류시
ABC:
VersionP = "1.0"
End Sub

Private Sub Timer1_Timer()
'버전 확인부
k = k + 1
If k = 5 Then
    Timer2.Enabled = False
    If Version = VersionP Then
        lblVer = "버전 확인 완료. 최신 버전입니다."
    Else
        MsgBox "최신 버전이 아닙니다."
        MsgBox "hajuu96123@naver.com에게 문의해 주세요!"
        End
    End If
ElseIf k >= 8 Then
    Frm_Main.Show
    Unload Me
End If
End Sub

Private Sub Timer2_Timer()
'그냥 보기 좋은 용도
If i = 1 Then
    lblVer = "버전을 확인중입니다 ..."
    i = i + 1
ElseIf i = 2 Then
    lblVer = "버전을 확인중입니다 .."
    i = i + 1
ElseIf i = 3 Then
    lblVer = "버전을 확인중입니다 ."
    i = i + 1
Else
    lblVer = "버전을 확인중입니다 .."
    i = 1
End If
End Sub
