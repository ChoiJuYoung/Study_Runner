VERSION 5.00
Begin VB.Form FrmRank 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4215
   Icon            =   "FrmRank.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   885
   ScaleWidth      =   4215
   StartUpPosition =   2  '화면 가운데
   Begin Study_Runner.jcbutton jcbutton1 
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
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
      Caption         =   "저장"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "이름을 입력하세요."
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblPoint 
      BackStyle       =   0  '투명
      Caption         =   "점수 : "
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "FrmRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Made By 서령고등학교 최 주영

Private Sub Form_Load()
'점수 표시
lblPoint = "점수 : " & SPoint
End Sub

Private Sub jcbutton1_Click()
'랭크 넣기 !
Dim temp1 As String, temp2 As Integer
temp1 = TxtName
temp2 = SPoint
If temp2 >= Val(RankValue(10)) Then
    RankValue(10) = temp2
    RankName(10) = temp1
    Dim j As Integer
    For i = 1 To 15
        For j = 1 To 9
            If RankValue(j) < RankValue(j + 1) Then
                temp2 = Val(RankValue(j))
                temp1 = RankName(j)
                RankValue(j) = Val(RankValue(j + 1))
                RankName(j) = RankName(j + 1)
                RankValue(j + 1) = temp2
                RankName(j + 1) = temp1
            End If
        Next
    Next
End If

'랭크 넣기 종료후 게임 종료
For i = 1 To 10
    SaveSetting "StudyRun", "RankName", i, RankName(i)
    SaveSetting "StudyRun", "RankValue", i, RankValue(i)
Next
MsgBox "랭킹 갱신이 완료되었습니다."
DoEvents
End
End Sub

Private Sub TxtName_Change()
'이름 글자수 제한
If Len(TxtName) >= 10 Then
    TxtName = Left(TxtName, 10)
End If
End Sub
