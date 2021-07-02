VERSION 5.00
Begin VB.Form Frm_Help 
   BackColor       =   &H00FFFFFF&
   Caption         =   "EBSi - 입시 뉴스"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   Icon            =   "Frm_Help.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   7200
   StartUpPosition =   2  '화면 가운데
   Begin VB.ListBox List1 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "EBSi 입시 뉴스 [더블클릭하면 오늘의 뉴스 페이지로 이동합니다.]"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "Frm_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'EBSi 입시정보 파싱부
Me.Show
On Error GoTo PSEr:
Dim winhttp As New WinHttpRequest
winhttp.Open "GET", "http://www.ebsi.co.kr/ebs/ent/enta/retrieveEntNwsLst.ebs"
winhttp.Send

Dim i As Integer
Dim 네임(15) As String
For i = 1 To 15
    네임(i) = Split(Split(Split(StrConv(winhttp.ResponseBody, vbUnicode), "<span class=" & """" & "cP" & """" & " onclick=" & """" & "javascript:fncViewArticle('")(i), "title=")(1), ">")(0)
    네임(i) = Replace(네임(i), "&#039;", "'")
    List1.AddItem (네임(i))
Next

Exit Sub

PSEr:
DoEvents
MsgBox "인터넷이 연결되어 있지 않습니다."
Unload Me
End Sub

Private Sub List1_DblClick()
'EBSi 이동부
Shell "explorer.exe http://www.ebsi.co.kr/ebs/ent/enta/retrieveEntNwsLst.ebs"
End Sub
