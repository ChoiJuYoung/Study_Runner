VERSION 5.00
Begin VB.Form Frm_Help 
   BackColor       =   &H00FFFFFF&
   Caption         =   "EBSi - �Խ� ����"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   Icon            =   "Frm_Help.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   7200
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.ListBox List1 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����"
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
      BackStyle       =   0  '����
      Caption         =   "EBSi �Խ� ���� [����Ŭ���ϸ� ������ ���� �������� �̵��մϴ�.]"
      BeginProperty Font 
         Name            =   "����"
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
'EBSi �Խ����� �Ľ̺�
Me.Show
On Error GoTo PSEr:
Dim winhttp As New WinHttpRequest
winhttp.Open "GET", "http://www.ebsi.co.kr/ebs/ent/enta/retrieveEntNwsLst.ebs"
winhttp.Send

Dim i As Integer
Dim ����(15) As String
For i = 1 To 15
    ����(i) = Split(Split(Split(StrConv(winhttp.ResponseBody, vbUnicode), "<span class=" & """" & "cP" & """" & " onclick=" & """" & "javascript:fncViewArticle('")(i), "title=")(1), ">")(0)
    ����(i) = Replace(����(i), "&#039;", "'")
    List1.AddItem (����(i))
Next

Exit Sub

PSEr:
DoEvents
MsgBox "���ͳ��� ����Ǿ� ���� �ʽ��ϴ�."
Unload Me
End Sub

Private Sub List1_DblClick()
'EBSi �̵���
Shell "explorer.exe http://www.ebsi.co.kr/ebs/ent/enta/retrieveEntNwsLst.ebs"
End Sub
