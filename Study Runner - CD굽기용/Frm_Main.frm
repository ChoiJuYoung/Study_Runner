VERSION 5.00
Begin VB.Form Frm_Main 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  '���� ����
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
   StartUpPosition =   2  'ȭ�� ���
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
      Caption         =   "���� ����"
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
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "��ŷ ����"
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
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "�Խ� News"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
'CopyRight �� ����
FrmCopy.Show
End Sub

Private Sub Cmd_Help_Click()
'EBSi �Խ����� �� ����
Frm_Help.Show
End Sub

Private Sub Cmd_Load_Click()
'�ҷ����� �� ����
Call load(Frm_Load, Frm_Main)
DoEvents
End Sub

Private Sub Cmd_New_Click()
'�� ���� ���� ^_^ ;
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
'��ŷ���� �� ����
FrmShowRank.Show
End Sub

Private Sub Form_Load()
'���� ���������
Call load(Frm_Loading, Frm_Loading)
DoEvents
'������ �ʱ�ȭ
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

'���� �Է�
''�ͼ���
Question(1, 2) = 30
Question(1, 3) = 0
Question(1, 4) = "����"
Question(1, 5) = "�߱ݾ߱�"
Question(1, 6) = "�̾��ð�"
Question(1, 7) = "ȫ�浿��"
Question(1, 8) = "���"
Question(1, 9) = "��������"
Question(1, 10) = "���"
Question(1, 11) = "�渶��"
Question(1, 12) = "���۸�����"

''����
Question(2, 1) = "��뼺"
Question(2, 2) = "��ȭ����"
Question(2, 3) = "������"
Question(2, 4) = "HR��"
Question(2, 5) = "����"
Question(2, 6) = "��ȭ��"
Question(2, 7) = "��ҿ뼳"
Question(2, 8) = "�𽺱����"

''��ȸ
Question(3, 1) = "�ڿ�����"
Question(3, 2) = "������Ģ"
Question(3, 3) = "�����Ģ"
Question(3, 4) = "�������µ�"
Question(3, 5) = "������"
Question(3, 6) = "��ȸ����"
Question(3, 7) = "��������"
Question(3, 8) = "������"
Question(3, 9) = "��ȸ��ȭ����"

End Sub
