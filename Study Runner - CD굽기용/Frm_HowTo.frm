VERSION 5.00
Begin VB.Form Frm_HowTo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ٽô� �� â�� ����� �ʽ��ϴ�."
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
      BackStyle       =   0  '����
      Caption         =   "YŰ�� ������ ���찳�� ���˴ϴ�."
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '����
      Caption         =   "���� ����� EscŰ�� �����ø� �˴ϴ�."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '����
      Caption         =   "�ٰ����� LOL�ΰ� ���ϼž� �մϴ�."
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '����
      Caption         =   "����� ��� ���Ǹ� ���� ������ �˴ϴ�."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   "3�� ���� : Shift + Enter"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "2�� ���� : Ctrl + Space Bar"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "1�� ���� : Space Bar"
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
Private Checking As String '������Ʈ�� ������ Ȯ�� ����

Private Sub Form_Load()
'�ٽ� ���� ���� Ȯ�κ�
Checking = GetSetting("STRE", "STRE", "STRE")
DoEvents
If Checking = "Yes" Then
    jcbutton2_Click
End If
End Sub
Private Sub jcbutton2_Click()
'���� �����
If Check1.Value = 1 Then
    SaveSetting "STRE", "STRE", "STRE", "Yes"
End If
Frm_Play.Show
Unload Me
End Sub
