VERSION 5.00
Begin VB.Form FrmShop 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Item Shop"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6045
   Icon            =   "FrmShop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   6045
   StartUpPosition =   2  'ȭ�� ���
   Begin Study_Runner.jcbutton jcbutton1 
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "���� �Ϸ�"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Frame FmeExp 
      BackColor       =   &H00C0C0FF&
      Caption         =   "������ ����"
      Height          =   2175
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
      Begin VB.Label lblExp 
         BackStyle       =   0  '����
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Label lblItem 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�������� : 000��"
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Life : 100p"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image ImgItem 
      Height          =   375
      Index           =   3
      Left            =   600
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblItem 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�������� : 000��"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   "���������� : 30p"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Image ImgItem 
      Height          =   480
      Index           =   2
      Left            =   240
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label lblItem 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�������� : 000��"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "Open Book : 50p"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image ImgItem 
      Height          =   465
      Index           =   1
      Left            =   360
      Top             =   960
      Width           =   750
   End
   Begin VB.Label lblItem 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�������� : 000��"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblPoint 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���� ����Ʈ : 1000"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���찳 : 45p"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image ImgItem 
      Height          =   645
      Index           =   0
      Left            =   360
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "FrmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private price(1 To 4) As Integer

Private Sub Form_Load()
'�ʱ� �� ǥ��
lblPoint = "���� ����Ʈ : " & GPoint
item(4) = LifPoint

For i = 0 To 3
    ImgItem(i) = LoadResPicture(i + 107, vbResBitmap)
    lblItem(i) = "�������� : " & item(i + 1)
Next

'���� ǥ��
price(1) = 45
price(2) = 50
price(3) = 30
price(4) = 100
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'lblExp �ʱ�ȭ
lblExp = ""
End Sub

Private Sub ImgItem_Click(Index As Integer)
'���� Ȯ��
If GPoint >= price(Index + 1) Then
    GPoint = GPoint - price(Index + 1)
    item(Index + 1) = item(Index + 1) + 1
    lblItem(Index) = "�������� : " & item(Index + 1)
    lblPoint = "���� ����Ʈ : " & GPoint
Else
    MsgBox "GPoint�� �����մϴ� �Ф�"
End If

If item(4) >= 10 Then
    Contri(5) = True
End If
End Sub

Private Sub jcbutton1_Click()
'���� �Ϸ�. ������ ����
LifPoint = item(4)
Unload Me
Frm_Save.Show
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���찳 ����
lblExp = "���찳�� ���� ȭ�鳻 ��� ��ֹ��� �����ݴϴ�. ���� �����ð����� ��ֹ��� �������� �ʵ��� ���ݴϴ�."
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Open Book ����
lblExp = "Open Book�� ��ֹ��� �ɷ� ������ Ǯ �� ���� ������ ���ڸ� ��Ʈ�� �����մϴ�."
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���������� ����
lblExp = "������������ ������ Ʋ���� �� �ڵ����� ���˴ϴ�. ������������ �� �� �̻� �������� �� ������ Ʋ���� ��� ����� ���� ���� �ڵ����� �� �� �� ��ȸ�� �־����ϴ�."
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Life ����
lblExp = "Life�� �Ѱ� �÷��ݴϴ�."
End Sub
