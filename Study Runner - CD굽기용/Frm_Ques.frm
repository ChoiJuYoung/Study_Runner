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
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame FmeSel 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '����
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
            Name            =   "����"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14800597
         Caption         =   "��Ÿ ����"
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
            Name            =   "����"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14800597
         Caption         =   "��ȸ ����"
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
            Name            =   "����"
            Size            =   20.25
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
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����"
      CaptionEffects  =   0
   End
   Begin VB.TextBox txtAns 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      Height          =   270
      Left            =   1440
      TabIndex        =   2
      Text            =   "Answer"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label lblItem 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���� ���� : 1111"
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblItem 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�������� : 1111"
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
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "             ��ġ ������ ������ ������ ���� ���ڸ� �����ּ���.           ��� ����� ���� ���� ���ּ���."
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Label lblQus 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Label1"
      Height          =   975
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label lblLev 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
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
Private m As Integer, L As Integer, v As Integer '������ ���� ������
Private k As Integer, non As Integer, ���� As Integer

Private Sub Cmd_Ans_Click()
'���� �Ǻ�
If k <= 3 Then
    If txtAns = Question(k, Q) Then
        MsgBox "����!"
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
    If txtAns = ���� Then
        MsgBox "����!"
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
'�ʱ� ����
lblQus_Click
For i = 0 To 1
    ImgItem(i) = LoadResPicture(i + 108, vbResBitmap)
Next

If LifPoint = 0 Then
    End
End If

'Ŀ�ǵ��ư ���� �κ�
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
    CmdNon.Caption = "��Ÿ ���� (������ : " & (11 - Qq(1)) & ")"
End If

If Qq(2) = 8 Then
    CmdSci.Enabled = False
    Contri(6) = True
Else
    CmdSci.Caption = "���� ���� (������ : " & (8 - Qq(2)) & ")"
End If

If Qq(3) = 9 Then
    CmdSoc.Enabled = False
    Contri(8) = True
Else
    CmdSoc.Caption = "��ȸ ���� (������ : " & (9 - Qq(3)) & ")"
End If



If CmdNon.Enabled = False And CmdSci.Enabled = False And CmdSoc.Enabled = False Then
    k = 4
    Call MakeQuestion(k, lblQus)
    FmeSel.Visible = False
End If




DoEvents
'�������� ǥ�� �κ�
If Stage = 1 Then
    lblLev = "Level : ����б� 1�г� ��"
ElseIf Stage = 2 Then
    lblLev = "Level : ����б� 1�г� ����"
ElseIf Stage = 3 Then
    lblLev = "Level : ����б� 1�г� ����"
ElseIf Stage = 4 Then
    lblLev = "Level : ����б� 1�г� �ܿ�"
ElseIf Stage = 5 Then
    lblLev = "Level : ����б� 2�г� ��"
ElseIf Stage = 6 Then
    lblLev = "Level : ����б� 2�г� ����"
ElseIf Stage = 7 Then
    lblLev = "Level : ����б� 2�г� ����"
ElseIf Stage = 8 Then
    lblLev = "Level : ����б� 2�г� �ܿ�"
ElseIf Stage = 9 Then
    lblLev = "Level : ����б� 3�г� ��"
ElseIf Stage = 10 Then
    lblLev = "Level : ����б� 3�г� ����"
ElseIf Stage = 11 Then
    lblLev = "Level : ����б� 3�г� ����"
Else
    lblLev = "Level : ����б� 3�г� �ܿ�"
End If
End Sub

Private Sub ImgItem_Click(Index As Integer)
'������(Open Book) ����
If Index = 0 Then
    If (item(Index + 2) >= 1) And k <= 3 Then
        MsgBox Left(Question(k, Q), 1)
        item(Index + 2) = item(Index + 2) - 1
        lblQus_Click
    ElseIf (item(Index + 2) >= 1) And k >= 4 Then
        MsgBox Left(����, 1)
        item(Index + 2) = item(Index + 2) - 1
        lblQus_Click
    Else
        MsgBox "�������� �����ϴ�."
    End If
End If
End Sub

Private Sub lblQus_Click()
'�����۰��� ��ǥ��
For i = 0 To 1
    lblItem(i) = "���� ���� : " & item(i + 2)
Next
End Sub

Private Function MakeQuestion(QuestionCode As Integer, QuestionLabel As Label)
If QuestionCode = 1 Then
'�ͼ��� ����
    Q = Int((11 * Rnd) + 2)
    Do Until QuestionB(k, Q) = 0
        Q = Int((11 * Rnd) + 2)
    Loop
    
    Select Case Q
    Case 2
        QuestionLabel = "<��� �ͼ���>�� ����� ���� 90���� ������ �ִ�. ���� �ǳʷ��� �ϴ� ������ �ǳ������� �ǳʰ��� ���� ���� ������� ����ߴ�. �� ������ ������ �ʰ� �ǳʰ��ٸ�, �������� �� ������ ���� �־�� �ϴ°�?"
    Case 3
        QuestionLabel = "<��ȣ �ͼ���>1 �� 2 �� 3 �� 4 �� 5 �� 6 �� 7 �� 8 �� 9 �� 10 �� 11 �� 12 �� 13 �� 14 �� 15 �� 16 �� 17 �� 18 �� 19 �� 20 �� 0  �� 21 �� 22 �� 23 �� 24 �� 25 �� 26 �� 27 �� 28 �� 29 �� 30 �� 31 �� 32 �� ... �� 100�� ���� ���ΰ�?"
    Case 4
        QuestionLabel = "<�׳� �ͼ���>��¥ ���� �̸��� �����ΰ�?"
    Case 5
        QuestionLabel = "<�׳� �ͼ���>���� ���ε� ���� ����̿��� ���� ��︮�� ����?"
    Case 6
        QuestionLabel = "<�׳� �ͼ���>��⸦ ���� �� ���� ������� ����?"
    Case 7
        QuestionLabel = "<�׳� �ͼ���>������ �� ���� ���� �ϳ��� �������ִ�. �� ������ �̸���?"
    Case 8
        QuestionLabel = "<�׳� �ͼ���>��� �� �� ����� ����ؼ� ����� �Ҹ��� ��� �����?"
    Case 9
        QuestionLabel = "<�׳� �ͼ���>�ο�ٰ� �ڰ� �������� �� �ڳ����� �̸���?"
    Case 10
        QuestionLabel = "<�׳� �ͼ���>����ƺ��� �Ƶ� �̸���?"
    Case 11
        QuestionLabel = "<�׳� �ͼ���>������ �������� �ִ� ����?"
    Case 12
        QuestionLabel = "<�׳� �ͼ���>���۸ǰ� �Բ� �ϴ��� ���� �ִ� ���� �̸��� �����ϱ�?"
    End Select
ElseIf QuestionCode = 2 Then
'���й���
    Q = Int((8 * Rnd) + 1)
    Do Until QuestionB(k, Q) = 0
        Q = Int((8 * Rnd) + 1)
    Loop
    
    Select Case Q
    Case 1
        QuestionLabel = "�ð� ��â�� ���ν�Ÿ���� Ư�� OOO�̷��� ������ �ϳ��̴�." '��뼺
    Case 2
        QuestionLabel = "��� ���Ⱑ �����Ͽ� ���� ���� �����ϴ� ������ �����ΰ�?" '��ȭ����
    Case 3
        QuestionLabel = "RNA�� �ܹ����� �̷���� ����ü�μ� ������ �ӿ��� �ܹ����� �ռ��ϴ� ������ �ϴ°��� �����ΰ�?" '������
    Case 4
        QuestionLabel = "���� �������� ������, �µ� �Ǵ� �б���, �������� �����࿡ ��� �� ���踦 ��Ÿ�� ��ǥ�� ���ϸ� ����ġ�Ƹ�������θ��� �̰��� �����ΰ�?(-�� �������� ���ÿ�.)"
    Case 5
        QuestionLabel = "���� �������� ������ ���� ���� ����ϴ� �ڼ��� ��ȭ�� �����ϴ� �������� ��Ÿ���ٴ� ��Ģ�� OO�� ��Ģ�̶�� �Ѵ�."
    Case 6
        QuestionLabel = "ȭ�չ��� �����ϴ� �� ���ڿ� ��ü ���ڸ� ������ ������� ����Ͽ��� ��, �� ���ڰ� ���� ������ ���� �����̶�� �ϴ°�?"
    Case 7
        QuestionLabel = "�������� ȯ�濡 ���� �������� �־�, ���� ����ϴ� ����� �ߴ��ϰ� ������� �ʴ� ����� ��ȭ�Ͽ� �������� �ȴٴ� �м��� �󸶸�ũ�� ��â�� ��ȭ���� �����ΰ�?"
    Case 8
        QuestionLabel = "���� ���� ���� 1�� �ϰ� ���� �ܴ��� ���� 10���� �Ͽ� 10���� ������ ���� ������� ��ȣ�� �ٿ� ���� ���� �����̶� �ϴ°�?"
    End Select
ElseIf QuestionCode = 3 Then
'��ȸ����
    Q = Int((9 * Rnd) + 1)
    Do Until QuestionB(k, Q) = 0
        Q = Int((9 * Rnd) + 1)
    Loop
    
    Select Case Q
    Case 1
        QuestionLabel = "�ڿ��� ���迡�� �ΰ��� ������ ������� �߻��ϴ� �پ��� ������ ���´� ��"
    Case 2
        QuestionLabel = "��ȸ��ȭ������ Ư¡���� ��ġ���༺, OOOO, �������� Ȯ���� ����, ������ Ư������ �ִ�."
    Case 3
        QuestionLabel = "�ڿ������� Ư¡���� ����ġ��, OOOO, �ʿ����� �ΰ� ��Ģ, ������ �ִ�."
    Case 4
        QuestionLabel = "��ȸ ��ȭ������ Ž���� �ʿ��� �µ��߿��� Ž�� �������� �����ڰ� �ڽ��� �ְ��� ��ġ�� ���, ���ذ������ �����ϰ� ��ȸ��ȭ ������ ���� ��Ƿμ��� Ư������ �ľ��ϴ� �µ��� �����ΰ�?"
    Case 5
        QuestionLabel = "��ư�� ��ȸ���� ����� �����߿��� ��ȭ�� ��ǥ�� ������ ��� �����ϴ� ������?"
    Case 6
        QuestionLabel = "��ȸ�����߿����� �� ��ǥ�� ��谡 �ѷ��ϰ�, �������� ������ ������ ��Ȯ�ϸ� �����޼��� ���� �������� �Թ��� ������ ü�������� �����Ǿ� �ִ� ������?"
    Case 7
        QuestionLabel = "���� ���� �ٹ��Ⱓ�� ���� ������ ����ȭ �Ǵ°���?"
    Case 8
        QuestionLabel = "��ȭ�� �Ӽ��߿��� �ǹ��ϴ� �ٰ� ��ȭ�� �� ��ȸ�� ���������� �������� ������ �ൿ �� ������̴�. �� �Ӽ��� ��ȭ�� OOO��� �Ѵ�."
    Case 9
        QuestionLabel = "�ΰ��� ����ü�� �̷�� ��ư��鼭 ���������� ������ ������?"
    End Select
'��ġ�� �ٲٴ� ������
Else
    Q = Int((5 * Rnd) + 1)
    Select Case Q
    Case 1
        m = 10 * Int((3 * Rnd) + 1)
        L = Int((100 * Rnd) + 1)
        v = 100 * Int((5 * Rnd) + 1)
        QuestionLabel = "�ѽ��� ���̰�" & L & "cm�� �ѿ��� ����" & 4 * m & "g�� �Ѿ��� �߻�� ���ָӴϿ�" & m & "cm��ŭ �� ������. �� �� ��� ���� ũ�⸦ ���Ͻÿ�. (��, ó�� �ӵ���" & v & "m/s�̸� ���� ������ N�̴�.)"
        ���� = v ^ 2 / 5
    Case 2
        m = Int((10 * Rnd) + 1)
        QuestionLabel = "������ �� 1 : " & m & "�̰�, �ܸ����� �� " & m & " : 1�� ���� ������ ����ü A,B�� ���ķ� �����Ͽ���. ����ü A�� �ɸ��� ������ 10V�̸� ����ü B�� �ɸ��� ������ �� V�ΰ�?"
        ���� = 10 * m ^ 2
    Case 3
        m = Int((20 * Rnd) + 1)
        QuestionLabel = "���� ����,ũ���� ���ö 5���� ���ķ� ����Ǿ��ִ�. �� ���ö �ϳ��� ���ö �����" & m & "N/m�̶� �� ��, ��ü�� ���ö ����� ���ΰ�?"
        ���� = 5 * m
    Case 4
        m = Int((10 * Rnd) + 1)
        v = Int((10 * Rnd) + 1)
        QuestionLabel = "��ü�� ������ " & m & "kg, �ӵ��� " & v & "m/s�� ��, �� ��ü�� ����� ���Ͻÿ�."
        ���� = v * m
    Case 5
        m = Int((10 * Rnd) + 1)
        QuestionLabel = "P�� ������ ���� V�� ������ �ִ�. ������ (1 / " & m & ")��� �ٿ��� ��, �սǵǴ� ������ ��谡 �Ǵ°�?"
        ���� = m ^ 2
    End Select
End If
End Function
