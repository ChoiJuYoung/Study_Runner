VERSION 5.00
Begin VB.Form Frm_Play 
   BorderStyle     =   0  '����
   Caption         =   "Form2"
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9930
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Timer TimItem 
      Enabled         =   0   'False
      Interval        =   5500
      Left            =   8280
      Top             =   7560
   End
   Begin VB.PictureBox Pic_School 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   11400
      Picture         =   "Frm_Play.frx":0000
      ScaleHeight     =   600
      ScaleWidth      =   990
      TabIndex        =   19
      Top             =   7320
      Width           =   1050
   End
   Begin VB.PictureBox Pic_SchoolM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   11520
      Picture         =   "Frm_Play.frx":1F84
      ScaleHeight     =   600
      ScaleWidth      =   990
      TabIndex        =   18
      Top             =   7920
      Width           =   1050
   End
   Begin VB.PictureBox Pic_SmiM2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   11280
      Picture         =   "Frm_Play.frx":3F08
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   17
      Top             =   6120
      Width           =   435
   End
   Begin VB.PictureBox Pic_Smi1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   11280
      Picture         =   "Frm_Play.frx":46B8
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   16
      Top             =   6240
      Width           =   435
   End
   Begin VB.PictureBox Pic_SmiM1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   11280
      Picture         =   "Frm_Play.frx":4E68
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   15
      Top             =   6360
      Width           =   435
   End
   Begin VB.PictureBox Pic_Smi2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   11280
      Picture         =   "Frm_Play.frx":5618
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   14
      Top             =   6480
      Width           =   435
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6840
      Top             =   7560
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   7320
      Top             =   7560
   End
   Begin VB.PictureBox B_Obj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   13800
      Picture         =   "Frm_Play.frx":5DC8
      ScaleHeight     =   375
      ScaleWidth      =   930
      TabIndex        =   13
      Top             =   7680
      Width           =   990
   End
   Begin VB.PictureBox B_ObjM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   13800
      Picture         =   "Frm_Play.frx":7068
      ScaleHeight     =   375
      ScaleWidth      =   930
      TabIndex        =   12
      Top             =   7680
      Width           =   990
   End
   Begin VB.PictureBox G_ObjM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   12240
      Picture         =   "Frm_Play.frx":8308
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   9
      Top             =   8880
      Width           =   810
   End
   Begin VB.PictureBox G_Obj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   14520
      Picture         =   "Frm_Play.frx":A0FC
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   8
      Top             =   8880
      Width           =   810
   End
   Begin VB.PictureBox PicMCha 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1995
      Index           =   0
      Left            =   10680
      ScaleHeight     =   1935
      ScaleWidth      =   1500
      TabIndex        =   7
      Top             =   4920
      Width           =   1560
      Begin VB.Label lblLeft 
         Caption         =   "Label1"
         Height          =   1335
         Index           =   3
         Left            =   480
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox PicCha 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1995
      Index           =   0
      Left            =   10680
      ScaleHeight     =   1935
      ScaleWidth      =   1500
      TabIndex        =   6
      Top             =   4920
      Width           =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7800
      Top             =   7560
   End
   Begin VB.PictureBox PicCha 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1995
      Index           =   2
      Left            =   10680
      ScaleHeight     =   1935
      ScaleWidth      =   1500
      TabIndex        =   5
      Top             =   4920
      Width           =   1560
      Begin VB.Label lblLeft 
         Caption         =   "Label2"
         Height          =   1815
         Index           =   1
         Left            =   480
         TabIndex        =   22
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox PicMCha 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1995
      Index           =   2
      Left            =   10680
      ScaleHeight     =   1935
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   4920
      Width           =   1560
   End
   Begin VB.PictureBox PicMCha 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1995
      Index           =   1
      Left            =   10680
      ScaleHeight     =   1935
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   4920
      Width           =   1560
      Begin VB.Label lblLeft 
         Caption         =   "lblL"
         Height          =   1695
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox PicCha 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1995
      Index           =   1
      Left            =   10680
      ScaleHeight     =   1935
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   4920
      Width           =   1560
   End
   Begin VB.PictureBox Pic_Scr 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6750
      Left            =   0
      Picture         =   "Frm_Play.frx":BEF0
      ScaleHeight     =   6690
      ScaleWidth      =   9915
      TabIndex        =   1
      Top             =   600
      Width           =   9975
      Begin VB.Image ImgItem 
         Height          =   645
         Left            =   4680
         Top             =   120
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lblErase 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "���찳 : 00000��"
         Height          =   255
         Left            =   8160
         TabIndex        =   26
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "���� On / Off : T"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblPoint 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "Point : 0"
         Height          =   255
         Left            =   8160
         TabIndex        =   24
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblStg 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "Stage : 1"
         Height          =   255
         Left            =   8160
         TabIndex        =   20
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblLPoint 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "Life Point : 0"
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
         Left            =   8160
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblLife 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "Life : 3"
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
         Left            =   8160
         TabIndex        =   10
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox Pic_Status 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "Frm_Play"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Const SRCAND = &H8800C6
Private Const SRCPAINT = &HEE0086
Private Const SRCCOPY = &HCC0020
'�� BitBlt�� ����� �����
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer 'GetAsyncKeyState API ����
Dim NMCha As PictureBox, MCha As PictureBox, Gobj As PictureBox, Bobj As PictureBox 'bitblt���� ������ �׸�
Private ChaNum As Integer '�޸��� ���� Index��
Private PlayT As Integer '�÷��� �ð�
Private JB As Boolean '���� ����
Private JSum As Boolean '���� ���� ����
Private JHei As Integer '���� ����
Private JB2 As Boolean '������ ����
Private JB3 As Boolean '�������� ����
Private Hei As Integer '����
Private Lim1 As Integer, Lim2 As Integer '���� ����Ʈ
Private k As Integer '�����
Private Gobw(14) As Long, Bobw(14) As Long    'obj���� x��
Private Goby(14) As Long, Boby(14) As Long   'obj���� y��
Private GobwVis(14) As Boolean, BobwVis(14) As Boolean   '���� ����
Private BodyLeft(3) As Integer, BodyTop(3) As Integer   '�׸� �� ��ũ��
Private t As Integer 'obj���� �ӵ� ��
Private h 'ShockWave �̸�
Private it As Boolean '���찳 ��� �����


Private Sub Form_Load()
'������� = Second Run
Set h = Controls.Add("shockwaveflash.shockwaveflash", "hh")

With h
.Movie = "http://down5.snoin.kr/datagf/hajuu96/1pv0s/ice.swf"
.Move 0, -180, 5000, 4000
End With

'�� �ʱ� ����
ImgItem = LoadResPicture(107, vbResBitmap)
Me.Show
Pic_Scr.SetFocus

PicCha(0) = LoadResPicture(106, vbResBitmap)
PicMCha(0) = LoadResPicture(101, vbResBitmap)
PicCha(1) = LoadResPicture(102, vbResBitmap)
PicMCha(1) = LoadResPicture(104, vbResBitmap)
PicCha(2) = LoadResPicture(103, vbResBitmap)
PicMCha(2) = LoadResPicture(105, vbResBitmap)

lblStg = "Stage : " & Stage
PlayT = 0
ChaNum = 1
JB = False
JHei = 300

'��ũ�� �� ����
For i = 1 To 3
    BodyLeft(i) = Int(lblLeft(i).Left / 15)
    BodyTop(i) = Int(lblLeft(i).Top / 15)
Next

For i = 1 To 10
    Gobw(i) = 700
    Bobw(i) = 700
    Goby(i) = 200
    Boby(i) = 275
Next
BitBlt Pic_Status.hdc, 0, 7, Pic_SmiM1.Width, Pic_SmiM1.Height, Pic_SmiM1.hdc, 0, 0, SRCPAINT
BitBlt Pic_Status.hdc, 0, 7, Pic_Smi1.Width, Pic_Smi1.Height, Pic_Smi1.hdc, 0, 0, SRCAND
BitBlt Pic_Status.hdc, 597, 0, Pic_SchoolM.Width, Pic_SchoolM.Height, Pic_SchoolM.hdc, 0, 0, SRCPAINT
BitBlt Pic_Status.hdc, 597, 0, Pic_School.Width, Pic_School.Height, Pic_School.hdc, 0, 0, SRCAND

'�Լ� ����
DoEvents

Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
End Sub

Private Sub Timer1_Timer()
'������Ʈ ����
For i = 1 To 10
    If GobwVis(i) = True Then
        If Gobw(i) > -81 Then
            Gobw(i) = Gobw(i) - 15
        Else
            GobwVis(i) = False
        End If
    End If

    If BobwVis(i) = True Then
        If Bobw(i) > 0 Then
            Bobw(i) = Bobw(i) - 15
        Else
            BobwVis(i) = False
            SPoint = Val(SPoint) + 1
            If SPoint >= 4000 Then
                Contri(4) = True
            ElseIf SPoint >= 2000 Then
                Contri(3) = True
            ElseIf SPoint >= 1000 Then
                Contri(2) = True
            End If
        End If
    End If
Next

'ĳ���� ����
If JB = False Then
    JHei = 300
    Set NMCha = PicCha(ChaNum)
    Set MCha = PicMCha(ChaNum)
    If ChaNum = 1 Then
        ChaNum = ChaNum + 1
    Else
        ChaNum = 1
    End If
Else
    ChaNum = 0
    Set NMCha = PicCha(ChaNum)
    Set MCha = PicMCha(ChaNum)
    If JSum = False Then
        If k > 0 Then
            k = k - 2.5
            JHei = JHei - k
        Else
             k = k + 2.5
            JHei = JHei + k
            JSum = True
        End If
    Else
        If JHei < 300 - Lim2 Then
            k = k + 2.5
            JHei = JHei + k
        Else
            JHei = 300
            JB = False
            JB2 = False
            JB3 = False
            k = 0
        End If
    End If
End If

'�ε��� �Ǻ�

If JB = False Then
    If ChaNum = 1 Then
        For i = 1 To 13
            If GobwVis(i) = True Then
                If ((Gobw(i) <= 30 + BodyLeft(1) And Gobw(i) + G_ObjM.Width / 20 >= 30 + BodyLeft(1)) Or (Gobw(i) <= 30 + BodyLeft(1) + lblLeft(1).Width / 20 And Gobw(i) + G_ObjM.Width / 20 >= 30 + BodyLeft(1) + lblLeft(1).Width / 20) Or (30 + BodyLeft(1) <= Gobw(i) And Gobw(i) + G_ObjM.Width / 20 <= 30 + BodyLeft(1) + lblLeft(1).Width / 20)) _
                And _
                ((Goby(i) + G_ObjM.Height / 20 >= JHei And Goby(i) <= JHei) Or (Goby(i) + G_ObjM.Height / 20 >= JHei + lblLeft(1).Height / 20 And Goby(i) <= JHei + lblLeft(1).Height / 20) Or (Goby(i) >= JHei And Goby(i) + G_ObjM.Height / 20 <= JHei + lblLeft(1).Height / 20)) Then
                    Gobw(i) = 700
                    GobwVis(i) = False
                    GPoint = GPoint + 1
                End If
            End If

            If BobwVis(i) = True Then
                If ((Bobw(i) <= 30 + BodyLeft(1) And Bobw(i) + (B_ObjM.Width / 20) >= 30 + BodyLeft(1)) Or (Bobw(i) <= 30 + BodyLeft(1) + lblLeft(1).Width / 20 And Bobw(i) + (B_ObjM.Width / 20) >= 30 + BodyLeft(1) + lblLeft(1).Width / 20) Or (30 + BodyLeft(1) <= Bobw(i) And Bobw(i) + (B_ObjM.Width / 20) <= 30 + BodyLeft(1) + lblLeft(1).Width / 20)) _
                And _
                ((Boby(i) + B_ObjM.Height / 20 >= JHei And Boby(i) <= JHei) Or (Boby(i) + B_ObjM.Height / 20 >= JHei + lblLeft(1).Height / 20 And Boby(i) <= JHei + lblLeft(1).Height / 20) Or (Boby(i) >= JHei And Boby(i) + B_ObjM.Height / 20 <= JHei + lblLeft(1).Height / 20)) Then
                    Timer1.Enabled = False
                    Timer2.Enabled = False
                    Timer3.Enabled = False
                    Frm_Ques.Show
                    For p = 1 To 13
                        If BobwVis(p) = True Then
                            Bobw(p) = 700
                            BobwVis(p) = False
                        End If
                    Next
                    BitBlt Pic_Scr.hdc, 30, JHei, PicMCha(ChaNum).Width, PicMCha(ChaNum).Height, PicMCha(ChaNum).hdc, 0, 0, SRCPAINT 'Character
                    BitBlt Pic_Scr.hdc, 30, JHei, PicCha(ChaNum).Width, PicCha(ChaNum).Height, PicCha(ChaNum).hdc, 0, 0, SRCAND 'Character
                    Exit Sub
                End If
            End If
        Next
    Else
        For i = 1 To 13
            If GobwVis(i) = True Then
                If ((Gobw(i) <= 30 And Gobw(i) + G_ObjM.Width / 20 >= 30) Or (Gobw(i) <= 30 + PicCha(ChaNum).Width / 20 And Gobw(i) + G_ObjM.Width / 20 >= 30 + PicCha(ChaNum).Width / 20) Or (30 <= Gobw(i) And Gobw(i) + G_ObjM.Width / 20 <= 30 + PicCha(ChaNum).Width / 20)) _
                And _
                ((Goby(i) + G_ObjM.Height / 20 >= JHei + BodyTop(2) And Goby(i) <= JHei + BodyTop(2)) Or (Goby(i) + G_ObjM.Height / 20 >= JHei + BodyTop(2) + PicCha(ChaNum).Height / 20 And Goby(i) <= JHei + BodyTop(2) + PicCha(ChaNum).Height / 20) Or (Goby(i) >= JHei + BodyTop(2) And Goby(i) + G_ObjM.Height / 20 <= JHei + BodyTop(2) + PicCha(ChaNum).Height / 20)) Then
                    Gobw(i) = 700
                    GobwVis(i) = False
                    GPoint = GPoint + 1
                End If
            End If

            If BobwVis(i) = True Then
                If ((Bobw(i) <= 30 And Bobw(i) + (B_ObjM.Width / 20) >= 30) Or (Bobw(i) <= 30 + PicCha(ChaNum).Width / 20 And Bobw(i) + (B_ObjM.Width / 20) >= 30 + PicCha(ChaNum).Width / 20) Or (30 <= Bobw(i) And Bobw(i) + (B_ObjM.Width / 20) <= 30 + PicCha(ChaNum).Width / 20)) _
                And _
                ((Boby(i) + B_ObjM.Height / 20 >= JHei + BodyTop(2) And Boby(i) <= JHei + BodyTop(2)) Or (Boby(i) + B_ObjM.Height / 20 >= JHei + BodyTop(2) + PicCha(ChaNum).Height / 20 And Boby(i) <= JHei + BodyTop(2) + PicCha(ChaNum).Height / 20) Or (Boby(i) >= JHei + BodyTop(2) And Boby(i) + B_ObjM.Height / 20 <= JHei + BodyTop(2) + PicCha(ChaNum).Height / 20)) Then
                    Timer1.Enabled = False
                    Timer2.Enabled = False
                    Timer3.Enabled = False
                    Frm_Ques.Show
                    For p = 1 To 13
                        If BobwVis(p) = True Then
                            Bobw(p) = 700
                            BobwVis(p) = False
                        End If
                    Next
                    BitBlt Pic_Scr.hdc, 30, JHei, PicMCha(ChaNum).Width, PicMCha(ChaNum).Height, PicMCha(ChaNum).hdc, 0, 0, SRCPAINT 'Character
                    BitBlt Pic_Scr.hdc, 30, JHei, PicCha(ChaNum).Width, PicCha(ChaNum).Height, PicCha(ChaNum).hdc, 0, 0, SRCAND 'Character
                    Exit Sub
                End If
            End If
        Next
    End If
Else
    For i = 1 To 13
        If GobwVis(i) = True Then
            If ((Gobw(i) <= 30 + BodyLeft(3) And Gobw(i) + G_ObjM.Width / 17 >= 30 + BodyLeft(3)) Or (Gobw(i) <= 30 + BodyLeft(3) + lblLeft(3).Width / 17 And Gobw(i) + G_ObjM.Width / 17 >= 30 + BodyLeft(3) + lblLeft(3).Width / 17) Or (30 + BodyLeft(3) <= Gobw(i) And Gobw(i) + G_ObjM.Width / 17 <= 30 + BodyLeft(3) + lblLeft(3).Width / 17)) _
            And _
            ((Goby(i) + G_ObjM.Height / 17 >= JHei + BodyTop(3) And Goby(i) <= JHei + BodyTop(3)) Or (Goby(i) + G_ObjM.Height / 17 >= JHei + BodyTop(3) + lblLeft(3).Height / 17 And Goby(i) <= JHei + BodyTop(3) + lblLeft(3).Height / 17) Or (Goby(i) >= JHei + BodyTop(3) And Goby(i) + G_ObjM.Height / 17 <= JHei + BodyTop(3) + lblLeft(3).Height / 17)) Then
                Gobw(i) = 700
                GobwVis(i) = False
                GPoint = GPoint + 1
            End If
        End If

        If BobwVis(i) = True Then
            If ((Bobw(i) <= 30 + BodyLeft(3) And Bobw(i) + (B_ObjM.Width / 17) >= 30 + BodyLeft(3)) Or (Bobw(i) <= 30 + BodyLeft(3) + lblLeft(3).Width / 17 And Bobw(i) + (B_ObjM.Width / 17) >= 30 + BodyLeft(3) + lblLeft(3).Width / 17) Or (30 + BodyLeft(3) <= Bobw(i) And Bobw(i) + (B_ObjM.Width / 17) <= 30 + BodyLeft(3) + lblLeft(3).Width / 17)) _
            And _
            ((Boby(i) + B_ObjM.Height / 17 >= JHei + BodyTop(3) And Boby(i) <= JHei + BodyTop(3)) Or (Boby(i) + B_ObjM.Height / 17 >= JHei + BodyTop(3) + lblLeft(3).Height / 17 And Boby(i) <= JHei + BodyTop(3) + lblLeft(3).Height / 17) Or (Boby(i) >= JHei + BodyTop(3) And Boby(i) + B_ObjM.Height / 17 <= JHei + BodyTop(3) + lblLeft(3).Height / 17)) Then
                Timer1.Enabled = False
                Timer2.Enabled = False
                Timer3.Enabled = False
                Frm_Ques.Show
                For p = 1 To 13
                    If BobwVis(p) = True Then
                        Bobw(p) = 700
                        BobwVis(p) = False
                    End If
                Next
                BitBlt Pic_Scr.hdc, 30, JHei, PicMCha(ChaNum).Width, PicMCha(ChaNum).Height, PicMCha(ChaNum).hdc, 0, 0, SRCPAINT 'Character
                BitBlt Pic_Scr.hdc, 30, JHei, PicCha(ChaNum).Width, PicCha(ChaNum).Height, PicCha(ChaNum).hdc, 0, 0, SRCAND 'Character
                Exit Sub
            End If
        End If
    Next
End If
'������ / ����Ʈ ��
lblLife = "Life : " & LifPoint
lblLPoint = "���� ����Ʈ : " & GPoint
lblPoint = "���� : " & SPoint
lblErase = "���찳 : " & item(1)

'���Ĺڽ� �ʱ�ȭ
Pic_Scr.Cls

' BitBlt �����
For i = 1 To 10
    If GobwVis(i) = True Then
        BitBlt Pic_Scr.hdc, Gobw(i), Goby(i), G_ObjM.Width, G_ObjM.Height, G_ObjM.hdc, 0, 0, SRCPAINT 'GObj
        BitBlt Pic_Scr.hdc, Gobw(i), Goby(i), G_Obj.Width, G_Obj.Height, G_Obj.hdc, 0, 0, SRCAND 'GObj
    End If
    If BobwVis(i) = True Then
        BitBlt Pic_Scr.hdc, Bobw(i), Boby(i), B_ObjM.Width, B_ObjM.Height, B_ObjM.hdc, 0, 0, SRCPAINT 'BObj
        BitBlt Pic_Scr.hdc, Bobw(i), Boby(i), B_Obj.Width, B_Obj.Height, B_Obj.hdc, 0, 0, SRCAND 'BObj
    End If
Next

BitBlt Pic_Scr.hdc, 30, JHei, PicMCha(ChaNum).Width, PicMCha(ChaNum).Height, PicMCha(ChaNum).hdc, 0, 0, SRCPAINT 'Character
BitBlt Pic_Scr.hdc, 30, JHei, PicCha(ChaNum).Width, PicCha(ChaNum).Height, PicCha(ChaNum).hdc, 0, 0, SRCAND 'Character
'BitBlt ����� ��

End Sub

Private Sub Pic_Scr_KeyPress(KeyAscii As Integer)
' Ű���� ��
Dim Quit As String
If JB = False And GetAsyncKeyState(32) <> 0 Then
    JB = True
    JSum = False
    JB2 = False
    JHei = 300
    Lim2 = 6
    k = 20
    DoEvents
ElseIf JB = True And JB2 = False And GetAsyncKeyState(32) <> 0 And GetAsyncKeyState(17) <> 0 Then
    If JSum = False Then
        k = k + 11
    Else
        JSum = False
        k = 21
    End If
    JB2 = True
ElseIf JB2 = True And JB3 = False And GetAsyncKeyState(13) <> 0 And GetAsyncKeyState(16) <> 0 Then
    If JSum = False Then
        k = k + 6
    Else
        JSum = False
        k = 15
    End If
    JB3 = True
ElseIf KeyAscii = vbKeyEscape Then
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    If MsgBox("�����Ͻðڽ��ϱ�?", vbYesNo, "Quit") = vbYes Then
        End
    Else
        Timer1.Enabled = True
        Timer2.Enabled = True
        Timer3.Enabled = True
    End If
ElseIf GetAsyncKeyState(84) <> 0 Then
    h.Playing = Not h.Playing
ElseIf GetAsyncKeyState(89) <> 0 Then
    If item(1) >= 1 Then
        it = True
        TimItem.Enabled = True
        ImgItem.Visible = True
        EraseOn = True
        item(1) = item(1) - 1
        For p = 1 To 13
            If BobwVis(p) = True Then
                Bobw(p) = 700
                BobwVis(p) = False
            End If
        Next
    End If
End If
'Ű���� �� ��
End Sub

Private Sub Timer2_Timer()
' OBJ ����
Dim Ran1 As Long, Ran2 As Long
Randomize Ran1

'GOBJ ����
For p = 1 To Stage + 1
    If GobwVis(p) = False Then
        For Ran2 = 1 To GetTickCount / 1000
            Ran1 = Int((100 * Rnd) + 1)
        Next
        If Ran1 >= 91 Then
            Gobw(p) = 700
            GobwVis(p) = True
            If Ran1 <= 95 Then
                Goby(p) = 200
            Else
                Goby(p) = 135
            End If
            Exit For
        End If
    End If
Next
'&&&

'BOBJ ����
For p = 1 To Stage + 1
    If EraseOn = False Then
        If BobwVis(p) = False Then
            For Ran2 = 1 To GetTickCount / 1000
                Ran1 = Int((100 * Rnd) + 1)
            Next
            If Ran1 >= 41 Then
                Bobw(p) = 700
                BobwVis(p) = True
                If Ran1 >= 40 + (2 * Stage) Then
                    If Ran1 <= (70 + Stage) Then
                        Boby(p) = 350
                    Else
                        Boby(p) = 275
                    End If
                Else
                    Boby(p) = 100
                End If
                Exit For
            End If
        End If
    End If
Next
'&&&

'OBJ ���� ��
End Sub

Private Sub Timer3_Timer()
'�б��� �޷����ȵ�� ~
PlayT = PlayT + 1
If Stage <= 12 Then
    Pic_Status.Cls
    BitBlt Pic_Status.hdc, 597, 0, Pic_SchoolM.Width, Pic_SchoolM.Height, Pic_SchoolM.hdc, 0, 0, SRCPAINT
    BitBlt Pic_Status.hdc, 597, 0, Pic_School.Width, Pic_School.Height, Pic_School.hdc, 0, 0, SRCAND
    BitBlt Pic_Status.hdc, PlayT * (Pic_Status.Width / 3000), 7, Pic_SmiM1.Width, Pic_SmiM1.Height, Pic_SmiM1.hdc, 0, 0, SRCPAINT
    BitBlt Pic_Status.hdc, PlayT * (Pic_Status.Width / 3000), 7, Pic_Smi1.Width, Pic_Smi1.Height, Pic_Smi1.hdc, 0, 0, SRCAND
    If PlayT >= 185 Then
        Pic_Status.Cls
        BitBlt Pic_Status.hdc, 597, 0, Pic_SchoolM.Width, Pic_SchoolM.Height, Pic_SchoolM.hdc, 0, 0, SRCPAINT
        BitBlt Pic_Status.hdc, 597, 0, Pic_School.Width, Pic_School.Height, Pic_School.hdc, 0, 0, SRCAND
        BitBlt Pic_Status.hdc, PlayT * (Pic_Status.Width / 3000), 7, Pic_SmiM2.Width, Pic_SmiM2.Height, Pic_SmiM2.hdc, 0, 0, SRCPAINT
        BitBlt Pic_Status.hdc, PlayT * (Pic_Status.Width / 3000), 7, Pic_Smi2.Width, Pic_Smi2.Height, Pic_Smi2.hdc, 0, 0, SRCAND
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        MsgBox "Stage" & Stage & " Clear!"
        Stg_Clear(Stage) = True
        For i = 1 To 10
            Gobw(i) = 700
            Bobw(i) = 700
            Goby(i) = 200
            Boby(i) = 275
        Next
        If Stage = 12 Then
            Contri(1) = True
        End If
        Stage = Stage + 1
        Unload Me
        Frm_Save.Show
    End If
End If
End Sub

Private Sub TimItem_Timer()
'���찳 ���
If it = True Then
    it = False
Else
    EraseOn = False
    ImgItem.Visible = False
End If
End Sub
