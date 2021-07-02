VERSION 5.00
Begin VB.Form Frm_Save 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   0  '없음
   Caption         =   "Save Or Not?"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   Icon            =   "Frm_Save.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame FmeSel 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  '없음
      Height          =   1575
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3015
      Begin Study_Runner.jcbutton jcbutton6 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
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
         Caption         =   "Save Or End"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   2
      End
      Begin Study_Runner.jcbutton jcbutton5 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
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
         Caption         =   "Item Shop"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   2
      End
   End
   Begin Study_Runner.jcbutton jcbutton4 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16761087
      Caption         =   "Non Save - END"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Study_Runner.jcbutton jcbutton3 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16761087
      Caption         =   "Save - END"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Study_Runner.jcbutton jcbutton2 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16761087
      Caption         =   "Non Save - GO"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
   Begin Study_Runner.jcbutton jcbutton1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16761087
      Caption         =   "Save - GO"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   3
   End
End
Attribute VB_Name = "Frm_Save"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub jcbutton1_Click()
'저장 / GO
Call save
DoEvents
Frm_Play.Show
Unload Me
End Sub

Private Sub jcbutton2_Click()
'저장없이 go
Frm_Play.Show
Unload Me
End Sub

Private Sub jcbutton3_Click()
'저장하고 끄기
Call save
DoEvents
End
End Sub

Private Sub jcbutton4_Click()
'저장없이 끄기
End
End Sub

Private Sub jcbutton5_Click()
'상점 오픈
FrmShop.Show
Unload Me
End Sub

Private Sub jcbutton6_Click()
'저장부 오픈
FmeSel.Visible = False
End Sub
