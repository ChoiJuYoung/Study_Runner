VERSION 5.00
Begin VB.Form FrmContri 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Contribution (°øÀû)"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12480
   Icon            =   "FrmContri.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   12480
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¹ÌÈ¹µæ"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   10200
      TabIndex        =   15
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¹ÌÈ¹µæ"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6840
      TabIndex        =   14
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¹ÌÈ¹µæ"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   13
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¹ÌÈ¹µæ"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¹ÌÈ¹µæ"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   10200
      TabIndex        =   11
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¹ÌÈ¹µæ"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6840
      TabIndex        =   10
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¹ÌÈ¹µæ"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   9
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¹ÌÈ¹µæ"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¹®°ú»ý"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   10200
      TabIndex        =   7
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Non Sense"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ÀÌ°ú»ý"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¹ÙÄû¹ú·¹"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Conqueror"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   10200
      TabIndex        =   3
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ãê"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Master"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "±Ù¼ºÀÇ ÇÑ±¹ÀÎ"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Image ImgCon 
      Height          =   3465
      Index           =   7
      Left            =   10200
      Top             =   4320
      Width           =   2190
   End
   Begin VB.Image ImgCon 
      Height          =   3465
      Index           =   6
      Left            =   6840
      Top             =   4320
      Width           =   2190
   End
   Begin VB.Image ImgCon 
      Height          =   3465
      Index           =   5
      Left            =   3480
      Top             =   4320
      Width           =   2190
   End
   Begin VB.Image ImgCon 
      Height          =   3465
      Index           =   4
      Left            =   120
      Top             =   4320
      Width           =   2190
   End
   Begin VB.Image ImgCon 
      Height          =   3465
      Index           =   3
      Left            =   10200
      Top             =   120
      Width           =   2190
   End
   Begin VB.Image ImgCon 
      Height          =   3465
      Index           =   2
      Left            =   6840
      Top             =   120
      Width           =   2190
   End
   Begin VB.Image ImgCon 
      Height          =   3465
      Index           =   1
      Left            =   3480
      Top             =   120
      Width           =   2190
   End
   Begin VB.Image ImgCon 
      Height          =   3465
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   2190
   End
End
Attribute VB_Name = "FrmContri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
For i = 0 To 7
    If Contri(i + 1) = True Then
        ImgCon(i) = LoadResPicture(112, vbResBitmap)
        Label2(i) = "È¹µæ"
        Label2(i).ForeColor = RGB(255, 0, 255)
    End If
Next
End Sub
