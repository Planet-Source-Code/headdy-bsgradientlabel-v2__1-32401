VERSION 5.00
Object = "{D1D34DFA-7707-4FB4-9AD2-68BBE11FECF5}#1.0#0"; "bsGradientLabel.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "bsGradientLabel Demo"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Select a font"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   2760
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmMain.frx":0442
      Left            =   1320
      List            =   "frmMain.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2280
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   1320
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   69891
   End
   Begin Project1.bsGradientLabel bsGradientLabel2 
      Height          =   255
      Left            =   3960
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      Caption         =   "Added from Version 1"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmMain.frx":0468
      Left            =   1320
      List            =   "frmMain.frx":0484
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
   Begin Project1.bsGradientLabel glTest 
      Height          =   3855
      Left            =   120
      Top             =   480
      Width           =   975
      _ExtentX        =   6800
      _ExtentY        =   1720
      GradientType    =   1
      Caption         =   "Options"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   12648384
      Colour1         =   12583104
      Colour2         =   12582912
      LabelType       =   1
      BorderStyle     =   5
      FlatBorderColour=   16777215
      TextShadow      =   -1  'True
      TextShadowYOffset=   -2
   End
   Begin Project1.bsGradientLabel bsGradientLabel1 
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      Caption         =   "Introducing the bsGradientLabel v2"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   -2147483633
      Colour2         =   13160664
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmMain.frx":04D2
      Top             =   3720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About this control"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmMain.frx":05C5
      Left            =   1320
      List            =   "frmMain.frx":05D2
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin Project1.bsGradientLabel bsGradientLabel3 
      Height          =   255
      Left            =   3960
      Top             =   1680
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      Caption         =   "But let's make this clear..."
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Text Alignment"
      Height          =   195
      Left            =   1320
      TabIndex        =   10
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   $"frmMain.frx":05F3
      Height          =   975
      Left            =   3960
      TabIndex        =   8
      Top             =   1920
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "On request, border styles with customisable colours. Text shadows are also made possible (see left)."
      Height          =   585
      Left            =   3960
      TabIndex        =   7
      Top             =   840
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Border Style"
      Height          =   195
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Gradient Type"
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShownWarning As Boolean

Private Sub Combo1_Click()
   glTest.GradientType = Combo1.ListIndex
End Sub

Private Sub Combo2_Click()
   glTest.BorderStyle = Combo2.ListIndex
End Sub

Private Sub Combo3_Click()
   glTest.CaptionAlignment = Combo3.ListIndex
End Sub

Private Sub Command1_Click()
   MsgBox Text1.Text, vbInformation, "About this control"
   glTest.ShowAbout
End Sub

Private Sub Command2_Click()
   End
End Sub

Private Sub Command3_Click()

   Dim temp As New StdFont
   If Not ShownWarning Then
      MsgBox "Only TrueType fonts can be rotated. See for yourself...", vbInformation, "A friendly warning"
      ShownWarning = True
   End If
   On Error GoTo forget_it
   dlgFont.ShowFont
   With temp
      .Name = dlgFont.FontName
      .Bold = dlgFont.FontBold
      .Italic = dlgFont.FontItalic
      .Underline = dlgFont.FontUnderline
      .Size = dlgFont.FontSize
   End With
   
   Set glTest.Fount = temp
forget_it:
   
End Sub

Private Sub Form_Load()
   Combo1.ListIndex = glTest.GradientType
   Combo2.ListIndex = glTest.BorderStyle
   Combo3.ListIndex = glTest.CaptionAlignment
   With glTest.Fount
      dlgFont.FontName = .Name
      dlgFont.FontBold = .Bold
      dlgFont.FontItalic = .Italic
      dlgFont.FontUnderline = .Underline
      dlgFont.FontSize = .Size
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub
