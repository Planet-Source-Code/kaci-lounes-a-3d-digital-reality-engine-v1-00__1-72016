VERSION 5.00
Begin VB.Form FRM_Primitives_Landscape 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Primatives : Add Landscape"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_Primitives_Landscape.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "&Create"
         Default         =   -1  'True
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Landscape (grid) parameters : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.Frame Frame4 
         Caption         =   "Texels filter : "
         Height          =   650
         Left            =   2160
         TabIndex        =   18
         Top             =   850
         Width           =   2175
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "FRM_Primitives_Landscape.frx":000C
            Left            =   120
            List            =   "FRM_Primitives_Landscape.frx":0022
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Text            =   "50"
         Top             =   1080
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Displacement map : "
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   4215
         Begin VB.CommandButton Command1 
            Caption         =   "&Browse for a displacement map..."
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Double-sided"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "Creation axe : "
         Height          =   615
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   2175
         Begin VB.OptionButton Option1 
            Caption         =   "YZ"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton Option2 
            Caption         =   "XZ"
            Height          =   255
            Left            =   840
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Option3 
            Caption         =   "XY"
            Height          =   255
            Left            =   1440
            TabIndex        =   5
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Text            =   "250"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "30"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Text            =   "250"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Depth : "
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Height : "
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Steps : "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Width : "
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "FRM_Primitives_Landscape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

 BrowseType = 1: FRM_BrowseMap.Show 1

End Sub
Private Sub Command2_Click()

 If (Browsed = False) Then MsgBox "You must select a displacement map !", vbCritical, "No map": Exit Sub

 Hide
 MousePointer = 11
 DoEvents

 Dim TexFilter As K3DE_TEXELS_FILTER_MODES

 Select Case Combo1.Text
  Case "Bilinear":              TexFilter = K3DE_XFM_BILINEAR
  Case "Bell":                  TexFilter = K3DE_XFM_BELL
  Case "Gaussian":              TexFilter = K3DE_XFM_GAUSSIAN
  Case "Cubic spline B":        TexFilter = K3DE_XFM_CUBIC_SPLINE_B
  Case "Cubic spline BC":       TexFilter = K3DE_XFM_CUBIC_SPLINE_BC
  Case "Cubic spline cardinal": TexFilter = K3DE_XFM_CUBIC_SPLINE_CARDINAL
 End Select

 Primitive_Landscape TheOutputMap, TexFilter, CSng(Text1.Text), CSng(Text2.Text), CSng(Text4.Text), CSng(Text3.Text), CheckAxe(Option1, Option2, Option3), CheckOut(Check1)

 MousePointer = 0
 Unload Me

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub
Private Sub Form_Load()

 Combo1.Text = "Bilinear"

End Sub
