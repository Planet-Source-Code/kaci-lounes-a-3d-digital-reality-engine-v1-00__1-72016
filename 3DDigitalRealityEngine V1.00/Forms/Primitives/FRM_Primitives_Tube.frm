VERSION 5.00
Begin VB.Form FRM_Primitives_Tube 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Tube"
   ClientHeight    =   2760
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
   Icon            =   "FRM_Primitives_Tube.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   184
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "&Create"
         Default         =   -1  'True
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tube parameters : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox Check1 
         Caption         =   "Double-sided"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "Creation axe : "
         Height          =   855
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   2175
         Begin VB.OptionButton Option1 
            Caption         =   "YZ"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   495
         End
         Begin VB.OptionButton Option2 
            Caption         =   "XZ"
            Height          =   255
            Left            =   840
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Option3 
            Caption         =   "XY"
            Height          =   255
            Left            =   1440
            TabIndex        =   6
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Text            =   "50"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Text            =   "25"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "300"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Text            =   "10"
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Radius 1 : "
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Radius 2 : "
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Base : "
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Steps : "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
   End
End
Attribute VB_Name = "FRM_Primitives_Tube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()

 Hide
 MousePointer = 11
 DoEvents

 Primitive_Tube False, CSng(Text1.Text), CSng(Text2.Text), CSng(Text3.Text), CSng(Text4.Text), CheckAxe(Option1, Option2, Option3), CheckOut(Check1)

 MousePointer = 0
 Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub
