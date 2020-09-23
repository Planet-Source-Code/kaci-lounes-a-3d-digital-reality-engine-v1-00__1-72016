VERSION 5.00
Begin VB.Form FRM_Primitives_Pyramid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Pyramid"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_Primitives_Pyramid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   152
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "&Create"
         Default         =   -1  'True
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pyramid parameters : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox Check1 
         Caption         =   "Double-sided"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "Creation axe : "
         Height          =   615
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   2175
         Begin VB.OptionButton Option1 
            Caption         =   "YZ"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton Option2 
            Caption         =   "XZ"
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Option3 
            Caption         =   "XY"
            Height          =   255
            Left            =   1440
            TabIndex        =   4
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "150"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Text            =   "150"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Radius : "
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Base : "
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "FRM_Primitives_Pyramid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()

 Hide
 MousePointer = 11
 DoEvents

 Primitive_Pyramid False, CSng(Text1.Text), CSng(Text2.Text), CheckAxe(Option1, Option2, Option3), CheckOut(Check1)

 MousePointer = 0
 Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub
