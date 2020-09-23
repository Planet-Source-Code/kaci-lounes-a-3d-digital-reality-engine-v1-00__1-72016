VERSION 5.00
Begin VB.Form FRM_Primitives_Box 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Box"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_Primitives_Box.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   159
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   3615
      Begin VB.CommandButton Command2 
         Caption         =   "&Create"
         Default         =   -1  'True
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Box dimensions :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CheckBox Check4 
         Caption         =   "Double-sided"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "100"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Text            =   "100"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "100"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Depth : "
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Height : "
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Width : "
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "FRM_Primitives_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()

 Hide
 MousePointer = 11
 DoEvents
 
 Primitive_Box False, VectorInput(CSng(Text1.Text), CSng(Text2.Text), CSng(Text3.Text)), CheckOut(Check4)

 MousePointer = 0
 Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub
