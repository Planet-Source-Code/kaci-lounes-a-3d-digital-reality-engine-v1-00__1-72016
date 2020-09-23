VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRM_Lights_Omni 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Omni light source"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_Lights_Omni.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Omni light parameters : "
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
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Text            =   "OmniLight "
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "500"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   5
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enable"
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   1080
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label : "
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Position : "
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Color"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         MousePointer    =   2  'Cross
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Z"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Y"
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Range : "
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   4335
      Begin MSComDlg.CommonDialog COMDLG 
         Left            =   120
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Create"
         Default         =   -1  'True
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Color"
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "FRM_Lights_Omni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()

 Hide
 MousePointer = 11

 If (OmniWindowMode = True) Then
  With TheSphereLights(OmniWindowIndex)
   .Label = Text5.Text
   .Color = ColorLongToRGB(Label1.BackColor)
   .Position = VectorInput(CSng(Text1.Text), CSng(Text2.Text), CSng(Text3.Text))
   .Range = CSng(Text4.Text)
   .Enable = CheckOut(Check1)
  End With
  MousePointer = 0
  Unload Me
  Exit Sub
 End If

 Dim Indx&: Indx = SphereLight3D_Add

 If (Indx = -1) Then
  MsgBox "The allocation is full for the Omni lights !", vbCritical, "Error"
 Else
  With TheSphereLights(Indx)
   .Label = Text5.Text
   .Color = ColorLongToRGB(Label1.BackColor)
   .Position = VectorInput(CSng(Text1.Text), CSng(Text2.Text), CSng(Text3.Text))
   .Range = CSng(Text4.Text)
   .Enable = CheckOut(Check1)
  End With
 End If

 MousePointer = 0
 Unload Me

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub
Private Sub Form_Load()

 If (OmniWindowMode = True) Then
  With TheSphereLights(OmniWindowIndex)
   Caption = "Properties of : '" & .Label & "'"
   Text5.Text = .Label
   Text1.Text = .Position.X
   Text2.Text = .Position.Y
   Text3.Text = .Position.Z
   Text4.Text = .Range
   If (.Enable = True) Then Check1.Value = vbChecked Else Check1.Value = vbUnchecked
   Label1.BackColor = ColorRGBToLong(.Color)
   Command2.Caption = "&Update"
  End With
 Else
  Label1.BackColor = vbWhite
  Text5.Text = Text5.Text & (TheSphereLightsCount + 1)
 End If

End Sub
Private Sub Label1_Click()

 COMDLG.ShowColor
 Label1.BackColor = COMDLG.Color

End Sub
