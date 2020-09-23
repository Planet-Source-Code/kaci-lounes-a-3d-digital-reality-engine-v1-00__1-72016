VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRM_Lights_Spot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Spot light source"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_Lights_Spot.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   4215
      Begin VB.CommandButton Command2 
         Caption         =   "&Create"
         Default         =   -1  'True
         Height          =   375
         Left            =   3120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin MSComDlg.CommonDialog COMDLG 
         Left            =   120
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spot light parameters : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   27
         Text            =   "40"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   25
         Text            =   "45"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   20
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   19
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enable"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   1800
         Value           =   1  'Checked
         Width           =   800
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Text            =   "-100"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Text            =   "100"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Text            =   "100"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "500"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Text            =   "SpotLight "
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Penumbra Ang : "
         Height          =   255
         Left            =   1920
         TabIndex        =   28
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Cone.Ang :"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Y"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Z"
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Direction : "
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Range : "
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Y"
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   720
         Width           =   135
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         MousePointer    =   2  'Cross
         TabIndex        =   10
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Color"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Position : "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label : "
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Color"
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "FRM_Lights_Spot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()

 Hide
 MousePointer = 11

 If (SpotWindowMode = True) Then
  With TheConeLights(SpotWindowIndex)
   .Label = Text5.Text
   .Position = VectorInput(CSng(Text1.Text), CSng(Text2.Text), CSng(Text3.Text))
   .Direction = VectorInput(CSng(Text6.Text), CSng(Text8.Text), CSng(Text7.Text))
   .Falloff = DegToRad(CSng(Text9.Text))
   .Hotspot = DegToRad(CSng(Text10.Text))
   .Range = CSng(Text4.Text)
   .Color = ColorLongToRGB(Label1.BackColor)
   .Enable = CheckOut(Check1)
  End With
  MousePointer = 0
  Unload Me
  Exit Sub
 End If

 Dim Indx&: Indx = ConeLight3D_Add

 If (Indx = -1) Then
  MsgBox "The allocation is full for the Spot lights !", vbCritical, "Error"
 Else
  With TheConeLights(Indx)
   .Label = Text5.Text
   .Position = VectorInput(CSng(Text1.Text), CSng(Text2.Text), CSng(Text3.Text))
   .Direction = VectorInput(CSng(Text6.Text), CSng(Text8.Text), CSng(Text7.Text))
   .Falloff = DegToRad(CSng(Text9.Text))
   .Hotspot = DegToRad(CSng(Text10.Text))
   .Range = CSng(Text4.Text)
   .Color = ColorLongToRGB(Label1.BackColor)
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

 If (SpotWindowMode = True) Then
  With TheConeLights(SpotWindowIndex)
   Caption = "Properties of : " & .Label & "  (Index " & SpotWindowIndex & ")"
   Text5.Text = .Label
   Text1.Text = .Position.X
   Text2.Text = .Position.Y
   Text3.Text = .Position.Z
   Text6.Text = .Direction.X
   Text8.Text = .Direction.Y
   Text7.Text = .Direction.Z
   Text9.Text = RadToDeg(.Falloff)
   Text10.Text = RadToDeg(.Hotspot)
   Text4.Text = .Range
   If (.Enable = True) Then Check1.Value = vbChecked Else Check1.Value = vbUnchecked
   Label1.BackColor = ColorRGBToLong(.Color)
   Command2.Caption = "&Update"
  End With
 Else
  Label1.BackColor = vbWhite
  Text5.Text = Text5.Text & (TheConeLightsCount + 1)
 End If

End Sub
Private Sub Label1_Click()

 COMDLG.ShowColor
 Label1.BackColor = COMDLG.Color

End Sub
