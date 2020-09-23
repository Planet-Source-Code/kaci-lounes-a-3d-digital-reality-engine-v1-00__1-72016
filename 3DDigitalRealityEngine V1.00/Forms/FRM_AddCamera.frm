VERSION 5.00
Begin VB.Form FRM_AddCamera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Camera"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_AddCamera.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Camera parameters : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   ".."
         Enabled         =   0   'False
         Height          =   230
         Left            =   3720
         TabIndex        =   30
         Top             =   1820
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   28
         Text            =   "1"
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "User.def cam-matrix"
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   1800
         Width           =   1875
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Text            =   "Camera "
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Text            =   "300"
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Text            =   "-100"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Text            =   "100"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Text            =   "-100"
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Back-face culling"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1515
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   6
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Text            =   "45"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   3
         Text            =   "0"
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dispersion :"
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Label : "
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Position : "
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Z"
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Y"
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Focal distance : "
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Direction : "
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Z"
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Y"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "FOV.Ang :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Roll (Z).Ang : "
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   4215
      Begin VB.CommandButton Command2 
         Caption         =   "&Create"
         Default         =   -1  'True
         Height          =   375
         Left            =   3120
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
      TabIndex        =   26
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "FRM_AddCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Informed As Boolean
Private Sub Check2_Click()

 If (CamWindowMode = True) Then
  Command1.Enabled = True
 Else
  If (Check2.Value = vbChecked) Then
   If (Informed = False) Then
    MsgBox "You have to create the camera first, before modifing the camera-matrix in the scene manager.", vbInformation, "User-defined matrix"
    Informed = True
   End If
  End If
 End If

End Sub
Private Sub Command1_Click()

 MatWindowMode = True
 MatWindowIndex = CamWindowIndex
 FRM_EditMatrix.Show 1

End Sub
Private Sub Command2_Click()

 Hide
 MousePointer = 11

 If (CamWindowMode = True) Then
  With TheCameras(CamWindowIndex)
   .Label = Text5.Text
   .Position = VectorInput(CSng(Text1.Text), CSng(Text2.Text), CSng(Text3.Text))
   .Direction = VectorInput(CSng(Text6.Text), CSng(Text8.Text), CSng(Text7.Text))
   .FOVAngle = DegToRad(CSng(Text9.Text))
   .RollAngle = DegToRad(CSng(Text10.Text))
   .BackFaceCulling = CheckOut(Check1)
   .MakeMatrix = Not CheckOut(Check2)
   .ClearDistance = CSng(Text4.Text)
   .Dispersion = CSng(Text11.Text)
  End With
  MousePointer = 0
  Unload Me
  Exit Sub
 End If

 Dim Indx&: Indx = Camera3D_Add

 If (Indx = -1) Then
  MsgBox "The allocation is full for the Cameras !", vbCritical, "Error"
 Else
  With TheCameras(Indx)
   .Label = Text5.Text
   .Position = VectorInput(CSng(Text1.Text), CSng(Text2.Text), CSng(Text3.Text))
   .Direction = VectorInput(CSng(Text6.Text), CSng(Text8.Text), CSng(Text7.Text))
   .FOVAngle = DegToRad(CSng(Text9.Text))
   .RollAngle = DegToRad(CSng(Text10.Text))
   .BackFaceCulling = CheckOut(Check1)
   .MakeMatrix = Not CheckOut(Check2)
   .ClearDistance = CSng(Text4.Text)
   .Dispersion = CSng(Text11.Text)
  End With
 End If

 MousePointer = 0
 Unload Me

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub
Private Sub Form_Load()

 If (CamWindowMode = True) Then
  With TheCameras(CamWindowIndex)
   Caption = "Properties of : '" & .Label & "'"
   Text5.Text = .Label
   Text1.Text = .Position.X
   Text2.Text = .Position.Y
   Text3.Text = .Position.Z
   Text6.Text = .Direction.X
   Text8.Text = .Direction.Y
   Text7.Text = .Direction.Z
   Text9.Text = RadToDeg(.FOVAngle)
   Text10.Text = RadToDeg(.RollAngle)
   Text4.Text = .ClearDistance
   Text11.Text = .Dispersion
   If (.BackFaceCulling = True) Then Check1.Value = vbChecked Else Check1.Value = vbUnchecked
   If (.MakeMatrix = False) Then Check2.Value = vbChecked Else Check2.Value = vbUnchecked
   Command2.Caption = "&Update"
   Command1.Visible = True
  End With
 Else
  Text5.Text = Text5.Text & (TheCamerasCount + 1)
 End If

End Sub
