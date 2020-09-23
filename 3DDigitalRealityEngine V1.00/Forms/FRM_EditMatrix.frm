VERSION 5.00
Begin VB.Form FRM_EditMatrix 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Matrix datas"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_EditMatrix.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   6975
      Begin VB.CommandButton Command1 
         Caption         =   "&Update"
         Default         =   -1  'True
         Height          =   375
         Left            =   5640
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Matrix datas (4x4) : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5640
         TabIndex        =   18
         Text            =   "0"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3960
         TabIndex        =   17
         Text            =   "0"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Text            =   "0"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   15
         Text            =   "0"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5640
         TabIndex        =   14
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3960
         TabIndex        =   13
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   12
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   11
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5640
         TabIndex        =   10
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   7
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5640
         TabIndex        =   6
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3960
         TabIndex        =   5
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "M44"
         Height          =   195
         Left            =   5280
         TabIndex        =   34
         Top             =   1560
         Width           =   300
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "M43"
         Height          =   195
         Left            =   3600
         TabIndex        =   33
         Top             =   1560
         Width           =   300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "M42"
         Height          =   195
         Left            =   1920
         TabIndex        =   32
         Top             =   1560
         Width           =   300
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "M41"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   1560
         Width           =   300
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "M34"
         Height          =   195
         Left            =   5280
         TabIndex        =   30
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "M33"
         Height          =   195
         Left            =   3600
         TabIndex        =   29
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "M32"
         Height          =   195
         Left            =   1920
         TabIndex        =   28
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "M31"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "M24"
         Height          =   195
         Left            =   5280
         TabIndex        =   26
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "M23"
         Height          =   195
         Left            =   3600
         TabIndex        =   25
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "M22"
         Height          =   195
         Left            =   1920
         TabIndex        =   24
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "M21"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "M14"
         Height          =   195
         Left            =   5280
         TabIndex        =   22
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "M13"
         Height          =   195
         Left            =   3600
         TabIndex        =   21
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "M12"
         Height          =   195
         Left            =   1920
         TabIndex        =   20
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "M11"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   300
      End
   End
End
Attribute VB_Name = "FRM_EditMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

 If (MatWindowMode = False) Then 'Object matrix (World matrix)
  TheMeshs(MatWindowIndex).WorldMatrix.M11 = Text1.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M12 = Text2.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M13 = Text3.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M14 = Text4.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M21 = Text5.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M22 = Text6.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M23 = Text7.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M24 = Text8.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M31 = Text9.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M32 = Text10.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M33 = Text11.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M34 = Text12.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M41 = Text13.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M42 = Text14.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M43 = Text15.Text
  TheMeshs(MatWindowIndex).WorldMatrix.M44 = Text16.Text
 Else 'Camera-space matrix (View matrix)
  TheCameras(MatWindowIndex).ViewMatrix.M11 = Text1.Text
  TheCameras(MatWindowIndex).ViewMatrix.M12 = Text2.Text
  TheCameras(MatWindowIndex).ViewMatrix.M13 = Text3.Text
  TheCameras(MatWindowIndex).ViewMatrix.M14 = Text4.Text
  TheCameras(MatWindowIndex).ViewMatrix.M21 = Text5.Text
  TheCameras(MatWindowIndex).ViewMatrix.M22 = Text6.Text
  TheCameras(MatWindowIndex).ViewMatrix.M23 = Text7.Text
  TheCameras(MatWindowIndex).ViewMatrix.M24 = Text8.Text
  TheCameras(MatWindowIndex).ViewMatrix.M31 = Text9.Text
  TheCameras(MatWindowIndex).ViewMatrix.M32 = Text10.Text
  TheCameras(MatWindowIndex).ViewMatrix.M33 = Text11.Text
  TheCameras(MatWindowIndex).ViewMatrix.M34 = Text12.Text
  TheCameras(MatWindowIndex).ViewMatrix.M41 = Text13.Text
  TheCameras(MatWindowIndex).ViewMatrix.M42 = Text14.Text
  TheCameras(MatWindowIndex).ViewMatrix.M43 = Text15.Text
  TheCameras(MatWindowIndex).ViewMatrix.M44 = Text16.Text
 End If

 Unload Me

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub
Private Sub Form_Load()

 If (MatWindowMode = False) Then 'Object matrix (World matrix)
  Caption = "Edit World matrix of the mesh : " & TheMeshs(MatWindowIndex).Label & "  (Index " & MatWindowIndex & ")"
  Text1.Text = TheMeshs(MatWindowIndex).WorldMatrix.M11
  Text2.Text = TheMeshs(MatWindowIndex).WorldMatrix.M12
  Text3.Text = TheMeshs(MatWindowIndex).WorldMatrix.M13
  Text4.Text = TheMeshs(MatWindowIndex).WorldMatrix.M14
  Text5.Text = TheMeshs(MatWindowIndex).WorldMatrix.M21
  Text6.Text = TheMeshs(MatWindowIndex).WorldMatrix.M22
  Text7.Text = TheMeshs(MatWindowIndex).WorldMatrix.M23
  Text8.Text = TheMeshs(MatWindowIndex).WorldMatrix.M24
  Text9.Text = TheMeshs(MatWindowIndex).WorldMatrix.M31
  Text10.Text = TheMeshs(MatWindowIndex).WorldMatrix.M32
  Text11.Text = TheMeshs(MatWindowIndex).WorldMatrix.M33
  Text12.Text = TheMeshs(MatWindowIndex).WorldMatrix.M34
  Text13.Text = TheMeshs(MatWindowIndex).WorldMatrix.M41
  Text14.Text = TheMeshs(MatWindowIndex).WorldMatrix.M42
  Text15.Text = TheMeshs(MatWindowIndex).WorldMatrix.M43
  Text16.Text = TheMeshs(MatWindowIndex).WorldMatrix.M44
 Else 'Camera-space matrix (View matrix)
  Caption = "Edit View matrix of the camera : " & TheCameras(MatWindowIndex).Label & "  (Index " & MatWindowIndex & ")"
  Text1.Text = TheCameras(MatWindowIndex).ViewMatrix.M11
  Text2.Text = TheCameras(MatWindowIndex).ViewMatrix.M12
  Text3.Text = TheCameras(MatWindowIndex).ViewMatrix.M13
  Text4.Text = TheCameras(MatWindowIndex).ViewMatrix.M14
  Text5.Text = TheCameras(MatWindowIndex).ViewMatrix.M21
  Text6.Text = TheCameras(MatWindowIndex).ViewMatrix.M22
  Text7.Text = TheCameras(MatWindowIndex).ViewMatrix.M23
  Text8.Text = TheCameras(MatWindowIndex).ViewMatrix.M24
  Text9.Text = TheCameras(MatWindowIndex).ViewMatrix.M31
  Text10.Text = TheCameras(MatWindowIndex).ViewMatrix.M32
  Text11.Text = TheCameras(MatWindowIndex).ViewMatrix.M33
  Text12.Text = TheCameras(MatWindowIndex).ViewMatrix.M34
  Text13.Text = TheCameras(MatWindowIndex).ViewMatrix.M41
  Text14.Text = TheCameras(MatWindowIndex).ViewMatrix.M42
  Text15.Text = TheCameras(MatWindowIndex).ViewMatrix.M43
  Text16.Text = TheCameras(MatWindowIndex).ViewMatrix.M44
 End If

End Sub
