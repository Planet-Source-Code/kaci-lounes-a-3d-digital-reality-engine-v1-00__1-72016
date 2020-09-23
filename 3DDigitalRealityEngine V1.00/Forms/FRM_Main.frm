VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRM_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D Digital Reality Engine, V1.00, by KACI Lounes"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1009
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   10680
      TabIndex        =   45
      Top             =   1320
      Width           =   4335
      Begin MSComDlg.CommonDialog COMDLG 
         Left            =   0
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command9 
         Caption         =   "About the author..."
         Height          =   735
         Left            =   2280
         TabIndex        =   53
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Documentation..."
         Height          =   735
         Left            =   240
         TabIndex        =   52
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Render settings..."
         Height          =   735
         Left            =   2280
         TabIndex        =   51
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Render scene..."
         Height          =   735
         Left            =   240
         TabIndex        =   50
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Scene manager..."
         Height          =   735
         Left            =   2280
         TabIndex        =   49
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save scene..."
         Height          =   735
         Left            =   240
         TabIndex        =   48
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Open scene..."
         Height          =   735
         Left            =   2280
         TabIndex        =   47
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reset Engine"
         Height          =   735
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current camera : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   10680
      TabIndex        =   36
      Top             =   5640
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "Select a camera..."
         Height          =   495
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1065
         Width           =   2550
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Back-faces culling : "
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   2880
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Roll angle : "
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   2520
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "FOV angle : "
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Direction : "
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Position : "
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Camera name : "
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   1140
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Preview elements : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10680
      TabIndex        =   29
      Top             =   120
      Width           =   4335
      Begin VB.CheckBox Check1 
         Caption         =   "Geometry"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Omni lights"
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Spot lights"
         Height          =   255
         Left            =   1560
         TabIndex        =   33
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Cameras"
         Height          =   255
         Left            =   2880
         TabIndex        =   32
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Photon map"
         Height          =   255
         Left            =   2880
         TabIndex        =   31
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Normals"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.Frame FRA_Views 
      Caption         =   "Scene views :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton Command15 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   5
         Top             =   0
         Width           =   855
      End
      Begin VB.PictureBox PIC_Front 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         Height          =   4095
         Left            =   120
         ScaleHeight     =   269
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   333
         TabIndex        =   4
         ToolTipText     =   "Front view (XY plane) of the scene"
         Top             =   360
         Width           =   5055
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   4080
            ScaleHeight     =   375
            ScaleMode       =   0  'User
            ScaleWidth      =   900
            TabIndex        =   14
            Top             =   3720
            Width           =   900
            Begin VB.OptionButton Option1 
               BackColor       =   &H00808080&
               Caption         =   "T"
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Value           =   -1  'True
               Width           =   375
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00808080&
               Caption         =   "S"
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   480
               TabIndex        =   15
               Top             =   0
               Width           =   375
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   615
            TabIndex        =   6
            Top             =   0
            Width           =   615
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               Caption         =   "Front"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   12
               Top             =   0
               Width           =   615
            End
         End
      End
      Begin VB.PictureBox PIC_Top 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         Height          =   4095
         Left            =   5280
         ScaleHeight     =   269
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   333
         TabIndex        =   3
         ToolTipText     =   "Top view (XZ plane) of the scene"
         Top             =   360
         Width           =   5055
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   4080
            ScaleHeight     =   375
            ScaleMode       =   0  'User
            ScaleWidth      =   900
            TabIndex        =   17
            Top             =   3720
            Width           =   900
            Begin VB.OptionButton Option4 
               BackColor       =   &H00808080&
               Caption         =   "T"
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   0
               TabIndex        =   19
               Top             =   0
               Value           =   -1  'True
               Width           =   375
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00808080&
               Caption         =   "S"
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   480
               TabIndex        =   18
               Top             =   0
               Width           =   375
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   495
            TabIndex        =   7
            Top             =   0
            Width           =   495
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               Caption         =   "Top"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Width           =   495
            End
         End
      End
      Begin VB.PictureBox PIC_Side 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         Height          =   4095
         Left            =   120
         ScaleHeight     =   269
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   333
         TabIndex        =   2
         ToolTipText     =   "Left view (YZ plane) of the scene"
         Top             =   4560
         Width           =   5055
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   4080
            ScaleHeight     =   375
            ScaleMode       =   0  'User
            ScaleWidth      =   900
            TabIndex        =   20
            Top             =   3720
            Width           =   900
            Begin VB.OptionButton Option6 
               BackColor       =   &H00808080&
               Caption         =   "T"
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   0
               TabIndex        =   22
               Top             =   0
               Value           =   -1  'True
               Width           =   375
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00808080&
               Caption         =   "S"
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   480
               TabIndex        =   21
               Top             =   0
               Width           =   375
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   495
            TabIndex        =   8
            Top             =   0
            Width           =   495
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               Caption         =   "Side"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   11
               Top             =   0
               Width           =   495
            End
         End
      End
      Begin VB.PictureBox PIC_Persp 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         Height          =   4095
         Left            =   5280
         ScaleHeight     =   269
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   333
         TabIndex        =   1
         ToolTipText     =   "Free Perspective view of the scene"
         Top             =   4560
         Width           =   5055
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2400
            ScaleHeight     =   375
            ScaleMode       =   0  'User
            ScaleWidth      =   2580
            TabIndex        =   23
            Top             =   3720
            Width           =   2580
            Begin VB.OptionButton Option11 
               BackColor       =   &H00808080&
               Caption         =   "Rz"
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   2040
               TabIndex        =   28
               Top             =   0
               Width           =   495
            End
            Begin VB.OptionButton Option10 
               BackColor       =   &H00808080&
               Caption         =   "R"
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   1560
               TabIndex        =   27
               Top             =   0
               Width           =   375
            End
            Begin VB.OptionButton Option9 
               BackColor       =   &H00808080&
               Caption         =   "T"
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   1080
               TabIndex        =   26
               Top             =   0
               Width           =   375
            End
            Begin VB.OptionButton Option8 
               BackColor       =   &H00808080&
               Caption         =   "FOV"
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   480
               TabIndex        =   25
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton Option7 
               BackColor       =   &H00808080&
               Caption         =   "Z"
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   0
               TabIndex        =   24
               Top             =   0
               Value           =   -1  'True
               Width           =   375
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1095
            TabIndex        =   9
            Top             =   0
            Width           =   1095
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               Caption         =   "Perspective"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   10
               Top             =   0
               Width           =   1095
            End
         End
      End
   End
   Begin VB.Menu MNU_File 
      Caption         =   "&File"
      Begin VB.Menu MNU_File_Reset 
         Caption         =   "Rese&t"
      End
      Begin VB.Menu MNU_File_Open 
         Caption         =   "&Open scene..."
      End
      Begin VB.Menu MNU_File_Save 
         Caption         =   "&Save scene..."
      End
      Begin VB.Menu Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu MNU_File_Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu MNU_Scene 
      Caption         =   "&Scene"
      Begin VB.Menu MNU_Scene_Objects 
         Caption         =   "Add &Object"
         Begin VB.Menu MNU_Scene_Objects_Prim 
            Caption         =   "&Primitives"
            Begin VB.Menu MNU_Scene_Objects_Prim_Box 
               Caption         =   "Box"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Capsule 
               Caption         =   "Capsule"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Cone 
               Caption         =   "Cone"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Cylinder 
               Caption         =   "Cylinder"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Disk 
               Caption         =   "Disk"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Grid 
               Caption         =   "Grid"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Hemisphere 
               Caption         =   "Hemisphere"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Octahedron 
               Caption         =   "Octaherdron"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Pyramid 
               Caption         =   "Pyramid"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Sphere 
               Caption         =   "Sphere"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Tetrahedron 
               Caption         =   "Tetrahedron"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Torus 
               Caption         =   "Torus"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Tube 
               Caption         =   "Tube"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_FlatTube 
               Caption         =   "Flat tube"
            End
            Begin VB.Menu Bar2 
               Caption         =   "-"
            End
            Begin VB.Menu MNU_Scene_Objects_Prim_Spe 
               Caption         =   "Special"
               Begin VB.Menu MNU_Scene_Objects_Prim_Spec_Landscape 
                  Caption         =   "Landscape"
               End
               Begin VB.Menu MNU_Scene_Objects_Prim_Spec_Cornell 
                  Caption         =   "Cornell(R) box"
               End
            End
         End
         Begin VB.Menu Bar21 
            Caption         =   "-"
         End
         Begin VB.Menu MNU_Scene_Objects_Import 
            Caption         =   "Import object..."
         End
         Begin VB.Menu Bar22 
            Caption         =   "-"
         End
         Begin VB.Menu MNU_Scene_Objects_ImportX 
            Caption         =   "&Import DirectX(R) object..."
         End
      End
      Begin VB.Menu MNU_Scene_Lights 
         Caption         =   "Add &Light"
         Begin VB.Menu MNU_Scene_Lights_Sphere 
            Caption         =   "&Omni light..."
         End
         Begin VB.Menu MNU_Scene_Lights_Cone 
            Caption         =   "&Spot light..."
         End
      End
      Begin VB.Menu MNU_Scene_AddCam 
         Caption         =   "Add &Camera"
      End
      Begin VB.Menu Bar4 
         Caption         =   "-"
      End
      Begin VB.Menu MNU_Scene_Manager 
         Caption         =   "Scene &manager..."
      End
      Begin VB.Menu MNU_Scene_SelCamera 
         Caption         =   "Select a came&ra..."
      End
   End
   Begin VB.Menu MNU_View 
      Caption         =   "&Views"
      Begin VB.Menu MNU_View_Refresh 
         Caption         =   "Refresh views"
      End
      Begin VB.Menu Bar5 
         Caption         =   "-"
      End
      Begin VB.Menu MNU_View_All 
         Caption         =   "&All views"
         Shortcut        =   ^A
      End
      Begin VB.Menu MNU_View_Persp 
         Caption         =   "&Perspective view"
         Shortcut        =   ^P
      End
      Begin VB.Menu MNU_View_Front 
         Caption         =   "&Front view"
         Shortcut        =   ^F
      End
      Begin VB.Menu MNU_View_Top 
         Caption         =   "&Top view"
         Shortcut        =   ^T
      End
      Begin VB.Menu MNU_View_Side 
         Caption         =   "&Side view"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu MNU_Render 
      Caption         =   "&Render"
      Begin VB.Menu MNU_Render_Render 
         Caption         =   "&Render scene..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu MNU_Render_Settings 
         Caption         =   "Render &settings..."
      End
   End
   Begin VB.Menu MNU_Help 
      Caption         =   "&Help"
      Begin VB.Menu MNU_Help_Doc 
         Caption         =   "&Documentation..."
      End
      Begin VB.Menu MNU_Help_About 
         Caption         =   "&About the author..."
      End
   End
End
Attribute VB_Name = "FRM_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const GridRes% = 20
Const ThinLinesColor& = 6579300
Const ThikLinesColor& = 255
Const ParallalViewsMouseFactor! = 0.01
Const RotationSpeedFactor! = 2000
Const RotationSpeedFactor2! = 5000

'Mouse point datas:
Dim MousePressed As Boolean, StartX%, StartY%, DiffX%, DiffY%

'Parallal view parameters:
'-------------------------
Dim FrontScale!, FrontMoveX!, FrontMoveY!, OldFrontScale!, OldFrontMoveX!, OldFrontMoveY!
Dim TopScale!, TopMoveX!, TopMoveY!, OldTopScale!, OldTopMoveX!, OldTopMoveY!
Dim SideScale!, SideMoveX!, SideMoveY!, OldSideScale!, OldSideMoveX!, OldSideMoveY!

'Perspective view parameters:
'----------------------------
Dim OldCamPosVec As Vector3D, OldCamDirVec As Vector3D, ZoomVec As Vector3D
Dim PerspFOV!, OldPerspFOV!, OldViewMatrix As Matrix4x4
Dim PerspMoveX!, PerspMoveY!, OldPerspMoveX!, OldPerspMoveY!, MoveVec As Vector3D
Dim RotationVec As Vector3D, RotationXAngle!, RotationYAngle!
Dim PerspRotateZ!, OldPerspRotateZ!
Sub ChooseDisplay(DispId As Integer)

 PIC_Persp.Visible = False
 PIC_Front.Visible = False
 PIC_Top.Visible = False
 PIC_Side.Visible = False

 Select Case DispId

  Case 0: 'All viewports
   Picture5.Left = 272: Picture5.Top = 248
   Picture6.Left = 272: Picture6.Top = 248
   Picture7.Left = 272: Picture7.Top = 248
   Picture8.Left = 160: Picture8.Top = 248
   PIC_Persp.Move 5280, 4560, 5055, 4095: PIC_Persp.Visible = True
   PIC_Front.Move 120, 360, 5055, 4095:   PIC_Front.Visible = True
   PIC_Top.Move 5280, 360, 5055, 4095:    PIC_Top.Visible = True
   PIC_Side.Move 120, 4560, 5055, 4095:   PIC_Side.Visible = True

  Case 1: 'Perspective
   Picture8.Left = 504: Picture8.Top = 528
   PIC_Persp.Move 120, 360, 10215, 8295:  PIC_Persp.Visible = True

  Case 2: 'Front
   Picture5.Left = 616: Picture5.Top = 528
   PIC_Front.Move 120, 360, 10215, 8295:  PIC_Front.Visible = True

  Case 3: 'Top
   Picture6.Left = 616: Picture6.Top = 528
   PIC_Top.Move 120, 360, 10215, 8295:    PIC_Top.Visible = True

  Case 4: 'Side
   Picture7.Left = 616: Picture7.Top = 528
   PIC_Side.Move 120, 360, 10215, 8295:   PIC_Side.Visible = True

 End Select

 RefreshViews

End Sub
Sub DisplayCameraParams()

 With TheCameras(TheCurrentCamera)
  Text1.Text = .Label
  Label6.Caption = "Position :   X " & .Position.X & " ,  Y " & .Position.Y & " ,  Z " & .Position.Z
  Label7.Caption = "Direction :   X " & .Direction.X & " ,  Y " & .Direction.Y & " ,  Z " & .Direction.Z
  Label8.Caption = "FOV angle :  " & RadToDeg(.FOVAngle) & " degrees"
  Label9.Caption = "Roll angle :  " & RadToDeg(.RollAngle) & " degrees"
  If (.BackFaceCulling = False) Then
   Label10.Caption = "Back-faces culling :  Disabled"
  Else
   Label10.Caption = "Back-faces culling :  Enabled"
  End If
 End With

End Sub
Sub DoReset(DisplayMsg As Boolean)

 If (DisplayMsg = False) Then
  MousePointer = 11
  'Set default viewports params:
  FrontScale = 1: TopScale = 1: SideScale = 1: OldFrontScale = 1
  FrontMoveX = 0: FrontMoveY = 0: TopMoveX = 0: TopMoveY = 0: SideMoveX = 0: SideMoveY = 0
  Option1.Value = True: Option4.Value = True: Option6.Value = True
  Check1.Value = vbChecked: Check2.Value = vbChecked: Check3.Value = vbChecked
  Check4.Value = vbChecked: Check5.Value = vbChecked: Check6.Value = vbChecked
  UserAllocation = False
  Engine_Reset
  DisplayMode = 0: ChooseDisplay DisplayMode
  PerspFOV = Wire_DefaultPerspectiveDistorsion
  RefreshViews
  MousePointer = 0
 Else
  If (MsgBox("Are you sure to reset ?", (vbYesNo + vbQuestion), "Reset") = vbYes) Then
   MousePointer = 11
   'Set default viewports params:
   FrontScale = 1: TopScale = 1: SideScale = 1: OldFrontScale = 1
   FrontMoveX = 0: FrontMoveY = 0: TopMoveX = 0: TopMoveY = 0: SideMoveX = 0: SideMoveY = 0
   Option1.Value = True: Option4.Value = True: Option6.Value = True
   Check1.Value = vbChecked: Check2.Value = vbChecked: Check3.Value = vbChecked
   Check4.Value = vbChecked: Check5.Value = vbChecked: Check6.Value = vbChecked
   UserAllocation = False
   Engine_Reset
   DisplayMode = 0: ChooseDisplay DisplayMode
   PerspFOV = Wire_DefaultPerspectiveDistorsion
   RefreshViews
   MousePointer = 0
  End If
 End If

End Sub
Sub DrawGrids()

 Dim CurX%, CurY%

 Select Case DisplayMode

  Case 0: 'All viewports
   'Draw grid:
   PIC_Persp.Cls: PIC_Persp.DrawWidth = 1
   For CurY = 0 To PIC_Persp.ScaleHeight Step GridRes
    For CurX = 0 To PIC_Persp.ScaleWidth Step GridRes
     PIC_Persp.Line (0, CurY)-(PIC_Persp.ScaleWidth, CurY), ThinLinesColor
     PIC_Persp.Line (CurX, 0)-(CurX, PIC_Persp.ScaleHeight), ThinLinesColor
    Next CurX
   Next CurY

   'Draw grid:
   PIC_Front.Cls: PIC_Front.DrawWidth = 1
   For CurY = 0 To PIC_Front.ScaleHeight Step GridRes
    For CurX = 0 To PIC_Front.ScaleWidth Step GridRes
     PIC_Front.Line (0, CurY)-(PIC_Front.ScaleWidth, CurY), ThinLinesColor
     PIC_Front.Line (CurX, 0)-(CurX, PIC_Front.ScaleHeight), ThinLinesColor
    Next CurX
   Next CurY
   'Draw thik crossed lines:
   PIC_Front.DrawWidth = 2
   PIC_Front.Line (0, FrontMoveY + (PIC_Front.ScaleHeight * 0.5))-(PIC_Front.ScaleWidth, FrontMoveY + (PIC_Front.ScaleHeight * 0.5)), ThikLinesColor
   PIC_Front.Line (FrontMoveX + (PIC_Front.ScaleWidth * 0.5), 0)-(FrontMoveX + (PIC_Front.ScaleWidth * 0.5), PIC_Front.ScaleHeight), ThikLinesColor
   PIC_Front.DrawWidth = 1

   'Draw grid:
   PIC_Top.Cls: PIC_Top.DrawWidth = 1
   For CurY = 0 To PIC_Top.ScaleHeight Step GridRes
    For CurX = 0 To PIC_Top.ScaleWidth Step GridRes
     PIC_Top.Line (0, CurY)-(PIC_Top.ScaleWidth, CurY), ThinLinesColor
     PIC_Top.Line (CurX, 0)-(CurX, PIC_Top.ScaleHeight), ThinLinesColor
    Next CurX
   Next CurY
   'Draw thik crossed lines:
   PIC_Top.DrawWidth = 2
   PIC_Top.Line (0, TopMoveY + (PIC_Top.ScaleHeight * 0.5))-(PIC_Top.ScaleWidth, TopMoveY + (PIC_Top.ScaleHeight * 0.5)), ThikLinesColor
   PIC_Top.Line (TopMoveX + (PIC_Top.ScaleWidth * 0.5), 0)-(TopMoveX + (PIC_Top.ScaleWidth * 0.5), PIC_Top.ScaleHeight), ThikLinesColor
   PIC_Top.DrawWidth = 1

   'Draw grid:
   PIC_Side.Cls: PIC_Side.DrawWidth = 1
   For CurY = 0 To PIC_Side.ScaleHeight Step GridRes
    For CurX = 0 To PIC_Side.ScaleWidth Step GridRes
     PIC_Side.Line (0, CurY)-(PIC_Side.ScaleWidth, CurY), ThinLinesColor
     PIC_Side.Line (CurX, 0)-(CurX, PIC_Side.ScaleHeight), ThinLinesColor
    Next CurX
   Next CurY
   'Draw thik crossed lines:
   PIC_Side.DrawWidth = 2
   PIC_Side.Line (0, SideMoveY + (PIC_Side.ScaleHeight * 0.5))-(PIC_Side.ScaleWidth, SideMoveY + (PIC_Side.ScaleHeight * 0.5)), ThikLinesColor
   PIC_Side.Line (SideMoveX + (PIC_Side.ScaleWidth * 0.5), 0)-(SideMoveX + (PIC_Side.ScaleWidth * 0.5), PIC_Side.ScaleHeight), ThikLinesColor
   PIC_Side.DrawWidth = 1

  Case 1: 'Perspective
   'Draw grid:
   PIC_Persp.Cls: PIC_Persp.DrawWidth = 1
   For CurY = 0 To PIC_Persp.ScaleHeight Step GridRes
    For CurX = 0 To PIC_Persp.ScaleWidth Step GridRes
     PIC_Persp.Line (0, CurY)-(PIC_Persp.ScaleWidth, CurY), ThinLinesColor
     PIC_Persp.Line (CurX, 0)-(CurX, PIC_Persp.ScaleHeight), ThinLinesColor
    Next CurX
   Next CurY

  Case 2: 'Front
   'Draw grid:
   PIC_Front.Cls: PIC_Front.DrawWidth = 1
   For CurY = 0 To PIC_Front.ScaleHeight Step GridRes
    For CurX = 0 To PIC_Front.ScaleWidth Step GridRes
     PIC_Front.Line (0, CurY)-(PIC_Front.ScaleWidth, CurY), ThinLinesColor
     PIC_Front.Line (CurX, 0)-(CurX, PIC_Front.ScaleHeight), ThinLinesColor
    Next CurX
   Next CurY
   'Draw thik crossed lines:
   PIC_Front.DrawWidth = 2
   PIC_Front.Line (0, FrontMoveY + (PIC_Front.ScaleHeight * 0.5))-(PIC_Front.ScaleWidth, FrontMoveY + (PIC_Front.ScaleHeight * 0.5)), ThikLinesColor
   PIC_Front.Line (FrontMoveX + (PIC_Front.ScaleWidth * 0.5), 0)-(FrontMoveX + (PIC_Front.ScaleWidth * 0.5), PIC_Front.ScaleHeight), ThikLinesColor
   PIC_Front.DrawWidth = 1

  Case 3: 'Top
   'Draw grid:
   PIC_Top.Cls: PIC_Top.DrawWidth = 1
   For CurY = 0 To PIC_Top.ScaleHeight Step GridRes
    For CurX = 0 To PIC_Top.ScaleWidth Step GridRes
     PIC_Top.Line (0, CurY)-(PIC_Top.ScaleWidth, CurY), ThinLinesColor
     PIC_Top.Line (CurX, 0)-(CurX, PIC_Top.ScaleHeight), ThinLinesColor
    Next CurX
   Next CurY
   'Draw thik crossed lines:
   PIC_Top.DrawWidth = 2
   PIC_Top.Line (0, TopMoveY + (PIC_Top.ScaleHeight * 0.5))-(PIC_Top.ScaleWidth, TopMoveY + (PIC_Top.ScaleHeight * 0.5)), ThikLinesColor
   PIC_Top.Line (TopMoveX + (PIC_Top.ScaleWidth * 0.5), 0)-(TopMoveX + (PIC_Top.ScaleWidth * 0.5), PIC_Top.ScaleHeight), ThikLinesColor
   PIC_Top.DrawWidth = 1

  Case 4: 'Side
   'Draw grid:
   PIC_Side.Cls: PIC_Side.DrawWidth = 1
   For CurY = 0 To PIC_Side.ScaleHeight Step GridRes
    For CurX = 0 To PIC_Side.ScaleWidth Step GridRes
     PIC_Side.Line (0, CurY)-(PIC_Side.ScaleWidth, CurY), ThinLinesColor
     PIC_Side.Line (CurX, 0)-(CurX, PIC_Side.ScaleHeight), ThinLinesColor
    Next CurX
   Next CurY
   'Draw thik crossed lines:
   PIC_Side.DrawWidth = 2
   PIC_Side.Line (0, SideMoveY + (PIC_Side.ScaleHeight * 0.5))-(PIC_Side.ScaleWidth, SideMoveY + (PIC_Side.ScaleHeight * 0.5)), ThikLinesColor
   PIC_Side.Line (SideMoveX + (PIC_Side.ScaleWidth * 0.5), 0)-(SideMoveX + (PIC_Side.ScaleWidth * 0.5), PIC_Side.ScaleHeight), ThikLinesColor
   PIC_Side.DrawWidth = 1

 End Select

End Sub
Sub DrawScene()

 Dim Geo As Boolean, Nrm As Boolean, Sph As Boolean
 Dim Con As Boolean, Cam As Boolean, Pht As Boolean
 Dim OldCameraFrom As Vector3D, OldCameraTo As Vector3D, Tmp!

 If (Check1.Value = vbChecked) Then Geo = True
 If (Check2.Value = vbChecked) Then Sph = True
 If (Check3.Value = vbChecked) Then Con = True
 If (Check4.Value = vbChecked) Then Cam = True
 If (Check5.Value = vbChecked) Then Pht = True
 If (Check6.Value = vbChecked) Then Nrm = True

 Select Case DisplayMode

  Case 0: 'All viewports
   'Perspective
   Wire_PerspectiveDistorsion = PerspFOV: Tmp = Wire_AddedDepth: Wire_AddedDepth = 0
   Engine_WireframePreview PIC_Persp, Geo, Nrm, Sph, Con, Cam, Pht, True
   Wire_AddedDepth = Tmp

   OldCameraFrom = TheCameras(TheCurrentCamera).Position
   OldCameraTo = TheCameras(TheCurrentCamera).Direction
   TheCameras(TheCurrentCamera).Direction = VectorNull

   'Front
   Wire_ParallalScale = FrontScale: Wire_ParallalMoveToX = FrontMoveX: Wire_ParallalMoveToY = FrontMoveY
   TheCameras(TheCurrentCamera).Position = VectorInput(ApproachVal, 0, -1000)
   Engine_WireframePreview PIC_Front, Geo, Nrm, Sph, Con, Cam, Pht, False

   'Top
   Wire_ParallalScale = TopScale: Wire_ParallalMoveToX = TopMoveX: Wire_ParallalMoveToY = TopMoveY
   TheCameras(TheCurrentCamera).Position = VectorInput(ApproachVal, -1000, 0)
   Engine_WireframePreview PIC_Top, Geo, Nrm, Sph, Con, Cam, Pht, False

   'Side
   Wire_ParallalScale = SideScale: Wire_ParallalMoveToX = SideMoveX: Wire_ParallalMoveToY = SideMoveY
   TheCameras(TheCurrentCamera).Position = VectorInput(-1000, 0, ApproachVal)
   Engine_WireframePreview PIC_Side, Geo, Nrm, Sph, Con, Cam, Pht, False

   Wire_ParallalScale = 1: Wire_ParallalMoveToX = 0: Wire_ParallalMoveToY = 0
   TheCameras(TheCurrentCamera).Position = OldCameraFrom
   TheCameras(TheCurrentCamera).Direction = OldCameraTo

  Case 1: 'Perspective
   Wire_PerspectiveDistorsion = PerspFOV: Tmp = Wire_AddedDepth: Wire_AddedDepth = 0
   Engine_WireframePreview PIC_Persp, Geo, Nrm, Sph, Con, Cam, Pht, True
   Wire_AddedDepth = Tmp

  Case 2: 'Front
   OldCameraFrom = TheCameras(TheCurrentCamera).Position
   OldCameraTo = TheCameras(TheCurrentCamera).Direction
   TheCameras(TheCurrentCamera).Position = VectorInput(ApproachVal, 0, -1000)
   TheCameras(TheCurrentCamera).Direction = VectorNull
   Wire_ParallalScale = FrontScale: Wire_ParallalMoveToX = FrontMoveX: Wire_ParallalMoveToY = FrontMoveY
   Engine_WireframePreview PIC_Front, Geo, Nrm, Sph, Con, Cam, Pht, False
   Wire_ParallalScale = 1: Wire_ParallalMoveToX = 0: Wire_ParallalMoveToY = 0
   TheCameras(TheCurrentCamera).Position = OldCameraFrom
   TheCameras(TheCurrentCamera).Direction = OldCameraTo

  Case 3: 'Top
   OldCameraFrom = TheCameras(TheCurrentCamera).Position
   OldCameraTo = TheCameras(TheCurrentCamera).Direction
   TheCameras(TheCurrentCamera).Position = VectorInput(ApproachVal, -1000, 0)
   TheCameras(TheCurrentCamera).Direction = VectorNull
   Wire_ParallalScale = TopScale: Wire_ParallalMoveToX = TopMoveX: Wire_ParallalMoveToY = TopMoveY
   Engine_WireframePreview PIC_Top, Geo, Nrm, Sph, Con, Cam, Pht, False
   Wire_ParallalScale = 1: Wire_ParallalMoveToX = 0: Wire_ParallalMoveToY = 0
   TheCameras(TheCurrentCamera).Position = OldCameraFrom
   TheCameras(TheCurrentCamera).Direction = OldCameraTo

  Case 4: 'Side
   OldCameraFrom = TheCameras(TheCurrentCamera).Position
   OldCameraTo = TheCameras(TheCurrentCamera).Direction
   TheCameras(TheCurrentCamera).Position = VectorInput(-1000, 0, ApproachVal)
   TheCameras(TheCurrentCamera).Direction = VectorNull
   Wire_ParallalScale = SideScale: Wire_ParallalMoveToX = SideMoveX: Wire_ParallalMoveToY = SideMoveY
   Engine_WireframePreview PIC_Side, Geo, Nrm, Sph, Con, Cam, Pht, False
   Wire_ParallalScale = 1: Wire_ParallalMoveToX = 0: Wire_ParallalMoveToY = 0
   TheCameras(TheCurrentCamera).Position = OldCameraFrom
   TheCameras(TheCurrentCamera).Direction = OldCameraTo

 End Select

End Sub
Sub RefreshViews()

 MousePointer = 11
 DoEvents

 DrawGrids
 DrawScene
 DisplayCameraParams

 MousePointer = 0

End Sub
Private Sub Check1_Click()

 RefreshViews

End Sub

Private Sub Check2_Click()

 RefreshViews

End Sub

Private Sub Check3_Click()

 RefreshViews

End Sub

Private Sub Check4_Click()

 RefreshViews

End Sub

Private Sub Check5_Click()

 RefreshViews

End Sub

Private Sub Check6_Click()

 RefreshViews

End Sub

Private Sub Command1_Click()

 FRM_SelectCamera.Show 1
 RefreshViews

End Sub
Private Sub Command15_Click()

 RefreshViews

End Sub

Private Sub Command2_Click()

 MNU_File_Reset_Click

End Sub

Private Sub Command3_Click()

 MNU_File_Open_Click

End Sub

Private Sub Command4_Click()

 MNU_File_Save_Click

End Sub
Private Sub Command5_Click()

 MNU_Scene_Manager_Click

End Sub

Private Sub Command6_Click()

 MNU_Render_Render_Click

End Sub

Private Sub Command7_Click()

 MNU_Render_Settings_Click

End Sub

Private Sub Command8_Click()

 MNU_Help_Doc_Click

End Sub
Private Sub Command9_Click()

 MNU_Help_About_Click

End Sub
Private Sub Form_Activate()

 ChooseDisplay DisplayMode

End Sub
Private Sub Form_Load()

 Engine_Start

 'Set default viewports params:
 FrontScale = 1: TopScale = 1: SideScale = 1
 OldFrontScale = 1: PerspFOV = Wire_DefaultPerspectiveDistorsion

 DisplayMode = 0

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

 Dim StrMsg As String

 StrMsg = "3D Digital Reality Engine V1.00, Pure VB ! & FREE (for non-commercial use), just give me some votes !" & vbNewLine & vbNewLine & _
          "                                                  KACI Lounes - 2009"

 MsgBox StrMsg, vbInformation, "So long !": End

End Sub
Private Sub Label1_Click()

 If (DisplayMode = 0) Then
  DisplayMode = 2: ChooseDisplay DisplayMode
 ElseIf (DisplayMode = 2) Then
  DisplayMode = 0: ChooseDisplay DisplayMode
 End If

End Sub

Private Sub Label2_Click()

 If (DisplayMode = 0) Then
  DisplayMode = 3: ChooseDisplay DisplayMode
 ElseIf (DisplayMode = 3) Then
  DisplayMode = 0: ChooseDisplay DisplayMode
 End If

End Sub

Private Sub Label3_Click()

 If (DisplayMode = 0) Then
  DisplayMode = 4: ChooseDisplay DisplayMode
 ElseIf (DisplayMode = 4) Then
  DisplayMode = 0: ChooseDisplay DisplayMode
 End If

End Sub
Private Sub Label4_Click()

 If (DisplayMode = 0) Then
  DisplayMode = 1: ChooseDisplay DisplayMode
 ElseIf (DisplayMode = 1) Then
  DisplayMode = 0: ChooseDisplay DisplayMode
 End If

End Sub
Private Sub MNU_File_Exit_Click()

 If (MsgBox("What are you talking about ?!", (vbYesNo + vbExclamation), "Exit") = vbYes) Then

  Dim StrMsg As String
  StrMsg = "3D Digital Reality Engine V1.00, Pure VB ! & FREE (for non-commercial use), just give me some votes !" & vbNewLine & vbNewLine & _
           "                                                  KACI Lounes - 2009"

  MsgBox StrMsg, vbInformation, "So long !": End

 End If

End Sub

Private Sub MNU_File_Open_Click()

 COMDLG.FileName = vbNullString
 COMDLG.Filter = "3D Digital Reality Engine scenes files (*" & SceneFileExtension & ")|*" & SceneFileExtension & "|"
 COMDLG.InitDir = App.Path & "\Datas\Scenes\"
 COMDLG.ShowOpen

 If (COMDLG.FileName <> vbNullString) Then Engine_LoadScene COMDLG.FileName

End Sub
Private Sub MNU_File_Reset_Click()

 DoReset True

End Sub

Private Sub MNU_File_Save_Click()

 COMDLG.FileName = vbNullString
 COMDLG.Filter = "3D Digital Reality Engine scenes files (*" & SceneFileExtension & ")|*" & SceneFileExtension & "|"
 COMDLG.InitDir = App.Path & "\Datas\Scenes\"
 COMDLG.ShowSave

 If (COMDLG.FileName <> vbNullString) Then Engine_SaveScene COMDLG.FileName

End Sub
Private Sub MNU_Help_About_Click()

 FRM_About.Show 1

End Sub

Private Sub MNU_Help_Doc_Click()

 MousePointer = 11
 ShellExecute 0, "open", (App.Path & "\Manual\3D Digital Reality Engine.htm"), "", "", 1
 MousePointer = 0

End Sub
Private Sub MNU_Render_Render_Click()

 On Error Resume Next

 Unload FRM_Manager
 FRM_Render.Show 1

End Sub
Private Sub MNU_Render_Settings_Click()

 FRM_RenderSettings.Show 1

End Sub
Private Sub MNU_Scene_AddCam_Click()

 CamWindowMode = False
 FRM_AddCamera.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Lights_Cone_Click()

 SpotWindowMode = False
 FRM_Lights_Spot.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Lights_Sphere_Click()

 OmniWindowMode = False
 FRM_Lights_Omni.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Manager_Click()

 FRM_Manager.Show

End Sub

Private Sub MNU_Scene_Objects_Import_Click()

 COMDLG.FileName = vbNullString
 COMDLG.Filter = "3D Digital Reality Engine objects files (*" & ObjectFileExtension & ")|*" & ObjectFileExtension & "|"
 COMDLG.InitDir = App.Path & "\Datas\Objects\"
 COMDLG.ShowOpen

 If (COMDLG.FileName <> vbNullString) Then Engine_LoadMesh COMDLG.FileName: RefreshViews

End Sub
Private Sub MNU_Scene_Objects_ImportX_Click()

 FRM_ImportXFile.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Box_Click()

 FRM_Primitives_Box.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Capsule_Click()

 FRM_Primitives_Capsule.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Cone_Click()

 FRM_Primitives_Cone.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Cylinder_Click()

 FRM_Primitives_Cylinder.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Disk_Click()

 FRM_Primitives_Disk.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_FlatTube_Click()

 FRM_Primitives_TubeFlat.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Grid_Click()

 FRM_Primitives_Grid.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Hemisphere_Click()

 FRM_Primitives_Hemisphere.Show 1
 RefreshViews

End Sub


Private Sub MNU_Scene_Objects_Prim_Octahedron_Click()

 FRM_Primitives_Octahedron.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Pyramid_Click()

 FRM_Primitives_Pyramid.Show 1
 RefreshViews

End Sub

Private Sub MNU_Scene_Objects_Prim_Spec_Cornell_Click()

 Primitive_CornellBox True, VectorNull, 0, False
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Spec_Landscape_Click()

 FRM_Primitives_Landscape.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Sphere_Click()

 FRM_Primitives_Sphere.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Tetrahedron_Click()

 FRM_Primitives_Tetrahedron.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Torus_Click()

 FRM_Primitives_Torus.Show 1
 RefreshViews

End Sub
Private Sub MNU_Scene_Objects_Prim_Tube_Click()

 FRM_Primitives_Tube.Show 1
 RefreshViews

End Sub

Private Sub MNU_Scene_SelCamera_Click()

 FRM_SelectCamera.Show 1
 RefreshViews

End Sub
Private Sub MNU_View_All_Click()

 DisplayMode = 0: ChooseDisplay DisplayMode

End Sub
Private Sub MNU_View_Front_Click()

 DisplayMode = 2: ChooseDisplay DisplayMode

End Sub
Private Sub MNU_View_Side_Click()

 DisplayMode = 4: ChooseDisplay DisplayMode

End Sub

Private Sub MNU_View_Persp_Click()

 DisplayMode = 1: ChooseDisplay DisplayMode

End Sub

Private Sub MNU_View_Refresh_Click()

 RefreshViews

End Sub
Private Sub MNU_View_Top_Click()

 DisplayMode = 3: ChooseDisplay DisplayMode

End Sub

Private Sub PIC_Front_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

 MousePressed = True
 StartX = X: StartY = Y

 OldFrontMoveX = FrontMoveX
 OldFrontMoveY = FrontMoveY
 OldFrontScale = FrontScale

End Sub
Private Sub PIC_Front_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If (MousePressed = True) Then
  DiffX = (X - StartX): DiffY = (Y - StartY)
  If (Option1.Value = True) Then 'Translate
   FrontMoveX = (OldFrontMoveX + DiffX)
   FrontMoveY = (OldFrontMoveY + DiffY)
  Else 'Scale
   FrontScale = (OldFrontScale + (DiffY * ParallalViewsMouseFactor))
  End If
  RefreshViews
 End If

End Sub
Private Sub PIC_Front_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 MousePressed = False

End Sub
Private Sub PIC_Persp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

 MousePressed = True
 StartX = X: StartY = Y

 OldCamPosVec = TheCameras(TheCurrentCamera).Position
 OldCamDirVec = TheCameras(TheCurrentCamera).Direction

 OldPerspFOV = PerspFOV
 OldPerspRotateZ = PerspRotateZ

 OldViewMatrix = MatrixView(VectorNormalize(VectorSubtract(TheCameras(TheCurrentCamera).Position, TheCameras(TheCurrentCamera).Direction)), VectorNull, TheCameras(TheCurrentCamera).RollAngle)

End Sub
Private Sub PIC_Persp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If (MousePressed = True) Then
  DiffX = (X - StartX): DiffY = (Y - StartY)

  If (Option7.Value = True) Then      'Zoom
   ZoomVec = VectorScale(VectorNormalize(VectorSubtract(OldCamDirVec, OldCamPosVec)), CSng(DiffY))
   TheCameras(TheCurrentCamera).Position = VectorAdd(OldCamPosVec, ZoomVec)
   TheCameras(TheCurrentCamera).Direction = VectorAdd(OldCamDirVec, ZoomVec)

  ElseIf (Option8.Value = True) Then  'FOV
   PerspFOV = Abs(OldPerspFOV + DiffY)

  ElseIf (Option9.Value = True) Then  'Translate
   MoveVec = MatrixMultiplyVector(VectorInput(0, CSng(-DiffY), CSng(DiffX)), OldViewMatrix)
   TheCameras(TheCurrentCamera).Position = VectorAdd(OldCamPosVec, MoveVec)
   TheCameras(TheCurrentCamera).Direction = VectorAdd(OldCamDirVec, MoveVec)

  ElseIf (Option10.Value = True) Then 'Rotate XY
   RotationXAngle = ((DiffY / RotationSpeedFactor) * Pi2)
   RotationYAngle = ((-DiffX / RotationSpeedFactor) * Pi2)
   RotationVec = VectorRotate(VectorSubtract(OldCamDirVec, OldCamPosVec), 0, RotationXAngle)
   RotationVec = VectorRotate(RotationVec, 1, RotationYAngle)
   TheCameras(TheCurrentCamera).Direction = VectorAdd(OldCamDirVec, RotationVec)

  ElseIf (Option11.Value = True) Then 'Rotate Z (Roll)
   PerspRotateZ = (OldPerspRotateZ + ((DiffY / RotationSpeedFactor) * Pi2))
   TheCameras(TheCurrentCamera).RollAngle = PerspRotateZ

  End If

  RefreshViews
 End If

End Sub
Private Sub PIC_Persp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 MousePressed = False

End Sub
Private Sub PIC_Side_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

 MousePressed = True
 StartX = X: StartY = Y

 OldSideMoveX = SideMoveX
 OldSideMoveY = SideMoveY
 OldSideScale = SideScale

End Sub
Private Sub PIC_Side_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If (MousePressed = True) Then
  DiffX = (X - StartX): DiffY = (Y - StartY)
  If (Option6.Value = True) Then 'Translate
   SideMoveX = (OldSideMoveX + DiffX)
   SideMoveY = (OldSideMoveY + DiffY)
  Else 'Scale
   SideScale = (OldSideScale + (DiffY * ParallalViewsMouseFactor))
  End If
  RefreshViews
 End If

End Sub
Private Sub PIC_Side_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 MousePressed = False

End Sub
Private Sub PIC_Top_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

 MousePressed = True
 StartX = X: StartY = Y

 OldTopMoveX = TopMoveX
 OldTopMoveY = TopMoveY
 OldTopScale = TopScale

End Sub
Private Sub PIC_Top_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 If (MousePressed = True) Then
  DiffX = (X - StartX): DiffY = (Y - StartY)
  If (Option4.Value = True) Then 'Translate
   TopMoveX = (OldTopMoveX + DiffX)
   TopMoveY = (OldTopMoveY + DiffY)
  Else 'Scale
   TopScale = (OldTopScale + (DiffY * ParallalViewsMouseFactor))
  End If
  RefreshViews
 End If

End Sub
Private Sub PIC_Top_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 MousePressed = False

End Sub
