VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRM_RenderSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Render settings"
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_RenderSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   664
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   576
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame13 
      Caption         =   "Output size : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   3840
      TabIndex        =   82
      Top             =   4320
      Width           =   4695
      Begin VB.CommandButton Command7 
         Caption         =   "Models..."
         Height          =   570
         Left            =   2880
         TabIndex        =   94
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   84
         Text            =   "640"
         Top             =   330
         Width           =   975
      End
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   83
         Text            =   "480"
         Top             =   690
         Width           =   975
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Image width :"
         Height          =   195
         Left            =   465
         TabIndex        =   86
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Image height : "
         Height          =   195
         Left            =   405
         TabIndex        =   85
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Fog effect : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   5160
      TabIndex        =   60
      Top             =   1440
      Width           =   3375
      Begin VB.CommandButton Command2 
         Caption         =   "Fog parameters..."
         Height          =   615
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   6
      Top             =   9120
      Width           =   4695
      Begin VB.CommandButton Command4 
         Caption         =   "&Update"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   80
         Top             =   240
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog COMDLG 
         Left            =   120
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Memory : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4095
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      Begin VB.Frame Frame9 
         Caption         =   "Memory usage : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2415
         Begin VB.Label Label45 
            Height          =   195
            Left            =   1215
            TabIndex        =   90
            Top             =   2880
            Width           =   1065
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BackGround : "
            Height          =   195
            Left            =   135
            TabIndex        =   89
            Top             =   2880
            Width           =   1005
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Lights : "
            Height          =   195
            Left            =   570
            TabIndex        =   35
            Top             =   1800
            Width           =   570
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vertices : "
            Height          =   195
            Left            =   420
            TabIndex        =   34
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Meshs : "
            Height          =   195
            Left            =   540
            TabIndex        =   33
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Faces : "
            Height          =   195
            Left            =   570
            TabIndex        =   32
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cameras : "
            Height          =   195
            Left            =   360
            TabIndex        =   31
            Top             =   2160
            Width           =   780
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Textures : "
            Height          =   195
            Left            =   345
            TabIndex        =   30
            Top             =   1440
            Width           =   795
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Photon-map : "
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   2520
            Width           =   1020
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total use : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   225
            TabIndex        =   28
            Top             =   3240
            Width           =   915
         End
         Begin VB.Label Label9 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1215
            TabIndex        =   27
            Top             =   3240
            Width           =   1065
         End
         Begin VB.Label Label8 
            Height          =   195
            Left            =   1215
            TabIndex        =   26
            Top             =   2520
            Width           =   1065
         End
         Begin VB.Label Label7 
            Height          =   195
            Left            =   1215
            TabIndex        =   25
            Top             =   1440
            Width           =   1065
         End
         Begin VB.Label Label6 
            Height          =   195
            Left            =   1215
            TabIndex        =   24
            Top             =   2160
            Width           =   1065
         End
         Begin VB.Label Label5 
            Height          =   195
            Left            =   1215
            TabIndex        =   23
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Label4 
            Height          =   195
            Left            =   1215
            TabIndex        =   22
            Top             =   1080
            Width           =   1065
         End
         Begin VB.Label Label3 
            Height          =   195
            Left            =   1215
            TabIndex        =   21
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label2 
            Height          =   195
            Left            =   1215
            TabIndex        =   20
            Top             =   1800
            Width           =   1065
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Allocations : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton Command1 
            Caption         =   "Re-allocate memory..."
            Height          =   615
            Left            =   120
            TabIndex        =   18
            Top             =   2880
            Width           =   1935
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            TabIndex        =   16
            Text            =   "100"
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            TabIndex        =   11
            Text            =   "5000"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            TabIndex        =   10
            Text            =   "100"
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            TabIndex        =   9
            Text            =   "5000"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            TabIndex        =   8
            Text            =   "100"
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cameras : "
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Faces : "
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Meshs : "
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Vertices : "
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Lights : "
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1920
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Background : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton Command3 
         Caption         =   "Browse..."
         Height          =   450
         Left            =   2040
         TabIndex        =   75
         Top             =   315
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use a background"
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   540
         Width           =   1695
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Color"
         Height          =   195
         Left            =   1800
         TabIndex        =   88
         Top             =   885
         Width           =   375
      End
      Begin VB.Label Label43 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         MousePointer    =   2  'Cross
         TabIndex        =   87
         Top             =   855
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Filtering : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3495
      Left            =   3840
      TabIndex        =   3
      Top             =   5640
      Width           =   4695
      Begin VB.CommandButton Command6 
         Caption         =   "Recreate MIP-maps"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3120
         TabIndex        =   93
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text19 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   71
         Text            =   "30"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   70
         Text            =   "5"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   68
         Text            =   "2"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   66
         Text            =   "0.5"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   64
         Text            =   "0.5"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   62
         Text            =   "-0.5"
         Top             =   2280
         Width           =   735
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FRM_RenderSettings.frx":000C
         Left            =   2640
         List            =   "FRM_RenderSettings.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   435
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FRM_RenderSettings.frx":0099
         Left            =   2640
         List            =   "FRM_RenderSettings.frx":00AF
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   1875
         Width           =   1935
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "MIP-maps min purcent : "
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "MIP-maps level : "
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "Kernel size : "
         Enabled         =   0   'False
         Height          =   255
         Left            =   2400
         TabIndex        =   69
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "Cubic C :"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   67
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Cubic B :"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   65
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "Cubic A :"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   63
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Texels filter (anti-magnification) : "
         Height          =   195
         Left            =   225
         TabIndex        =   59
         Top             =   1920
         Width           =   2430
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Textures filter (anti-minification) : "
         Height          =   195
         Left            =   195
         TabIndex        =   58
         Top             =   480
         Width           =   2460
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Global illumination :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   3615
      Begin VB.OptionButton Option4 
         Caption         =   "Photon mapping (slow)"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   720
         Width           =   1935
      End
      Begin VB.Frame Frame11 
         Height          =   2415
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   3375
         Begin VB.TextBox Text22 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            TabIndex        =   91
            Text            =   "30"
            Top             =   1560
            Width           =   975
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Allocate"
            Height          =   255
            Left            =   2280
            TabIndex        =   81
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            TabIndex        =   53
            Text            =   "20"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            TabIndex        =   52
            Text            =   "30"
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            TabIndex        =   49
            Text            =   "20"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2160
            TabIndex        =   48
            Text            =   "30"
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            Caption         =   "Color-bleeding distance  : "
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Search radius : "
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Estimate multiplier : "
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Photons count : "
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Bounces per photon path : "
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Flat ambiant"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Label Label39 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         MousePointer    =   2  'Cross
         TabIndex        =   77
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Caption         =   "Ambiant color"
         Height          =   255
         Left            =   1920
         TabIndex        =   76
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Shadows mode : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   3615
      Begin VB.OptionButton Option2 
         Caption         =   "Area shadows (soft && very slow)"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   720
         Width           =   2655
      End
      Begin VB.Frame Frame10 
         Height          =   1335
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   3375
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1920
            TabIndex        =   43
            Text            =   "30"
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1920
            TabIndex        =   42
            Text            =   "20"
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Shadow width : "
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "Shadow rays count : "
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Flat shadows (1 shadow ray)"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stochastic path-tracer parameters : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1455
      Left            =   5160
      TabIndex        =   0
      Top             =   2760
      Width           =   3375
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   38
         Text            =   "5000"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   36
         Text            =   "5000"
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Bounces per view path : "
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "View paths per pixel : "
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   960
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FRM_RenderSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DontAsk1 As Boolean, DontAsk2 As Boolean
Sub RefreshForm()

 Dim BytesV&, BytesF&, BytesM&, BytesT&, BytesL&, BytesC&, BytesP&, BytesB&, BytesAll&
 Dim StrBytesV$, StrBytesF$, StrBytesM$, StrBytesT$, StrBytesL$, StrBytesC$, StrBytesP$, StrBytesB$, StrBytesAll$

 BytesV = Engine_ComputeUsedMemory(1)
 BytesF = Engine_ComputeUsedMemory(2)
 BytesM = Engine_ComputeUsedMemory(3)
 BytesT = Engine_ComputeUsedMemory(4)
 BytesL = Engine_ComputeUsedMemory(5)
 BytesC = Engine_ComputeUsedMemory(6)
 BytesP = Engine_ComputeUsedMemory(7)
 BytesB = Engine_ComputeUsedMemory(8)
 BytesAll = (BytesV + BytesF + BytesM + BytesT + BytesL + BytesC + BytesP + BytesB)

 If (BytesV < 1024) Then
  StrBytesV = BytesV & " Bytes"
 Else
  BytesV = (BytesV / 1024)
  If (BytesV < 1024) Then
   StrBytesV = BytesV & " Kb"
  Else
   BytesV = (BytesV / 1024)
   If (BytesV < 1024) Then
    StrBytesV = BytesV & " Mb"
   Else
    BytesV = (BytesV / 1024)
    If (BytesV < 1024) Then
     StrBytesV = BytesV & " Gb"
    End If
   End If
  End If
 End If

 If (BytesF < 1024) Then
  StrBytesF = BytesF & " Bytes"
 Else
  BytesF = (BytesF / 1024)
  If (BytesF < 1024) Then
   StrBytesF = BytesF & " Kb"
  Else
   BytesF = (BytesF / 1024)
   If (BytesF < 1024) Then
    StrBytesF = BytesF & " Mb"
   Else
    BytesF = (BytesF / 1024)
    If (BytesF < 1024) Then
     StrBytesF = BytesF & " Gb"
    End If
   End If
  End If
 End If

 If (BytesM < 1024) Then
  StrBytesM = BytesM & " Bytes"
 Else
  BytesM = (BytesM / 1024)
  If (BytesM < 1024) Then
   StrBytesM = BytesM & " Kb"
  Else
   BytesM = (BytesM / 1024)
   If (BytesM < 1024) Then
    StrBytesM = BytesM & " Mb"
   Else
    BytesM = (BytesM / 1024)
    If (BytesM < 1024) Then
     StrBytesM = BytesM & " Gb"
    End If
   End If
  End If
 End If

 If (BytesT < 1024) Then
  StrBytesT = BytesT & " Bytes"
 Else
  BytesT = (BytesT / 1024)
  If (BytesT < 1024) Then
   StrBytesT = BytesT & " Kb"
  Else
   BytesT = (BytesT / 1024)
   If (BytesT < 1024) Then
    StrBytesT = BytesT & " Mb"
   Else
    BytesT = (BytesT / 1024)
    If (BytesT < 1024) Then
     StrBytesT = BytesT & " Gb"
    End If
   End If
  End If
 End If

 If (BytesL < 1024) Then
  StrBytesL = BytesL & " Bytes"
 Else
  BytesL = (BytesL / 1024)
  If (BytesL < 1024) Then
   StrBytesL = BytesL & " Kb"
  Else
   BytesL = (BytesL / 1024)
   If (BytesL < 1024) Then
    StrBytesL = BytesL & " Mb"
   Else
    BytesL = (BytesL / 1024)
    If (BytesL < 1024) Then
     StrBytesL = BytesL & " Gb"
    End If
   End If
  End If
 End If

 If (BytesC < 1024) Then
  StrBytesC = BytesC & " Bytes"
 Else
  BytesC = (BytesC / 1024)
  If (BytesC < 1024) Then
   StrBytesC = BytesC & " Kb"
  Else
   BytesC = (BytesC / 1024)
   If (BytesC < 1024) Then
    StrBytesC = BytesC & " Mb"
   Else
    BytesC = (BytesC / 1024)
    If (BytesC < 1024) Then
     StrBytesC = BytesC & " Gb"
    End If
   End If
  End If
 End If

 If (BytesP < 1024) Then
  StrBytesP = BytesP & " Bytes"
 Else
  BytesP = (BytesP / 1024)
  If (BytesP < 1024) Then
   StrBytesP = BytesP & " Kb"
  Else
   BytesP = (BytesP / 1024)
   If (BytesP < 1024) Then
    StrBytesP = BytesP & " Mb"
   Else
    BytesP = (BytesP / 1024)
    If (BytesP < 1024) Then
     StrBytesP = BytesP & " Gb"
    End If
   End If
  End If
 End If

 If (BytesB < 1024) Then
  StrBytesB = BytesB & " Bytes"
 Else
  BytesB = (BytesB / 1024)
  If (BytesB < 1024) Then
   StrBytesB = BytesB & " Kb"
  Else
   BytesB = (BytesB / 1024)
   If (BytesB < 1024) Then
    StrBytesB = BytesB & " Mb"
   Else
    BytesB = (BytesB / 1024)
    If (BytesB < 1024) Then
     StrBytesB = BytesB & " Gb"
    End If
   End If
  End If
 End If

 If (BytesAll < 1024) Then
  StrBytesAll = BytesAll & " Bytes"
 Else
  BytesAll = (BytesAll / 1024)
  If (BytesAll < 1024) Then
   StrBytesAll = BytesAll & " Kb"
  Else
   BytesAll = (BytesAll / 1024)
   If (BytesAll < 1024) Then
    StrBytesAll = BytesAll & " Mb"
   Else
    BytesAll = (BytesAll / 1024)
    If (BytesAll < 1024) Then
     StrBytesAll = BytesAll & " Gb"
    End If
   End If
  End If
 End If

 Label3.Caption = StrBytesV: Label5.Caption = StrBytesF: Label4.Caption = StrBytesM
 Label7.Caption = StrBytesT: Label2.Caption = StrBytesL: Label6.Caption = StrBytesC
 Label8.Caption = StrBytesP: Label45.Caption = StrBytesB: Label9.Caption = StrBytesAll

 '---------------------------------------------------

 Text1.Text = MaxVertices: Text2.Text = MaxFaces: Text3.Text = MaxMeshs
 Text4.Text = MaxLights: Text5.Text = MaxCameras

 '---------------------------------------------------

 Text7.Text = SamplesPerViewPath: Text6.Text = ViewPathsPerPixel

 '---------------------------------------------------

 If (UseBackGround = True) Then
  Check1.Value = vbChecked: Command3.Enabled = True
  Label43.Enabled = False: Label44.Enabled = False
 Else
  Check1.Value = vbUnchecked: Command3.Enabled = False
  Label43.Enabled = True: Label44.Enabled = True
  Label43.BackColor = ColorRGBToLong(TheBackGroundColor)
 End If

 '---------------------------------------------------

 Text8.Text = ShadowsApproxRadius
 Text9.Text = ShadowRaysCount

 If (EnableAreaShadows = False) Then
  Option1.Value = True
  Label24.Enabled = False
  Label25.Enabled = False
  Text8.Enabled = False
  Text9.Enabled = False
 Else
  Option2.Value = True
  Label24.Enabled = True
  Label25.Enabled = True
  Text8.Enabled = True
  Text9.Enabled = True
 End If

 '---------------------------------------------------

 Text10.Text = SamplesPerPhotonPath
 Text11.Text = MaximumAllocatedPhotons
 Text12.Text = EstimateMultiplier
 Text13.Text = PhotonsSearchRadius
 Text22.Text = BleedingDistance

 If (EnablePhotonMapping = False) Then
  DontAsk1 = True
  Option3.Value = True
  Label38.Enabled = True
  Label39.Enabled = True
  Label26.Enabled = False
  Label27.Enabled = False
  Label28.Enabled = False
  Label29.Enabled = False
  Label46.Enabled = False
  Text10.Enabled = False
  Text11.Enabled = False
  Text12.Enabled = False
  Text13.Enabled = False
  Text22.Enabled = False
  Command5.Enabled = False
  Label39.BackColor = ColorRGBToLong(TheAmbiantLight)
 Else
  DontAsk2 = True
  Option4.Value = True
  Label38.Enabled = False
  Label39.Enabled = False
  Label26.Enabled = True
  Label27.Enabled = True
  Label28.Enabled = True
  Label29.Enabled = True
  Label46.Enabled = True
  Text10.Enabled = True
  Text11.Enabled = True
  Text12.Enabled = True
  Text13.Enabled = True
  Text22.Enabled = True
  Command5.Enabled = True
 End If

 '---------------------------------------------------

 Text14.Text = CubicA: Text15.Text = CubicB
 Text16.Text = CubicC: Text17.Text = KernelSize

 Select Case TheTexturesFilter
  Case K3DE_TFM_NEAREST:               Combo2.Text = "Nearest"
  Case K3DE_TFM_NEAREST_MIP_NEAREST:   Combo2.Text = "Nearest mip nearest"
  Case K3DE_TFM_NEAREST_MIP_LINEAR:    Combo2.Text = "Nearest mip linear"
  Case K3DE_TFM_FILTERED:               Combo2.Text = "FILTERED"
  Case K3DE_TFM_FILTERED_MIP_NEAREST:   Combo2.Text = "FILTERED mip nearest"
  Case K3DE_TFM_FILTERED_MIP_LINEAR:    Combo2.Text = "FILTERED mip linear (Trilinear)"
 End Select

 Select Case TheTexelsFilter
  Case K3DE_XFM_BILINEAR:              Combo1.Text = "Bilinear"
  Case K3DE_XFM_BELL:                  Combo1.Text = "Bell"
  Case K3DE_XFM_GAUSSIAN:              Combo1.Text = "Gaussian"
  Case K3DE_XFM_CUBIC_SPLINE_B:        Combo1.Text = "Cubic spline B"
  Case K3DE_XFM_CUBIC_SPLINE_BC:       Combo1.Text = "Cubic spline BC"
  Case K3DE_XFM_CUBIC_SPLINE_CARDINAL: Combo1.Text = "Cubic spline cardinal"
 End Select

 '---------------------------------------------------

 Text21.Text = OutputWidth: Text20.Text = OutputHeight

End Sub
Private Sub Check1_Click()

 UseBackGround = CheckOut(Check1)
 Command3.Enabled = CheckOut(Check1)
 Label43.Enabled = Not CheckOut(Check1)
 Label44.Enabled = Not CheckOut(Check1)

 If ((UseBackGround = False) And ((OriginalBackGround.Dimensions.X > 0) And (OriginalBackGround.Dimensions.Y > 0))) Then
  If (MsgBox("Confirm to delete the the current background ?", (vbQuestion + vbYesNo), "Delete background") = vbYes) Then
   BitMap2D_Delete OriginalBackGround
  Else
   Check1.Value = vbChecked
  End If
 End If

End Sub
Private Sub Combo1_Click()

 Label32.Enabled = False: Text14.Enabled = False
 Label33.Enabled = False: Text15.Enabled = False
 Label34.Enabled = False: Text16.Enabled = False
 Label35.Enabled = False: Text17.Enabled = False

 Select Case Combo1.Text

  Case "Bilinear":
   TheTexelsFilter = K3DE_XFM_BILINEAR

  Case "Bell":
   TheTexelsFilter = K3DE_XFM_BELL

  Case "Gaussian":
   TheTexelsFilter = K3DE_XFM_GAUSSIAN
   Label35.Enabled = True: Text17.Enabled = True

  Case "Cubic spline B":
   TheTexelsFilter = K3DE_XFM_CUBIC_SPLINE_B

  Case "Cubic spline BC":
   TheTexelsFilter = K3DE_XFM_CUBIC_SPLINE_BC
   Label33.Enabled = True: Text15.Enabled = True
   Label34.Enabled = True: Text16.Enabled = True

  Case "Cubic spline cardinal":
   TheTexelsFilter = K3DE_XFM_CUBIC_SPLINE_CARDINAL
   Label32.Enabled = True: Text14.Enabled = True

 End Select

End Sub
Private Sub Combo2_Click()

 Label36.Enabled = False: Text18.Enabled = False
 Label37.Enabled = False: Text19.Enabled = False
 Command6.Enabled = False
 Text18.Text = MipMapsLevel
 Text19.Text = MipMapsMinPurcent
 Combo1.Enabled = True
 Select Case TheTexelsFilter
  Case K3DE_XFM_BILINEAR:              Combo1.Text = "Bilinear"
  Case K3DE_XFM_BELL:                  Combo1.Text = "Bell"
  Case K3DE_XFM_GAUSSIAN:              Combo1.Text = "Gaussian"
  Case K3DE_XFM_CUBIC_SPLINE_B:        Combo1.Text = "Cubic spline B"
  Case K3DE_XFM_CUBIC_SPLINE_BC:       Combo1.Text = "Cubic spline BC"
  Case K3DE_XFM_CUBIC_SPLINE_CARDINAL: Combo1.Text = "Cubic spline cardinal"
 End Select

 Select Case Combo2.Text
  Case "Nearest":
   TheTexturesFilter = K3DE_TFM_NEAREST
   Label36.Enabled = False: Text18.Enabled = False
   Label37.Enabled = False: Text19.Enabled = False
   Combo1.Enabled = False

  Case "Nearest mip nearest":
   TheTexturesFilter = K3DE_TFM_NEAREST_MIP_NEAREST
   Label36.Enabled = True: Text18.Enabled = True
   Label37.Enabled = True: Text19.Enabled = True
   Command6.Enabled = True: Combo1.Enabled = False

  Case "Nearest mip linear":
   TheTexturesFilter = K3DE_TFM_NEAREST_MIP_LINEAR
   Label36.Enabled = True: Text18.Enabled = True
   Label37.Enabled = True: Text19.Enabled = True
   Command6.Enabled = True: Combo1.Enabled = False

  Case "Filtered":
   TheTexturesFilter = K3DE_TFM_FILTERED
   Label36.Enabled = False: Text18.Enabled = False
   Label37.Enabled = False: Text19.Enabled = False

  Case "Filtered mip nearest":
   TheTexturesFilter = K3DE_TFM_FILTERED_MIP_NEAREST
   Label36.Enabled = True: Text18.Enabled = True
   Label37.Enabled = True: Text19.Enabled = True
   Command6.Enabled = True

  Case "Filtered mip linear (Trilinear)":
   TheTexturesFilter = K3DE_TFM_FILTERED_MIP_LINEAR
   Label36.Enabled = True: Text18.Enabled = True
   Label37.Enabled = True: Text19.Enabled = True
   Command6.Enabled = True

 End Select

End Sub
Private Sub Command1_Click()

 If (MsgBox("Re-allocating the memory MEAN'S DESTROY all the current 3D database, and the engine will reset, continue ?", (vbQuestion + vbYesNo), "Reset") = vbYes) Then
  MaxVertices = Text1.Text
  MaxFaces = Text2.Text
  MaxMeshs = Text3.Text
  MaxLights = Text4.Text
  MaxCameras = Text5.Text
  UserAllocation = True: Engine_Reset: Unload Me
 End If

End Sub

Private Sub Command2_Click()

 FRM_Fog.Show 1

End Sub
Private Sub Command3_Click()

 BrowseType = 0
 DisplayAMap OriginalBackGround

 If ((OriginalBackGround.Dimensions.X > 0) And (OriginalBackGround.Dimensions.Y > 0)) Then Exit Sub
 If (Browsed = False) Then
  MsgBox "No background image was selected !", vbInformation, "Background"
  Command3.Enabled = False
  Check1.Value = vbUnchecked
 Else
  'Fill the viewport with background (optional)
  MousePointer = 11
  OriginalBackGround = TheOutputMap
  RefreshForm
  MousePointer = 0
 End If

End Sub
Private Sub Command4_Click()

 If ((Option4.Value = True) And (Text11.Text <= 100)) Then
  MsgBox "You must specify a valid photons count !", vbCritical, "Photons count"
  Exit Sub
 End If

 Select Case TheTexelsFilter
  Case K3DE_XFM_GAUSSIAN:              KernelSize = Text17.Text
  Case K3DE_XFM_CUBIC_SPLINE_BC:       CubicB = Text15.Text: CubicC = Text16.Text
  Case K3DE_XFM_CUBIC_SPLINE_CARDINAL: CubicA = Text14.Text
 End Select

 SamplesPerViewPath = Text7.Text: ViewPathsPerPixel = Text6.Text
 ShadowRaysCount = Text9.Text: ShadowsApproxRadius = Text8.Text
 OutputWidth = Text21.Text: OutputHeight = Text20.Text
 SamplesPerPhotonPath = Text10.Text: PhotonsSearchRadius = Text13.Text
 EstimateMultiplier = Text12.Text: BleedingDistance = Text22.Text

 Unload Me

End Sub
Private Sub Command5_Click()

 If (Text11.Text > 100) Then
  MaximumAllocatedPhotons = Text11.Text
  ReDim ThePhotonMap(MaximumAllocatedPhotons)
  RefreshForm
 Else
  If (MsgBox("You must type at least 100 photons in the photon map, use a default value ?", (vbInformation + vbYesNo), "No enough photons") = vbYes) Then
   Text11.Text = "10000"
   MaximumAllocatedPhotons = Text11.Text
   ReDim ThePhotonMap(MaximumAllocatedPhotons)
   RefreshForm
  End If
 End If

End Sub

Private Sub Command6_Click()

 If (MsgBox("Are you sure to re-create the mip-maps ?", (vbQuestion + vbYesNo), "Mip mapping") = vbYes) Then
  MousePointer = 11
  MipMapsLevel = Text18.Text
  MipMapsMinPurcent = Text19.Text
  ReCreateMipMaps
  MousePointer = 0
 Else
  Text18.Text = MipMapsLevel
  Text19.Text = MipMapsMinPurcent
 End If

End Sub

Private Sub Command7_Click()

 Dim StrMsg As String, UserChoise As String

 StrMsg = vbNewLine & "1- QCIF (176x144)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "2- CIF (352x288)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "3- QVGA (320x240)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "4- VGA (640x480)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "5- NTSC Video CD (352x240)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "6- NTSC Standard (720x486)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "7- NTSC DV (720x480)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "8- NTSC Cropped (704x480)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "9- PAL DV (720x576)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "10- PAL Square Pixel (768x576)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "11- Screen (800x600)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "12- Screen (1024x768)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "13- Screen (1152x864)" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "14- 1.3 Méga-pixel camera (1280x1024)" & vbNewLine

 UserChoise = InputBox(StrMsg, "Select a model for the output size :")

 Select Case UserChoise
  Case "1":  Text21.Text = "176": Text20.Text = "144"
  Case "2":  Text21.Text = "352": Text20.Text = "288"
  Case "3":  Text21.Text = "320": Text20.Text = "240"
  Case "4":  Text21.Text = "640": Text20.Text = "480"
  Case "5":  Text21.Text = "352": Text20.Text = "240"
  Case "6":  Text21.Text = "720": Text20.Text = "486"
  Case "7":  Text21.Text = "720": Text20.Text = "480"
  Case "8":  Text21.Text = "704": Text20.Text = "480"
  Case "9":  Text21.Text = "720": Text20.Text = "576"
  Case "10": Text21.Text = "768": Text20.Text = "576"
  Case "11": Text21.Text = "800": Text20.Text = "600"
  Case "12": Text21.Text = "1024": Text20.Text = "768"
  Case "13": Text21.Text = "1152": Text20.Text = "864"
  Case "14": Text21.Text = "1280": Text20.Text = "1024"
 End Select

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub

Private Sub Form_Load()

 RefreshForm

End Sub

Private Sub Label39_Click()

 COMDLG.ShowColor
 Label39.BackColor = COMDLG.Color
 TheAmbiantLight = ColorLongToRGB(COMDLG.Color)

End Sub

Private Sub Label43_Click()

 COMDLG.ShowColor
 Label43.BackColor = COMDLG.Color
 TheBackGroundColor = ColorLongToRGB(COMDLG.Color)

End Sub
Private Sub Option1_Click()

 Label24.Enabled = False
 Label25.Enabled = False
 Text8.Enabled = False
 Text9.Enabled = False

 EnableAreaShadows = False

End Sub
Private Sub Option2_Click()

 Label24.Enabled = True
 Label25.Enabled = True
 Text8.Enabled = True: Text8.Text = ShadowsApproxRadius
 Text9.Enabled = True: Text9.Text = ShadowRaysCount

 EnableAreaShadows = True

End Sub

Private Sub Option3_Click()

 DontAsk2 = False: If (DontAsk1 = True) Then Exit Sub

 If (MsgBox("You confirm to destroy the photon map ?", (vbQuestion + vbYesNo), "Photon mapping") = vbYes) Then

  ReDim ThePhotonMap(0): MaximumAllocatedPhotons = 0

  Label38.Enabled = True
  Label39.Enabled = True

  Label26.Enabled = False
  Label27.Enabled = False
  Label28.Enabled = False
  Label29.Enabled = False
  Label46.Enabled = False
  Text10.Enabled = False
  Text11.Enabled = False
  Text12.Enabled = False
  Text13.Enabled = False
  Text22.Enabled = False
  Command5.Enabled = False

  EnablePhotonMapping = False
  DontAsk2 = False

  RefreshForm

 Else
  DontAsk2 = True
  Option4.Value = True
 End If

End Sub
Private Sub Option4_Click()

 DontAsk1 = False: If (DontAsk2 = True) Then Exit Sub

 If (MsgBox("The photon map will use a large amout of extra memory, continue ?", (vbQuestion + vbYesNo), "Photon mapping") = vbYes) Then

  Label38.Enabled = False
  Label39.Enabled = False

  Label26.Enabled = True
  Label27.Enabled = True
  Label28.Enabled = True
  Label29.Enabled = True
  Label46.Enabled = True
  Text10.Enabled = True
  Text11.Enabled = True
  Text12.Enabled = True
  Text13.Enabled = True
  Text22.Enabled = True
  Command5.Enabled = True

  Text10.Text = SamplesPerPhotonPath
  Text11.Text = MaximumAllocatedPhotons
  Text12.Text = EstimateMultiplier
  Text13.Text = PhotonsSearchRadius
  Text22.Text = BleedingDistance

  EnablePhotonMapping = True
  DontAsk1 = False

 Else
  DontAsk1 = True
  Option3.Value = True
 End If

End Sub
