VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRM_Manager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scene manager"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_Manager.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   640
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Meshs : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   7215
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   240
         TabIndex        =   105
         Top             =   480
         Width           =   4935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Material..."
         Height          =   375
         Left            =   5400
         TabIndex        =   104
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Attach 2 meshs..."
         Height          =   375
         Left            =   5400
         TabIndex        =   103
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Remove mesh"
         Height          =   375
         Left            =   5400
         TabIndex        =   102
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Make a copy"
         Height          =   375
         Left            =   5400
         TabIndex        =   101
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Remove all meshs"
         Height          =   375
         Left            =   5400
         TabIndex        =   100
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Frame Frame8 
         Caption         =   "World coordinates : "
         Height          =   4455
         Left            =   3120
         TabIndex        =   41
         Top             =   3000
         Width           =   3975
         Begin VB.TextBox Text20 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3240
            TabIndex        =   78
            Text            =   "0.1"
            Top             =   3600
            Width           =   495
         End
         Begin VB.CommandButton Command24 
            Caption         =   "-"
            Height          =   255
            Left            =   2520
            TabIndex        =   77
            Top             =   3600
            Width           =   255
         End
         Begin VB.CommandButton Command23 
            Caption         =   "+"
            Height          =   255
            Left            =   2160
            TabIndex        =   76
            Top             =   3600
            Width           =   255
         End
         Begin VB.TextBox Text19 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3240
            TabIndex        =   75
            Text            =   "0.1"
            Top             =   3240
            Width           =   495
         End
         Begin VB.CommandButton Command22 
            Caption         =   "-"
            Height          =   255
            Left            =   2520
            TabIndex        =   74
            Top             =   3240
            Width           =   255
         End
         Begin VB.CommandButton Command21 
            Caption         =   "+"
            Height          =   255
            Left            =   2160
            TabIndex        =   73
            Top             =   3240
            Width           =   255
         End
         Begin VB.TextBox Text18 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3240
            TabIndex        =   72
            Text            =   "0.1"
            Top             =   2880
            Width           =   495
         End
         Begin VB.CommandButton Command20 
            Caption         =   "-"
            Height          =   255
            Left            =   2520
            TabIndex        =   71
            Top             =   2880
            Width           =   255
         End
         Begin VB.CommandButton Command19 
            Caption         =   "+"
            Height          =   255
            Left            =   2160
            TabIndex        =   70
            Top             =   2880
            Width           =   255
         End
         Begin VB.TextBox Text17 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3240
            TabIndex        =   69
            Text            =   "5"
            Top             =   2400
            Width           =   495
         End
         Begin VB.CommandButton Command18 
            Caption         =   "-"
            Height          =   255
            Left            =   2520
            TabIndex        =   68
            Top             =   2400
            Width           =   255
         End
         Begin VB.CommandButton Command17 
            Caption         =   "+"
            Height          =   255
            Left            =   2160
            TabIndex        =   67
            Top             =   2400
            Width           =   255
         End
         Begin VB.TextBox Text16 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3240
            TabIndex        =   66
            Text            =   "5"
            Top             =   2040
            Width           =   495
         End
         Begin VB.CommandButton Command16 
            Caption         =   "-"
            Height          =   255
            Left            =   2520
            TabIndex        =   65
            Top             =   2040
            Width           =   255
         End
         Begin VB.CommandButton Command15 
            Caption         =   "+"
            Height          =   255
            Left            =   2160
            TabIndex        =   64
            Top             =   2040
            Width           =   255
         End
         Begin VB.TextBox Text15 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3240
            TabIndex        =   63
            Text            =   "5"
            Top             =   1680
            Width           =   495
         End
         Begin VB.CommandButton Command14 
            Caption         =   "-"
            Height          =   255
            Left            =   2520
            TabIndex        =   62
            Top             =   1680
            Width           =   255
         End
         Begin VB.CommandButton Command13 
            Caption         =   "+"
            Height          =   255
            Left            =   2160
            TabIndex        =   61
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox Text14 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3240
            TabIndex        =   60
            Text            =   "10"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3240
            TabIndex        =   59
            Text            =   "10"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3240
            TabIndex        =   58
            Text            =   "10"
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command12 
            Caption         =   "-"
            Height          =   255
            Left            =   2520
            TabIndex        =   57
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton Command11 
            Caption         =   "+"
            Height          =   255
            Left            =   2160
            TabIndex        =   56
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton Command10 
            Caption         =   "-"
            Height          =   255
            Left            =   2520
            TabIndex        =   55
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton Command9 
            Caption         =   "+"
            Height          =   255
            Left            =   2160
            TabIndex        =   54
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton Command8 
            Caption         =   "-"
            Height          =   255
            Left            =   2520
            TabIndex        =   53
            Top             =   480
            Width           =   255
         End
         Begin VB.CommandButton Command7 
            Caption         =   "+"
            Height          =   255
            Left            =   2160
            TabIndex        =   52
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   960
            TabIndex        =   51
            Text            =   "0"
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   960
            TabIndex        =   50
            Text            =   "0"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   960
            TabIndex        =   49
            Text            =   "0"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   960
            TabIndex        =   48
            Text            =   "0"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   960
            TabIndex        =   47
            Text            =   "0"
            Top             =   2040
            Width           =   1095
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   960
            TabIndex        =   46
            Text            =   "0"
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   960
            TabIndex        =   45
            Text            =   "0"
            Top             =   2880
            Width           =   1095
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   960
            TabIndex        =   44
            Text            =   "0"
            Top             =   3240
            Width           =   1095
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   960
            TabIndex        =   43
            Text            =   "0"
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Uniform scaling"
            Height          =   255
            Left            =   960
            TabIndex        =   42
            Top             =   4080
            Width           =   1575
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Step "
            Height          =   255
            Left            =   2640
            TabIndex        =   99
            Top             =   3600
            Width           =   615
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Step "
            Height          =   255
            Left            =   2640
            TabIndex        =   98
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Step "
            Height          =   255
            Left            =   2640
            TabIndex        =   97
            Top             =   2880
            Width           =   615
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Step "
            Height          =   255
            Left            =   2640
            TabIndex        =   96
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Step "
            Height          =   195
            Left            =   2880
            TabIndex        =   95
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Step "
            Height          =   255
            Left            =   2640
            TabIndex        =   94
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Step "
            Height          =   255
            Left            =   2640
            TabIndex        =   93
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Step "
            Height          =   255
            Left            =   2640
            TabIndex        =   92
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Step "
            Height          =   255
            Left            =   2640
            TabIndex        =   91
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Origin X :"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Origin Y :"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Origin Z :"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Angle X :"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Angle Y :"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Angle Z :"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Scale Y :"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Scale X :"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Scale Z :"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   3600
            Width           =   735
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "°"
            Height          =   195
            Left            =   3780
            TabIndex        =   81
            Top             =   1680
            Width           =   75
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "°"
            Height          =   195
            Left            =   3780
            TabIndex        =   80
            Top             =   2040
            Width           =   75
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "°"
            Height          =   195
            Left            =   3780
            TabIndex        =   79
            Top             =   2400
            Width           =   75
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Parameters : "
         Height          =   4455
         Left            =   120
         TabIndex        =   29
         Top             =   3000
         Width           =   2895
         Begin VB.CommandButton Command27 
            Caption         =   "..."
            Height          =   255
            Left            =   2450
            TabIndex        =   35
            Top             =   840
            Width           =   300
         End
         Begin VB.TextBox Text4 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   720
            TabIndex        =   34
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Visible"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Width           =   735
         End
         Begin VB.CheckBox Check3 
            Caption         =   "User-defined object-matrix"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton Command43 
            Caption         =   "Update"
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   3960
            Width           =   2415
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Export to file..."
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   2520
            Width           =   2415
         End
         Begin MSComDlg.CommonDialog COMDLG 
            Left            =   2280
            Top             =   2040
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Label :"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Vertices count : "
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   1680
            Width           =   1170
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Faces count : "
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   2040
            Width           =   1020
         End
         Begin VB.Label Label25 
            Caption         =   "0"
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
            Height          =   255
            Left            =   1440
            TabIndex        =   37
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "0"
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
            Height          =   255
            Left            =   1440
            TabIndex        =   36
            Top             =   2040
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   8760
      Width           =   7215
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
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
         Left            =   5880
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cameras : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   7215
      Begin VB.CommandButton Command32 
         Caption         =   "Remove all cameras"
         Height          =   375
         Left            =   5400
         TabIndex        =   15
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Make a copy"
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command30 
         Caption         =   "Remove camera"
         Height          =   375
         Left            =   5400
         TabIndex        =   13
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Modify camera..."
         Height          =   375
         Left            =   5400
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Add camera..."
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox List2 
         Height          =   2400
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   4935
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Spot lights : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   7215
      Begin VB.CommandButton Command37 
         Caption         =   "Remove all spotlights"
         Height          =   495
         Left            =   5400
         TabIndex        =   21
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command36 
         Caption         =   "Make a copy"
         Height          =   375
         Left            =   5400
         TabIndex        =   20
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command35 
         Caption         =   "Remove spotlight"
         Height          =   375
         Left            =   5400
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command34 
         Caption         =   "Modify spotlight..."
         Height          =   375
         Left            =   5400
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command33 
         Caption         =   "Add spotlight..."
         Height          =   375
         Left            =   5400
         TabIndex        =   17
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox List3 
         Height          =   2400
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   4935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Omni lights : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   7215
      Begin VB.CommandButton Command42 
         Caption         =   "Remove all omnilights"
         Height          =   495
         Left            =   5400
         TabIndex        =   27
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Make a copy"
         Height          =   375
         Left            =   5400
         TabIndex        =   26
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command40 
         Caption         =   "Remove omnilight"
         Height          =   375
         Left            =   5400
         TabIndex        =   25
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Modify omnilight..."
         Height          =   375
         Left            =   5400
         TabIndex        =   24
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command38 
         Caption         =   "Add omnilight..."
         Height          =   375
         Left            =   5400
         TabIndex        =   23
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox List4 
         Height          =   2400
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   4935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scene elements : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.OptionButton Option4 
         Caption         =   "Cameras"
         Height          =   255
         Left            =   5160
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Spot lights"
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Omni lights"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Geometry"
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FRM_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AttachMode As Boolean, AttachMesh1&
Sub RefreshForm()

 Dim CurElem&

 List1.Clear: List2.Clear: List3.Clear: List4.Clear

 If (Option1.Value = True) Then
  Frame2.Visible = True:  Frame3.Visible = False
  Frame4.Visible = False: Frame5.Visible = False
  Frame6.Top = 584: Height = 10080
  If (TheMeshsCount = -1) Then
   SetMeshParamsEnable False
  Else
   SetMeshParamsEnable True: If (TheMeshsCount = 0) Then Command3.Enabled = False
   If (TheMeshsCount <> -1) Then
    For CurElem = 0 To TheMeshsCount
     List1.AddItem TheMeshs(CurElem).Label
    Next CurElem
    If (TheMeshsCount > -1) Then List1.ListIndex = 0
   End If
  End If

 ElseIf (Option2.Value = True) Then
  Frame2.Visible = False: Frame3.Visible = True
  Frame4.Visible = False: Frame5.Visible = False
  Frame6.Top = 287: Height = 5640
  If (TheSphereLightsCount = -1) Then
   Command38.Enabled = True:  Command39.Enabled = False
   Command40.Enabled = False: Command41.Enabled = False
   Command42.Enabled = False
  Else
   Command38.Enabled = True: Command39.Enabled = True
   Command40.Enabled = True: Command41.Enabled = True
   Command42.Enabled = True
   For CurElem = 0 To TheSphereLightsCount
    List4.AddItem TheSphereLights(CurElem).Label
   Next CurElem
   List4.ListIndex = 0
  End If

 ElseIf (Option3.Value = True) Then
  Frame2.Visible = False: Frame3.Visible = False
  Frame4.Visible = True:  Frame5.Visible = False
  Frame6.Top = 287: Height = 5640
  If (TheConeLightsCount = -1) Then
   Command33.Enabled = True:  Command34.Enabled = False
   Command35.Enabled = False: Command36.Enabled = False
   Command37.Enabled = False
  Else
   Command33.Enabled = True: Command34.Enabled = True
   Command35.Enabled = True: Command36.Enabled = True
   Command37.Enabled = True
   For CurElem = 0 To TheConeLightsCount
    List3.AddItem TheConeLights(CurElem).Label
   Next CurElem
   List3.ListIndex = 0
  End If

 ElseIf (Option4.Value = True) Then
  Frame2.Visible = False: Frame3.Visible = False
  Frame4.Visible = False: Frame5.Visible = True
  Frame6.Top = 287: Height = 5640
  If (TheCamerasCount = -1) Then
   Command28.Enabled = True:  Command29.Enabled = False
   Command30.Enabled = False: Command31.Enabled = False
   Command32.Enabled = False
  Else
   Command28.Enabled = True: Command29.Enabled = True
   Command30.Enabled = True: Command31.Enabled = True
   Command32.Enabled = True
   For CurElem = 0 To TheCamerasCount
    List2.AddItem TheCameras(CurElem).Label
   Next CurElem
   List2.ListIndex = TheCurrentCamera
  End If

 End If

End Sub
Sub SetMeshParamsEnable(AValue As Boolean)

 Frame8.Enabled = AValue: Frame9.Enabled = AValue
 Check1.Enabled = AValue: Check2.Enabled = AValue: Check3.Enabled = AValue

 Label1.Enabled = AValue: Label2.Enabled = AValue
 Label3.Enabled = AValue: Label4.Enabled = AValue
 Label5.Enabled = AValue: Label6.Enabled = AValue
 Label7.Enabled = AValue: Label8.Enabled = AValue
 Label9.Enabled = AValue: Label10.Enabled = AValue
 Label11.Enabled = AValue: Label12.Enabled = AValue
 Label13.Enabled = AValue: Label14.Enabled = AValue
 Label15.Enabled = AValue: Label16.Enabled = AValue
 Label17.Enabled = AValue: Label18.Enabled = AValue
 Label19.Enabled = AValue: Label20.Enabled = AValue
 Label21.Enabled = AValue: Label22.Enabled = AValue
 Label23.Enabled = AValue: Label24.Enabled = AValue
 Label25.Enabled = AValue: Label26.Enabled = AValue

 Text1.Enabled = AValue: Text2.Enabled = AValue
 Text3.Enabled = AValue: Text4.Enabled = AValue
 Text5.Enabled = AValue: Text6.Enabled = AValue
 Text7.Enabled = AValue: Text8.Enabled = AValue
 Text9.Enabled = AValue: Text10.Enabled = AValue
 Text11.Enabled = AValue: Text13.Enabled = AValue
 Text14.Enabled = AValue: Text15.Enabled = AValue
 Text16.Enabled = AValue: Text17.Enabled = AValue
 Text18.Enabled = AValue: Text19.Enabled = AValue
 Text20.Enabled = AValue

 Command7.Enabled = AValue: Command8.Enabled = AValue
 Command9.Enabled = AValue: Command10.Enabled = AValue
 Command11.Enabled = AValue: Command12.Enabled = AValue
 Command13.Enabled = AValue: Command14.Enabled = AValue
 Command15.Enabled = AValue: Command16.Enabled = AValue
 Command17.Enabled = AValue: Command18.Enabled = AValue
 Command19.Enabled = AValue: Command20.Enabled = AValue
 Command21.Enabled = AValue: Command22.Enabled = AValue
 Command23.Enabled = AValue: Command24.Enabled = AValue
 Command27.Enabled = AValue: Command43.Enabled = AValue
 Command2.Enabled = AValue: Command3.Enabled = AValue
 Command4.Enabled = AValue: Command5.Enabled = AValue
 Command6.Enabled = AValue: Command25.Enabled = AValue

End Sub
Sub ViewMeshParameters(TheMeshIndex As Long)

 Text4.Text = TheMeshs(TheMeshIndex).Label

 Label25.Caption = TheMeshs(TheMeshIndex).Vertices.Length
 Label26.Caption = TheMeshs(TheMeshIndex).Faces.Length

 If (TheMeshs(TheMeshIndex).MakeMatrix = False) Then
  Check3.Value = vbChecked
  Command27.Enabled = True
 Else
  Check3.Value = vbUnchecked
  Command27.Enabled = False
 End If

 If (TheMeshs(TheMeshIndex).Visible = True) Then
  Check1.Value = vbChecked
 Else
  Check1.Value = vbUnchecked
 End If

 Text5.Text = TheMeshs(TheMeshIndex).Position.X
 Text2.Text = TheMeshs(TheMeshIndex).Position.Y
 Text3.Text = TheMeshs(TheMeshIndex).Position.Z

 '(in degrees)
 Text9.Text = TheMeshs(TheMeshIndex).Angles.X
 Text10.Text = TheMeshs(TheMeshIndex).Angles.Y
 Text11.Text = TheMeshs(TheMeshIndex).Angles.Z

 Text6.Text = TheMeshs(TheMeshIndex).Scales.X
 Text7.Text = TheMeshs(TheMeshIndex).Scales.Y
 Text8.Text = TheMeshs(TheMeshIndex).Scales.Z

End Sub

Private Sub Check1_Click()

 If (Check1.Value = vbChecked) Then
  TheMeshs(List1.ListIndex).Visible = True
  FRM_Main.RefreshViews
 Else
  TheMeshs(List1.ListIndex).Visible = False
 End If

 FRM_Main.RefreshViews

End Sub
Private Sub Check3_Click()

 If (Check3.Value = vbChecked) Then
  TheMeshs(List1.ListIndex).MakeMatrix = False
  Command27.Enabled = True
 Else
  TheMeshs(List1.ListIndex).MakeMatrix = True
  Command27.Enabled = False
 End If

 FRM_Main.RefreshViews

End Sub

Private Sub Command1_Click()

 Unload Me

End Sub
Private Sub Command10_Click()

 Text2.Text = CSng(Text2.Text) - Text13.Text
 TheMeshs(List1.ListIndex).Position.Y = Text2.Text

 FRM_Main.RefreshViews

End Sub
Private Sub Command11_Click()

 Text3.Text = CSng(Text3.Text) + Text14.Text
 TheMeshs(List1.ListIndex).Position.Z = Text3.Text

 FRM_Main.RefreshViews

End Sub
Private Sub Command12_Click()

 Text3.Text = CSng(Text3.Text) - Text14.Text
 TheMeshs(List1.ListIndex).Position.Z = Text3.Text

 FRM_Main.RefreshViews

End Sub

Private Sub Command13_Click()

 Text9.Text = CSng(Text9.Text) + (Text15.Text * Deg)
 TheMeshs(List1.ListIndex).Angles.X = Text9.Text

 FRM_Main.RefreshViews

End Sub

Private Sub Command14_Click()

 Text9.Text = CSng(Text9.Text) - (Text15.Text * Deg)
 TheMeshs(List1.ListIndex).Angles.X = Text9.Text

 FRM_Main.RefreshViews

End Sub

Private Sub Command15_Click()

 Text10.Text = CSng(Text10.Text) + (Text16.Text * Deg)
 TheMeshs(List1.ListIndex).Angles.Y = Text10.Text

 FRM_Main.RefreshViews

End Sub

Private Sub Command16_Click()

 Text10.Text = CSng(Text10.Text) - (Text16.Text * Deg)
 TheMeshs(List1.ListIndex).Angles.Y = Text10.Text

 FRM_Main.RefreshViews

End Sub

Private Sub Command17_Click()

 Text11.Text = CSng(Text11.Text) + (Text17.Text * Deg)
 TheMeshs(List1.ListIndex).Angles.Z = Text11.Text

 FRM_Main.RefreshViews

End Sub

Private Sub Command18_Click()

 Text11.Text = CSng(Text11.Text) - (Text17.Text * Deg)
 TheMeshs(List1.ListIndex).Angles.Z = Text11.Text

 FRM_Main.RefreshViews

End Sub

Private Sub Command19_Click()

 Text6.Text = CSng(Text6.Text) + Text18.Text
 TheMeshs(List1.ListIndex).Scales.X = Text6.Text

 If (Check2.Value = vbChecked) Then
  Text7.Text = CSng(Text7.Text) + Text19.Text
  TheMeshs(List1.ListIndex).Scales.Y = Text7.Text
  Text8.Text = CSng(Text8.Text) + Text20.Text
  TheMeshs(List1.ListIndex).Scales.Z = Text8.Text
 End If

 FRM_Main.RefreshViews

End Sub

Private Sub Command2_Click()

 MaterialWindowIndex = List1.ListIndex
 FRM_Materials.Show 1

 FRM_Main.RefreshViews

End Sub
Private Sub Command20_Click()

 Text6.Text = CSng(Text6.Text) - Text18.Text
 TheMeshs(List1.ListIndex).Scales.X = Text6.Text

 If (Check2.Value = vbChecked) Then
  Text7.Text = CSng(Text7.Text) - Text19.Text
  TheMeshs(List1.ListIndex).Scales.Y = Text7.Text
  Text8.Text = CSng(Text8.Text) - Text20.Text
  TheMeshs(List1.ListIndex).Scales.Z = Text8.Text
 End If

 FRM_Main.RefreshViews

End Sub

Private Sub Command21_Click()

 Text7.Text = CSng(Text7.Text) + Text19.Text
 TheMeshs(List1.ListIndex).Scales.Y = Text7.Text

 If (Check2.Value = vbChecked) Then
  Text6.Text = CSng(Text6.Text) + Text18.Text
  TheMeshs(List1.ListIndex).Scales.X = Text6.Text
  Text8.Text = CSng(Text8.Text) + Text20.Text
  TheMeshs(List1.ListIndex).Scales.Z = Text8.Text
 End If

 FRM_Main.RefreshViews

End Sub

Private Sub Command22_Click()

 Text7.Text = CSng(Text7.Text) - Text19.Text
 TheMeshs(List1.ListIndex).Scales.Y = Text7.Text

 If (Check2.Value = vbChecked) Then
  Text6.Text = CSng(Text6.Text) - Text18.Text
  TheMeshs(List1.ListIndex).Scales.X = Text6.Text
  Text8.Text = CSng(Text8.Text) - Text20.Text
  TheMeshs(List1.ListIndex).Scales.Z = Text8.Text
 End If

 FRM_Main.RefreshViews

End Sub
Private Sub Command23_Click()

 Text8.Text = CSng(Text8.Text) + Text20.Text
 TheMeshs(List1.ListIndex).Scales.Z = Text8.Text

 If (Check2.Value = vbChecked) Then
  Text6.Text = CSng(Text6.Text) + Text18.Text
  TheMeshs(List1.ListIndex).Scales.X = Text6.Text
  Text7.Text = CSng(Text7.Text) + Text19.Text
  TheMeshs(List1.ListIndex).Scales.Y = Text7.Text
 End If

 FRM_Main.RefreshViews

End Sub

Private Sub Command24_Click()

 Text8.Text = CSng(Text8.Text) - Text20.Text
 TheMeshs(List1.ListIndex).Scales.Z = Text8.Text

 If (Check2.Value = vbChecked) Then
  Text6.Text = CSng(Text6.Text) - Text18.Text
  TheMeshs(List1.ListIndex).Scales.X = Text6.Text
  Text7.Text = CSng(Text7.Text) - Text19.Text
  TheMeshs(List1.ListIndex).Scales.Y = Text7.Text
 End If

 FRM_Main.RefreshViews

End Sub

Private Sub Command25_Click()

 COMDLG.FileName = vbNullString
 COMDLG.Filter = "3D Digital Reality Engine objects files (*" & ObjectFileExtension & ")|*" & ObjectFileExtension & "|"
 COMDLG.InitDir = App.Path & "\Datas\Objects\"
 COMDLG.ShowSave

 If (COMDLG.FileName <> vbNullString) Then Engine_SaveMesh List1.ListIndex, COMDLG.FileName

End Sub
Private Sub Command27_Click()

 MatWindowMode = False
 MatWindowIndex = List1.ListIndex

 FRM_EditMatrix.Show 1
 FRM_Main.RefreshViews

End Sub
Private Sub Command28_Click()

 CamWindowMode = False
 FRM_AddCamera.Show 1
 RefreshForm
 FRM_Main.RefreshViews

End Sub
Private Sub Command29_Click()

 CamWindowMode = True
 CamWindowIndex = List2.ListIndex
 FRM_AddCamera.Show 1
 RefreshForm
 FRM_Main.RefreshViews

End Sub
Private Sub Command3_Click()

 If (MsgBox("Only the geometry is concerned in the attachement proccess, the textures and materials are ignored, continue ?", (vbInformation + vbYesNo), "Attach") = vbYes) Then
  AttachMesh1 = List1.ListIndex
  MsgBox "Please select the second mesh", vbInformation, "Attach"
  Command4.Enabled = False: Command6.Enabled = False
  AttachMode = True
 End If

End Sub
Private Sub Command30_Click()

 If (TheCamerasCount = 0) Then
  If (MsgBox("Confim to remove the LAST & CURRENT camera ?", (vbQuestion + vbYesNo), "Remove") = vbYes) Then
   Camera3D_Remove List2.ListIndex
   MsgBox "Sorry, but the system must create a default camera !", vbInformation, "Camera"
   TheCurrentCamera = Camera3D_Add  'Add default camera
   TheCameras(TheCurrentCamera).Label = "Default camera"
   RefreshForm
   FRM_Main.RefreshViews
  End If
 ElseIf (TheCamerasCount > 0) Then
  If (List2.ListIndex = TheCurrentCamera) Then
   If (MsgBox("The selected one is the current camera used to view the scene, deleting this one mean's that you have to select another camera, confirm ?", (vbQuestion + vbYesNo), "Remove") = vbYes) Then
    Camera3D_Remove List2.ListIndex
    FRM_SelectCamera.Show 1
    RefreshForm
    FRM_Main.RefreshViews
   End If
  Else
   If (MsgBox("Confim to remove the camera ?", (vbQuestion + vbYesNo), "Remove") = vbYes) Then
    Camera3D_Remove List2.ListIndex
    RefreshForm
    FRM_Main.RefreshViews
   End If
  End If
 End If

End Sub
Private Sub Command31_Click()

 If (MsgBox("Confim ?", (vbQuestion + vbYesNo), "Copy") = vbYes) Then
  Camera3D_Copy List2.ListIndex
  RefreshForm
  FRM_Main.RefreshViews
 End If

End Sub

Private Sub Command32_Click()

 If (MsgBox("Are you sure to delete all the cameras in this scene ?", (vbQuestion + vbYesNo), "Remove all") = vbYes) Then
  Camera3D_Clear
  MsgBox "Sorry, but the system must create a default camera !", vbInformation, "Camera"
  TheCurrentCamera = Camera3D_Add  'Add default camera
  TheCameras(TheCurrentCamera).Label = "Default camera"
  RefreshForm
  FRM_Main.RefreshViews
 End If

End Sub
Private Sub Command33_Click()

 SpotWindowMode = False
 FRM_Lights_Spot.Show 1
 RefreshForm
 FRM_Main.RefreshViews

End Sub

Private Sub Command34_Click()

 SpotWindowMode = True
 SpotWindowIndex = List3.ListIndex
 FRM_Lights_Spot.Show 1
 RefreshForm
 FRM_Main.RefreshViews

End Sub

Private Sub Command35_Click()

 If (MsgBox("Confim to remove the spot-light source ?", (vbQuestion + vbYesNo), "Remove") = vbYes) Then
  ConeLight3D_Remove List3.ListIndex
  RefreshForm
  FRM_Main.RefreshViews
 End If

End Sub

Private Sub Command36_Click()

 If (MsgBox("Confim ?", (vbQuestion + vbYesNo), "Copy") = vbYes) Then
  ConeLight3D_Copy List3.ListIndex
  RefreshForm
  FRM_Main.RefreshViews
 End If

End Sub

Private Sub Command37_Click()

 If (MsgBox("Are you sure to delete all the spot-lights in this scene ?", (vbQuestion + vbYesNo), "Remove all") = vbYes) Then
  ConeLight3D_Clear
  RefreshForm
  FRM_Main.RefreshViews
 End If

End Sub
Private Sub Command38_Click()

 OmniWindowMode = False
 FRM_Lights_Omni.Show 1
 RefreshForm
 FRM_Main.RefreshViews

End Sub

Private Sub Command39_Click()

 OmniWindowMode = True
 OmniWindowIndex = List4.ListIndex
 FRM_Lights_Omni.Show 1
 RefreshForm
 FRM_Main.RefreshViews

End Sub
Private Sub Command4_Click()

 If (MsgBox("Confim to remove the mesh ?", (vbQuestion + vbYesNo), "Remove") = vbYes) Then
  Mesh3D_Remove List1.ListIndex
  RefreshForm
  FRM_Main.RefreshViews
 End If

End Sub
Private Sub Command40_Click()

 If (MsgBox("Confim to remove the omni-light source ?", (vbQuestion + vbYesNo), "Remove") = vbYes) Then
  SphereLight3D_Remove List4.ListIndex
  RefreshForm
  FRM_Main.RefreshViews
 End If

End Sub
Private Sub Command41_Click()

 If (MsgBox("Confim ?", (vbQuestion + vbYesNo), "Copy") = vbYes) Then
  SphereLight3D_Copy List4.ListIndex
  RefreshForm
  FRM_Main.RefreshViews
 End If

End Sub
Private Sub Command42_Click()

 If (MsgBox("Are you sure to delete all the omni-lights in this scene ?", (vbQuestion + vbYesNo), "Remove all") = vbYes) Then
  SphereLight3D_Clear
  RefreshForm
  FRM_Main.RefreshViews
 End If

End Sub
Private Sub Command43_Click()

 TheMeshs(List1.ListIndex).Label = Text4.Text

End Sub

Private Sub Command5_Click()

 If (MsgBox("Confim ?", (vbQuestion + vbYesNo), "Copy") = vbYes) Then
  Mesh3D_Copy List1.ListIndex
  'Engine_SaveMesh List1.ListIndex, (App.Path & "\Datas\Objects\Tmpobject" & ObjectFileExtension)
  'Engine_LoadMesh (App.Path & "\Datas\Objects\Tmpobject" & ObjectFileExtension)
  'Kill (App.Path & "\Datas\Objects\Tmpobject" & ObjectFileExtension)
  RefreshForm
  FRM_Main.RefreshViews
 End If

End Sub
Private Sub Command6_Click()

 If (MsgBox("Are you sure to delete all the geometery in this scene ?", (vbQuestion + vbYesNo), "Remove all") = vbYes) Then
  Mesh3D_Clear
  RefreshForm
  FRM_Main.RefreshViews
 End If

End Sub
Private Sub Command7_Click()

 Text5.Text = CSng(Text5.Text) + Text1.Text
 TheMeshs(List1.ListIndex).Position.X = Text5.Text

 FRM_Main.RefreshViews

End Sub
Private Sub Command8_Click()

 Text5.Text = CSng(Text5.Text) - Text1.Text
 TheMeshs(List1.ListIndex).Position.X = Text5.Text

 FRM_Main.RefreshViews

End Sub
Private Sub Command9_Click()

 Text2.Text = CSng(Text2.Text) + Text13.Text
 TheMeshs(List1.ListIndex).Position.Y = Text2.Text

 FRM_Main.RefreshViews

End Sub
Private Sub Form_Activate()

 RefreshForm

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub

Private Sub List1_Click()

 If (AttachMode = True) Then
  If (List1.ListIndex = AttachMesh1) Then
   If (MsgBox("Select another mesh (no), or just abort (yes) ?", (vbQuestion + vbYesNo), "Attach") = vbYes) Then
    AttachMode = False
    Command4.Enabled = True: Command6.Enabled = True
   End If
  Else
   If (MsgBox("Attaching mesh N° " & AttachMesh1 & " , with mesh N° " & List1.ListIndex & " , Confirm ?", (vbQuestion + vbYesNo), "Attach") = vbYes) Then
    Mesh3D_Attach AttachMesh1, List1.ListIndex, True
    Command4.Enabled = True: Command6.Enabled = True
    AttachMode = False
    RefreshForm
   Else
    MsgBox "Attachement proccess is aborted.", vbCritical, "Abort"
    Command4.Enabled = True: Command6.Enabled = True
    AttachMode = False
   End If
  End If
 Else
  ViewMeshParameters List1.ListIndex
  Command4.Enabled = True: Command6.Enabled = True
  FRM_Main.RefreshViews
 End If

End Sub
Private Sub List1_DblClick()

 If (AttachMode = False) Then Command2_Click

End Sub

Private Sub List2_DblClick()

 Command29_Click

End Sub

Private Sub List3_DblClick()

 Command34_Click

End Sub

Private Sub List4_DblClick()

 Command39_Click

End Sub
Private Sub Option1_Click()

 RefreshForm

End Sub

Private Sub Option2_Click()

 RefreshForm

End Sub

Private Sub Option3_Click()

 RefreshForm

End Sub

Private Sub Option4_Click()

 RefreshForm

End Sub
