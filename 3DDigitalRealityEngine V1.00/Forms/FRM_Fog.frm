VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRM_Fog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fog effect"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_Fog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Enable Fog"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   5535
      Begin MSComDlg.CommonDialog COMDLG 
         Left            =   0
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
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
         Left            =   4440
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   5535
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         TabIndex        =   15
         Text            =   "0.25"
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Exponential"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   3135
         Begin VB.CommandButton Command2 
            Caption         =   "Display"
            Height          =   615
            Left            =   2160
            TabIndex        =   14
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   12
            Text            =   "0.75"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   10
            Text            =   "0.25"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Exp factor 2 : "
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   750
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Exp factor 1 : "
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   390
            Width           =   1035
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Linear"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Range : "
         Height          =   195
         Left            =   3720
         TabIndex        =   16
         Top             =   1110
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Color : "
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fog shape preview : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   120
         ScaleHeight     =   199
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   351
         TabIndex        =   1
         Top             =   360
         Width           =   5295
         Begin VB.Line Line2 
            BorderColor     =   &H00E0E0E0&
            X1              =   0
            X2              =   464
            Y1              =   96
            Y2              =   96
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            X1              =   176
            X2              =   176
            Y1              =   0
            Y2              =   200
         End
      End
   End
End
Attribute VB_Name = "FRM_Fog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub DisplayFog()

 Picture1.Cls

 Dim X!, Y!, T!

 For X = 0 To (Picture1.ScaleWidth - 1)
  T = (1 - (X / (Picture1.ScaleWidth - 1)))
  If (Option2.Value = True) Then
   T = ExpScale(T, Text1.Text, Text2.Text)
  End If
  Y = (T * (Picture1.ScaleHeight - 1))
  Picture1.PSet (X, Y), vbBlue
 Next X

End Sub
Private Sub Check1_Click()

 If (CheckOut(Check1) = False) Then
  Label3.Enabled = False
  Option1.Enabled = False
  Option2.Enabled = False
  Text1.Enabled = False
  Text2.Enabled = False
  Text3.Enabled = False
  Label1.Enabled = False
  Label2.Enabled = False
  Label4.Enabled = False
  Label5.Enabled = False
  Command2.Enabled = False
 Else
  Label3.Enabled = True
  Option1.Enabled = True
  Option2.Enabled = True
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Label1.Enabled = True
  Label2.Enabled = True
  Label4.Enabled = True
  Label5.Enabled = True
  Command2.Enabled = True
  Label3.BackColor = ColorRGBToLong(FogColor)
  If (FogMode = K3DE_FM_LINEAR) Then
   Option1.Value = True
  Else
   Option2.Value = True
  End If
  Text1.Text = FogExpFactor1
  Text2.Text = FogExpFactor2
  Text3.Text = FogRange
 End If

End Sub
Private Sub Command1_Click()

 If (CheckOut(Check1) = False) Then
  FogEnable = False
 Else
  FogEnable = True
  FogColor = ColorLongToRGB(Label3.BackColor)
  If (Option1.Value = True) Then
   FogMode = K3DE_FM_LINEAR
  Else
   FogMode = K3DE_FM_EXP
  End If
  FogExpFactor1 = Text1.Text
  FogExpFactor2 = Text2.Text
  FogRange = Text3.Text
 End If

 Unload Me

End Sub
Private Sub Command2_Click()

 DisplayFog

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub
Private Sub Form_Load()

 If (FogEnable = True) Then Check1.Value = vbChecked
 Check1_Click

End Sub
Private Sub Label3_Click()

 COMDLG.ShowColor
 Label3.BackColor = COMDLG.Color

End Sub
Private Sub Option1_Click()

 Label1.Enabled = False
 Label2.Enabled = False
 Text1.Enabled = False
 Text2.Enabled = False
 Command2.Enabled = False
 DisplayFog

End Sub
Private Sub Option2_Click()

 Label1.Enabled = True
 Label2.Enabled = True
 Text1.Enabled = True
 Text2.Enabled = True
 Command2.Enabled = True
 DisplayFog

End Sub
