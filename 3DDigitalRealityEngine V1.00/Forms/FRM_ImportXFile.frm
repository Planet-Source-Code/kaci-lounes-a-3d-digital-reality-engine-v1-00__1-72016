VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRM_ImportXFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import geometry from DirectX(R) file"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_ImportXFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5775
      Begin VB.CommandButton Command2 
         Caption         =   "&Load"
         Default         =   -1  'True
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DirectX(R) File name (*.x) :"
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
      Width           =   5775
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Text            =   "1"
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Browse..."
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "[None]"
         Top             =   360
         Width           =   4455
      End
      Begin MSComDlg.CommonDialog COMDLG 
         Left            =   120
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Import DirectX(R) file"
         Filter          =   "DiectX(R) file (*.X)|*.X|"
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Initial uniform scaling factor :"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   900
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FRM_ImportXFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

 COMDLG.FileName = vbNullString
 COMDLG.InitDir = App.Path & "\Datas\XFiles objects\"
 COMDLG.ShowOpen

 If (COMDLG.FileName = vbNullString) Then
  Text1.Alignment = 2
  Text1.Text = "[None]"
 Else
  Text1.Alignment = 0
  Text1.Text = COMDLG.FileName
 End If

End Sub
Private Sub Command2_Click()

 Hide
 MousePointer = 11
 DoEvents

 ImportMeshFromXFile Text1.Text, Text2.Text

 MousePointer = 0
 Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub
