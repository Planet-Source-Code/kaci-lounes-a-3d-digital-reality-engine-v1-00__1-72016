VERSION 5.00
Begin VB.Form FRM_SelectCamera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select a camera"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_SelectCamera.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   279
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   3975
      Begin VB.CommandButton Command1 
         Caption         =   "&Select"
         Default         =   -1  'True
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cameras list : "
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
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
   End
End
Attribute VB_Name = "FRM_SelectCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

 TheCurrentCamera = List1.ListIndex
 Unload Me

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub
Private Sub Form_Load()

 Dim CurCam As Long

 For CurCam = 0 To TheCamerasCount
  List1.AddItem TheCameras(CurCam).Label
 Next CurCam

 List1.ListIndex = TheCurrentCamera

End Sub
Private Sub List1_DblClick()

 TheCurrentCamera = List1.ListIndex
 Unload Me

End Sub
