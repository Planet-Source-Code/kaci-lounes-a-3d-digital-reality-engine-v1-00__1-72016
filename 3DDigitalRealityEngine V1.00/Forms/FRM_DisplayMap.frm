VERSION 5.00
Begin VB.Form FRM_DisplayMap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Map visualization"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9105
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   536
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   607
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   7200
      Width           =   8895
      Begin VB.CommandButton Command3 
         Caption         =   "&Browse for another map..."
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Close"
         Default         =   -1  'True
         Height          =   375
         Left            =   7560
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map display : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   6255
         Left            =   8520
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   10
         Left            =   120
         TabIndex        =   3
         Top             =   6600
         Width           =   8415
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         Height          =   6255
         Left            =   120
         ScaleHeight     =   413
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   557
         TabIndex        =   1
         Top             =   360
         Width           =   8415
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2295
            Left            =   0
            ScaleHeight     =   153
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   545
            TabIndex        =   2
            Top             =   0
            Width           =   8175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "[None]"
            Height          =   195
            Left            =   3960
            TabIndex        =   8
            Top             =   3120
            Visible         =   0   'False
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "FRM_DisplayMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

 Unload Me

End Sub
Private Sub Command3_Click()

 FRM_BrowseMap.Show 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me: Unload FRM_DisplayMap

End Sub
