VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRM_Render 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Render scene"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
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
   ScaleHeight     =   648
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   729
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Render display : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin MSComDlg.CommonDialog COMDLG 
         Left            =   3840
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   7815
         Left            =   10320
         TabIndex        =   7
         Top             =   840
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   8640
         Width           =   10215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save render..."
         Default         =   -1  'True
         Enabled         =   0   'False
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
         Left            =   8760
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   10395
         TabIndex        =   2
         Top             =   9000
         Width           =   10455
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   7740
            TabIndex        =   10
            Top             =   60
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   7920
            TabIndex        =   9
            Top             =   100
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   7920
            TabIndex        =   8
            Top             =   100
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   0
            TabIndex        =   3
            Top             =   45
            Width           =   45
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   7695
         Left            =   120
         ScaleHeight     =   7635
         ScaleWidth      =   10035
         TabIndex        =   1
         Top             =   840
         Width           =   10095
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   75
            Left            =   0
            ScaleHeight     =   5
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   2
            TabIndex        =   5
            Top             =   0
            Width           =   30
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               X1              =   0
               X2              =   100
               Y1              =   0
               Y2              =   0
            End
         End
      End
   End
End
Attribute VB_Name = "FRM_Render"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

 COMDLG.FileName = vbNullString
 COMDLG.Filter = "Windows bitmaps file (*.BMP)|*.BMP|"
 COMDLG.InitDir = App.Path & "\Renders\"
 COMDLG.ShowSave

 If (COMDLG.FileName = vbNullString) Then
  MsgBox "Saving file aborted!", vbCritical, "Abort"
 Else
  If (FileExist(COMDLG.FileName) = True) Then
   If (MsgBox("Overwrite the file ?", (vbQuestion + vbYesNo), "Overwrite") = vbYes) Then
    SavePicture Picture3.Image, COMDLG.FileName
    MsgBox "File saved.", vbInformation, "Saved"
   Else
    MsgBox "Saving file aborted!", vbCritical, "Abort"
   End If
  Else
   SavePicture Picture3.Image, COMDLG.FileName
   MsgBox "File saved.", vbInformation, "Saved"
  End If
 End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (PreviewMode = True) Then Exit Sub
 If (KeyCode = vbKeyEscape) Then
  If (Command1.Enabled = False) Then
   StopRender = True
  Else
   Unload Me
  End If
 End If

End Sub
Private Sub Form_Load()

 Show
 StopRender = False
 Previewed = False

 Dim OldOutputWidth%, OldOutputHeight%

 If (MsgBox("Take a preview first ?", (vbQuestion + vbYesNo), "Preview") = vbYes) Then
  PreviewMode = True
  If (EnablePhotonMapping = True) Then
   Label1.Caption = "Emitting photons from light sources...": MousePointer = 11
   FRM_Render.Label2.Visible = True
   FRM_Render.Label3.Visible = True
   FRM_Render.Label4.Visible = True
  End If
  OldOutputWidth = OutputWidth: OldOutputHeight = OutputHeight
  OutputWidth = 100: OutputHeight = ((OldOutputHeight / OldOutputWidth) * OutputWidth)
  Engine_Render BitMap2D_Dummy
  OutputWidth = OldOutputWidth: OutputHeight = OldOutputHeight
  Caption = "Render scene"
  PreviewMode = False: Previewed = True
  If (MsgBox("Continue the render ?", (vbQuestion + vbYesNo), "Render") = vbYes) Then
   Engine_Render BitMap2D_Dummy
  Else
   Unload Me
  End If
 Else
  If (EnablePhotonMapping = True) Then Label1.Caption = "Emitting photons from light sources...": MousePointer = 11
  Engine_Render BitMap2D_Dummy
 End If

End Sub
Private Sub HScroll1_Change()

 Picture3.Left = HScroll1.Value

End Sub
Private Sub HScroll1_Scroll()

 HScroll1_Change

End Sub
Private Sub VScroll1_Change()

 Picture3.Top = VScroll1.Value

End Sub
Private Sub VScroll1_Scroll()

 VScroll1_Change

End Sub
