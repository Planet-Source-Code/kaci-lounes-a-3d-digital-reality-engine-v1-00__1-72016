VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRM_BrowseMap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browsing a map (texture)"
   ClientHeight    =   9720
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
   ScaleHeight     =   648
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   607
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   120
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   561
      TabIndex        =   15
      Top             =   9840
      Width           =   8415
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   8880
      Width           =   8895
      Begin MSComDlg.CommonDialog COMDLG 
         Left            =   0
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "All supported image files"
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Ok"
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
         Left            =   7560
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Channel selection : "
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
      Left            =   120
      TabIndex        =   4
      Top             =   8160
      Width           =   8895
      Begin VB.OptionButton Option4 
         Caption         =   "Grey-scale channel"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Blue channel"
         Height          =   255
         Left            =   6480
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Geen channel"
         Height          =   255
         Left            =   4680
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Red channel"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   360
         Width           =   1335
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
      TabIndex        =   3
      Top             =   1080
      Width           =   8895
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         Height          =   6255
         Left            =   120
         ScaleHeight     =   413
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   557
         TabIndex        =   13
         Top             =   360
         Width           =   8415
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   6255
            Left            =   0
            ScaleHeight     =   417
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   561
            TabIndex        =   14
            Top             =   0
            Width           =   8415
         End
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   10
         Left            =   120
         TabIndex        =   12
         Top             =   6600
         Width           =   8415
      End
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   6255
         Left            =   8520
         TabIndex        =   11
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map file name : "
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
      Width           =   8895
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "[None]"
         Top             =   360
         Width           =   7575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Browse..."
         Height          =   375
         Left            =   7800
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "FRM_BrowseMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ImageW%, ImageH%
Sub RefreshView()

 On Error Resume Next

 Dim CurX&, CurY&, Col As ColorRGB

 Picture2.Width = ImageW
 Picture2.Height = ImageH

 If (BrowseType = 0) Then
  For CurY = 0 To ImageH
   For CurX = 0 To ImageW
    Col = ColorLongToRGB(Picture1.Point(CurX, CurY))
    TheOutputMap.Datas(0, CurX, CurY) = Col.R
    TheOutputMap.Datas(1, CurX, CurY) = Col.G
    TheOutputMap.Datas(2, CurX, CurY) = Col.B
    Picture2.PSet (CurX, CurY), RGB(TheOutputMap.Datas(0, CurX, CurY), TheOutputMap.Datas(1, CurX, CurY), TheOutputMap.Datas(2, CurX, CurY))
   Next CurX
  Next CurY
 Else
  For CurY = 0 To ImageH
   For CurX = 0 To ImageW
    Col = ColorLongToRGB(Picture1.Point(CurX, CurY))
    If (Option4.Value = True) Then
     TheOutputMap.Datas(0, CurX, CurY) = ColorFilterGreyScale(Col).R
    ElseIf (Option1.Value = True) Then
     TheOutputMap.Datas(0, CurX, CurY) = Col.R
    ElseIf (Option2.Value = True) Then
     TheOutputMap.Datas(0, CurX, CurY) = Col.G
    ElseIf (Option3.Value = True) Then
     TheOutputMap.Datas(0, CurX, CurY) = Col.B
    End If
    Picture2.PSet (CurX, CurY), RGB(TheOutputMap.Datas(0, CurX, CurY), TheOutputMap.Datas(0, CurX, CurY), TheOutputMap.Datas(0, CurX, CurY))
   Next CurX
  Next CurY
 End If

 Browsed = True

End Sub
Private Sub Command1_Click()

 COMDLG.FileName = vbNullString
 COMDLG.InitDir = App.Path & "\Datas\Textures\"
 COMDLG.ShowOpen

 If (COMDLG.FileName = vbNullString) Then
  Text1.Alignment = 2: Text1.Text = "[None]": ImageW = 0: ImageH = 0
 Else

  On Error Resume Next
  Picture1.Picture = LoadPicture(COMDLG.FileName)
  If (Err.Number <> 0) Then MsgBox "Bad image file !", vbCritical, "Bad file": Exit Sub

  ImageW = Picture1.Width: ImageH = Picture1.Height
  If ((ImageW < MinBitMapWidth) Or (ImageW > MaxBitMapWidth) Or (ImageH < MinBitMapHeight) Or (ImageH > MaxBitMapHeight)) Then
   MsgBox "Too small or too big picture ! choose another one !", vbCritical, "Bad dimensions"
   Exit Sub
  End If

  MousePointer = 11

  Text1.Alignment = 0: Text1.Text = COMDLG.FileName
  BitMap2D_Delete TheOutputMap
  Select Case BrowseType
   Case 0: BitMap2D_Create TheOutputMap, "ColorMap", 24, ImageW, ImageH, ColorBlack
   Case 1: BitMap2D_Create TheOutputMap, "AlphaMap", 8, ImageW, ImageH, ColorBlack
   Case 2: BitMap2D_Create TheOutputMap, "ReflectionMap", 8, ImageW, ImageH, ColorBlack
   Case 3: BitMap2D_Create TheOutputMap, "RefractionMap", 8, ImageW, ImageH, ColorBlack
   Case 4: BitMap2D_Create TheOutputMap, "RefractionNMap", 8, ImageW, ImageH, ColorBlack
  End Select

  RefreshView

  'Ajust scroll bars:
  If (Picture2.ScaleWidth > Picture3.ScaleWidth) Then
   HScroll1.Enabled = True: HScroll1.Max = (Picture3.ScaleWidth - Picture2.ScaleWidth)
  Else
   HScroll1.Enabled = False
  End If
  If (Picture2.ScaleHeight > Picture3.ScaleHeight) Then
   VScroll1.Enabled = True: VScroll1.Max = (Picture3.ScaleHeight - Picture2.ScaleHeight)
  Else
   VScroll1.Enabled = False
  End If

  MousePointer = 0

 End If

End Sub
Private Sub Command3_Click()

 Unload Me
 Unload FRM_DisplayMap

End Sub
Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub
Private Sub Form_Load()

 Browsed = False

 If (BrowseType = 0) Then
  Frame3.Visible = False
  Frame4.Left = 8: Frame4.Top = 536
  Height = 9345
 Else
  Frame3.Visible = True
  Frame3.Left = 8: Frame3.Top = 544
  Frame4.Left = 8: Frame4.Top = 592
  Height = 10200
  If (BrowseType = 1) Then
   Option4.Value = True
  ElseIf (BrowseType = 2) Then
   Option1.Value = True
  ElseIf (BrowseType = 3) Then
   Option2.Value = True
  ElseIf (BrowseType = 4) Then
   Option3.Value = True
  End If
 End If

End Sub
Private Sub HScroll1_Change()

 Picture2.Left = HScroll1.Value

End Sub
Private Sub HScroll1_Scroll()

 HScroll1_Change

End Sub
Private Sub Option1_Click()

 RefreshView

End Sub
Private Sub Option2_Click()

 RefreshView

End Sub
Private Sub Option3_Click()

 RefreshView

End Sub
Private Sub Option4_Click()

 RefreshView

End Sub
Private Sub VScroll1_Change()

 Picture2.Top = VScroll1.Value

End Sub
Private Sub VScroll1_Scroll()

 VScroll1_Change

End Sub
