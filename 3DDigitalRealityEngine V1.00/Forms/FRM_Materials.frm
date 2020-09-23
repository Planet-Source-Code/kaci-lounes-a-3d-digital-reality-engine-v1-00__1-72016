VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRM_Materials 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Materials manager"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRM_Materials.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   361
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   31
      Top             =   4560
      Width           =   5895
      Begin MSComDlg.CommonDialog COMDLG 
         Left            =   0
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command6 
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
         Left            =   4800
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Material options :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      Begin VB.Frame Frame7 
         Caption         =   "Specular hightlight :"
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
         Height          =   1095
         Left            =   3720
         TabIndex        =   11
         Top             =   360
         Width           =   2055
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1200
            TabIndex        =   28
            Text            =   "0"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1200
            TabIndex        =   27
            Text            =   "0"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Specular N"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Specular K"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Color :"
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
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   5655
         Begin VB.OptionButton Option1 
            Caption         =   "Use Color value"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Use Color map"
            Height          =   255
            Left            =   3240
            TabIndex        =   8
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Left            =   4800
            TabIndex        =   7
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            MousePointer    =   2  'Cross
            TabIndex        =   10
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Alpha map (clip mapping) :"
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
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3495
         Begin VB.CheckBox Check1 
            Caption         =   "Use Alpha map"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Browse a map..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   4
            Top             =   400
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Reflection && refraction :"
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
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   5655
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   5415
            TabIndex        =   18
            Top             =   720
            Width           =   5415
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2040
               TabIndex        =   22
               Text            =   "0"
               Top             =   0
               Width           =   615
            End
            Begin VB.CommandButton Command4 
               Caption         =   "..."
               Height          =   255
               Left            =   4800
               TabIndex        =   21
               Top             =   45
               Width           =   495
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Use Refraction map"
               Height          =   255
               Left            =   2880
               TabIndex        =   20
               Top             =   45
               Width           =   1815
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Use Refraction value"
               Height          =   255
               Left            =   0
               TabIndex        =   19
               Top             =   45
               Value           =   -1  'True
               Width           =   2055
            End
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   5415
            TabIndex        =   17
            Top             =   1080
            Width           =   5415
            Begin VB.OptionButton Option7 
               Caption         =   "Use RefractionN value"
               Height          =   255
               Left            =   0
               TabIndex        =   26
               Top             =   45
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton Option8 
               Caption         =   "Use RefractionN map"
               Height          =   255
               Left            =   2880
               TabIndex        =   25
               Top             =   45
               Width           =   1815
            End
            Begin VB.CommandButton Command5 
               Caption         =   "..."
               Height          =   255
               Left            =   4800
               TabIndex        =   24
               Top             =   45
               Width           =   495
            End
            Begin VB.TextBox Text3 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2040
               TabIndex        =   23
               Text            =   "0"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   5415
            TabIndex        =   12
            Top             =   360
            Width           =   5415
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   2040
               TabIndex        =   16
               Text            =   "0"
               Top             =   0
               Width           =   615
            End
            Begin VB.CommandButton Command3 
               Caption         =   "..."
               Height          =   255
               Left            =   4800
               TabIndex        =   15
               Top             =   45
               Width           =   495
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Use Reflection map"
               Height          =   255
               Left            =   2880
               TabIndex        =   14
               Top             =   45
               Width           =   1695
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Use Reflection value"
               Height          =   255
               Left            =   0
               TabIndex        =   13
               Top             =   45
               Value           =   -1  'True
               Width           =   1935
            End
         End
      End
   End
End
Attribute VB_Name = "FRM_Materials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub RefreshForm()

 With TheMaterials(MaterialWindowIndex)

  If (.UseAlphaTexture = False) Then Check1.Value = vbUnchecked Else Check1.Value = vbChecked

  Label1.BackColor = ColorRGBToLong(TheMaterials(MaterialWindowIndex).Color)

  If (.UseColorTexture = True) Then
   If ((TheColorTextures(MaterialWindowIndex).Dimensions.X = 0) And (TheColorTextures(MaterialWindowIndex).Dimensions.Y = 0)) Then
    .UseColorTexture = False
   End If
  End If
  If (.UseColorTexture = False) Then
   Option1.Value = True: Label1.Enabled = True: Command1.Enabled = False
  Else
   Option2.Value = True: Label1.Enabled = False: Command1.Enabled = True
  End If

  If (.UseReflectionTexture = True) Then
   If ((TheReflectionTextures(MaterialWindowIndex).Dimensions.X = 0) And (TheReflectionTextures(MaterialWindowIndex).Dimensions.Y = 0)) Then
    .UseReflectionTexture = False
   End If
  End If
  If (.UseReflectionTexture = False) Then
   Option3.Value = True: Text1.Enabled = True
   Text1.Text = TheMaterials(MaterialWindowIndex).Reflection
   Command3.Enabled = False
  Else
   Option4.Value = True: Text1.Enabled = False: Command3.Enabled = True
  End If

  If (.UseRefractionTexture = True) Then
   If ((TheRefractionTextures(MaterialWindowIndex).Dimensions.X = 0) And (TheRefractionTextures(MaterialWindowIndex).Dimensions.Y = 0)) Then
    .UseRefractionTexture = False
   End If
  End If
  If (.UseRefractionTexture = False) Then
   Option6.Value = True: Text2.Enabled = True
   Text2.Text = TheMaterials(MaterialWindowIndex).Refraction
   Command4.Enabled = False
  Else
   Option5.Value = True: Text2.Enabled = False: Command4.Enabled = True
  End If

  If (.UseRefractionNTexture = True) Then
   If ((TheRefractionNTextures(MaterialWindowIndex).Dimensions.X = 0) And (TheRefractionNTextures(MaterialWindowIndex).Dimensions.Y = 0)) Then
    .UseRefractionNTexture = False
   End If
  End If
  If (.UseRefractionNTexture = False) Then
   Option7.Value = True: Text3.Enabled = True
   Text3.Text = TheMaterials(MaterialWindowIndex).RefractionN
   Command5.Enabled = False
  Else
   Option8.Value = True: Text3.Enabled = False: Command5.Enabled = True
  End If

  Text4.Text = .SpecularPowerK: Text5.Text = .SpecularPowerN

 End With

End Sub
Private Sub Check1_Click()

 If ((TheAlphaTextures(MaterialWindowIndex).Dimensions.X > 0) And (TheAlphaTextures(MaterialWindowIndex).Dimensions.Y > 0)) Then
  If (MsgBox("Confirm to delete the the alpha-map ?", (vbQuestion + vbYesNo), "Delete the alpha-map") = vbYes) Then
   TEX_Alpha_Remove MaterialWindowIndex
   TheMaterials(MaterialWindowIndex).UseAlphaTexture = False
   Command2.Enabled = False
  Else
   Check1.Value = vbChecked: Command2.Enabled = True
  End If
 Else
  TheMaterials(MaterialWindowIndex).UseAlphaTexture = CheckOut(Check1)
  Command2.Enabled = CheckOut(Check1)
 End If

End Sub
Private Sub Command1_Click()

 Browsed = False: BrowseType = 0
 DisplayAMap TheColorTextures(MaterialWindowIndex)
 If ((TheColorTextures(MaterialWindowIndex).Dimensions.X > 0) And (TheColorTextures(MaterialWindowIndex).Dimensions.Y > 0)) Then Exit Sub

 If (Browsed = False) Then
  MsgBox "No color-map was selected !", vbInformation, "Color texture"
  Label1.Enabled = True: Command1.Enabled = False: Option1.Value = True
 Else
  MousePointer = 11
  If (TheColorUsed(MaterialWindowIndex) = False) Then TEX_Color_Add
  TEX_Color_Set MaterialWindowIndex, TheOutputMap
  MousePointer = 0
 End If

End Sub
Private Sub Command2_Click()

 Browsed = False: BrowseType = 1
 DisplayAMap TheAlphaTextures(MaterialWindowIndex)
 If ((TheAlphaTextures(MaterialWindowIndex).Dimensions.X > 0) And (TheAlphaTextures(MaterialWindowIndex).Dimensions.Y > 0)) Then Exit Sub

 If (Browsed = False) Then
  MsgBox "No alpha-map was selected !", vbInformation, "Alpha map"
  TheMaterials(MaterialWindowIndex).UseAlphaTexture = False
  Command2.Enabled = False: Check1.Value = vbUnchecked
 Else
  MousePointer = 11
  If (TheAlphaUsed(MaterialWindowIndex) = False) Then TEX_Alpha_Add
  TEX_Alpha_Set MaterialWindowIndex, TheOutputMap
  TheMaterials(MaterialWindowIndex).UseAlphaTexture = True
  MousePointer = 0
 End If

End Sub
Private Sub Command3_Click()

 Browsed = False: BrowseType = 1
 DisplayAMap TheReflectionTextures(MaterialWindowIndex)
 If ((TheReflectionTextures(MaterialWindowIndex).Dimensions.X > 0) And (TheReflectionTextures(MaterialWindowIndex).Dimensions.Y > 0)) Then Exit Sub

 If (Browsed = False) Then
  MsgBox "No reflection-map was selected !", vbInformation, "Reflection texture"
  Text1.Enabled = True: Command3.Enabled = False: Option3.Value = True
 Else
  MousePointer = 11
  If (TheReflectionUsed(MaterialWindowIndex) = False) Then TEX_Reflection_Add
  TEX_Reflection_Set MaterialWindowIndex, TheOutputMap
  MousePointer = 0
 End If

End Sub

Private Sub Command4_Click()

 Browsed = False: BrowseType = 1
 DisplayAMap TheRefractionTextures(MaterialWindowIndex)
 If ((TheRefractionTextures(MaterialWindowIndex).Dimensions.X > 0) And (TheRefractionTextures(MaterialWindowIndex).Dimensions.Y > 0)) Then Exit Sub

 If (Browsed = False) Then
  MsgBox "No refraction-map was selected !", vbInformation, "Refraction texture"
  Text2.Enabled = True: Command4.Enabled = False: Option6.Value = True
 Else
  MousePointer = 11
  If (TheRefractionUsed(MaterialWindowIndex) = False) Then TEX_Refraction_Add
  TEX_Refraction_Set MaterialWindowIndex, TheOutputMap
  MousePointer = 0
 End If

End Sub

Private Sub Command5_Click()

 Browsed = False: BrowseType = 1
 DisplayAMap TheRefractionNTextures(MaterialWindowIndex)
 If ((TheRefractionNTextures(MaterialWindowIndex).Dimensions.X > 0) And (TheRefractionNTextures(MaterialWindowIndex).Dimensions.Y > 0)) Then Exit Sub

 If (Browsed = False) Then
  MsgBox "No refractionN-map was selected !", vbInformation, "RefractionN (index) texture"
  Text3.Enabled = True: Command5.Enabled = False: Option7.Value = True
 Else
  MousePointer = 11
  If (TheRefractionNUsed(MaterialWindowIndex) = False) Then TEX_RefractionN_Add
  TEX_RefractionN_Set MaterialWindowIndex, TheOutputMap
  MousePointer = 0
 End If

End Sub
Private Sub Command6_Click()

 If (Option1.Value = True) Then
  TheMaterials(MaterialWindowIndex).Color = ColorLongToRGB(Label1.BackColor)
 Else
  If ((TheColorTextures(MaterialWindowIndex).Dimensions.X = 0) And (TheColorTextures(MaterialWindowIndex).Dimensions.Y = 0)) Then
   MsgBox "You must select a color-map !", vbCritical, "Color map"
   Exit Sub
  End If
 End If

 If (Option3.Value = True) Then
  On Error Resume Next
  TheMaterials(MaterialWindowIndex).Reflection = Text1.Text
  If (Err.Number <> 0) Then
   MsgBox "Type a positif integer value between 0 and 255 in the reflection field !", vbCritical, "Wrong value"
   Exit Sub
  End If
  If ((Text1.Text < 0) Or (Text1.Text > 255)) Then
   MsgBox "Type a positif integer value between 0 and 255 in the reflection field !", vbCritical, "Wrong value"
   Exit Sub
  End If
 Else
  If ((TheReflectionTextures(MaterialWindowIndex).Dimensions.X = 0) And (TheReflectionTextures(MaterialWindowIndex).Dimensions.Y = 0)) Then
   MsgBox "You must select a reflection-map !", vbCritical, "Reflection map"
   Exit Sub
  End If
 End If

 If (Option6.Value = True) Then
  If ((Text2.Text < 0) Or (Text2.Text > 255)) Then
   MsgBox "Type a positif integer value between 0 and 255 in the refraction field !", vbCritical, "Wrong value"
   Exit Sub
  End If
  On Error Resume Next
  TheMaterials(MaterialWindowIndex).Refraction = Text2.Text
  If (Err.Number <> 0) Then
   MsgBox "Type a positif integer value between 0 and 255 in the refraction field !", vbCritical, "Wrong value"
   Exit Sub
  End If
 Else
  If ((TheRefractionTextures(MaterialWindowIndex).Dimensions.X = 0) And (TheRefractionTextures(MaterialWindowIndex).Dimensions.Y = 0)) Then
   MsgBox "You must select a refraction-map !", vbCritical, "Refraction map"
   Exit Sub
  End If
 End If

 If (Option7.Value = True) Then
  On Error Resume Next
  TheMaterials(MaterialWindowIndex).RefractionN = Text3.Text
  If (Err.Number <> 0) Then
   MsgBox "Type a positif value in the refractionN (refraction index) field !", vbCritical, "Wrong value"
   Exit Sub
  End If
 Else
  If ((TheRefractionNTextures(MaterialWindowIndex).Dimensions.X = 0) And (TheRefractionNTextures(MaterialWindowIndex).Dimensions.Y = 0)) Then
   MsgBox "You must select a refractionN-map !", vbCritical, "RefractionN map"
   Exit Sub
  End If
 End If

 If (Check1.Value = vbChecked) Then
  If ((TheAlphaTextures(MaterialWindowIndex).Dimensions.X = 0) And (TheAlphaTextures(MaterialWindowIndex).Dimensions.Y = 0)) Then
   MsgBox "You must select a alpha-map !", vbCritical, "Alpha map"
   Exit Sub
  End If
 End If

 On Error Resume Next
 TheMaterials(MaterialWindowIndex).SpecularPowerK = Text4.Text
 If (Err.Number <> 0) Then
  MsgBox "Type a valid positif value in the specularK field !", vbCritical, "Wrong value"
  Exit Sub
 End If

 On Error Resume Next
 TheMaterials(MaterialWindowIndex).SpecularPowerN = Text5.Text
 If (Err.Number <> 0) Then
  MsgBox "Type a valid positif value in the specularN field !", vbCritical, "Wrong value"
  Exit Sub
 End If

 If (((Text1.Enabled = True) And (Text1.Text = "0")) And ((Text2.Enabled = True) And (Text2.Text = "0"))) Then
 Else
  MsgBox "For objects with some reflection and some refraction, you must to incease the view-paths count per pixel, to better approximate the result, the image must surly take's a very long time to render.", vbInformation, "Stochastic path-tracer"
 End If

 Unload Me

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then Unload Me

End Sub
Private Sub Form_Load()

 RefreshForm

End Sub

Private Sub Label1_Click()

 COMDLG.ShowColor
 Label1.BackColor = COMDLG.Color
 TheMaterials(MaterialWindowIndex).Color = ColorLongToRGB(Label1.BackColor)

End Sub
Private Sub Option1_Click()

 If ((TheColorTextures(MaterialWindowIndex).Dimensions.X > 0) And (TheColorTextures(MaterialWindowIndex).Dimensions.Y > 0)) Then
  If (MsgBox("Confirm to delete the the color-map ?", (vbQuestion + vbYesNo), "Delete the color-map") = vbYes) Then
   TEX_Color_Remove MaterialWindowIndex
   TheMaterials(MaterialWindowIndex).UseColorTexture = False
   Label1.Enabled = True
   Label1.BackColor = ColorRGBToLong(TheMaterials(MaterialWindowIndex).Color)
   Command1.Enabled = False
  Else
   Option2.Value = True
  End If
 Else
  Label1.Enabled = True
  Label1.BackColor = ColorRGBToLong(TheMaterials(MaterialWindowIndex).Color)
  Command1.Enabled = False
 End If

End Sub
Private Sub Option2_Click()

 TheMaterials(MaterialWindowIndex).UseColorTexture = True

 Label1.Enabled = False: Command1.Enabled = True

End Sub
Private Sub Option3_Click()

 If ((TheReflectionTextures(MaterialWindowIndex).Dimensions.X > 0) And (TheReflectionTextures(MaterialWindowIndex).Dimensions.Y > 0)) Then
  If (MsgBox("Confirm to delete the the reflection-map ?", (vbQuestion + vbYesNo), "Delete the reflection-map") = vbYes) Then
   TEX_Reflection_Remove MaterialWindowIndex: TheMaterials(MaterialWindowIndex).UseReflectionTexture = False
   Text1.Enabled = True: Text1.Text = TheMaterials(MaterialWindowIndex).Reflection: Command3.Enabled = False
  Else
   Option4.Value = True
  End If
 Else
  TheMaterials(MaterialWindowIndex).UseReflectionTexture = False
  Text1.Enabled = True: Text1.Text = TheMaterials(MaterialWindowIndex).Reflection: Command3.Enabled = False
 End If

End Sub
Private Sub Option4_Click()

 TheMaterials(MaterialWindowIndex).UseReflectionTexture = True

 Text1.Enabled = False: Command3.Enabled = True

End Sub
Private Sub Option5_Click()

 TheMaterials(MaterialWindowIndex).UseRefractionTexture = True

 Text2.Enabled = False: Command4.Enabled = True

End Sub
Private Sub Option6_Click()

 If ((TheRefractionTextures(MaterialWindowIndex).Dimensions.X > 0) And (TheRefractionTextures(MaterialWindowIndex).Dimensions.Y > 0)) Then
  If (MsgBox("Confirm to delete the the refraction-map ?", (vbQuestion + vbYesNo), "Delete the refraction-map") = vbYes) Then
   TEX_Refraction_Remove MaterialWindowIndex: TheMaterials(MaterialWindowIndex).UseRefractionTexture = False
   Text2.Enabled = True: Text2.Text = TheMaterials(MaterialWindowIndex).Refraction: Command4.Enabled = False
  Else
   Option5.Value = True
  End If
 Else
  TheMaterials(MaterialWindowIndex).UseRefractionTexture = False
  Text2.Enabled = True: Text2.Text = TheMaterials(MaterialWindowIndex).Refraction: Command4.Enabled = False
 End If

End Sub
Private Sub Option7_Click()

 If ((TheRefractionNTextures(MaterialWindowIndex).Dimensions.X > 0) And (TheRefractionNTextures(MaterialWindowIndex).Dimensions.Y > 0)) Then
  If (MsgBox("Confirm to delete the the refractionN-map ?", (vbQuestion + vbYesNo), "Delete the refractionN-map") = vbYes) Then
   TEX_RefractionN_Remove MaterialWindowIndex: TheMaterials(MaterialWindowIndex).UseRefractionNTexture = False
   Text3.Enabled = True: Text3.Text = TheMaterials(MaterialWindowIndex).RefractionN: Command5.Enabled = False
  Else
   Option8.Value = True
  End If
 Else
  TheMaterials(MaterialWindowIndex).UseRefractionNTexture = False
  Text3.Enabled = True: Text3.Text = TheMaterials(MaterialWindowIndex).RefractionN: Command5.Enabled = False
 End If

End Sub
Private Sub Option8_Click()

 TheMaterials(MaterialWindowIndex).UseRefractionNTexture = True

 Text3.Enabled = False: Command5.Enabled = True

End Sub
