VERSION 5.00
Begin VB.Form FRM_Pack 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Unpack"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1815
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   121
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Unpack"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FRM_Pack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub UnpackFile(TheIndex As Byte)

 Dim SFileName$, DFileName$, AByte As Byte

 Select Case TheIndex
  Case 1:
   SFileName = (App.Path & "\Package files\File1.pck")
   DFileName = (App.Path & "\Quran.exe")
  Case 2:
   SFileName = (App.Path & "\Package files\File2.pck")
   DFileName = (App.Path & "\ARef32.dll")
  Case 3:
   SFileName = (App.Path & "\Package files\File3.pck")
   DFileName = (App.Path & "\NQTRef32.dll")
  Case 4:
   SFileName = (App.Path & "\Package files\File4.pck")
   DFileName = (App.Path & "\SIRef32.dll")
  Case 5:
   SFileName = (App.Path & "\Package files\File5.pck")
   DFileName = (App.Path & "\TLRef32.dll")
 End Select

 Open SFileName For Binary As 1
  Open DFileName For Binary As 2
   Do
    Get 1, , AByte: Put 2, , CByte(255 - AByte)
   Loop Until (EOF(1) = True)
  Close 2
 Close 1

End Sub
Private Sub Command1_Click()

 Command1.Enabled = False
 Command1.Caption = "(Please wait...)"
 MousePointer = 11
 Caption = "Unpacking...(0/5)": DoEvents
 UnpackFile 1: Caption = "Unpacking...(1/5)": DoEvents
 UnpackFile 2: Caption = "Unpacking...(2/5)": DoEvents
 UnpackFile 3: Caption = "Unpacking...(3/5)": DoEvents
 UnpackFile 4: Caption = "Unpacking...(4/5)": DoEvents
 UnpackFile 5: Caption = "Unpacking...(5/5)": DoEvents
 MousePointer = 0

 Hide
 ShellExecute 0, "open", (App.Path & "\Quran.exe"), "", "", 1
 End

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If (KeyCode = vbKeyEscape) Then End

End Sub
