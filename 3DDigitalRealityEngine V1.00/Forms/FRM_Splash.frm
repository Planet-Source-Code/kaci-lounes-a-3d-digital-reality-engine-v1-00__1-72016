VERSION 5.00
Begin VB.Form FRM_Splash 
   BorderStyle     =   0  'None
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6300
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
   Picture         =   "FRM_Splash.frx":0000
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "FRM_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Click()

 Unload Me
 FRM_Main.Show

End Sub
Private Sub Timer1_Timer()

 Unload Me
 FRM_Main.Show

End Sub
