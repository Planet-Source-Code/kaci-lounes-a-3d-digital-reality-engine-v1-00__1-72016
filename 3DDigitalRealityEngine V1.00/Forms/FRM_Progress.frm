VERSION 5.00
Begin VB.Form FRM_Progress 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4995
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
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   333
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   1980
      TabIndex        =   7
      Top             =   165
      Width           =   15
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   195
      Left            =   4440
      TabIndex        =   6
      Top             =   405
      Width           =   165
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   195
      Left            =   4425
      TabIndex        =   5
      Top             =   165
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   405
      Width           =   45
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   1980
      TabIndex        =   3
      Top             =   405
      Width           =   15
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   1980
      TabIndex        =   2
      Top             =   405
      Width           =   2295
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   1980
      TabIndex        =   1
      Top             =   165
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   255
      TabIndex        =   0
      Top             =   165
      Width           =   45
   End
End
Attribute VB_Name = "FRM_Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub DisplaySaveProgress(Elem As Integer, Purcent1!, Purcent2!)

 Caption = "Saving scene to file..."
 Label4.Caption = "Saving progress..."
 Label3.Width = (Purcent1 * Label2.Width)
 Label7.Caption = Fix(Purcent1 * 100) & " %"

 Select Case Elem
  Case 1:
   Label1.Caption = "Saving settings..."
   Label5.Width = (Purcent2 * Label6.Width)
   Label8.Caption = Fix(Purcent2 * 100) & " %"

  Case 2:
   Label1.Caption = "Saving background..."
   Label5.Width = (Purcent2 * Label6.Width)
   Label8.Caption = Fix(Purcent2 * 100) & " %"

  Case 3:
   Label1.Caption = "Saving meshs..."
   Label5.Width = (Purcent2 * Label6.Width)
   Label8.Caption = Fix(Purcent2 * 100) & " %"

  Case 4:
   Label1.Caption = "Saving omnilights..."
   Label5.Width = (Purcent2 * Label6.Width)
   Label8.Caption = Fix(Purcent2 * 100) & " %"

  Case 5:
   Label1.Caption = "Saving spotlights..."
   Label5.Width = (Purcent2 * Label6.Width)
   Label8.Caption = Fix(Purcent2 * 100) & " %"

  Case 6:
   Label1.Caption = "Saving cameras..."
   Label5.Width = (Purcent2 * Label6.Width)
   Label8.Caption = Fix(Purcent2 * 100) & " %"

 End Select

 DoEvents

End Sub
Sub DisplayLoadProgress(Elem As Integer, Purcent1!, Purcent2!)

 Caption = "Loading scene from file..."
 Label4.Caption = "Loading progress..."
 Label3.Width = (Purcent1 * Label2.Width)
 Label7.Caption = Fix(Purcent1 * 100) & " %"

 Select Case Elem
  Case 1:
   Label1.Caption = "Loading settings..."
   Label5.Width = (Purcent2 * Label6.Width)
   Label8.Caption = Fix(Purcent2 * 100) & " %"

  Case 2:
   Label1.Caption = "Loading background..."
   Label5.Width = (Purcent2 * Label6.Width)
   Label8.Caption = Fix(Purcent2 * 100) & " %"

  Case 3:
   Label1.Caption = "Loading meshs..."
   Label5.Width = (Purcent2 * Label6.Width)
   Label8.Caption = Fix(Purcent2 * 100) & " %"

  Case 4:
   Label1.Caption = "Loading omnilights..."
   Label5.Width = (Purcent2 * Label6.Width)
   Label8.Caption = Fix(Purcent2 * 100) & " %"

  Case 5:
   Label1.Caption = "Loading spotlights..."
   Label5.Width = (Purcent2 * Label6.Width)
   Label8.Caption = Fix(Purcent2 * 100) & " %"

  Case 6:
   Label1.Caption = "Loading cameras..."
   Label5.Width = (Purcent2 * Label6.Width)
   Label8.Caption = Fix(Purcent2 * 100) & " %"

 End Select

 DoEvents

End Sub
