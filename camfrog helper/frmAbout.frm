VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Camfrog Helper"
   ClientHeight    =   1545
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   3150
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1066.386
   ScaleMode       =   0  'User
   ScaleWidth      =   2958.013
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Coded By Jack Laidlaw"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FF0000&
      Caption         =   "Camfrog Helper v1.0"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1605
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
  Unload Me
End Sub
