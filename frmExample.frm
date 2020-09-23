VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmExample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sunero Four Colour Gradient Test"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExTwo 
      Caption         =   "Second Example"
      Height          =   375
      Left            =   3060
      TabIndex        =   6
      Top             =   7140
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cDLG 
      Left            =   6660
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picHandleD 
      BackColor       =   &H00FF00FF&
      Height          =   375
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   6900
      Width           =   375
   End
   Begin VB.PictureBox picHandleC 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   10140
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   6900
      Width           =   375
   End
   Begin VB.PictureBox picHandleB 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   10140
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   60
      Width           =   375
   End
   Begin VB.PictureBox picHandleA 
      BackColor       =   &H00FFFF00&
      Height          =   375
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   60
      Width           =   375
   End
   Begin VB.PictureBox picBox 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   480
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   0
      Top             =   60
      Width           =   9615
   End
   Begin VB.Label lblRender 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   7320
      Width           =   465
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub cmdExTwo_Click()
    frmExampleTwo.Show
End Sub

Private Sub Form_Load()
    DrawG
End Sub

Private Sub picHandleA_Click()
    cDLG.Color = picHandleA.BackColor
    cDLG.ShowColor
    picHandleA.BackColor = cDLG.Color
    DrawG
End Sub

Private Function DrawG()
    Dim iTick As Long
    iTick = GetTickCount
    DrawGradient picBox.hDC, 0, 0, picBox.ScaleWidth, picBox.ScaleHeight, picHandleA.BackColor, picHandleB.BackColor, picHandleD.BackColor, picHandleC.BackColor
    picBox.Refresh
    lblRender = "Rendering took: " & ((GetTickCount - iTick) / 1000) & " seconds."
End Function

Private Sub picHandleB_Click()
    cDLG.Color = picHandleB.BackColor
    cDLG.ShowColor
    picHandleB.BackColor = cDLG.Color
    DrawG
End Sub

Private Sub picHandleC_Click()
    cDLG.Color = picHandleC.BackColor
    cDLG.ShowColor
    picHandleC.BackColor = cDLG.Color
    DrawG
End Sub

Private Sub picHandleD_Click()
    cDLG.Color = picHandleD.BackColor
    cDLG.ShowColor
    picHandleD.BackColor = cDLG.Color
    DrawG
End Sub
