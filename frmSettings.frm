VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   6060
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHard 
      Caption         =   "Hard"
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdMedium 
      Caption         =   "Medium"
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdEasy 
      Caption         =   "Easy"
      Height          =   495
      Left            =   840
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Preset Sizing:"
      Height          =   2655
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox txtMines 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtHeight 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtWidth 
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Custom Sizing:"
      Height          =   2655
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   2775
      Begin VB.Label Label3 
         Caption         =   "Mines:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Height:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Width:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    If txtWidth.Text <> "" And txtHeight.Text <> "" And txtMines.Text <> "" Then
        intWidth = txtWidth.Text
        intHeight = txtHeight.Text
        intMines = txtMines.Text
        If intMines < intWidth * intHeight Then
            frmMain.Show
            Unload Me
        Else
            MsgBox ("Too many mines")
        End If
    Else
        MsgBox ("Please Enter All Values")
    End If
End Sub

Private Sub cmdEasy_Click()
    intWidth = 9
    intHeight = 9
    intMines = 10
    frmMain.Show
    Unload Me
End Sub

Private Sub cmdHard_Click()
    intWidth = 30
    intHeight = 16
    intMines = 99
    frmMain.Show
    Unload Me
End Sub

Private Sub cmdMedium_Click()
    intWidth = 16
    intHeight = 16
    intMines = 40
    frmMain.Show
    Unload Me
End Sub
