VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minesweeper"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   3270
   FillColor       =   &H00808000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3270
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   720
   End
   Begin VB.CommandButton cmdReset 
      Height          =   495
      Left            =   1560
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      Height          =   255
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblMines 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Quartz"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label number 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private booMine() As Boolean
Private intNeighbors() As Integer
Private intTouchingMines() As Integer
Private halt As Boolean
Private intMineList() As Integer
Private booChecked() As Boolean
Private checked As Integer
Private clicks As Integer
Private time As Integer
Option Explicit

Private Sub cmdButton_Click(Index As Integer)
    If halt = False And booChecked(Index) = False Then
        Timer1.Enabled = True
        If cmdButton(Index).Visible = True And booMine(Index) = False Then
            clicks = clicks + 1
            Dim i As Integer
            cmdButton(Index).Visible = False
                If clicks = (intHeight * intWidth) - intMines Then
                    Timer1.Enabled = False
                    halt = True
                    cmdReset.Picture = LoadPicture(App.Path & "\won.bmp")
                    For i = 1 To intMines
                        cmdButton(intMineList(i)).Picture = LoadPicture(App.Path & "\flag.bmp")
                    Next i
                    lblMines.Caption = "000"
                End If
            If intTouchingMines(Index) = 0 And booMine(Index) = False Then
                For i = 1 To 8
                    cmdButton_Click (intNeighbors(Index, i))
                Next i
            End If
        End If
    
        If booMine(Index) = True Then
            halt = True
            Timer1.Enabled = False
            cmdReset.Picture = LoadPicture(App.Path & "\dead.bmp")
            For i = 1 To intMines
                cmdButton(intMineList(i)).Picture = LoadPicture(App.Path & "\mine.bmp")
            Next i
            For i = 1 To intHeight * intWidth
                If booChecked(i) = True Then
                    If booMine(i) = True Then
                        cmdButton(i).Picture = LoadPicture(App.Path & "/flag.bmp")
                    Else
                        cmdButton(i).Picture = LoadPicture(App.Path & "/false.bmp")
                    End If
                End If
            Next i
            cmdButton(Index).Picture = LoadPicture(App.Path & "\redmine.bmp")
        End If
    End If
End Sub



Private Sub cmdButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If halt = False Then
        If Button = 2 Then
            If booChecked(Index) = False Then
                cmdButton(Index).Picture = LoadPicture(App.Path & "\flag.bmp")
                booChecked(Index) = True
                checked = checked + 1
            Else
                cmdButton(Index).Picture = LoadPicture()
                booChecked(Index) = False
                checked = checked - 1
            End If
        End If
    lblMines.Caption = Format(intMines - checked, "000")
    End If
End Sub


Private Sub cmdReset_Click()
    Dim i As Integer
    Dim e As Integer
    Dim intCurrentIndex As Integer
    Timer1.Enabled = False
    time = 0
    lblTime.Caption = "000"
    clicks = 0
    checked = 0
    lblMines.Caption = Format(intMines, "000")
    intCurrentIndex = 1
    Dim intIndex() As Integer
    Dim mine As Integer
    halt = False
    cmdReset.Picture = LoadPicture(App.Path & "\smile.bmp")
    For i = 1 To intHeight * intWidth
        booMine(i) = False
        booChecked(i) = False
        cmdButton(i).Picture = LoadPicture()
        intTouchingMines(i) = 0
        cmdButton(i).Visible = True
    Next i
    For i = 1 To intMines
        Do
            mine = Rnd * (intWidth * intHeight - 1) + 1
        Loop Until booMine(mine) = False
        booMine(mine) = True
        For e = 1 To 8
            intTouchingMines(intNeighbors(mine, e)) = intTouchingMines(intNeighbors(mine, e)) + 1
        Next e
        intMineList(i) = mine
    Next i
    For i = 1 To intHeight * intWidth
        If intTouchingMines(i) > 0 Then
            number(i).Caption = intTouchingMines(i)
            If intTouchingMines(i) = 1 Then
                number(i).ForeColor = vbBlue
            ElseIf intTouchingMines(i) = 2 Then
                number(i).ForeColor = &HC000&
            ElseIf intTouchingMines(i) = 3 Then
                number(i).ForeColor = vbRed
            ElseIf intTouchingMines(i) = 4 Then
                number(i).ForeColor = &H800000
            ElseIf intTouchingMines(i) = 5 Then
                number(i).ForeColor = &H80&
            ElseIf intTouchingMines(i) = 6 Then
                number(i).ForeColor = &H808000
            End If
        Else
            number(i).Caption = ""
        End If
    Next i
            
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim e As Integer
    ReDim intMineList(intMines)
    Dim intCurrentIndex As Integer
    lblMines.Caption = Format(intMines, "000")
    intCurrentIndex = 1
    Dim intIndex() As Integer
    ReDim booChecked(intWidth * intHeight)
    ReDim intIndex(intWidth + 1, intHeight + 1) As Integer
    ReDim intNeighbors(intWidth * intHeight, 8)
    ReDim booMine(intWidth * intHeight)
    ReDim intTouchingMines(intWidth * intHeight)
    'build interface
    For i = 1 To intWidth
        For e = 1 To intHeight
            Load cmdButton(intCurrentIndex)
            Load number(intCurrentIndex)
            intIndex(i, e) = intCurrentIndex
            cmdButton(intCurrentIndex).Top = 480 + (e * 255)
            cmdButton(intCurrentIndex).Left = (i * 255)
            number(intCurrentIndex).Top = 480 + (e * 255)
            number(intCurrentIndex).Left = (i * 255)
            number(intCurrentIndex).Visible = True
            cmdButton(intCurrentIndex).Visible = True
            intCurrentIndex = intCurrentIndex + 1
        Next e
    Next i
    
    Me.Width = intWidth * 255 + 600
    Me.Height = intHeight * 255 + 1565
    cmdReset.Left = (Me.Width / 2) - 375
    lblMines.Left = cmdReset.Left - 960
    lblTime.Left = cmdReset.Left + 720
    
    For i = 1 To intWidth
        For e = 1 To intHeight
            intNeighbors(intIndex(i, e), 1) = intIndex(i, e + 1)
            intNeighbors(intIndex(i, e), 2) = intIndex(i, e - 1)
            intNeighbors(intIndex(i, e), 3) = intIndex(i + 1, e)
            intNeighbors(intIndex(i, e), 4) = intIndex(i + 1, e + 1)
            intNeighbors(intIndex(i, e), 5) = intIndex(i + 1, e - 1)
            intNeighbors(intIndex(i, e), 6) = intIndex(i - 1, e)
            intNeighbors(intIndex(i, e), 7) = intIndex(i - 1, e + 1)
            intNeighbors(intIndex(i, e), 8) = intIndex(i - 1, e - 1)
        Next e
    Next i
    Dim mine As Integer
    For i = 1 To intMines
        Do
            mine = Rnd * (intWidth * intHeight - 1) + 1
        Loop Until booMine(mine) = False
        booMine(mine) = True
        For e = 1 To 8
            intTouchingMines(intNeighbors(mine, e)) = intTouchingMines(intNeighbors(mine, e)) + 1
        Next e
        intMineList(i) = mine
    Next i
    For i = 1 To intHeight * intWidth
        If intTouchingMines(i) > 0 Then
            number(i).Caption = intTouchingMines(i)
            If intTouchingMines(i) = 1 Then
                number(i).ForeColor = vbBlue
            ElseIf intTouchingMines(i) = 2 Then
                number(i).ForeColor = &HC000&
            ElseIf intTouchingMines(i) = 3 Then
                number(i).ForeColor = vbRed
            ElseIf intTouchingMines(i) = 4 Then
                number(i).ForeColor = &H800000
            ElseIf intTouchingMines(i) = 5 Then
                number(i).ForeColor = &H80&
            ElseIf intTouchingMines(i) = 6 Then
                number(i).ForeColor = &H808000
            End If
        Else
            number(i).Caption = ""
        End If
    Next i
                
            
End Sub

Private Sub Timer1_Timer()
    time = time + 1
    lblTime.Caption = Format(time, "000")
End Sub
