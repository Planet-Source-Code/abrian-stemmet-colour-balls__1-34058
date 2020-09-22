VERSION 5.00
Begin VB.Form frmBalls 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balls"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBlueTop 
      BackColor       =   &H00000000&
      Caption         =   "Blue Top"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkBlueLeft 
      BackColor       =   &H00000000&
      Caption         =   "Blue Left"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkRedLeft 
      BackColor       =   &H00000000&
      Caption         =   "Red Left"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkRedTop 
      BackColor       =   &H00000000&
      Caption         =   "Red Top"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Timer BlueTop 
      Interval        =   2
      Left            =   5040
      Top             =   2760
   End
   Begin VB.Timer BlueLeft 
      Interval        =   1
      Left            =   4560
      Top             =   2760
   End
   Begin VB.Timer RedTopTimer 
      Interval        =   2
      Left            =   480
      Top             =   2640
   End
   Begin VB.Timer RedLeftTimer 
      Interval        =   1
      Left            =   0
      Top             =   2640
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape picBlue 
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape picRed 
      BorderColor     =   &H000000FF&
      Height          =   495
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmBalls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BlueLeft_Timer()
    If BlueLeft.Interval = 1 Then
        picBlue.Left = picBlue.Left - 50
            If picBlue.Left < 0 Then
                BlueLeft.Interval = 2
            End If
        picBlue.Left = picBlue.Left - 50
    End If
    
    If BlueLeft.Interval = 2 Then
        picBlue.Left = picBlue.Left + 50
            If picBlue.Left > 5000 Then
                BlueLeft.Interval = 1
            End If
        picBlue.Left = picBlue.Left + 50
    End If
End Sub

Private Sub BlueTop_Timer()
    If BlueTop.Interval = 1 Then
        picBlue.Top = picBlue.Top - 50
            If picBlue.Top < 0 Then
                BlueTop.Interval = 2
            End If
        picBlue.Top = picBlue.Top - 50
    End If
    
    If BlueTop.Interval = 2 Then
        picBlue.Top = picBlue.Top + 50
            If picBlue.Top > 3000 Then
                BlueTop.Interval = 1
                    If picBlue.BorderColor = vbRed Then
                        picBlue.BorderColor = vbGreen
                        Label2.Caption = "Green"
                    ElseIf picBlue.BorderColor = vbGreen Then
                        picBlue.BorderColor = vbBlue
                        Label2.Caption = "Blue"
                    ElseIf picBlue.BorderColor = vbBlue Then
                        picBlue.BorderColor = vbYellow
                        Label2.Caption = "Yellow"
                    ElseIf picBlue.BorderColor = vbYellow Then
                        picBlue.BorderColor = vbRed
                        Label2.Caption = "Red"
                    End If
            End If
        picBlue.Top = picBlue.Top + 50
    End If
End Sub

Private Sub chkBlueLeft_Click()
    If chkBlueLeft.Value = Unchecked Then
        BlueLeft.Enabled = False
    Else
        BlueLeft.Enabled = True
    End If
End Sub

Private Sub chkBlueTop_Click()
    If chkBlueTop.Value = Unchecked Then
        BlueTop.Enabled = False
    Else
        BlueTop.Enabled = True
    End If
End Sub

Private Sub chkRedLeft_Click()
    If chkRedLeft.Value = Unchecked Then
        RedLeftTimer.Enabled = False
    Else
        RedLeftTimer.Enabled = True
    End If
End Sub

Private Sub chkRedTop_Click()
    If chkRedTop.Value = Unchecked Then
        RedTopTimer.Enabled = False
    Else
        RedTopTimer.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    picRed.Left = 0
    picRed.Top = 0
    picRed.BorderColor = vbRed
    Label1.Caption = "Red"
    picBlue.BorderColor = vbYellow
    Label2.Caption = "Yellow"
End Sub

Private Sub redLeftTimer_Timer()

    If RedLeftTimer.Interval = 1 Then
        picRed.Left = picRed.Left - 50
            If picRed.Left < 0 Then
                RedLeftTimer.Interval = 2
            End If
        picRed.Left = picRed.Left - 50
    End If
    
    If RedLeftTimer.Interval = 2 Then
        picRed.Left = picRed.Left + 50
            If picRed.Left > 5000 Then
                RedLeftTimer.Interval = 1
            End If
        picRed.Left = picRed.Left + 50
    End If
        
End Sub

Private Sub RedTopTimer_Timer()
    If RedTopTimer.Interval = 1 Then
        picRed.Top = picRed.Top - 50
            If picRed.Top < 0 Then
                RedTopTimer.Interval = 2
            End If
        picRed.Top = picRed.Top - 50
    End If
    
    If RedTopTimer.Interval = 2 Then
        picRed.Top = picRed.Top + 50
            If picRed.Top > 3000 Then
                RedTopTimer.Interval = 1
                    If picRed.BorderColor = vbRed Then
                        picRed.BorderColor = vbGreen
                        Label1.Caption = "Green"
                    ElseIf picRed.BorderColor = vbGreen Then
                        picRed.BorderColor = vbBlue
                        Label1.Caption = "Blue"
                    ElseIf picRed.BorderColor = vbBlue Then
                        picRed.BorderColor = vbYellow
                        Label1.Caption = "Yellow"
                    ElseIf picRed.BorderColor = vbYellow Then
                        picRed.BorderColor = vbRed
                        Label1.Caption = "Red"
                    End If
            End If
        picRed.Top = picRed.Top + 50
    End If
End Sub
