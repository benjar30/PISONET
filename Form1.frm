VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "howto_countdown_timer"
   ClientHeight    =   660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   ScaleHeight     =   660
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1200
   End
   Begin VB.TextBox txtDuration 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblRemaining 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_StopTime As Date
Public Sub cmdGo()


Dim fields() As String
Dim hours As Long
Dim minutes As Long
Dim seconds As Long
    fields = Split(txtDuration.Text, ":")
    seconds = fields(0)

    m_StopTime = Now
    m_StopTime = DateAdd("h", hours, m_StopTime)
    m_StopTime = DateAdd("n", minutes, m_StopTime)
    m_StopTime = DateAdd("s", seconds, m_StopTime)

    tmrWait.Enabled = True
    

End Sub


Private Sub tmrWait_Timer()
Dim time_now As Date
Dim hours As Long
Dim minutes As Long
Dim seconds As Long

    time_now = Now
    If time_now >= m_StopTime Then
        tmrWait.Enabled = False
        'frmTerminal.tunoff
        If frmTerminal.Text2.Text = "0" Then
        frmTerminal.tunoff
        Else
        frmTerminal.logout
        End If
        lblRemaining.Caption = "0:00:00"
    Else
        seconds = DateDiff("s", time_now, m_StopTime)
        minutes = seconds \ 60
        seconds = seconds - minutes * 60
        hours = minutes \ 60
        minutes = minutes - hours * 60

        lblRemaining.Caption = _
            Format$(hours) & ":" & _
            Format$(minutes, "00") & ":" & _
            Format$(seconds, "00")
    End If
End Sub
