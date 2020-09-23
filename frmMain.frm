VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Time Difference"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   285
      Left            =   2580
      TabIndex        =   5
      Top             =   540
      Width           =   1245
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1350
      Top             =   0
   End
   Begin VB.CommandButton cmdDiff 
      Caption         =   "&Get Difference"
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   540
      Width           =   1245
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   540
      Width           =   1245
   End
   Begin VB.Label lblDiff 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   270
      Width           =   45
   End
   Begin VB.Label lblClock 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2970
      TabIndex        =   2
      Top             =   30
      Width           =   45
   End
   Begin VB.Label lblStartTime 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   45
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDiff_Click()
Timer1.Enabled = False
lblDiff.Caption = "Time Difference= " & GetHours(StartTime, Format(Now, "Short Time")) & " hr(s) " _
& GetMins(StartTime, Format(Now, "Short Time")) & " min(s)"
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdStart_Click()
StartTime = Format(Now, "Short Time")
lblStartTime.Caption = Format(Now, "Long Time")
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
lblClock.Caption = Format(Now, "Long Time")
End Sub
