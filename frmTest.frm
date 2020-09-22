VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmTest 
   Caption         =   "   Test the RoundTime Routine     by    BB WS1O"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider sldInterval 
      Height          =   450
      Left            =   150
      TabIndex        =   3
      ToolTipText     =   "Move pointer to select interval you want to round time down to, in minutes"
      Top             =   1785
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   794
      _Version        =   393216
      Max             =   60
      SelStart        =   1
      TickFrequency   =   5
      Value           =   1
      TextPosition    =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Enough"
      Height          =   360
      Left            =   1875
      TabIndex        =   2
      Top             =   2640
      Width           =   885
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3720
      Top             =   2595
   End
   Begin VB.Label lblSelectedInterval 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3405
      TabIndex        =   12
      Top             =   1485
      Width           =   105
   End
   Begin VB.Label lblSelectInterval 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rounding Interval in Minutes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1155
      TabIndex        =   11
      Top             =   1485
      Width           =   2115
   End
   Begin VB.Label lblInterval 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      Height          =   210
      Index           =   6
      Left            =   3585
      TabIndex        =   10
      Top             =   2295
      Width           =   195
   End
   Begin VB.Label lblInterval 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      Height          =   210
      Index           =   5
      Left            =   2925
      TabIndex        =   9
      Top             =   2295
      Width           =   195
   End
   Begin VB.Label lblInterval 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      Height          =   210
      Index           =   4
      Left            =   1590
      TabIndex        =   8
      Top             =   2295
      Width           =   195
   End
   Begin VB.Label lblInterval 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   210
      Index           =   3
      Left            =   900
      TabIndex        =   7
      Top             =   2295
      Width           =   195
   End
   Begin VB.Label lblInterval 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      Height          =   210
      Index           =   2
      Left            =   2235
      TabIndex        =   6
      Top             =   2295
      Width           =   195
   End
   Begin VB.Label lblInterval 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      Height          =   210
      Index           =   1
      Left            =   4275
      TabIndex        =   5
      Top             =   2295
      Width           =   195
   End
   Begin VB.Label lblInterval 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   210
      Index           =   0
      Left            =   345
      TabIndex        =   4
      Top             =   2295
      Width           =   105
   End
   Begin VB.Label lblRealTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Real Time:  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   330
      Left            =   690
      TabIndex        =   1
      Top             =   210
      Width           =   1620
   End
   Begin VB.Label lblRoundedTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rounded Time:  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   75
      TabIndex        =   0
      Top             =   780
      Width           =   2280
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
   
    '----------------------------------------------
    ' Purpose   : close frmTest, close down the app
    ' Modified  : 5/24/2001 By BB
    '----------------------------------------------

    On Error GoTo Err_cmdClose_Click

    Unload Me

Exit_cmdClose_Click:
    
    End
    On Error GoTo 0
    Exit Sub

Err_cmdClose_Click:
     
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmTest" & " during " & "cmdClose_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdClose_Click
    End Select

End Sub
Private Sub Form_Load()
   
    '---------------------------------------------------------------
    ' Purpose   : set the initial value of the rounding interval
    ' Modified  : 5/24/2001 By BB
    '---------------------------------------------------------------

    On Error GoTo Err_Form_Load

    ' start with interval set to 10 minutes
    sldInterval.Value = 10

Exit_Form_Load:
    
    On Error GoTo 0
    Exit Sub

Err_Form_Load:
     
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmTest" & " during " & "Form_Load" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Form_Load
    End Select

End Sub
Private Sub sldInterval_Change()
   
    '---------------------------------------------------------------
    ' Purpose   : changes Rounding Interval to minutes selected by user
    ' Modified  : 5/24/2001 By BB
    '---------------------------------------------------------------

    On Error GoTo Err_sldInterval_Change
    
        ' had to set the slider control's minimum value to 0,
        ' otherwise the tick marks don't line up right on the
        ' 5-minute marks (is this a bug, Microsoft?), but
        ' because we don't want the actual minimum value to be
        ' set to 0 (which causes a divide by zero error), we
        ' just force it to 1 if the user tries to set it to 0
        
        '  make sure we don't allow the user to
        '  select a value less than 1 or more than 60
        ' (above 60 wouldn't make sense, and less than
        '  1 would cause a divide by zero error)
        If sldInterval.Value < gcintMin Then
            sldInterval.Value = 1
            glngInterval = 1
        ElseIf sldInterval.Value > gcintMax Then
            sldInterval.Value = 60
            glngInterval = 60
        Else
            glngInterval = sldInterval.Value
        End If
    lblSelectedInterval.Caption = glngInterval

Exit_sldInterval_Change:
    
    On Error GoTo 0
    Exit Sub

Err_sldInterval_Change:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmTest" & " during " & "sldInterval_Change" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_sldInterval_Change
    End Select

End Sub
Private Sub Timer1_Timer()
   
    '---------------------------------------------------------------
    ' Purpose   : generate a formatted time string to process
    ' Modified  : 5/24/2001 By BB
    '---------------------------------------------------------------

    On Error GoTo Err_Timer1_Timer

    On Error Resume Next
    
    '    get time Now and pass to the function to return the
    ' current time rounded down to the nearest 10-minute interval
    '        (that's what the "10" in the line below means)
    'gstrRoundedTime = Format$(dhRoundTime(Now(), 10), "DDDD hh:nn")
    
    ' for this demo program, we'll let the user set the interval
    ' by selecting it from the slider control on the form, which we
    ' assign to the public glngInterval variable
    gstrRoundedTime = Format$(dhRoundTime(Now(), glngInterval), "DDDD hh:nn")
    
    ' get the curent time Now without any rounding
    gstrRealTime = Format$(Now(), "DDDD hh:nn")
    
    ' display real time and rounded time on form's labels
    lblRealTime.Caption = "Real Time: " & gstrRealTime
    lblRoundedTime.Caption = "Rounded Time: " & gstrRoundedTime
    
    ' show real time on form's title bar, just for reference
    Me.Caption = "Test RoundTime Function by BB WS1O  " & Format$(Now(), "h:nn:ss AMPM")
    
Exit_Timer1_Timer:
    
    On Error GoTo 0
    Exit Sub

Err_Timer1_Timer:
     
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmTest" & " during " & "Timer1_Timer" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Timer1_Timer
    End Select

End Sub
