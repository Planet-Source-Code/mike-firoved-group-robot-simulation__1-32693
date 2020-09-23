VERSION 5.00
Begin VB.Form frmField 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Robot Behavior"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   4440
      Top             =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Log"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   7260
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Interval        =   188
      Left            =   4020
      Top             =   2640
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmField.frx":0000
      Height          =   375
      Left            =   1140
      TabIndex        =   2
      Top             =   7140
      Width           =   8775
   End
   Begin VB.Label robotText 
      BackStyle       =   0  'Transparent
      Caption         =   " 00 "
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image robotOverlay 
      Height          =   255
      Index           =   0
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   7035
      Left            =   120
      Top             =   120
      Width           =   9795
   End
   Begin VB.Shape robot 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C2898D&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   2160
      Top             =   1440
      Width           =   255
   End
End
Attribute VB_Name = "frmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
     frmStatus.Show
End Sub

Private Sub Form_Load()
    On Error GoTo ERRForm_Load
    Randomize Timer

    'Initialize robots
    For xa = 0 To 100
        If xa >= 1 Then
            Load robotText(xa)
            Load robot(xa)
            Load robotOverlay(xa)
        End If
        robotText(xa).Visible = True
        robot(xa).Visible = True
        robotOverlay(xa).Visible = True
        robotText(xa).Visible = True
        robotText(xa).Caption = " " & CStr(xa) & " "
        l = Int((9495 - 120 + 1) * Rnd(2) + 120)
        t = Int((6855 - 120 + 1) * Rnd(2) + 120)
        robot(xa).Tag = CStr(l) & "|" & CStr(t) & "|0|+"
        l = Int((9495 - 120 + 1) * Rnd(2) + 120)
        t = Int((6855 - 120 + 1) * Rnd(2) + 120)
        robot(xa).Left = l
        robotOverlay(xa).Left = l
        robotText(xa).Left = l
        robot(xa).Top = t
        robotOverlay(xa).Top = t
        robotText(xa).Top = t

    Next
    'Clear the log
    frmStatus.List1.Clear
    Timer2_Timer
    Exit Sub
    
ERRForm_Load:
    'save errors in registry to view later
    SaveSetting "GrpRobot", "error", "2nd", GetSetting("GrpRobot", "error", "Last", ""): SaveSetting "GrpRobot", "error", "3rd", GetSetting("GrpRobot", "error", "2nd", "")
    SaveSetting "GrpRobot", "error", "Last", "GrpRobot.frmField.Form_Load." & App.ThreadID & " Produced the following error:  " & Err.Description
    Resume Next

End Sub

Private Sub Timer1_Timer()
    On Error GoTo ERRTimer1_Timer
    For xa = 0 To robot.Count - 1
        DoEvents
        
        'get robot coordinates
        inf = Split(robot(xa).Tag, "|")
        dl = inf(0)
        dt = inf(1)
        grpID = inf(2)
        role = inf(3)
        Set inf = Nothing
        diffl = Val(dl) - robot(xa).Left
        difft = Val(dt) - robot(xa).Top
        
        'frmStatus.List1.List(xa) = "Robot ID " & xa & " " & robot(xa).Left & "," & robot(xa).Top & " to " & dl & "," & dt & ""

        robotOverlay(xa).ToolTipText = "Robot ID " & xa & " " & robot(xa).Left & "," & robot(xa).Top & " to " & dl & "," & dt & ""
        
        If Abs(diffl) > 50 Then
            diffl = diffl / 20
            robot(xa).Left = robot(xa).Left + diffl
            robotOverlay(xa).Left = robotOverlay(xa).Left + diffl
            robotText(xa).Left = robotText(xa).Left + diffl

        Else
            robot(xa).Left = robot(xa).Left + diffl
            robotOverlay(xa).Left = robotOverlay(xa).Left + diffl
            robotText(xa).Left = robotText(xa).Left + diffl
       
        End If
        If Abs(difft) > 50 Then
            difft = difft / 20
            robot(xa).Top = robot(xa).Top + difft
            robotOverlay(xa).Top = robotOverlay(xa).Top + difft
            robotText(xa).Top = robotText(xa).Top + difft
           
        Else
            robot(xa).Top = robot(xa).Top + difft
            robotOverlay(xa).Top = robotOverlay(xa).Top + difft
            robotText(xa).Top = robotText(xa).Top + difft
           
        End If
        aa = Abs(difft) + Abs(diffl)
        DoEvents
        'Change color based open how far one is to a target
        robotText(xa).ForeColor = black
        Select Case aa
            Case 161 To 16350
                robot(xa).FillColor = vbRed
                
            Case 50 To 160
                robot(xa).FillColor = vbMagenta
                
            Case 30 To 49
                robot(xa).FillColor = vbBlue
                robotText(xa).ForeColor = vbWhite
                
            Case 15 To 29
                robot(xa).FillColor = vbYellow
                robotText(xa).ForeColor = vbBlue
                
            Case 1 To 14
                robot(xa).FillColor = vbCyan
                
            Case 0
                If robot(xa).FillColor = vbGreen Then
                    el = Int((9495 - 120 + 1) * Rnd(2) + 120)
                    et = Int((6855 - 120 + 1) * Rnd(2) + 120)
                    robot(xa).Tag = CStr(el) & "|" & CStr(et) & "|0|+"
                End If
                robot(xa).FillColor = vbGreen
        End Select
    Next
    aa = 0
    Exit Sub
    
ERRTimer1_Timer:
    SaveSetting "GrpRobot", "error", "2nd", GetSetting("GrpRobot", "error", "Last", ""): SaveSetting "GrpRobot", "error", "3rd", GetSetting("GrpRobot", "error", "2nd", "")
    SaveSetting "GrpRobot", "error", "Last", "GrpRobot.frmField.Timer1_Timer." & App.ThreadID & " Produced the following error:  " & Err.Description
    Resume Next

    
End Sub

Private Sub Timer2_Timer()
    On Error GoTo ERRTimer2_Timer
    frmStatus.List1.Visible = False
    For xa = 0 To robot.Count - 1
        DoEvents
        inf = Split(robot(xa).Tag, "|")
        dl = inf(0)
        dt = inf(1)
        Set inf = Nothing
        frmStatus.List1.List(xa) = "Robot ID " & xa & " " & robot(xa).Left & "," & robot(xa).Top & " to " & dl & "," & dt & ""
    Next
    frmStatus.List1.Visible = True
    Exit Sub
    
ERRTimer2_Timer:
    SaveSetting "GrpRobot", "error", "2nd", GetSetting("GrpRobot", "error", "Last", ""): SaveSetting "GrpRobot", "error", "3rd", GetSetting("GrpRobot", "error", "2nd", "")
    SaveSetting "GrpRobot", "error", "Last", "GrpRobot.frmField.Timer2_Timer." & App.ThreadID & " Produced the following error:  " & Err.Description
    Resume Next

End Sub
