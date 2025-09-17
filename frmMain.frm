VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00202020&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ShutdownTimer2"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10860
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ShutdownTimer2.StatusWindow wStatus 
      Height          =   1185
      Left            =   6240
      TabIndex        =   5
      Top             =   5805
      Width           =   3555
      _ExtentX        =   5080
      _ExtentY        =   1667
      Title           =   "Will do something"
      Caption         =   "00:00:00"
   End
   Begin ShutdownTimer2.FloodBar flProgress 
      Height          =   840
      Left            =   840
      TabIndex        =   4
      Top             =   5505
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   1164
   End
   Begin ShutdownTimer2.ShutdownOptions optShutdown 
      Height          =   1590
      Left            =   555
      TabIndex        =   3
      Top             =   3660
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   2805
      Title           =   "Options"
      Enabled         =   0   'False
   End
   Begin ShutdownTimer2.Button cmdStart 
      Height          =   705
      Left            =   7740
      TabIndex        =   2
      Top             =   2100
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   1244
      Caption         =   "Start"
      BackColor       =   5023791
      HoverColor      =   6339651
      PressedColor    =   4033061
      ForeColor       =   15463654
      DisabledBackColor=   6455130
      DisabledTextColor=   13355947
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShutdownTimer2.ShutdownSelector cmbShutdown 
      Height          =   1500
      Left            =   2235
      TabIndex        =   1
      Top             =   330
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   3440
      Title           =   "System Action"
   End
   Begin ShutdownTimer2.DelaySelector txtDelay 
      Height          =   840
      Left            =   1635
      TabIndex        =   0
      Top             =   1995
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   1482
      Title           =   "Delay"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FixedFormSize As POINTAPI

Dim WithEvents CountdownTimer As ShutdownTimer
Attribute CountdownTimer.VB_VarHelpID = -1

Dim IsRunning As Boolean

Friend Sub SetForm()
  ChangedByCode = True
  ChangedByCode = False
  PrepareObjects
  MoveObjects
  Me.Caption = APP_NAME
  LoadSettings
  ApplySettings
  cmbShutdown.AllowHibernate = CanHibernate()
  If o_AlwaysOnTop = True Then Call optShutdown_AlwaysOnTopClick
  CheckReadyButton
  optShutdown.UpdatesEnabled = IsWindowsVistaOrHigher()
  If optShutdown.Enabled = False Then optShutdown.InstallUpdates = False
  Me.Show
End Sub

Private Sub PrepareObjects()
  With cmdStart
    With .Font
      .Name = "Segoe UI"
      .Size = 11
      .Bold = True
    End With
  End With
  Dim c As Object
  For Each c In Me.Controls
    c.Refresh True
  Next
End Sub

Private Sub MoveObjects()
  ChangedByCode = True
  cmbShutdown.Move (Screen.TwipsPerPixelX * 3), (Screen.TwipsPerPixelY * 3)
  txtDelay.Left = cmbShutdown.Left
  cmdStart.Move ((txtDelay.Left + txtDelay.Width) + (Screen.TwipsPerPixelX * 10)), ((cmbShutdown.Top + cmbShutdown.Height) + (Screen.TwipsPerPixelY * 7))
  txtDelay.Top = ((cmdStart.Top + cmdStart.Height) - txtDelay.Height)
  cmbShutdown.Width = ((cmdStart.Left + cmdStart.Width) - cmbShutdown.Left)
  Me.Width = ((Me.Width - Me.ScaleWidth) + ((cmdStart.Left + cmdStart.Width) + cmbShutdown.Left))
  optShutdown.Move txtDelay.Left, ((txtDelay.Top + txtDelay.Height) + (Screen.TwipsPerPixelY * 1)), (Me.ScaleWidth - (txtDelay.Left * 2))
  Me.Height = ((Me.Height - Me.ScaleHeight) + (((optShutdown.Top + optShutdown.Height) + cmbShutdown.Top) + (Screen.TwipsPerPixelY * 1)))
  With FixedFormSize
    .X = Me.Width
    .Y = Me.Height
  End With
  flProgress.Move txtDelay.Left, (txtDelay.Top + (Screen.TwipsPerPixelY * 5)), txtDelay.Width, (txtDelay.Height - (Screen.TwipsPerPixelY * 5))
  flProgress.Visible = False
  wStatus.Move cmbShutdown.Left, cmbShutdown.Top, cmbShutdown.Width, cmbShutdown.Height
  wStatus.Visible = False
  ChangedByCode = False
End Sub

Private Sub ApplySettings()
  optShutdown.ForceExit = o_ForceExit
  optShutdown.InstallUpdates = o_InstallUpdates
  optShutdown.AlwaysOnTop = o_AlwaysOnTop
  With o_Pos
    Me.Move .X, .Y
  End With
  cmbShutdown.SelectedShutdownMethod = o_Method
  DelayTime = o_Time
  With DelayTime
    txtDelay.Hours = .Hours
    txtDelay.Minutes = .Minutes
    txtDelay.Seconds = .Seconds
  End With
End Sub

Private Sub CollectDelayTime()
  With DelayTime
    .Hours = txtDelay.Hours
    .Minutes = txtDelay.Minutes
    .Seconds = txtDelay.Seconds
  End With
End Sub

Private Function CountTotalDelayUnit(cTime As Time_Info, UnitType As String) As Long
  Select Case LCase$(UnitType)
    Case "h"
      CountTotalDelayUnit = CountTotalHours(cTime)
    Case "m"
      CountTotalDelayUnit = CountTotalMinutes(cTime)
    Case "s"
      CountTotalDelayUnit = CountTotalSeconds(cTime)
  End Select
End Function

Private Function CountTotalHours(cTime As Time_Info) As Long
  With cTime
    CountTotalHours = .Hours + ((.Minutes + (.Seconds \ 60)) \ 60)
  End With
End Function

Private Function CountTotalMinutes(cTime As Time_Info) As Long
  With cTime
    CountTotalMinutes = (.Hours * 60) + .Minutes + (.Seconds \ 60)
  End With
End Function

Private Function CountTotalSeconds(cTime As Time_Info) As Long
  With cTime
    CountTotalSeconds = (.Hours * 3600) + (.Minutes * 60) + .Seconds
  End With
End Function

Private Sub CheckReadyButton()
  If IsRunning Then cmdStart.Enabled = True: Exit Sub
  With DelayTime
    If Not ((.Hours = 0 And .Minutes = 0) And .Seconds = 0) Then
      If o_Method > smNone Then
        cmdStart.Enabled = True
        Exit Sub
      End If
    End If
  End With
  cmdStart.Enabled = False
End Sub

Private Sub SetTimer()
  KillTimer
  Set CountdownTimer = New ShutdownTimer
  CountdownTimer.Interval = 1000 '1 second
  CountdownTimer.Enabled = True
End Sub

Private Sub KillTimer()
  If Not CountdownTimer Is Nothing Then
    CountdownTimer.Enabled = False
    Set CountdownTimer = Nothing
  End If
End Sub

Private Sub UpdateStatusTime(sTime As Time_Info)
  With sTime
    wStatus.Hours = .Hours
    wStatus.Minutes = .Minutes
    wStatus.Seconds = .Seconds
  End With
End Sub

Private Sub SetEndTime()
  EndTime = DateAdd("h", DelayTime.Hours, DateAdd("n", DelayTime.Minutes, DateAdd("s", DelayTime.Seconds, Now)))
End Sub

Private Function CalculateRemainingTime() As Time_Info
  With CalculateRemainingTime
    .Seconds = DateDiff("s", Now, EndTime, vbMonday)
    .Hours = (.Seconds \ 3600)
    .Minutes = (.Seconds \ 60) Mod 60
    .Seconds = .Seconds Mod 60
  End With
End Function

Private Function GetStatusTitle() As String
  Select Case cmbShutdown.SelectedShutdownMethod
    Case Shutdown_Method.smShutdown
      GetStatusTitle = "System will shut down in"
    Case Shutdown_Method.smReboot
      GetStatusTitle = "System will restart in"
    Case Shutdown_Method.smHibernate
      GetStatusTitle = "System will hibernate in"
    Case Else
      GetStatusTitle = "Something went wrong"
  End Select
End Function

Private Sub StartCountdown()
  cmbShutdown.Visible = False
  wStatus.Visible = True
  txtDelay.Visible = False
  With flProgress
    .Visible = True
    .Min = 0
    .Max = CountTotalSeconds(DelayTime)
    .Value = .Min
  End With
  SetEndTime
  wStatus.Title = GetStatusTitle()
  UpdateStatusTime CalculateRemainingTime()
  With cmdStart
    .ButtonStyle = bsRed
    .Caption = "Stop"
  End With
  SetTimer
  IsRunning = True
End Sub

Private Sub EndCountdown()
  KillTimer
  cmbShutdown.Visible = True
  wStatus.Visible = False
  txtDelay.Visible = True
  flProgress.Visible = False
  With cmdStart
    .ButtonStyle = bsGreen
    .Caption = "Start"
  End With
  IsRunning = False
End Sub

Private Sub cmbShutdown_ButtonSelected(Button As Shutdown_Method)
  o_Method = Button
  CheckReadyButton
End Sub

Private Sub cmdStart_Click()
  If Not IsRunning Then
    StartCountdown
  Else
    EndCountdown
  End If
End Sub

Private Sub CountdownTimer_Timer()
  Dim rT As Time_Info, s As Long
  If IsRunningInIDE = True Then CountdownTimer.Enabled = False
  rT = CalculateRemainingTime()
  UpdateStatusTime rT
  s = CountTotalSeconds(rT)
  flProgress.Value = (flProgress.Max - s)
  If Now >= EndTime Then
    UnloadForm
    ShutDownNow
    Exit Sub
  End If
  If IsRunningInIDE = True Then If Not CountdownTimer Is Nothing Then CountdownTimer.Enabled = True
End Sub

Private Sub Form_Resize()
  If ChangedByCode Then Exit Sub
  ChangedByCode = True
  With FixedFormSize
    If Me.Width <> .X Then Me.Width = .X
    If Me.Height <> .Y Then Me.Height = .Y
  End With
  ChangedByCode = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If UnloadedByCode = True Then GoTo UnloadExit
  If IsRunning = True Then
    CountdownTimer.Enabled = False
    Select Case MsgBox("A shutdown has been scheduled." & vbCrLf & "Close the window and stop the countdown?", vbYesNo Or vbInformation Or vbMsgBoxSetForeground, APP_NAME)
      Case vbYes
        'do nothing...
      Case Else
        Cancel = 1
        CountdownTimer.Enabled = True
        Exit Sub
    End Select
  End If
UnloadExit:
  KillTimer
  SaveSettings
End Sub

Private Sub optShutdown_AlwaysOnTopClick()
  o_AlwaysOnTop = (optShutdown.AlwaysOnTop = True)
  WindowOnTop Me.hWnd, o_AlwaysOnTop
End Sub

Private Sub optShutdown_ForceExitClick()
  o_ForceExit = (optShutdown.ForceExit = True)
End Sub

Private Sub optShutdown_InstallUpdatesClick()
  o_InstallUpdates = (optShutdown.InstallUpdates = True)
End Sub

Private Sub txtDelay_ValueChanged()
  CollectDelayTime
  CheckReadyButton
End Sub
