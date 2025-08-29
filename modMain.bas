Attribute VB_Name = "modMain"
Option Explicit

Global Const APP_NAME = "Ubehage's ShutdownTimer v2"

Global Const FONT_MAIN As String = "Segoe UI"
Global Const FONT_SECONDARY As String = "Consolas"
Global Const FONTSIZE_MAIN As Integer = 11
Global Const FONTSIZE_SECONDARY As Integer = 9

Global Const COLOR_BACKGROUND As Long = 2105376
Global Const COLOR_CONTROLS As Long = 2763306
Global Const COLOR_BUTTON_HOVER As Long = 3684408
Global Const COLOR_BUTTON_PRESSED As Long = 3289650
Global Const COLOR_BACKGROUND_DISABLED As Long = 5263440

Global Const COLOR_TEXT As Long = 14737632
Global Const COLOR_TEXT_HOVER As Long = 15790320
Global Const COLOR_TEXT_DISABLED As Long = 7895160
Global Const COLOR_TEXT_ONGREEN As Long = 15463654
Global Const COLOR_TEXT_ONRED As Long = 15395579
Global Const COLOR_TEXT_DISABLED_ONGREEN As Long = 13355947
Global Const COLOR_TEXT_DISABLED_ONRED As Long = 10592542

Global Const COLOR_GREEN As Long = 5023791
Global Const COLOR_GREEN_HOVER As Long = 6339651
Global Const COLOR_GREEN_PRESSED As Long = 4033061
Global Const COLOR_GREEN_DISABLED As Long = 6455130
Global Const COLOR_YELLOW As Long = 4965861
Global Const COLOR_RED As Long = 4539862
Global Const COLOR_RED_HOVER As Long = 6513642
Global Const COLOR_RED_PRESSED As Long = 3223992
Global Const COLOR_RED_DISABLED As Long = 6776730
Global Const COLOR_OUTLINE As Long = 3815994
Global Const COLOR_OUTLINE_LIGHT As Long = 7368816

Private Const SETTINGS_APPNAME As String = "UbeSDTimer2"
Private Const SETTINGS_SECTION As String = "Settings"
Private Const SETTINGS_FORCE As String = "ForceExit"
Private Const SETTINGS_UPDATES As String = "InstallUpdates"
Private Const SETTINGS_ONTOP As String = "AlwaysOnTop"
Private Const SETTINGS_POS_X As String = "PosX"
Private Const SETTINGS_POS_Y As String = "PosY"
Private Const SETTINGS_METHOD As String = "ShutdownMethod"
Private Const SETTINGS_TIME_HOURS As String = "Hours"
Private Const SETTINGS_TIME_MINUTES As String = "Minutes"
Private Const SETTINGS_TIME_SECONDS As String = "Seconds"

Enum Font_Defaults
  fdMain = &H1
  fdControls = &H2
End Enum

Public Type Time_Info
  Hours As Long
  Minutes As Long
  Seconds As Long
End Type

Global ChangedByCode As Boolean

Global DelayTime As Time_Info
Global EndTime As Date

Global o_ForceExit As Boolean
Global o_InstallUpdates As Boolean
Global o_AlwaysOnTop As Boolean
Global o_Pos As POINTAPI
Global o_Method As Shutdown_Method
Global o_Time As Time_Info

Global IsRunningInIDE As Boolean

Global UnloadedByCode As Boolean

Sub Main()
  InitCommonControls
  IsRunningInIDE = IsInIDE()
  If App.PrevInstance = True Then
    MsgBox "Error:" & vbCrLf & "Another instance of this app is already running!", vbOKOnly Or vbInformation, APP_NAME
    Exit Sub
  End If
  LoadForm
End Sub

Private Sub LoadForm()
  Load frmMain
  frmMain.SetForm
End Sub

Public Sub UnloadForm()
  UnloadedByCode = True
  Unload frmMain
  UnloadedByCode = False
  Set frmMain = Nothing
End Sub

Public Sub LoadSettings()
  o_ForceExit = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_FORCE, False)
  o_InstallUpdates = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_UPDATES, False)
  o_AlwaysOnTop = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_ONTOP, True)
  With o_Pos
    .X = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_POS_X, frmMain.Left)
    .Y = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_POS_Y, frmMain.Top)
  End With
  o_Method = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_METHOD, Shutdown_Method.smNone)
  With o_Time
    .Hours = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_TIME_HOURS, 0)
    .Minutes = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_TIME_MINUTES, 0)
    .Seconds = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_TIME_SECONDS, 0)
  End With
End Sub

Public Sub SaveSettings()
  SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_FORCE, o_ForceExit
  SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_UPDATES, o_InstallUpdates
  SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_ONTOP, o_AlwaysOnTop
  With frmMain
    SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_POS_X, .Left
    SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_POS_Y, .Top
  End With
  SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_METHOD, o_Method
  With DelayTime
    SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_TIME_HOURS, .Hours
    SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_TIME_MINUTES, .Minutes
    SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_TIME_SECONDS, .Seconds
  End With
End Sub

Public Function IsInIDE() As Boolean
  Dim inIDE As Boolean
  inIDE = False
  On Error Resume Next
  Debug.Assert MakeIDECheck(inIDE)
  On Error GoTo 0
  IsInIDE = inIDE
End Function

Private Function MakeIDECheck(bSet As Boolean) As Boolean
  bSet = True
  MakeIDECheck = True
End Function
