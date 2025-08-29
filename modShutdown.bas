Attribute VB_Name = "modShutdown"
Option Explicit

Private Const EWX_SHUTDOWN As Long = &H1
Private Const EWX_REBOOT As Long = &H2
Private Const EWX_FORCE As Long = &H4
Private Const EWX_POWEROFF As Long = &H8
Private Const EWX_LOGOFF As Long = &H0

Private Const EWX_HIBERNATE As Long = &H20 'Internal only

Private Const SHUTDOWN_INSTALL_UPDATES As Long = &H40
Private Const SHUTDOWN_RESTART As Long = &H4
Private Const SHUTDOWN_POWEROFF As Long = &H8

Private Const SHUTDOWN_FORCE_OTHERS = &H1
Private Const SHUTDOWN_FORCE_SELF = &H2

Private Const SHTDN_REASON_MAJOR_OPERATINGSYSTEM As Long = &H20000
Private Const SHTDN_REASON_MINOR_UPGRADE As Long = &H3

Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_QUERY As Long = &H8
Private Const SE_PRIVILEGE_ENABLED As Long = &H2

Public Enum Shutdown_Method
  smNone = &H0
  smShutdown = &H1
  smReboot = &H2
  smHibernate = &H4
  smLogOut = &H8
End Enum

Private Type LUID
  dwLowPart As Long
  dwHighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
  udtLUID As LUID
  dwAttributes As Long
End Type

Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  laa As LUID_AND_ATTRIBUTES
End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Long) As Long

Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long
Private Declare Function SetSuspendState Lib "PowrProf" (ByVal bHibernate As Long, ByVal bForce As Long, bWakeUpEventsDisabled) As Long

Private Declare Function InitiateShutdown Lib "advapi32.dll" Alias "InitiateShutdownW" (ByVal lpMachineName As Long, ByVal lpMessage As Long, ByVal dwGracePeriod As Long, ByVal dwShutdownFlags As Long, ByVal dwReason As Long) As Long

Public Sub ShutDownNow()
  Dim m As Long
  Call EnableShutdownPrivileges
  m = GetShutdownMethod(o_Method)
  If m = EWX_HIBERNATE Then
    Call SetSuspendState(True, 0, True)
    Exit Sub
  End If
  If o_InstallUpdates = True Then
    If (m = EWX_POWEROFF Or m = EWX_REBOOT) Then
      InstallAndShutdown m
      Exit Sub
    End If
  End If
  If o_ForceExit Then m = m Or EWX_FORCE
  Call ExitWindowsEx(m, 0&)
End Sub

Private Function GetShutdownMethod(ShutdownMethod As Shutdown_Method) As Long
  If (ShutdownMethod And smShutdown) Then
    GetShutdownMethod = EWX_POWEROFF
  ElseIf (ShutdownMethod And smReboot) Then
    GetShutdownMethod = EWX_REBOOT
  ElseIf (ShutdownMethod And smHibernate) Then
    GetShutdownMethod = EWX_HIBERNATE
  ElseIf (ShutdownMethod And smLogOut) Then
    GetShutdownMethod = EWX_LOGOFF
  End If
End Function

Private Sub InstallAndShutdown(ShutdownMethod As Long)
  Dim m As Long
  m = SHUTDOWN_INSTALL_UPDATES
  If (ShutdownMethod And EWX_POWEROFF) Then
    m = m Or SHUTDOWN_POWEROFF
  ElseIf (ShutdownMethod And EWX_REBOOT) Then
    m = m Or SHUTDOWN_RESTART
  End If
  If o_ForceExit Then m = m Or SHUTDOWN_FORCE_SELF Or SHUTDOWN_FORCE_OTHERS
  Call InitiateShutdown(0&, 0&, 0&, m, SHTDN_REASON_MAJOR_OPERATINGSYSTEM Or SHTDN_REASON_MAJOR_OPERATINGSYSTEM)
End Sub

Private Function EnableShutdownPrivileges() As Boolean
  Dim hProcessHandle As Long
  Dim hTokenHandle As Long
  Dim lpv_la As LUID
  Dim Token As TOKEN_PRIVILEGES
  hProcessHandle = GetCurrentProcess
  If Not hProcessHandle = 0 Then
    If Not OpenProcessToken(hProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hTokenHandle) = 0 Then
      If Not LookupPrivilegeValue(vbNullString, "SeShutdownPrivilege", lpv_la) = 0 Then
        With Token
          .PrivilegeCount = 1
          With .laa
            .udtLUID = lpv_la
            .dwAttributes = SE_PRIVILEGE_ENABLED
          End With
        End With
        If Not AdjustTokenPrivileges(hTokenHandle, False, Token, ByVal 0&, ByVal 0&, ByVal 0&) = 0 Then
          EnableShutdownPrivileges = True
        End If
      End If
    End If
  End If
End Function

Public Function IsRebootRequired() As Boolean
  On Error Resume Next
  Dim shell As Object
  Set shell = CreateObject("WScript.Shell")
  
  If shell Is Nothing Then Exit Function
  
  Dim keys
  keys = Array( _
    "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending\", _
    "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired\", _
    "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations", _
    "HKLM\SOFTWARE\Microsoft\Updates\UpdateExeVolatile", _
    "HKLM\SYSTEM\CurrentControlSet\Services\Netlogon\JoinDomain", _
    "HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName\ComputerName", _
    "HKLM\SYSTEM\CurrentControlSet\Control\Windows\SystemRestore\RPRebootRequired" _
  )
  Dim i As Long
  For i = 0 To UBound(keys)
    Err.Clear
    shell.RegRead keys(i)
    If Err.Number = 0 Then
      IsRebootRequired = True
      Exit For
    End If
  Next
  On Error GoTo 0
End Function
