VERSION 5.00
Begin VB.UserControl ShutdownSelector 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4260
   ScaleWidth      =   10095
   Begin ShutdownTimer2.RadioButton rbLogout 
      Height          =   660
      Left            =   7635
      TabIndex        =   3
      Top             =   705
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1164
      Caption         =   "rbLogout"
      AutoSize        =   -1  'True
   End
   Begin ShutdownTimer2.RadioButton rbHibernate 
      Height          =   660
      Left            =   5460
      TabIndex        =   2
      Top             =   705
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1164
      Caption         =   "rbHibernate"
      AutoSize        =   -1  'True
   End
   Begin ShutdownTimer2.RadioButton rbReboot 
      Height          =   660
      Left            =   3135
      TabIndex        =   1
      Top             =   660
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1164
      Caption         =   "rbReboot"
      AutoSize        =   -1  'True
   End
   Begin ShutdownTimer2.RadioButton rbShutdown 
      Height          =   660
      Left            =   420
      TabIndex        =   0
      Top             =   645
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1164
      Caption         =   "rbShutdown"
      AutoSize        =   -1  'True
   End
End
Attribute VB_Name = "ShutdownSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PROPNAME_SHUTDOWN As String = "AllowShutdown"
Private Const PROPNAME_REBOOT As String = "AllowReboot"
Private Const PROPNAME_LOGOUT As String = "AllowLogout"
Private Const PROPNAME_HIBERNATE As String = "AllowHibernate"
Private Const PROPNAME_TITLE As String = "Title"

Private Const OPT_TEXT_SHUTDOWN As String = "Shut Down"
Private Const OPT_TEXT_REBOOT As String = "Restart"
Private Const OPT_TEXT_LOGOUT As String = "Log Out"
Private Const OPT_TEXT_HIBERNATE As String = "Hibernate"

Private Const BUTTON_SEP_PIXELS = 3

Dim m_AllowShutdown As Boolean
Dim m_AllowReboot As Boolean
Dim m_AllowLogout As Boolean
Dim m_AllowHibernate As Boolean
Dim m_Title As String

Dim WindowRect As RECT
Dim ButtonsRect As RECT

Dim ButtonsChanged As Boolean
Dim LoadedButtons As Integer

Dim m_HasFocus As Boolean

Public Event ButtonSelected(Button As Shutdown_Method)

Public Property Get AllowShutdown() As Boolean
  AllowShutdown = m_AllowShutdown
End Property
Public Property Get AllowReboot() As Boolean
  AllowReboot = m_AllowReboot
End Property
Public Property Get AllowLogout() As Boolean
  AllowLogout = m_AllowLogout
End Property
Public Property Get AllowHibernate() As Boolean
  AllowHibernate = m_AllowHibernate
End Property
Public Property Get Controls() As Object
  Set Controls = UserControl.Controls
End Property
Public Property Get SelectedShutdownMethod() As Shutdown_Method
  SelectedShutdownMethod = GetSelectedShutdownMethod
End Property
Public Property Get Title() As String
  Title = m_Title
End Property
Public Property Let AllowShutdown(New_AllowShutdown As Boolean)
  If m_AllowShutdown = New_AllowShutdown Then Exit Property
  m_AllowShutdown = New_AllowShutdown
  UserControl.PropertyChanged PROPNAME_SHUTDOWN
  ButtonsChanged = True
  Refresh
End Property
Public Property Let AllowReboot(New_AllowReboot As Boolean)
  If m_AllowReboot = New_AllowReboot Then Exit Property
  m_AllowReboot = New_AllowReboot
  UserControl.PropertyChanged PROPNAME_REBOOT
  ButtonsChanged = True
  Refresh
End Property
Public Property Let AllowLogout(New_AllowLogout As Boolean)
  If m_AllowLogout = New_AllowLogout Then Exit Property
  m_AllowLogout = New_AllowLogout
  UserControl.PropertyChanged PROPNAME_LOGOUT
  ButtonsChanged = True
  Refresh
End Property
Public Property Let AllowHibernate(New_AllowHibernate As Boolean)
  If m_AllowHibernate = New_AllowHibernate Then Exit Property
  m_AllowHibernate = New_AllowHibernate
  UserControl.PropertyChanged PROPNAME_HIBERNATE
  ButtonsChanged = True
  Refresh
End Property
Public Property Let SelectedShutdownMethod(New_Method As Shutdown_Method)
  If GetSelectedShutdownMethod = New_Method Then Exit Property
  If New_Method = smShutdown Then
    rbShutdown.Value = True
  ElseIf New_Method = smReboot Then
    rbReboot.Value = True
  ElseIf New_Method = smHibernate Then
    rbHibernate.Value = True
  ElseIf New_Method = smLogOut Then
    rbLogout.Value = True
  Else
    rbShutdown.Value = False
    rbReboot.Value = False
    rbHibernate.Value = False
    rbLogout.Value = False
  End If
End Property
Public Property Let Title(New_Title As String)
  If m_Title = New_Title Then Exit Property
  m_Title = New_Title
  UserControl.PropertyChanged PROPNAME_TITLE
  Refresh True
End Property

Public Sub Refresh(Optional FullRefresh As Boolean = False)
  UserControl.Cls
  If (ButtonsChanged = True Or FullRefresh = True) Then
    MoveShutdownButtons
    ButtonsChanged = False
  End If
  RefreshButtons
  DrawBorder
  DrawTitle
End Sub

Private Sub RefreshButtons()
  Dim c As Object
  For Each c In UserControl.Controls
    If c.Visible = True Then c.Refresh
  Next
End Sub

Private Sub DrawBorder()
  Dim cX As Long, cY As Long, cX2 As Long, cY2 As Long
  cX = 0
  cY = (Screen.TwipsPerPixelY * 10)
  cX2 = (UserControl.ScaleWidth - Screen.TwipsPerPixelX)
  cY2 = (UserControl.ScaleHeight - Screen.TwipsPerPixelY)
  UserControl.Line (cX, cY)-(cX2, cY2), COLOR_OUTLINE, B
  UserControl.Line ((cX + Screen.TwipsPerPixelX), (cY + Screen.TwipsPerPixelY))-((cX2 - Screen.TwipsPerPixelX), (cY2 - Screen.TwipsPerPixelY)), COLOR_OUTLINE, B
  With WindowRect
    .Left = cX
    .Top = cY
  End With
End Sub

Private Sub DrawTitle()
  UserControl.CurrentX = (Screen.TwipsPerPixelX * 10)
  UserControl.CurrentY = 15
  UserControl.ForeColor = COLOR_TEXT
  UserControl.Print m_Title
  ButtonsRect.Top = (UserControl.CurrentY + (Screen.TwipsPerPixelY * 7))
End Sub

Private Sub MoveShutdownButtons()
  Dim cX As Long, b As Integer, s As Long
  With ButtonsRect
    .Left = (WindowRect.Left + (Screen.TwipsPerPixelX * (BUTTON_SEP_PIXELS * 3)))
    '.Top = (WindowRect.Top + (Screen.TwipsPerPixelY * (BUTTON_SEP_PIXELS * 3)))
    .Right = (WindowRect.Right - (Screen.TwipsPerPixelX * BUTTON_SEP_PIXELS))
    .Bottom = (.Top + rbShutdown.Height)
    cX = .Left
    b = GetActiveButtonsCount()
    If b > 1 Then
      s = (((.Right - .Left) - GetAllButtonsWidth) \ (b - 1))
    ElseIf b = 1 Then
      s = (Screen.TwipsPerPixelX * BUTTON_SEP_PIXELS)
    End If
    MoveShutdownButton rbShutdown, m_AllowShutdown, cX, s, b
    MoveShutdownButton rbReboot, m_AllowReboot, cX, s, b
    MoveShutdownButton rbHibernate, m_AllowHibernate, cX, s, b
    MoveShutdownButton rbLogout, m_AllowLogout, cX, s, b
    WindowRect.Bottom = (.Bottom + (Screen.TwipsPerPixelY * (BUTTON_SEP_PIXELS * 3)))
    .Right = cX
    Call CheckWindowSize
  End With
End Sub

Private Sub MoveShutdownButton(ByRef sButton As Object, ByVal sAllow As Boolean, ByRef xPos As Long, ByVal pSpacer As Long, ByRef bCounter As Integer)
  With sButton
    If sAllow = True Then
      .Visible = True
      .Move xPos, ButtonsRect.Top
      If bCounter > 1 Then
        xPos = (xPos + (.Width + pSpacer))
      ElseIf bCounter = 1 Then
        xPos = (xPos + (Screen.TwipsPerPixelX * BUTTON_SEP_PIXELS))
      End If
      bCounter = (bCounter - 1)
    Else
      .Visible = False
    End If
  End With
End Sub

Private Function GetMinimumWindowWidth() As Long
  If GetActiveButtonsCount = 0 Then
    GetMinimumWindowWidth = (Screen.TwipsPerPixelX * (BUTTON_SEP_PIXELS * 25))
  Else
    GetMinimumWindowWidth = (GetAllButtonsWidth + (Screen.TwipsPerPixelX * (BUTTON_SEP_PIXELS * 2)))
  End If
End Function

Private Function GetWindowHeight() As Long
  GetWindowHeight = (rbShutdown.Height + (rbShutdown.Top * 2))
End Function

Private Function GetAllButtonsWidth() As Long
  Dim r As Long
  If m_AllowShutdown Then r = (r + rbShutdown.Width)
  If m_AllowReboot Then r = (r + rbReboot.Width)
  If m_AllowHibernate Then r = (r + rbHibernate.Width)
  If m_AllowLogout Then r = (r + rbLogout.Width)
  GetAllButtonsWidth = r
End Function

Private Function GetActiveButtonsCount() As Integer
  Dim r As Integer
  If m_AllowShutdown Then r = (r + 1)
  If m_AllowReboot Then r = (r + 1)
  If m_AllowHibernate Then r = (r + 1)
  If m_AllowLogout Then r = (r + 1)
  GetActiveButtonsCount = r
End Function

Private Function GetSelectedShutdownMethod() As Shutdown_Method
  If rbShutdown.Value = True Then
    GetSelectedShutdownMethod = smShutdown
  ElseIf rbReboot.Value = True Then
    GetSelectedShutdownMethod = smReboot
  ElseIf rbHibernate.Value = True Then
    GetSelectedShutdownMethod = smHibernate
  ElseIf rbLogout.Value = True Then
    GetSelectedShutdownMethod = smLogOut
  End If
End Function

Private Function CheckWindowSize() As Boolean
  CheckWindowSize = (CheckWindowWidth() = True And CheckWindowHeight() = True)
End Function

Private Function CheckWindowWidth() As Boolean
  If UserControl.ScaleWidth >= GetMinimumWindowWidth() Then
    CheckWindowWidth = True
    Exit Function
  End If
  UserControl.Width = ((UserControl.Width - UserControl.ScaleWidth) + GetMinimumWindowWidth())
End Function

Private Function CheckWindowHeight() As Boolean
  If UserControl.ScaleHeight = GetWindowHeight() Then
    CheckWindowHeight = True
    Exit Function
  End If
  UserControl.Height = ((UserControl.Height - UserControl.ScaleHeight) + GetWindowHeight)
End Function

Private Sub SetWindowRect()
  With WindowRect
    .Left = 0
    .Top = 0
    .Right = UserControl.ScaleWidth
    .Bottom = UserControl.ScaleHeight
  End With
End Sub

Private Sub rbHibernate_Click()
  RaiseEvent ButtonSelected(smHibernate)
End Sub

Private Sub rbLogout_Click()
  RaiseEvent ButtonSelected(smLogOut)
End Sub

Private Sub rbReboot_Click()
  RaiseEvent ButtonSelected(smReboot)
End Sub

Private Sub rbShutdown_Click()
  RaiseEvent ButtonSelected(smShutdown)
End Sub

Private Sub UserControl_GotFocus()
  m_HasFocus = True
  Refresh
End Sub

Private Sub UserControl_Initialize()
  ButtonsChanged = True
  rbShutdown.Caption = OPT_TEXT_SHUTDOWN
  rbReboot.Caption = OPT_TEXT_REBOOT
  rbHibernate.Caption = OPT_TEXT_HIBERNATE
  rbLogout.Caption = OPT_TEXT_LOGOUT
End Sub

Private Sub UserControl_InitProperties()
  m_AllowShutdown = True
  m_AllowReboot = True
  m_AllowHibernate = True
  m_AllowLogout = True
End Sub

Private Sub UserControl_LostFocus()
  m_HasFocus = False
  Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_AllowShutdown = PropBag.ReadProperty(PROPNAME_SHUTDOWN, True)
  m_AllowReboot = PropBag.ReadProperty(PROPNAME_REBOOT, True)
  m_AllowHibernate = PropBag.ReadProperty(PROPNAME_HIBERNATE, True)
  m_AllowLogout = PropBag.ReadProperty(PROPNAME_LOGOUT, True)
  m_Title = PropBag.ReadProperty(PROPNAME_TITLE, "Title")
End Sub

Private Sub UserControl_Resize()
  If CheckWindowSize() = True Then SetWindowRect: Refresh True
End Sub

Private Sub UserControl_Show()
  Refresh True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_SHUTDOWN, m_AllowShutdown, True
  PropBag.WriteProperty PROPNAME_REBOOT, m_AllowReboot, True
  PropBag.WriteProperty PROPNAME_HIBERNATE, m_AllowHibernate, True
  PropBag.WriteProperty PROPNAME_LOGOUT, m_AllowLogout, True
  PropBag.WriteProperty PROPNAME_TITLE, m_Title, "Title"
End Sub
