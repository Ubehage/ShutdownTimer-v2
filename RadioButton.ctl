VERSION 5.00
Begin VB.UserControl RadioButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "RadioButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PROPNAME_VALUE As String = "Value"
Private Const PROPNAME_CAPTION As String = "Caption"
Private Const PROPNAME_ENABLED As String = "Enabled"
Private Const PROPNAME_AUTOSIZE As String = "AutoSize"

Private Const BUTTON_RADIO_SIZE As Long = 18

Private Const COLOR_RADIO_BACKGROUND As Long = COLOR_CONTROLS
Private Const COLOR_RADIO_INSIDE As Long = COLOR_BACKGROUND
Private Const COLOR_RADIO_SELECTED As Long = COLOR_GREEN
Private Const COLOR_RADIO_HOVER As Long = COLOR_BUTTON_HOVER
Private Const COLOR_RADIO_PRESSED As Long = COLOR_BUTTON_PRESSED

Dim m_Value As Boolean
Dim m_Caption As String
Dim m_Enabled As Boolean
Dim m_AutoSize As Boolean

Dim ButtonRect As RECT
Dim TextPos As RECT
Dim CirclePos As Circle_Pos

Dim m_ScreenRect As RECT
Dim m_IsCapturing As Boolean
Dim m_Hovering As Boolean
Dim m_IsPressed As Boolean
Dim m_MouseIsDown As Boolean
Dim m_KeyIsDown As Boolean
Dim m_PreviewSelect As Boolean
Dim m_HasFocus As Boolean

Dim DontChangeVal As Boolean

Public Event Click()

Public Property Get Value() As Boolean
  Value = m_Value
End Property
Public Property Get Caption() As String
  Caption = m_Caption
End Property
Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property
Public Property Get ContainerhWnd() As Long
  ContainerhWnd = UserControl.ContainerhWnd
End Property
Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property
Public Property Get AutoSize() As Boolean
  AutoSize = m_AutoSize
End Property
Public Property Let Value(New_Value As Boolean)
  If DontChangeVal = True Then Exit Property
  If m_Value = New_Value Then Exit Property
  m_Value = New_Value
  UserControl.PropertyChanged PROPNAME_VALUE
  If m_Value = True Then TurnOffSiblings
  Refresh
End Property
Public Property Let Caption(New_Caption As String)
  If m_Caption = New_Caption Then Exit Property
  m_Caption = New_Caption
  UserControl.PropertyChanged PROPNAME_CAPTION
  SetTextPosition
  Refresh
End Property
Public Property Let Enabled(New_Enabled As Boolean)
  If m_Enabled = New_Enabled Then Exit Property
  m_Enabled = New_Enabled
  UserControl.PropertyChanged PROPNAME_ENABLED
  If m_Enabled = False Then
    EndHover
    m_MouseIsDown = False
    m_KeyIsDown = False
    m_IsPressed = False
  End If
  Refresh
End Property
Public Property Let AutoSize(New_AutoSize As Boolean)
  If m_AutoSize = New_AutoSize Then Exit Property
  m_AutoSize = New_AutoSize
  If m_AutoSize = True Then Refresh
End Property

Public Sub Refresh()
  UserControl.Line (0, 0)-(UserControl.ScaleWidth, UserControl.ScaleHeight), COLOR_BACKGROUND, BF
  SetTextPosition
  DrawButton
  DrawText
  If m_HasFocus Then DrawFocusBorder
End Sub

Private Sub DrawFocusBorder()
  UserControl.Line (0, 0)-((UserControl.ScaleWidth - Screen.TwipsPerPixelX), (UserControl.ScaleHeight - Screen.TwipsPerPixelY)), COLOR_OUTLINE_LIGHT, B
End Sub

Private Function GetButtonBackColor() As Long
  If m_Enabled = False Then
    GetButtonBackColor = COLOR_BACKGROUND_DISABLED
  ElseIf m_IsPressed Then
    GetButtonBackColor = COLOR_RADIO_PRESSED
  ElseIf m_Hovering = True Then
    GetButtonBackColor = COLOR_RADIO_HOVER
  Else
    GetButtonBackColor = COLOR_RADIO_INSIDE
  End If
End Function

Private Function GetButtonDotColor() As Long
  If m_Enabled = False Then
    GetButtonDotColor = COLOR_GREEN_DISABLED
  ElseIf m_IsPressed Then
    GetButtonDotColor = COLOR_GREEN_PRESSED
  ElseIf m_Hovering Then
    GetButtonDotColor = COLOR_GREEN_HOVER
  Else
    GetButtonDotColor = COLOR_RADIO_SELECTED
  End If
End Function

Private Function GetTextColor() As Long
  If m_Enabled = True Then GetTextColor = COLOR_TEXT Else GetTextColor = COLOR_TEXT_DISABLED
End Function

Private Sub DrawButton()
  Dim r As Long
  With CirclePos
    If .Radius = 0 Then Exit Sub
    UserControl.FillColor = GetButtonBackColor()
    UserControl.FillStyle = 0
    UserControl.Circle (.X, .Y), .Radius
    UserControl.FillStyle = 1
    r = (.Radius - Screen.TwipsPerPixelX)
    UserControl.Circle (.X, .Y), r, COLOR_OUTLINE_LIGHT
    UserControl.Circle (.X, .Y), (r - Screen.TwipsPerPixelX), COLOR_OUTLINE_LIGHT
    If (m_Value = True Or m_PreviewSelect = True) Then
      r = (.Radius - (Screen.TwipsPerPixelX * 4))
      UserControl.FillColor = GetButtonDotColor
      UserControl.FillStyle = 0
      UserControl.Circle (.X, .Y), r
      UserControl.FillStyle = 1
    End If
  End With
End Sub

Private Sub DrawText()
  With TextPos
    UserControl.CurrentX = .Left
    UserControl.CurrentY = .Top
  End With
  UserControl.ForeColor = GetTextColor()
  UserControl.Print m_Caption
End Sub

Private Sub SetButtonRect()
  With ButtonRect
    .Left = ((UserControl.ScaleWidth - (BUTTON_RADIO_SIZE * Screen.TwipsPerPixelX)) \ 2)
    .Top = 30
    .Right = (.Left + (BUTTON_RADIO_SIZE * Screen.TwipsPerPixelX))
    .Bottom = (.Top + (BUTTON_RADIO_SIZE * Screen.TwipsPerPixelY))
    CirclePos.X = (.Left + ((.Right - .Left) \ 2))
    CirclePos.Y = (.Top + ((.Bottom - .Top) \ 2))
  End With
  CirclePos.Radius = ((BUTTON_RADIO_SIZE / 2) * Screen.TwipsPerPixelX)
End Sub

Private Sub SetTextPosition()
  With TextPos
    .Right = UserControl.TextWidth(m_Caption)
    If m_AutoSize = True Then
      If (.Right + Screen.TwipsPerPixelX) <> UserControl.ScaleWidth Then
        UserControl.Width = ((UserControl.Width - UserControl.ScaleWidth) + (.Right + Screen.TwipsPerPixelX))
        Exit Sub
      End If
    End If
    .Left = ((UserControl.ScaleWidth - .Right) \ 2)
    .Top = (ButtonRect.Bottom + (Screen.TwipsPerPixelY * 3))
    .Right = (.Left + .Right)
    .Bottom = (.Top + UserControl.TextHeight(m_Caption))
    UserControl.Height = ((UserControl.Height - UserControl.ScaleHeight) + (.Bottom + (Screen.TwipsPerPixelX * 3)))
  End With
End Sub

Private Sub TurnOffSiblings()
  Dim c As Object
  DontChangeVal = True
  For Each c In UserControl.Parent.Controls
    If TypeOf c Is RadioButton Then
      If c.ContainerhWnd = ContainerhWnd Then
        If Not c.hWnd = hWnd Then
          c.Value = False
        End If
      End If
    End If
  Next
  DontChangeVal = False
End Sub

Private Sub SetScreenRect()
  Dim r As RECT, p As POINTAPI
  Call GetClientRect(UserControl.hWnd, r)
  p.X = r.Left
  p.Y = r.Top
  Call ClientToScreen(UserControl.hWnd, p)
  With m_ScreenRect
    .Left = p.X
    .Top = p.Y
    .Right = (.Left + r.Right)
    .Bottom = (.Top + r.Bottom)
  End With
End Sub

Private Sub StartHover(Optional DoNotRefresh As Boolean = False)
  If m_Hovering = False Then
    m_Hovering = True
    SetScreenRect
  End If
  If DoNotRefresh = False Then Refresh
  If m_IsCapturing = True Then Exit Sub
  Call SetCapture(UserControl.hWnd)
  m_IsCapturing = True
End Sub

Private Sub EndHover(Optional DoNotRefresh As Boolean = False)
  If m_Hovering = True Then m_Hovering = False
  If DoNotRefresh = False Then Refresh
  If m_IsCapturing = False Or m_MouseIsDown = True Then Exit Sub
  EndCapture
End Sub

Private Sub EndCapture()
  If m_IsCapturing Then
    Call ReleaseCapture
    m_IsCapturing = False
  End If
End Sub

Private Function IsCursorOnButton() As Boolean
  Dim p As POINTAPI, hTop As Long
  Call GetCursorPos(p)
  If IsPointInRect(m_ScreenRect, p) = False Then If m_MouseIsDown = False Then Exit Function
  hTop = WindowFromPoint(p.X, p.Y)
  If hTop <> UserControl.hWnd Then If m_MouseIsDown = False Then Exit Function
  IsCursorOnButton = True
End Function

Private Sub DoTheClick()
  If m_Value = True Then Exit Sub
  Value = True
  RaiseEvent Click
End Sub

Private Function CanInteractNow() As Boolean
  CanInteractNow = UserControl.Ambient.UserMode And m_Enabled And UserControl.Extender.Visible And UserControl.hWnd <> 0
End Function

Private Sub UserControl_GotFocus()
  m_HasFocus = True
  Refresh
End Sub

Private Sub UserControl_Initialize()
  With UserControl.Font
    .Name = FONT_SECONDARY
    .Size = FONTSIZE_MAIN
  End With
  Refresh
End Sub

Private Sub UserControl_InitProperties()
  m_Value = False
  m_Caption = "Button"
  m_Enabled = True
  m_AutoSize = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  If m_Enabled = False Then Exit Sub
  If KeyCode = vbKeySpace Then
    m_IsPressed = True
    m_KeyIsDown = True
    Refresh
  End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  If m_Enabled = False Then Exit Sub
  If KeyCode = vbKeySpace Then
    If m_IsPressed Then
      If m_KeyIsDown = True Then
        m_KeyIsDown = False
        m_IsPressed = False
        Refresh
        DoTheClick
      End If
    ElseIf m_KeyIsDown = True Then
      m_KeyIsDown = False
    End If
  End If
End Sub

Private Sub UserControl_LostFocus()
  m_HasFocus = False
  Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If m_Enabled = False Then Exit Sub
  If (Button = vbLeftButton And m_MouseIsDown = False) Then
    m_MouseIsDown = True
    m_IsPressed = True
    If m_Value = False Then m_PreviewSelect = True
    StartHover
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If m_Enabled = False Then Exit Sub
  StartHover
  If m_IsCapturing = False Then Exit Sub
  If IsCursorOnButton() = False Then EndHover
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim HadCapture As Boolean
  If m_Enabled = False Then Exit Sub
  If (Button = vbLeftButton And m_MouseIsDown = True) Then
    HadCapture = m_IsCapturing
    m_MouseIsDown = False
    m_IsPressed = False
    m_PreviewSelect = False
    EndHover True
    If IsCursorOnButton() Then
      DoTheClick
      If CanInteractNow() = True Then If HadCapture = True Then StartHover True
    End If
    Refresh
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Value = PropBag.ReadProperty(PROPNAME_VALUE, False)
  m_Caption = PropBag.ReadProperty(PROPNAME_CAPTION, "Button")
  m_Enabled = PropBag.ReadProperty(PROPNAME_ENABLED, True)
  m_AutoSize = PropBag.ReadProperty(PROPNAME_AUTOSIZE, False)
  Refresh
End Sub

Private Sub UserControl_Resize()
  SetButtonRect
  SetTextPosition
  Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_VALUE, m_Value, False
  PropBag.WriteProperty PROPNAME_CAPTION, m_Caption, "Button"
  PropBag.WriteProperty PROPNAME_ENABLED, m_Enabled, True
  PropBag.WriteProperty PROPNAME_AUTOSIZE, m_AutoSize, False
End Sub
