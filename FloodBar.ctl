VERSION 5.00
Begin VB.UserControl FloodBar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "FloodBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PROPNAME_MIN As String = "Min"
Private Const PROPNAME_MAX As String = "Max"
Private Const PROPNAME_VALUE As String = "Value"

Private Const FLOOD_HEIGHT As Long = 15

Dim m_Min As Long
Dim m_Max As Long
Dim m_Value As Long

Dim WindowRect As RECT
Dim FloodRect As RECT

Public Property Get Min() As Long
  Min = m_Min
End Property
Public Property Get Max() As Long
  Max = m_Max
End Property
Public Property Get Value() As Long
  Value = m_Value
End Property
Public Property Let Min(New_Min As Long)
  If m_Min = New_Min Then Exit Property
  m_Min = New_Min
  If m_Value < m_Min Then m_Value = m_Min
  If m_Max <= m_Min Then m_Max = (m_Min + 1)
  UserControl.PropertyChanged PROPNAME_MIN
  Refresh
End Property
Public Property Let Max(New_Max As Long)
  If m_Max = New_Max Then Exit Property
  m_Max = New_Max
  If m_Value > m_Max Then m_Value = m_Max
  If m_Min > m_Max Then m_Min = (m_Max - 1)
  UserControl.PropertyChanged PROPNAME_MAX
  Refresh
End Property
Public Property Let Value(New_Value As Long)
  If m_Value = New_Value Then Exit Property
  m_Value = New_Value
  If m_Value > m_Max Then m_Value = m_Max
  If m_Value < m_Min Then m_Value = m_Min
  UserControl.PropertyChanged PROPNAME_VALUE
  Refresh
End Property

Public Sub Refresh(Optional FullRedraw As Boolean = False)
  If FullRedraw = True Then UserControl.Cls
  DrawFlood
  DrawBorder
End Sub

Private Sub DrawBorder()
  With WindowRect
    UserControl.Line ((.Left + (Screen.TwipsPerPixelX * 2)), (.Top + (Screen.TwipsPerPixelY * 2)))-((.Right - (Screen.TwipsPerPixelX * 2)), (.Bottom - (Screen.TwipsPerPixelY * 2))), COLOR_OUTLINE, B
    UserControl.Line ((.Left + (Screen.TwipsPerPixelX * 3)), (.Top + (Screen.TwipsPerPixelY * 3)))-((.Right - (Screen.TwipsPerPixelX * 3)), (.Bottom - (Screen.TwipsPerPixelY * 3))), COLOR_OUTLINE, B
  End With
End Sub

Private Sub DrawFlood()
  Dim w As Long
  With FloodRect
    w = GetFloodWidth()
    If w > 0 Then UserControl.Line (.Left, .Top)-((.Left + w), .Bottom), GetBarColor(w), BF
    If w < (.Right - .Left) Then UserControl.Line ((.Left + w), .Top)-(.Right, .Bottom), COLOR_CONTROLS, BF
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), COLOR_OUTLINE_LIGHT, B
  End With
End Sub

Private Function GetBarColor(Optional fWidth As Long = 0) As Long
  Dim w As Long
  If fWidth = 0 Then w = GetFloodWidth() Else w = fWidth
  Select Case (fWidth / (FloodRect.Right - FloodRect.Left))
    Case Is < 0.7
      GetBarColor = COLOR_GREEN
    Case Is < 0.9
      GetBarColor = COLOR_YELLOW
    Case Else
      GetBarColor = COLOR_RED
  End Select
End Function

Private Function GetFloodWidth() As Long
  Dim w As Long
  With FloodRect
    w = (m_Max - m_Min)
    If w = 0 Then GetFloodWidth = 0: Exit Function
    w = (((.Right - .Left) / w) * (m_Value - m_Min))
    If w > (.Right - .Left) Then w = (.Right - .Left)
  End With
  GetFloodWidth = w
End Function

Private Sub SetRects()
  SetWindowRect
  SetFloodRect
End Sub

Private Sub SetFloodRect()
  Dim h As Long
  h = (FLOOD_HEIGHT * Screen.TwipsPerPixelY)
  If (WindowRect.Bottom - WindowRect.Top) < h Then UserControl.Height = ((UserControl.Height - UserControl.ScaleHeight) + (h + (Screen.TwipsPerPixelY * 6))): Exit Sub
  With FloodRect
    .Top = (WindowRect.Top + (((WindowRect.Bottom - WindowRect.Top) - h) \ 2))
    .Bottom = (.Top + h)
    .Left = (WindowRect.Left + (Screen.TwipsPerPixelX * 10))
    .Right = (WindowRect.Right - (Screen.TwipsPerPixelX * 11))
  End With
End Sub

Private Sub SetWindowRect()
  With WindowRect
    .Left = 0
    .Top = 0
    .Right = UserControl.ScaleWidth
    .Bottom = UserControl.ScaleHeight
  End With
End Sub

Private Sub UserControl_InitProperties()
  m_Min = 0
  m_Max = 100
  m_Value = m_Min
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Min = PropBag.ReadProperty(PROPNAME_MIN, 0)
  m_Max = PropBag.ReadProperty(PROPNAME_MAX, 100)
  m_Value = PropBag.ReadProperty(PROPNAME_VALUE, 0)
End Sub

Private Sub UserControl_Resize()
  SetRects
  Refresh True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_MIN, m_Min, 0
  PropBag.WriteProperty PROPNAME_MAX, m_Max, 100
  PropBag.WriteProperty PROPNAME_VALUE, m_Value, 0
End Sub
