VERSION 5.00
Begin VB.UserControl DelaySelector 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
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
   ScaleHeight     =   3540
   ScaleWidth      =   6105
   Begin VB.TextBox txtVal 
      BackColor       =   &H002A2A2A&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   2
      Left            =   5130
      TabIndex        =   5
      Top             =   1110
      Width           =   510
   End
   Begin VB.TextBox txtVal 
      BackColor       =   &H002A2A2A&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   1
      Left            =   3255
      TabIndex        =   4
      Top             =   1065
      Width           =   510
   End
   Begin VB.TextBox txtVal 
      BackColor       =   &H002A2A2A&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   0
      Left            =   1110
      TabIndex        =   3
      Top             =   1065
      Width           =   510
   End
   Begin VB.Label lblSeconds 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   4245
      TabIndex        =   2
      Top             =   1095
      Width           =   840
   End
   Begin VB.Label lblMinutes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   2235
      TabIndex        =   1
      Top             =   1065
      Width           =   840
   End
   Begin VB.Label lblHours 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hours:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   360
      TabIndex        =   0
      Top             =   1065
      Width           =   630
   End
End
Attribute VB_Name = "DelaySelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PROPNAME_ENABLED As String = "Enabled"
Private Const PROPNAME_TITLE As String = "Title"

Private Const BUTTON_MIN_WIDTH = 1215

Dim m_Enabled As Boolean
Dim m_Title As String

Dim BorderRect As RECT

Dim DontResize As Boolean

Dim LastGoodVal(2) As String

Public Event ValueChanged()

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property
Public Property Get Title() As String
  Title = m_Title
End Property
Public Property Get Hours() As Integer
  Hours = CInt(Val(txtVal(0).Text))
End Property
Public Property Get Minutes() As Integer
  Minutes = CInt(Val(txtVal(1).Text))
End Property
Public Property Get Seconds() As Integer
  Seconds = CInt(Val(txtVal(2).Text))
End Property
Public Property Get TotalSeconds() As Long
  TotalSeconds = ((((Hours * 60) + Minutes) * 60) + Seconds)
End Property
Public Property Let Enabled(New_Enabled As Boolean)
  If m_Enabled = New_Enabled Then Exit Property
  m_Enabled = New_Enabled
  UserControl.PropertyChanged PROPNAME_ENABLED
  SetEnabled
  Refresh
End Property
Public Property Let Title(New_Title As String)
  If m_Title = New_Title Then Exit Property
  m_Title = New_Title
  UserControl.PropertyChanged PROPNAME_TITLE
  Refresh
End Property
Public Property Let Hours(New_Hours As Integer)
  If Hours = New_Hours Then Exit Property
  ChangedByCode = True
  txtVal(0).Text = CStr(New_Hours)
  ChangedByCode = False
End Property
Public Property Let Minutes(New_Minutes As Integer)
  If Minutes = New_Minutes Then Exit Property
  ChangedByCode = True
  txtVal(1).Text = CStr(New_Minutes)
  ChangedByCode = False
End Property
Public Property Let Seconds(New_Seconds As Integer)
  If Seconds = New_Seconds Then Exit Property
  ChangedByCode = True
  txtVal(2).Text = CStr(New_Seconds)
  ChangedByCode = False
End Property
Public Property Let TotalSeconds(New_TotalSeconds As Long)
  If TotalSeconds = New_TotalSeconds Then Exit Property
  Dim s As Long, t As Long
  t = (New_TotalSeconds \ 60)
  txtVal(2).Text = CStr((New_TotalSeconds - (t * 60)))
  s = (t \ 60)
  txtVal(0).Text = CStr(s)
  txtVal(1).Text = CStr((t - (s * 60)))
End Property

Public Sub Refresh(Optional FullRefresh As Boolean = False)
  UserControl.Cls
  If FullRefresh = True Then SetBorderRect
  If FullRefresh = True Then AlignControlsTop
  DrawBorder
  DrawTitle
  Call CheckWindowSize
  RefreshObjects
End Sub

Private Sub RefreshObjects()
  Dim c As Object
  For Each c In UserControl.Controls
    c.Refresh
  Next
End Sub

Private Function CheckWindowSize() As Boolean
  If UserControl.ScaleWidth <> BorderRect.Right Then
    If BorderRect.Right > 0 Then
      UserControl.Width = ((UserControl.Width - UserControl.ScaleWidth) + BorderRect.Right)
      Exit Function
    End If
  End If
  If UserControl.ScaleHeight <> BorderRect.Bottom Then
    If BorderRect.Bottom > 0 Then
      UserControl.Height = ((UserControl.Height - UserControl.ScaleHeight) + BorderRect.Bottom)
      Exit Function
    End If
  End If
  If (BorderRect.Right Or BorderRect.Bottom) = 0 Then
    Refresh True
    Exit Function
  End If
  CheckWindowSize = True
End Function

Private Sub SetBorderRect()
  With BorderRect
    .Left = 0
    .Top = (Screen.TwipsPerPixelY * 10)
    AlignControlsLeft
    .Right = ((txtVal(2).Left + txtVal(2).Width) + (Screen.TwipsPerPixelX * 7))
  End With
End Sub

Private Sub AlignControlsLeft()
  lblHours.Left = (BorderRect.Left + (Screen.TwipsPerPixelX * 7))
  txtVal(0).Left = ((lblHours.Left + lblHours.Width) + (Screen.TwipsPerPixelX * 5))
  lblMinutes.Left = ((txtVal(0).Left + txtVal(0).Width) + (Screen.TwipsPerPixelX * 10))
  txtVal(1).Left = ((lblMinutes.Left + lblMinutes.Width) + (Screen.TwipsPerPixelX * 5))
  lblSeconds.Left = ((txtVal(1).Left + txtVal(1).Width) + (Screen.TwipsPerPixelX * 10))
  txtVal(2).Left = ((lblSeconds.Left + lblSeconds.Width) + (Screen.TwipsPerPixelX * 3))
End Sub

Private Sub AlignControlsTop()
  txtVal(1).Top = txtVal(0).Top
  txtVal(2).Top = txtVal(1).Top
  lblHours.Top = (txtVal(0).Top + ((txtVal(0).Height - lblHours.Height) \ 2))
  lblMinutes.Top = lblHours.Top
  lblSeconds.Top = lblMinutes.Top
  BorderRect.Bottom = ((txtVal(2).Top + txtVal(2).Height) + (Screen.TwipsPerPixelY * 7))
End Sub

Private Sub DrawBorder()
  Dim cX As Long, cY As Long, cX2 As Long, cY2 As Long
  cX = 0
  cY = (Screen.TwipsPerPixelY * 10)
  cX2 = (UserControl.ScaleWidth - Screen.TwipsPerPixelX)
  cY2 = (UserControl.ScaleHeight - Screen.TwipsPerPixelY)
  UserControl.Line (cX, cY)-(cX2, cY2), COLOR_OUTLINE, B
  UserControl.Line ((cX + Screen.TwipsPerPixelX), (cY + Screen.TwipsPerPixelY))-((cX2 - Screen.TwipsPerPixelX), (cY2 - Screen.TwipsPerPixelY)), COLOR_OUTLINE, B
End Sub

Private Sub DrawTitle(Optional DoNotMove As Boolean = False)
  UserControl.CurrentX = (Screen.TwipsPerPixelX * 10)
  UserControl.CurrentY = 15
  UserControl.ForeColor = COLOR_TEXT
  UserControl.Print m_Title
  If DoNotMove = False Then txtVal(0).Top = (UserControl.CurrentY + (Screen.TwipsPerPixelY * 7))
End Sub

Private Sub SetEnabled()
  Dim i As Integer, c As Long
  If m_Enabled = True Then c = COLOR_CONTROLS Else c = COLOR_BACKGROUND_DISABLED
  For i = 0 To 2
    txtVal(i).Enabled = m_Enabled
    txtVal(i).BackColor = c
  Next
End Sub

Private Function GetTextRect(TextIndex As Integer) As RECT
  With GetTextRect
    .Left = txtVal(TextIndex).Left
    .Top = txtVal(TextIndex).Top
    .Right = (.Left + txtVal(TextIndex).Width)
    .Bottom = (.Top + txtVal(TextIndex).Height)
  End With
End Function

Private Function CanInteractNow() As Boolean
  CanInteractNow = UserControl.Ambient.UserMode And m_Enabled And UserControl.Extender.Visible And UserControl.hWnd <> 0
End Function

Private Sub txtVal_Change(Index As Integer)
  If m_Enabled = False Then Exit Sub
  If ChangedByCode = True Then Exit Sub
  With txtVal(Index)
    If (IsNumeric(.Text) = False And .Text <> "") Then
      ChangedByCode = True
      .Text = LastGoodVal(Index)
      ChangedByCode = False
    Else
      LastGoodVal(Index) = .Text
      RaiseEvent ValueChanged
    End If
  End With
End Sub

Private Sub txtVal_GotFocus(Index As Integer)
  If m_Enabled = False Then Exit Sub
  With txtVal(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtVal_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If m_Enabled = False Then Exit Sub
  Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyDelete, vbKeyBack
      'do nothing...
    Case Else
      KeyCode = 0
  End Select
End Sub

Private Sub txtVal_KeyPress(Index As Integer, KeyAscii As Integer)
  If m_Enabled = False Then Exit Sub
  Select Case KeyAscii
    Case vbKeyBack, vbKeyDelete
      Exit Sub
  End Select
  If KeyAscii = vbKeyBack Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub UserControl_Initialize()
  SetEnabled
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Enabled = PropBag.ReadProperty(PROPNAME_ENABLED, True)
  m_Title = PropBag.ReadProperty(PROPNAME_TITLE, "Title")
  SetEnabled
End Sub

Private Sub UserControl_Resize()
  DontResize = True
  If CheckWindowSize = True Then Refresh True
  DontResize = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_ENABLED, m_Enabled, True
  PropBag.WriteProperty PROPNAME_TITLE, m_Title, "Title"
End Sub
