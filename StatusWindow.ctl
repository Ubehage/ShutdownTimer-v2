VERSION 5.00
Begin VB.UserControl StatusWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label lStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   510
      Left            =   825
      TabIndex        =   0
      Top             =   930
      Width           =   1440
   End
End
Attribute VB_Name = "StatusWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PROPNAME_TITLE As String = "Title"
Private Const PROPNAME_CAPTION As String = "Caption"

Private Const CAPTION_STATUS = "h:m:s"

Dim m_Title As String
Dim m_Caption As String
Dim m_Hours As Long
Dim m_Minutes As Long
Dim m_Seconds As Long

Dim WindowRect As RECT

Public Property Get Title() As String
  Title = m_Title
End Property
Public Property Get Caption() As String
  Caption = m_Caption
End Property
Public Property Get Hours() As Long
  Hours = m_Hours
End Property
Public Property Get Minutes() As Long
  Minutes = m_Minutes
End Property
Public Property Get Seconds() As Long
  Seconds = m_Seconds
End Property
Public Property Let Title(New_Title As String)
  If m_Title = New_Title Then Exit Property
  m_Title = New_Title
  UserControl.PropertyChanged PROPNAME_TITLE
  Refresh True
End Property
Public Property Let Caption(New_Caption As String)
  If m_Caption = New_Caption Then Exit Property
  m_Caption = New_Caption
  SplitCaption
  UserControl.PropertyChanged PROPNAME_CAPTION
  Refresh
End Property
Public Property Let Hours(New_Hours As Long)
  If m_Hours = New_Hours Then Exit Property
  m_Hours = New_Hours
  SetCaption
End Property
Public Property Let Minutes(New_Minutes As Long)
  If m_Minutes = New_Minutes Then Exit Property
  m_Minutes = New_Minutes
  SetCaption
End Property
Public Property Let Seconds(New_Seconds As Long)
  If m_Seconds = New_Seconds Then Exit Property
  m_Seconds = New_Seconds
  SetCaption
End Property

Public Sub Refresh(Optional FullRedraw As Boolean = False)
  If FullRedraw = True Then UserControl.Cls
  If FullRedraw = True Then DrawBorder
  If FullRedraw = True Then DrawTitle
  lStatus.Caption = m_Caption
  lStatus.Move (UserControl.ScaleWidth - lStatus.Width) \ 2, (((UserControl.ScaleHeight - lStatus.Height) \ 2) + (Screen.TwipsPerPixelY * 6))
End Sub

Private Sub DrawBorder()
  Dim cX As Long, cY As Long
  With WindowRect
    cX = .Left
    cY = (.Top + (Screen.TwipsPerPixelY * 10))
    UserControl.Line (cX, cY)-((.Right - Screen.TwipsPerPixelX), (.Bottom - Screen.TwipsPerPixelY)), COLOR_OUTLINE, B
    UserControl.Line ((cX + Screen.TwipsPerPixelX), (cY + Screen.TwipsPerPixelY))-((.Right - (Screen.TwipsPerPixelX * 2)), (.Bottom - (Screen.TwipsPerPixelY * 2))), COLOR_OUTLINE, B
  End With
End Sub

Private Sub DrawTitle()
  UserControl.CurrentX = (Screen.TwipsPerPixelX * 10)
  UserControl.CurrentY = Screen.TwipsPerPixelY
  UserControl.ForeColor = COLOR_TEXT
  UserControl.Print m_Title
End Sub

Private Sub SetCaption()
  m_Caption = Replace(Replace(Replace(CAPTION_STATUS, "h", NumToDigitString(m_Hours)), "m", NumToDigitString(m_Minutes)), "s", NumToDigitString(m_Seconds))
  UserControl.PropertyChanged PROPNAME_CAPTION
  Refresh
End Sub

Private Function NumToDigitString(sNum As Long) As String
  Dim r As String
  r = CStr(sNum)
  If Len(r) = 1 Then NumToDigitString = "0" & r Else NumToDigitString = r
End Function

Private Sub SplitCaption()
  Dim i As Long, cArr() As String, c As Integer
  m_Hours = 0
  m_Minutes = 0
  m_Seconds = 0
  cArr = Split(m_Caption, ":")
  For i = UBound(cArr) To LBound(cArr) Step -1
    If c = 0 Then
      m_Seconds = Val(cArr(i))
    ElseIf c = 1 Then
      m_Minutes = Val(cArr(i))
    ElseIf c = 2 Then
      m_Hours = Val(cArr(i))
    Else
      Exit For
    End If
    c = (c + 1)
  Next
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
  m_Title = "Title"
  m_Caption = "00:00:00"
  SplitCaption
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Title = PropBag.ReadProperty(PROPNAME_TITLE, "Title")
  m_Caption = PropBag.ReadProperty(PROPNAME_CAPTION, "Caption")
  SplitCaption
End Sub

Private Sub UserControl_Resize()
  SetWindowRect
  Refresh True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_TITLE, m_Title, "Title"
  PropBag.WriteProperty PROPNAME_CAPTION, m_Caption, "Caption"
End Sub
