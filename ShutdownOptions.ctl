VERSION 5.00
Begin VB.UserControl ShutdownOptions 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   7485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9630
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
   ScaleHeight     =   7485
   ScaleWidth      =   9630
   Begin ShutdownTimer2.CheckBox chkOnTop 
      Height          =   300
      Left            =   1035
      TabIndex        =   0
      Top             =   3420
      Width           =   3975
      _extentx        =   7011
      _extenty        =   529
      caption         =   "Always keep this window on top"
   End
   Begin ShutdownTimer2.CheckBox chkUpdate 
      Height          =   300
      Left            =   1170
      TabIndex        =   1
      Top             =   2235
      Width           =   5655
      _extentx        =   9975
      _extenty        =   529
      caption         =   "Install pending Windows updates if available"
   End
   Begin ShutdownTimer2.CheckBox chkForce 
      Height          =   300
      Left            =   1170
      TabIndex        =   2
      Top             =   1005
      Width           =   5055
      _extentx        =   2143
      _extenty        =   529
      caption         =   "Force running apps to close immediately"
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Unsaved work will be lost)"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   1665
      TabIndex        =   3
      Top             =   1515
      Width           =   2835
   End
End
Attribute VB_Name = "ShutdownOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PROPNAME_TITLE = "Title"
Private Const PROPNAME_ENABLED = "Enabled"

Private Const LABELCAPTION_UPDATES = "Pending updates are waiting"
Private Const LABELCAPTION_NOUPDATES = "There are no pending updates"

Dim m_Title As String
Dim m_Enabled As Boolean

Dim BoxTop As Long

Dim WindowRect As RECT

Public Event ForceExitClick()
Public Event InstallUpdatesClick()
Public Event AlwaysOnTopClick()

Public Property Get Title() As String
  Title = m_Title
End Property
Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property
Public Property Get ForceExit() As Boolean
  ForceExit = (chkForce.Value = vbChecked)
End Property
Public Property Get InstallUpdates() As Boolean
  InstallUpdates = (chkUpdate.Value = vbChecked)
End Property
Public Property Get AlwaysOnTop() As Boolean
  AlwaysOnTop = (chkOnTop.Value = vbChecked)
End Property
Public Property Let Title(New_Title As String)
  If m_Title = New_Title Then Exit Property
  m_Title = New_Title
  UserControl.PropertyChanged PROPNAME_TITLE
  Refresh
End Property
Public Property Let Enabled(New_Enabled As Boolean)
  If m_Enabled = New_Enabled Then Exit Property
  m_Enabled = New_Enabled
  UserControl.PropertyChanged PROPNAME_ENABLED
  Refresh
End Property
Public Property Let ForceExit(New_ForceExit As Boolean)
  chkForce.Value = IIf(New_ForceExit, vbChecked, vbUnchecked)
End Property
Public Property Let InstallUpdates(New_InstallUpdates As Boolean)
  chkUpdate.Value = IIf(New_InstallUpdates, vbChecked, vbUnchecked)
End Property
Public Property Let AlwaysOnTop(New_AlwaysOnTop As Boolean)
  chkOnTop.Value = IIf(New_AlwaysOnTop, vbChecked, vbUnchecked)
End Property

Public Sub Refresh(Optional RedrawAll As Boolean = False)
  UserControl.Cls
  If RedrawAll = True Then Call CheckWindowSize
  DrawBorder
  DrawTitle
  RefreshObjects
End Sub

Private Sub RefreshObjects()
  Dim c As Object
  For Each c In UserControl.Controls
    c.Refresh
  Next
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
  BoxTop = (UserControl.CurrentY + (Screen.TwipsPerPixelY * 7))
End Sub

Private Sub SetWindowRect()
  With WindowRect
    .Left = 0
    .Top = 0
    MoveObjects
    .Right = GetWidestLabel()
    If .Right < UserControl.ScaleWidth Then .Right = UserControl.ScaleWidth
    .Bottom = ((chkOnTop.Top + chkOnTop.Height) + (Screen.TwipsPerPixelY * 5))
  End With
End Sub

Private Sub MoveObjects()
  Dim t As Long
  If BoxTop = 0 Then t = (WindowRect.Top + (Screen.TwipsPerPixelY * 15)) Else t = BoxTop
  chkForce.Move (WindowRect.Left + (Screen.TwipsPerPixelX * 5)), t
  lblWarning.Move ((chkForce.Left + chkForce.Height) + (Screen.TwipsPerPixelX * 3)), ((chkForce.Top + chkForce.Height) + (Screen.TwipsPerPixelY * 2))
  chkUpdate.Move chkForce.Left, ((lblWarning.Top + lblWarning.Height) + (Screen.TwipsPerPixelY * 5))
  chkOnTop.Move chkUpdate.Left, ((chkUpdate.Top + chkUpdate.Height) + (Screen.TwipsPerPixelY * 5))
End Sub

Private Function GetWidestLabel() As Long
  Dim r As Long
  r = chkForce.Width
  If chkUpdate.Width > r Then r = chkUpdate.Width
  If chkOnTop.Width > r Then r = chkOnTop.Width
  GetWidestLabel = (chkForce.Left + r)
End Function

Private Function CheckWindowSize() As Boolean
  SetWindowRect
  If UserControl.ScaleWidth < WindowRect.Right Then
    UserControl.Width = ((UserControl.Width - UserControl.ScaleWidth) + WindowRect.Right)
    Exit Function
  End If
  If UserControl.ScaleHeight <> WindowRect.Bottom Then
    UserControl.Height = ((UserControl.Height - UserControl.ScaleHeight) + WindowRect.Bottom)
    Exit Function
  End If
  CheckWindowSize = True
End Function

Private Sub chkForce_Click()
  RaiseEvent ForceExitClick
End Sub

Private Sub chkOnTop_Click()
  RaiseEvent AlwaysOnTopClick
End Sub

Private Sub chkUpdate_Click()
  RaiseEvent InstallUpdatesClick
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Title = PropBag.ReadProperty(PROPNAME_TITLE, "Title")
  m_Enabled = PropBag.ReadProperty(PROPNAME_ENABLED, True)
End Sub

Private Sub UserControl_Resize()
  If CheckWindowSize() = True Then Refresh True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_TITLE, m_Title, "Title"
  PropBag.WriteProperty PROPNAME_ENABLED, m_Enabled, True
End Sub
