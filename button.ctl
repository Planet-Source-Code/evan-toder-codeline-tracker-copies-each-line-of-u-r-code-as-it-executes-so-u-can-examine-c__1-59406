VERSION 5.00
Begin VB.UserControl button 
   AutoRedraw      =   -1  'True
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1245
   ScaleHeight     =   345
   ScaleWidth      =   1245
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   8325
      ScaleHeight     =   510
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   5490
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   405
      Top             =   540
   End
End
Attribute VB_Name = "button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
 
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Const SRCCOPY = &HCC0020

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type RECT
   Left As Long
   Top  As Long
   Right As Long
   Bottom As Long
End Type

Enum enCaptionAlign
    caCenter = 0
    caLeft = 1
    caRight = 2
End Enum


Dim m_forecolor    As Long
Dim btnRgn         As Long
Dim R              As RECT
Dim m_bEntered     As Boolean

Event Click()
Event MouseDown()
Event MouseUp()
Event MouseEnter()
Event MouseExit()
 
'default property values
Const m_def_Caption = "caption"
Const m_def_CaptionAlign = 0
Const m_def_ShowClickAnimation = 0

'Property Variables:
Dim m_MouseOverBorderColor As OLE_COLOR
Dim m_ShowMouseoverColor  As Boolean
Dim m_MouseOverColor      As OLE_COLOR
Dim m_ShowClickAnimation  As Boolean
Dim m_IsBusy              As Boolean
Dim m_Caption             As String
Dim m_CaptionAlign        As enCaptionAlign

'Default Property Values:
Const m_def_MouseOverBorderColor = &HFF8080
Const m_def_ShowMouseoverColor = &HFFC0C0
Const m_def_MouseOverColor = False
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
 
 













 

 

 

 

Private Sub UserControl_Click()
 Dim Click                     As Control
  If m_IsBusy Then Exit Sub
  Call ActivateEffect
  RaiseEvent Click
End Sub
Private Sub UserControl_Terminate()
     Timer1.Interval = 0
     Timer1.Enabled = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_IsBusy Then Exit Sub
    Call Paint(2)
    RaiseEvent MouseDown
End Sub

'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  12/26/2004 1:29:12 AM
'
'     UserControl_MouseMove: On this event the control is
'     repainted with the raised appearance, font is made bold, and
'     timer1 is turned on which montors for when the mouse leaves
'     this control
'
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If m_IsBusy Then Exit Sub
  
  If Not (m_bEntered) Then
     m_bEntered = True
     Timer1.Interval = 200
     Timer1 = True
     Call Paint(1)
     RaiseEvent MouseEnter
  End If
  
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If m_IsBusy Then Exit Sub
  Call Paint(1)
  RaiseEvent MouseUp
End Sub
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  12/26/2004 1:27:57 AM
'
'     Timer1_Timer: this timer is activated as soon as the mouse
'     moves within the bounds of this control. This timer checks,
'     every 1/10 second or so, if the mouse pointer is still over
'     this control.  As soon as its not, it repaints, fires the
'     MouseExit event, and turns itself off
'
'
'     PART                DESCRIPTION
'     -------------------------------------------------------
'
Private Sub Timer1_Timer()
  Dim pt As POINTAPI
  
  Call GetCursorPos(pt)
  If WindowFromPoint(pt.X, pt.Y) <> hwnd Then
     Timer1.Interval = 0
     Timer1.Enabled = False
     m_bEntered = False
     Call Paint(0)
     RaiseEvent MouseExit
  End If
 
End Sub
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  12/26/2004 1:30:39 AM
'
'     UserControl_Resize: The rect that defines where the drawing
'     of the buttons text is is defined/redefined in this event
'
Private Sub UserControl_Resize()
  Call Paint(0)
End Sub
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  12/26/2004 1:26:11 AM
'
'     ActivateEffect:  this is an effect that provides clear
'     feedback to the user that he successfully clicked the button
'     and that "something is happening"  this effect is acheived
'     by creating an expanding focus rect and the cursor turns
'     into a watch.  This feedback effect lasts for about 1
'     second.  This effect can be made longer or shorter by
'     changing the Step command to a larger or smaller number
'
Private Sub ActivateEffect()
   With UserControl
     'get the center point of the control
     Dim centerX                              As Long
     Dim centerY                              As Long
     centerX = twip2PixX((.Width * 0.5))
     centerY = twip2PixY((.Height * 0.5))
     Dim sCnt                                 As Single
     Dim newR                                 As RECT
     'temp disables any further input to the button
     'til this effect is done and helps to prevent
     'an accidental dbl click
     m_IsBusy = True
     
     If m_ShowClickAnimation Then
        For sCnt = 1 To 12 Step 0.03
          Call Paint(1)
          'expanding focus rect effect
          Call SetRect(newR, centerX - sCnt, centerY - sCnt, centerX + sCnt, centerY + sCnt)
          Call DrawFocusRect(hdc, newR)
          DoEvents
        Next sCnt
     End If
     
     '"re-enable"
     m_IsBusy = False
     Call Paint(1)
     'in case mouse has left b4 end of this effect
     Call Timer1_Timer
  End With
End Sub
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  12/26/2004 1:32:56 AM
'
'     Paint:  Paints the visual appearance of the button
'
'
'     PART                DESCRIPTION
'     -------------------------------------------------------
'     lVal                [Required | long]
'                         the parameter lVal is used to determine
'                         if we paint the button with visual
'                         appearance of mousehover, mousedown, or
'                         otherwise...i.e in the
'                         Usercontrol_Mousedown..this sub is
'                         called with lVal=2, which signals this
'                         function to paint a mousedown visual
'                         effect
'     -------------------------------------------------------
'
Private Function Paint(lVal&)
  Dim clr&
  
  Cls
  
  Select Case lVal
    Case Is = 0 'flat
        UserControl.ForeColor = m_forecolor
    Case Is = 1 'mouseover
        Call DrawLines(1)
        Call DrawHilight
        UserControl.ForeColor = vbBlue
    Case Is = 2 'mousedown
        Call DrawLines(2)
        Call DrawHilight
        UserControl.ForeColor = vbBlue
  End Select
 
  If Picture <> 0 Then
      Dim xWid As Long, xHei As Long
      
      With Picture1
         xWid = twip2PixX(.Width)
         xHei = twip2PixY(.Height)
         Call BitBlt(hdc, 2, 2, xWid, xHei, .hdc, 0, 0, SRCCOPY)
      End With
  End If
  
  Call PrintText
  
End Function

Private Sub DrawHilight()
  '
  'draws the focus hilight rect
  '
  If m_ShowMouseoverColor = True Then
     Dim hBrush   As Long
     Dim hBrush2  As Long
     Dim hRgn     As Long
     Dim R        As RECT
     
     hBrush = CreateSolidBrush(m_MouseOverColor)
     hBrush2 = CreateSolidBrush(m_MouseOverBorderColor)
     hRgn = CreateRoundRectRgn(2, 2, twip2PixX(Width - 15), twip2PixY(Height - 15), 1, 1)
     FillRgn hdc, hRgn, hBrush
     FrameRgn hdc, hRgn, hBrush2, 1, 1
     
     DeleteObject hRgn
     DeleteObject hBrush2
     DeleteObject hBrush
  End If
  
End Sub
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  12/26/2004 1:35:59 AM
'
'     DrawLines:  Is  called by sub [Paint]. actually draws the
'     four lines that create the visual box defining the borders
'     of the button
'
Private Sub DrawLines(lVal&)
   Dim lclr(1) As Long
   Dim W As Long, H As Long
 
   If lVal = 1 Then
      lclr(0) = vbWhite
      lclr(1) = RGB(180, 180, 190)
   ElseIf lVal = 2 Then
      lclr(1) = vbWhite
      lclr(0) = RGB(180, 180, 190)
   End If
   
   W = (Width - 10)
   H = (Height - 10)
   UserControl.Line (5, 5)-(W, 5), lclr(0)
   UserControl.Line (5, 5)-(5, H), lclr(0)
   UserControl.Line (W, 10)-(W, H), lclr(1)
   UserControl.Line (10, H)-(W, H), lclr(1)
End Sub
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  12/26/2004 1:33:53 AM
'
'     PrintText: Prints the text to the control within the
'     boundries of rect [R] which is set in the Usercontrol_Resize
'     event
'
Private Sub PrintText()
  Dim dtCalcVal  As Long, dtVal   As Long
  Const DT_CALCRECT As Long = &H400
  Const DT_CENTER As Long = &H1
  Const DT_RIGHT As Long = &H2
  Const DT_LEFT As Long = &H0
  Const DT_VCENTER As Long = &H4
  Const DT_SINGLELINE As Long = &H20
 
  If m_CaptionAlign = caLeft Then
       dtVal = (DT_SINGLELINE Or DT_VCENTER Or DT_LEFT)
  ElseIf m_CaptionAlign = caCenter Then
       dtVal = (DT_SINGLELINE Or DT_VCENTER Or DT_CENTER)
  ElseIf m_CaptionAlign = caRight Then
       dtVal = (DT_SINGLELINE Or DT_VCENTER Or DT_RIGHT)
  End If
 
   'this is the rect where we will be drawing the caption
  Call SetRect(R, 5, 5, (twip2PixX(Width - 100)), twip2PixY(Height - 50))
  Call DrawText(hdc, m_Caption, Len(m_Caption), R, dtVal)
End Sub
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'PURPOSE:                 Coded on  12/26/2004 1:37:30 AM
'
'     twip2PixX:  twip2PixY:  Converts twips to pixels which is
'     required for API processing compliance
'
'
'     PART                DESCRIPTION
'     -------------------------------------------------------
'     twipVal&            [Required | long]
'                         This is the twips value that is to be
'                         converted
'     -------------------------------------------------------
'
Private Function twip2PixX(twipVal&) As Long
    twip2PixX = (twipVal / Screen.TwipsPerPixelX)
End Function
Private Function twip2PixY(twipVal&) As Long
    twip2PixY = (twipVal / Screen.TwipsPerPixelY)
End Function














'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'  FROM HERE DOWN IS PROPERTY SET AND GET RELATED CODE
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº




'BACKCOLOR
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Call UserControl_Resize
End Property
'CAPTION
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Call UserControl_Resize
End Property
'CaptionAlign
Public Property Get CaptionAlign() As enCaptionAlign
    CaptionAlign = m_CaptionAlign
End Property
Public Property Let CaptionAlign(ByVal New_CaptionAlign As enCaptionAlign)
    m_CaptionAlign = New_CaptionAlign
    PropertyChanged "CaptionAlign"
    Call UserControl_Resize
End Property
'FONT
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Call UserControl_Resize
End Property
'FORECOLOR
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    m_forecolor = New_ForeColor
    Call UserControl_Resize
End Property
'hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
'IsBusy (public read only)
Public Property Get IsBusy() As Boolean
    IsBusy = m_IsBusy
End Property
Private Property Let IsBusy(ByVal New_IsBusy As Boolean)
    m_IsBusy = New_IsBusy
End Property
'MouseOverColor
Public Property Get MouseOverColor() As OLE_COLOR
    MouseOverColor = m_MouseOverColor
End Property
Public Property Let MouseOverColor(ByVal New_MouseOverColor As OLE_COLOR)
    m_MouseOverColor = New_MouseOverColor
    PropertyChanged "MouseOverColor"
End Property
'MouseOverBorderColor
Public Property Get MouseOverBorderColor() As OLE_COLOR
    MouseOverBorderColor = m_MouseOverBorderColor
End Property
Public Property Let MouseOverBorderColor(ByVal New_MouseOverBorderColor As OLE_COLOR)
    m_MouseOverBorderColor = New_MouseOverBorderColor
    PropertyChanged "MouseOverBorderColor"
End Property
'MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
'MousePointer
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property
'Picture
Public Property Get Picture() As Picture
    Set Picture = Picture1.Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    Set Picture1.Picture = New_Picture
    PropertyChanged "Picture"
    Call UserControl_Resize
End Property
'ShowClickAnimation
Public Property Get ShowClickAnimation() As Boolean
    ShowClickAnimation = m_ShowClickAnimation
End Property
Public Property Let ShowClickAnimation(ByVal New_ShowClickAnimation As Boolean)
    m_ShowClickAnimation = New_ShowClickAnimation
    PropertyChanged "ShowClickAnimation"
End Property
'ShowMouseoverColor
Public Property Get ShowMouseoverColor() As Boolean
    ShowMouseoverColor = m_ShowMouseoverColor
End Property
Public Property Let ShowMouseoverColor(ByVal New_ShowMouseoverColor As Boolean)
    m_ShowMouseoverColor = New_ShowMouseoverColor
    PropertyChanged "ShowMouseoverColor"
End Property

Private Sub UserControl_InitProperties()
   m_Caption = m_def_Caption
   Call UserControl_Resize
    m_CaptionAlign = m_def_CaptionAlign
    m_ShowClickAnimation = m_def_ShowClickAnimation
    m_MouseOverColor = m_def_MouseOverColor
    m_ShowMouseoverColor = m_def_ShowMouseoverColor
    m_MouseOverBorderColor = m_def_MouseOverBorderColor
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    Set Picture1.Picture = PropBag.ReadProperty("Picture", Nothing)
    m_CaptionAlign = PropBag.ReadProperty("CaptionAlign", m_def_CaptionAlign)
    m_ShowClickAnimation = PropBag.ReadProperty("ShowClickAnimation", m_def_ShowClickAnimation)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_MouseOverColor = PropBag.ReadProperty("MouseOverColor", m_def_MouseOverColor)
    m_ShowMouseoverColor = PropBag.ReadProperty("ShowMouseoverColor", m_def_ShowMouseoverColor)
    m_MouseOverBorderColor = PropBag.ReadProperty("MouseOverBorderColor", m_def_MouseOverBorderColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    
    m_forecolor = UserControl.ForeColor
    Call UserControl_Resize
    Call Paint(0)

End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Picture", Picture1.Picture, Nothing)
    Call PropBag.WriteProperty("CaptionAlign", m_CaptionAlign, m_def_CaptionAlign)
    Call PropBag.WriteProperty("ShowClickAnimation", m_ShowClickAnimation, m_def_ShowClickAnimation)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseOverColor", m_MouseOverColor, m_def_MouseOverColor)
    Call PropBag.WriteProperty("ShowMouseoverColor", m_ShowMouseoverColor, m_def_ShowMouseoverColor)
    Call PropBag.WriteProperty("MouseOverBorderColor", m_MouseOverBorderColor, m_def_MouseOverBorderColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

