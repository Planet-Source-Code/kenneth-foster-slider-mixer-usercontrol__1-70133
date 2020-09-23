VERSION 5.00
Begin VB.UserControl Fader 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1080
   ScaleHeight     =   82
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   72
   Begin VB.PictureBox DSH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1845
      Picture         =   "Fader.ctx":0000
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   4
      Top             =   1710
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox DSV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   1845
      Picture         =   "Fader.ctx":00D1
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   3
      Top             =   1065
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DraggerSourceV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   1350
      Picture         =   "Fader.ctx":0192
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   2
      Top             =   1665
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox DraggerSourceH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1455
      Picture         =   "Fader.ctx":01F8
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox DraggerSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   105
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   0
      Top             =   165
      Visible         =   0   'False
      Width           =   390
   End
End
Attribute VB_Name = "Fader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Original code by Mike Payne

'To use T.O.P for a search , highlight just the word or words you want to search for.
'Press Ctrl + F3
' To continue search , just use F3 until all uses are found.

'***************** Table of Procedures *************
'   Private Sub UserControl_Initialize
'   Private Sub UserControl_MouseMove
'   Private Sub UserControl_MouseDown
'   Private Sub UserControl_ReadProperties
'   Private Sub UserControl_Resize
'   Private Sub UserControl_WriteProperties
'   Public Function OneIncrement
'   Public Sub DrawPicAtValue
'   Private Sub Redraw
'   Public Property Get BackColor
'   Public Property Let BackColor
'   Public Property Get ButSz
'   Public Property Let ButSz
'   Public Property Get HalfMark
'   Public Property Let HalfMark
'   Public Property Get HalfMarkColor
'   Public Property Let HalfMarkColor
'   Public Property Get Max
'   Public Property Let Max
'   Public Property Get TickMarks
'   Public Property Get Style
'   Public Property Let Style
'   Public Property Let TickMarks
'   Public Property Get TickMarkCnt
'   Public Property Let TickMarkCnt
'   Public Property Get Value
'   Public Property Let Value
'***************** End of Table ********************

Event Scrolling()
    Private Enum RasterOps
        srccopy = &HCC0020
         SRCAND = &H8800C6
         SRCINVERT = &H660046
         SRCPAINT = &HEE0086
         SRCERASE = &H4400328
         WHITENESS = &HFF0062
         BLACKNESS = &H42
    End Enum
    
    Public Enum eStyle
       Vertical = 0
       Horizontal = 1
    End Enum
    
    Public Enum eButSz
       Small = 0
       Large = 1
    End Enum
    
     Private Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As RasterOps _
        ) As Long

Const m_def_Max = 100
Const m_def_m_Value = 0
Const m_def_Style = Vertical
Const m_def_HalfMark = False
Const m_def_HalfMarkColor = vbRed
Const m_def_BackColor = &HC0C0C0
Const m_def_TickMarks = True
Const m_def_TickMarkCnt = 5
Const m_def_ButSz = 0

Dim m_Value As Integer
Dim m_Max As Long
Dim m_m_Value As Long
Dim m_Style As eStyle
Dim m_HalfMark As Boolean
Dim m_HalfMarkColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_TickMarks As Boolean
Dim m_TickMarkCnt As Integer
Dim m_ButSz As eButSz

Dim Faderheight As Integer
Dim BottomBoundary As Integer
Dim Faderwidth As Integer
Dim LeftBoundary As Integer

Private Sub UserControl_Initialize()
   Max = m_def_Max
   Style = m_def_Style
   HalfMark = m_def_HalfMark
   HalfMarkColor = m_def_HalfMarkColor
   BackColor = m_def_BackColor
   TickMarks = m_def_TickMarks
   TickMarkCnt = m_def_TickMarkCnt
   ButSz = m_def_ButSz
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim realY As Single
   Dim realX As Single
   
    If Button = 1 Then
       If Style = Vertical Then
          If ButSz = 0 Then
             realY = (Y - 5) / OneIncrement 'use the Y value of the mouse to calculate the value
             m_Value = (m_Max - realY)   'set the value
          Else
              realY = (Y - 15) / OneIncrement 'use the Y value of the mouse to calculate the value
              m_Value = (m_Max - realY)   'set the value
          End If
       Else
          If ButSz = 0 Then
             realX = (X - 5) / OneIncrement 'use the X value of the mouse to calculate the value
             m_Value = m_Max - (m_Max - realX)   'set the value
          Else
             realX = (X - 15) / OneIncrement 'use the X value of the mouse to calculate the value
             m_Value = m_Max - (m_Max - realX)   'set the value
          End If
       End If
       Redraw
    RaiseEvent Scrolling
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim realY As Single
    Dim realX As Single
    
    If Button = 1 Then
       If Style = Vertical Then
          If ButSz = 0 Then
             realY = (Y - 5) / OneIncrement 'use the Y value of the mouse to calculate the value
             m_Value = (m_Max - realY)   'set the value
          Else
              realY = (Y - 15) / OneIncrement 'use the Y value of the mouse to calculate the value
              m_Value = (m_Max - realY)   'set the value
          End If
       Else
          If ButSz = 0 Then
             realX = (X - 5) / OneIncrement 'use the X value of the mouse to calculate the value
             m_Value = m_Max - (m_Max - realX)   'set the value
          Else
             realX = (X - 15) / OneIncrement 'use the X value of the mouse to calculate the value
             m_Value = m_Max - (m_Max - realX)   'set the value
          End If
       End If
       Redraw
    RaiseEvent Scrolling
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Max = PropBag.ReadProperty("Max", m_def_Max)
   Value = PropBag.ReadProperty("Value", m_def_m_Value)
   Style = PropBag.ReadProperty("Style", m_def_Style)
   HalfMark = PropBag.ReadProperty("HalfMark", m_def_HalfMark)
   HalfMarkColor = PropBag.ReadProperty("HalfMarkColor", m_def_HalfMarkColor)
   BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
   TickMarks = PropBag.ReadProperty("TickMarks", m_def_TickMarks)
   TickMarkCnt = PropBag.ReadProperty("TickMarkCnt", m_def_TickMarkCnt)
   ButSz = PropBag.ReadProperty("ButSz", m_def_ButSz)
End Sub

Private Sub UserControl_Resize()
 
 If m_Style = Vertical Then
    If ButSz = 0 Then
       BottomBoundary = UserControl.ScaleHeight - 10
       Faderheight = UserControl.ScaleHeight - 1
       UserControl.Width = 335
    Else
       BottomBoundary = UserControl.ScaleHeight - 30
       Faderheight = UserControl.ScaleHeight - 1
       UserControl.Width = 335
    End If
 Else
    If ButSz = 0 Then
       LeftBoundary = UserControl.ScaleWidth - 15
       Faderwidth = UserControl.ScaleWidth - 1
       UserControl.Height = 335
    Else
       LeftBoundary = UserControl.ScaleWidth - 30
       Faderwidth = UserControl.ScaleWidth - 1
       UserControl.Height = 335
    End If
 End If
 Redraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
   Call PropBag.WriteProperty("Value", m_Value, m_def_m_Value)
   Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
   Call PropBag.WriteProperty("HalfMark", m_HalfMark, m_def_HalfMark)
   Call PropBag.WriteProperty("HalfMarkColor", m_HalfMarkColor, m_def_HalfMarkColor)
   Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
   Call PropBag.WriteProperty("TickMarks", m_TickMarks, m_def_TickMarks)
   Call PropBag.WriteProperty("TickMarkCnt", m_TickMarkCnt, m_def_TickMarkCnt)
   Call PropBag.WriteProperty("ButSz", m_ButSz, m_def_ButSz)
End Sub

Public Function OneIncrement() As Double
   If Style = Vertical Then
      If ButSz = 0 Then
        OneIncrement = (BottomBoundary - 1) / m_Max
      Else
        OneIncrement = (BottomBoundary - 22) / m_Max
      End If
   Else
     If ButSz = 0 Then
        OneIncrement = (LeftBoundary + 4) / m_Max
     Else
        OneIncrement = (LeftBoundary - 2) / m_Max
     End If
   End If
End Function

Public Sub DrawPicAtValue(tValue As Integer)
   Dim NewY As Integer
   Dim NewX As Integer
   
   If tValue > m_Max Then tValue = m_Max 'if new value is too big, make it max size
   If tValue < 0 Then tValue = 0 'if new value is too small, make it 0
   If Style = Vertical Then
        If ButSz = 0 Then
           NewY = ((m_Max - tValue) * OneIncrement) 'calculate new y value of slider
           BitBlt UserControl.hDC, 1, NewY + 1, DraggerSource.Width, DraggerSource.Height, DraggerSource.hDC, 0, 0, srccopy 'draw it
        Else
           NewY = ((m_Max - tValue) * OneIncrement) 'calculate new y value of slider
           BitBlt UserControl.hDC, 1, NewY, DSV.Width, DSV.Height, DSV.hDC, 0, 0, srccopy  'draw it
        End If
   Else
      If ButSz = 0 Then
         NewX = 2 + ((m_Max - tValue) * OneIncrement)  'calculate new X value of slider
         BitBlt UserControl.hDC, (tValue * OneIncrement) + 1, 1, DraggerSource.Width, DraggerSource.Height, DraggerSource.hDC, 0, 0, srccopy  'draw it
      Else
         NewX = 1 + ((m_Max - tValue) * OneIncrement) 'calculate new X value of slider
         BitBlt UserControl.hDC, (tValue * OneIncrement), 1, DSH.Width, DSH.Height, DSH.hDC, 0, 0, srccopy    'draw it
      End If
   End If
End Sub

Private Sub Redraw()
    Dim LightColor As Long, DarkColor As Long
    Dim X As Integer
    
    UserControl.Cls
    'change colors here::
    LightColor = vbWhite
    DarkColor = &H404040
    If Style = Vertical Then
       'draws the 3D looking border lines (no need to change this code)
       UserControl.Line (0, 0)-(20, 0), DarkColor, BF      'top
       UserControl.Line (0, 0)-(0, Faderheight - 1), DarkColor, BF   'left
       UserControl.Line (21, 0)-(21, Faderheight), LightColor, BF    'right
       UserControl.Line (0, Faderheight)-(21, Faderheight), LightColor, BF   'bottom
       
       'draw tick marks
       If TickMarks = True Then
          For X = 3 To Faderheight - 3 Step TickMarkCnt
             UserControl.Line (5, 1 + X)-(16, 1 + X), vbBlack, BF
          Next X
       End If
       If HalfMark = True Then
          UserControl.DrawWidth = 2
          UserControl.Line (1, Faderheight / 2)-(22, Faderheight / 2), HalfMarkColor, BF
          UserControl.DrawWidth = 1
       End If
       'draws the dark rectangle (no need to change this code) ...center bar
       UserControl.Line (9, 2)-(12, 2), DarkColor, BF
       UserControl.Line (9, 2)-(9, Faderheight - 3), DarkColor, BF
       UserControl.Line (12, 3)-(12, Faderheight - 3), LightColor, BF
       UserControl.Line (10, 3)-(11, Faderheight - 3), vbBlack, BF
    Else
       'draws the 3D looking border lines (no need to change this code)
       UserControl.Line (0, 0)-(Faderwidth, 0), DarkColor, BF   'top
       UserControl.Line (0, 0)-(0, Faderheight), DarkColor, BF    'left
        UserControl.Line (Faderwidth, 0)-(Faderwidth, Faderheight), LightColor, BF    'right
       UserControl.Line (0, Faderheight)-(Faderwidth, Faderheight), LightColor, BF   'bottom
       
       'draw tick marks
       If TickMarks = True Then
          For X = 3 To Faderwidth - 3 Step TickMarkCnt
             UserControl.Line (1 + X, 6)-(1 + X, 16), vbBlack, BF
          Next X
       End If
       If HalfMark = True Then
          UserControl.DrawWidth = 2
          UserControl.Line (Faderwidth / 2, 1)-(Faderwidth / 2, 22), HalfMarkColor, BF
          UserControl.DrawWidth = 1
       End If
       'draws the dark rectangle (no need to change this code)...center bar
       UserControl.Line (4, 10)-(Faderwidth - 4, 10), vbBlack, BF
       UserControl.Line (4, 11)-(Faderwidth - 4, 11), vbBlack, BF
       UserControl.Line (4, 12)-(Faderwidth - 4, 12), DarkColor, BF
       UserControl.Line (4, 12)-(Faderwidth - 4, 12), LightColor, BF
    End If
    DrawPicAtValue m_Value
End Sub

Public Property Get BackColor() As OLE_COLOR
   Let BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
   Let m_BackColor = NewBackColor
   UserControl.BackColor = m_BackColor
   PropertyChanged "BackColor"
   Redraw
End Property

Public Property Get ButSz() As eButSz
   Let ButSz = m_ButSz
End Property

Public Property Let ButSz(ByVal NewButSz As eButSz)
   Let m_ButSz = NewButSz
   PropertyChanged "ButSz"
   Redraw
End Property

Public Property Get HalfMark() As Boolean
Attribute HalfMark.VB_Description = "Show half tick mark or not."
   Let HalfMark = m_HalfMark
End Property

Public Property Let HalfMark(ByVal NewHalfMark As Boolean)
   Let m_HalfMark = NewHalfMark
   PropertyChanged "HalfMark"
   Redraw
End Property

Public Property Get HalfMarkColor() As OLE_COLOR
Attribute HalfMarkColor.VB_Description = "Color of the half mark tick"
   Let HalfMarkColor = m_HalfMarkColor
End Property

Public Property Let HalfMarkColor(ByVal NewHalfMarkColor As OLE_COLOR)
   Let m_HalfMarkColor = NewHalfMarkColor
   PropertyChanged "HalfMarkColor"
   Redraw
End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "Scale maximum you want eg. 100"
   Max = m_Max
End Property

Public Property Let Max(NewMax As Long)
   m_Max = NewMax
   PropertyChanged "Max"
End Property

Public Property Get TickMarks() As Boolean
Attribute TickMarks.VB_Description = "Show tick marks or not."
   Let TickMarks = m_TickMarks
End Property

Public Property Get Style() As eStyle
Attribute Style.VB_Description = "Vertical or Horizontal"
   Let Style = m_Style
End Property

Public Property Let Style(ByVal NewStyle As eStyle)
   Let m_Style = NewStyle
   If m_Style = Vertical Then
      If ButSz = 0 Then
         DraggerSource.Picture = DraggerSourceV.Picture
      Else
         DraggerSource.Picture = DSV.Picture
      End If
   Else
      If ButSz = 0 Then
         DraggerSource.Picture = DraggerSourceH.Picture
      Else
         DraggerSource.Picture = DSH.Picture
      End If
   End If
   PropertyChanged "Style"
   UserControl_Resize
End Property

Public Property Let TickMarks(ByVal NewTickMarks As Boolean)
   Let m_TickMarks = NewTickMarks
   PropertyChanged "TickMarks"
   Redraw
End Property

Public Property Get TickMarkCnt() As Integer
Attribute TickMarkCnt.VB_Description = "The spacing between the marks. The larger the number , the farther apart they are."
   Let TickMarkCnt = m_TickMarkCnt
End Property

Public Property Let TickMarkCnt(ByVal NewTickMarkCnt As Integer)
   Let m_TickMarkCnt = NewTickMarkCnt
   PropertyChanged "TickMarkCnt"
   Redraw
End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "The current cursor position value."
    Value = m_Value
End Property

Public Property Let Value(NewValue As Long)
    PropertyChanged "Value"
    m_Value = NewValue
    Redraw
End Property
