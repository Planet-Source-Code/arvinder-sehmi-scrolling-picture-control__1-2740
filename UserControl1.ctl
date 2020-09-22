VERSION 5.00
Begin VB.UserControl ScrollControl 
   BackColor       =   &H00000000&
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   1500
   Begin VB.PictureBox ParentPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   495
      ScaleHeight     =   1320
      ScaleWidth      =   2265
      TabIndex        =   2
      Top             =   1845
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1560
      Left            =   1890
      ScaleHeight     =   1560
      ScaleWidth      =   2145
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Scroll Control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   -225
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image PrevImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   1305
      Left            =   120
      Stretch         =   -1  'True
      Top             =   270
      Width           =   1245
   End
End
Attribute VB_Name = "ScrollControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------'
' The BitBlt Api Call'
'--------------------'
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'---------------------------------'
' Declare some groups of constants'
'---------------------------------'
Public Enum ScrollDirection 'These Are Used in the Scroll Direction Property
 Right_To_Left = 1
 Left_To_Right = 2
 Bottom_To_Top = 3
 Top_To_Bottom = 4
 BottomRight_To_TopLeft = 5
 TopLeft_To_BottomRight = 6
 TopRight_To_BottomLeft = 7
 BottomLeft_To_TopRight = 8
End Enum
Public Enum BitBlt_dwROP    'The Different BitBlt Types
 Src_And = &H8800C6
 Src_Copy = &HCC0020
 Src_Invert = &H660046
 Src_Paint = &HEE0086
 Src_Erase = &H440328
 Not_Src_Copy = &H330008
 Not_Src_Erase = &H1100A6
End Enum
'--------------------'
'Dim Some Stuff'
'--------------------'
Dim X As Integer, Y As Integer 'Position Of The Scrolling Image
Dim Back As Integer, Wdth As Integer, Hght As Integer 'Holds Info for The Parents Size
Dim PicWdth As Integer, PicHght As Integer, Direct As Integer ' Holds Info From the Picture to Blt
'Property Variables:
Dim m_BitBltStyle As Variant
Dim m_Direction As ScrollDirection
Dim ExitIt As Boolean ' Tells The Scrolling When To Stop
'Default Property Values:
Const m_def_BitBltStyle = &HCC0020
Const m_def_Direction = 1
Public Sub Stop_Scroll() ' Stop Scrolling
 ExitIt = True
End Sub
Public Sub Start_Scroll() 'Start Scrolling
  '---------------------'
  'Get Some Picture Info'
  '---------------------'
  PicWdth = Int(Pic.Width / Screen.TwipsPerPixelX)
  PicHght = Int(Pic.Height / Screen.TwipsPerPixelY)
  Pic.BackColor = UserControl.Parent.BackColor
  ExitIt = False
  On Error GoTo Nd:
  Do 'Start The Loop
   For Back = 0 To PicWdth
    For Y = -PicHght To Hght Step PicHght ' The Y-Blt Pos
     For X = -PicWdth To Wdth Step PicWdth ' The X-Blt Pos
      Wdth = Int(UserControl.Parent.Width / Screen.TwipsPerPixelX) + PicWdth  'Update Parents Width Variable
      Hght = Int(UserControl.Parent.Height / Screen.TwipsPerPixelY) + PicHght 'Update Parents Height Variable
      BltType = m_BitBltStyle
      Select Case m_Direction 'Only Blt To The Correct Direction
       Case 1 ' Right->Left
        BitBlt UserControl.Parent.hDC, X - Back, Y, PicWdth, PicHght, Pic.hDC, 0, 0, BltType
       Case 2 ' Left->Right
        BitBlt UserControl.Parent.hDC, X + Back, Y, PicWdth, PicHght, Pic.hDC, 0, 0, BltType
       Case 3 ' Bottom->Top
        BitBlt UserControl.Parent.hDC, X, Y - Back, PicWdth, PicHght, Pic.hDC, 0, 0, BltType
       Case 4 ' Top->Bottom
        BitBlt UserControl.Parent.hDC, X, Y + Back, PicWdth, PicHght, Pic.hDC, 0, 0, BltType
       Case 5 ' BottomRight->TopLeft
        BitBlt UserControl.Parent.hDC, X - Back, Y - Back, PicWdth, PicHght, Pic.hDC, 0, 0, BltType
       Case 6 ' TopLeft->BottomRight
        BitBlt UserControl.Parent.hDC, X + Back, Y + Back, PicWdth, PicHght, Pic.hDC, 0, 0, BltType
       Case 7 ' TopRight->BottomLeft
        BitBlt UserControl.Parent.hDC, X - Back, Y + Back, PicWdth, PicHght, Pic.hDC, 0, 0, BltType
       Case 8 ' BottomLeft->TopRight
        BitBlt UserControl.Parent.hDC, X + Back, Y - Back, PicWdth, PicHght, Pic.hDC, 0, 0, BltType
      End Select
      DoEvents
     Next X
    Next Y
   If ExitIt = True Then Exit Sub
   Next Back
  Loop Until ExitIt = True
Nd:
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,0,0,1
Public Property Get Direction() As ScrollDirection
    Direction = m_Direction
End Property
Public Property Let Direction(ByVal New_Direction As ScrollDirection)
    m_Direction = New_Direction
    PropertyChanged "Direction"
End Property
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Direction = m_def_Direction
    m_BitBltStyle = m_def_BitBltStyle
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Direction = PropBag.ReadProperty("Direction", m_def_Direction)
    m_BitBltStyle = PropBag.ReadProperty("BitBltStyle", m_def_BitBltStyle)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub
Private Sub UserControl_Resize()
 UserControl.Width = 1500
 UserControl.Height = 1740
End Sub
Private Sub UserControl_Terminate()
 ExitIt = True 'Exit The Scrolling Loop
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Direction", m_Direction, m_def_Direction)
    Call PropBag.WriteProperty("BitBltStyle", m_BitBltStyle, m_def_BitBltStyle)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

'-------------------------------------'
'Enables the User To Change Blt Styles'
'  and To Change The Scrolling Image  '
'-------------------------------------'
Public Property Get BitBltStyle() As BitBlt_dwROP
Attribute BitBltStyle.VB_Description = "The dwRop That the BitBlt Api Uses."
    BitBltStyle = m_BitBltStyle
End Property

Public Property Let BitBltStyle(ByVal New_BitBltStyle As BitBlt_dwROP)
    m_BitBltStyle = New_BitBltStyle
    PropertyChanged "BitBltStyle"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = Pic.Picture
    Set PrevImage = Pic.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Pic.Picture = New_Picture
    PropertyChanged "Picture"
End Property

