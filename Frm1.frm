VERSION 5.00
Object = "*\AProject2.vbp"
Begin VB.Form Form1 
   Caption         =   "Test Picture Scrolling Project "
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin PictureScrollControl.ScrollControl ScrollControl2 
      Left            =   1980
      Top             =   270
      _ExtentX        =   2646
      _ExtentY        =   3069
      Picture         =   "Frm1.frx":0000
   End
   Begin PictureScrollControl.ScrollControl ScrollControl1 
      Left            =   270
      Top             =   540
      _ExtentX        =   2646
      _ExtentY        =   3069
      Picture         =   "Frm1.frx":08E3
   End
   Begin VB.ComboBox Direction 
      Height          =   315
      Left            =   2385
      TabIndex        =   9
      Text            =   "Direction"
      Top             =   4050
      Width           =   2265
   End
   Begin VB.ComboBox BitBltStyle 
      Height          =   315
      Left            =   90
      TabIndex        =   8
      Text            =   "BitBlt Style"
      Top             =   4050
      Width           =   2220
   End
   Begin VB.PictureBox Picture2 
      Height          =   780
      Left            =   2385
      ScaleHeight     =   720
      ScaleWidth      =   2205
      TabIndex        =   1
      Top             =   3195
      Width           =   2265
      Begin VB.CommandButton Command4 
         Caption         =   "Stop"
         Height          =   330
         Left            =   1170
         TabIndex        =   5
         Top             =   315
         Width           =   870
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Start"
         Height          =   330
         Left            =   135
         TabIndex        =   3
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Transparent Gif"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   45
         Width           =   2040
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   780
      Left            =   90
      ScaleHeight     =   720
      ScaleWidth      =   2160
      TabIndex        =   0
      Top             =   3195
      Width           =   2220
      Begin VB.CommandButton Command3 
         Caption         =   "Stop"
         Height          =   330
         Left            =   1170
         TabIndex        =   4
         Top             =   315
         Width           =   870
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   330
         Left            =   90
         TabIndex        =   2
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Standard Picture"
         Height          =   195
         Left            =   45
         TabIndex        =   6
         Top             =   45
         Width           =   2040
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'for help contact arvinder@bigfoot.com
Private Sub BitBltStyle_Click()
Select Case BitBltStyle.Text
 Case "Src_And"
  ScrollControl1.BitBltStyle = Src_And
 Case "Src_Copy"
  ScrollControl1.BitBltStyle = Src_Copy
 Case "Src_Invert"
  ScrollControl1.BitBltStyle = Src_Invert
 Case "Src_Paint"
  ScrollControl1.BitBltStyle = Src_Paint
 Case "Src_Erase"
  ScrollControl1.BitBltStyle = Src_Erase
 Case "Not_Src_Copy"
  ScrollControl1.BitBltStyle = Not_Src_Copy
 Case "Not_Src_Erase"
  ScrollControl1.BitBltStyle = Not_Src_Erase
End Select
ScrollControl2.BitBltStyle = ScrollControl1.BitBltStyle
End Sub
Private Sub Command1_Click()
 ScrollControl1.Start_Scroll
End Sub
Private Sub Command2_Click()
 ScrollControl2.Start_Scroll
End Sub
Private Sub Command3_Click()
 ScrollControl1.Stop_Scroll
End Sub
Private Sub Command4_Click()
 ScrollControl2.Stop_Scroll
End Sub
Private Sub Direction_Click()
Select Case Direction.Text
 Case "1 Right_To_Left"
  ScrollControl1.Direction = Right_To_Left
 Case "2 Left_To_Right"
  ScrollControl1.Direction = Left_To_Right
 Case "3 Bottom_To_Top"
  ScrollControl1.Direction = Bottom_To_Top
 Case "4 Top_To_Bottom"
  ScrollControl1.Direction = Top_To_Bottom
 Case "5 BottomRight_To_TopLeft"
  ScrollControl1.Direction = BottomRight_To_TopLeft
 Case "6 TopLeft_To_BottomRight"
  ScrollControl1.Direction = TopLeft_To_BottomRight
 Case "7 TopRight_To_BottomLeft"
  ScrollControl1.Direction = TopRight_To_BottomLeft
 Case "8 BottomLeft_To_TopRight"
  ScrollControl1.Direction = BottomLeft_To_TopRight
End Select
ScrollControl2.Direction = ScrollControl1.Direction
End Sub
Private Sub Form_Load()
With Direction
 .AddItem "1 Right_To_Left"
 .AddItem "2 Left_To_Right"
 .AddItem "3 Bottom_To_Top"
 .AddItem "4 Top_To_Bottom"
 .AddItem "5 BottomRight_To_TopLeft"
 .AddItem "6 TopLeft_To_BottomRight"
 .AddItem "7 TopRight_To_BottomLeft"
 .AddItem "8 BottomLeft_To_TopRight"
End With
With BitBltStyle
 .AddItem "Src_And"
 .AddItem "Src_Copy"
 .AddItem "Src_Invert"
 .AddItem "Src_Paint"
 .AddItem "Src_Erase"
 .AddItem "Not_Src_Copy"
 .AddItem "Not_Src_Erase"
End With
End Sub

Private Sub ScrollControl2_Click()

End Sub
