VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Grad_API 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Gradient_API"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3390
   Icon            =   "Grad_API.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   226
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSpeed 
      Height          =   1500
      Left            =   75
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   3210
      Begin MSComctlLib.Slider sldGrad_API 
         Height          =   675
         Left            =   75
         TabIndex        =   1
         Top             =   495
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1191
         _Version        =   393216
         LargeChange     =   1000
         SmallChange     =   100
         Min             =   10
         Max             =   10000
         SelStart        =   1000
         TickFrequency   =   1000
         Value           =   1000
      End
      Begin VB.Label lblFasterSlower 
         Alignment       =   2  'Center
         Caption         =   "  Faster          Slower"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   3
         Top             =   1110
         Width           =   3000
      End
      Begin VB.Label lblDisplaySpeed 
         Alignment       =   2  'Center
         Caption         =   "Display Speed!"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   75
         TabIndex        =   2
         Top             =   75
         Width           =   3000
      End
   End
   Begin VB.Timer tmrGrad_API 
      Interval        =   1000
      Left            =   75
      Top             =   1650
   End
End
Attribute VB_Name = "Grad_API"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Gradient_API uses the 'GradientFillTriangle' API to          ****
'   ****    create two opposed Gradient triangles.  The effect is        ****
'   ****    that you get a screen sized display that is a Gradient in    ****
'   ****    all directions.  By using the API, I can achieve very        ****
'   ****    fast graphics.                                               ****
'   ****                                                                 ****
'   ****    I set it up to use a Timer (tmrGrad_API) to continually      ****
'   ****    refresh the display every second.  I included a popup        ****
'   ****    Speed control, accessible by right-clicking the form, so     ****
'   ****    that the user may control the display speed.  I set the      ****
'   ****    minimum time to 10 milliseconds, but I find that to be       ****
'   ****    too fast.  It causes an annoying flashing.  On the other     ****
'   ****    hand, the maximum of 10 seconds is very slow (too slow       ****
'   ****    for me).                                                     ****
'   ****                                                                 ****
'   ****    The colors of the 4 corners, and the direction that the      ****
'   ****    triangles face, are chosen at random, so the display will    ****
'   ****    be constantly changing.                                      ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

'** Variables **
Private blnButton                      As Boolean

Private GradTriangle                   As GRADIENT_TRIANGLE

Private lngColorNum(0 To 3)            As Long
Private lngCorner                      As Long
Private lngCornerNum(0 To 3)           As Long
Private lngDiagonal                    As Long
Private lngLoop                        As Long
Private lngNum                         As Long
Private lngRGB(0 To 3)                 As Long
Private lngStartTime                   As Long
Private lngStopTime                    As Long
Private lngUorD                        As Long

Private TriVert(0 To 3)                As TRIVERTEX

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    In this mode, an array of TRIVERTEX structures is passed     ****
'   ****    to GDI along with a list of array indexes that describe      ****
'   ****    separate triangles.  GDI performs linear interpolation       ****
'   ****    between triangle vertices and fills the interior.            ****
'   ****    Drawing is done directly in 24- and 32-bpp modes.            ****
'   ****    Dithering is performed in 16-, 8-, 4-, and 1-bpp mode.       ****
'   ****                                                                 ****
'   *************************************************************************
'

Private Const GRADIENT_FILL_TRIANGLE   As Long = &H2

Private Type TRIVERTEX
    X                                  As Long
    Y                                  As Long
    Red                                As Integer   '  Ushort Red value
    Green                              As Integer   '  Ushort Green value
    Blue                               As Integer   '  Ushort Blue value
    Alpha                              As Integer   '  Ushort Alpha value
End Type

Private Type GRADIENT_TRIANGLE
    Vertex1                            As Long      '  In reality this is an UNSIGNED Long
    Vertex2                            As Long      '  In reality this is an UNSIGNED Long
    Vertex3                            As Long      '  In reality this is an UNSIGNED Long
End Type

Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Sub Form_Load()

    Me.Show
    Randomize Timer
    SetCursorPos (Me.Width \ Screen.TwipsPerPixelX) - 99, 10
    lngNum = 1
    lngUorD = -1

'   Define the four corners of the form.
    With TriVert(0)
        .X = 0
        .Y = 0
        .Alpha = 0&
    End With  '  TriVert(0)

    With TriVert(1)
        .X = Me.ScaleWidth
        .Y = 0
        .Alpha = 0&
    End With  '  TriVert(1)

    With TriVert(2)
        .X = Me.ScaleWidth
        .Y = Me.ScaleHeight
        .Alpha = 0&
    End With  '  TriVert(2)

    With TriVert(3)
        .X = 0
        .Y = Me.ScaleHeight
        .Alpha = 0&
    End With  '  TriVert(3)

    For lngLoop = 1 To 3
        lngColorNum(lngLoop) = Int(128 * Rnd) + 128
        lngRGB(lngLoop) = CLng(lngColorNum(lngLoop)) * 256
    Next lngLoop

'    blnButton = False
    subDoubleTriangle

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    A Right Click opens up the Speed Menu.                       ****
'   ****    A Left Click will Freeze (Pause) the display.                ****
'   ****    Left Click again to Thaw the display.                        ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

    If Button = vbRightButton Then
        picSpeed.Visible = True
        lngStartTime = Timer
    Else      '  NOT Button....
        If blnButton = False Then
            blnButton = True
            tmrGrad_API.Enabled = False
            subFreeze
        Else  '  NOT blnButton...
            blnButton = False
            tmrGrad_API.Enabled = True
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Function LongToUShort(Unsigned As Long) As Long

    LongToUShort = CInt(Unsigned - &H10000)

End Function

Private Sub subColor()

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    VB uses the color groups: Red, Green & Blue.                 ****
'   ****    The values for each color group is chosen at random and      ****
'   ****    can be anywhere from 0 to 255.                               ****
'   ****                                                                 ****
'   ****    The 'Select Case' below is used to make the colors even      ****
'   ****    more random by randomly selecting which of the 3 color       ****
'   ****    groups to include in each display.  The possible choices     ****
'   ****    in this program are:                                         ****
'   ****                                                                 ****
'   ****    Red, Green, Blue                                             ****
'   ****    Red                                                          ****
'   ****    Green                                                        ****
'   ****    Blue                                                         ****
'   ****    Red & Blue                                                   ****
'   ****    Green & Blue                                                 ****
'   ****    Red & Green                                                  ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

    Select Case lngCorner
    Case 0
        With TriVert(lngLoop)
            .Red = LongToUShort(lngRGB(1))
            .Green = LongToUShort(lngRGB(2))
            .Blue = LongToUShort(lngRGB(3))
        End With             '  TriVert(lngLoop)

    Case 1
        With TriVert(lngLoop)
            .Red = LongToUShort(lngRGB(1))
            .Green = 0&
            .Blue = 0&
        End With             '  TriVert(lngLoop)

    Case 2
        With TriVert(lngLoop)
            .Red = 0&
            .Green = LongToUShort(lngRGB(2))
            .Blue = 0&
        End With             '  TriVert(lngLoop)

    Case 3
        With TriVert(lngLoop)
            .Red = 0&
            .Green = 0&
            .Blue = LongToUShort(lngRGB(3))
        End With             '  TriVert(lngLoop)

    Case 4
        With TriVert(lngLoop)
            .Red = LongToUShort(lngRGB(1))
            .Green = 0&
            .Blue = LongToUShort(lngRGB(3))
        End With             '  TriVert(lngLoop)

    Case 5
        With TriVert(lngLoop)
            .Red = 0&
            .Green = LongToUShort(lngRGB(2))
            .Blue = LongToUShort(lngRGB(3))
        End With             '  TriVert(lngLoop)

    Case 6
        With TriVert(lngLoop)
            .Red = LongToUShort(lngRGB(1))
            .Green = LongToUShort(lngRGB(2))
            .Blue = 0&
        End With             '  TriVert(lngLoop)
    End Select

    lngColorNum(lngNum) = lngColorNum(lngNum) + lngUorD

End Sub

Private Sub subDoubleTriangle()

'   Select new random colors
    For lngLoop = 1 To 3
        lngColorNum(lngLoop) = Int(128 * Rnd) + 128
        lngRGB(lngLoop) = CLng(lngColorNum(lngLoop)) * 256
    Next lngLoop

'   Choose which colors are to be used in the display.
    For lngLoop = 0 To 3
        lngCorner = Int(7 * Rnd)
        lngCornerNum(lngLoop) = lngCorner
        subColor
    Next lngLoop

    If GetInputState Then
        DoEvents       '  Yield for other processing.
    End If

    lngDiagonal = Int(2 * Rnd) + 1

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    This routine is where I draw the                             ****
'   ****    Double Gradient Triangles.                                   ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

    Select Case lngDiagonal
    Case 1
        With GradTriangle
            .Vertex1 = 0
            .Vertex2 = 1
            .Vertex3 = 2
            GradientFillTriangle Me.hdc, TriVert(0), 4, GradTriangle, 1, GRADIENT_FILL_TRIANGLE
            .Vertex1 = 2
            .Vertex2 = 3
            .Vertex3 = 0
            GradientFillTriangle Me.hdc, TriVert(0), 4, GradTriangle, 1, GRADIENT_FILL_TRIANGLE
        End With             '  GradTriangle

    Case 2
        With GradTriangle
            .Vertex1 = 1
            .Vertex2 = 2
            .Vertex3 = 3
            GradientFillTriangle Me.hdc, TriVert(0), 4, GradTriangle, 1, GRADIENT_FILL_TRIANGLE
            .Vertex1 = 3
            .Vertex2 = 0
            .Vertex3 = 1
            GradientFillTriangle Me.hdc, TriVert(0), 4, GradTriangle, 1, GRADIENT_FILL_TRIANGLE
        End With             '  GradTriangle
    End Select

    Me.Refresh

    If GetInputState Then
        DoEvents       '  Yield for other processing.
    End If

'   Hide the Speed control if it's been visible for 4 seconds.
    lngStopTime = Timer

    If lngStopTime > lngStartTime + 4 Then
        picSpeed.Visible = False
    End If

End Sub

Private Sub subFreeze()

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    The Pause routine is activated by pressing a Mouse           ****
'   ****    button.  Press a button once to Freeze (Pause) the           ****
'   ****    screen.  Press a button again to Thaw the screen.            ****
'   ****                                                                 ****
'   *************************************************************************
'

    Do While blnButton
        DoEvents             '  Yield for other processing.
    Loop

End Sub

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    These are the 'Slider' routines for the Speed control.       ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

Private Sub sldGrad_API_Change()

    tmrGrad_API.Interval = sldGrad_API.Value

End Sub

Private Sub sldGrad_API_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lngStartTime = Timer

End Sub

Private Sub tmrGrad_API_Timer()

    subDoubleTriangle

End Sub
':)Code Fixer V3.0.7 (3/24/2005 2:50:13 PM) 48 + 172 = 220 Lines Thanks Ulli for inspiration and lots of code.
