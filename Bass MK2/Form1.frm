VERSION 5.00
Begin VB.Form frmBassMK2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Bass Maker 2 - Updated Version«"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7185
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pboxGraphWaveNOScaleForAmplitude 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   5295
      TabIndex        =   14
      ToolTipText     =   "This is the Wave Number scale, move the scroll bar above to change this scale"
      Top             =   4470
      Width           =   5295
   End
   Begin VB.PictureBox pboxLargeGraphAmplitude 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2430
      Left            =   1440
      MouseIcon       =   "Form1.frx":1272
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":13C4
      ScaleHeight     =   2430
      ScaleWidth      =   5295
      TabIndex        =   13
      Top             =   4770
      Width           =   5295
      Begin VB.Label labelTip2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click here to make a graph of the Amplitude/Volume"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1575
         Left            =   660
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   3945
      End
      Begin VB.Line lineV2CrossFollowingMouseInAmplitude 
         BorderColor     =   &H00008000&
         X1              =   -1000
         X2              =   -1001
         Y1              =   1080
         Y2              =   1500
      End
      Begin VB.Line lineV1CrossFollowingMouseInAmplitude 
         BorderColor     =   &H00008000&
         X1              =   -1000
         X2              =   -1001
         Y1              =   750
         Y2              =   1320
      End
      Begin VB.Line lineH2CrossFollowingMouseInAmplitude 
         BorderColor     =   &H00008000&
         X1              =   -1000
         X2              =   -1001
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line lineH1CrossFollowingMouseInAmplitude 
         BorderColor     =   &H00008000&
         X1              =   -1000
         X2              =   -1001
         Y1              =   780
         Y2              =   900
      End
   End
   Begin VB.PictureBox pboxLargeGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   1440
      MouseIcon       =   "Form1.frx":2AEAA
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":2AFFC
      ScaleHeight     =   2415
      ScaleWidth      =   5295
      TabIndex        =   0
      Top             =   1560
      Width           =   5295
      Begin VB.Label labelTip1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click here to make a graph of the frequency"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1245
         Left            =   690
         TabIndex        =   19
         Top             =   810
         Visible         =   0   'False
         Width           =   3945
      End
      Begin VB.Line lineH2CrossFollowingMouse 
         BorderColor     =   &H00008000&
         X1              =   9100
         X2              =   8640
         Y1              =   9340
         Y2              =   9340
      End
      Begin VB.Line lineV2CrossFollowingMouse 
         BorderColor     =   &H00008000&
         X1              =   9100
         X2              =   9100
         Y1              =   2340
         Y2              =   3480
      End
      Begin VB.Line lineV1CrossFollowingMouse 
         BorderColor     =   &H00008000&
         X1              =   7260
         X2              =   7260
         Y1              =   9940
         Y2              =   9080
      End
      Begin VB.Line lineH1CrossFollowingMouse 
         BorderColor     =   &H00008000&
         X1              =   5580
         X2              =   9120
         Y1              =   3600
         Y2              =   3600
      End
   End
   Begin VB.HScrollBar hscrolGraphWaveNO 
      Height          =   195
      LargeChange     =   1000
      Left            =   1410
      Max             =   9700
      SmallChange     =   100
      TabIndex        =   6
      Top             =   4050
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.PictureBox pboxGraphWaveNOScale 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   5295
      TabIndex        =   3
      ToolTipText     =   "This is the Wave Number scale, move the scroll bar below to change this scale"
      Top             =   3960
      Width           =   5295
   End
   Begin VB.PictureBox pboxMiniGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1440
      ScaleHeight     =   735
      ScaleWidth      =   5295
      TabIndex        =   2
      ToolTipText     =   "This is a Mini Graph of your whole project"
      Top             =   840
      Width           =   5295
   End
   Begin VB.PictureBox pboxGraphScaleHertz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   180
      ScaleHeight     =   2415
      ScaleWidth      =   1275
      TabIndex        =   1
      ToolTipText     =   "This is the Hertz Scale, double click to change this scale"
      Top             =   1560
      Width           =   1275
   End
   Begin VB.PictureBox pboxWavePreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   1035
      Left            =   60
      Picture         =   "Form1.frx":54AE2
      ScaleHeight     =   1035
      ScaleWidth      =   1815
      TabIndex        =   9
      Top             =   60
      Width           =   1815
      Begin VB.Timer timerWavePreview 
         Interval        =   200
         Left            =   1140
         Top             =   120
      End
   End
   Begin VB.PictureBox pboxGraphScaleAmplitude 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   180
      ScaleHeight     =   2415
      ScaleWidth      =   1275
      TabIndex        =   15
      ToolTipText     =   "This is the Amplitude (or Volume) of the current wave"
      Top             =   4770
      Width           =   1275
   End
   Begin VB.PictureBox pboxMiniGraphAmplitude 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1440
      ScaleHeight     =   735
      ScaleWidth      =   5295
      TabIndex        =   18
      ToolTipText     =   "This is a Mini Graph of your whole project"
      Top             =   7185
      Width           =   5295
   End
   Begin VB.Label labelTip4 
      BackStyle       =   0  'Transparent
      Caption         =   "When your done, click file then make wave"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1245
      Left            =   1860
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label labelTip3 
      BackStyle       =   0  'Transparent
      Caption         =   "< Slide this bar to change the wave no."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   795
      Left            =   2580
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   4485
   End
   Begin VB.Label labelGraphAmplitude 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amplitude (Volume)"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   270
      TabIndex        =   17
      Top             =   7410
      Width           =   1275
   End
   Begin VB.Label labelWaveNOScaleAmplitudeLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Wave Number >"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4530
      Width           =   1275
   End
   Begin VB.Label labelHScrollMovingBar 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1620
      TabIndex        =   12
      Top             =   4260
      Width           =   915
   End
   Begin VB.Shape shapeHScrollMovingBar 
      BackColor       =   &H0000C000&
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   1620
      Top             =   4275
      Width           =   915
   End
   Begin VB.Label labelHScrollRight 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   6540
      TabIndex        =   11
      Top             =   4260
      Width           =   195
   End
   Begin VB.Label labelHScrollLeft 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1440
      TabIndex        =   10
      Top             =   4260
      Width           =   195
   End
   Begin VB.Line lineHScrollBarLine12 
      BorderColor     =   &H0000FF00&
      X1              =   6700
      X2              =   6600
      Y1              =   4365
      Y2              =   4395
   End
   Begin VB.Line lineHScrollBarLine11 
      BorderColor     =   &H0000FF00&
      X1              =   6695
      X2              =   6600
      Y1              =   4355
      Y2              =   4320
   End
   Begin VB.Line lineHScrollBarLine10 
      BorderColor     =   &H0000FF00&
      X1              =   6600
      X2              =   6600
      Y1              =   4320
      Y2              =   4410
   End
   Begin VB.Line lineHScrollBarLine9 
      BorderColor     =   &H0000FF00&
      X1              =   6540
      X2              =   6540
      Y1              =   4260
      Y2              =   4440
   End
   Begin VB.Line lineHScrollBarLine8 
      BorderColor     =   &H0000FF00&
      X1              =   6720
      X2              =   6720
      Y1              =   4260
      Y2              =   4440
   End
   Begin VB.Line lineHScrollBarLine7 
      BorderColor     =   &H0000FF00&
      X1              =   1560
      X2              =   1485
      Y1              =   4395
      Y2              =   4365
   End
   Begin VB.Line lineHScrollBarLine6 
      BorderColor     =   &H0000FF00&
      X1              =   1560
      X2              =   1455
      Y1              =   4320
      Y2              =   4380
   End
   Begin VB.Line lineHScrollBarLine5 
      BorderColor     =   &H0000FF00&
      X1              =   1560
      X2              =   1560
      Y1              =   4320
      Y2              =   4410
   End
   Begin VB.Line lineHScrollBarLine4 
      BorderColor     =   &H0000FF00&
      X1              =   1620
      X2              =   1620
      Y1              =   4260
      Y2              =   4440
   End
   Begin VB.Line lineHScrollBarLine3 
      BorderColor     =   &H0000FF00&
      X1              =   1440
      X2              =   1440
      Y1              =   4260
      Y2              =   4440
   End
   Begin VB.Line lineHScrollBarLine2 
      BorderColor     =   &H0000FF00&
      X1              =   1440
      X2              =   6735
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line lineHScrollBarLine1 
      BorderColor     =   &H0000FF00&
      X1              =   1440
      X2              =   6720
      Y1              =   4275
      Y2              =   4275
   End
   Begin VB.Label labelCurrentValue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Value = 0"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   510
      Width           =   2235
   End
   Begin VB.Line lineDotFixerLine 
      BorderColor     =   &H0000FF00&
      Tag             =   "fixes a little spot on the axes where the X-Y meet"
      X1              =   1425
      X2              =   3075
      Y1              =   3975
      Y2              =   2805
   End
   Begin VB.Label labelGraphHertz 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hertz"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label labelWaveNOScaleLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Wave Number >"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4020
      Width           =   1275
   End
   Begin VB.Image imageBassMK2 
      Height          =   600
      Left            =   2340
      Picture         =   "Form1.frx":57FA4
      Stretch         =   -1  'True
      ToolTipText     =   "Welcome to BASS MK2"
      Top             =   60
      Width           =   2340
   End
   Begin VB.Label labelCurrentCords 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "( 0 , 0 )"
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   4680
      TabIndex        =   7
      Top             =   60
      Width           =   2475
   End
   Begin VB.Menu menuFileMenu 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu menuNewProj 
         Caption         =   "New Project"
         Index           =   8
         Shortcut        =   ^N
      End
      Begin VB.Menu menuSaveMenu 
         Caption         =   "Save Project"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu menuLoadMenu 
         Caption         =   "Load Project"
         Index           =   3
         Shortcut        =   ^L
      End
      Begin VB.Menu menuMakeWave 
         Caption         =   "Make Wave File"
         Index           =   4
         Shortcut        =   ^W
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
         Index           =   5
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu menugraph 
      Caption         =   "Graph"
      Index           =   6
      Begin VB.Menu menucngHertzScale 
         Caption         =   "Change Hertz Scale"
         Index           =   7
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu menuOptions 
      Caption         =   "Options"
      Index           =   8
      Begin VB.Menu menuAmplitudeClear 
         Caption         =   "Clear Amplitude Graph"
         Index           =   11
      End
      Begin VB.Menu menuClearFreq 
         Caption         =   "Clear Frequency Graph"
         Index           =   12
      End
      Begin VB.Menu menuAminStopStart 
         Caption         =   "Stop/Start Speaker Animation"
         Index           =   9
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "Help"
      Index           =   10
   End
End
Attribute VB_Name = "frmBassMK2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
'Bass Maker 2 (Updated Version)
'----------------------------------------------------------
'
'Updates
'----------------------------------------------------------
'» Use of AllAPI api's code for open dialogs instead of
'  Windows CommonDialog
'» Fixed Code, Algorithms run Much faster
'» A few new visual features added including a new Scroll
'  bar to suit the rest of the Form
'» User now able to change Amplitude
'» Better Drawing methods, you can click and drag to plot
'  points instead of only clicking
'----------------------------------------------------------
'
'Known Glitches in program
'----------------------------------------------------------
'« The Amplitude graph does not always draw properly
'----------------------------------------------------------


Dim arrOldFreqVAlue&(1)
Dim arrOldAmplitudeVAlue&(1)
'stores the current X,Y mouse cords in terms of hertz and
'wave no.
Dim intCurrentYCordS%
Dim intCurrentXCordS%
Dim intCurrentYCordSAmplitude%
Dim intCurrentXCordSAmplitude%
'used to store length of the Divisions on the X-axis of
'the graph
Dim longintDivision#
'temp integer (it does come in handy for returning values
'over two different subs)
Dim intTempHolder%
Dim inttempHolder2%

'set up the form
Private Sub Form_Load()
'set the wave No scale
hscrolGraphWaveNO_Change
'set the hertz scale
intMaxHertzScale% = 100
subFixHertzAxes
For I = 0 To 10000
    arrPointsForAmplitude&(I) = 100
Next I
'this If statement stops the program from redrawing the
'amplitude scale again when you select new project from the
'start menu (otherwise it just keeps getting bigger)
If Me.Tag = "" Then
    'clear the picture box
    pboxGraphScaleAmplitude.Cls
    'just to make coding easier for me I incresed the
    'amplitude's scale picture box by 0.1%
    longintTempHeight% = pboxGraphScaleAmplitude.Height
    pboxGraphScaleAmplitude.Height = pboxGraphScaleAmplitude.Height * 1.1
    For I = 0 To 1 Step 0.1
        'set the point where the label is going to be placed
        pboxGraphScaleAmplitude.ForeColor = RGB(0, 0, 0)
        pboxGraphScaleAmplitude.PSet (15, longintTempHeight% * I)
        'set the text for the Amplitude axis
        pboxGraphScaleAmplitude.ForeColor = RGB(0, 255, 0)
        pboxGraphScaleAmplitude.Print Str$(I * 100) + "%"
        'draw some lines to make it easier to see where you are
        'on the graph
        pboxGraphScaleAmplitude.Line (pboxGraphScaleAmplitude.Width - 800, longintTempHeight% * (I + 0.1))-(pboxGraphScaleAmplitude.Width - 15, longintTempHeight% * (I + 0.1))
        pboxGraphScaleAmplitude.Line (pboxGraphScaleAmplitude.Width - 200, longintTempHeight% * I - longintTempHeight% * 0.05)-(pboxGraphScaleAmplitude.Width - 15, longintTempHeight% * I - longintTempHeight% * 0.05)
    Next I
    'draw the vertical line to separate the scale from the
    'graph
    pboxGraphScaleAmplitude.Line (pboxGraphScaleAmplitude.Width - 30, 0)-(pboxGraphScaleAmplitude.Width - 30, longintTempHeight%)
    Me.Tag = "Already Drawn Amplitude Scale"
    'fixes a problem with the thumb view, doesn't initially draw
    'properly
    subDrawThumbViewAmplitude
End If
End Sub

'---------------------------------------------------------
'when a user clicks and drag's the bar on the Hscroll bar
'then move the bar's centre to where the mouse is, then
'adjust the new value of the Hscroll bar
'---------------------------------------------------------
Private Sub labelHScrollMovingBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Call labelHScrollMovingBar_MouseMove(1, 0, X, 0)
End If
End Sub
Private Sub labelHScrollMovingBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    labelHScrollMovingBar.Left = labelHScrollMovingBar.Left + X - (labelHScrollMovingBar.Width / 2)
    If labelHScrollMovingBar.Left < 1620 Then labelHScrollMovingBar.Left = 1620
    If labelHScrollMovingBar.Left > 5640 Then labelHScrollMovingBar.Left = 5640
    shapeHScrollMovingBar.Left = labelHScrollMovingBar.Left
    hscrolGraphWaveNO.Value = (labelHScrollMovingBar.Left - 1620) / 402 * 970
End If
End Sub
'---------------------------------------------------------

'---------------------------------------------------------
'These are when the user clicks on the left or right arrows
'on the "Graphical Style" Hscroll bar (it still uses the
'old Hscroll bar's subroutines)
'---------------------------------------------------------
Private Sub labelHScrollLeft_Click()
If hscrolGraphWaveNO.Value - 100 < 0 Then hscrolGraphWaveNO.Value = 100
hscrolGraphWaveNO.Value = hscrolGraphWaveNO.Value - 100
End Sub
Private Sub labelHScrollLeft_DblClick()
labelHScrollLeft_Click
labelHScrollLeft_Click
End Sub
Private Sub labelHScrollRight_Click()
If hscrolGraphWaveNO.Value + 100 > 9700 Then hscrolGraphWaveNO.Value = 9600
hscrolGraphWaveNO.Value = hscrolGraphWaveNO.Value + 100
End Sub
Private Sub labelHScrollRight_DblClick()
labelHScrollRight_Click
labelHScrollRight_Click
End Sub
'---------------------------------------------------------

'---------------------------------------------------------
' if the user moves the mouse over the tips, then make them
'dissappear
'---------------------------------------------------------
Private Sub labelTip1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
labelTip1.Visible = False
End Sub
Private Sub labelTip2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
labelTip2.Visible = False
End Sub
Private Sub labelTip3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
labelTip3.Visible = False
End Sub
Private Sub labelTip4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
labelTip4.Visible = False
End Sub
'---------------------------------------------------------


'enable/disable the speaker animation
Private Sub menuAminStopStart_Click(Index As Integer)
If timerWavePreview = True Then
    timerWavePreview.Enabled = False
Else
    timerWavePreview.Enabled = True
End If
End Sub

Private Sub menuAmplitudeClear_Click(Index As Integer)
'reset the amplitude array
For I = 0 To 10000
    arrPointsForAmplitude&(I) = 100
Next I
'redraw the form
hscrolGraphWaveNO_Change
End Sub

Private Sub menuClearFreq_Click(Index As Integer)
'reset the frequency array
For I = 0 To 10000
    arrPointsForFreq&(I) = 0
Next I
'redraw the form
hscrolGraphWaveNO_Change
End Sub

'---------------------------------------------------------
'Exit Bass Maker
'---------------------------------------------------------
'User Selects Exit from the drop down menu
Private Sub menuExit_Click(Index As Integer)
'initiate unload form (jumps to Form_Unload sub)
Unload Me
End Sub
'When unloading the form, show a confirmation message
Private Sub Form_Unload(Cancel As Integer)
'get confirmation from the user
Response$ = MsgBox("Are You Sure?", vbYesNo, "Confirm Exit")
'if user clicks yes then exit bass maker
If Response$ = vbYes Then End
'stop the form unloading
Cancel = 1
End Sub
'----------------------------------------------------------

'----------------------------------------------------------
'Change the Hertz Scale
'----------------------------------------------------------
Private Sub menucngHertzScale_Click(Index As Integer)
pboxGraphScaleHertz_DblClick
End Sub

Private Sub menuHelp_Click(Index As Integer)
'timerTipsDissapear.Enabled = True

'if the user wants help then show the tips
labelTip1.Visible = True
labelTip2.Visible = True
labelTip3.Visible = True
labelTip4.Visible = True
End Sub

Private Sub menuLoadMenu_Click(Index As Integer)
On Error Resume Next

'show the save project form
strFileToOpen$ = "Bass Maker 2 Files (*.bs2)"
strTypeToOpen$ = "*.bs2"
Me.Enabled = False
frmLoad.Show

'if cancel was pressed then exit sub
If strFileToOpen$ = "////" Then Exit Sub

'close all files
Close
'open the selected file for reading
Open strFileToOpen$ For Input As #1
'read all values from the file
Input #1, A$
intMaxHertzScale% = Val(A$)
For I = 0 To 10000
    Input #1, A$
    arrPointsForFreq&(I) = Val(A$)
Next I
For I = 0 To 10000
    Input #1, A$
    arrPointsForAmplitude&(I) = Val(A$)
Next I
'close all files
Close
'refresh form
subFixHertzAxes
hscrolGraphWaveNO_Change
End Sub

Private Sub menuMakeWave_Click(Index As Integer)
'if you don't put error trapping on, you will get and error
'when you cancel selecting the wave file to save to
On Error Resume Next
strFileToOpen$ = "Wave Files (*.wav)"
strTypeToOpen$ = "*.wav"
Me.Enabled = False
frmSave.Show

'if cancel was pressed then exit sub
If strFileToOpen$ = "////" Then Exit Sub

'if filename doesn't end with .wav then change it
If LCase$(Mid$(strFileToOpen$, Len(strFileToOpen$) - 4, 4)) <> ".wav" Then
    strFileToOpen$ = Mid(strFileToOpen$, 1, Len(strFileToOpen$) - 1) + ".wav"
End If

Me.Enabled = False
frmMakeWave.Show
End Sub

Private Sub menuNewProj_Click(Index As Integer)
'set all points to zero
For I = 0 To 10000
    arrPointsForFreq&(I) = 0
Next I
For I = 0 To 10000
    arrPointsForAmplitude&(I) = 100
Next I
'redraw the form
Form_Load
End Sub


Private Sub menuSaveMenu_Click(Index As Integer)
On Error Resume Next

'show the save project form
strFileToOpen$ = "Bass Maker 2 Files (*.bs2)"
strTypeToOpen$ = "*.bs2"
Me.Enabled = False
frmSave.Show

'if cancel was pressed then exit sub
If strFileToOpen$ = "////" Then Exit Sub

'close all files
Close
'if filename doesn't end with .bs2 then change it
If LCase$(Mid$(strFileToOpen$, Len(strFileToOpen$) - 4, 4)) <> ".bs2" Then
    strFileToOpen$ = Mid(strFileToOpen$, 1, Len(strFileToOpen$) - 1) + ".bs2"
End If
'open the selected file for writing to
Open strFileToOpen$ For Output As #1
'save all data to a file, the ';' stops vb from putting
'an enter for each print statement
Print #1, Trim$(Str$(intMaxHertzScale%)) + ",";
For I = 0 To 10000
    Print #1, Trim$(Str$(arrPointsForFreq&(I))) + ",";
Next I
For I = 0 To 10000
    Print #1, Trim$(Str$(arrPointsForAmplitude&(I))) + ",";
Next I
'close all files
Close
'notify the user
Call MsgBox("Project Saved to " + strFileToOpen$, vbInformation, "Save Complete")
End Sub


Private Sub pboxGraphScaleHertz_DblClick()
120 'show a inputbox to get a value for the new scale
TempString$ = InputBox("What is the new Hertz scale you wish set? (100-minimum, 22050-maximum)", "New Hertz Scale", "500Hz")
'if the value is out of bounds then ask for it again
If Val(TempString$) = 0 Then Exit Sub
If Val(TempString$) < 100 Then GoTo 120
If Val(TempString$) > 22050 Then GoTo 120
'set new value
intMaxHertzScale% = Val(TempString$)
subFixHertzAxes
'redraw graph
subDrawGraph
End Sub
'draw the new hertz scale
Private Sub subFixHertzAxes()
'clear the pciture box
pboxGraphScaleHertz.Cls
For I = 0 To 1 Step 0.1
    'set the point where the label is going to be placed
    '(you used to use locate in Qbasic and Gwbasic instead
    'of pset ,but the locate command has since gone in VB)
    pboxGraphScaleHertz.ForeColor = RGB(0, 0, 0)
    pboxGraphScaleHertz.PSet (15, pboxGraphScaleHertz.Height - pboxGraphScaleHertz.Height * I)
    'set the text for the hertz axis
    pboxGraphScaleHertz.ForeColor = RGB(0, 255, 0)
    pboxGraphScaleHertz.Print (I * intMaxHertzScale%)
    'draw some lines to make it easier to see where you are
    'on the graph
    pboxGraphScaleHertz.Line (pboxGraphScaleHertz.Width - 800, pboxGraphScaleHertz.Height - pboxGraphScaleHertz.Height * I)-(pboxGraphScaleHertz.Width, pboxGraphScaleHertz.Height - pboxGraphScaleHertz.Height * I)
    pboxGraphScaleHertz.Line (pboxGraphScaleHertz.Width - 200, pboxGraphScaleHertz.Height - pboxGraphScaleHertz.Height * I - pboxGraphScaleHertz.Height * 0.05)-(pboxGraphScaleHertz.Width, pboxGraphScaleHertz.Height - pboxGraphScaleHertz.Height * I - pboxGraphScaleHertz.Height * 0.05)
Next I
'draw the vertical line to separate the scale from the
'graph
pboxGraphScaleHertz.Line (pboxGraphScaleHertz.Width - 30, 0)-(pboxGraphScaleHertz.Width - 30, pboxGraphScaleHertz.Height)
End Sub
'----------------------------------------------------------

'change wave number scale
Private Sub hscrolGraphWaveNO_Change()
'draw the new wave no axes
pboxGraphWaveNOScale.Cls
pboxGraphWaveNOScaleForAmplitude.Cls
For I = 0 To 3
    'set the point for the text to be placed in the scale
    pboxGraphWaveNOScale.ForeColor = RGB(0, 0, 0)
    pboxGraphWaveNOScale.PSet (pboxGraphWaveNOScale.Width / 3 * I, 100)
    pboxGraphWaveNOScaleForAmplitude.ForeColor = RGB(0, 0, 0)
    pboxGraphWaveNOScaleForAmplitude.PSet (pboxGraphWaveNOScale.Width / 3 * I, pboxGraphWaveNOScale.Height - 200)
    'print the text in the picture box
    pboxGraphWaveNOScale.ForeColor = RGB(0, 255, 0)
    pboxGraphWaveNOScale.Print hscrolGraphWaveNO.Value + 100 * I
    pboxGraphWaveNOScaleForAmplitude.ForeColor = RGB(0, 255, 0)
    pboxGraphWaveNOScaleForAmplitude.Print hscrolGraphWaveNO.Value + 100 * I
    'draw lines to divide up the graph (makes graph easier
    'to read
    pboxGraphWaveNOScale.Line (pboxGraphWaveNOScale.Width / 3 * I, 200)-(pboxGraphWaveNOScale.Width / 3 * I, 0)
    pboxGraphWaveNOScale.Line (pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width / 6), 100)-(pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width / 6), 0)
    pboxGraphWaveNOScale.Line (pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width / 12), 60)-(pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width / 12), 0)
    pboxGraphWaveNOScale.Line (pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width * 18 / 72), 60)-(pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width * 18 / 72), 0)
    pboxGraphWaveNOScaleForAmplitude.Line (pboxGraphWaveNOScale.Width / 3 * I, pboxGraphWaveNOScale.Height - 200)-(pboxGraphWaveNOScale.Width / 3 * I, pboxGraphWaveNOScale.Height)
    pboxGraphWaveNOScaleForAmplitude.Line (pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width / 6), pboxGraphWaveNOScale.Height - 100)-(pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width / 6), pboxGraphWaveNOScale.Height)
    pboxGraphWaveNOScaleForAmplitude.Line (pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width / 12), pboxGraphWaveNOScale.Height - 60)-(pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width / 12), pboxGraphWaveNOScale.Height)
    pboxGraphWaveNOScaleForAmplitude.Line (pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width * 18 / 72), pboxGraphWaveNOScale.Height - 60)-(pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width * 18 / 72), pboxGraphWaveNOScale.Height)
Next I

'draw the horizontal line to separate the scale from the
'graph
pboxGraphWaveNOScale.Line (0, 15)-(pboxGraphWaveNOScale.Width, 15)
pboxGraphWaveNOScaleForAmplitude.Line (0, pboxGraphWaveNOScale.Height - 15)-(pboxGraphWaveNOScale.Width, pboxGraphWaveNOScale.Height - 15)
subDrawGraph
subDrawGraphAmplitude
'move the "graphical style" Hscroll bar
shapeHScrollMovingBar.Left = 1620 + (hscrolGraphWaveNO.Value / 9700) * 4020
labelHScrollMovingBar.Left = shapeHScrollMovingBar.Left
End Sub

Private Sub pboxLargeGraphClick()
'if the point is at zero then set it to -1 (for use in the
'loop, zero is disguarded)
If intCurrentYCordS% = 0 Then intCurrentYCordS% = -1
'set the values of the frequency at this point
arrPointsForFreq&(intCurrentXCordS%) = intCurrentYCordS%
'set the current Y coordinate back to zero (you will get an
'error if you don't)
If intCurrentYCordS% = -1 Then intCurrentYCordS% = 0
'draw the graph (don't draw it if the user is dragging the
'mouse, improves performance)
If pboxLargeGraph.Tag = "" Then subDrawGraph
End Sub

'draw the graph
Private Sub subDrawGraph()
On Error Resume Next
'set the values to non-zero (to state that is is the
'first time it is used
arrOldFreqVAlue&(0) = -1
arrOldFreqVAlue&(1) = -1
'set up the graph picture box's colour and refresh it
pboxLargeGraph.Cls
pboxLargeGraph.ForeColor = RGB(0, 255, 0)
'store the length of the divisions on the x-axis (for some
'reason they are not 15? but anyways this fixes that
'problem)
longintDivision# = pboxLargeGraph.Width / 300
'set temp integer to -5 (forces it to check for previous
'points)
intTempHolder% = -5
'This loop does most of the graphing
For I = 0 To 300
    'if value is zero, then skip drawing it (take this part
    'out and see what happens, makes it really hard to draw
    'a graph)
    If arrPointsForFreq&(I + hscrolGraphWaveNO.Value) = 0 Then GoTo 10
    'if the value is -1 then set it to zero, so that if the
    'user wants a zero value, then it will plot a zero
    'value
    If arrPointsForFreq&(I + hscrolGraphWaveNO.Value) = -1 Then arrPointsForFreq&(I + hscrolGraphWaveNO.Value) = 0
        'if temp integer is non-zero then do graphing as
        'per normal, otherwise find a previous value to
        'graph (otherwise the first line will be between
        'the first point and (0,0), try removing this if
        'statement and 'un-remark' the pbox...line
        'statement that is remarked below
        If intTempHolder% <> -5 Then
            'draw a line between the current and previous
            'point
            pboxLargeGraph.Line (arrOldFreqVAlue&(0) * longintDivision#, pboxLargeGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxLargeGraph.Height)-(I * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(I + hscrolGraphWaveNO) / intMaxHertzScale%) * pboxLargeGraph.Height)
        Else
            'find previous values (i.e. before
            'hscoll1.value)
            subFindPreviousValue
                If intTempHolder% <> -5 Then
                'if a point does exist before the current
                'point then draw a line between it and
                'the current point
                    pboxLargeGraph.Line ((intTempHolder% - hscrolGraphWaveNO.Value) * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(intTempHolder%) / intMaxHertzScale%) * pboxLargeGraph.Height)-(I * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(I + hscrolGraphWaveNO) / intMaxHertzScale%) * pboxLargeGraph.Height)
                Else
                    'set the fore colour to red
                    pboxLargeGraph.ForeColor = RGB(255, 0, 0)
                    'draw a line between the current point
                    'and (0,0)
                    pboxLargeGraph.Line (0, pboxLargeGraph.Height)-(I * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(I + hscrolGraphWaveNO) / intMaxHertzScale%) * pboxLargeGraph.Height)
                    'set the temp integer to non-zero (so
                    'that the sub doesn't keep checking
                    'for previous values, probably not
                    'necessary but will save a fraction
                    'of the processing time)
                    intTempHolder% = 1
                    'set the forecolour back to green
                    pboxLargeGraph.ForeColor = RGB(0, 255, 0)
                End If
        End If
'    pboxLargeGraph.Line (arrOldFreqVAlue&(0) * longintDivision#, pboxLargeGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxLargeGraph.Height)-(I * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(I + hscrolGraphWaveNO) / intMaxHertzScale%) * pboxLargeGraph.Height)
    'set the current values as the old values
    arrOldFreqVAlue&(0) = I
    arrOldFreqVAlue&(1) = arrPointsForFreq&(I + hscrolGraphWaveNO)
    'if the point is zero, set it back to the original -1
    'value
    If arrPointsForFreq&(I + hscrolGraphWaveNO) = 0 Then arrPointsForFreq&(I + hscrolGraphWaveNO) = -1
'the line number 10 (I hope everyone still understands line
'numbers, if not, it does the same as puting a 'goto start'
'then ':start' in a batch file (.bat) (or the other way
'round :-) ))
10
Next I
'find the next value that is not on the graph's current
'scale
If intTempHolder% = -5 Then
    subFindPreviousValue
    inttempHolder2% = intTempHolder%
    subFindNextValue
    If intTempHolder% <> -5 And inttempHolder2% <> -5 Then
        'if there is a point off the scale then draw a line to
        'it
        pboxLargeGraph.Line ((inttempHolder2% - hscrolGraphWaveNO.Value) * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(inttempHolder2%) / intMaxHertzScale%) * pboxLargeGraph.Height)-((intTempHolder% - hscrolGraphWaveNO.Value) * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(intTempHolder%) / intMaxHertzScale%) * pboxLargeGraph.Height)
    Else
        'if there isn't a line off the scale then draw a red
        'line to (pbox.width,0)
        pboxLargeGraph.ForeColor = RGB(255, 0, 0)
        pboxLargeGraph.Line (inttempHolder2% * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(inttempHolder2%) / intMaxHertzScale%) * pboxLargeGraph.Height)-(intTempHolder% * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(intTempHolder%) / intMaxHertzScale%) * pboxLargeGraph.Height)
    End If
Else
    subFindNextValue
    If intTempHolder% <> -5 Then
        'if there is a point off the scale then draw a line to
        'it
        pboxLargeGraph.Line (arrOldFreqVAlue&(0) * longintDivision#, pboxLargeGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxLargeGraph.Height)-(intTempHolder% * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(intTempHolder%) / intMaxHertzScale%) * pboxLargeGraph.Height)
    Else
        'if there isn't a line off the scale then draw a red
        'line to (pbox.width,0)
        pboxLargeGraph.ForeColor = RGB(255, 0, 0)
        pboxLargeGraph.Line (arrOldFreqVAlue&(0) * longintDivision#, pboxLargeGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxLargeGraph.Height)-(I * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(I + hscrolGraphWaveNO) / intMaxHertzScale%) * pboxLargeGraph.Height)
    End If
End If
'draw the thumb view of the WHOLE graph
subDrawThumbView
End Sub

Private Sub subFindPreviousValue()
'set the temp integer to zero
intTempHolder% = -5
'step backwards from the start of the current scale to zero
'(note:- I don't fully understand why you have to switch
'both the values around (the ones separated by the 'TO')
'and put step -1, no doubt its microsoft crazy logic)
For J = hscrolGraphWaveNO.Value To 0 Step -1
    'if a point has been found then store it in the temp
    'integer and exit the loop
    If arrPointsForFreq&(J) <> 0 Then
        intTempHolder% = J
        Exit For
    End If
Next J
End Sub

Private Sub subFindNextValue()
'set the temp integer to zero
intTempHolder% = -5
'step from the end of the scale to the very last value
'in the array
For J = 300 + hscrolGraphWaveNO.Value To 10000
    If arrPointsForFreq&(J) <> 0 Then
    'if a point has been found then store it in the temp
    'integer and exit the loop
        intTempHolder% = J
        Exit For
    End If
Next J
End Sub

Private Sub subDrawThumbView()
'set the values to non-zero (to state that is is the
'first time it is used
arrOldFreqVAlue&(0) = -1
arrOldFreqVAlue&(1) = -1
'set up the pbox
pboxMiniGraph.Cls
pboxMiniGraph.ForeColor = RGB(0, 155, 0)
'store the division lengths in vbpixels
longintDivision# = pboxMiniGraph.Width / 10000
'this loop draw the thumb view graph
For I = 0 To 10000
    'if the value of the current point is zero then skip it
    If arrPointsForFreq&(I) = 0 Then GoTo 10
    'if the current value is -1 then set it to zero
    If arrPointsForFreq&(I) = -1 Then arrPointsForFreq&(I) = 0
    If arrOldFreqVAlue&(0) <> -1 Then
        'if it is not the first point then draw a line
        'between the current point and the previous point
        pboxMiniGraph.Line (I * longintDivision#, pboxMiniGraph.Height - (arrPointsForFreq&(I) / intMaxHertzScale%) * pboxMiniGraph.Height)-(arrOldFreqVAlue&(0) * longintDivision#, pboxMiniGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxMiniGraph.Height)
    Else
        'if it is the first point then draw a red line
        'betweent the current point and (0,0)
        pboxMiniGraph.ForeColor = RGB(155, 0, 0)
        pboxMiniGraph.Line (I * longintDivision#, pboxMiniGraph.Height - (arrPointsForFreq&(I) / intMaxHertzScale%) * pboxMiniGraph.Height)-(arrOldFreqVAlue&(0) * longintDivision#, pboxMiniGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxMiniGraph.Height)
        pboxMiniGraph.ForeColor = RGB(0, 155, 0)
    End If
    'set the current point to the old points
    arrOldFreqVAlue&(0) = I
    arrOldFreqVAlue&(1) = arrPointsForFreq&(I)
    'if the current point is zero then set it back to -1
    If arrPointsForFreq&(I) = 0 Then arrPointsForFreq&(I) = -1
10
Next I
'draw a red line between the current point and
'(pbox.width,0)
pboxMiniGraph.ForeColor = RGB(155, 0, 0)
pboxMiniGraph.Line (arrOldFreqVAlue&(0) * longintDivision#, pboxMiniGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxMiniGraph.Height)-(10000 * longintDivision#, pboxMiniGraph.Height)
End Sub


'stop redrawing the graph (for performance)
Private Sub pboxLargeGraph_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pboxLargeGraph.Tag = "Stop Re-Drawing Graph"
'call the mouse move sub (other wise the point won't be
'added to the array)
Call pboxLargeGraph_MouseMove(Button, Shift, X, Y)
End Sub

'When mouse moves in the graph picture box
Private Sub pboxLargeGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'----------------------------------------------
'Hide the four line in the other graph
'----------------------------------------------
lineH1CrossFollowingMouse.Visible = True
lineH2CrossFollowingMouse.Visible = True
lineV1CrossFollowingMouse.Visible = True
lineV2CrossFollowingMouse.Visible = True
lineH1CrossFollowingMouseInAmplitude.Visible = False
lineH2CrossFollowingMouseInAmplitude.Visible = False
lineV1CrossFollowingMouseInAmplitude.Visible = False
lineV2CrossFollowingMouseInAmplitude.Visible = False
'----------------------------------------------
'----------------------------------------------
'make the four lines move with the mouse
'----------------------------------------------
lineH1CrossFollowingMouse.X1 = -160
lineH1CrossFollowingMouse.X2 = X - 120
lineH1CrossFollowingMouse.Y1 = Y
lineH1CrossFollowingMouse.Y2 = Y
lineV1CrossFollowingMouse.X1 = X
lineV1CrossFollowingMouse.X2 = X
lineV1CrossFollowingMouse.Y1 = -160
lineV1CrossFollowingMouse.Y2 = Y - 120
lineH2CrossFollowingMouse.X1 = X + 120
lineH2CrossFollowingMouse.X2 = pboxLargeGraph.Width
lineH2CrossFollowingMouse.Y1 = Y
lineH2CrossFollowingMouse.Y2 = Y
lineV2CrossFollowingMouse.X1 = X
lineV2CrossFollowingMouse.X2 = X
lineV2CrossFollowingMouse.Y1 = Y + 120
lineV2CrossFollowingMouse.Y2 = pboxLargeGraph.Height
'----------------------------------------------
'display the coordinates of the mouse cursor
intCurrentXCordS% = hscrolGraphWaveNO.Value + Int((X / (pboxLargeGraph.Width - 15)) * (300))
intCurrentYCordS% = Int(intMaxHertzScale% - (intMaxHertzScale% * (Y / (pboxLargeGraph.Height - 15))))
'a bit of error trapping
If intCurrentXCordS% < 0 Then intCurrentXCordS% = 0
If intCurrentYCordS% < 0 Then intCurrentYCordS% = 0
If intCurrentXCordS% > 10000 Then intCurrentXCordS% = 10000
If intCurrentYCordS% > intMaxHertzScale% Then intCurrentYCordS% = intMaxHertzScale%
'display value of point at this spot
labelCurrentCords.Caption = "( Wave no.=" + Mid$(Str$(intCurrentXCordS%), 2, 100) + " ," + " Hertz=" + Mid$(Str$(intCurrentYCordS%), 2, 100) + " )"
If arrPointsForFreq&(intCurrentXCordS%) = -1 Then labelCurrentValue.Caption = "Hertz Value = Silence" Else labelCurrentValue.Caption = "Hertz Value =" + Str$(arrPointsForFreq&(intCurrentXCordS%))
'if the user has the mouse button down then set the point in
'the array (draw a point instead of the whole graph to improve
'overall performance)
If Button = 1 Then
    pboxLargeGraph.ForeColor = RGB(0, 255, 0)
    pboxLargeGraph.PSet (X, Y)
    pboxLargeGraphClick
End If
'if the user right clicks then delete the current point
If Button = 2 Then
    'set the value of the frequency at this point to null
    arrPointsForFreq&(intCurrentXCordS%) = 0
    pboxLargeGraph.ForeColor = RGB(255, 0, 0)
    pboxLargeGraph.PSet (X, Y)
End If
'I originally just set the tag to a null value every time the
'user lifts their mouse button, but a problem occured when the
'mouse left the edge of the pbox without lifting the button,
'so I came up with this solution
If pboxLargeGraph.Tag = "Draw Graph Now" Then
    pboxLargeGraph.Tag = ""
    subDrawGraph
End If
End Sub

'Continue re-drawing the graph
Private Sub pboxLargeGraph_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pboxLargeGraph.Tag = "Draw Graph Now"
End Sub

'stop redrawing the graph (for performance)
Private Sub pboxLargeGraphAmplitude_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pboxLargeGraphAmplitude.Tag = "Stop Re-Drawing Graph"
'call the mouse move sub (other wise the point won't be
'added to the array)
Call pboxLargeGraphAmplitude_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub pboxLargeGraphAmplitude_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'----------------------------------------------
'Hide the four line in the other graph
'----------------------------------------------
lineH1CrossFollowingMouse.Visible = False
lineH2CrossFollowingMouse.Visible = False
lineV1CrossFollowingMouse.Visible = False
lineV2CrossFollowingMouse.Visible = False
lineH1CrossFollowingMouseInAmplitude.Visible = True
lineH2CrossFollowingMouseInAmplitude.Visible = True
lineV1CrossFollowingMouseInAmplitude.Visible = True
lineV2CrossFollowingMouseInAmplitude.Visible = True
'----------------------------------------------
'----------------------------------------------
'make the four lines move with the mouse in the amplitude
'graph
'----------------------------------------------
lineH1CrossFollowingMouseInAmplitude.X1 = -160
lineH1CrossFollowingMouseInAmplitude.X2 = X - 120
lineH1CrossFollowingMouseInAmplitude.Y1 = Y
lineH1CrossFollowingMouseInAmplitude.Y2 = Y
lineV1CrossFollowingMouseInAmplitude.X1 = X
lineV1CrossFollowingMouseInAmplitude.X2 = X
lineV1CrossFollowingMouseInAmplitude.Y1 = -160
lineV1CrossFollowingMouseInAmplitude.Y2 = Y - 120
lineH2CrossFollowingMouseInAmplitude.X1 = X + 120
lineH2CrossFollowingMouseInAmplitude.X2 = pboxLargeGraphAmplitude.Width
lineH2CrossFollowingMouseInAmplitude.Y1 = Y
lineH2CrossFollowingMouseInAmplitude.Y2 = Y
lineV2CrossFollowingMouseInAmplitude.X1 = X
lineV2CrossFollowingMouseInAmplitude.X2 = X
lineV2CrossFollowingMouseInAmplitude.Y1 = Y + 120
lineV2CrossFollowingMouseInAmplitude.Y2 = pboxLargeGraphAmplitude.Height
'----------------------------------------------

'display the coordinates of the mouse cursor
intCurrentXCordSAmplitude% = hscrolGraphWaveNO.Value + Int((X / (pboxLargeGraphAmplitude.Width - 15)) * (300))
intCurrentYCordSAmplitude% = Int((100 * (Y / (pboxLargeGraphAmplitude.Height - 15))))
'a bit of error trapping
If intCurrentXCordSAmplitude% < 0 Then intCurrentXCordSAmplitude% = 0
If intCurrentYCordSAmplitude% < 0 Then intCurrentYCordSAmplitude% = 0
If intCurrentXCordSAmplitude% > 10000 Then intCurrentXCordSAmplitude% = 10000
If intCurrentYCordSAmplitude% > 100 Then intCurrentYCordSAmplitude% = 100
'display value of point at this spot
labelCurrentCords.Caption = "( Wave no.=" + Mid$(Str$(intCurrentXCordSAmplitude%), 2, 100) + " ," + " Amplitude=" + Mid$(Str$(intCurrentYCordSAmplitude%), 2, 100) + "% )"
If arrPointsForAmplitude&(intCurrentXCordSAmplitude%) = -1 Then labelCurrentValue.Caption = "Amplitude Value = 100%" Else labelCurrentValue.Caption = "Amplitude Value =" + Str$(arrPointsForAmplitude&(intCurrentXCordSAmplitude%)) + "%"

'if the user has the mouse button down then set the point in
'the array (draw a point instead of the whole graph to improve
'overall performance)
If Button = 1 Then
    If intCurrentYCordSAmplitude% = 100 Then intCurrentYCordSAmplitude% = -1
    pboxLargeGraphAmplitude.ForeColor = RGB(0, 255, 0)
    pboxLargeGraphAmplitude.PSet (X, Y)
    'if the point is at zero then set it to -1 (for use in the
    'loop, zero is disguarded)
    If intCurrentYCordSAmplitude% = 0 Then intCurrentYCordSAmplitude% = -1
    'set the values of the amplitude at this point
    arrPointsForAmplitude&(intCurrentXCordSAmplitude%) = intCurrentYCordSAmplitude%
    'set the current Y coordinate back to zero (you will get
    'an error if you don't)
    If intCurrentYCordSAmplitude% = -1 Then intCurrentYCordSAmplitude% = 0
    'draw the graph (don't draw it if the user is dragging the
    'mouse, improves performance)
'    If pboxLargeGraphAmplitude.Tag = "" Then subDrawGraph
    If intCurrentYCordSAmplitude% = -1 Then intCurrentYCordSAmplitude% = 100
End If

'if the user right clicks then delete the current point
If Button = 2 Then
    'set the value of the frequency at this point to null
    arrPointsForAmplitude&(intCurrentXCordSAmplitude%) = 100
    pboxLargeGraphAmplitude.ForeColor = RGB(255, 0, 0)
    pboxLargeGraphAmplitude.PSet (X, Y)
End If
'I originally just set the tag to a null value every time the
'user lifts their mouse button, but a problem occured when the
'mouse left the edge of the pbox without lifting the button,
'so I came up with this solution
If pboxLargeGraphAmplitude.Tag = "Draw Graph Now" Then
    pboxLargeGraphAmplitude.Tag = ""
    subDrawGraphAmplitude
End If
If intCurrentYCordSAmplitude% = 100 Then intCurrentYCordSAmplitude% = -1
End Sub

'allow drawing of the amplitude curve
Private Sub pboxLargeGraphAmplitude_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pboxLargeGraphAmplitude.Tag = "Draw Graph Now"
End Sub

'display the current frequency by means of a moving speaker
Private Sub timerWavePreview_Timer()
pboxWavePreview.Cls
'number of frames (the higher the better looking)
intNoOfFrames% = 10

'if the hertz value is 0 then draw a non-moving speaker
If intCurrentYCordS% = 0 Then
    pboxWavePreview.Line (0, 50 + 100)-(600, 350 + 100)
    pboxWavePreview.Line (1000, 350 + 100)-(1600, 50 + 100)
    pboxWavePreview.Line (1000, 350 + 100)-(585, 350 + 100)
    pboxWavePreview.Line (900, 350 + 100)-(900, 550 + 100)
    pboxWavePreview.Line (685, 350 + 100)-(685, 550 + 100)
    Exit Sub
End If

'set the time intervals according to the hertz value
If intCurrentYCordS% > 1000 Then
    timerWavePreview.Interval = 1
Else
    timerWavePreview.Interval = (1000 / (intCurrentYCordS% * intNoOfFrames%)) + 1
End If
'I've never actually used the Tags to store anything in the
'past, but I've found out that they do come in handy!...
'increment tag by 1 (used for calculating its vertical
'position)
timerWavePreview.Tag = Val(timerWavePreview.Tag) + 1
'if the value is getting a bit high, then set it back to 0
If Val(timerWavePreview.Tag) > 1000 Then timerWavePreview.Tag = 0

'Draw the cone and the coil of the speaker
pboxWavePreview.Line (0, 100 + 100 * Sin(2 * 3.1415926654 / intNoOfFrames% * timerWavePreview.Tag))-(600, 400 + 100 * Sin(2 * 3.1415926654 / intNoOfFrames% * timerWavePreview.Tag))
pboxWavePreview.Line (1000, 400 + 100 * Sin(2 * 3.1415926654 / intNoOfFrames% * timerWavePreview.Tag))-(1600, 100 + 100 * Sin(2 * 3.1415926654 / intNoOfFrames% * timerWavePreview.Tag))
pboxWavePreview.Line (1000, 400 + 100 * Sin(2 * 3.1415926654 / intNoOfFrames% * timerWavePreview.Tag))-(585, 400 + 100 * Sin(2 * 3.1415926654 / intNoOfFrames% * timerWavePreview.Tag))
pboxWavePreview.Line (900, 400 + 100 * Sin(2 * 3.1415926654 / intNoOfFrames% * timerWavePreview.Tag))-(900, 600 + 100 * Sin(2 * 3.1415926654 / intNoOfFrames% * timerWavePreview.Tag))
pboxWavePreview.Line (685, 400 + 100 * Sin(2 * 3.1415926654 / intNoOfFrames% * timerWavePreview.Tag))-(685, 600 + 100 * Sin(2 * 3.1415926654 / intNoOfFrames% * timerWavePreview.Tag))
End Sub

'----------------------------------------------------------
'----------------------------------------------------------
'----------------------------------------------------------
'These last few subroutines are the same as the ones for the
'frequency, but draw the graph of the amplitude instead
'----------------------------------------------------------
'draw the Amplitude graph
Private Sub subDrawGraphAmplitude()
On Error Resume Next
'set the apparent height of pboxLargeGraphAmplitude
longintApparentHeight& = pboxLargeGraphAmplitude.Height - 15
'set the values to non-zero (to state that is is the
'first time it is used
arrOldAmplitudeVAlue&(0) = -1
arrOldAmplitudeVAlue&(1) = -1
'set up the graph picture box's colour and refresh it
pboxLargeGraphAmplitude.Cls
pboxLargeGraphAmplitude.ForeColor = RGB(0, 255, 0)
'store the length of the divisions on the x-axis (for some
'reason they are not 15? but anyways this fixes that
'problem)
longintDivision# = pboxLargeGraphAmplitude.Width / 300
'set temp integer to -5 (forces it to check for previous
'points)
intTempHolder% = -5
'This loop does most of the graphing
For I = 0 To 300
    'if value is zero, then skip drawing it (take this part
    'out and see what happens, makes it really hard to draw
    'a graph)
    If arrPointsForAmplitude&(I + hscrolGraphWaveNO.Value) = 100 Then GoTo 10
    'if the value is -1 then set it to zero, so that if the
    'user wants a zero value, then it will plot a zero
    'value
    If arrPointsForAmplitude&(I + hscrolGraphWaveNO.Value) = -1 Then arrPointsForAmplitude&(I + hscrolGraphWaveNO.Value) = 100
        'if temp integer is non-zero then do graphing as
        'per normal, otherwise find a previous value to
        'graph (otherwise the first line will be between
        'the first point and (0,0), try removing this if
        'statement and 'un-remark' the pbox...line
        'statement that is remarked below
        If intTempHolder% <> -5 Then
            'draw a line between the current and previous
            'point
            pboxLargeGraphAmplitude.Line (arrOldAmplitudeVAlue&(0) * longintDivision#, (arrOldAmplitudeVAlue&(1) / 100) * longintApparentHeight&)-(I * longintDivision#, (arrPointsForAmplitude&(I + hscrolGraphWaveNO) / 100) * longintApparentHeight&)
        Else
            'find previous values (i.e. before
            'hscoll1.value)
            subFindPreviousValueAmplitude
                If intTempHolder% <> -5 Then
                'if a point does exist before the current
                'point then draw a line between it and
                'the current point
                    pboxLargeGraphAmplitude.Line ((intTempHolder% - hscrolGraphWaveNO.Value) * longintDivision#, (arrPointsForAmplitude&(intTempHolder%) / 100) * longintApparentHeight&)-(I * longintDivision#, (arrPointsForAmplitude&(I + hscrolGraphWaveNO) / 100) * longintApparentHeight&)
                Else
                    'set the fore colour to red
                    pboxLargeGraphAmplitude.ForeColor = RGB(255, 0, 0)
                    'draw a line between the current point
                    'and (0,0)
                    pboxLargeGraphAmplitude.Line (0, longintApparentHeight&)-(I * longintDivision#, (arrPointsForAmplitude&(I + hscrolGraphWaveNO) / 100) * longintApparentHeight&)
                    'set the temp integer to non-zero (so
                    'that the sub doesn't keep checking
                    'for previous values, probably not
                    'necessary but will save a fraction
                    'of the processing time)
                    intTempHolder% = 1
                    'set the forecolour back to green
                    pboxLargeGraphAmplitude.ForeColor = RGB(0, 255, 0)
                End If
        End If
'    pboxlargegraphamplitude.Line (arrOldAmplitudeVAlue&(0) * longintDivision#, longintApparentHeight& - (arrOldAmplitudeVAlue&(1) / 100) * longintApparentHeight&)-(I * longintDivision#, longintApparentHeight& - (arrPointsForAmplitude&(I + hscrolGraphWaveNO) / 100) * longintApparentHeight&)
    'set the current values as the old values
    arrOldAmplitudeVAlue&(0) = I
    arrOldAmplitudeVAlue&(1) = arrPointsForAmplitude&(I + hscrolGraphWaveNO)
    'if the point is zero, set it back to the original -1
    'value
    If arrPointsForAmplitude&(I + hscrolGraphWaveNO) = 100 Then arrPointsForAmplitude&(I + hscrolGraphWaveNO) = -1
'the line number 10 (I hope everyone still understands line
'numbers, if not, it does the same as puting a 'goto start'
'then ':start' in a batch file (.bat) (or the other way
'round :-) ))
10
Next I
'find the next value that is not on the graph's current
'scale
If intTempHolder% = -5 Then
    subFindPreviousValueAmplitude
    inttempHolder2% = intTempHolder%
    subFindNextValueAmplitude
    If intTempHolder% <> -5 And inttempHolder2% <> -5 Then
        'if there is a point off the scale then draw a line to
        'it
        pboxLargeGraphAmplitude.Line ((inttempHolder2% - hscrolGraphWaveNO.Value) * longintDivision#, (arrPointsForAmplitude&(inttempHolder2%) / 100) * longintApparentHeight&)-((intTempHolder% - hscrolGraphWaveNO.Value) * longintDivision#, (arrPointsForAmplitude&(intTempHolder%) / 100) * longintApparentHeight&)
    Else
        'if there isn't a line off the scale then draw a red
        'line to (pbox.width,0)
        pboxLargeGraphAmplitude.ForeColor = RGB(255, 0, 0)
        pboxLargeGraphAmplitude.Line (inttempHolder2% * longintDivision#, (arrPointsForAmplitude&(inttempHolder2%) / 100) * longintApparentHeight&)-(intTempHolder% * longintDivision#, (arrPointsForAmplitude&(intTempHolder%) / 100) * longintApparentHeight&)
    End If
Else
    subFindNextValueAmplitude
    If intTempHolder% <> -5 Then
        'if there is a point off the scale then draw a line to
        'it
        pboxLargeGraphAmplitude.Line (arrOldAmplitudeVAlue&(0) * longintDivision#, (arrOldAmplitudeVAlue&(1) / 100) * longintApparentHeight&)-(intTempHolder% * longintDivision#, (arrPointsForAmplitude&(intTempHolder%) / 100) * longintApparentHeight&)
    Else
        'if there isn't a line off the scale then draw a red
        'line to (pbox.width,0)
        pboxLargeGraphAmplitude.ForeColor = RGB(255, 0, 0)
        pboxLargeGraphAmplitude.Line (arrOldAmplitudeVAlue&(0) * longintDivision#, (arrOldAmplitudeVAlue&(1) / 100) * longintApparentHeight&)-(I * longintDivision#, (arrPointsForAmplitude&(I + hscrolGraphWaveNO) / 100) * longintApparentHeight&)
    End If
End If
'draw the thumb view of the WHOLE graph
subDrawThumbViewAmplitude
End Sub
Private Sub subFindNextValueAmplitude()
'set the temp integer to zero
intTempHolder% = -5
'step from the end of the scale to the very last value
'in the array
For J = 300 + hscrolGraphWaveNO.Value To 10000
    If arrPointsForAmplitude&(J) <> 100 Then
    'if a point has been found then store it in the temp
    'integer and exit the loop
        intTempHolder% = J
        Exit For
    End If
Next J
End Sub
Private Sub subFindPreviousValueAmplitude()
'set the temp integer to zero
intTempHolder% = -5
'step backwards from the start of the current scale to zero
For J = hscrolGraphWaveNO.Value To 0 Step -1
    'if a point has been found then store it in the temp
    'integer and exit the loop
    If arrPointsForAmplitude&(J) <> 100 Then
        intTempHolder% = J
        Exit For
    End If
Next J
End Sub
Private Sub subDrawThumbViewAmplitude()
longintApparentHeight& = pboxMiniGraphAmplitude.Height - 15
'set the values to non-zero (to state that is is the
'first time it is used
arrOldAmplitudeVAlue&(0) = -1
arrOldAmplitudeVAlue&(1) = -1
'set up the pbox
pboxMiniGraphAmplitude.Cls
pboxMiniGraphAmplitude.ForeColor = RGB(0, 155, 0)
'store the division lengths in vbpixels
longintDivision# = pboxMiniGraphAmplitude.Width / 10000
'this loop draw the thumb view graph
For I = 0 To 10000
    'if the value of the current point is zero then skip it
    If arrPointsForAmplitude&(I) = 100 Then GoTo 10
    'if the current value is -1 then set it to zero
    If arrPointsForAmplitude&(I) = -1 Then arrPointsForAmplitude&(I) = 100
    If arrOldAmplitudeVAlue&(0) <> -1 Then
        'if it is not the first point then draw a line
        'between the current point and the previous point
        pboxMiniGraphAmplitude.Line (I * longintDivision#, (arrPointsForAmplitude&(I) / 100) * longintApparentHeight&)-(arrOldAmplitudeVAlue&(0) * longintDivision#, (arrOldAmplitudeVAlue&(1) / 100) * longintApparentHeight&)
    Else
        'if it is the first point then draw a red line
        'between the current point and (0,0)
        pboxMiniGraphAmplitude.ForeColor = RGB(155, 0, 0)
        pboxMiniGraphAmplitude.Line (I * longintDivision#, (arrPointsForAmplitude&(I) / 100) * longintApparentHeight&)-((arrOldAmplitudeVAlue&(0) / 100) * arrOldAmplitudeVAlue&(0) * longintDivision#, longintApparentHeight&)
        pboxMiniGraphAmplitude.ForeColor = RGB(0, 155, 0)
    End If
    'set the current point to the old points
    arrOldAmplitudeVAlue&(0) = I
    arrOldAmplitudeVAlue&(1) = arrPointsForAmplitude&(I)
    'if the current point is zero then set it back to -1
    If arrPointsForAmplitude&(I) = 100 Then arrPointsForAmplitude&(I) = -1
10
Next I
'draw a red line between the current point and
'(pbox.width,0)
pboxMiniGraphAmplitude.ForeColor = RGB(155, 0, 0)
If arrOldAmplitudeVAlue&(1) = -1 Then arrOldAmplitudeVAlue&(1) = 100
pboxMiniGraphAmplitude.Line (arrOldAmplitudeVAlue&(0) * longintDivision#, (arrOldAmplitudeVAlue&(1) / 100) * longintApparentHeight&)-(10000 * longintDivision#, longintApparentHeight&)
End Sub
'----------------------------------------------------------

