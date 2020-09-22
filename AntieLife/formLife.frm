VERSION 5.00
Begin VB.Form formLife 
   BackColor       =   &H00000000&
   Caption         =   "Evolving Artificial Life"
   ClientHeight    =   6150
   ClientLeft      =   1170
   ClientTop       =   915
   ClientWidth     =   7680
   FillColor       =   &H008080FF&
   ForeColor       =   &H000000FF&
   Icon            =   "formLife.frx":0000
   LinkTopic       =   "formLife"
   ScaleHeight     =   6150
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Remove Animals"
      Height          =   492
      Left            =   6240
      TabIndex        =   9
      Top             =   5520
      Width           =   1092
   End
   Begin VB.CommandButton cmdDelPlants 
      Caption         =   "Remove Plants"
      Height          =   492
      Left            =   360
      TabIndex        =   5
      Top             =   5520
      Width           =   1092
   End
   Begin VB.Timer tmrMultiTask 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   5040
   End
   Begin VB.PictureBox picLife 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4812
      Left            =   120
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   4
      Top             =   120
      Width           =   7452
   End
   Begin VB.CommandButton cmdAddAnts 
      Caption         =   "Add Animals"
      Height          =   372
      Left            =   4680
      TabIndex        =   8
      Top             =   5520
      Width           =   1332
   End
   Begin VB.CommandButton cmdAddPlants 
      Caption         =   "Add Plants"
      Height          =   372
      Left            =   1680
      TabIndex        =   6
      Top             =   5520
      Width           =   1332
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Restart"
      Height          =   492
      Left            =   3360
      TabIndex        =   7
      Top             =   5520
      Width           =   972
   End
   Begin VB.Label lblAntCountLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Number of animals:"
      ForeColor       =   &H00C0C0FF&
      Height          =   252
      Left            =   4200
      TabIndex        =   3
      Top             =   5160
      Width           =   2052
   End
   Begin VB.Label lblPlantCountLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Number of plants:"
      ForeColor       =   &H00C0C0FF&
      Height          =   252
      Left            =   360
      TabIndex        =   2
      Top             =   5160
      Width           =   2052
   End
   Begin VB.Label lblAntCount 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00C0C0FF&
      Height          =   252
      Left            =   6360
      TabIndex        =   1
      Top             =   5160
      Width           =   1212
   End
   Begin VB.Label lblPlantCount 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00C0C0FF&
      Height          =   252
      Left            =   2520
      TabIndex        =   0
      Top             =   5160
      Width           =   1212
   End
   Begin VB.Menu mOptions 
      Caption         =   "&Options"
      Begin VB.Menu mChkEnableSounds 
         Caption         =   "Enable sound effects"
      End
      Begin VB.Menu mChkInvertRed 
         Caption         =   "Invert red level"
      End
      Begin VB.Menu mChkInvertGreen 
         Caption         =   "Invert green level"
      End
      Begin VB.Menu mChkInvertBlue 
         Caption         =   "Invert blue level"
      End
      Begin VB.Menu mChkNoFatalAge 
         Caption         =   "Prohibit death by old age"
      End
      Begin VB.Menu mChkNoFatalHunger 
         Caption         =   "Prohibit death by starvation"
      End
      Begin VB.Menu mChkNoCloneAnt 
         Caption         =   "Prohibit self-reproducing animals"
      End
      Begin VB.Menu mChkNoClonePlant 
         Caption         =   "Prohibit self-reproducing plants"
      End
      Begin VB.Menu mChkNoMateAntAnt 
         Caption         =   "Prohibit animals from mating with animals"
      End
      Begin VB.Menu mChkNoMateAntPlant 
         Caption         =   "Prohibit animals from mating with plants"
      End
      Begin VB.Menu mChkNoMatePlantAnt 
         Caption         =   "Prohibit plants from mating with animals"
      End
      Begin VB.Menu mChkNoMatePlantPlant 
         Caption         =   "Prohibit plants from mating with plants"
      End
      Begin VB.Menu mChkNoAntEatAnt 
         Caption         =   "Prohibit animals from eating animals"
      End
      Begin VB.Menu mChkNoAntEatPlant 
         Caption         =   "Prohibit animals from eating plants"
      End
      Begin VB.Menu mChkNoPlantEatAnt 
         Caption         =   "Prohibit plants from eating animals"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkNoPlantEatPlant 
         Caption         =   "Prohibit plants from eating plants"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkNoAntEatCarion 
         Caption         =   "Prohibit animals from scavenging"
      End
      Begin VB.Menu mChkNoPlantEatCarion 
         Caption         =   "Prohibit plants from scavenging"
      End
      Begin VB.Menu mChkNoAntEatDirt 
         Caption         =   "Prohibit ambient feeding for animals"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkAllowSpores 
         Caption         =   "Allow animals to mate over a distance"
      End
      Begin VB.Menu mChkAllowPollen 
         Caption         =   "Allow plants to mate over a distance"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkAllowPlagues 
         Caption         =   "Allow plagues"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkQuickStart 
         Caption         =   "Quick-start evolution"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkFastEvolve 
         Caption         =   "High rate of mutation"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkBreakRules 
         Caption         =   "Occasionally break selected rules"
         Checked         =   -1  'True
      End
      Begin VB.Menu mChkZoomFactor 
         Caption         =   "&Zoom Factor"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "formLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim initialHeight As Single
Dim initialWidth As Single

Private ExitForm As Boolean

Private Sub cmdAddAnts_Click()
    'Add a user specified number of primateve live animals.
    Dim s As String
    Dim n
    s = InputBox("Add how many random primative animals?" + vbCrLf + vbCrLf + "0 = cancel", "Add Animals", "1")
    If s = "" Then Exit Sub
    'On Error Resume Next
    n = -Val(s)
    'On Error GoTo 0
    If n = 0 Then Exit Sub
    firstLife n
End Sub

Private Sub cmdAddPlants_Click()
    'Add a user specified number of primateve live plants.
    Dim s As String
    Dim n
    s = InputBox("Add how many primative plants?" + vbCrLf + vbCrLf + "0 = cancel", "Add Plants", "1")
    If s = "" Then Exit Sub
    'On Error Resume Next
    n = Val(s)
    'On Error GoTo 0
    If n = 0 Then Exit Sub
    firstLife n
End Sub

Private Sub cmdDelPlants_Click()
    'Delete a user specified number of plants.
    Dim s As String
    Dim n
    s = InputBox("Remove how many plants?" + vbCrLf + vbCrLf + "0 = cancel", "Delete Plants", "1")
    If s = "" Then Exit Sub
    'On Error Resume Next
    n = Val(s)
    'On Error GoTo 0
    If n = 0 Then Exit Sub
    deleteLife n
End Sub

Private Sub cmdRestart_Click()
    'Start over with a user specified number of primateve live plants.  (choosing a negative number will start with animals.)
    Dim s As String
    Dim n
    picLife.BorderStyle = 1 'show the pictureBox border.
    s = InputBox("Restart with how many primative plants?" + vbCrLf + vbCrLf + "0 = random", "Restart", "1")
    If s = "" Then
        picLife.BorderStyle = 0 'hide the pictureBox border.
        Exit Sub
    End If
    'On Error Resume Next
    n = Val(s)
    'On Error GoTo 0
    If n = 0 Then n = Int(Rnd * (Rnd + 1) * 500 + 1) 'pick a random number of plants and finish re-initializing.
    postInitialize n 're-initialize life form environment.
End Sub

Private Sub Command1_Click()
    'Delete a user specified number of Animals.
    Dim s As String
    Dim n
    s = InputBox("Remove how many animals?" + vbCrLf + vbCrLf + "0 = cancel", "Delete Animals", "1")
    If s = "" Then Exit Sub
    'On Error Resume Next
    n = Val(s)
    'On Error GoTo 0
    If n = 0 Then Exit Sub
    deleteLife -n
End Sub

Private Sub Form_Activate()
    Static postinitialized As Boolean 'initially false
    
    'multitasking core.
    Static timeOfLastRedraw As Single
    Static taskNumber As Integer
    Static X As Single
    Static Y As Single
    
    Dim XX As Single
    Dim YY As Single
    Dim i As Integer 'loop counter
    Dim lngLastTick As Long
    
    
    If Not postinitialized Then
        postInitialize Int(Rnd * Rnd * 500 + 1) 'pick a random number of plants and finish initializing.
        postinitialized = True
        
        
        Do
            
            
            For i = 0 To 25 'arbitrary number of loops.  Low numbers are slower. High numbers are less responsive.
                taskNumber = taskNumber + 1
                If taskNumber > 8 Then taskNumber = 0
                If taskNumber <> 0 Then
                    If Not ((X >= LowerX) And (X <= UpperX) And (Y >= LowerY) And (Y <= UpperY)) Then
                        taskNumber = -1  'Do not process life forms out of the zoom area.
                    End If
                End If
                Select Case taskNumber
                Case 0
                    GetNextItemXY X, Y
                Case 1
                    moveItemXY X, Y
                Case 2
                    mateItemXY X, Y
                Case 3
                    cloneItemXY X, Y
                Case 4
                    feedItemXY X, Y
                Case 5
                    ageItemXY X, Y
                Case 6
                    healItemXY X, Y
                Case 7
                    processNextGeneXY X, Y
                Case Else
                    
                    'If (Abs(timeOfLastRedraw - Timer) > 0.5) Then
                    '**using Long as a data type is a lot faster than using Single
                    If ((GetTickCount - lngLastTick) >= 500) Then
                        
                        '** all graphical drawing needs to be inside the IF statement, including the
                        'DoEvents as this is one of the things that really kills an app when in a
                        'loop
                        DoEvents
                        lngLastTick = GetTickCount
                    End If
                End Select
            Next i
            If lifeCount < 1 Then 'Odds are "supposed to be" against all of this...
                'Simulate long term interactions of biochemicals in "lifeless" organic materials.
                If Grid(X, Y).Energy < 0 Then
                    If Not Grid(X, Y).Alive Then
                        XX = GetRndInt(RangeX - 1) + LowerX
                        YY = GetRndInt(RangeY - 1) + LowerY
                        If Sqr((XX - X) ^ 2 + (YY - Y) ^ 2) < Abs(Grid(X, Y).Energy) Then
                            If (Grid(X, Y).RGB = Grid(XX, YY).RGB) And Not Grid(XX, YY).Alive Then
                                Grid(X, Y).Energy = -Grid(X, Y).Energy
                                BirthItemXY XX, YY
                            End If
                        End If
                    End If
                End If
            End If
            If (mChkAllowPlagues.Checked Xor breakRules) Then
                If lifeCount > 200 Then 'Don't have plagues when the population is low.
                    If neighborCount(X, Y) > 4 - Sgn(Grid(X, Y).Speed) Then '(over-crowded)
                        plagueItemXY X, Y
                    End If
                End If
            End If
        Loop Until ExitForm
        Unload Me
    End If
End Sub

Sub postInitialize(n) 'Initialization of environment after the form has been initialized.
    resetGrid 'Prepare the grid for life to grow in it.
    picLife.Cls 'Clear the pictureBox.
    picLife.BorderStyle = 0 'hide the pictureBox border.
    firstLife n 'Seed first n primative life forms.
    tmrMultiTask.Enabled = True 'Activate the multi-tasking core.
End Sub

Private Sub Form_Initialize()
    ZoomFactor = 1.5
    mChkZoomFactor.Caption = "Zoom Factor = " + Trim(Str(ZoomFactor))
    OffsetX = 0
    OffsetY = 0
End Sub

Private Sub Form_Load()
    Dim ctrl As Control
    Dim sTmp As String
    
    formLife.Caption = "AntieLife - v." + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + ".0." + Trim(Str(App.Revision)) + " - TechnoZeus"
    initialHeight = formLife.ScaleHeight
    initialWidth = formLife.ScaleWidth
    
    'On Error Resume Next
    For Each ctrl In formLife.Controls
        
        '** why is this here?
        If ctrl.Tag > "" Then Stop
        
        
        'stop trying to read a timer control which does not have a run time interface
        If (Not TypeOf ctrl Is Timer) And (Not TypeOf ctrl Is Menu) Then
            sTmp = ""
            sTmp = sTmp + CStr(ctrl.Top)
            sTmp = sTmp + vbCrLf
            sTmp = sTmp + CStr(ctrl.Left)
            sTmp = sTmp + vbCrLf
            sTmp = sTmp + CStr(ctrl.Height)
            sTmp = sTmp + vbCrLf
            sTmp = sTmp + CStr(ctrl.Width)
            sTmp = sTmp + vbCrLf
            sTmp = sTmp + CStr(ctrl.Font.Size)
            ctrl.Tag = sTmp
        End If
    Next ctrl
    Randomize
End Sub

Sub drawGrid()
    'draw all life forms.
    Dim X As Single
    Dim Y As Single
    Dim LastX As Single
    Dim LastY As Single
    Dim W As Single
    
    GetNextItemXY LastX, LastY
    lifeCount = 0
    PlantCount = 0
    AntCount = 0
    
    Do
        GetNextItemXY X, Y
        If (X >= LowerX) And (X <= UpperX) And (Y >= LowerY) And (Y <= UpperY) Then
            DrawItemXY X, Y
            If Grid(X, Y).Alive Then
                lifeCount = lifeCount + 1
                If Grid(X, Y).Speed > 0 Then
                    AntCount = AntCount + 1
                Else
                    PlantCount = PlantCount + 1
                End If
            End If
        End If
    Loop Until (X = LastX) And (Y = LastY)
    
    lblPlantCount.Caption = Str(PlantCount)
    lblAntCount.Caption = Str(AntCount)
End Sub

Sub DrawItemXY(X As Single, Y As Single)
    'draw a single life form.
    Dim W As Single
    
    Call GridPos(X, Y)
    
    With Grid(X, Y)
        If .Alive Then
            W = .Width * ZoomFactor
            If W < 1 Then W = 1
            picLife.DrawWidth = W  ' Set DrawWidth.
            picLife.PSet ((.NextX - LowerX) * mulX, (.NextY - LowerY) * mulY), .RGB
            picLife.Line -((X - LowerX) * mulX, (Y - LowerY) * mulY), .RGB 'Draw to anchor point.
        End If
    End With
End Sub

Private Sub Form_Paint()
    'force a redraw when necessary
    Call Form_Resize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'exit the form
    ExitForm = True
    Unload Me
End Sub

Private Sub Form_Resize()
    'Make sure everything fits in the window.
    Dim ctrl As Control
    Dim tmpSa() As String
    Dim tmpHeight As Single
    Dim tmpWidth As Single
    Dim tmpSize As Single
    
    tmpHeight = heightRatio * 0.97
    tmpWidth = widthRatio * 0.99
    
    If tmpHeight < 0.25 Then tmpHeight = 0.25
    If tmpWidth < 0.25 Then tmpWidth = 0.25
    
    tmpSize = lessNotZero(tmpHeight, tmpWidth)
    
    'On Error Resume Next
    For Each ctrl In formLife.Controls
        
        'is there anything to process
        If (ctrl.Tag <> "") And (Not TypeOf ctrl Is Timer) And (Not TypeOf ctrl Is Menu) Then
            tmpSa() = Split(ctrl.Tag, vbCrLf)
            ctrl.Top = (tmpHeight * Val(tmpSa(0)))
            ctrl.Left = (tmpWidth * Val(tmpSa(1)))
            ctrl.Height = (tmpHeight * Val(tmpSa(2)))
            ctrl.Width = (tmpWidth * Val(tmpSa(3)))
            ctrl.Font.Name = "Arial"
            ctrl.Font.Size = (tmpSize * Val(tmpSa(4)))
            ctrl.Refresh
        End If
    Next ctrl
    'On Error GoTo 0
    
    mulX = picLife.ScaleWidth / GridSizeX * ZoomFactor
    mulY = picLife.ScaleHeight / GridSizeY * ZoomFactor
    LowerX = (GridSizeX / ZoomFactor * (ZoomFactor - 1)) / 2
    UpperX = (GridSizeX - LowerX)
    RangeX = (UpperX - LowerX)
    LowerY = (GridSizeY / ZoomFactor * (ZoomFactor - 1)) / 2
    UpperY = (GridSizeY - LowerY)
    RangeY = (UpperY - LowerY)
    
    If Abs(OffsetX) > RangeX * 4 Then OffsetX = OffsetX * 0.85
    If Abs(OffsetY) > RangeY * 4 Then OffsetY = OffsetY * 0.85
    
    LowerX = LowerX + OffsetX
    LowerY = LowerY + OffsetY
    UpperX = UpperX + OffsetX
    UpperY = UpperY + OffsetY
    
    picLife.Cls
    drawGrid
End Sub

Private Sub mChkAllowPlagues_Click()
    mChkAllowPlagues.Checked = Not mChkAllowPlagues.Checked
End Sub

Private Sub mChkAllowPollen_Click()
    mChkAllowPollen.Checked = Not mChkAllowPollen.Checked
End Sub

Private Sub mChkAllowSpores_Click()
    mChkAllowSpores.Checked = Not mChkAllowSpores.Checked
End Sub

Private Sub mChkBreakRules_Click()
    mChkBreakRules.Checked = Not mChkBreakRules.Checked
End Sub

Private Sub mChkEnableSounds_Click()
    mChkEnableSounds.Checked = Not mChkEnableSounds.Checked
End Sub

Private Sub mChkFastEvolve_Click()
    mChkFastEvolve.Checked = Not mChkFastEvolve.Checked
End Sub

Private Sub mChkInvertGreen_Click()
    mChkInvertGreen.Checked = Not mChkInvertGreen.Checked
    setBackgroundColor
    drawGrid
End Sub

Private Sub mChkInvertBlue_Click()
    mChkInvertBlue.Checked = Not mChkInvertBlue.Checked
    setBackgroundColor
    drawGrid
End Sub

Private Sub mChkInvertRed_Click()
    mChkInvertRed.Checked = Not mChkInvertRed.Checked
    setBackgroundColor
    drawGrid
End Sub

Private Sub mChkNoAntEatAnt_Click()
    mChkNoAntEatAnt.Checked = Not mChkNoAntEatAnt.Checked
End Sub

Private Sub mChkNoAntEatCarion_Click()
    mChkNoAntEatCarion.Checked = Not mChkNoAntEatCarion.Checked
End Sub

Private Sub mChkNoAntEatDirt_Click()
    mChkNoAntEatDirt.Checked = Not mChkNoAntEatDirt.Checked
End Sub

Private Sub mChkNoAntEatPlant_Click()
    mChkNoAntEatPlant.Checked = Not mChkNoAntEatPlant.Checked
End Sub

Private Sub mChkNoCloneAnt_Click()
    mChkNoCloneAnt.Checked = Not mChkNoCloneAnt.Checked
End Sub

Private Sub mChkNoClonePlant_Click()
    mChkNoClonePlant.Checked = Not mChkNoClonePlant.Checked
End Sub

Private Sub mChkNoFatalAge_Click()
    mChkNoFatalAge.Checked = Not mChkNoFatalAge.Checked
End Sub

Private Sub mChkNoFatalHunger_Click()
    mChkNoFatalHunger.Checked = Not mChkNoFatalHunger.Checked
End Sub

Private Sub mChkNoMateAntAnt_Click()
    mChkNoMateAntAnt.Checked = Not mChkNoMateAntAnt.Checked
End Sub

Private Sub mChkNoMateAntPlant_Click()
    mChkNoMateAntPlant.Checked = Not mChkNoMateAntPlant.Checked
End Sub

Private Sub mChkNoMatePlantAnt_Click()
    mChkNoMatePlantAnt.Checked = Not mChkNoMatePlantAnt.Checked
End Sub

Private Sub mChkNoMatePlantPlant_Click()
    mChkNoMatePlantPlant.Checked = Not mChkNoMatePlantPlant.Checked
End Sub

Private Sub mChkNoPlantEatAnt_Click()
    mChkNoPlantEatAnt.Checked = Not mChkNoPlantEatAnt.Checked
End Sub

Private Sub mChkNoPlantEatCarion_Click()
    mChkNoPlantEatCarion.Checked = Not mChkNoPlantEatCarion.Checked
End Sub

Private Sub mChkNoPlantEatPlant_Click()
    mChkNoPlantEatPlant.Checked = Not mChkNoPlantEatPlant.Checked
End Sub

Private Sub mChkQuickStart_Click()
    mChkQuickStart.Checked = Not mChkQuickStart.Checked
End Sub

Private Sub mChkZoomFactor_Click()
    Dim s As String
    Dim c As String
    'On Error Resume Next
    If ZoomFactor = 1 Then
        c = "You may choose to zoom in on the center, but note that life-forms outside of the area displayed will be suspended in time until you zoom back out."
        c = c + vbCrLf + vbCrLf + "Please enter the desired zoom factor, in the range of 1 to 20"
        s = InputBox(c, "Set Zoom Factor", "2")
        ZoomFactor = Val(s)
        If (ZoomFactor < 1) Or (ZoomFactor > 20) Then ZoomFactor = 1
    Else
        ZoomFactor = 1
    End If
    'On Error GoTo 0
    mChkZoomFactor.Checked = (ZoomFactor <> 1)
    
    If mChkZoomFactor.Checked Then
        mChkZoomFactor.Caption = "Zoom Factor = " + Trim(Str(ZoomFactor))
    Else
        mChkZoomFactor.Caption = "Zoom In (normal = 1)"
    End If
    Form_Resize
End Sub


Private Sub picLife_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 97
        OffsetX = OffsetX - RangeX / 10
        OffsetY = OffsetY + RangeY / 10
    Case 98, 40 'Down
        OffsetY = OffsetY + RangeY / 10
    Case 99
        OffsetX = OffsetX + RangeX / 10
        OffsetY = OffsetY + RangeY / 10
    Case 100, 37 'Left
        OffsetX = OffsetX - RangeX / 10
    Case 101
        OffsetX = 0
        OffsetY = 0
    Case 102, 39 'Right
        OffsetX = OffsetX + RangeX / 10
    Case 103
        OffsetX = OffsetX - RangeX / 10
        OffsetY = OffsetY - RangeY / 10
    Case 104, 38 'Up
        OffsetY = OffsetY - RangeY / 10
    Case 105
        OffsetX = OffsetX + RangeX / 10
        OffsetY = OffsetY - RangeY / 10
    Case 107, 187 'Plus
        ZoomFactor = ZoomFactor * 1.01 + 0.1
    Case 109, 189 'Minus
        ZoomFactor = (ZoomFactor - 0.1) / 1.01
    End Select
    If ZoomFactor < 0.1 Then ZoomFactor = 0.1
    If ZoomFactor > 50 Then ZoomFactor = 50
    mChkZoomFactor.Checked = (ZoomFactor <> 1)
    If mChkZoomFactor.Checked Then
        mChkZoomFactor.Caption = "Zoom Factor = " + Trim(Str(ZoomFactor))
    Else
        mChkZoomFactor.Caption = "Zoom In (normal = 1)"
    End If
    Form_Resize
End Sub

Private Sub picLife_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Lf As Creature
    Dim XX As Single
    Dim YY As Single
    
    XX = CLng(LowerX + X / mulX)
    YY = CLng(LowerY + Y / mulY)
    
    'On Error Resume Next 'To avoid errors while scrolling the display
    
    If Button = vbRightButton Then
        If Not Grid(XX, YY).Alive Then
            GetNearestLiveNeighborXY XX, YY, False, False, True, True, True
        End If
        If Grid(XX, YY).Alive Then
            Lf = Grid(XX, YY)
            killItemXY XX, YY
        End If
    
    ElseIf Button = vbLeftButton Then
        If Not Grid(XX, YY).Alive Then
            Grid(XX, YY) = Lf
            BirthItemXY XX, YY
            Grid(XX, YY).Energy = Abs(Grid(XX, YY).Energy) + 0.3 * Rnd
            Grid(XX, YY).redEnergy = Grid(XX, YY).redEnergy + 0.2 * Rnd
            Grid(XX, YY).greenEnergy = Grid(XX, YY).greenEnergy + 0.2 * Rnd
            Grid(XX, YY).blueEnergy = Grid(XX, YY).blueEnergy + 0.2 * Rnd
        End If
    End If
    
    'On Error GoTo 0
End Sub

Private Sub tmrMultiTask_Timer()
    'update the display
    lblPlantCount.Caption = Str(PlantCount)
    lblAntCount.Caption = Str(AntCount)
    picLife.Cls
    Call drawGrid
End Sub

Function heightRatio() As Double
    heightRatio = formLife.ScaleHeight / initialHeight
End Function

Function widthRatio() As Double
    widthRatio = formLife.ScaleWidth / initialWidth
End Function

Function lessNotZero(ByVal A, ByVal b)
    'returns the lesser value, provided that value is not zero.
    'returns zero only if both values provided are zero.
    If (A = 0) Or ((b < A) And (b <> 0)) Then A = b
    lessNotZero = A
End Function

