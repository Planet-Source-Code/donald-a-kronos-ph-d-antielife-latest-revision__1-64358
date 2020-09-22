Attribute VB_Name = "LifeForm"
Option Explicit

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Type Creature
    'attributes for life forms...
    Alive As Boolean 'Initialize phenotype when Alive gets set to True.
    
    Chromosome As String 'Genes are stored here.
    activeGene As Long 'Gene number waiting to be processed next.
    geneProcessAge As Long 'Earliest age at which next gene may be processed.
        
    Age As Long 'Count of times this life form has been processed.
    MaturityAge As Long 'Minimum age to produce offspring.
    MatingAge As Long 'Minimum age to mate.
    CloningAge As Long 'Minimum age to self reproduce.
    CertainDeathAge As Long 'Absolute maximum life-span.
    
    Health As Single 'Meant to be an indiator of wellness in the range of 0="critical" to 1="thriving".
    Speed As Single 'Speed of motion expected
    Velocity As Single 'Speed of motion in effect
    Red As Integer 'Display this level of Red. (See Note.)
    Green As Integer 'Display this level of Green. (See Note.)
    Blue As Integer 'Display this level of Blue. (See Note.)
    'Note... Adjusted automatically before use to ensure visibility.
    RGB As Long 'Stores actual calculated DisplayColor.
    
    Width As Single 'Pen width, or thinckness of creature.
    Orientation As Single
    minMates As Integer 'Minimum number of mates needed to reproduce.  The default of 0 allows self reproduction.
    
    NextX As Single
    NextY As Single
    
    NearSenseAngle As Single
    FarSenseAngle As Single
    
    'if speed is zero, nothing from here down is used.
    length As Integer 'Additional length to be acquired before moving to NextX, NextY.
    Energy As Double 'Energy available for use at this time.
    Hungry As Double 'Below this level of energy, life form is hungry.
    EnergyToMate As Double 'Mimimum Energy Needed Before Mating.
    EnergyToClone As Double  'Mimimum Energy Needed Before Self Reproduction.
    redEnergy As Double 'Energy from Red food, not yet ready for use.
    greenEnergy As Double 'Energy from Green food, not yet ready for use.
    blueEnergy As Double 'Energy from Blue food, not yet ready for use.
    CloneCountDown As Integer  ' off = 0.  Clone now = 1. Higher numbers = counter.
    mateRed As Integer 'Will see this level of Red as a potential mate.
    mateGreen As Integer 'Will see this level of Green as a potential mate.
    mateBlue As Integer 'Will see this level of Blue as a potential mate.
    foodRed As Integer 'Will see this level of Red as potential food.
    foodGreen As Integer 'Will see this level of Green as potential food.
    foodBlue As Integer 'Will see this level of Blue as potential food.
End Type

'GridSizeX and GridSizeY define the dimensions of the grid on which the life forms live.
Public Const GridSizeX = 700
Public Const GridSizeY = 600

'used in positioning the display...
Public LowerX As Single
Public RangeX As Single
Public UpperX As Single
Public LowerY As Single
Public RangeY As Single
Public UpperY As Single
Public OffsetX As Single
Public OffsetY As Single
Public mulX As Double
Public mulY As Double

Public ZoomFactor As Single

Public Grid(GridSizeX - 1, GridSizeY - 1) As Creature
Public ListXY As String
Public lifeCount As Long
Public PlantCount As Long
Public AntCount As Long

Private itemPosition As Long

Function setBackgroundColor()
    Dim R As Integer
    Dim G As Integer
    Dim b As Integer
    Dim bColor As Long
    Dim tColor As Long 'Label text color
    R = IIf(formLife.mChkInvertRed.Checked, 255, 0)
    G = IIf(formLife.mChkInvertGreen.Checked, 255, 0)
    b = IIf(formLife.mChkInvertBlue.Checked, 255, 0)
    bColor = RGB(R, G, b)
    'PictureBox background...
    formLife.picLife.BackColor = bColor
    'Form and label colors...
    tColor = RGB(255 - R, Abs(192 - G), Abs(192 - b))
    formLife.BackColor = bColor
    formLife.lblAntCountLbl.BackColor = bColor
    formLife.lblPlantCountLbl.BackColor = bColor
    formLife.lblAntCount.BackColor = bColor
    formLife.lblPlantCount.BackColor = bColor
    formLife.lblAntCountLbl.ForeColor = tColor
    formLife.lblPlantCountLbl.ForeColor = tColor
    formLife.lblAntCount.ForeColor = tColor
    formLife.lblPlantCount.ForeColor = tColor
End Function

Function DisplayColor(X As Single, Y As Single) As Long
    Dim R As Integer
    Dim G As Integer
    Dim b As Integer
    With Grid(X, Y)
        R = .Red
        G = .Green
        b = .Blue
    End With
    If (R = G) And (G = b) Then
        R = R + 32
        G = G + 32
        b = b + 32
    ElseIf (R < G) And (R < b) Then
        G = G + 48
        b = b + 48
    ElseIf (G < R) And (G < b) Then
        R = R + 48
        b = b + 48
    ElseIf (b < R) And (b < G) Then
        R = R + 48
        G = G + 48
    ElseIf (R > G) And (R > b) Then
        R = R + 64
    ElseIf (G > R) And (G > b) Then
        G = G + 64
    ElseIf (b > R) And (b > G) Then
        b = b + 64
    End If
    'Note: RGB values may be higher than 255. That's why Abs() is used here...
    If formLife.mChkInvertRed.Checked Then R = Abs(255 - R)
    If formLife.mChkInvertGreen.Checked Then G = Abs(255 - G)
    If formLife.mChkInvertBlue.Checked Then b = Abs(255 - b)
    DisplayColor = RGB(R, G, b)
End Function

Sub AddItemXY(X As Single, Y As Single)
    Static entry As String
    entry = makeEntryXY(X, Y)
    If InStr(ListXY, entry) Then Exit Sub
    ListXY = ListXY + entry
    lifeCount = lifeCount + 1
    If Grid(X, Y).Speed <= 0 Then
        PlantCount = PlantCount + 1
    Else
        AntCount = AntCount + 1
    End If
End Sub

Function makeEntryXY(X As Single, Y As Single)
    makeEntryXY = "[" + Hex(inGridX(CLng(X))) + "," + Hex(inGridY(CLng(Y))) + "]"
End Function

Function RemoveItemEntry(entry) As Boolean
    Static i As Long
    If entry = "" Then
        RemoveItemEntry = False
        Exit Function
    End If
    i = InStr(ListXY, entry)
    If i = 0 Then
        RemoveItemEntry = False
        Exit Function
    End If
    If itemPosition > i Then itemPosition = itemPosition - Len(entry)
    ListXY = Left(ListXY, i - 1) + Mid(ListXY, i + Len(entry))
    lifeCount = lifeCount - 1
    RemoveItemEntry = True
    If itemPosition > 1 Then itemPosition = itemPosition - 1
End Function

Sub RemoveItemXY(X As Single, Y As Single)
    If RemoveItemEntry(makeEntryXY(X, Y)) Then
        If Grid(X, Y).Speed <= 0 Then
            PlantCount = PlantCount - 1
        Else
            AntCount = AntCount - 1
        End If
    End If
End Sub

Sub GetNextItemXY(ByRef X As Single, ByRef Y As Single)
    Static ii As Long
    Static entry
    If ListXY < "[" Then Exit Sub
    itemPosition = InStr(itemPosition + 1, ListXY, "[")
    If itemPosition = 0 Then itemPosition = 1 'Return to first item
    ii = InStr(itemPosition, ListXY, "]")
    entry = Mid(ListXY, itemPosition, ii - itemPosition + 1)
    ii = InStr(entry, ",")
    X = Val("&h" + Mid(entry, 2))
    Y = Val("&h" + Mid(entry, ii + 1))
End Sub

Sub ResetItemXY(X As Single, Y As Single)
    Dim tmp As Creature
    Grid(X, Y) = tmp
End Sub

Sub reBirthItemXY(X As Single, Y As Single)
    Dim baby As Creature
    baby = Grid(X, Y)
    BirthItemXY X, Y
    With Grid(X, Y)
        .Energy = baby.Energy
        .redEnergy = baby.redEnergy
        .greenEnergy = baby.greenEnergy
        .blueEnergy = baby.blueEnergy
        .Orientation = baby.Orientation
    End With
End Sub
Sub BirthItemXY(X As Single, Y As Single)
    Dim Genes As String
    Dim baby As Creature
    Genes = Grid(X, Y).Chromosome
    Grid(X, Y) = baby
    With Grid(X, Y)
        .Chromosome = Genes
        .Alive = True
        .Health = 1 'Thriving
        'Set to default phenotype...
        .Red = 77
        .Green = 99
        .Blue = 82
        .mateRed = 77
        .mateGreen = 99
        .mateBlue = 82
        .foodRed = 77
        .foodGreen = 99
        .foodBlue = 82
        .CertainDeathAge = 23
        .MatingAge = 0
        .CloningAge = 0
        .MaturityAge = 0
        .Hungry = 0
        .Orientation = Rnd * pi * 2 'random facing direction
        .activeGene = 0 'Reset to default active gene.
        'Process a few genes normally (if available) before birth...
        processNextGeneXY X, Y
        processNextGeneXY X, Y
        processNextGeneXY X, Y
        processNextGeneXY X, Y
        processNextGeneXY X, Y
        'Make sure item will display...
        If .Width < 0.1 Then .Width = 1
        .NextX = X + Cos(.Orientation) * .Width
        .NextY = Y + Sin(.Orientation) * .Width
        If .CertainDeathAge < 1 Then .CertainDeathAge = 1
        If .Width > 33 Then .Width = 33
        .RGB = DisplayColor(X, Y)
    End With
    AddItemXY X, Y 'Add coordinates of newborn to list
End Sub

Function biteItemXY(X As Single, Y As Single, Optional biteCount As Single = 1) As Boolean
    ' take a bite out of grid(x,y) and return a boolean indicating whether or not .health is still greater than 0.
    Dim sTmp As Single
    With Grid(X, Y)
        'note: biteCount could also simulate mouth size.
        sTmp = 1 / (2 + .Width + .length) * biteCount
        .Health = .Health - sTmp
        biteItemXY = (.Health > 0)
    End With
End Function

Function healItemXY(X As Single, Y As Single) As Boolean
    ' Restore some health to grid(x,y) and return a boolean indicating whether or not .health has reached or exceded 1.
    Dim sTmp As Single
    If Grid(X, Y).Health >= 1 Then
        healItemXY = True
        Exit Function
    Else
        With Grid(X, Y)
            sTmp = 1 / (3 + .Width + .length)
            If sTmp >= .Energy / 2 Then sTmp = .Energy / 2 + 0.1
            .Health = .Health + sTmp
            .Energy = .Energy - sTmp 'Energy used to heal.
            healItemXY = (.Health >= 1)
        End With
    End If
End Function

Function getNextGeneValueXY(X As Single, Y As Single)
    Dim t() As Byte
    With Grid(X, Y)
        If .activeGene < 1 Then .activeGene = 1
        If .activeGene > Len(.Chromosome) Then
            getNextGeneValueXY = 0
            Exit Function
        End If
        t() = Mid(.Chromosome, .activeGene, 1)
        getNextGeneValueXY = PickGeneValue(t(0), t(1))
        .activeGene = .activeGene + 1
        'On Error GoTo 0
    End With
    Exit Function
End Function

Sub processNextGeneXY(X As Single, Y As Single)
    Dim value As Integer
    Dim mode As Integer
    With Grid(X, Y)
        If .Age < .geneProcessAge Then Exit Sub
        If .activeGene > Len(.Chromosome) - 2 Then Exit Sub
        mode = getNextGeneValueXY(X, Y)
'        If mode > 32 Then mode = mode - 32 'Don't waste genes.
        value = getNextGeneValueXY(X, Y)
        Select Case mode
        Case 1
            .Red = Abs(.Red + value)
        Case 2
            .mateRed = Abs(.mateRed + value)
        Case 3
            .Green = Abs(.Green + value)
        Case 4
            .mateGreen = Abs(.mateGreen + value)
        Case 5
            .Blue = Abs(.Blue + value)
        Case 6
            .mateBlue = Abs(.mateBlue + value)
        Case 7
            .foodRed = Abs(.foodRed + value)
        Case 8
            .foodGreen = Abs(.foodGreen + value)
        Case 9
            .foodBlue = Abs(.foodBlue + value)
        Case 10
            .Speed = .Speed + 0.01 * value
        Case 11
            .MaturityAge = .MaturityAge + value
        Case 12
            .Hungry = .Hungry + value
        Case 13
            .CertainDeathAge = .CertainDeathAge + value * 3
        Case 14
            .CertainDeathAge = .CertainDeathAge + value
            .Speed = .Speed + Log(value + 1) 'alternate animal gene
        Case 15
            .CloningAge = .CloningAge + value
        Case 16
            .MatingAge = .MatingAge + value
        Case 17
            .length = greaterOf(.length / 2 + value, .length + 1)
        Case 18
            .minMates = 1
        Case 19
            .activeGene = .activeGene + value
        Case 20
            .geneProcessAge = .Age + value + 2
        Case 21
            .NearSenseAngle = .NearSenseAngle + value / 8 - 1
        Case 22
            .FarSenseAngle = .FarSenseAngle + value / 8 - 1
        Case 23
            .Red = Abs(.Red + value * 2 - 16)
        Case 24
            .mateRed = Abs(.mateRed + value * 2 - 16)
        Case 25
            .Green = Abs(.Green + value * 2 - 16)
        Case 26
            .mateGreen = Abs(.mateGreen + value * 2 - 16)
        Case 27
            .Blue = Abs(.Blue + value * 2 - 16)
        Case 28
            .mateBlue = Abs(.mateBlue + value * 2 - 16)
        Case 29
            .foodRed = Abs(.foodRed + value * 2 - 16)
        Case 30
            .foodGreen = Abs(.foodGreen + value * 2 - 16)
        Case 31
            .foodBlue = Abs(.foodBlue + value * 2 - 16)
        Case 32
            .Width = .Width + 1
        Case 33
            .Red = Abs(.Red - value)
        Case 34
            .mateRed = Abs(.mateRed - value)
        Case 35
            .Green = Abs(.Green - value)
        Case 36
            .mateGreen = Abs(.mateGreen - value)
        Case 37
            .Blue = Abs(.Blue - value)
        Case 38
            .mateBlue = Abs(.mateBlue - value)
        Case 39
            .foodRed = Abs(.foodRed - value)
        Case 40
            .foodGreen = Abs(.foodGreen - value)
        Case 41
            .foodBlue = Abs(.foodBlue - value)
        End Select
        'On Error GoTo 0
    End With
    Exit Sub
End Sub

Function PickGeneValue(gM, gP)  'Decides which of two alleles to use.
    PickGeneValue = IIf(gM And &H80, IIf(gP And &H80, (gM Or gP) And &H7F, gM And &H7F), IIf(gP And &H80, gP And &H7F, (gM And gP) And &H7F))
End Function

Function greaterOf(A, b)
    greaterOf = IIf(A > b, A, b)
End Function

Function lesserOf(A, b)
    lesserOf = IIf(A < b, A, b)
End Function

Sub firstLife(Optional ByVal number As Long = 1)      '** NEVER use Variant
    Dim baby As Creature
    Dim X As Single
    Dim Y As Single
    Dim i As Long           '** NEVER use Variant
    Dim ii As Long          '** NEVER use Variant
    Dim t() As Byte
    Dim s As String
    Dim ct As String
    Dim errorCount As Long
    
    ReDim t(3)
    t(0) = 10 'Speed adjust mode... to produce an animal.
    t(1) = 10
    t(2) = 3 'Adjustment value
    t(3) = 3
    ct = Left(t(), 1)
    
    If number < 0 Then
        baby.Chromosome = t()
    End If
    If (formLife.mChkQuickStart.Checked Xor breakRules) Then
        'add some random genes...
        s = Chr(Int(Rnd * 64)) + Chr(Int(Rnd * 64)) + Chr(Int(Rnd * 64)) + Chr(Int(Rnd * 64))
        s = Replace(s, ct, Chr(0)) ' Make sure these extra genes don't turn plants into animals.
        baby.Chromosome = baby.Chromosome + s
    End If
    If number < 0 Then
        s = t() 'speed gene
    End If
    baby.Alive = True 'So that mutations can happen.
    
    'On Error Resume Next 'To avoid errors while scrolling the display
    
    For i = 1 To Abs(number)
        ii = 0
getNewXY:
        Do
            X = inGridX(GetRndInt(RangeX - 1) + LowerX)
            Y = inGridY(GetRndInt(RangeY - 1) + LowerY)
            ii = ii + 1
        Loop Until (ii > GridSizeX + GridSizeY) Or Not Grid(X, Y).Alive
        Grid(X, Y) = baby
        
        
        '** I'm assuming this has a 10% chance of mutation...?
        If (formLife.mChkQuickStart.Checked Xor breakRules) Or (Rnd > 0.9) Then
            mutateItemXY X, Y
            If number < 1 Then
                Grid(X, Y).activeGene = (InStr(Grid(X, Y).Chromosome, Left(s, 1)) + 1) / 2
                Grid(X, Y).Speed = 0 'plant speed.
                processNextGeneXY X, Y
                If Grid(X, Y).Speed <= 0 Then
                    Grid(X, Y).Chromosome = s + Grid(X, Y).Chromosome 'restore initial animal gene.
                End If
            Else
                baby.Chromosome = Replace(baby.Chromosome, ct, Chr(0)) ' Make sure these extra genes don't turn plants into animals.
            End If
            baby = Grid(X, Y) 'So that the next generated baby will carry any mutations gained by this one.
        End If
        
        BirthItemXY X, Y
        'On Error GoTo handleErrorWend 'To keep from getting caught in a loop if out of bounds.
        Do While Sgn(Grid(X, Y).Speed) = Sgn(number) 'wrong kind of life form.  Try again.
            mutateItemXY X, Y
            reBirthItemXY X, Y
        Loop
        
afterWend:
        'On Error Resume Next 'To avoid errors while scrolling the display
        Grid(X, Y).Energy = Grid(X, Y).Energy + 0.3 * Rnd
        Grid(X, Y).redEnergy = Grid(X, Y).redEnergy + 0.2 * Rnd
        Grid(X, Y).greenEnergy = Grid(X, Y).greenEnergy + 0.2 * Rnd
        Grid(X, Y).blueEnergy = Grid(X, Y).blueEnergy + 0.2 * Rnd
    Next i
    
    'On Error GoTo 0
    
    Exit Sub
handleErrorWend:
    errorCount = errorCount + 1
    If errorCount > Abs(number) Then
        Resume afterWend
    Else
        Resume getNewXY
    End If
End Sub

Sub deleteLife(Optional number = 0)
    Dim delCount As Long
    Dim entry As String
    Dim X As Single
    Dim Y As Single
    Dim delAnts As Boolean
    Dim n As Long
    If number = 0 Then Exit Sub
    If ListXY < "[" Then Exit Sub
    delAnts = (number < 0)
    n = Abs(number)
    n = IIf(delAnts, lesserOf(n, AntCount), lesserOf(n, PlantCount))
    delCount = 0
    itemPosition = 1
    Do
        GetNextItemXY X, Y
        If (X >= LowerX) And (X <= UpperX) And (Y >= LowerY) And (Y <= UpperY) Then
            If (delAnts) Eqv (Grid(X, Y).Speed > 0) Then
                killItemXY X, Y
                delCount = delCount + 1
            End If
        End If
    Loop Until (itemPosition <= 1) Or (delCount >= n)
End Sub

Function neighborCount(X As Single, Y As Single) As Integer
    Dim XX As Single
    Dim YY As Single
    Dim count As Integer
    Dim W As Single
    Dim iw As Single
    Dim i As Single
    
    Call GridPos(X, Y)
    
    If Not Grid(X, Y).Alive Then
        neighborCount = 0
        Exit Function
    End If
    'On Error Resume Next 'To avoid errors while scrolling the display
    With Grid(X, Y)
        W = .Width
        iw = 1 / (W + .length) * 0.75
        For i = iw To pi2 Step iw
            XX = inGridX(X + Cos(i + .Orientation) * (W * sqr2) + .length)
            YY = inGridY(Y + Sin(i + .Orientation) * (W * sqr2) + .length)
            Call GridPos(XX, YY)
            If Grid(XX, YY).Alive Then count = count + 1
        Next i
    End With
    'On Error GoTo 0
    neighborCount = count
End Function

Function GetLiveNeighborXY(ByRef X As Single, ByRef Y As Single, distance, Optional mustBeFood As Boolean = False, Optional mustBePotentialMate As Boolean = False, Optional mayBeAnt As Boolean = True, Optional mayBePlant As Boolean = True, Optional clickSearch As Boolean = False) As Boolean
    Dim pX As Single
    Dim pY As Single
    Dim XX As Single
    Dim YY As Single
    Dim count As Integer
    Dim W As Single
    Dim iw As Single
    Dim i As Single
    Dim ii As Integer
    Dim startTime As Single
    'On Error Resume Next 'To avoid errors while scrolling the display
    If Not Grid(X, Y).Alive And Not clickSearch Then
        GetLiveNeighborXY = False
        Exit Function
    End If
    'On Error GoTo 0
    If Not (mayBeAnt Or mayBePlant) Then
        GetLiveNeighborXY = False
        Exit Function
    End If
    With Grid(X, Y)
        W = distance
        iw = 1 / W * 0.5
        i = 0
        startTime = Timer
        'On Error Resume Next 'To avoid errors while scrolling the display
        Do
            Do
                Do
                    Do
                        Do
                            pX = Cos(.Orientation + i) * W
                            pY = Sin(.Orientation + i) * W
                            XX = inGridX(X + pX)
                            YY = inGridY(Y + pY)
                            
                            Call GridPos(XX, YY)
                            
                            If Not Grid(XX, YY).Alive Then
                                pX = Cos(.Orientation - i) * W
                                pY = Sin(.Orientation - i) * W
                                XX = inGridX(X + pX)
                                YY = inGridY(Y + pY)
                            End If
                            i = i + iw
                            
                            Call GridPos(XX, YY)
                        Loop Until CInt(XX) <> CInt(X) And CInt(YY) <> CInt(Y)
                    Loop Until (i > pi) Or Grid(XX, YY).Alive Or (Abs(startTime - Timer) > 2)
                Loop Until (i > pi) Or isMate(X, Y, XX, YY) Or (Abs(startTime - Timer) > 2) Or Not mustBePotentialMate
            Loop Until (i > pi) Or isFood(X, Y, XX, YY) Or (Abs(startTime - Timer) > 2) Or Not mustBeFood
        Loop Until (i > pi) Or (isAnt(X, Y) And mayBeAnt) Or (isPlant(X, Y) And mayBePlant) Or (Abs(startTime - Timer) > 2)
        If Grid(XX, YY).Alive Then
            X = XX
            Y = YY
            GetLiveNeighborXY = True
        Else
            GetLiveNeighborXY = False
        End If
        'On Error GoTo 0
    End With
End Function

Function isAnt(X As Single, Y As Single) As Boolean
    isAnt = Grid(X, Y).Speed > 0
End Function

Function isPlant(X As Single, Y As Single) As Boolean
    isPlant = Grid(X, Y).Speed <= 0
End Function

Function isMate(sourceX As Single, sourceY As Single, mateX As Single, mateY As Single) As Boolean
    Dim match As Boolean
    If Not Grid(mateX, mateY).Alive Then
        isMate = False
        Exit Function
    End If
    match = isMateColor(sourceX, sourceY, Grid(mateX, mateY).Red, Grid(mateX, mateY).Green, Grid(mateX, mateY).Blue)
    If Abs(Grid(sourceX, sourceY).Width - Grid(mateX, mateY).Width) > 2 Then
        match = False 'Don't mate is width is too different.
    ElseIf Abs(Grid(sourceX, sourceY).length - Grid(mateX, mateY).length) > 3 Then
        match = False 'Don't mate is length is too different.
    ElseIf Abs(Grid(sourceX, sourceY).Speed - Grid(mateX, mateY).Speed) > 0.4 Then
        match = False 'Don't mate is speed is too different.
    End If
End Function

Function isMateColor(sourceX, sourceY, Red, Green, Blue) As Boolean
    Dim match As Boolean
    match = isSameColor(Grid(sourceX, sourceY).mateRed, Grid(sourceX, sourceY).mateGreen, Grid(sourceX, sourceY).mateBlue, Red, Green, Blue)
    isMateColor = match
End Function

Function isSame(sourceX As Single, sourceY As Single, targetX As Single, targetY As Single) As Boolean
    'Intended for species matching, by color.
    isSame = isSameColor(Grid(sourceX, sourceY).Red, Grid(sourceX, sourceY).Green, Grid(sourceX, sourceY).Blue, Grid(targetX, targetY).Red, Grid(targetX, targetY).Green, Grid(targetX, targetY).Blue)
End Function

Function isNearlySameColor(sourceRed, sourceGreen, sourceBlue, matchRed, matchGreen, matchBlue) As Boolean
    'Alternative to isSameColor.
    Dim match As Boolean
    match = False
    If Abs(sourceRed - matchRed) < 3 Then match = True
    If Abs(sourceGreen - matchGreen) < 3 Then match = True
    If Abs(sourceBlue - matchBlue) < 3 Then match = True
    isNearlySameColor = match
End Function

Function isSameColor(sourceRed, sourceGreen, sourceBlue, matchRed, matchGreen, matchBlue, Optional maxDiff As Integer = 9) As Boolean
    'Three dimensional match by absolute distance.
    Dim match As Boolean
    Dim diffRed As Long
    Dim diffGreen As Long
    Dim diffBlue As Long
    diffRed = sourceRed - matchRed
    diffGreen = sourceGreen - matchGreen
    diffBlue = sourceBlue - matchBlue
    match = False
    If Sqr(diffRed * diffRed + diffGreen * diffGreen + diffBlue * diffBlue) < maxDiff Then match = True
    isSameColor = match
End Function

Function isFood(sourceX As Single, sourceY As Single, foodX As Single, foodY As Single) As Boolean
    If Not Grid(foodX, foodY).Alive Then
        isFood = False
        Exit Function
    End If
    isFood = isFoodColor(sourceX, sourceY, Grid(foodX, foodY).Red, Grid(foodX, foodY).Green, Grid(foodX, foodY).Blue)
End Function

Function isFoodColor(sourceX As Single, sourceY As Single, Red, Green, Blue) As Boolean
    Dim match As Boolean
    match = isNearlySameColor(Grid(sourceX, sourceY).foodRed, Grid(sourceX, sourceY).foodGreen, Grid(sourceX, sourceY).foodBlue, Red, Green, Blue)
    isFoodColor = match
End Function

Function GetNearestLiveNeighborXY(ByRef X As Single, ByRef Y As Single, Optional mustBeFood As Boolean = False, Optional mustBePotentialMate As Boolean = False, Optional mayBeAnt As Boolean = True, Optional mayBePlant As Boolean = True, Optional clickSearch As Boolean = False) As Boolean
    Dim XX As Single
    Dim YY As Single
    Dim count As Integer
    Dim W As Single
    Dim iw As Single
    Dim i As Single
    Dim expandSearch As Boolean
    Dim startTime As Single
    Dim extraRange As Single
    expandSearch = False
    If clickSearch Then
        extraRange = 7
    Else
        extraRange = 0
        If Not Grid(X, Y).Alive Then
            GetNearestLiveNeighborXY = False
            Exit Function
        End If
    End If
    If Not (mayBeAnt Or mayBePlant) Then
        GetNearestLiveNeighborXY = False
        Exit Function
    End If
    If mustBePotentialMate Then
        If Grid(X, Y).Speed > 0 Then
            If (formLife.mChkAllowSpores.Checked Xor breakRules) Then
                If Rnd > 0.3 Then
                    expandSearch = True
                End If
            End If
        Else
            If (formLife.mChkAllowPollen.Checked Xor breakRules) Then
                If Rnd > 0.3 Then
                    expandSearch = True
                End If
            End If
        End If
    End If
    W = Grid(X, Y).Width / 2 + 1
    XX = X
    YY = Y
    startTime = Timer
    While (W <= Grid(X, Y).Width + 1 + extraRange) And (Abs(startTime - Timer) < 5)
        If GetLiveNeighborXY(XX, YY, W, mustBeFood, mustBePotentialMate, mayBeAnt, mayBePlant, clickSearch) Then W = Grid(X, Y).Width + 1 + extraRange
        W = W + 1
    Wend
    If expandSearch And (XX = X) And (YY = Y) Then
        startTime = Timer
        While (W / 15 <= Grid(X, Y).Width + 1 + extraRange) And (Abs(startTime - Timer) < 5)
            If GetLiveNeighborXY(XX, YY, W, mustBeFood, mustBePotentialMate, mayBeAnt, mayBePlant, clickSearch) Then
                W = Grid(X, Y).Width * 15 + 1 + extraRange
            End If
            W = W + 5
        Wend
    End If
    
    'On Error Resume Next 'To avoid errors while scrolling the display
    If ((XX = X) And (YY = Y)) Or Not Grid(XX, YY).Alive Then
        GetNearestLiveNeighborXY = False
    Else
        X = XX
        Y = YY
        GetNearestLiveNeighborXY = True
    End If
    'On Error GoTo 0
End Function

Sub inGridXY(ByRef X As Single, ByRef Y As Single)
'make sure X and Y are in the grid
    While CLng(X) < 0
        X = X + GridSizeX
    Wend
    While CLng(X) >= GridSizeX
        X = X - GridSizeX
    Wend
    While CLng(Y) < 0
        Y = Y + GridSizeY
    Wend
    While CLng(Y) >= GridSizeY
        Y = Y - GridSizeY
    Wend
    While CLng(X) < LowerX
        X = X + RangeX
    Wend
    While CLng(X) >= UpperX
        X = X - RangeX
    Wend
    While CLng(Y) < LowerY
        Y = Y + RangeY
    Wend
    While CLng(Y) >= UpperY
        Y = Y - RangeY
    Wend
End Sub

Function inGridX(ByVal X As Single) As Single
'make sure X is in the grid
    If (CLng(X) >= LowerX) And (CLng(X) < UpperX) Then
        inGridX = X
        Exit Function
    End If
    While CLng(X) < 0
        X = X + GridSizeX
    Wend
    While CLng(X) >= GridSizeX
        X = X - GridSizeX
    Wend
    While CLng(X) < LowerX
        X = X + RangeX
    Wend
    While CLng(X) >= UpperX
        X = X - RangeX
    Wend
    inGridX = X
End Function

Function inGridY(ByVal Y As Single) As Single
'make sure Y is in the grid
    If (CLng(Y) >= LowerY) And (CLng(Y) < UpperY) Then
        inGridY = Y
        Exit Function
    End If
    While CLng(Y) < 0
        Y = Y + GridSizeY
    Wend
    While CLng(Y) >= GridSizeY
        Y = Y - GridSizeY
    Wend
    While CLng(Y) < LowerY
        Y = Y + RangeY
    Wend
    While CLng(Y) >= UpperY
        Y = Y - RangeY
    Wend
    inGridY = Y
End Function

Function sense(X As Single, Y As Single, Direction, distance, Optional mate As Boolean = True, Optional food As Boolean = True) As Boolean
    ' Returns whether or not the indicated item type was sensed in the specified direction.
    ' Direction = 0 means steraight ahead.
    ' Distance = 1 means starting at width distance away and up to 2 width distance away.
    Dim XX As Single
    Dim YY As Single
    Dim Color As Long
    Dim matchFound As Boolean
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer
    Dim i As Single
    matchFound = False
    With Grid(X, Y)
        i = distance * .Width + 1
        Do
            XX = inGridX(X + i * Cos(.Orientation + Direction))
            YY = inGridY(Y + i * Sin(.Orientation + Direction))
            Red = Grid(XX, YY).Red
            Green = Grid(XX, YY).Green
            Blue = Grid(XX, YY).Blue
            If food Then
                If isFoodColor(X, Y, Red, Green, Blue) Then matchFound = True
            End If
            If mate Then
                If isMateColor(X, Y, Red, Green, Blue) Then matchFound = True
            End If
            i = i + 1
        Loop Until matchFound Or (i > distance * (.Width + 1) * 2)
    End With
End Function

Sub AimXY(X As Single, Y As Single, Optional mate As Boolean = True, Optional food As Boolean = True)
    Dim Direction As Single
    Dim Aim As Single
    Aim = 0
    With Grid(X, Y)
        Direction = .NearSenseAngle
        If Direction <> 0 Then
            If sense(X, Y, Direction, 1, mate, food) Then Aim = Aim + Direction
            If sense(X, Y, -Direction, 1, mate, food) Then Aim = Aim - Direction
        End If
        Direction = .FarSenseAngle
        If Direction <> 0 Then
            If sense(X, Y, Direction, 2, mate, food) Then Aim = Aim + Direction
            If sense(X, Y, -Direction, 2, mate, food) Then Aim = Aim - Direction
        End If
        Direction = .FarSenseAngle + .NearSenseAngle - 1
        If Direction <> 0 Then
            If sense(X, Y, Direction, 3, mate, food) Then Aim = Aim + Direction
            If sense(X, Y, -Direction, 3, mate, food) Then Aim = Aim - Direction
        End If
        If Aim <> 0 Then .Orientation = .Orientation + Aim
    End With
End Sub

Sub moveNextXY()
    Static X As Single
    Static Y As Single
    GetNextItemXY X, Y
    moveItemXY X, Y
End Sub

Sub killItemXY(X As Single, Y As Single)
    With Grid(X, Y)
        .Alive = False
        .Chromosome = ""
        RemoveItemXY X, Y
        If formLife.mChkEnableSounds.Checked Then
            killSound
        End If
    End With
End Sub

Sub killSound()
    Static note As Double
    Static timeSinceLast As Single
    Static lastTime As Single
    timeSinceLast = Abs(lastTime - Timer)
    If timeSinceLast > 2 Then
        note = 16
    ElseIf timeSinceLast > 1 Then
        note = note + 8
    ElseIf timeSinceLast > 0.5 Then
        note = note + 4
    Else
        note = note - 1
    End If
    If note > 20 Then note = 20
    If note < -8 Then note = -8
    playNote note, 2, 1, 3000
    lastTime = Timer
End Sub

Sub mateSound()
    Static note As Double
    Static timeSinceLast As Single
    Static lastTime As Single
    timeSinceLast = Abs(lastTime - Timer)
    If timeSinceLast > 2 Then
        note = 16
    ElseIf timeSinceLast > 1 Then
        note = note - 8
    ElseIf timeSinceLast > 0.5 Then
        note = note - 4
    Else
        note = note + 1
    End If
    If note > 20 Then note = 20
    If note < -8 Then note = -8
    playNote note, 4, 1, 2000
    playNote note + 2, 3, 1, 4000
    lastTime = Timer
End Sub

Sub cloneSound()
    Static note As Double
    Static timeSinceLast As Single
    Static lastTime As Single
    timeSinceLast = Abs(lastTime - Timer)
    If timeSinceLast > 2 Then
        note = 8
    ElseIf timeSinceLast > 1 Then
        note = note - 8
    ElseIf timeSinceLast > 0.5 Then
        note = note - 4
    Else
        note = note + 1
    End If
    If note > 12 Then note = 12
    If note < -16 Then note = -16
    playNote note, 3, 1, 2000
    lastTime = Timer
End Sub

Sub plagueItemXY(X As Single, Y As Single)
    Dim XX As Single
    Dim YY As Single
    
    Static i As Integer
    
    '** make sure that we're not outside the bounds of the grid
    Call GridPos(X, Y)
    
    With Grid(X, Y)
        XX = X 'because GetNearestLiveNeighborXY returns values in its parameters.
        YY = Y
        If Not .Alive Then Exit Sub 'can't kill a corpse.
        If Not GetNearestLiveNeighborXY(X, Y) Then Exit Sub 'Nobody around... so illness can't spread.
        If Not isSame(X, Y, XX, YY) Then Exit Sub 'Nearest neighbor is not the same species.
        If neighborCount(X, Y) < 2 Then Exit Sub 'Three's a crowd?
        killItemXY X, Y
        For i = 0 To 2
            XX = X - Int(Rnd * 3 - 1)
            YY = Y - Int(Rnd * 3 - 1)
            If (XX = X) And (YY = Y) Then Exit Sub
            plagueItemXY inGridX(XX), inGridY(YY)
        Next i
    End With
End Sub

Sub moveItemXY(X As Single, Y As Single)
    Dim AimForFood As Boolean
    Dim AimForMate As Boolean
    Dim digestionEnergy As Single
    Dim clearAhead As Boolean
    Dim headToTail
    
    Call GridPos(X, Y)
    
    With Grid(X, Y)
        If .Speed <= 0 Then Exit Sub
        If .Energy + .redEnergy + .greenEnergy + .blueEnergy <= .Hungry Then AimForFood = True
        If (.Age >= .MatingAge) And (.Age > .MaturityAge) And (.Energy >= .EnergyToMate) Then AimForMate = True
        AimXY X, Y, AimForMate, AimForFood
        If .Energy > 0 Then
            .NextX = .NextX + .Velocity * Cos(.Orientation)
            .NextY = .NextY + .Velocity * Sin(.Orientation)
        Else
            'Recover energy from momentum (stop to rest).
            .Energy = .Energy + .Velocity
            .Energy = .Energy + Sqr(squareStepSizeXY(X, Y))
            .Velocity = 0
            .NextX = X
            .NextY = Y
        End If
        clearAhead = Not (Grid(inGridX(CInt(.NextX)), inGridY(CInt(.NextY))).Alive)
        headToTail = squareStepSizeXY(X, Y) 'The square of the distance from head to tail.
        If headToTail > 10000 Then 'has to be wrapped around to other side of universe
            If clearAhead Then
                .Velocity = .Speed
                .NextX = inGridX(CInt(.NextX))
                .NextY = inGridY(CInt(.NextY))
                Grid(.NextX, .NextY) = Grid(X, Y)
                RemoveItemXY X, Y
                AddItemXY .NextX, .NextY
                .Alive = False
                Grid(.NextX, .NextY).Alive = True
                X = inGridX(.NextX)
                Y = inGridY(.NextY)
                Exit Sub
            Else 'Snap back to anchor point if no room to enter other side of universe.
                .NextX = inGridX(X)
                .NextY = inGridY(Y)
            End If
        End If
        If headToTail > .length * .length + 1 Then
            If (.Energy > headToTail) And (clearAhead) Then
                inGridXY .NextX, .NextY
                If Grid(.NextX, .NextY).Alive Then
                    If isMate(X, Y, .NextX, .NextY) Then Exit Sub  'Do not consume a potential mate.
                    If isSame(X, Y, .NextX, .NextY) Then Exit Sub   'Do not consume a possible relative.
                    If Not isFood(X, Y, .NextX, .NextY) Then
                        Exit Sub 'Do not consume a life form that isn't food.
                    Else
                        eatFood X, Y, .NextX, .NextY
                    End If
                End If
                If Not (formLife.mChkNoAntEatCarion.Checked Xor breakRules) Then 'Scavenge organic matter in new location before moving there...
                    scavenge X, Y, inGridX(.NextX - X), inGridY(.NextY - Y)
                End If
                Grid(.NextX, .NextY) = Grid(X, Y)
                RemoveItemXY X, Y
                AddItemXY .NextX, .NextY
                .Alive = False
                Grid(.NextX, .NextY).Alive = True
                X = inGridX(.NextX)
                Y = inGridY(.NextY)
            Else
                .Orientation = .Orientation + Rnd - 0.5 'Randomly adjust facing direction.
                'turn stored food into raw energy...
                digestionEnergy = (.redEnergy + .greenEnergy + .blueEnergy) / 8 + 0.01
                If .redEnergy > digestionEnergy Then
                    .Energy = .Energy + digestionEnergy
                    .redEnergy = .redEnergy - digestionEnergy
                End If
                If .greenEnergy > digestionEnergy Then
                    .Energy = .Energy + digestionEnergy
                    .greenEnergy = .greenEnergy - digestionEnergy
                End If
                If .blueEnergy > digestionEnergy Then
                    .Energy = .Energy + digestionEnergy
                    .blueEnergy = .blueEnergy - digestionEnergy
                End If
            End If
            If Not clearAhead Then
                'Something was in the way of further motion...
                .Velocity = 0
            End If
        End If
        If .Velocity < .Speed Then
            .Velocity = .Velocity + .Speed * 0.3
            .Energy = .Energy - .Speed * 0.4 'Some loss in energy conversion.
        End If
    End With
End Sub

Sub mateItemXY(X As Single, Y As Single)
    Dim R As Single
    Dim XX As Single
    Dim YY As Single
    Dim xXX As Single
    Dim yYY As Single
    Dim mayMateWithAnt As Boolean
    Dim mayMateWithPlant As Boolean
    R = Rnd * 4 - 2
    With Grid(X, Y)
        If .MatingAge + Rnd * 4 > .Age Then Exit Sub
        If .EnergyToMate > .Energy Then Exit Sub
        If .Speed > 0 Then
            mayMateWithAnt = Not (formLife.mChkNoMateAntAnt.Checked Xor breakRules)
            mayMateWithPlant = Not (formLife.mChkNoMateAntPlant.Checked Xor breakRules)
        Else
            mayMateWithAnt = Not (formLife.mChkNoMatePlantAnt.Checked Xor breakRules)
            mayMateWithPlant = Not (formLife.mChkNoMatePlantPlant.Checked Xor breakRules)
        End If
        If .Energy < 0.2 Then Exit Sub 'Absolute minimum energy to mate.
        xXX = X
        yYY = Y
        XX = X
        YY = Y
        If findEmptyItemAhead(xXX, yYY) Then
            If GetNearestLiveNeighborXY(XX, YY, False, True, mayMateWithAnt, mayMateWithPlant) Then
                Grid(xXX, yYY) = Grid(XX, YY)
                
                .Energy = .Energy / 2
                .redEnergy = .redEnergy / 2
                .greenEnergy = .greenEnergy / 2
                .blueEnergy = .blueEnergy / 2
                
                Grid(XX, YY).Energy = Grid(XX, YY).Energy / 2
                Grid(XX, YY).redEnergy = Grid(XX, YY).redEnergy / 2
                Grid(XX, YY).greenEnergy = Grid(XX, YY).greenEnergy / 2
                Grid(XX, YY).blueEnergy = Grid(XX, YY).blueEnergy / 2
                
                Grid(xXX, yYY).Energy = .Energy + Grid(XX, YY).Energy
                Grid(xXX, yYY).redEnergy = .redEnergy + Grid(XX, YY).redEnergy
                Grid(xXX, yYY).greenEnergy = .greenEnergy + Grid(XX, YY).greenEnergy
                Grid(xXX, yYY).blueEnergy = .blueEnergy + Grid(XX, YY).blueEnergy
                
                Grid(xXX, yYY).Chromosome = breedGenes(.Chromosome, Grid(XX, YY).Chromosome)
                
                If formLife.mChkEnableSounds.Checked Then
                    mateSound
                End If
                
                reBirthItemXY xXX, yYY
                Grid(XX, YY).Orientation = .Orientation + R
                .Orientation = .Orientation - R
                If (Rnd > 0.95) Or (formLife.mChkFastEvolve.Checked Xor breakRules) Then
                    mutateItemXY xXX, yYY
                End If
            End If
        End If
    End With
End Sub

Sub cloneItemXY(X As Single, Y As Single)
    Dim R As Single
    Dim XX As Single
    Dim YY As Single
    R = Rnd * 4 - 2
    'On Error Resume Next 'To avoid errors while scrolling the display
    With Grid(X, Y)
        If .minMates > 0 Then Exit Sub
        If .CloneCountDown > 0 Then
            .CloneCountDown = .CloneCountDown - 1
            Exit Sub
        End If
        If .Speed > 0 Then
            If (formLife.mChkNoCloneAnt.Checked Xor breakRules) Then Exit Sub
        Else
            If (formLife.mChkNoClonePlant.Checked Xor breakRules) Then Exit Sub
        End If
        If .CloningAge + Rnd * 4 > .Age Then Exit Sub
        If .EnergyToClone > .Energy Then Exit Sub
        If .Energy < 0.2 Then Exit Sub 'Absolute minimum energy to clone.
        XX = X
        YY = Y
        If findEmptyItemAhead(XX, YY) Then
        
            If formLife.mChkEnableSounds.Checked Then
                cloneSound
            End If
        
            .Energy = .Energy / 2
            .redEnergy = .redEnergy / 2
            .greenEnergy = .greenEnergy / 2
            .blueEnergy = .blueEnergy / 2
            Grid(XX, YY) = Grid(X, Y)
            reBirthItemXY XX, YY
            reBirthItemXY X, Y
            Grid(XX, YY).Orientation = .Orientation + R
            .Orientation = .Orientation - R
            If (Rnd > 0.95) Or (formLife.mChkFastEvolve.Checked Xor breakRules) Then
                mutateItemXY XX, YY
            End If
        End If
    End With
    'On Error GoTo 0
End Sub

Sub mutateItemXY(X As Single, Y As Single)
    Dim L As Long
    Dim b As Integer
    Dim M As Integer
    Dim t() As Byte
    Dim tS As String
    M = 1
    With Grid(X, Y)
        If Not .Alive Then Exit Sub
        If Rnd > 0.995 Then
            .Chromosome = .Chromosome + Chr(Int(Rnd * 256)) 'Gross mutation
        End If
        If Len(.Chromosome) < 20 Then
            If Rnd > 0.9 Then  'insert duplicate genes
                L = Int(Rnd * Len(.Chromosome)) + 1
                .Chromosome = Left(.Chromosome, L - 1) + Mid(.Chromosome, L, 4) + Mid(.Chromosome, L)
            End If
        Else
            If Rnd > 0.99 Then 'insert duplicate gene
                L = Int(Rnd * Len(.Chromosome)) + 1
                .Chromosome = Left(.Chromosome, L - 1) + Mid(.Chromosome, L, 2) + Mid(.Chromosome, L)
            End If
        End If
        If .Chromosome = "" Then
            If (formLife.mChkQuickStart.Checked Xor breakRules) Then
                'start with some random genes...
                .Chromosome = Chr(Int(Rnd * 255)) + Chr(Int(Rnd * 255)) + Chr(Int(Rnd * 255)) + Chr(Int(Rnd * 255))
                .Chromosome = Replace(.Chromosome, Chr(10), Chr(0)) ' Make sure initial random genes don't produce an animal.
            Else
                .Chromosome = Chr(0) + Chr(0) + Chr(0) + Chr(0) 'Make room for genes to start evolving
            End If
        End If
        L = Int(Rnd * Len(.Chromosome)) + 1
        If Mid(.Chromosome, L, 2) = Chr(0) + Chr(0) Then M = M + 1
        If M * Rnd > 0.11 Then 'Mutate a random gene
            L = Int(Rnd * Len(.Chromosome)) + 1
            t() = Mid(.Chromosome, L, 1)
            b = t(0)
            b = b + Int(Rnd * 5) - 2
            If b < 0 Then b = 0
            If b > 255 Then b = 255
            If Rnd > 0.1 Then
                t(0) = b
            Else 'swap after mutation
                t(0) = t(1)
                t(1) = b
            End If
            tS = t()
            .Chromosome = Left(.Chromosome, L - 1) + tS + Mid(.Chromosome, L + 1)
        End If
        If M * Rnd > 0.11 Then 'Mutate a random gene with bias toward earlier genes
            L = Int(Rnd * Len(.Chromosome)) + 1
            L = Int(Rnd * L) + 1
            t() = Mid(.Chromosome, L, 1)
            b = t(0)
            b = b + Int(Rnd * 5) - 2
            If b < 0 Then b = 0
            If b > 255 Then b = 255
            If Rnd > 0.1 Then
                t(0) = b
            Else 'swap after mutation
                t(0) = t(1)
                t(1) = b
            End If
            tS = t()
            .Chromosome = Left(.Chromosome, L - 1) + tS + Mid(.Chromosome, L + 1)
        End If
        If Rnd > 0.99 Then 'remove gene
            L = Int(Rnd * Len(.Chromosome)) + 1
            .Chromosome = Left(.Chromosome, L - 1) + Mid(.Chromosome, L + 3)
        End If
        .Energy = .Energy - 0.0001 * Len(.Chromosome) 'Energy expended for gene reproduction.
    End With
End Sub

Function findEmptyItemAhead(ByRef X As Single, ByRef Y As Single)
    Dim XX As Single
    Dim YY As Single
    Dim i As Integer
    'On Error Resume Next 'To avoid errors while scrolling the display
    With Grid(X, Y)
        i = 1
        Do
            Do
                XX = inGridX(X + Cos(.Orientation) * i)
                YY = inGridY(Y + Sin(.Orientation) * i)
                
                Call GridPos(XX, YY)
                
                i = i + 1
            Loop Until (i > .Width + .length + 3) Or ((Not Grid(XX, YY).Alive) And (Grid(X, Y).RGB <> Grid(XX, YY).RGB))
            'inner loop is intended to simulate toxins left behind by decaying organic matter.
        Loop Until (i > .Width * 2 + .length + 4) Or Not Grid(XX, YY).Alive
    End With
    If Not Grid(XX, YY).Alive Then
        X = XX
        Y = YY
        findEmptyItemAhead = True
    Else
        findEmptyItemAhead = False
    End If
    'On Error GoTo 0
End Function
Sub ageItemXY(X As Single, Y As Single)
    Dim foodEnergy As Single
    With Grid(X, Y)
        If Not .Alive Then Exit Sub
        If (.Age > .CertainDeathAge) And Not (formLife.mChkNoFatalAge.Checked Xor breakRules) Then
            'Death by old age.
            killItemXY X, Y
            Exit Sub
        End If
        .Age = .Age + 1
        If .Energy <= 0 Then ' Inefficient attempt not to starve.
            foodEnergy = .redEnergy + .greenEnergy + .blueEnergy / 3
            If foodEnergy > 0 Then
                .redEnergy = .redEnergy - foodEnergy
                .greenEnergy = .greenEnergy - foodEnergy
                .blueEnergy = .blueEnergy - foodEnergy
                'Yes, I know I subtracted the same energy three times.  This is meant to be wasteful.
                .Energy = .Energy + foodEnergy
            End If
        End If
        If (.Energy <= 0.1) And Not (formLife.mChkNoFatalHunger.Checked Xor breakRules) Then
            ' Starvation causes tissue destruction...
            biteItemXY X, Y
        End If
        If .Health < 0 Then
            killItemXY X, Y 'Starve to death.
            Exit Sub
        End If
    End With
End Sub

Sub scavenge(X As Single, Y As Single, Optional relativeX = 0, Optional relativeY = 0)
    Dim XX As Single
    Dim YY As Single
    
    'Pick a location near (x,y)...
    If (relativeX = 0) And (relativeY = 0) Then
        XX = X + relativeX
        YY = Y + relativeY
    Else
        XX = inGridX(X + Int(Rnd * 3) - 1)
        YY = inGridY(Y + Int(Rnd * 3) - 1)
    End If
    
    If (XX = X) And (YY = Y) Then 'One out of nine times, try again and allow a small gap...
        If Rnd < 0.5 Then
            XX = inGridX(X + Int(Rnd * 5) - 2)
        Else
            YY = inGridY(Y + Int(Rnd * 5) - 2)
        End If
    End If
    
    If (XX = X) And (YY = Y) Then Exit Sub 'Give up for now.
    
    Call GridPos(XX, YY)
    
    With Grid(XX, YY) 'Found organic matter to eat.
        Select Case Int(Rnd * 4) 'Don't eat it all at once.
        Case 0
             Grid(X, Y).Energy = Grid(X, Y).Energy + .Energy / 2 'Waste part of the scavenved energy.
             .Energy = 0
        Case 1
             Grid(X, Y).redEnergy = Grid(X, Y).redEnergy + .redEnergy / 2
             .redEnergy = 0
        Case 2
             Grid(X, Y).redEnergy = Grid(X, Y).redEnergy + .redEnergy / 2
             .redEnergy = 0
        Case 3
             Grid(X, Y).blueEnergy = Grid(X, Y).blueEnergy + .blueEnergy / 2
             .blueEnergy = 0
        End Select
    End With
End Sub

Sub eatFood(X As Single, Y As Single, foodX As Single, foodY As Single)
    With Grid(X, Y)
        .Energy = .Energy + Grid(foodX, foodY).Energy
        .redEnergy = .redEnergy + Grid(foodX, foodY).redEnergy
        .greenEnergy = .greenEnergy + Grid(foodX, foodY).greenEnergy
        .blueEnergy = .blueEnergy + Grid(foodX, foodY).blueEnergy
        Grid(foodX, foodY).Energy = 0
        Grid(foodX, foodY).redEnergy = 0
        Grid(foodX, foodY).greenEnergy = 0
        Grid(foodX, foodY).blueEnergy = 0
    End With
End Sub

Sub feedItemXY(X As Single, Y As Single)
    Dim XX As Single
    Dim YY As Single
    Dim redEnergy As Single
    Dim greenEnergy As Single
    Dim blueEnergy As Single
    Dim whiteEnergy As Single
    Dim onlyEatFood As Boolean
    Dim mayEatAnt As Boolean
    Dim mayEatPlant As Boolean
    XX = X
    YY = Y
    mayEatAnt = True
    mayEatPlant = True
    With Grid(X, Y)
        If .Speed > 0 Then 'Animal...
            mayEatAnt = Not (formLife.mChkNoAntEatAnt.Checked Xor breakRules)
            mayEatPlant = Not (formLife.mChkNoAntEatPlant.Checked Xor breakRules)
            If Not (formLife.mChkNoAntEatCarion.Checked Xor breakRules) Then
                If Rnd > 0.93 Then scavenge X, Y
            End If
        Else 'Plant...
            mayEatAnt = Not (formLife.mChkNoPlantEatAnt.Checked Xor breakRules)
            mayEatPlant = Not (formLife.mChkNoPlantEatPlant.Checked Xor breakRules)
            If Not (formLife.mChkNoPlantEatCarion.Checked Xor breakRules) Then
                If Rnd > 0.93 Then scavenge X, Y
            End If
        End If
        If .Energy < .Hungry Then
            onlyEatFood = (.Energy * 10 > .Hungry)
            .Energy = .Energy - 0.01 'Energy expended looking for food.
            If GetNearestLiveNeighborXY(XX, YY, onlyEatFood, False, mayEatAnt, mayEatPlant) Then
                If biteItemXY(XX, YY, .Width) Then
                    'take a bite out of grid(XX,YY).
                    redEnergy = Grid(XX, YY).redEnergy * .foodRed / 255 + 0.1
                    greenEnergy = Grid(XX, YY).greenEnergy * .foodGreen / 255 + 0.1
                    blueEnergy = Grid(XX, YY).blueEnergy * .foodBlue / 255 + 0.1
                    whiteEnergy = Grid(XX, YY).Energy
                    If redEnergy + greenEnergy + blueEnergy + whiteEnergy > .Hungry * 2 Then
                        whiteEnergy = whiteEnergy / 2 + 0.3
                    End If
                    If whiteEnergy < Grid(XX, YY).Energy Then
                        redEnergy = redEnergy / 2
                        greenEnergy = greenEnergy / 2
                        blueEnergy = blueEnergy / 2
                    Else
                        whiteEnergy = Grid(XX, YY).Energy
                        redEnergy = Grid(XX, YY).redEnergy
                        greenEnergy = Grid(XX, YY).greenEnergy
                        blueEnergy = Grid(XX, YY).blueEnergy
                    End If
                    .redEnergy = .redEnergy + redEnergy
                    .greenEnergy = .greenEnergy + greenEnergy
                    .blueEnergy = .blueEnergy + blueEnergy
                    .Energy = .Energy + whiteEnergy
                    Grid(XX, YY).redEnergy = Grid(XX, YY).redEnergy - redEnergy
                    Grid(XX, YY).greenEnergy = Grid(XX, YY).greenEnergy - greenEnergy
                    Grid(XX, YY).blueEnergy = Grid(XX, YY).blueEnergy - blueEnergy
                    Grid(XX, YY).Energy = Grid(XX, YY).Energy - whiteEnergy
                Else
                    'Eat grid(XX,YY) whole.
                    .Energy = .Energy + Grid(XX, YY).Energy
                    .redEnergy = .redEnergy + Grid(XX, YY).redEnergy
                    .greenEnergy = .greenEnergy + Grid(XX, YY).greenEnergy
                    .blueEnergy = .blueEnergy + Grid(XX, YY).blueEnergy
                    'Note... negative levels of energy in food simulate toxins.
                    Grid(XX, YY).Energy = 0
                    Grid(XX, YY).redEnergy = 0
                    Grid(XX, YY).greenEnergy = 0
                    Grid(XX, YY).blueEnergy = 0
                    killItemXY XX, YY 'death by ingestion.
                End If
            End If
        End If
        If (.Speed <= 0) Or Not (formLife.mChkNoAntEatDirt.Checked Xor breakRules) Then
            If .Age * 3 <= .CertainDeathAge + 77 Then 'No ambient feeding in old-age.
                .Energy = .Energy + 0.005 'Individual ambient energy absorption.
                If .Speed <= 0 Then 'Extra ration for young plants (root feeding)...
                    .Energy = .Energy + 0.007 / (Abs(.Age) + 1)
                End If
                If lifeCount > 0 Then .Energy = .Energy + 0.055 / Log(lifeCount + 1) 'Distributed ambient energy absorption.
            End If
        End If
        If .Age + .Age > .CertainDeathAge + 17 Then
            .Energy = .Energy - 0.01 'Energy drain in old-age, simulates organ system failure.
        End If
    End With
End Sub

Function squareStepSizeXY(X As Single, Y As Single)
    Dim XX As Single
    Dim YY As Single
    With Grid(X, Y)
        XX = .NextX - X
        YY = .NextY - Y
        squareStepSizeXY = XX * XX + YY * YY
    End With
End Function

Sub resetGrid()
    Dim X As Integer
    Dim Y As Integer
    Dim baby As Creature
    For X = 0 To GridSizeX - 1
        For Y = 0 To GridSizeY - 1
            Grid(X, Y) = baby
            'Start with some "organic matter" in the environment...
            Grid(X, Y).Energy = (Rnd + 1) / 400 'raw energy. (Calories)
            Grid(X, Y).redEnergy = 0.0005 'Needs to be digested before use.
            Grid(X, Y).greenEnergy = 0.0005 'Needs to be digested before use.
            Grid(X, Y).blueEnergy = 0.0005 'Needs to be digested before use.
        Next Y
    Next X
    ListXY = ""
    lifeCount = 0
    AntCount = 0
    PlantCount = 0
End Sub

Sub playNote(Optional note As Double = 0, Optional Octave As Double = 4, Optional beats As Double = 1, Optional beatsPerMinute As Double = 100)
    'Tempo conversions into approximate beats per minute where a quarter note note equals one beat...
    'Grave >=20
    'Adagio >=30
    'Largo >=40
    'Lento >=60
    'Andante >=66
    'Moderato >=76
    'Allegretto >=120
    'Allegro >=140
    'Vivace >=160
    'Presto >=180
    'Prestissimo >=200
    'Based on these approximations, I  have chosen 100 beats per minute as a standard tempo.
    Beep 16.35 * 2 ^ ((note / 12) + Octave), beats * 60000 / beatsPerMinute
End Sub

Function breakRules() As Boolean
    If formLife.mChkBreakRules.Checked Then
        breakRules = (Rnd < 0.01) 'break rules to be broken less than 1% of the time.
    Else
        breakRules = False
    End If
End Function
