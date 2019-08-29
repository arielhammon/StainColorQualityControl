# StainColorQualityControl

## Introduction
This file contains descriptions and code samples of a Visual Basic application I built for a cabinet manufacturer to help them maintain consistency in their stain colors with other component manufacturers. It maintained a library of hundreds of stain colors spanning multiple years. It imported readings taken from a spectrophotometer and compared them to a standard in order to keep production on track. It was important to allow for considerable variation as naturally occurs in products like wood. Therefore a standard was actually a collection of separate readings which met certain statistical requirements.

## My Role
I was the sole developer of this project and worked on it, off-and-on along with other manufacturing systems for the same company, for about five years or so. In this capacity, not only did I write the code for this project, I also identified the need, engineered the solution, gathered the necessary equipment (spectrophotometers), and implemented the solution.

## Contents
Sample Components
1. [Routine to Build Chart for Comparing Current Trial to Standard](#routine-to-build-chart-for-comparing-current-trial-to-standard)
2. [ColorCombo_Change Event Handler](#colorcombo_change-event-handler)
3. [DatesLB_KeyDown Event Handler](#dateslb_keydown-event-handler)
4. [Create New Standard](#create-new-standard)
5. [Reading Class](#reading-class) Instances of this class represent a spectrophotometer reading.
6. [ReadingSet Class](#readingset-class) Instances of this class represent a collection of spectrophotometer readings such as a standard or a collection of readings used to create a standard.

## Routine to Build Chart for Comparing Current Trial to Standard
This routine retrieves a trial reading from the spectrophotometer (a collection of reflectance percentages vs. wavelengths of light). It then retrieves the standard from a database of historical data and displays a chart which compares the two.

    Public Sub RebuildChart()
        Dim k As Integer, i As Integer, npts As Integer
        Dim c As Range
        Dim spm As Series 'spectrophotometric
        Dim cs As ColorSet
        Dim r As Reading
        Dim crs As New ReadingSet 'chart reading set
        Dim RS As ReadingSet
        Dim SCIb As Boolean, Extb As Boolean
        Dim LowLambda As Long, HighLambda As Long

        On Error GoTo Problem
        ChartPage.HideTrialMeasureBox
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False

        SCIb = SCIOB.Value
        Extb = ExtendedTB.Value
        If ShowCurrentCB.Value Then
            With TrialSet
                .ColorName = "Trial"
                .ColorDate = "Today"
                .GetTrialReadings
                If SCIb Then .SCMode = SCI Else .SCMode = SCE
                Set RS = .SCModeSet
                For k = 1 To RS.Readings.Count
                    crs.Readings.Add RS.Readings(k)
                Next k
            End With
        End If
        For i = 1 To ColorSets.Count
            Set cs = ColorSets(i)
            With cs
                If SCIb Then .SCMode = SCI Else .SCMode = SCE
                Set RS = .SCModeSet
                For k = 1 To RS.Readings.Count
                    crs.Readings.Add RS.Readings(k)
                Next k
            End With
        Next i
        If crs.Readings.Count = 0 Then Exit Sub
        crs.GetLambdas LowLambda, HighLambda
        If Extb Then
            If LowLambda >= 400 Or HighLambda <= 700 Then _
                MsgBox "Extended wavelengths are not available for this color. To view them, select that option in OnColor and re-export."
        Else
            If LowLambda < 400 Or HighLambda > 700 Then
                LowLambda = 400
                HighLambda = 700
            End If
        End If
        npts = ((HighLambda - LowLambda) \ 10) + 1
        With ChartDataSheet
            .Cells.Clear
            For i = 0 To npts - 1
                .Cells(i + 4).Value = (LowLambda + i * 10) & "nm"
            Next i
            Set c = .Range("A2")
            For i = 1 To crs.Readings.Count
                Set r = crs.Readings(i)
                r.PlacePartHeaderInRange c
                r.PlacePartPtsInRange c.offset(0, 3), (LowLambda - r.StartLambda) \ 10
                Set c = c.offset(1, 0)
            Next i
        End With
        With ChartPage
            .SetSourceData Source:=ChartDataSheet.Range("A2").CurrentRegion, PlotBy:=xlRows
            With .SeriesCollection
                For i = 1 To crs.Readings.Count
                    Set spm = .Item(i)
                    Set r = crs.Readings(i)
                    r.SetSeries Line:=spm
                Next i
            End With
        End With
        Set crs = Nothing
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Exit Sub

    Problem:
        ResetApp
        Err.Raise Err.Number
    End Sub

## ColorCombo_Change Event Handler
This event handler for a ComboBox searched the database of historical data for color readings which matched the selected color and then populated a ListBox of matching colors grouped by date.

    Private Sub ColorCombo_Change()
        Dim ColorName As String, ColorDate As String, PrevColorDate As String, temp As String
        Dim i As Long, j As Long, k As Long, n As Long, m As Long
        Dim l As Range
        Dim cs As ColorSet, cs2 As ColorSet
        Dim AlreadyIn As Boolean

        If Not SkipEvents Then
        If Not ColorCombo.MatchFound Then
            SkipEvents = True
            DatesLB.Clear
            SkipEvents = False
        Else
            SkipEvents = True
            If ClearWithNewColorCB.Value And ClearColorSet Then
                For j = ColorSets.Count To 1 Step -1
                    ColorSets.Remove (j)
                Next j
            End If
            With DatesLB
                .Clear
                n = DispColorSets.Count
                For k = n To 1 Step -1
                    DispColorSets.Remove (k)
                Next k
                ColorName = ColorCombo.Text
                Dim block As Range
                Set block = GetColorRange(ColorName)
                n = block.Rows.Count
                PrevColorDate = ""
                m = 0
                For k = n To 1 Step -1 'so that dates appear most recent first
                    Set l = block.Rows(k)
                    ColorDate = l.Cells(2).Value
                    If ColorDate <> PrevColorDate Then
                        .AddItem ColorDate
                        AlreadyIn = False
                        For i = 1 To ColorSets.Count
                            Set cs = ColorSets(i)
                            If cs.ColorName = ColorName Then
                                If cs.FormattedDate = ColorDate Then
                                    AlreadyIn = True
                                    m = m + 1
                                    Exit For
                                End If
                            End If
                        Next i
                        If Not AlreadyIn Then
                            Set cs = New ColorSet
                            With cs
                                .ColorName = ColorName
                                .ColorDate = l.Cells(2).Value2
                                .FormattedDate = ColorDate
                            End With
                        End If
                        DispColorSets.Add cs
                    End If
                    PrevColorDate = ColorDate
                Next k
                If m > 1 Then
                    MultiDatesCB.Value = True
                    DatesLB.MultiSelect = fmMultiSelectMulti
                End If
                For j = 1 To DispColorSets.Count
                    Set cs = DispColorSets(j)
                    .Selected(j - 1) = cs.Selected
                Next j
            End With
            SkipEvents = False
            DatesLB_Change
            LabelChartForm.ColorTB = ColorName
        End If
        End If
    End Sub

## DatesLB_KeyDown Event Handler
This event handler for a ListBox allowed the user to delete, insert, or rename color standards.

    Private Sub DatesLB_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        Dim cs As ColorSet
        Dim k As Integer
        If KeyCode = 46 And Shift = 0 Then 'the user has pressed the delete key
            If Not User.Standards Then
                DoLoginQuestion
                Exit Sub
            End If
            Dim result As VbMsgBoxResult
            Dim style As VbMsgBoxStyle
            style = vbCritical + vbYesNo + vbDefaultButton2
            result = MsgBox("Are you sure you want to delete these colorsets?", style, "WARNING")
            If result = vbYes Then
                For k = 1 To DispColorSets.Count
                    Set cs = DispColorSets(k)
                    If cs.Selected = True Then
                        cs.DelStdHistory True
                    End If
                Next k
            End If
            Repopulate ColorCombo.Text, True
        End If
        If KeyCode = 45 And Shift = 0 Then 'the user has pressed the insert key
            Dim tempcs As ColorSet
            For k = 1 To DispColorSets.Count
                Set cs = DispColorSets(k)
                If cs.Selected = True Then
                    Set tempcs = cs.Copy
                    tempcs.ColorDate = tempcs.FormattedDate & " Copy"
                    tempcs.FormattedDate = tempcs.ColorDate
                    tempcs.PlaceStdHistory
                    Repopulate ColorCombo.Text, False
                    Exit For 'only copy the first selected colorset
                End If
            Next k
        End If
        If KeyCode = 82 And Shift = 2 Then 'the user has pressed Ctrl-R to rename the standard
            If Not User.Standards Then
                DoLoginQuestion
                Exit Sub
            End If
            For k = 1 To DispColorSets.Count
                Set cs = DispColorSets(k)
                If cs.Selected = True Then
                    Dim t As String
                    t = InputBox("Please enter the new name.", "Rename ColorSet", cs.FormattedDate)
                    If t <> "" Then
                        Dim rng As Range, l As Range
                        Set rng = HistData.GetColorDateRange(cs.ColorName, cs.ColorDate)
                        For Each l In rng.Rows
                            l.Cells(2).Value = t
                        Next l
                        cs.ColorDate = rng.Cells(2).Value2
                        cs.FormattedDate = rng.Cells(2).Value
                        HistData.IsSorted = False
                        Repopulate ColorCombo.Text, False
                    End If
                    Exit For 'only rename the first selected colorset
                End If
            Next k
        End If
    End Sub

## Create New Standard
The following two routines are examples of a collection of routines that helped the user create a new color standard. The first routine below takes a collection of many sample spectrophotometer readings and removes the lightest and darkest readings. The second routine below takes values from a Form containing sliders where the user could fine tune the standard by weighting individual readings; it would then normalize the weights for each reading in the standard.

    Private Sub UnflagUnweightExtremes()
        Dim white As New Reading
        Dim black As New Reading
        white.SetDimSameAs AllCSTrim.SCIs.Readings(1)
        black.SetDimSameAs white

        Dim k As Integer, m As Integer
        Dim npts As Integer
        npts = white.Pts.Count
        For k = 1 To npts
            white.Pts(k).d = 200 'beaming white
            black.Pts(k).d = 0.001 'ultra black, can't be zero
        Next k

        'remove extremes:
        Dim Nearest As Collection
        Dim r As Reading
        Dim RS As ReadingSet, RS2 As ReadingSet
        If StdCS.SCMode = SCI Then
            Set RS = AllCSTrim.SCIs
            Set RS2 = AllCSTrim.SCEs
        Else
            Set RS = AllCSTrim.SCEs
            Set RS2 = AllCSTrim.SCIs
        End If

        m = LightestSB.Value
        If m > 0 Then
            Set Nearest = RS.IndicesOfNearest(white, m)
            For k = 1 To m
                Set r = RS.Readings(Nearest(k).i)
                r.Flag = False
                r.weight = 0
                Set r = RS2.Readings(Nearest(k).i)
                r.Flag = False
                r.weight = 0
            Next k
        End If

        m = DarkestSB.Value
        If m > 0 Then
            Set Nearest = RS.IndicesOfNearest(black, m)
            For k = 1 To m
                Set r = RS.Readings(Nearest(k).i)
                r.Flag = False
                r.weight = 0
                Set r = RS2.Readings(Nearest(k).i)
                r.Flag = False
                r.weight = 0
            Next k
        End If
    End Sub
    
    Private Sub NormalizeWeights()
        Dim k As Integer
        Dim n As Integer
        Dim r As Reading, r2 As Reading
        Dim RS As ReadingSet, RS2 As ReadingSet

        If StdCS.SCMode = SCI Then
            Set RS = AllCSTrim.SCIs
            Set RS2 = AllCSTrim.SCEs
        Else
            Set RS = AllCSTrim.SCEs
            Set RS2 = AllCSTrim.SCIs
        End If

        n = StdCS.SCIs.Readings.Count
        Dim tw As Double
        tw = 0
        For k = 1 To n
            Set r = RS.Readings(k)
            tw = tw + r.weight
        Next k
        For k = 1 To n
            Set r = RS.Readings(k)
            Set r2 = RS2.Readings(k)
            r.weight = r.weight / tw
            r2.weight = r.weight
        Next k
    End Sub

## Reading Class
Instances of this class represent a spectrophotometer reading. The class contains a helper function to compute the root mean squares (RMS) in reference to a "nearby" standard reading. Because a standard was a collection of readings, it was necessary to transform each reading in a trial to its nearest reading in the standard before computing any deviation. Therefore, this class also contains several methods to transform a reading to a position where it could best be compared to the standard. Dealing with all of this variation is what set this method of QC'ing parts ahead of other methods.

    Option Explicit

    Public ColorName As String
    Public ColorDate As String
    Public Reading As String
    Public desc As String
    Public Mode As String
    Public StartLambda As Double
    Public DeltaLambda As Double
    Public Pts As Collection 'of double wrapper
    Public Flag As Boolean 'used to mark this reading for removal from the current standard
    Public Format As SeriesFormat
    Public PtFormat As SeriesFormat
    Public LineSeries As Series
    Public weight As Double
    Public Index As Integer

    Public Function Copy() As Reading
        Dim r As New Reading
        With r
            .ColorName = ColorName
            .ColorDate = ColorDate
            .Reading = Reading
            .desc = desc
            .Mode = Mode
            .StartLambda = StartLambda
            .DeltaLambda = DeltaLambda
            Dim k As Integer
            Dim d As Dbl
            For k = 1 To Pts.Count
                Set d = New Dbl
                d.d = Pts(k).d
                .Pts.Add d
            Next k
            .Flag = Flag
            Set .Format = Format.Copy
            Set .PtFormat = PtFormat.Copy
            Set .LineSeries = LineSeries
            .weight = weight
            .Index = Index
        End With
        Set Copy = r
    End Function

    Public Sub TrimPts(ByVal LowLambda As Long, ByVal HighLambda As Long)
        Dim ll As Long, hl As Long
        ll = StartLambda
        hl = ll + (Pts.Count - 1) * DeltaLambda
        Dim k As Integer
        For k = 1 To (LowLambda - ll) \ DeltaLambda
            Pts.Remove (1)
            StartLambda = StartLambda + DeltaLambda
        Next k
        Dim n As Integer
        n = Pts.Count
        For k = 1 To (hl - HighLambda) \ DeltaLambda
            Pts.Remove (n)
            n = n - 1
        Next k
    End Sub

    Public Sub ClearPts()
        Set Pts = Nothing
        Set Pts = New Collection
    End Sub
    Public Sub SetSeries(Line As Series)
        Set LineSeries = Line
        With LineSeries
            .Smooth = True
            .MarkerStyle = Format.MarkerStyle
            If .MarkerStyle <> xlMarkerStyleNone Then
                .MarkerForegroundColor = Format.MarkerForegroundColor
                .MarkerBackgroundColor = Format.MarkerBackgroundColor
                .MarkerSize = Format.MarkerSize
            End If
            With .Border
                .Color = Format.BorderColor
                .weight = Format.BorderWeight
            End With
        End With
    End Sub

    Public Sub FillHeaderFromRange(ByVal cell As Range)
        'fills the header from a single row starting at "cell"
        ColorName = cell.offset(0, 0).Value2
        ColorDate = cell.offset(0, 1).Value2
        Reading = cell.offset(0, 2).Value2
        desc = cell.offset(0, 3).Value2
        Mode = cell.offset(0, 4).Value2
    End Sub

    Public Sub FillPartHeaderFromRange(ByVal cell As Range)
        'fills the header from a single row starting at "cell"
        Reading = cell.offset(0, 0).Value2
        desc = cell.offset(0, 1).Value2
        Mode = cell.offset(0, 2).Value2
    End Sub

    Public Sub PlaceHeaderInRange(ByVal cell As Range)
        cell.offset(0, 0).Value = ColorName
        cell.offset(0, 1).Value = ColorDate
        cell.offset(0, 2).Value = Reading
        cell.offset(0, 3).Value = desc
        cell.offset(0, 4).Value = Mode
    End Sub

    Public Sub PlacePartHeaderInRange(ByVal cell As Range)
        cell.offset(0, 0).Value = Reading
        cell.offset(0, 1).Value = desc
        cell.offset(0, 2).Value = Mode
    End Sub

    Public Sub FillPtsFromRange(ByVal cell As Range, ByVal n As Integer)
        'fills the collection of points from a single row starting at "cell", n cells wide.
        Dim i As Integer
        Dim d As Dbl
        Set Pts = New Collection
        For i = 1 To n
            Set d = New Dbl
            d.d = cell.Value
            Pts.Add d
            Set cell = cell.offset(0, 1)
        Next i
    End Sub

    Public Sub PlacePtsInRange(ByVal cell As Range)
            'fills the collection of points from a single row starting at "cell", n cells wide.
        Dim i As Integer
        For i = 1 To Pts.Count
            cell.Value = Pts(i).d
            Set cell = cell.offset(0, 1)
        Next i
    End Sub

    Public Sub PlacePartPtsInRange(ByVal cell As Range, ByVal nOffEnds As Integer)
        'fills the collection of points from a single row starting at "cell", n cells wide.
        Dim i As Integer
        For i = 1 + nOffEnds To Pts.Count - nOffEnds
            cell.Value = Pts(i).d
            Set cell = cell.offset(0, 1)
        Next i
    End Sub

    Public Function MeanDifference(ByVal Other As Reading) As Double
        Dim rv As Double
        rv = 0
        If Not Me.SameDim(Other) Then
            Err.Raise 2000, "Reading.MeanDifference", "Readings are not of the same dimension."
        Else
            Dim i As Integer
            Dim d1 As Double, d2 As Double
            For i = 1 To Pts.Count
                d1 = Pts(i).d
                d2 = Other.Pts(i).d
                rv = rv + (d2 - d1) / d1
            Next i
            rv = rv / Pts.Count
        End If
        VectorRMS = rv
    End Function

    Public Function RootMeanSquares(ByVal Other As Reading) As Double
        Dim rv As Double
        rv = 0
        If Not SameDim(Other) Then
            Err.Raise 2000, "Reading.RootMeanSquares", "Readings are not of the same dimension."
        Else
            Dim i As Integer
            Dim d1 As Double, d2 As Double, d As Double
            For i = 1 To Pts.Count
                d1 = Pts(i).d
                d2 = Other.Pts(i).d
                d = (d2 - d1) / d1
                rv = rv + d * d
            Next i
            rv = rv / Pts.Count
            rv = Sqr(rv)
        End If
        RootMeanSquares = rv
    End Function

    Public Sub TransformBy(ByVal FromReading As Reading, _
        ByVal Shift As Double, ByVal SFx As Double, ByVal PFy As Double)

        If Not Me.SameDim(FromReading) Then
            Me.SetDimSameAs FromReading
        End If

        Dim k As Integer
        For k = 1 To Pts.Count
            Pts(k).d = SFx * FromReading.Pts(k).d ^ PFy + Shift
        Next k
    End Sub

    Public Sub cTransform(ByVal FromReading As Reading, ByVal ToReading As Reading, ByVal cReading As Reading, _
        Optional ByRef Shift As Double, Optional ByRef SFx As Double, Optional ByRef PFy As Double, _
        Optional ByRef RMS As Double)

        If Not FromReading.SameDim(ToReading) Then
            Err.Raise 2000, "Transform", "Some readings are not the same dimension."
        End If
        If Not Me.SameDim(FromReading) Then
            Me.SetDimSameAs FromReading
        End If

        Dim tws As Worksheet
        Set tws = TransWB.Sheets("Fitting")
        Dim k As Integer, npts As Integer
        npts = Pts.Count
        Dim ShConst As Double, PFyConst As Double, cRMS As Double
        Transform FromReading, cReading, Shift, SFx, PFy, RMS
        'handle case SFx = 1
        PFyConst = (1 - PFy) / Log(SFx)
        'handle case SFx*PFy = 1
        ShConst = Shift / (1 - SFx * PFy)
        cRMS = RMS
        Dim d As Double
        With tws
            .Select
            Dim l As Range
            Set l = .Range("B2", "C2")
            For k = 1 To npts
                l(1).Value = FromReading.Pts(k).d
                l(2).Value = ToReading.Pts(k).d
                Set l = l.offset(1, 0)
            Next k

            .Range("K5").Value2 = ShConst
            .Range("M5").Value2 = PFyConst
            .Range("K2").Formula = "=K5*(1-L2*M2)"
            .Range("M2").Formula = "=IF(L2>0,1-M5*LOG(L2),2)"
            SolverLoad LoadArea:=.Range("SolverOnCurve")
            SolverSolve True
            Shift = .Range("K2").Value2
            SFx = .Range("L2").Value2
            PFy = .Range("M2").Value2

            Set l = .Range("B2", "E2")
            For k = 1 To npts
                d = l.Cells(3).Value2
                Pts(k).d = d
                l.Cells(4).Value2 = d
                Set l = l.offset(1, 0)
            Next k

            cRMS = (cRMS + 0.0025) * FromReading.RootMeanSquares(Me) / FromReading.RootMeanSquares(cReading)
            .Range("J5").Value2 = cRMS
            Dim errcount As Integer
            errcount = 0
    TryAgain:
    On Error GoTo Problem
            SolverLoad LoadArea:=.Range("SolverWithTarget")
            SolverSolve True

            RMS = .Range("J2").Value2
            Shift = .Range("K2").Value2
            SFx = .Range("L2").Value2
            PFy = .Range("M2").Value2
            Set l = .Range("D2")
            For k = 1 To Pts.Count
                Pts(k).d = l.Value2
                Set l = l.offset(1, 0)
            Next k
        End With
        Exit Sub
    Problem:
        If Err.Number = 13 Then 'type mismatch, most likely overflow from too large of a PFy value
            With tws
                'reset transform values
                .Range("K2").Value2 = (-1) ^ errcount * 0.05 * errcount
                .Range("L2").Value2 = 1 + (-1) ^ errcount * 0.05 * errcount
                .Range("M2").Value2 = 1 + (-1) ^ (errcount + 1) * 0.05 * errcount
                errcount = errcount + 1
                If errcount < 10 Then GoTo TryAgain
            End With
        End If
        Err.Raise Err.Number
    End Sub

    Public Sub Transform(ByVal FromReading As Reading, ByVal ToReading As Reading, _
        Optional ByRef Shift As Double, Optional ByRef SFx As Double, Optional ByRef PFy As Double, _
        Optional ByRef RMS As Double)

        If Not FromReading.SameDim(ToReading) Then
            Err.Raise 2000, "Transform", "Some readings are not the same dimension."
        End If
        If Not Me.SameDim(FromReading) Then
            Me.SetDimSameAs FromReading
        End If

        Dim tws As Worksheet
        Dim l As Range
        Dim k As Integer, npts As Integer

        npts = Pts.Count
        Set tws = TransWB.Sheets("Fitting")
        Set l = tws.Range("B2", "C2")
        For k = 1 To npts
            l(1).Value = FromReading.Pts(k).d
            l(2).Value = ToReading.Pts(k).d
            Set l = l.offset(1, 0)
        Next k
        SFx = 1
        PFy = 1
        Shift = 0
        With tws
            .Select
            .Range("K2").Value2 = Shift
            .Range("L2").Value2 = SFx
            .Range("M2").Value2 = PFy
            SolverLoad LoadArea:=.Range("SolverNoTarget")
            SolverSolve True
            RMS = .Range("J2").Value2
            Shift = .Range("K2").Value2
            SFx = .Range("L2").Value2
            PFy = .Range("M2").Value2
        End With
        Set l = tws.Range("D2")
        For k = 1 To Pts.Count
            Pts(k).d = l.Value2
            Set l = l.offset(1, 0)
        Next k
    End Sub

    Public Function SameDim(ByVal Other As Reading) As Boolean
        If Other.Pts.Count <> Pts.Count Or Other.StartLambda <> StartLambda Or Other.DeltaLambda <> DeltaLambda Then
            SameDim = False
        Else
            SameDim = True
        End If
    End Function

    Public Sub SetNumPts(ByVal n As Integer)
        Set Pts = Nothing
        Set Pts = New Collection
        Dim i As Integer
        Dim d As Dbl
        For i = 1 To n
            Set d = New Dbl
            Pts.Add d
        Next i
    End Sub

    Public Sub SetDimSameAs(ByVal Other As Reading)
        StartLambda = Other.StartLambda
        DeltaLambda = Other.DeltaLambda
        SetNumPts (Other.Pts.Count)
    End Sub

    Private Sub Class_Initialize()
        ColorName = ""
        ColorDate = ""
        Reading = ""
        desc = ""
        Mode = ""
        StartLambda = 360
        DeltaLambda = 10
        Flag = False
        weight = 1
        Index = 0
        Set Pts = New Collection
        Set Format = New SeriesFormat
        Set PtFormat = New SeriesFormat
    End Sub

## ReadingSet Class
Instances of this class represent a collection of spectrophotometer readings such as a standard or a collection of readings used to create a standard.
