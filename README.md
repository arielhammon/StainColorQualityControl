# StainColorQualityControl

## Introduction
This file contains descriptions and code samples of a Visual Basic application I built for a cabinet manufacturer to help them maintain consistency in their stain colors with other component manufacturers. It maintained a library of hundreds of stain colors spanning multiple years. It imported readings taken from a spectrophotometer and compared them to the standard in order to keep production on track. It was important to allow for considerable variation as naturally occurs in products like wood.

## My Role
I was the sole developer of this project and worked on it, off-and-on along with other manufacturing systems for the same company, for about five years or so. In this capacity, not only did I write the code for this project, I also identified the need, engineered the solution, gathered the necessary equipment (spectrophotometers), and implemented the solution.

## Contents
Sample Components
1. [Routine Building Chart to Compare Current Trial Reading to Standard](#routine-building-chart-to-compare-current-trial-reading-to-standard)
2. [ColorCombo_Change Event Handler](#colorcombo_change-event-handler)

## Routine Building Chart to Compare Current Trial Reading to Standard
This routine retrieves a trial reading from the spectrophotometer (a collection of reflectance percentages vs. wavelengths of light). It then retrieves the standard from a database of historic data and displays a chart which compares the two.

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

