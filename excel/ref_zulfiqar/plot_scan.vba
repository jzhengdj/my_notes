Sub Button1_Click()
    Dim strFile As String, strPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Const xlDelimited = 1
    Const xlTextQualifierNone = -4142
     
     'Change this path for your own file location:
    strPath = "D:\TiM100\TSD\17420111\"
    
    Application.DisplayAlerts = False
    'Delete all worksheets except the first
    Do While ThisWorkbook.Worksheets.Count > 1
        ThisWorkbook.Worksheets(2).Delete
    Loop
    
    Do While ThisWorkbook.Charts.Count > 0
        ThisWorkbook.Charts(1).Delete
    Loop
    Application.DisplayAlerts = True
    
     'this returns an empty string "" if the file cannot be found and will error if the folder is incorrect
    strFile = Dir(strPath & "*.txt")
    Do While strFile <> ""
         'add a new worksheet at the end of the macro workbook to paste into
        ThisWorkbook.Worksheets.Add After:=Sheets(Sheets.Count)
        Set currentSheet = ThisWorkbook.Worksheets(Sheets.Count)
        currentSheet.Name = Right(strFile, Len(strFile) - 25)
         
         'open the csv file and assign it to a variable so that we can easily reference it later
        Set file = fso.OpenTextFile(strPath + strFile, 1)
        StrData = file.ReadLine
        
        Row = 1
        Do Until file.AtEndOfStream
             'paste it
            currentSheet.Cells(Row, 1) = StrData
            StrData = file.ReadLine
            Row = Row + 1
        Loop
        
        With currentSheet
            .Columns("A:A").TextToColumns .Range("A1"), xlDelimited, xlDoubleQuote, True, , True, , False
        End With
        
        'Chart for 1.5m targets
        ThisWorkbook.Charts.Add
        ThisWorkbook.ActiveChart.HasTitle = True
        ThisWorkbook.ActiveChart.ChartTitle.Text = "Scan View Visualization Around Zero Degrees"
        ThisWorkbook.ActiveChart.ChartType = xlXYScatterLines
        ThisWorkbook.ActiveChart.SetSourceData Source:=currentSheet.Range("A:C"), PlotBy:=xlColumns
        ThisWorkbook.ActiveChart.Axes(xlCategory).MinimumScale = -1200
        ThisWorkbook.ActiveChart.Axes(xlCategory).MaximumScale = 1200
        ThisWorkbook.ActiveChart.Axes(xlValue).MaximumScale = 1600
        ThisWorkbook.ActiveChart.Location xlLocationAsObject, currentSheet.Name
        ThisWorkbook.ActiveChart.SeriesCollection(1).AxisGroup = 2  'set the axis to secondary for RSSI plot
        ThisWorkbook.ActiveChart.Legend.Position = xlLegendPositionBottom
        'highlight the search angle
        Dim idx1 As Integer
        idx1 = 97
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx1).MarkerSize = 10
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx1).MarkerStyle = 2
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx1).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx1).MarkerSize = 10
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx1).MarkerStyle = 2
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx1).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)

        'Chart for 50mm target
        ThisWorkbook.Charts.Add
        ThisWorkbook.ActiveChart.HasTitle = True
        ThisWorkbook.ActiveChart.ChartTitle.Text = "Scan View Visualization Around -90° degrees"
        ThisWorkbook.ActiveChart.ChartType = xlXYScatterLines
        ThisWorkbook.ActiveChart.SetSourceData Source:=currentSheet.Range("A:C"), PlotBy:=xlColumns
        ThisWorkbook.ActiveChart.Axes(xlCategory).MinimumScale = -10000
        ThisWorkbook.ActiveChart.Axes(xlCategory).MaximumScale = -8000
        ThisWorkbook.ActiveChart.Axes(xlValue).MaximumScale = 100
        ThisWorkbook.ActiveChart.Location xlLocationAsObject, currentSheet.Name
        ThisWorkbook.ActiveChart.SeriesCollection(1).AxisGroup = 2  'set the axis to secondary for RSSI plot
        ThisWorkbook.ActiveChart.Legend.Position = xlLegendPositionBottom
        'highlight the search angle
        Dim idx2 As Integer
        idx2 = 11
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx2).MarkerSize = 10
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx2).MarkerStyle = 2
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx2).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx2).MarkerSize = 10
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx2).MarkerStyle = 2
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx2).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)

        'Chart for 2.9m target
        ThisWorkbook.Charts.Add
        ThisWorkbook.ActiveChart.HasTitle = True
        ThisWorkbook.ActiveChart.ChartTitle.Text = "Scan View Visualization Around 18° degrees"
        ThisWorkbook.ActiveChart.ChartType = xlXYScatterLines
        ThisWorkbook.ActiveChart.SetSourceData Source:=currentSheet.Range("A:C"), PlotBy:=xlColumns
        ThisWorkbook.ActiveChart.Axes(xlCategory).MinimumScale = 1200
        ThisWorkbook.ActiveChart.Axes(xlCategory).MaximumScale = 2200
        ThisWorkbook.ActiveChart.Axes(xlValue).MaximumScale = 3000
        ThisWorkbook.ActiveChart.Location xlLocationAsObject, currentSheet.Name
        ThisWorkbook.ActiveChart.SeriesCollection(1).AxisGroup = 2  'set the axis to secondary for RSSI plot
        ThisWorkbook.ActiveChart.Legend.Position = xlLegendPositionBottom
        'highlight the search angle
        Dim idx3 As Integer
        idx3 = 119
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx3).MarkerSize = 10
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx3).MarkerStyle = 2
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx3).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx3).MarkerSize = 10
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx3).MarkerStyle = 2
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx3).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)

        'Chart for 1m target
        ThisWorkbook.Charts.Add
        ThisWorkbook.ActiveChart.HasTitle = True
        ThisWorkbook.ActiveChart.ChartTitle.Text = "Scan View Visualization Around 90° degrees"
        ThisWorkbook.ActiveChart.ChartType = xlXYScatterLines
        ThisWorkbook.ActiveChart.SetSourceData Source:=currentSheet.Range("A:C"), PlotBy:=xlColumns
        ThisWorkbook.ActiveChart.Axes(xlCategory).MinimumScale = 8000
        ThisWorkbook.ActiveChart.Axes(xlCategory).MaximumScale = 10000
        ThisWorkbook.ActiveChart.Axes(xlValue).MaximumScale = 1100
        ThisWorkbook.ActiveChart.Location xlLocationAsObject, currentSheet.Name
        ThisWorkbook.ActiveChart.SeriesCollection(1).AxisGroup = 2  'set the axis to secondary for RSSI plot
        ThisWorkbook.ActiveChart.Legend.Position = xlLegendPositionBottom
        'highlight the search angle
        Dim idx4 As Integer
        idx4 = 191
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx4).MarkerSize = 10
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx4).MarkerStyle = 2
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx4).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx4).MarkerSize = 10
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx4).MarkerStyle = 2
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx4).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)

        'Chart for 0.7m target
        ThisWorkbook.Charts.Add
        ThisWorkbook.ActiveChart.HasTitle = True
        ThisWorkbook.ActiveChart.ChartTitle.Text = "Scan View Visualization Around 70° degrees"
        ThisWorkbook.ActiveChart.ChartType = xlXYScatterLines
        ThisWorkbook.ActiveChart.SetSourceData Source:=currentSheet.Range("A:C"), PlotBy:=xlColumns
        ThisWorkbook.ActiveChart.Axes(xlCategory).MinimumScale = 6000
        ThisWorkbook.ActiveChart.Axes(xlCategory).MaximumScale = 8000
        ThisWorkbook.ActiveChart.Axes(xlValue).MaximumScale = 1400
        ThisWorkbook.ActiveChart.Location xlLocationAsObject, currentSheet.Name
        ThisWorkbook.ActiveChart.SeriesCollection(1).AxisGroup = 2  'set the axis to secondary for RSSI plot
        ThisWorkbook.ActiveChart.Legend.Position = xlLegendPositionBottom
        'highlight the search angle
        Dim idx5 As Integer
        idx5 = 171
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx5).MarkerSize = 10
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx5).MarkerStyle = 2
        ThisWorkbook.ActiveChart.SeriesCollection(1).Points(idx5).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx5).MarkerSize = 10
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx5).MarkerStyle = 2
        ThisWorkbook.ActiveChart.SeriesCollection(2).Points(idx5).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        
        For j = 1 To currentSheet.Shapes.Count
            If currentSheet.Shapes(j).Type = msoChart Then
                currentSheet.Shapes(j).Width = 4 * 120
                currentSheet.Shapes(j).Height = 3 * 120
                currentSheet.Shapes(j).Top = currentSheet.Shapes(j).Top + (j * 400) - 500
            End If
        Next j
        
         'find the next file - Dir without parameters will look for the next file in the folder that matches the first Dir call
        strFile = Dir
    Loop
End Sub
Sub Button2_Click()
    'Delete all worksheets except the first
    Application.DisplayAlerts = False
    Do While ThisWorkbook.Worksheets.Count > 1
        ThisWorkbook.Worksheets(2).Delete
    Loop
    
    Do While ThisWorkbook.Charts.Count > 0
        ThisWorkbook.Charts(1).Delete
    Loop
    Application.DisplayAlerts = True
End Sub
