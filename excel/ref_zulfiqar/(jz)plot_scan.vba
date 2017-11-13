Sub Button1_Click()
    Dim strFile As String, strPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Const xlDelimited = 1
    Const xlTextQualifierNone = -4142
     
     'Change this path for your own file location:
    strPath = "U:\TIM1xx\TIM100\TSD Data\17420111\"
    
    
    '*******************
    'Delete all worksheets except the first
    Application.DisplayAlerts = False
    Do While ThisWorkbook.Worksheets.Count > 1
        ThisWorkbook.Worksheets(2).Delete
    Loop
    
    Do While ThisWorkbook.Charts.Count > 0
        ThisWorkbook.Charts(1).Delete
    Loop
    Application.DisplayAlerts = True
    '*******************
    
    
     'add a new worksheet at the end of the macro workbook to paste into
    ThisWorkbook.Worksheets.Add After:=Sheets(Sheets.Count)
    Set currentSheet = ThisWorkbook.Worksheets(Sheets.Count)
    currentSheet.Name = Mid(strPath, Len(strPath) - 8, 8)
    
     'this returns an empty string "" if the file cannot be found and will error if the folder is incorrect
    strFile = Dir(strPath & "*.txt")
    
    Dim file_number As Integer
    file_number = 0
    
    Do While strFile <> ""
        
         'open the csv file and assign it to a variable so that we can easily reference it later
        Set file = fso.OpenTextFile(strPath + strFile, 1)
        StrData = file.ReadLine
        
        Row = 1
        Do Until file.AtEndOfStream
             'paste it
            currentSheet.Cells(Row, 1 + 4 * file_number) = StrData
            StrData = file.ReadLine
            Row = Row + 1
        Loop
        
        'With currentSheet
            '.Columns("A:A").TextToColumns .Range("A1"), xlDelimited, xlDoubleQuote, True, , True, , False
        'End With
        currentSheet.Columns(1 + 4 * file_number).TextToColumns currentSheet.Cells(1, 1 + 4 * file_number), xlDelimited, xlDoubleQuote, True, , True, , False
        
        plot_scan_chart currentSheet.Columns(1 + 4 * file_number).Resize(, 3)

         'find the next file - Dir without parameters will look for the next file in the folder that matches the first Dir call
        strFile = Dir
        file_number = file_number + 1
    Loop
    
    'Arrange the charts position
    For j = 1 To currentSheet.Shapes.Count
        If currentSheet.Shapes(j).Type = msoChart Then
            currentSheet.Shapes(j).Width = 4 * 100
            currentSheet.Shapes(j).Height = 3 * 100
            'currentSheet.Shapes(j).Top = currentSheet.Shapes(j).Top + (j * 315) - 520
            'currentSheet.Shapes(j).Left = currentSheet.Shapes(j).Left - 100
            currentSheet.Shapes(j).Top = (((j - 1) Mod 5) * 300)
            currentSheet.Shapes(j).Left = ((j - 1) \ 5) * 400
        End If
    Next j
    
End Sub

Function plot_scan_chart(plot_range As Range)
        createNewChart plot_range, -1200, 1200, 3000, 97, "0°"  'Chart for 1.5m targets
        createNewChart plot_range, 1200, 2200, 3000, 119, "18°" 'Chart for 2.9m target
        createNewChart plot_range, 8000, 10000, 2000, 191, "90°" 'Chart for 1m target
        createNewChart plot_range, -10000, -8000, 100, 11, "-90°" 'Chart for 50mm target
        createNewChart plot_range, 6000, 8000, 1600, 171, "70°" 'Chart for 0.7m target
End Function


Function createNewChart(plot_range As Range, lower_x As Integer, upper_x As Integer, upper_y As Integer, mark As Integer, dg As String)
    Dim chrt As Chart
    With ActiveSheet.Shapes.AddChart.Chart
        .Parent.Height = 10
        .Parent.Width = 20
        .HasTitle = True
        .ChartTitle.Text = "Plot Scan " + dg + " Degrees"
        .ChartType = xlXYScatterLines
        '.SetSourceData Source:=ActiveSheet.Range("A:C"), PlotBy:=xlColumns
        .SetSourceData Source:=plot_range, PlotBy:=xlColumns
        .Axes(xlCategory).MinimumScale = lower_x
        .Axes(xlCategory).MaximumScale = upper_x
        .Axes(xlValue).MaximumScale = upper_y
        .Axes(xlValue).MajorGridlines.Format.Line.Visible = msoFalse
        
        .Location xlLocationAsObject, ActiveSheet.Name
        .SeriesCollection(1).AxisGroup = 2  'set the axis to secondary for RSSI plot
        .Legend.Position = xlLegendPositionBottom
        
        
        'plotArea style
        
        .ChartArea.Format.Line.Visible = msoCTrue
        .ChartArea.Format.Line.Style = msoLineSingle
        .ChartArea.Format.Line.Weight = 0.75
        
        .PlotArea.Format.Line.Visible = msoCTrue
        .PlotArea.Format.Line.Style = msoLineSingle
        .PlotArea.Format.Line.Weight = 1.75
        
        
        'line chart style
        .SeriesCollection(1).MarkerSize = 5
        .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
        .SeriesCollection(1).Format.Line.Weight = 1
        
        .SeriesCollection(2).MarkerSize = 5
        .SeriesCollection(2).MarkerStyle = xlMarkerStyleCircle
        .SeriesCollection(2).Format.Line.Weight = 1
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(237, 125, 49)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        
        'highlight the search angle
        .SeriesCollection(1).Points(mark).MarkerSize = 10
        .SeriesCollection(1).Points(mark).MarkerStyle = 2
        .SeriesCollection(1).Points(mark).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(2).Points(mark).MarkerSize = 10
        .SeriesCollection(2).Points(mark).MarkerStyle = 2
        .SeriesCollection(2).Points(mark).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
    End With
    

End Function



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
