Functions to conveniently analyse data.

# TextToColumns
```vba
    Range("A1").Select
    ActiveWindow.SmallScroll Down:=-9
    ActiveSheet.Paste
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
    Selection.NumberFormat = "General"
```
??error_message: Microsoft can convert only 1 column at a time.

# String manipulation
Left, Right, Mid
Right(strFile, Len(strFile) - 25)
        
test

# 'Delete all worksheets except the first
```vba
    Application.DisplayAlerts = False
    Do While ThisWorkbook.Worksheets.Count > 1
        ThisWorkbook.Worksheets(2).Delete
    Loop
    
    Do While ThisWorkbook.Charts.Count > 0
        ThisWorkbook.Charts(1).Delete
    Loop
    Application.DisplayAlerts = True
```


# TODO:
* compare the graph with ideal case.
* delete all sheets?
* format autogenerated graphs, to the same style with Steffan one.
* 
