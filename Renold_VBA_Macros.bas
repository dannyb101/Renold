Attribute VB_Name = "Module4"

Sub Convert_Data_2_Numbers()
Attribute Convert_Data_2_Numbers.VB_Description = "This macro will convert all columns in a sheet from text to columns, as well as set the format of all cells to ""General"""
Attribute Convert_Data_2_Numbers.VB_ProcData.VB_Invoke_Func = " \n14"
'
'This macro will convert all columns in A SINGLE SHEET from text to columns, as well as set the format of all cells to "General"
'


    Dim cell As Range
    Set cell = Range("A1")
    
    cell.Select
    Do Until IsEmpty(cell)
        ActiveCell.EntireColumn.Select
        Selection.TextToColumns , DataType:=xlDelimited _
            , TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Selection.NumberFormat = "General"
        Set cell = cell.Offset(0, 1)
        cell.Select
    Loop

End Sub

Sub All_Sheets_Convert_Data_2_Numbers()
Attribute All_Sheets_Convert_Data_2_Numbers.VB_Description = "This macro will convert all columns in ALL SHEETS from text to columns, as well as set the format of all cells to ""General"""
Attribute All_Sheets_Convert_Data_2_Numbers.VB_ProcData.VB_Invoke_Func = "d\n14"
'
'This macro will convert all columns in ALL SHEETS from text to columns, as well as set the format of all cells to "General"
'

    Dim Current As Worksheet

    ' Loop through all of the worksheets in the active workbook.
    For Each Current In Worksheets
        Current.Select
        Call Convert_Data_2_Numbers
    Next
End Sub

Sub Summary_Sheet()
'
' This macro will insert a summary sheet at the beginning of a workbook, giving average dynamic stiffness, magnifier and amplitude
'
    Dim stiffCell As Range
    Dim magCell As Range
    Dim angCell As Range
    Dim wS As Worksheet
    Dim i As Integer
    Dim xScreen As Boolean
    Dim rng As Range
    Dim cell As Range
    Dim stiffTotal As Double
    Dim magTotal As Double
        
    xScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    'Create summary sheet
    Sheets.Add Before:=Sheet1
    Worksheets(1).Select
    ActiveSheet.Name = "Summary"
    Set stiffCell = Range("B4")
    Set magCell = Range("D4")
    Set angCell = Range("F4")
    
    For Each wS In Worksheets
        If wS.Name = "Summary" Then
        
        Else
            wS.Select
            Set rng = Range("A1:J1")
            For Each cell In rng
                If Left(cell.Value, 5) = "Angle" Then
                    'Copy and paste Angle column
                    cell.Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.Copy
                    Sheets("Summary").Select
                    angCell.Select
                    ActiveSheet.Paste
                
                ElseIf Left(cell.Value, 19) = "Magnifier (DIN 740)" Then
                    'Copy and paste Magnifier column
                    Range(cell, cell.Offset(8000, 0)).Select
                    Selection.Copy
                    Sheets("Summary").Select
                    magCell.Select
                    ActiveSheet.Paste
                        
                ElseIf Left(cell.Value, 17) = "Dynamic Stiffness" Then
                    'Copy and paste dynamic stiffness column
                    cell.Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.Copy
                    Sheets("Summary").Select
                    stiffCell.Select
                    ActiveSheet.Paste
                Else
                End If
                wS.Select
            Next cell
            
            Sheets("Summary").Select
            
            'Write stiffness averaging formula
            stiffCell.Offset(2, 1).Value = "Average"
            stiffCell.Offset(3, 1).FormulaR1C1 = "=AVERAGE(R[5994]C[-1]:R[7993]C[-1])"
            stiffTotal = stiffTotal + stiffCell.Offset(3, 1).Value
            
            'Write magnifier formula
            magCell.Offset(2, 1).Select
            ActiveCell.FormulaR1C1 = "Average"
            ActiveCell.Offset(1, 0).Select
            ActiveCell.FormulaR1C1 = "=AVERAGE(R[5994]C[-1]:R[7993]C[-1])"
            magTotal = magTotal + ActiveCell.Value
            
            'Write angle formula
            angCell.Offset(2, 1).Value = "Amplitude"
            angCell.Offset(3, 1).FormulaR1C1 = "=AVERAGE((MAX(R[5994]C[-1]:R[6194]C[-1])-MIN(R[5994]C[-1]:R[6194]C[-1])) ,(MAX(R[6196]C[-1]:R[6396]C[-1])-MIN(R[6196]C[-1]:R[6396]C[-1]))    ,(MAX(R[6396]C[-1]:R[6596]C[-1])-MIN(R[6396]C[-1]:R[6596]C[-1]))    ,(MAX(R[6596]C[-1]:R[6796]C[-1])-MIN(R[6596]C[-1]:R[6796]C[-1])),(MAX(R[6796]C[-1]:R[6996]C[-1])-MIN(R[6796]C[-1]:R[6996]C[-1])),(MAX(R[6996]C[-1]:R[7196]C[-1])-MIN(R[6996]C[-1]:R[7196]C[-1])),(MAX(R[7196]C[-1]:R[7396]C[-1])-MIN(R[7196]C[-1]:R[7396]C[-1])),(MAX(R[7396]C[-1]:R[7596]C[-1])-MIN(R[7396]C[-1]:R[7596]C[-1])),(MAX(R[7596]C[-1]:R[7796]C[-1])-MIN(R[7596]C[-1]:R[7796]C[-1])),(MAX(R[7796]C[-1]:R[7996]C[-1])-MIN(R[7796]C[-1]:R[7996]C[-1])))"
            angCell.Offset(4, 1).FormulaR1C1 = "=R[-1]C[0]*0.0174533"
            angCell.Offset(3, 2).Value = "Deg"
            angCell.Offset(4, 2).Value = "Rad"
            
            stiffCell.Offset(-1, 0).Value = wS.Name
            Range(stiffCell.Offset(-1, 0), angCell.Offset(-1, 2)).Merge
            
            Set stiffCell = stiffCell.Offset(0, 7)
            Set angCell = angCell.Offset(0, 7)
            Set magCell = magCell.Offset(0, 7)
        End If
    Next
    
    'Create average read out at top of screen for stiffness and magnifier
    Rows("1:3").HorizontalAlignment = xlCenter
    Range("B1").Value = "Stiffness Average:"
    Range("B1:C1").Select
    Selection.Merge
    Range("B2").Value = (stiffTotal / (Application.Sheets.Count - 1))
    Range("B2:C2").Select
    
    Range("B1:C2").Select
    Selection.BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    Selection.Interior.ColorIndex = 27
    Range("G1").Value = "Magnifier Average:"
    Range("G1:H1").Select
    Selection.Merge
    Range("G2").Value = (magTotal / (Application.Sheets.Count - 1))
    Range("G2:H2").Select
    
    Range("G1:H2").Select
    Selection.BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    Selection.Interior.ColorIndex = 27
    
    Application.ScreenUpdating = xScreen
    
End Sub
Sub Test_Summary_Sheet()
'
' This macro will insert a summary sheet at the beginning of a workbook, giving average dynamic stiffness, magnifier and amplitude
'
    Dim stiffCell As Range
    Dim magCell As Range
    Dim angCell As Range
    Dim wS As Worksheet
    Dim i As Long
    Dim xScreen As Boolean
    Dim rng As Range
    Dim cell As Range
    Dim stiffTotal As Double
    Dim magTotal As Double
    Dim dataTitle As Variant
    Dim selectionList As Variant
    
        
    xScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    'Create summary sheet
    Sheets.Add Before:=Sheet1
    Worksheets(1).Select
    ActiveSheet.Name = "Summary"
    Set stiffCell = Range("B4")
    Set magCell = Range("D4")
    Set angCell = Range("F4")
    
    For Each wS In Worksheets
        If wS.Name = "Summary" Then
        
        Else
            wS.Select
            Set rng = Range("A1")
            i = 0
            Do While Not IsEmpty(rng.Value)
                
                
                rng = rng.Offset(0, 1)
                ListBox1.AddItem cell.Value
                i = i + 1
            Loop
            
            UserForm1.Show
            
            'selectionList = SelectionBoxMulti(List:=dataTitle, Prompt:="Please select all data you would like to process", SelectionType:=fmMultiSelectMulti, title:="Select multiple")
            
            Set rng = Range("A1")
            'Do While Not IsEmpty(rng.Value)
                'If rng.Value =
                
'                If Left(cell.Value, 5) = "Angle" Then
'                    'Copy and paste Angle column
'                    cell.Select
'                    Range(Selection, Selection.End(xlDown)).Select
'                    Selection.Copy
'                    Sheets("Summary").Select
'                    angCell.Select
'                    ActiveSheet.Paste
'
'                ElseIf Left(cell.Value, 19) = "Magnifier (DIN 740)" Then
'                    'Copy and paste Magnifier column
'                    Range(cell, cell.Offset(8000, 0)).Select
'                    Selection.Copy
'                    Sheets("Summary").Select
'                    magCell.Select
'                    ActiveSheet.Paste
'
'                ElseIf Left(cell.Value, 17) = "Dynamic Stiffness" Then
'                    'Copy and paste dynamic stiffness column
'                    cell.Select
'                    Range(Selection, Selection.End(xlDown)).Select
'                    Selection.Copy
'                    Sheets("Summary").Select
'                    stiffCell.Select
'                    ActiveSheet.Paste
'                Else
'                End If
'                wS.Select
            
            
            Sheets("Summary").Select
            
            'Write stiffness averaging formula
            stiffCell.Offset(2, 1).Value = "Average"
            stiffCell.Offset(3, 1).FormulaR1C1 = "=AVERAGE(R[5994]C[-1]:R[7993]C[-1])"
            stiffTotal = stiffTotal + stiffCell.Offset(3, 1).Value
            
            'Write magnifier formula
            magCell.Offset(2, 1).Select
            ActiveCell.FormulaR1C1 = "Average"
            ActiveCell.Offset(1, 0).Select
            ActiveCell.FormulaR1C1 = "=AVERAGE(R[5994]C[-1]:R[7993]C[-1])"
            magTotal = magTotal + ActiveCell.Value
            
            'Write angle formula
            angCell.Offset(2, 1).Value = "Amplitude"
            angCell.Offset(3, 1).FormulaR1C1 = "=AVERAGE((MAX(R[5994]C[-1]:R[6194]C[-1])-MIN(R[5994]C[-1]:R[6194]C[-1])) ,(MAX(R[6196]C[-1]:R[6396]C[-1])-MIN(R[6196]C[-1]:R[6396]C[-1]))    ,(MAX(R[6396]C[-1]:R[6596]C[-1])-MIN(R[6396]C[-1]:R[6596]C[-1]))    ,(MAX(R[6596]C[-1]:R[6796]C[-1])-MIN(R[6596]C[-1]:R[6796]C[-1])),(MAX(R[6796]C[-1]:R[6996]C[-1])-MIN(R[6796]C[-1]:R[6996]C[-1])),(MAX(R[6996]C[-1]:R[7196]C[-1])-MIN(R[6996]C[-1]:R[7196]C[-1])),(MAX(R[7196]C[-1]:R[7396]C[-1])-MIN(R[7196]C[-1]:R[7396]C[-1])),(MAX(R[7396]C[-1]:R[7596]C[-1])-MIN(R[7396]C[-1]:R[7596]C[-1])),(MAX(R[7596]C[-1]:R[7796]C[-1])-MIN(R[7596]C[-1]:R[7796]C[-1])),(MAX(R[7796]C[-1]:R[7996]C[-1])-MIN(R[7796]C[-1]:R[7996]C[-1])))"
            angCell.Offset(4, 1).FormulaR1C1 = "=R[-1]C[0]*0.0174533"
            angCell.Offset(3, 2).Value = "Deg"
            angCell.Offset(4, 2).Value = "Rad"
            
            stiffCell.Offset(-1, 0).Value = wS.Name
            Range(stiffCell.Offset(-1, 0), angCell.Offset(-1, 2)).Merge
            
            Set stiffCell = stiffCell.Offset(0, 7)
            Set angCell = angCell.Offset(0, 7)
            Set magCell = magCell.Offset(0, 7)
        End If
    Next
    
    'Create average read out at top of screen for stiffness and magnifier
    Rows("1:3").HorizontalAlignment = xlCenter
    Range("B1").Value = "Stiffness Average:"
    Range("B1:C1").Select
    Selection.Merge
    Range("B2").Value = (stiffTotal / (Application.Sheets.Count - 1))
    Range("B2:C2").Select
    
    Range("B1:C2").Select
    Selection.BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    Selection.Interior.ColorIndex = 27
    Range("G1").Value = "Magnifier Average:"
    Range("G1:H1").Select
    Selection.Merge
    Range("G2").Value = (magTotal / (Application.Sheets.Count - 1))
    Range("G2:H2").Select
    
    Range("G1:H2").Select
    Selection.BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    Selection.Interior.ColorIndex = 27
    
    Application.ScreenUpdating = xScreen
    
End Sub

Private Sub UserForm_Initialize()

    

End Sub
Sub Static_Summary_Sheet()
'
' This macro will insert a summary sheet at the beginning of a workbook, giving average dynamic stiffness, magnifier and amplitude
'
    Dim stiffCell As Range
    Dim magCell As Range
    Dim angCell As Range
    Dim wS As Worksheet
    Dim i As Integer
    Dim xScreen As Boolean
    Dim rng As Range
    Dim cell As Range
    Dim stiffTotal As Double
    Dim magTotal As Double
        
    xScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    'Create summary sheet
    Sheets.Add Before:=Sheet1
    Worksheets(1).Select
    ActiveSheet.Name = "Summary"
    Set stiffCell = Range("B4")
    Set magCell = Range("D4")
    Set angCell = Range("F4")
    
    For Each wS In Worksheets
        If wS.Name = "Summary" Then
        
        Else
            wS.Select
            Set rng = Range("A1:J1")
            For Each cell In rng
                If Left(cell.Value, 5) = "Angle" Then
                    'Copy and paste Angle column
                    cell.Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.Copy
                    Sheets("Summary").Select
                    angCell.Select
                    ActiveSheet.Paste
                
                ElseIf Left(cell.Value, 19) = "Magnifier (DIN 740)" Then
                    'Copy and paste Magnifier column
                    Range(cell, cell.Offset(6000, 0)).Select
                    Selection.Copy
                    Sheets("Summary").Select
                    magCell.Select
                    ActiveSheet.Paste
                        
                ElseIf Left(cell.Value, 20) = "Torque (Compensated)" Then
                    'Copy and paste dynamic stiffness column
                    cell.Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.Copy
                    Sheets("Summary").Select
                    stiffCell.Select
                    ActiveSheet.Paste
                Else
                End If
                wS.Select
            Next cell
            
            Sheets("Summary").Select
            
            'Write torque amplitudes
            stiffCell.Offset(2, 1).Value = "Amplitude"
            stiffCell.Offset(3, 1).FormulaR1C1 = "=R[500]C[-1]-R[1000]C[-1]"
            stiffCell.Offset(4, 1).FormulaR1C1 = "=R[1000]C[-1]-R[1500]C[-1]"
            stiffCell.Offset(5, 1).FormulaR1C1 = "=R[2000]C[-1]-R[2500]C[-1]"
            stiffTotal = stiffTotal + stiffCell.Offset(3, 1).Value
            
            
            'Write magnifier formula
            magCell.Offset(2, 1).Select
            ActiveCell.FormulaR1C1 = "Average"
            ActiveCell.Offset(1, 0).Select
            ActiveCell.FormulaR1C1 = "=AVERAGE(R[29]C[-1]:R[49]C[-1])"
            magTotal = magTotal + ActiveCell.Value
            
            'Write angle amplitudes
            angCell.Offset(2, 1).Value = "Amplitude"
            angCell.Offset(3, 1).FormulaR1C1 = "=R[500]C[-1]-R[1000]C[-1]"
            angCell.Offset(4, 1).FormulaR1C1 = "=R[1000]C[-1]-R[1500]C[-1]"
            angCell.Offset(5, 1).FormulaR1C1 = "=R[2000]C[-1]-R[2500]C[-1]"
            
            angCell.Offset(7, 1).FormulaR1C1 = "=RADIANS(R[-4]C[0])"
            angCell.Offset(8, 1).FormulaR1C1 = "=RADIANS(R[-4]C[0])"
            angCell.Offset(9, 1).FormulaR1C1 = "=RADIANS(R[-4]C[0])"
            
            'angCell.Offset(3, 1).FormulaR1C1 = "=AVERAGE((MAX(R[2994]C[-1]:R[3194]C[-1])-MIN(R[2994]C[-1]:R[3194]C[-1])) ,(MAX(R[3196]C[-1]:R[3396]C[-1])-MIN(R[3196]C[-1]:R[3396]C[-1]))    ,(MAX(R[3396]C[-1]:R[3596]C[-1])-MIN(R[3396]C[-1]:R[3596]C[-1]))    ,(MAX(R[3596]C[-1]:R[3796]C[-1])-MIN(R[3596]C[-1]:R[3796]C[-1])),(MAX(R[3796]C[-1]:R[3996]C[-1])-MIN(R[3796]C[-1]:R[3996]C[-1])),(MAX(R[3996]C[-1]:R[4196]C[-1])-MIN(R[3996]C[-1]:R[4196]C[-1])),(MAX(R[4196]C[-1]:R[4396]C[-1])-MIN(R[4196]C[-1]:R[4396]C[-1])),(MAX(R[4396]C[-1]:R[4596]C[-1])-MIN(R[4396]C[-1]:R[4596]C[-1])),(MAX(R[4596]C[-1]:R[4796]C[-1])-MIN(R[4596]C[-1]:R[4796]C[-1])),(MAX(R[4796]C[-1]:R[4996]C[-1])-MIN(R[4796]C[-1]:R[4996]C[-1])))"
            'angCell.Offset(4, 1).FormulaR1C1 = "=R[-1]C[0]*0.0174533"
            angCell.Offset(3, 2).Value = "Deg"
            angCell.Offset(7, 2).Value = "Rad"
            
            angCell.Offset(11, 2).Value = "Stiffness"
            angCell.Offset(12, 2).Formula2R1C1 = "=AVERAGE( (R[-9]C[-5]/R[-5]C[-1]), (R[-8]C[-5]/R[-4]C[-1]), (R[-7]C[-5]/R[-3]C[-1]) )"
            
            stiffCell.Offset(-1, 0).Value = wS.Name
            Range(stiffCell.Offset(-1, 0), angCell.Offset(-1, 2)).Merge
            
            Set stiffCell = stiffCell.Offset(0, 7)
            Set angCell = angCell.Offset(0, 7)
            Set magCell = magCell.Offset(0, 7)
        End If
    Next
    
    'Create average read out at top of screen for stiffness and magnifier
    Rows("1:3").HorizontalAlignment = xlCenter
    Range("B1").Value = "Stiffness Average:"
    Range("B1:C1").Select
    Selection.Merge
    Range("B2").Value = (stiffTotal / (Application.Sheets.Count - 1))
    Range("B2:C2").Select
    
    Range("B1:C2").Select
    Selection.BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    Selection.Interior.ColorIndex = 27
    Range("G1").Value = "Magnifier Average:"
    Range("G1:H1").Select
    Selection.Merge
    Range("G2").Value = (magTotal / (Application.Sheets.Count - 1))
    Range("G2:H2").Select
    
    Range("G1:H2").Select
    Selection.BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    Selection.Interior.ColorIndex = 27
    
    Application.ScreenUpdating = xScreen
    
End Sub

Function Import_Text_Files() As Boolean
'This macro will allow you to import text files saved on the Servotest rig into separate sheets of a single workbook
'The macro will work whether or not the header lines are left in
    Dim xFilesToOpen As Variant
    Dim i As Integer
    Dim xWb As Workbook
    Dim xTempWb As Workbook
    Dim xDelimiter As String
    Dim xScreen As Boolean
    On Error GoTo ErrHandler
    Import_Text_Files = True
    xScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    xDelimiter = ","
    xFilesToOpen = Application.GetOpenFilename("Text Files (*.txt), *.txt", , , , True)
    If TypeName(xFilesToOpen) = "Boolean" Then
        MsgBox "No files were selected"
        Import_Text_Files = False
        GoTo ExitHandler
    End If
    i = 1
    Set xTempWb = Workbooks.Open(xFilesToOpen(i)) 'Opens the text file to an excel doc (all info still in column A)
    xTempWb.Sheets(1).Copy  'Copies the sheet to a new workbook
    Set xWb = Application.ActiveWorkbook
    xTempWb.Close False
    xWb.Worksheets(i).Columns("A:A").TextToColumns _
      Destination:=Range("A1"), DataType:=xlDelimited, _
      TextQualifier:=xlDoubleQuote, _
      ConsecutiveDelimiter:=False, _
      Tab:=False, Semicolon:=False, _
      Comma:=True, Space:=False, _
      Other:=True, OtherChar:="|"
    While Range("B1").Value = ""
        xWb.Worksheets(i).Rows(1).EntireRow.Delete
    Wend
    Do While i < UBound(xFilesToOpen)
        i = i + 1
        Set xTempWb = Workbooks.Open(xFilesToOpen(i))
        With xWb
            xTempWb.Sheets(1).Move After:=.Sheets(.Sheets.Count)
            .Worksheets(i).Columns("A:A").TextToColumns _
              Destination:=Range("A1"), DataType:=xlDelimited, _
              TextQualifier:=xlDoubleQuote, _
              ConsecutiveDelimiter:=False, _
              Tab:=False, Semicolon:=False, _
              Comma:=False, Space:=False, _
              Other:=True, OtherChar:=xDelimiter
        End With
        While Range("B1").Value = ""
            xWb.Worksheets(i).Rows(1).EntireRow.Delete
        Wend
    Loop
ExitHandler:
    Application.ScreenUpdating = xScreen
    Set xWb = Nothing
    Set xTempWb = Nothing
    Exit Function
ErrHandler:
    MsgBox Err.Description
    Resume ExitHandler
End Function

Sub Dynamic_Servotest_Sheet()

Dim Import As Boolean

Import = Import_Text_Files

If Import = True Then
    Call All_Sheets_Convert_Data_2_Numbers
    Call Summary_Sheet
End If

End Sub
Sub Static_Servotest_Sheet()

Dim Import As Boolean

Import = Import_Text_Files

If Import = True Then
    Call All_Sheets_Convert_Data_2_Numbers
    Call Static_Summary_Sheet
End If

End Sub

Sub Change_Axes_Titles_Single()
'
' AxisT Macro
    
    ActiveSheet.ChartObjects(1).Activate
    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.text = "Angle, Degrees"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.text = "Torque, kNm"
    End With

'
End Sub

Sub Change_Axes_Titles_All()

Dim wS As Worksheet


For Each wS In Worksheets
wS.Select

Call Change_Axes_Titles_Single

Next wS

End Sub

Sub Change_Chart_Titles_All()

Dim wS As Worksheet
Dim title As String
title = "HTB 10000 Dynamic Torque Deflections" & Chr(10) & "±kNm @ Hz"

For Each wS In Worksheets
wS.Select

ActiveSheet.ChartObjects(1).Activate
ActiveChart.HasTitle = True
ActiveChart.ChartTitle.text = title


Next wS

End Sub

Sub Create_Graphs_All()

'Create all graphs A768:B872

Dim wS As Worksheet

For Each wS In Worksheets
wS.Select

ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
ActiveChart.SetSourceData Source:=Range("$A$768:$B$872")

Next wS

End Sub

Sub Change_All_Titles()

Dim wS As Worksheet
Dim title As String
title = "HTB 10000 Dynamic Torque Deflections" & Chr(10) & "±kNm @ Hz"

For Each wS In Worksheets
wS.Select

Call Change_Axes_Titles_Single

'ActiveSheet.ChartObjects(1).Activate
ActiveChart.HasTitle = True
ActiveChart.ChartTitle.text = title


Next wS

End Sub

Sub Highlight_Higher_Than()

Dim myRange As Range
Dim myVal As Range
Dim cellVal As Range


Range("af6").Select
Set myRange = Range(Selection, Selection.End(xlDown))


For Each myVal In myRange
'Set cellVal = myVal.Offset(-1, 0)
If myVal.Value > 1.56 Or myVal.Value < 0.93 Then
    myVal.Interior.ColorIndex = 27
End If
Next

End Sub

Sub graphtitle()
'
' graphtitle Macro
'

'
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.text = "Angle, Degrees"
    Selection.Format.TextFrame2.TextRange.Characters.text = "Angle, Degrees"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 14).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 14).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.text = "Torque, kNm"
    Selection.Format.TextFrame2.TextRange.Characters.text = "Torque, kNm"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 12).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 12).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementLegendNone)
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveSheet.Shapes("Chart 1").TextFrame2.TextRange.Characters.text = ""
End Sub





