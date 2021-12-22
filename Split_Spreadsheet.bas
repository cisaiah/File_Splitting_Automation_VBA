Attribute VB_Name = "Module1"
Option Explicit

Sub Split_Spreadsheet()

    ' This is a macro to automate splitting an excel worksheet into different workbooks
    ' The macro will allow the user to select your split column
    ' The user can also select whether or not to hide the split column, activate filter buttons or freeze panes
      
    Dim WBook As Workbook
    Dim WSheet As Worksheet
    Dim NewWB As Workbook
    Dim NewSH As Worksheet
    Dim TempSheet As Worksheet
    Dim WBPath As String
    Dim WBName As String
    Dim ValidSplitRange As Boolean
    Dim SplitRange As Range
    Dim SplitRow As Integer
    Dim SplitRowHeight As Integer
    Dim SplitCol As Integer
    Dim SplitArray() As Variant
    Dim HideSplitColumn As String
    Dim FilterButton As String
    Dim FreezePanes As String
    Dim FreezeCell As Range
    Dim counter As Integer
    Dim NewFile As String
    Dim hiddencol As String
    Dim allhiddencols As String
    Dim i As Integer
    Dim InputErrorMsg As String
    Dim StartTimer As Double
  
    
UserInput:

    'Get inputs from users
     On Error GoTo UserInputErrorHandler:
    
        '1. Get the Split Range
        ValidSplitRange = False
        Do Until ValidSplitRange
                      
            'Allow users to select the column the want to split
            Set SplitRange = Application.InputBox(Prompt:="Please click on the header of the column you want to be split into different files", _
                Title:="Select Split Range", Type:=8)
                                  
                'Error Handling where user chooses an empty cell
                If IsEmpty(SplitRange.Value) Then
                    MsgBox Prompt:="You have selected an empty cell. Please select the cell that contains the column header", _
                    Buttons:=vbExclamation, Title:="Error!"
                Else
                    ValidSplitRange = True
                End If
        Loop
        
        '2. Ask user if they want to hide the split column after splitting the source file
        HideSplitColumn = MsgBox(Prompt:="Do you want to hide the split column on the split files?", _
                Buttons:=vbYesNoCancel + vbQuestion, Title:="Hide Split Column?")
            If HideSplitColumn = vbCancel Then Exit Sub
        
        '3. Ask user if they want a filter button on the split files
        FilterButton = MsgBox(Prompt:="Do you want to activate the filter button on the split files?", _
                Buttons:=vbYesNoCancel + vbQuestion, Title:="Filter Button Required?")
            If FilterButton = vbCancel Then Exit Sub
        
        ' 4. 'Ask the user if (and where) they want freeze panes to apply
        FreezePanes = MsgBox(Prompt:="Do you want to activate a freeze panes on the split files?", _
                Buttons:=vbYesNoCancel + vbQuestion, Title:="Freeze Panes Required?")
            If FreezePanes = vbYes Then
                'Ask what cell or column the freeze panes should apply on
                Set FreezeCell = Application.InputBox(Prompt:="Please click on the cell where the freeze pane should apply", _
                    Title:="Select Split Range", Type:=8)
            Else:
                  If FreezePanes = vbCancel Then Exit Sub
            End If
            
     On Error GoTo 0
    'End of getting user inputs
        
    
    'Disable screen updating
    Application.ScreenUpdating = False
    
    'Get start of runtime
    StartTimer = Timer
        
    Set WBook = ActiveWorkbook
    Set WSheet = ActiveSheet
    
    WBPath = WBook.Path
    WBName = Left(WBook.Name, InStrRev(WBook.Name, ".") - 1)   'Remove the filename extension
    
    'Choose the first cell in case a user select multiple cells as split column header
    Set SplitRange = SplitRange.Cells(1, 1)
    
    'Clear any existing filter
    WSheet.AutoFilterMode = False
    
    'Sort the spreadsheet based on column to be split
    SplitRange.Sort Key1:=SplitRange.Cells(1, 1), Order1:=xlAscending, Header:=xlYes
            
    'Define the split array, split row, and split column
    Set SplitRange = Range(SplitRange, Cells((SplitRange.End(xlDown).Row), SplitRange.Column))
        SplitRow = SplitRange.Row
        SplitCol = SplitRange.Column
        SplitRowHeight = SplitRange.Cells(1, 1).RowHeight
    
    'Determine the hidden columns. I will unhide them for the splitting process and re-hide them afterwards
    allhiddencols = ""
    For i = 1 To 16384
        If Columns(i).Hidden Then
            hiddencol = Columns(i).Address(False, False)
            allhiddencols = allhiddencols & " " & hiddencol & ","     'Concatenate all hidden columns together
        End If
    Next i
    If Len(allhiddencols) > 1 Then
        allhiddencols = Left(allhiddencols, Len(allhiddencols) - 1)    'Remove the comma after the last hidden column
    End If
    
    'Unhide hidden columns
    WSheet.UsedRange.EntireColumn.Hidden = False

    'Get unique list of the field to be split by copying the SplitRange into a new (temporary) sheet and removing the duplicates
    SplitRange.Copy
    Set TempSheet = Sheets.Add
    TempSheet.Range("A1").PasteSpecial xlPasteAll
    TempSheet.Range("A1", Range("A1").End(xlDown)).RemoveDuplicates Columns:=Array(1), Header:=xlYes
    
    'Create an array and assign values from unique list to the array
    SplitArray = TempSheet.Range("A2", Range("A1").End(xlDown)).Value
    
    'Delete TempSheet (the temporary worksheet)
    Application.DisplayAlerts = False
    TempSheet.Delete
    Application.DisplayAlerts = True
    
    
    'THE LOOP TO SPLIT THE FILES
        
        For counter = LBound(SplitArray) To UBound(SplitArray)
               
            'Create new workbook and worksheet, and rename new worksheet
            Set NewWB = Workbooks.Add
            Set NewSH = NewWB.Worksheets(1)
                NewSH.Name = WSheet.Name
                
            'Filter the data on the split column of the Spreadsheet
            SplitRange.Cells(1, 1).AutoFilter Field:=SplitCol, Criteria1:=SplitArray(counter, 1)
            
            'Copy filtered/visible data and paste into new workbook created
            WSheet.UsedRange.SpecialCells(xlCellTypeVisible).Copy
            NewSH.Range("A1").PasteSpecial xlPasteColumnWidths
            NewSH.Range("A1").PasteSpecial xlPasteAllUsingSourceTheme
            
            'Fix header row height to be same as that of the source file
            SplitRange.Cells(1, 1).RowHeight = SplitRowHeight
            
            'Hide back columns that were hidden in source file
            If Len(allhiddencols) > 1 Then
                NewSH.Range(allhiddencols).EntireColumn.Hidden = True
            End If
            
            'Hide Split Column if user requested to do so
            If HideSplitColumn = vbYes Then
                NewSH.Cells(SplitRow, SplitCol).EntireColumn.Hidden = True
            End If
            
            'Apply freeze panes if user requested to do so
            If FreezePanes = vbYes Then
                NewSH.Range(FreezeCell.Address).Select
                ActiveWindow.FreezePanes = True
            End If
                       
            'Apply Autofilter on column headers if user requested to do so
            If FilterButton = vbYes Then
                NewSH.Range(Cells(SplitRow, 1), Cells(SplitRow, NewSH.UsedRange.Columns.Count)).AutoFilter
            End If
            
            'Select the cell you want the cursor to appear on when the file is opened
            NewSH.Range("A" & SplitRow).Select
            
            'Save the file as "Source FileName " + Cell content of the split column + ".xlsx"
            Application.EnableEvents = False
                NewFile = WBPath & "\" & WBName & " - " & Replace(SplitArray(counter, 1), "/", " or ") & ".xlsx"
                NewWB.SaveAs NewFile
                NewWB.Close False
            Application.EnableEvents = True
            
         Next counter
     'END OF LOOP
     
     'Clear the filter on the source sheet
     WSheet.AutoFilterMode = False
     WSheet.Range(Cells(SplitRow, 1), Cells(SplitRow, WSheet.UsedRange.Columns.Count)).AutoFilter
     
     'Hide back columns that need to be hidden on the Source File so it appears just the way it was before running the macro
      If Len(allhiddencols) > 1 Then
            WSheet.Range(allhiddencols).EntireColumn.Hidden = True
      End If
       
     'Re-enable screen updating
     Application.ScreenUpdating = True
     
     MsgBox Prompt:=Str(UBound(SplitArray)) + " file(s) were created in " + WBPath + vbNewLine + vbNewLine + "Runtime: " + _
                  Str(Application.WorksheetFunction.Floor((Timer - StartTimer) / 60, 1)) + " mins and " + Str((Timer - StartTimer) Mod 60) + _
                  " secs", Buttons:=vbInformation, Title:="Completed"
     
     Exit Sub
     
UserInputErrorHandler:
      InputErrorMsg = MsgBox(Prompt:="There is an error with your input/selection. Please try again.", Buttons:=vbExclamation + vbRetryCancel, Title:="Error")
      
      If InputErrorMsg = vbRetry Then
            GoTo UserInput
      End If
      
End Sub

