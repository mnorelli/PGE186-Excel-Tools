Attribute VB_Name = "Macros"
Public Const ERROR_BLANK_CHANGEID As Long = vbObjectError + 513
'Version: 1.5

'**************************************************************
'RIBBONSETUP TEMPLATE AUTHOR: Chris Newman, TheSpreadsheetGuru
'Instructions on how to use this template can be found at:
'https://www.thespreadsheetguru.com/blog/step-by-step-instructions-create-first-excel-ribbon-vba-addin
'https://www.thespreadsheetguru.com/myfirstaddin-help/
'**************************************************************

'**************************************************************
'Macros herein: Michael Norelli, Celerity Consulting Group, Inc.
'
'Change history:
'1.5
'- swaps Source Docs macros button names to correctly run the applicable macro
'- renames Create_Source_Doc_Table to MakeSourceDocsRefTable
'- renames CreateRatingsRequestedTable to MakeRatingsRequestedTable
'- abstracts away table creation to a general sub receving passed parameters and changes search for "=" to "Like" to allow for more search flexibility
'- adds buttons and code to make tables for Relays, Equipment Added, Equipment Retired
'- changes order of buttons to make tables in match Word doc order
'- checks for Change IDs as a prerequisite for running MakeSourceDocsRefTable
'- prevents Highlight macro from changing colors in legend at bottom below data rows
'1.4
'- Adds tool to make Source Documents Used table
'- Adds tool to make Ratings Requested table
'1.3
'- Removes "Retired" rows from SourceDocRefTable.
'- Rearranges tools into three groups.
'- Adds ability to update the add-in to the current version in P:\PGE186\Code Tricks
'Previous
'- adds Dan's code to create Source Doc Ref Table
'- adds code to move update rows to main table for Source Doc Ref Table
'- removes the general FindDistinctSubstrings() macro and replaced
'  with the more specific AddZeroes()
'- adds code to find last Date field for Paint()
'**************************************************************
Option Explicit
Dim LatestAddIn As String
Dim LastInstalledAddIn As String
Public Sub Paint()
'https://stackoverflow.com/questions/29085029/excel-macro-change-the-row-based-on-value
' This macro looks for the Type of Change values in the first column of the active spreadsheet
' and sets the background and font to the indicated color for the row containing the indicated text.
' "None" clears font and background, "Red" makes text red.
' Designed to make set up for PGE186 T-Line Line Data Sheets easier to set up
'M. Norelli 1/11/2019

  Application.GoTo Sheets("CAISO Update").Cells(1, 1)
  highlight "", "None"
  highlight "Create", 142, 169, 219
  highlight "Update", 244, 176, 132
  highlight "Retire", 255, 255, 0
  highlight "Previously Submitted*", 146, 208, 80
  highlight "Verification", "None"

End Sub
Private Sub highlight(word As String, r, Optional G As Integer, Optional B As Integer)
'This colors rows based on the Type of Change entry and RGB values passed from the Paint() macro
'M. Norelli 1/11/2019
    Dim LastColumnNum As Long

    'LastColumn = BX, as of Feb 2019
    LastColumnNum = FindLastCol("Date")
    'paint only to last Date colun, not the ChangeIDs
    With ActiveSheet
        With .UsedRange
            If .AutoFilter Then .AutoFilter
            .AutoFilter field:=1, Criteria1:=word
            If word = "" Then .AutoFilter field:=2, Criteria1:=Array("<>*"), Operator:=xlAnd
            'With .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count)
            With .Offset(1, 0).Resize(.Rows.count - 1, LastColumnNum)
                If CBool(Application.Subtotal(103, .Cells)) Then
                    If r = "None" Then
                        '.SpecialCells(xlCellTypeVisible).EntireRow.Font.ColorIndex = 0
                        .SpecialCells(xlCellTypeVisible).Interior.ColorIndex = xlNone
                        If word = "Verification" Then .SpecialCells(xlCellTypeVisible).Font.ColorIndex = 3
                    Else:
                        '.SpecialCells(xlCellTypeVisible).EntireRow.Font.ColorIndex = 0
                        .SpecialCells(xlCellTypeVisible).Interior.Color = RGB(r, G, B)
                    End If
                End If
            End With
        .AutoFilter
        End With
    End With
End Sub

Public Sub AddChangeIDCode()
'Fills the last three colomns in the tables with values and formaulas to calculate Change IDs,
'based on the Type of Change in the first column
'M. Norelli 1/11/2019

    Dim c13, c21, c22, c31, c32, c33 As String
    Dim LastColNum As Long

    'LastColumn = "BX" 'as of Feb 2019
    LastColNum = Cells(1, Columns.count).End(xlToLeft).Column
    'put ChangeIDs in last blank column, to prevent overwriting ChangeIDs the builder may have personalized already
    'will therefore create endless numbers of ChangeID columns with repeated macro button presses...

    c13 = Cells(1, LastColNum + 3).Address(0, 0)
    c21 = Cells(2, LastColNum + 1).Address(0, 0)
    c22 = Cells(2, LastColNum + 2).Address(0, 0)
    c31 = Cells(3, LastColNum + 1).Address(0, 0)
    c32 = Cells(3, LastColNum + 2).Address(0, 0)
    c33 = Cells(3, LastColNum + 3).Address(0, 0)


    Range(c13).Value = "Change ID"
    Range(c31).Formula = "=IF(H3=""TLS"",H3,IF(ISBLANK(A3),""A"",IF(ISERR(FIND("" "",TRIM(A3))),upper(LEFT(A3,1)),upper(LEFT(A3,1)&MID(A3,FIND("" "",A3)+1,1)))))"
    Range(c32).Formula = "=MAXIFS(INDIRECT(""" & c22 & ":""&ADDRESS(ROW()-1,COLUMN())),INDIRECT(""" & c21 & ":""&ADDRESS(ROW()-1,COLUMN()-1))," & c31 & ")+1"
    Range(c33).Formula = "=" & c31 & "&" & c32
    Range(c33).Font.Bold = True

    Dim lRow As Long

    'Find the first blank cell in column F (after the two header rows)
    lRow = Range("F3:F" & Range("F1").End(xlDown).Row).count
    If lRow > 500 Then lRow = 500 'basic error-trapping if macro run on blank worksheet

    'Fill down to last non-blank row
    Range(c31).Resize(lRow, 3).FillDown

End Sub

Sub DummyMacro()

MsgBox "Yes, this button works!"

End Sub

Sub AddZeroes()
'reformats tower numbers, like "37/12" to the standard "037/012" format, and color those changes red
'M. Norelli 2/8/2019

    Dim cell As Range, word As Variant, words As Variant, reword As Variant
    Dim c%, num$, letterPos%, letter$, msg(0 To 1) As String, w%, x%
    Dim wordCount%, content$, newWord$, newContent$, rewordCount%
    Dim startend(0 To 1) As Integer

    For Each cell In Selection

        content = cell.Value
        newContent = ""

        word = Split(content)
        wordCount = UBound(word) - LBound(word)

        For w = 0 To wordCount

            If InStr(word(w), "/") > 0 Then
                words = Split(word(w), "/")
                For x = 0 To 1
                    num = words(x)
                    letterPos = FirstNonDigit(num)
                    If letterPos > 0 Then
                        letter = Mid(num, letterPos, 1)
                        If letterPos > 1 Then
                            num = Left(words(x), letterPos - 1)
                            msg(x) = Format(num, "000") & letter
                        End If
                        If letterPos = 1 Then
                            num = Mid(words(x), 2)
                            msg(x) = letter & Format(num, "000")
                        End If
                    Else
                        msg(x) = Format(num, "000")
                    End If
                Next

                newWord = msg(0) & "/" & msg(1)
            Else
                newWord = word(w)
            End If

            newContent = newContent & " " & newWord

        Next

        cell.Value = Trim(newContent)

        'color red the newly formatted tower numbers

        reword = Split(newContent)
        rewordCount = UBound(reword) - LBound(reword)

        For w = 0 To rewordCount
            If InStr(reword(w), "/") > 0 Then
                With cell.Characters(Start:=InStr(newContent, reword(w)) - 1, Length:=Len(reword(w))).Font
                    .Color = vbRed
                End With
            End If
        Next

    Next

End Sub

Function FirstNonDigit(xStr As Variant) As Long
'https://www.extendoffice.com/documents/excel/3790-excel-find-position-of-first-letter-in-string.html
'used to separate out tower numbers containing letters
    Dim xChar As Integer
    Dim xPos As Integer
    Dim i As Integer
    Application.Volatile
    For i = 1 To Len(xStr)
        xChar = Asc(Mid(xStr, i, 1))
        If xChar <= 47 Or _
           xChar >= 58 Then
            xPos = i
            Exit For
        End If
    Next
    FirstNonDigit = xPos
End Function

Function FindLastCol(header As String) As Integer
Dim LColNum&, LColValue$
LColNum = Worksheets("CAISO Update").Cells(1, Columns.count).End(xlToLeft).Column

While Worksheets("CAISO Update").Range(Col_Letter(LColNum) & "1").Value <> header
    LColNum = LColNum - 1
Wend

FindLastCol = LColNum

End Function
Function Col_Letter(lngCol As Long) As String
'https://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

'Version 1.0
'**************************************************************
'Dan Kaufman, Celerity Consulting Group, Inc.
'3/20/2019
'3/21/2019 edit M. Norelli: create tab if it doesn't exist, refactor copy-paste
'**************************************************************
Sub MakeSourceDocsRefTable()
    Dim wsTest As Worksheet

If ChangeIDsQC Then

    Set wsTest = Nothing
    On Error Resume Next
    Set wsTest = ActiveWorkbook.Worksheets("Source Doc Ref Tbl")
    On Error GoTo 0
     
'   Make sheet if it doesn't exist
    If wsTest Is Nothing Then
        Worksheets.Add.Name = "Source Doc Ref Tbl"
'   Move Source Doc Ref Tbl sheet after CAISO Update
        With Sheets("Source Doc Ref Tbl")
            .Move after:=Sheets("CAISO Update")
'   Format columns and first row to fit neatly when pasted to Word
            .Columns("A").ColumnWidth = 14.57
            .Columns("B").ColumnWidth = 15.43
            .Columns("C").ColumnWidth = 9.43
            .Columns("D").ColumnWidth = 44.86
            .Columns("E").ColumnWidth = 5.86
            .Rows(1).RowHeight = 25.5
'   Format first row text
            With .Range("A1:E1")
                .Font.Bold = True
                .Font.Size = 10
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .WrapText = True
            End With
        End With
    End If

'   Unmerge any merged cells (like Required Updates row) in CAISO Update, to prevent errors
    Dim cell As Range
    For Each cell In ThisWorkbook.Sheets("CAISO Update").UsedRange
        If cell.MergeCells Then
            cell.MergeCells = False
        End If
    Next
    
'   Copy Columns  Station Name, Component Description, Component Type

'    Sheets("CAISO Update").Select
'    Columns("F:H").Select
'    Selection.Copy
'    Sheets("Source Doc Ref Tbl").Select
'    Columns("A:A").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
    Worksheets("Source Doc Ref Tbl").Range("A:C").Value = Worksheets("CAISO Update").Range("F:H").Value

    
'   Copy column "Celerity Analysis Comments"

'    Sheets("CAISO Update").Select
'    Columns("BT:BT").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Sheets("Source Doc Ref Tbl").Select
'    Columns("D:D").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
    Worksheets("Source Doc Ref Tbl").Range("D:D").Value = Worksheets("CAISO Update").Range("BT:BT").Value

'   Copy column "Change ID"

'    Sheets("CAISO Update").Select
'    Columns("BY:BY").Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Sheets("Source Doc Ref Tbl").Select
'    Columns("E:E").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
    Worksheets("Source Doc Ref Tbl").Range("E:E").Value = Worksheets("CAISO Update").Range("BY:BY").Value

'   Search for "Required Updates" label

    Application.GoTo ActiveWorkbook.Sheets("Source Doc Ref Tbl").Cells(1, 1)
    ' this will fail if "Required Updates" is not found
'    Cells.Find(What:="Required Updates", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
'   See UpdateTopSection() for code that removes Required Updates row

  UpdateTopSection

End If

End Sub

Private Sub Clear_Source_Doc_Table()

'   Clear all the date from the Source Doc Ref Tbl. Keep the columns to retain the proper size and formatting.

    Sheets("Source Doc Ref Tbl").Select
    Columns("A:E").Select
    Selection.ClearContents
    Range("A1").Select

End Sub

Private Sub UpdateTopSection()
'Michael Norelli, Celerity Consulting Group, Inc.
' 3/20/2018
'This code assumes:
' - the active sheet is called "Source Doc Ref Tbl"
' - the Source Doc Ref Tbl has been created from a CAISO Update tab with all information
'   filled out, including Change IDs in Column BY
' - no CAISO Update rows are merged
' - useful data start on Row 3 (after the header row and the mostly blank "TL" row)
' - there is one blank cell in Column E that separates the top and bottom sections

Dim TopRange, MidRange, BotRange, SearchRange, lcell, Rng As Range
Dim topArray, botArray As Variant
Dim StartRow, Lastrow, f, t, l, col, FirstRowTop, LastRowTop, lRow As Integer
Dim ChangeVal$, TopVal$, foundReqUpd$

StartRow = "3"
Lastrow = Range("E" & StartRow).End(xlDown).Row
Set TopRange = Range("A" & StartRow & ":E" & Lastrow)

' checks for middle range of any number of blank or nearly blank rows
' between top and bottom sections.
' REQUIRED:  this works only while last column is blank for all rows between
' top and bottom sections
StartRow = Lastrow
Lastrow = Range("E" & StartRow).End(xlDown).Row

StartRow = Lastrow
Lastrow = Range("E" & StartRow).End(xlDown).Row
Set BotRange = Range("A" & StartRow & ":E" & Lastrow)

'Debug.Print BotRange.Cells.Count
'TO DO: include error checking for BotRange and TopRange containing rows

topArray = TopRange
botArray = BotRange

For f = 1 To UBound(botArray)                               'iterate through bottom range
    ChangeVal = botArray(f, 5)                              'look at fifth column
    
    If ChangeVal Like "U*" Or ChangeVal Like "TLS*" Then    ' ...until a "U" or "TLS" Change ID is found
        For t = 1 To UBound(topArray)                       'then in top range
            TopVal = topArray(t, 5)
            If ChangeVal = TopVal Then                      ' ...find a matching Change ID
                For col = 1 To 4
                    TopRange(t, col).Value = BotRange(f, col).Value ' ...and update the values in first four columns
                Next col
            End If
        Next t
    End If

Next f

BotRange.EntireRow.Delete

'   Remove Retired rows
'   https://www.rondebruin.nl/win/winfiles/MoreDeleteCode.txt
'       Set the first and last row to loop through
FirstRowTop = TopRange.Cells(1).Row
LastRowTop = TopRange.Rows(TopRange.Rows.count).Row
'       We loop from Lastrow to Firstrow (bottom to top)
For lRow = LastRowTop To FirstRowTop Step -1
'       We check the Change ID values in the E column
    With Cells(lRow, "E")
        If Not IsError(.Value) Then
            If .Value Like "R*" Then 'This will delete each row with the Change ID starting with "R"
                If Rng Is Nothing Then
                    Set Rng = .Cells
                Else
                    Set Rng = Application.Union(Rng, .Cells)
                End If
            End If
        End If
    End With
Next lRow
'       Delete all rows in one time
If Not Rng Is Nothing Then Rng.EntireRow.Delete

    
'   Remove "Required Updates" row
Dim NewLastrow%
NewLastrow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.count).Row
Set SearchRange = Range("A" & NewLastrow & ":E" & NewLastrow)
foundReqUpd = 0
For Each lcell In SearchRange
    'MsgBox lcell.Address & " " & lcell.Value
    If lcell.Value Like "*equired*" Then foundReqUpd = 1
Next
If foundReqUpd = 1 Then SearchRange.EntireRow.Delete

End Sub


Sub CompareAddIns()
Dim LocalPGE186AddIns As New Collection
Dim totalAddIns%, n%
Dim addinName$, AddInStorageLocation$, pattern$
Dim currentVersion$, availVersion$

' in network storage
AddInStorageLocation = "P:\PGE186\Code Tricks"
pattern = "Excel - PGE186 Tools v*.xlam"
LatestAddIn = LastFile(AddInStorageLocation, pattern)

' in current workbook
totalAddIns = Application.AddIns.count
For n = 1 To totalAddIns
    addinName = Application.AddIns(n).Name
    If addinName Like pattern And AddIns(n).Installed Then LocalPGE186AddIns.Add addinName
Next

If LatestAddIn = "" Then
    MsgBox "No PGE186 tools available to install from " & AddInStorageLocation & "."
Else
    availVersion = Version(LatestAddIn)

    If LocalPGE186AddIns.count = 0 Then
        currentVersion = 0
        LastInstalledAddIn = ""
        MsgBox "Installing version " & availVersion & "."
        Call InstallAddIn
    Else
        LastInstalledAddIn = LocalPGE186AddIns(LocalPGE186AddIns.count) '   assumes add-in list is sorted A-Z ascending
        currentVersion = Version(LastInstalledAddIn)
    
        If availVersion > currentVersion Then
            MsgBox "Update needed." & vbCrLf & "Installing new version, " & availVersion & "."
            Call InstallAddIn
        Else
            LatestAddIn = ""
            MsgBox "You have the latest version, " & currentVersion & ".  No update needed."
        End If
    End If
End If

End Sub
Sub InstallAddIn()
' https://andreilungu.com/how-to-automatically-install-and-activate-an-excel-addin-using-vba-code/
' Dependencies:
' Requires LatestAddIn global variable set to the filename (no path) to the latest
'  add-in stored in the network location.
' Requires LastInstalledAddIn global variable set to the filename (no path) to the latest
'  local PGE186 add-in stored in the workbook.
' PGE186 Add-ins are stored in P:\PGE186\Code Tricks\ and named in the format:
'  Excel - PGE186 Tools vN.xlam  where N = the version number as any length ASCII including dots
'  but not including spaces, which would throw off the sorting
Dim eai As Excel.AddIn
Dim toolpath$, filetoinstall$, addinName$, x$, file_to_copy$, folder_to_copy$, copied_file$
Dim pattern$, n%, d%

toolpath = "P:\PGE186\Code Tricks\"
addinName = Split(LatestAddIn, ".xlam")(0)
filetoinstall = "" & LatestAddIn

file_to_copy = toolpath & filetoinstall

folder_to_copy = Application.UserLibraryPath

copied_file = folder_to_copy & filetoinstall

'Check if add-in is in %APPDATA%\Microsoft\AddIns
If Len(Dir(copied_file)) = 0 Then

'if add-in does not exist then copy the file
FileCopy file_to_copy, copied_file
Set eai = Application.AddIns.Add(fileName:=copied_file)
eai.Installed = True
MsgBox "Add-in installed." & vbCrLf & "If you don't see version " & Split(addinName, "Excel - PGE186 Tools v")(1) & " tools, save, close, and re-open Excel."
'remove old add-in
pattern = LastInstalledAddIn
For d = 1 To Application.AddIns.count
    If AddIns(d).Name = pattern Then
        Debug.Print "Removed " & AddIns(d).Name
        AddIns(d).Installed = False
        Kill Application.UserLibraryPath & AddIns(d).Name
    End If
Next


Else

'if add-in already exists then the user will decide if will replace it or not
x = MsgBox("Add-in already exists! Replace?", vbYesNo)

    If x = vbNo Then
        Exit Sub
    ElseIf x = vbYes Then

        'deactivate the add-in if it is activated
        pattern = "Excel - PGE186 Tools v*.xlam"
        For n = 1 To Application.AddIns.count
            If AddIns(n).Name Like pattern And AddIns(n).Installed Then
                AddIns(n).Installed = False
            End If
        Next

        'delete the existing file
        Kill copied_file

        'copy the new file
        FileCopy file_to_copy, copied_file
        Set eai = Application.AddIns.Add(fileName:=copied_file)
        eai.Installed = True
        MsgBox "New Add-in installed."

    End If

End If

End Sub
Function LastFile(path, pattern) As String
Dim fname$

    fname = Dir(path & "\" & pattern)
    Do While fname <> ""
        LastFile = fname
        fname = Dir()
    Loop

End Function
Function Version(fileName) As String
Dim before$, after$

  before = "Excel - PGE186 Tools v"
  after = ".xl"
  Version = Left(Mid(fileName, Len(before) + 1), InStr(Mid(fileName, Len(before) + 1), after) - 1)
End Function

Private Sub whataddins()
Dim i%, x$
Debug.Print String(65535, vbCr)
For i = 1 To Application.AddIns.count
    x = "."
    If AddIns(i).Installed Then x = "Installed"
    Debug.Print AddIns(i).Name & ": " & x
Next
End Sub
Private Sub uninstallallPGE186()
Dim pattern$, i%
pattern = "Excel - PGE186 Tools v*.xlam"
Debug.Print String(65535, vbCr)
For i = 1 To Application.AddIns.count
    If AddIns(i).Name Like pattern Then
        AddIns(i).Installed = False
        If Len(Dir(Application.UserLibraryPath & AddIns(i).Name)) <> 0 Then
            Kill Application.UserLibraryPath & AddIns(i).Name
            Debug.Print "Removed " & AddIns(i).Name
        End If
    End If
Next
End Sub
Sub MakeSourceDocsUsedTable()
'Develops the data necessary for copying into the Source Documents Used table.
'REQUIRES:
' - Comments column should be called "Celerity Analysis Comments"
' - A delimiter of ";" is needed to split grouped comments
' - Rows with "TRANSMISSION" in the Station Name column will be deleted
' - Cleans, sorts, and replaces text in the reamining comments
' - All Change IDs filled in for each row containing data in Column A

    ClearCreateTable
    SplitCommentRows
    CleanCommentRows
    
End Sub


Sub ClearCreateTable()
'Creates a Source Docs Used table, if needed, and copies in comments column
'Do this in CAISO Update
'REQUIRES: CAISO Update tab must have a column of comments called "Celerity Analysis Comments"

Dim wsTest As Worksheet
Dim LastColumnNum%, Lastrow%

Worksheets("CAISO Update").Activate

Set wsTest = Nothing
On Error Resume Next
Set wsTest = ActiveWorkbook.Worksheets("Source Docs Used")
On Error GoTo 0

If wsTest Is Nothing Then
    Worksheets.Add.Name = "Source Docs Used"
    With Sheets("Source Docs Used")
        .Move after:=Sheets("CAISO Update")
    End With
End If

Sheets("Source Docs Used").Columns("A:B").Delete
Range("A1").Select

'LastColumn = BX, as of Feb 2019
LastColumnNum = FindLastCol("Celerity Analysis Comments")
Lastrow = Worksheets("CAISO Update").Range("f1").End(xlDown).Row

Worksheets("CAISO Update").Activate
Worksheets("Source Docs Used").Range("A1:A" & Lastrow - 2).Value = Worksheets("CAISO Update").Range("F3", Range("F1").End(xlDown)).Value
Worksheets("Source Docs Used").Range("B1:B" & Lastrow - 2).Value = Worksheets("CAISO Update").Range(Cells(3, LastColumnNum), Cells(Lastrow, LastColumnNum)).Value
Worksheets("Source Docs Used").Activate
End Sub
Sub SplitCommentRows()
Dim LR As Long, i As Long
Dim x As Variant, f%
Dim Delimiter$

Application.ScreenUpdating = False

Delimiter = ";"

LR = Range("B" & Rows.count).End(xlUp).Row

'Delete Transmission Line rows
For f = LR To 1 Step -1
    If InStr(1, Cells(f, 1), "TRANSMISSION", vbBinaryCompare) <> 0 Then Rows(f).Delete
Next

LR = Range("A" & Rows.count).End(xlUp).Row

Columns("B").Insert
For i = LR To 1 Step -1
    With Range("C" & i)
        If InStr(.Value, Delimiter) = 0 Then
            .Offset(, -1).Value = .Value
        Else
            x = Split(.Value, Delimiter)
            .Offset(1).Resize(UBound(x)).EntireRow.Insert
            .Offset(, -1).Resize(UBound(x) - LBound(x) + 1).Value = Application.Transpose(x)
        End If
    End With
Next i
Columns("C").Delete
LR = Range("B" & Rows.count).End(xlUp).Row
With Range("A1:B" & LR)
    On Error Resume Next
    .SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
    On Error GoTo 0
    .Value = .Value
End With
Application.ScreenUpdating = True
End Sub

Sub CleanCommentRows()
'With the comments broken up into a single column, performs the following clean up:
' - Trims extra spaces
' - Removes whole row that doesn't contain SLD|BOM|GAD
' - Removes text "Verified By " and "(typical to "
' - Erases parts of BOM comments after the drawing number
' - Replaces document abbreviations with full names
' - Sorts and finds unique values

Dim LR As Long, i As Long
Dim result$, Flag%
Dim Keep, k, CommentArray As Variant
LR = Range("A" & Rows.count).End(xlUp).Row

'Trim
For i = LR To 1 Step -1
    With Range("B" & i)
        .Value = Trim(.Value)
    End With
Next i


For i = LR To 1 Step -1
'Remove whole row that doesn't contain SLD|BOM|GAD
    Keep = Array("BOM", "GAD", "SLD")
    Flag = 1

    For Each k In Keep
        If InStr(1, Cells(i, 2), k, vbBinaryCompare) <> 0 Then Flag = 0
    Next

    If Flag = 1 Then
        Rows(i).Delete
    End If
Next

'Remove text: "Verified By ", etc.
result = ReplaceText("Verified by ", "")
result = ReplaceText("(typical to ", "")
'add for Nargis
result = ReplaceText("Rating verified via ", "")

'Add headers
Rows(1).Insert shift:=xlShiftDown
Range("A1").Value = "Station"
Range("B1").Value = "Document"

LR = Range("A" & Rows.count).End(xlUp).Row

'Sort
ActiveSheet.Sort.SortFields.Clear
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("A2"), Order:=xlAscending
     .SortFields.Add Key:=Range("B2"), Order:=xlAscending
     .SetRange ActiveSheet.Range("A1:B" & LR)
     .header = xlYes
     .Apply
End With
ActiveSheet.Sort.SortFields.Clear

'Remove from BOM lines anything after document number (pulls off Rev., Sheet, Item)
'Remove from GAD lines anything after revision number (pulls off Section, Elevation, Detail, Plan, etc.)
For i = LR To 1 Step -1
    With Range("B" & i)
    If InStr(1, .Value, "BOM ", vbBinaryCompare) = 1 Then
        CommentArray = Split(.Value, " ")
        .Value = CommentArray(0) & " " & CommentArray(1)
    End If
    If InStr(1, .Value, "GAD ", vbBinaryCompare) = 1 Then
        CommentArray = Split(.Value, " ")
        .Value = CommentArray(0) & " " & CommentArray(1) & " " & CommentArray(2) & " " & CommentArray(3)
    End If
    End With
Next i

'Unique
With Worksheets("Source Docs Used")
    .Range("F:G").Delete
    .Range("A1:B" & LR).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=.Range("F:G"), Unique:=True
    .Range("A:E").Delete
    .Rows(1).Delete
    .Columns("A:B").AutoFit
End With

'Change BOM to "Bill of Materials, Dwg"
'Change GAD to "General Arrangement Diagram, Dwg"
'Change SLD to "Single Line Diagram, Dwg"
result = ReplaceText("BOM", "Bill of Materials, Dwg")
result = ReplaceText("GAD", "General Arrangement Diagram, Dwg")
result = ReplaceText("SLD", "Single Line Diagram, Dwg")


End Sub

Function ReplaceText(str As String, repl As String) As String
'Replace text throughout one column
Dim UsedRangeCol$, ReplaceRangeCol$
Dim Lastrow%, r%
UsedRangeCol = "A"
ReplaceRangeCol = "B"

Lastrow = ActiveSheet.Range(UsedRangeCol & Rows.count).End(xlUp).Row

For r = 1 To Lastrow
    With Range(ReplaceRangeCol & r)
        .Value = Replace(.Value, str, repl, , , vbBinaryCompare)
    End With
Next

ReplaceText = "Done"

End Function

Sub MakeRatingsRequestedTable()
 Call MakeTable("Ratings Requested", "Verification", "A", Array("C", "F", "G", "H", "BT"))
End Sub

Sub MakeEquipAddedTable()
 Call MakeTable("Equipment Added", "Create", "A", Array("F", "G", "H"))
End Sub
Sub MakeEquipRetiredTable()
 Call MakeTable("Equipment Retired", "Retire", "A", Array("F", "G", "H"))
End Sub
Sub MakeRelayRequestTable()
 Call MakeTable("TLS with Relay Request", "RLY", "H", Array("C", "F", "G", "H", "O", "Q", "R", "T", "U"))
End Sub
Sub MakeTable(NewSheet As String, LookingFor As String, LookInCAISOCol As String, GetCAISOCols As Variant)
'Creates table suitable for pasting into Summary Report
'of all rows marked for verification where:
'NewSheet = name of the new tab storing the data created
'LookingFor = text to search for in each row
'LookInCAISOCol = column in the CAISO Update tab to search for the LookingFor text
'GetCAISOCols = list of columns to pull from CAISO Update to make the table


Dim TopRange As Range
Dim topArray, c As Variant
Dim StartRow, Lastrow, NewSheetCol, NewSheetRow, f As Integer
Dim CurrentVal$
Dim wsTest As Worksheet

'   Set sheet name and what parts of CAISO Update sheet should be summarized for it
'   Make sheet if it doesn't exis
Set wsTest = Nothing
On Error Resume Next
Set wsTest = ActiveWorkbook.Worksheets(NewSheet)
On Error GoTo 0

If wsTest Is Nothing Then
    Worksheets.Add.Name = NewSheet
    With Sheets(NewSheet)
        .Move after:=Sheets("CAISO Update")
    End With
End If

Sheets(NewSheet).Columns("A:BY").Delete

'Copy out neeeded data from CAISO Update
Worksheets("CAISO Update").Activate

StartRow = "2"
Lastrow = Range("F" & StartRow).End(xlDown).Row
Set TopRange = Range(LookInCAISOCol & StartRow & ":" & LookInCAISOCol & Lastrow)

topArray = TopRange
NewSheetRow = 1
NewSheetCol = 1

'Add headers
For Each c In GetCAISOCols
    Worksheets(NewSheet).Cells(NewSheetRow, NewSheetCol).Value = Worksheets("CAISO Update").Range(c & "1").Value
    NewSheetCol = NewSheetCol + 1
Next
NewSheetRow = NewSheetRow + 1

For f = 1 To UBound(topArray)
    CurrentVal = topArray(f, 1)
    If CurrentVal Like LookingFor Then
        NewSheetCol = 1
        For Each c In GetCAISOCols
            Worksheets(NewSheet).Cells(NewSheetRow, NewSheetCol).Value = Worksheets("CAISO Update").Range(c & f + StartRow - 1).Value
            NewSheetCol = NewSheetCol + 1
        Next
        NewSheetRow = NewSheetRow + 1
    End If
Next f

Worksheets(NewSheet).UsedRange.Columns.AutoFit
Worksheets(NewSheet).Activate
End Sub
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Function ChangeIDsQC() As Boolean
'Check Change ID column vs. other columns to ensure all values filled in
Dim BotRange As Range
Dim topArray, botArray As Variant
Dim StartRow, Lastrow, B As Integer
Dim ChangeIDVal$, ChangeIDCol$
Dim arr As Variant
Dim resultArr(0 To 2) As Long, x%, maxBlanks%

On Error GoTo handler

StartRow = "3"
ChangeIDCol = "BY"
arr = Array("F", "G", "H")
    
With Worksheets("CAISO Update")
    .Activate
    
    'Count blanks in Station Name, Component Type, and Component Desc columns and compare the top two
    For x = 0 To 2
        resultArr(x) = .Range(arr(x) & ":" & arr(x)).Cells.SpecialCells(xlCellTypeBlanks).count
    Next
    
    'stop if the two maximum blank counts are not equal, meaning one of those columns is incomplete
    If WorksheetFunction.Large(resultArr, 1) <> WorksheetFunction.Large(resultArr, 2) Then _
    Err.Raise ERROR_BLANK_CHANGEID, "CheckChangeID", "There are empty cells in the Station Name, Component Type, or Component Desc. columns.  Please fill them before continuing."
    
    'Count blanks in Change ID column and stop if the blanks minus one (skipping the normally blank "BY2" cell) doesn't equal the count of the other columns, meaning a Change ID is missing
    If WorksheetFunction.Large(resultArr, 1) <> .Range(ChangeIDCol & ":" & ChangeIDCol).Cells.SpecialCells(xlCellTypeBlanks).count - 1 Then _
    Err.Raise ERROR_BLANK_CHANGEID, "CheckChangeID", "The count of values between the Station Name, Component Type, Component Desc., and Change Description column." & vbCrLf & "Please check these columns and fill any blank values before continuing."
    
    'get all ChangeID values up to the first blank into an array, hopefully representing the top section
    Lastrow = Range(ChangeIDCol & StartRow).End(xlDown).Row
    topArray = Application.Transpose(Range(ChangeIDCol & StartRow & ":" & ChangeIDCol & Lastrow).Value)
    
    'skip first blank value, hopefully the border between top and bottom sections
    StartRow = Lastrow
    Lastrow = Range(ChangeIDCol & StartRow).End(xlDown).Row
    StartRow = Lastrow
    
    'get the remaining Change IDs, hopfully the bottom section
    Lastrow = Range(ChangeIDCol & StartRow).End(xlDown).Row
    Set BotRange = Range(ChangeIDCol & StartRow & ":" & ChangeIDCol & Lastrow)
    
    botArray = BotRange
    
    'stop if top section is smaller than bottom section
    If UBound(topArray) < UBound(botArray) Then _
    Err.Raise ERROR_BLANK_CHANGEID, "CheckChangeID", "Please complete all ChangeIDs before generating the Source Docs Ref Table."
    
    'stop if a bottom section value does not exist in top section
    For B = 1 To UBound(botArray)
        ChangeIDVal = botArray(B, 1)
        If Not IsInArray(ChangeIDVal, topArray) Then _
        Err.Raise ERROR_BLANK_CHANGEID, "CheckChangeID", "Please complete all ChangeIDs before generating the Source Docs Ref Table."
    Next B

End With

Done:
    ChangeIDsQC = True
    Exit Function
    
handler:
    MsgBox Err.Description
    ChangeIDsQC = False
    
End Function
