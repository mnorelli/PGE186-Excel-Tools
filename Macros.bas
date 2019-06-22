Attribute VB_Name = "Macros"
'Version: 2.3

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
'2.3
'- Changes Change ID numbers to fill down only to last row of top section
'- Adds process to remove old add-in files from user directory when updating
'- In Change Descriptions, changes DISC to SW for transmission line switches
'2.2
'- Corrects Change Descriptions to look for TLSs first before concatenating other fields
'2.1
'- Adds TLS numbers for Change Descriptions of transmission line rows
'- Adds TLS numbers at bottom from TLS of row in top section with same OID
'- For Create rows, adds Change Description at bottom from row in top section with same Component Description
'- Adds public variable for splitrow between top and bottom sections
'- Removes trailing commas from document references to fix problem in making Source Docs Used table
'- In Source Doc Used processing, skips doc name parts found containing a semicolon
'- Sorts Source Docs Used table by type of doc (descending), substation, then by numeric doc number
'2.0
'- Refactors to create top and bottom sections once as public 2D objects
'- Refactors MakeTable to use public top array
'- Extends to bottom section the adding of the new Change Description
'- Adds two-row Change Descriptions for rows whose Description contains a space
'- Checks top section for duplicate Change Descriptions
'- Alerts if no tower numbers are found when Update Tower Numbers is run
'- Refactors Source Doc Ref Table to copy faster, refactor to use OID instead of Change ID, and delete Change Description column
'- Deletes Change Description column from Source Doc Ref Table
'- Removes unneeded routines
'- Adds formatting to Info Request tables
'1.9
'- Implements change to Change Descriptions from numbers to descriptions, checking for identical Change Descriptions, and displaying new button
'- Refactors Color Row by Change Type to use conditional formatting dynamically instead of static background change
'- Edits Source Docs Used sorting
'1.8
'- Corrects error in Source Docs Used sorting
'- Changes alignment in Source Docs Ref Table, Ratings Requested, Equipment Added, and Equipment Retired.
'1.7
'- Allows Ratings Requested table to collect Create rows in red text that need Verification
'- Info Requests show only the rows appropriate for the kind of request
'- Makes 9px font default for table exports
'1.6
'- Adds new error checking for creating top and bottom sections
'- Checks for blank or missing ChangeIDs
'- Adds better error text
'- Fixes Ratings Requested tool to pull Additional Info instead of Celerity Comment
'- Adds tools for selecting and formatting info for Information Requests
'- Arranges Source Docs Used according to current template
'- Adds a High Rating column to Source Docs Ref Table, to help finding most-limiting component
'- Allows Ratings Requested table to collect Create rows in red text that need Verification
'1.5
'- Swaps Source Docs macros button names to correctly run the applicable macro
'- Renames Create_Source_Doc_Table to MakeSourceDocsRefTable
'- Renames CreateRatingsRequestedTable to MakeRatingsRequestedTable
'- Abstracts away table creation to a general sub receving passed parameters and changes search for "=" to "Like" to allow for more search flexibility
'- Adds buttons and code to make tables for Relays, Equipment Added, Equipment Retired
'- Changes order of buttons to make tables in match Word doc order
'- Checks for Change IDs as a prerequisite for running MakeSourceDocsRefTable
'- Prevents Highlight macro from changing colors in legend at bottom below data rows
'1.4
'- Adds tool to make Source Documents Used table
'- Adds tool to make Ratings Requested table
'1.3
'- Removes "Retired" rows from SourceDocRefTable.
'- Rearranges tools into three groups.
'- Adds ability to update the add-in to the current version in P:\PGE186\Code Tricks
'Previous
'- Adds Dan's code to create Source Doc Ref Table
'- Adds code to move update rows to main table for Source Doc Ref Table
'- Removes the general FindDistinctSubstrings() macro and replaced with the more specific AddZeroes()
'- Adds code to find last Date field for Paint()
'**************************************************************

Option Explicit
Public Const ERROR_BLANK_CHANGEID As Long = vbObjectError + 513
Public Const ERROR_TOP_BOTTOM_ARRAY As Long = vbObjectError + 5131
Public LatestAddIn As String
Public LastInstalledAddIn As String
Public topArray As Variant
Public botArray As Variant
Public splitrow As Variant
Public TopRange As Range
Public BotRange As Range

Public Sub Paint()
'https://stackoverflow.com/questions/29085029/excel-macro-change-the-row-based-on-value
' This macro looks for the Type of Change values in the first column of the active spreadsheet
' and sets the background and font to the indicated color for the row containing the indicated text.
' "None" clears font and background, "Red" makes text red.
' Designed to make set up for PGE186 T-Line Line Data Sheets easier to set up
'M. Norelli 1/11/2019
Application.Sheets("CAISO Update").Activate
Application.ScreenUpdating = False
Sheets("CAISO Update").Cells.FormatConditions.Delete

    Call CondFormat("=$A2=""Create""", 142, 169, 219)
    Call CondFormat("=$A2=""Update""", 244, 176, 132)
    Call CondFormat("=$A2=""Retire""", 255, 255, 0)
    Call CondFormat("=$A2=""Previously Submitted""", 146, 208, 80)
    With Sheets("CAISO Update").Range("A2:BY200")
        'red top and bottom borders
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=CELL(""row"")=ROW()"
        With .FormatConditions(.FormatConditions.count)
            .StopIfTrue = False
            .SetFirstPriority
            With .Borders(xlTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbRed
            End With
            With .Borders(xlBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbRed
            End With
        End With
        'red text
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$A2=""Verification"""
        With .FormatConditions(.FormatConditions.count)
            .StopIfTrue = False
            .SetFirstPriority
            .Font.Color = vbRed
        End With
    End With

Application.ScreenUpdating = True
End Sub

Private Sub CondFormat(cond As String, r As Integer, g As Integer, b As Integer)
'Changes conditional formatting based on the Type of Change entry and RGB values passed from the Paint() macro
'M. Norelli 1/11/2019

    With Sheets("CAISO Update").Range("A2:BY200")
        .FormatConditions.Add Type:=xlExpression, Formula1:=cond
        With .FormatConditions(.FormatConditions.count)
            .StopIfTrue = False
            .SetFirstPriority
            With .Interior
                .PatternColorIndex = xlAutomatic
                .Color = RGB(r, g, b)
                .TintAndShade = 0
            End With
        End With
    End With

End Sub


Public Sub AddChangeDescCode()
'Fills the last column in the table with a Change description concatenated
' from Component Type and Component Descrition
Dim cTypeT$, cDescT$, cTypeB$, cDescB$, cTypeCol$, cDescCol
Dim arr As Variant, a, b As Variant, count
Dim BotFirstRow%, StartRow%, ReportCell As Range
Dim ChangeDescCol$, ChangeDescColNum%, ChangeTopRange As Variant, ChangeBotRange As Variant
Dim c13, c21, c22, c31, c32, c33 As String
Dim StationNameColNum, OIDColNum
Dim cStationT, cStationB, cChangeUp1, cChangeLeft4, cOID, rTopOIDs, rBotOIDs, AllChangeCol
Dim CompDescColNum, rTopCompDesc

    StartRow = 3  'First row to start processing
    cTypeCol = "H"  'Component Type found in column
    cDescCol = "G"  'Component Description found in column
    ChangeDescCol = "BY"  'where to write new Change Descriptions

    'Create formula to concatenate Component Type and Description
    ChangeDescColNum = FindLastCol("Change Description")
    StationNameColNum = FindLastCol("Station Name")
    OIDColNum = FindLastCol("OID")
    CompDescColNum = FindLastCol("Component Description")


    cTypeT = Cells(StartRow, cTypeCol).Address(0, 0)
    cDescT = Cells(StartRow, cDescCol).Address(0, 0)
    cStationT = Cells(StartRow, StationNameColNum).Address(0, 0)
    cChangeUp1 = Cells(StartRow, ChangeDescColNum).Offset(-1, 0).Address(0, 0)
    cChangeLeft4 = Cells(StartRow, ChangeDescColNum).Offset(0, 4).Address(0, 0)

    If MakeTopBottomArray("CAISO Update", ChangeDescCol, StartRow) Then

        '***********************  ChangeIDs
        '*
        'put ChangeIDs in three columns after Change Description column + 1.  Multiple runs will overwrite, not append new columns
        c13 = Cells(1, ChangeDescColNum + 4).Address(0, 0)
        c21 = Cells(2, ChangeDescColNum + 2).Address(0, 0)
        c22 = Cells(2, ChangeDescColNum + 3).Address(0, 0)
        c31 = Cells(3, ChangeDescColNum + 2).Address(0, 0)
        c32 = Cells(3, ChangeDescColNum + 3).Address(0, 0)
        c33 = Cells(3, ChangeDescColNum + 4).Address(0, 0)

        Range(c13).value = "Change ID"
        Range(c31).Formula = "=IF(H3=""TLS"",H3,IF(ISBLANK(A3),""A"",IF(ISERR(FIND("" "",TRIM(A3))),upper(LEFT(A3,1)),upper(LEFT(A3,1)&MID(A3,FIND("" "",A3)+1,1)))))"
        Range(c32).Formula = "=MAXIFS(INDIRECT(""" & c22 & ":""&ADDRESS(ROW()-1,COLUMN())),INDIRECT(""" & c21 & ":""&ADDRESS(ROW()-1,COLUMN()-1))," & c31 & ")+1"
        Range(c33).Formula = "=" & c31 & "&" & c32
        Range(c33).Font.Bold = True

        Range(c31).Resize(splitrow - StartRow, 3).FillDown
        '*
        '***********************   ChangeIDs

        Set ChangeTopRange = TopRange.Columns(ChangeDescColNum)
        Set ChangeBotRange = BotRange.Columns(ChangeDescColNum)

        BotFirstRow = Split(BotRange.Cells.Item(1).Address(1, 0), "$")(1)
        cTypeB = Cells(BotFirstRow, cTypeCol).Address(0, 0)
        cDescB = Cells(BotFirstRow, cDescCol).Address(0, 0)
        cStationB = Cells(BotFirstRow, StationNameColNum).Address(0, 0)
        cOID = Cells(BotFirstRow, OIDColNum).Address(0, 0)
        rTopOIDs = TopRange.Columns(OIDColNum).Address(1, 0)
        rBotOIDs = BotRange.Columns(OIDColNum).Address(1, 0)
        rTopCompDesc = TopRange.Columns(CompDescColNum).Address(1, 0)

        'To implement:
        '
        ' TOP SECTION
        '
        '=IF(LEFT(F20,6)="[TRANS",       'if this is a transmission line
        '    IF(LEFT(CC20,2)<>"TL",      'if this is not a TLS
        '                                'construct a description from these pieces:
        '        IF(H20="DISC","SW",H20) 'first, replace SW for DISC type, where present
        '            &MID(BY19,FIND(CHAR(10),BY19),5)&" - "&G20,   'and add the TLS number
        '                            'from the cell above with a dash and the description
        '            H20&CHAR(10)&CC20&" - "&G20),  'if it is a TLS
        '            'use the type ("TLS") and on a new line, the TLS's change ID number
        '            'plus the description
        '            IF(ISERROR(FIND(" ",G20)),H20&" "&G20,H20&CHAR(10)&G20))
        '            'if it's inside a sub (not transmission line feature), then assemble
        '            'the usual elements in the usual way, adding a line feed if the
        '            'description has a space
        '
        'Replace:
        '    cStationT for F20
        '    cChangeLeft4 for CC20
        '    cTypeT for H20
        '    cDescT for G20
        '    cChangeUp1 for BY19
        '
        '
        ' BOTTOM SECTION
        '
        '=IF(LEFT(F75,6)="[TRANS",   'if this is a transmission line
        '   IFERROR(IF(H75="DISC","SW",H75)&TRIM(MID(INDEX($BY$3:$BY$55,MATCH(C75,C$3:C$55,0)),
        '       FIND(CHAR(10),INDEX($BY$3:$BY$55,MATCH(C75,C$3:C$55,0))),5))&" - "&G75,
        '           'use the OID to get the TLS number from the text from the top section and
        '           ' add it to the description (to allow for description updates)
        '       INDEX($BY$3:$BY$55,MATCH(G75,G$3:G$55,0))),
        '           'otherwise use the description to get the whole text from the top section
        '           ' (there is no OID, it's a Create row, ad the description should be identical)
        '    IF(ISERROR(FIND(" ",G75)),H75&" "&G75,H75&CHAR(10)&G75))
        '          'if it's not a transmission line, make the in the usual way
        '
        ' Replace as above, and:
        '       ChangeTopRange.Address for $BY$3:$BY$55
        '       cOID for C75
        '       cStationB for F75
        '       cTypeB for H75
        '       rTopOIDs for C$3:C$55
        '       cDescB for G75
        '       rTopCompDesc for G$3:G$55
        '       cChangeUp1 Not used

        ChangeTopRange.value = "=IF(LEFT(" & cStationT & ",6)=""[TRANS"",IF(LEFT(" & cChangeLeft4 & ",2)<>""TL"",IF(" & cTypeT & "=""DISC"",""SW""," & cTypeT & ")&MID(" & cChangeUp1 & ",FIND(CHAR(10)," & cChangeUp1 & "),5)&"" - ""&" & cDescT & "," & cTypeT & "&CHAR(10)&" & cChangeLeft4 & "&"" - ""&" & cDescT & "),IF(ISERROR(FIND("" ""," & cDescT & "))," & cTypeT & "&"" ""&" & cDescT & "," & cTypeT & "&CHAR(10)&" & cDescT & "))"
        ChangeBotRange.value = "=IF(LEFT(" & cStationB & ",6)=""[TRANS"",IFERROR(IF(" & cTypeB & "=""DISC"",""SW""," & cTypeB & ")&TRIM(MID(INDEX(" & ChangeTopRange.Address & ",MATCH(" & cOID & "," & rTopOIDs & ",0)),FIND(CHAR(10),INDEX(" & ChangeTopRange.Address & ",MATCH(" & cOID & "," & rTopOIDs & ",0))),5))&"" - ""&" & cDescB & ",INDEX(" & ChangeTopRange.Address & ",MATCH(" & cDescB & "," & rTopCompDesc & ",0))),IF(ISERROR(FIND("" ""," & cDescB & "))," & cTypeB & "&"" ""&" & cDescB & "," & cTypeB & "&CHAR(10)&" & cDescB & "))"
        AllChangeCol = ChangeTopRange.Cells.Item(1).Address & ":" & ChangeBotRange.Cells.Item(ChangeBotRange.Cells.count).Address
        With Sheets("CAISO Update").Range(ChangeDescCol & ":" & ChangeDescCol).Columns
            .AutoFit
            .WrapText = True
        End With

    'check for duplicates
        arr = ChangeTopRange
        For a = 1 To UBound(arr)
            count = 0
            For Each b In arr
                If IsError(arr(a, 1)) = False And IsError(b) = False Then
                    If LCase(arr(a, 1)) = LCase(b) Then count = count + 1
                End If
            Next
            Set ReportCell = ChangeTopRange.Offset(0, 1)
            'clear previous reports while preserving other user-entered data in this column
            If ReportCell.Cells.Item(a).value = "<--duplicate" Then ReportCell.Cells.Item(a).value = ""
            If count > 1 Then
                With ReportCell.Cells.Item(a)
                    .value = "<--duplicate"
                    .Font.Color = vbRed
                End With
            End If
        Next
    End If
End Sub

Sub AddZeroes()
'reformats tower numbers, like "37/12" to the standard "037/012" format, and color those changes red
'M. Norelli 2/8/2019

    Dim cell As Range, word As Variant, words As Variant, reword As Variant
    Dim c%, num$, letterPos%, letter$, msg(0 To 1) As String, w%, x%
    Dim wordCount%, content$, newWord$, newContent$, rewordCount%
    Dim startend(0 To 1) As Integer

    If Selection.Find("/") Is Nothing Then
        MsgBox "No tower numbers found in the current selecton."
        Exit Sub
    End If

    For Each cell In Selection

        content = cell.value
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

        cell.value = Trim(newContent)

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
Dim FindCol

    Set FindCol = Sheets("CAISO Update").UsedRange.Rows(1).Find(header, lookat:=xlPart)
    If Not FindCol Is Nothing Then
        FindLastCol = FindCol.Column
    Else
        MsgBox "Column " & header & " not found."
    End If


End Function

Function Col_Letter(lngCol As Integer) As String
'https://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter
    Col_Letter = Split(Cells(1, 81).Address(1, 0), "$")(0)
End Function

'**************************************************************
'Dan Kaufman, Celerity Consulting Group, Inc.
'3/20/2019
'3/21/2019 edits by M. Norelli
'**************************************************************
Sub MakeSourceDocsRefTable()
    Dim wsTest As Worksheet
    Dim arr As Variant, Dest As Range
    Dim ColumnLtrs As Variant, ColumnNums() As Long
    Dim col As Long, SliceStr$, Separator$

'If ChangeIDsQC Then

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
            '.Columns("E").ColumnWidth = 5.86
            .Rows(1).RowHeight = 25.5
        End With
    End If

    Sheets("Source Doc Ref Tbl").UsedRange.ClearContents

'   Unmerge any merged cells (like Required Updates row) in CAISO Update, to prevent errors
'    Dim cell As Range
'    For Each cell In ThisWorkbook.Sheets("CAISO Update").UsedRange
'        If cell.MergeCells Then
'            cell.MergeCells = False
'        End If
'    Next
'
    'Columns needed to pull from CAISO Update tab
    ColumnLtrs = Array("F", "G", "H", "BT", "C", "R", "A")

    ReDim ColumnNums(UBound(ColumnLtrs)) As Long
    For col = 0 To UBound(ColumnLtrs)
        ColumnNums(col) = Range(ColumnLtrs(col) & 1).Column
    Next

    arr = Sheets("CAISO Update").Range("A1").CurrentRegion
    With Worksheets("Source Doc Ref Tbl")
        Set Dest = .Range("A1")
        'https://usefulgyaan.wordpress.com/2013/06/12/vba-trick-of-the-week-slicing-an-array-without-loop-application-index/
        '(go to last comment)
        'The first dimension is the number of rows and the second dimension is the number of columns
        'Usage: Destination Range, sized to all CAISO rows by ColumnLtrs columns equals
        '    the whole CAISO Range, all the rows, using the columns specified in ColumnLtrs
        Dest.Resize(UBound(arr, 1), UBound(ColumnLtrs) + 1) = _
            Application.Index(arr, Evaluate("=" & "Row(1:" & UBound(arr, 1) & ")"), ColumnNums)
        .Activate
    End With

  UpdateTopSection

  With Sheets("Source Doc Ref Tbl")
'   Format all rows
    With .Range("A:D")
        .Font.Bold = False
        .Font.Size = 9
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With
'   Format first row text
    With .Range("A1:D1")
        .Font.Bold = True
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
  End With

'End If

End Sub

Private Sub UpdateTopSection()
'This code assumes:
' - the active sheet is called "Source Doc Ref Tbl"
' - the Source Doc Ref Tbl has been created from a CAISO Update tab with all information
'   filled out, including Change IDs in Column BY
' - an OID column exists to allow updating to occur

Dim SearchRange, lcell, rng As Range
Dim f, t, col, FirstRowTop, LastRowTop, lRow As Integer
Dim ChangeVal$, TopVal$, foundReqUpd$
Dim FindCol As Range, TypeOfChangeCol, n, Dest As Range

If MakeTopBottomArray("Source Doc Ref Tbl", "F", 2) Then

    For f = 1 To UBound(botArray, 2)                              'iterate through bottom range
        ChangeVal = botArray(5, f)                              'look at fifth column
        If ChangeVal <> "" Then

            For t = 1 To UBound(topArray, 2)                      'then in top range
                TopVal = topArray(5, t)
                If ChangeVal = TopVal Then                      ' ...find a matching Change ID
                    'Debug.Print "found " & TopVal
                    For col = 1 To UBound(topArray, 1)       ' ...update each column
                        TopRange(t, col).value = BotRange(f, col).value ' ...and update the values in first four columns
                    Next col
                End If
            Next t
        End If

    Next f

    BotRange.EntireRow.Delete

    '   Remove Retired rows
    '   https://www.rondebruin.nl/win/winfiles/MoreDeleteCode.txt
    '       Set the first and last row to loop through
    FirstRowTop = TopRange.Cells(1).row
    LastRowTop = TopRange.Rows(TopRange.Rows.count).row

    ' find "Type of Change" column
    Set FindCol = Sheets("Source Doc Ref Tbl").UsedRange.Rows(1).Find("Type of Change", lookat:=xlPart)
    If Not FindCol Is Nothing Then
        TypeOfChangeCol = FindCol.Column
    End If

    '       We loop from Lastrow to Firstrow (bottom to top)
    For lRow = LastRowTop To FirstRowTop Step -1
    '       We check the Change ID values in the E column
        With Cells(lRow, TypeOfChangeCol)
            If Not IsError(.value) Then

                If .value = "Retire" Then 'This will delete each Retired
                    If rng Is Nothing Then
                        Set rng = .Cells
                    Else
                        Set rng = Application.Union(rng, .Cells)
                    End If
                End If
            End If
        End With
    Next lRow
    '       Delete all rows in one time
    If Not rng Is Nothing Then rng.EntireRow.Delete


    '   Remove "Required Updates" row
    Dim NewLastrow%
    NewLastrow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.count).row
    Set SearchRange = Range("A" & NewLastrow & ":E" & NewLastrow)
    foundReqUpd = 0
    For Each lcell In SearchRange
        'MsgBox lcell.Address & " " & lcell.Value
        If lcell.value Like "*equired*" Then foundReqUpd = 1
    Next
    If foundReqUpd = 1 Then SearchRange.EntireRow.Delete

    ' clear Type of Change and OID columns
    Columns(TypeOfChangeCol).ClearContents
    Set FindCol = Sheets("Source Doc Ref Tbl").UsedRange.Rows(1).Find("OID", lookat:=xlPart)
    If Not FindCol Is Nothing Then
        Columns(FindCol.Column).ClearContents
    End If
    ' fill with blanks to show clear separation of High Rating from main table
    With Worksheets("Source Doc Ref Tbl")
        Set Dest = .Range("E1")
        'The first dimension is the number of rows and the second dimension is the number of columns
        'Usage: Destination Range, sized to all CAISO rows by ColumnLtrs columns equals
        '    the whole CAISO Range, all the rows, using the columns specified in ColumnLtrs
        Dest.Resize(UBound(topArray, 2), 1) = " "
        .Activate
    End With

End If      ' MakeTopBottomArray

End Sub

Sub UpdatePGETools()
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

'Compare PGETools add-ins
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
    'delete old versions
    DeleteAllButLast ("Excel - PGE186 Tools v*.xlam")

Else  'if add-in already exists then the user will decide if will replace it or not
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
Sub zztest()
  Debug.Print DeleteAllButLast("Excel - PGE186 Tools v*.xlam")
End Sub

Function DeleteAllButLast(pattern As String) As Integer

Dim fname$, aFiles() As Variant
Dim fs, f, fc, s, PGEfile, count As Long

Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(Application.UserLibraryPath)
Set fc = f.Files

'fill array with filenames
count = 1
For Each PGEfile In fc
    If PGEfile.Name Like pattern Then
        ReDim Preserve aFiles(count)
        aFiles(count) = PGEfile.Name
       count = count + 1
    End If
Next

'drop last item
ReDim Preserve aFiles(UBound(aFiles) - 1)

'delete all remaining
For s = 1 To UBound(aFiles)
    On Error Resume Next
    Kill (f & "\" & aFiles(s))
    DeleteAllButLast = s
    On Error GoTo 0
Next

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
'for testing only, incorporated in InstallAddIn sub
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
On Error Resume Next
Kill Application.UserLibraryPath & pattern
On Error GoTo 0
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
    ParseCommentRows
    SplitCommentRows
    CleanCommentRows

End Sub

Sub ClearCreateTable()
'Creates a Source Docs Used table, if needed, and copies in comments column
'Do this in CAISO Update
'REQUIRES: CAISO Update tab must have a column of comments called "Celerity Analysis Comments"

Dim wsTest As Worksheet
Dim LastColumnNum%, LastRow%

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
LastRow = Worksheets("CAISO Update").Range("f1").End(xlDown).row

Worksheets("CAISO Update").Activate
Worksheets("Source Docs Used").Range("A1:A" & LastRow - 2).value = Worksheets("CAISO Update").Range("F3", Range("F1").End(xlDown)).value
Worksheets("Source Docs Used").Range("B1:B" & LastRow - 2).value = Worksheets("CAISO Update").Range(Cells(3, LastColumnNum), Cells(LastRow, LastColumnNum)).value
Worksheets("Source Docs Used").Activate
End Sub

Sub ParseCommentRows()
'For each cell in a column of Comments, finds any of SLD|BOM|GAD and assembles a new value from
'that text and a number of words after it, removing anything that is not a drawing reference
Dim c%, num$, letterPos%, letter$, msg(0 To 1) As String, w%, x%
Dim content$, newWord$, newContent$, rewordCount%
Dim startend(0 To 1) As Integer
Dim r, cell, word As Variant, s
Dim LastRow

LastRow = Worksheets("Source Docs Used").Range("a1").End(xlDown).row
Set r = Range("B1:B" & LastRow)

    For Each cell In r

        content = cell.value
        newContent = ""
        newWord = ""

        word = Split(content)
        For w = 0 To UBound(word)
            If word(w) = "SLD" Then
                newWord = "SLD"
                For s = 1 To 3
                    newWord = newWord & " " & word(w + s)
                    If InStr(word(w + s), ";") > 0 Then s = 3
                Next
                If InStr(Len(newWord), newWord, ".") = Len(newWord) Then newWord = Left(newWord, Len(newWord) - 1)
                If InStr(Len(newWord), newWord, ";") = 0 Then newWord = newWord & ";"
            ElseIf word(w) = "GAD" Then
                newWord = "GAD"
                For s = 1 To 3
                    newWord = newWord & " " & word(w + s)
                    If InStr(word(w + s), ";") > 0 Then s = 3
                Next
                If InStr(Len(newWord), newWord, ".") = Len(newWord) Then newWord = Left(newWord, Len(newWord) - 1)
                If InStr(Len(newWord), newWord, ";") = 0 Then newWord = newWord & ";"
            ElseIf word(w) = "BOM" Then
                newWord = "BOM"
                For s = 1 To 1
                    newWord = newWord & " " & word(w + s)
                Next
                If InStr(Len(newWord), newWord, ".") = Len(newWord) Then newWord = Left(newWord, Len(newWord) - 1)
                If InStr(Len(newWord), newWord, ";") = 0 Then newWord = newWord & ";"
            Else
                newWord = ""
            End If

            newContent = newContent & newWord

        Next

        newContent = Trim(newContent)
        If Len(newContent) > 0 Then
            If InStr(Len(newContent), newContent, ";") = Len(newContent) Then newContent = Left(newContent, Len(newContent) - 1)
        Else
        newContent = "<blank>"
        End If
        cell.value = newContent

    Next


End Sub

Sub SplitCommentRows()
Dim LR As Long, i As Long
Dim x As Variant, f%
Dim Delimiter$

Application.ScreenUpdating = False

Delimiter = ";"

LR = Range("B" & Rows.count).End(xlUp).row

'Delete Transmission Line rows
For f = LR To 1 Step -1
    If InStr(1, Cells(f, 1), "TRANSMISSION", vbBinaryCompare) <> 0 Then Rows(f).Delete
Next

LR = Range("A" & Rows.count).End(xlUp).row

Columns("B").Insert
For i = LR To 1 Step -1
    With Range("C" & i)
        If InStr(.value, Delimiter) = 0 Then
            .Offset(, -1).value = .value
        Else
            x = Split(.value, Delimiter)
            .Offset(1).Resize(UBound(x)).EntireRow.Insert
            .Offset(, -1).Resize(UBound(x) - LBound(x) + 1).value = Application.Transpose(x)
        End If
    End With
Next i
Columns("C").Delete

'Fill in blanks with station names in Col A left by transposing values from Col B
LR = Range("B" & Rows.count).End(xlUp).row
With Range("A1:B" & LR)
    On Error Resume Next
    .SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
    On Error GoTo 0
    .value = .value
End With
Application.ScreenUpdating = True
End Sub

Sub CleanCommentRows()
'With the comments broken up into a single column, performs the following clean up:
' - Removes whole row that doesn't contain SLD|BOM|GAD
' - Replaces document abbreviations with full names
' - Sorts and finds unique values

Dim LR As Long, i As Long
Dim result$, Flag%
Dim Keep, k, CommentArray As Variant
Dim rng As Range, r

LR = Range("A" & Rows.count).End(xlUp).row

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

'Add headers
Rows(1).Insert shift:=xlShiftDown
Range("A1").value = "Station"
Range("B1").value = "Document"

LR = Range("A" & Rows.count).End(xlUp).row

'Unique
With Worksheets("Source Docs Used")
    .Range("F:G").Delete
    .Range("A1:B" & LR).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=.Range("F:G"), Unique:=True
    .Range("A:E").Delete
    .Rows(1).Delete
    .Columns("A:B").AutoFit

'Remove trailing commas
Set rng = .Range("B1:B" & LR)
For Each r In rng
    r.value = Replace(r.value, ",", "")
Next

'Change BOM to "Bill of Materials, Dwg"
'       GAD to "General Arrangement Diagram, Dwg"
'       SLD to "Single Line Diagram, Dwg"
result = ReplaceText("BOM", "Bill of Materials~ Dwg")
result = ReplaceText("GAD", "General Arrangement Diagram~ Dwg")
result = ReplaceText("SLD", "Single Line Diagram~ Dwg")

'Text to Columns
LR = Range("A" & Rows.count).End(xlUp).row
    .Range("B1:B" & LR).TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Other:=True, OtherChar:="~", FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

'Trim
    Set rng = .Range("C1:C" & LR)
    rng.value = Application.Trim(rng)
End With

'Sort
ActiveSheet.Sort.SortFields.Clear
With ActiveSheet.Sort
    .SortFields.Add key:=Range("B1"), Order:=xlAscending
    .SortFields.Add key:=Range("A1"), Order:=xlDescending
    .SortFields.Add key:=Range("C1"), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    .SetRange ActiveSheet.Range("A1:C" & LR)
    .header = xlNo
    .Apply
ActiveSheet.Sort.SortFields.Clear

'Concatenate
For r = 1 To LR
    Cells(r, 3).value = Cells(r, 3).value & " - " & Cells(r, 1).value
Next
End With

End Sub

Function ReplaceText(str As String, repl As String) As String
'Replace text throughout one column
Dim UsedRangeCol$, ReplaceRangeCol$
Dim LastRow%, r%
UsedRangeCol = "A"
ReplaceRangeCol = "B"

LastRow = ActiveSheet.Range(UsedRangeCol & Rows.count).End(xlUp).row

For r = 1 To LastRow
    With Range(ReplaceRangeCol & r)
        .value = Replace(.value, str, repl, , , vbBinaryCompare)
    End With
Next

ReplaceText = "Done"

End Function

Sub MakeRatingsRequestedTable()
 Call MakeTable("Ratings Requested", "Verification", "A", Array("C", "F", "G", "H", "O"))
 With Sheets("Ratings Requested")
'   Format all rows
    With .Range("A:E")
        .Font.Bold = False
        .Font.Size = 9
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With
'   Format first row text
    With .Range("A1:E1")
        .Font.Bold = True
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
End With
End Sub

Sub MakeEquipAddedTable()
 Call MakeTable("Equipment Added", "Create", "A", Array("F", "G", "H"))
 With Sheets("Equipment Added")
'   Format all rows
    With .Range("A:E")
        .Font.Bold = False
        .Font.Size = 9
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With
'   Format first row text
    With .Range("A1:E1")
        .Font.Bold = True
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
End With
End Sub
Sub MakeEquipRetiredTable()
 Call MakeTable("Equipment Retired", "Retire", "A", Array("F", "G", "H"))
  With Sheets("Equipment Retired")
'   Format all rows
    With .Range("A:E")
        .Font.Bold = False
        .Font.Size = 9
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With
'   Format first row text
    With .Range("A1:E1")
        .Font.Bold = True
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
End With
End Sub

Sub MakeRelayRequestTable()
 Call MakeTable("TLS with Relay Request", "RLY", "H", Array("C", "F", "G", "H", "O", "Q", "R", "T", "U"))
End Sub

Sub MakeInfoRequestTable()
Dim rng As Range, cell
  With Worksheets("CAISO Update")
    .Range("XX1").value = "Line Name"
    .Range("XX2:XX" & .Range("F1", .Cells(Rows.count, 1).End(xlUp)).Rows.count).value = Split(.Parent.Name, "Rev")(0)
    .Range("XY1").value = "Field Verification Completion"

  Call MakeTable("Info Request", "Verification", "A", Array("BU", "BV", "XX", "F", "C", "G", "H", "O", "BT", "XY"))

    .Columns("XX:XY").Delete
  End With
  With Worksheets("Info Request")
    .Columns("B").HorizontalAlignment = xlCenter
    .Columns("E").HorizontalAlignment = xlCenter
    .Columns("F").HorizontalAlignment = xlLeft

    'Remove CT rows
    Set rng = Nothing
    For Each cell In .UsedRange.Columns("G").Offset(1, 0).Cells
        If Not IsError(cell.value) Then
            If cell.value Like "CT*" Then 'This will delete any row that has "CT" as Component Type
                If rng Is Nothing Then
                    Set rng = cell
                Else
                    Set rng = Application.Union(rng, cell)
                End If
            End If
        End If
    Next cell
    rng.Select
    If Not rng Is Nothing Then rng.EntireRow.Delete

        '   Format first row text
    With .Range("A1:P1")
        .Font.Bold = True
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
End With

End Sub
Sub MakeInfoRequestCTTable()
Dim Allrows%, col, cell, rng As Range
  With Worksheets("CAISO Update")
    Allrows = .Range("F1", .Cells(Rows.count, 1).End(xlUp)).Rows.count
    .Range("XX1").value = "Line Name"
    .Range("XX2:XX" & Allrows).value = Split(.Parent.Name, "Rev")(0)
    .Range("XY1").value = "Date Resolved"
    .Range("XZ1").value = "PG&E Verified Information"

  Call MakeTable("Info Request - CT", "Verification", "A", Array("BV", "XY", "C", "XX", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "XZ", "BU"))
  For Each col In Array("A", "C", "G", "H", "P")
    Worksheets("Info Request - CT").Range(col & ":" & col).HorizontalAlignment = xlCenter
  Next

    .Columns("XX:XZ").Delete
  End With

  With Worksheets("Info Request - CT")
    'Remove non-CT rows
    Set rng = Nothing
    For Each cell In .UsedRange.Columns("G").Offset(1, 0).Cells
        If Not IsError(cell.value) Then
            If Not cell.value Like "CT*" Then 'This will delete any row that doesn't have "CT" as Component Type
                If rng Is Nothing Then
                    Set rng = cell
                Else
                    Set rng = Application.Union(rng, cell)
                End If
            End If
        End If
    Next cell
    rng.Select
    If Not rng Is Nothing Then rng.EntireRow.Delete

    '   Format first row text
     With .Range("A1:P1")
         .Font.Bold = True
         .Font.Size = 10
         .VerticalAlignment = xlCenter
         .HorizontalAlignment = xlCenter
         .WrapText = True
     End With
End With

End Sub

Sub MakeTable(NewSheet As String, LookingFor As String, LookInCAISOCol As String, GetCAISOCols As Variant)
'Creates table suitable for pasting into Summary Report
'of all rows marked for verification where:
'NewSheet = name of the new tab storing the data created
'LookingFor = text to search for in each row
'LookInCAISOCol = column in the CAISO Update tab to search for the LookingFor text
'GetCAISOCols = list of columns to pull from CAISO Update to make the table

Dim c As Variant
Dim StartRow%, ChangeIDCol$, NewSheetCol, NewSheetRow, f As Integer
Dim CurrentVal$, red%
Dim wsTest As Worksheet
Dim col As Long, Dest As Range, ColumnNums() As Long
Dim DamnArray As Variant

StartRow = 1
ChangeIDCol = "XZ"

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

If MakeTopBottomArray("CAISO Update", ChangeIDCol, StartRow) Then


'    NewSheetCol = 1

    'Add headers
'    For Each c In GetCAISOCols
'        Worksheets(NewSheet).Cells(NewSheetRow, NewSheetCol).value = Worksheets("CAISO Update").Range(c & "1").value
'        NewSheetCol = NewSheetCol + 1
'    Next
'    NewSheetRow = NewSheetRow + 1


    ReDim ColumnNums(UBound(GetCAISOCols)) As Long
    For col = 0 To UBound(GetCAISOCols)
        ColumnNums(col) = Range(GetCAISOCols(col) & 1).Column
    Next
    With Worksheets(NewSheet)
        Set Dest = .Range("A1")
        'https://usefulgyaan.wordpress.com/2013/06/12/vba-trick-of-the-week-slicing-an-array-without-loop-application-index/
        '(go to last comment)
        'The first dimension is the number of rows and the second dimension is the number of columns
        'Usage: Destination Range, sized to 1 row by GetCAISOCols columns equals
        '    the whole CAISO Range, first row, using the ColumnNums specified in GetCAISOCols
        DamnArray = Application.Index(topArray, ColumnNums, 1)
        Dest.Resize(1, UBound(GetCAISOCols) + 1) = DamnArray
        .Activate
    End With

    NewSheetRow = 2

    If LookingFor = "Verification" Then
        For f = 1 To UBound(topArray, 2)
            red = Range(LookInCAISOCol & StartRow + f - 1).Font.Color ' collect rows that have read text, even if not "Verification"
            CurrentVal = topArray(1, f)
            If CurrentVal Like LookingFor Or red = 255 Then
                NewSheetCol = 1
                For Each c In GetCAISOCols
                    Worksheets(NewSheet).Cells(NewSheetRow, NewSheetCol).value = Worksheets("CAISO Update").Range(c & f + StartRow - 1).value
                    NewSheetCol = NewSheetCol + 1
                Next
                NewSheetRow = NewSheetRow + 1
            End If
        Next f
    Else
        For f = 1 To UBound(topArray, 2)
            CurrentVal = topArray(1, f)
            'Debug.Print CurrentVal & ": color " & Range(LookInCAISOCol & StartRow + f - 1).Font.Color
            If CurrentVal Like LookingFor Then
                NewSheetCol = 1
                For Each c In GetCAISOCols
                    Worksheets(NewSheet).Cells(NewSheetRow, NewSheetCol).value = Worksheets("CAISO Update").Range(c & f + StartRow - 1).value
                    NewSheetCol = NewSheetCol + 1
                Next
                NewSheetRow = NewSheetRow + 1
            End If
        Next f
    End If

    Worksheets(NewSheet).UsedRange.Columns.AutoFit
    Worksheets(NewSheet).Activate
End If 'MakeTopBotArray

End Sub
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Function ChangeIDsQC() As Boolean
 'Assure that there are two sections in the CurrentRegion split by a single mostly blank row,
 'with less rows in bottom than top, and all bottom section OIDs exist in the top
 Dim current_region As Range
 Dim all_array, connector, splitrow, k As Variant
 Dim ChangeIDVal$, ChangeIDCol$
 Dim EmptyCounter, FilledCount, row, col, StartRow, shortRowListText, s, b, botCount
 Dim shortRowList As Object
 Dim FoundCell As Range

 StartRow = "3"
 ChangeIDCol = "BY"

 On Error GoTo handler

 With Worksheets("CAISO Update")
     .Activate

     If MakeTopBottomArray("CAISO Update", ChangeIDCol, StartRow) Then
         If IsArray(botArray) Then botCount = UBound(botArray) Else botCount = 1
         For b = 1 To botCount
             If botCount = 1 Then ChangeIDVal = botArray Else ChangeIDVal = botArray(b, 1)
             Set FoundCell = BotRange.Cells(b, 1)
         'stop if a ChangeID is blank
             If ChangeIDVal = "" Then
                 FoundCell.Select
                 Err.Raise ERROR_BLANK_CHANGEID, "CheckChangeID", "Can't make SourceDocRef table." & vbCrLf & vbCrLf & _
                 "Found a blank cell in bottom range " & BotRange.Address & " at Row " & FoundCell.Address & vbCrLf & vbCrLf & _
                 "Fill this cell and rerun the macro."
             End If
    '     'stop if a bottom section value does not exist in top section
    '         If Not IsInArray(ChangeIDVal, topArray) Then
    '             FoundCell.Select
    '             Err.Raise ERROR_BLANK_CHANGEID, "CheckChangeID", "Can't make SourceDocRef table." & vbCrLf & vbCrLf & _
    '             "A ChangeID " & ChangeIDVal & " at Cell " & FoundCell.Address & " was not found in top section " & TopRange.Address & vbCrLf & vbCrLf & _
    '             "Edit this cell and rerun the macro."
    '         End If
         Next b
    End If

 End With

Done:
     ChangeIDsQC = True
     Exit Function

handler:
     MsgBox Err.Description
     ChangeIDsQC = False

 End Function

Function MakeTopBottomArray(wrksheet As String, LastCol As String, StartRow As Integer) As Boolean
 'Assure that there are two sections in the CurrentRegion split by a single mostly blank row,
 'with less rows in bottom than top, and all bottom section OIDs exist in the top

 Dim current_region As Range
 Dim all_array, connector, k As Variant
 Dim EmptyCounter, row, col, shortRowListText, s, b, botCount
 Dim shortRowList As Object
 Dim FilledCount, filledMax, totalCols

1
3
4 On Error GoTo handler
5
6 With Worksheets(wrksheet)
7     .Activate
8
9     Set current_region = Cells.CurrentRegion
10     Debug.Print "Current Region: " & current_region.Address
11     all_array = current_region
12
13     EmptyCounter = 0
14     Set shortRowList = CreateObject("Scripting.Dictionary")

     'Count filled rows to find least-filled
15     For row = 1 To UBound(all_array, 1)             'iterate through CurrentRegion rows
16         For col = 1 To UBound(all_array, 2)         'check each column
17             If IsError(all_array(row, col)) = False Then
                    If (all_array(row, col) = "") Then      'for blanks
18                      EmptyCounter = EmptyCounter + 1     'count them
                    End If
19             End If
20         Next

            'determine the maximum number of filled cells that represts a divider row,
            ' based on the total number of columns
            totalCols = Range(LastCol & 1).Column
            If totalCols < 7 Then
                filledMax = 2
            Else
                filledMax = 7
            End If
21
22         FilledCount = UBound(all_array, 2) - EmptyCounter
23         If FilledCount < filledMax Then   'look for any row that has less than the filled max specified.
                                             'The smallest number of filled cells
                                             ' in a typical CAISO export (TL row) has eight blanks
24             shortRowList.Add row, FilledCount
25         End If
26         EmptyCounter = 0
27     Next

     'stop if no row has less than seven filled cells
28     If shortRowList.count = 0 Then
29         current_region.Select
30         Err.Raise ERROR_TOP_BOTTOM_ARRAY, "CheckTopBottomArray", "Can't identify top and bottom sections in the highlighted range " & current_region.Address & "." & vbCrLf & vbCrLf & _
         "In that range, a mostly blank ('Required Updates') row that separates the top and bottom sections is not found." & vbCrLf & vbCrLf & _
         " - Does the highlighted area contain all your data?  If not, erase or fill in the row below the highlighted section." & vbCrLf & vbCrLf & _
         " - Is the row between top and bottom sections completely blank?  It needs at least one cell filled to allow comparison between top and bottom sections." & vbCrLf & vbCrLf & _
         " - Is there no mostly blank row between sections?  Add one, and put 'Required Updates', etc. into at least one cell." & vbCrLf & vbCrLf & _
         " - Does the separator row have seven or more filled cells?  Delete any unneeded data so that between one and seven cells have data."
36     End If

     'stop if more than one row has less than seven filled cells
37     If shortRowList.count > 1 Then
38         shortRowListText = ""
39         s = shortRowList.count
40         For Each k In shortRowList.keys
41             connector = Switch(s >= 3, ", ", s = 2, " and ", s = 1, ".")
42             shortRowListText = shortRowListText & k & connector
43             s = s - 1
44         Next
45
46         current_region.Select
         Err.Raise ERROR_TOP_BOTTOM_ARRAY, "CheckTopBottomArray", "More than one dividing row between top and bottom sections found in the highlighted range " & current_region.Address & "." & vbCrLf & vbCrLf & _
         "Found more than one possible dividing row between top and bottom sections, in rows " & shortRowListText & vbCrLf & vbCrLf & _
         "Complete filling in one of these rows so that there is one 'divider row' between top and bottom sections, i.e. contains one to seven filled cells."
50     End If

     'Find top and bottom sections, when shortRowList.Count = 1
51     splitrow = 0
52     For Each k In shortRowList.keys
53         splitrow = k
54     Next
55
56    Set TopRange = Range("A" & StartRow & ":" & LastCol & splitrow - 1)
57    topArray = Application.Transpose(TopRange.value)  'This converts 2d array into 1d
58    Set BotRange = Range("A" & splitrow + 1 & ":" & LastCol & UBound(all_array, 1))
59    botArray = Application.Transpose(BotRange.value)
    'botArray = BotRange
60        If IsArray(botArray) Then botCount = UBound(botArray) Else botCount = 1
61
62    Debug.Print "Top: " & TopRange.Address
63    Debug.Print "Bottom: " & BotRange.Address

    'stop if top section is smaller than bottom section
64     If UBound(topArray) < botCount Then
65         current_region.Select
         Err.Raise ERROR_TOP_BOTTOM_ARRAY, "CheckTopBottomArray", "Bottom section larger than Top section in highlighted range " & current_region.Address & "." & vbCrLf & vbCrLf & _
             "Top section in range " & TopRange.Address & " has fewer rows than bottom section in range " & BotRange.Address & "." & vbCrLf & vbCrLf & _
             "- Please make sure there is one mostly blank ('Required Updates') line with moe filled-in rows above than below"
69     End If
70
71
72 End With

Done:
     MakeTopBottomArray = True
     Exit Function

handler:
     MsgBox Err.Description & vbCrLf & "Line " & Err.Number
     MakeTopBottomArray = False
End Function

Sub DummyMacro()

MsgBox "Yes, this button works!"

End Sub
