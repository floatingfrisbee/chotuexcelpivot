VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Type filterInfo
    columnName As String
    values(5) As String
End Type

Private Type pivotTableInfo
    sheetName As String
    tableName As String
    dataField As String
    rowFields(5) As String
    columnFields(5) As String
    filters(5) As filterInfo
    
End Type

Private Type pivotInfo
    dataSheetName As String
    tables(100) As pivotTableInfo
End Type

Function IsInArray(arr() As String, stringToBeFound As String) As Boolean
  IsInArray = (UBound(filter(arr, stringToBeFound)) > -1)
End Function

Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
End Function

Function PivotTableExists(tableName As String, sheet As Worksheet) As Boolean
    Dim table As PivotTable
    On Error Resume Next
    Set table = sheet.PivotTables(tableName)
    On Error GoTo 0
    PivotTableExists = Not table Is Nothing
End Function

' This cleans all auto generated sheets. That is sheets that end with "AutoGen"
Public Sub CleanWorkbook()
    
    Application.DisplayAlerts = False
    
    For Each sheet In Sheets
        If sheet.Name Like "*AutoGen" Then
            sheet.Delete
        End If
    Next sheet
    
    Application.DisplayAlerts = True
    
End Sub

Public Sub Test()
    
    Dim filter As filterInfo
    filter.columnName = "Outcome_Outcome"
    filter.values(0) = "Family Decline"
    
    Dim ncreasonTable As pivotTableInfo
    ncreasonTable.sheetName = "FDs - Reason NC"
    ncreasonTable.tableName = "FDs - Reason NC Table"
    ncreasonTable.dataField = "Reason_NC"
    ncreasonTable.columnFields(0) = "year"
    ncreasonTable.filters(0) = filter
    
    Dim pi As pivotInfo
    pi.dataSheetName = "CLEAN_REFERRALS"
    pi.tables(0) = ncreasonTable
    
    On Error GoTo printError
    
    For i = 0 To UBound(pi.tables)
        If Not pi.tables(i).sheetName = "" Then
            CreatePivotTable pi.dataSheetName, pi.tables(i)
        Else
            Exit For
        End If
    Next i
    
printError:
    If Not Err.Description = "" Then
        MsgBox Err.Description
    End If
    
End Sub

' This is the function that should be run to generate
' the pivot tables.
Public Sub Main()
    
    Dim consentTable As pivotTableInfo
    consentTable.sheetName = "NYODN Consent Rates"
    consentTable.tableName = "NYODN Consent Table"
    consentTable.dataField = "nyodn_consent"
    consentTable.columnFields(0) = "year"
    
    Dim donationTable As pivotTableInfo
    donationTable.sheetName = "CMS Donation Rates"
    donationTable.tableName = "CMS Donation Table"
    donationTable.dataField = "cms_donation"
    donationTable.columnFields(0) = "year"
    
    Dim filter As filterInfo
    filter.columnName = "Outcome_Outcome"
    filter.values(0) = "Family Decline"
    filter.values(1) = "Donor"
    filter.values(2) = "CNR"
    
    Dim ncreasonTable As pivotTableInfo
    ncreasonTable.sheetName = "FDs - Reason NC"
    ncreasonTable.tableName = "FDs - Reason NC Table"
    ncreasonTable.dataField = "Reason_NC"
    ncreasonTable.columnFields(0) = "year"
    ncreasonTable.filters(0) = filter
    
    ' or I could say
    'ncreasonTable.filters(0).columnName = "outcome_outcome"
    'ncreasonTable.filters(0).value(0) = "Family Decline"
    
    Dim ncreasonByTeamTable As pivotTableInfo
    ncreasonByTeamTable.sheetName = "FDs - NC Reason by Team"
    ncreasonByTeamTable.tableName = "FDs - NC Reason by Team Table"
    ncreasonByTeamTable.dataField = "Reason_NC"
    ncreasonByTeamTable.rowFields(0) = "Team"
    ncreasonByTeamTable.columnFields(0) = "year"
    ncreasonByTeamTable.filters(0) = filter
        
    Dim outcomeTable As pivotTableInfo
    outcomeTable.sheetName = "Outcome"
    outcomeTable.tableName = "Outcome Table"
    outcomeTable.dataField = "Outcome_Outcome"
    outcomeTable.columnFields(0) = "year"
    
    Dim msrTable As pivotTableInfo
    msrTable.sheetName = "MSR"
    msrTable.tableName = "MSR Table"
    msrTable.dataField = "msr"
    msrTable.columnFields(0) = "year"
            
    Dim raceTable As pivotTableInfo
    raceTable.sheetName = "Race"
    raceTable.tableName = "Race Table"
    raceTable.dataField = "R"
    raceTable.columnFields(0) = "year"
    
    Dim codTable As pivotTableInfo
    codTable.sheetName = "COD"
    codTable.tableName = "COD Table"
    codTable.dataField = "CauseOfDeathType"
    codTable.columnFields(0) = "year"
    
    Dim ageTable As pivotTableInfo
    ageTable.sheetName = "Age"
    ageTable.tableName = "Age Table"
    ageTable.dataField = "age_cat5"
    ageTable.columnFields(0) = "year"
    
    Dim consentMSRTable As pivotTableInfo
    consentMSRTable.sheetName = "MSR - NYODN Consent Rates"
    consentMSRTable.tableName = "MSR - NYODN Consent Table"
    consentMSRTable.dataField = "nyodn_consent"
    consentMSRTable.columnFields(0) = "year"
    consentMSRTable.filters(0).columnName = "msr"
    consentMSRTable.filters(0).values(0) = "MSR Referral"
    
    Dim donationMSRTable As pivotTableInfo
    donationMSRTable.sheetName = "MSR - CMS Donation Rates"
    donationMSRTable.tableName = "MSR - CMS Donation Table"
    donationMSRTable.dataField = "cms_donation"
    donationMSRTable.columnFields(0) = "year"
    donationMSRTable.filters(0).columnName = "msr"
    donationMSRTable.filters(0).values(0) = "MSR Referral"
    
    Dim codMSRTable As pivotTableInfo
    codMSRTable.sheetName = "MSR - COD"
    codMSRTable.tableName = "MSR - COD Table"
    codMSRTable.dataField = "CauseOfDeathType"
    codMSRTable.columnFields(0) = "year"
    codMSRTable.filters(0).columnName = "msr"
    codMSRTable.filters(0).values(0) = "MSR Referral"
    
    Dim codMSRByTeamTable As pivotTableInfo
    codMSRByTeamTable.sheetName = "MSR - COD by Team"
    codMSRByTeamTable.tableName = "MSR - COD by Team Table"
    codMSRByTeamTable.dataField = "CauseOfDeathType"
    codMSRByTeamTable.rowFields(0) = "Team"
    codMSRByTeamTable.columnFields(0) = "year"
    codMSRByTeamTable.filters(0).columnName = "msr"
    codMSRByTeamTable.filters(0).values(0) = "MSR Referral"
    
    Dim codByTeamTable As pivotTableInfo
    codByTeamTable.sheetName = "COD by Team"
    codByTeamTable.tableName = "COD by Team Table"
    codByTeamTable.dataField = "CauseOfDeathType"
    codByTeamTable.rowFields(0) = "Team"
    codByTeamTable.columnFields(0) = "year"
    
    Dim consentAgeTable As pivotTableInfo
    consentAgeTable.sheetName = "NYODN Consent by Age"
    consentAgeTable.tableName = "NYODN Consent by Age Table"
    consentAgeTable.dataField = "nyodn_consent"
    consentAgeTable.rowFields(0) = "age_cat5"
    consentAgeTable.columnFields(0) = "year"
    
    Dim donationAgeTable As pivotTableInfo
    donationAgeTable.sheetName = "CMS Donation by Age"
    donationAgeTable.tableName = "CMS Donation by Age Table"
    donationAgeTable.dataField = "cms_donation"
    donationAgeTable.rowFields(0) = "age_cat5"
    donationAgeTable.columnFields(0) = "year"
    
    Dim consentRaceTable As pivotTableInfo
    consentRaceTable.sheetName = "NYODN Consent by Race"
    consentRaceTable.tableName = "NYODN Consent by Race Table"
    consentRaceTable.dataField = "nyodn_consent"
    consentRaceTable.rowFields(0) = "R"
    consentRaceTable.columnFields(0) = "year"
    
    Dim donationRaceTable As pivotTableInfo
    donationRaceTable.sheetName = "CMS Donation by Race"
    donationRaceTable.tableName = "CMS Donation by Race Table"
    donationRaceTable.dataField = "cms_donation"
    donationRaceTable.rowFields(0) = "R"
    donationRaceTable.columnFields(0) = "year"
    
    Dim consentAgeRaceTable As pivotTableInfo
    consentAgeRaceTable.sheetName = "NYODN Consent by Age and Race"
    consentAgeRaceTable.tableName = "NYODN Consent by Age and Race Table"
    consentAgeRaceTable.dataField = "nyodn_consent"
    consentAgeRaceTable.rowFields(0) = "age_cat5"
    consentAgeRaceTable.rowFields(1) = "R"
    consentAgeRaceTable.columnFields(0) = "year"
    
    Dim donationAgeRaceTable As pivotTableInfo
    donationAgeRaceTable.sheetName = "CMS Donation by Age and Race"
    donationAgeRaceTable.tableName = "CMS Donation by Age and Race Table"
    donationAgeRaceTable.dataField = "cms_donation"
    donationAgeRaceTable.rowFields(0) = "age_cat5"
    donationAgeRaceTable.rowFields(1) = "R"
    donationAgeRaceTable.columnFields(0) = "year"
    
    Dim consentTeamTable As pivotTableInfo
    consentTeamTable.sheetName = "NYODN Consent by Team"
    consentTeamTable.tableName = "NYODN Consent by Team Table"
    consentTeamTable.dataField = "nyodn_consent"
    consentTeamTable.rowFields(0) = "Team"
    consentTeamTable.columnFields(0) = "year"
    
    Dim consentTeamQtrTable As pivotTableInfo
    consentTeamQtrTable.sheetName = "NYODN Consent by Team Quarterly"
    consentTeamQtrTable.tableName = "NYODN Consent by Team Quarterly Table"
    consentTeamQtrTable.dataField = "nyodn_consent"
    consentTeamQtrTable.rowFields(0) = "Team"
    consentTeamQtrTable.columnFields(0) = "qtr"
    
    Dim consentRaceByTeamTable As pivotTableInfo
    consentRaceByTeamTable.sheetName = "NYODN Consent by Team and Race"
    consentRaceByTeamTable.tableName = "NYODN Consent by Team and Race Table"
    consentRaceByTeamTable.dataField = "nyodn_consent"
    consentRaceByTeamTable.rowFields(0) = "Team"
    consentRaceByTeamTable.rowFields(1) = "R"
    consentRaceByTeamTable.columnFields(0) = "year"
    
    Dim consentAgeByTeamTable As pivotTableInfo
    consentAgeByTeamTable.sheetName = "NYODN Consent by Team and Age"
    consentAgeByTeamTable.tableName = "NYODN Consent by Team and Age Table"
    consentAgeByTeamTable.dataField = "nyodn_consent"
    consentAgeByTeamTable.rowFields(0) = "Team"
    consentAgeByTeamTable.rowFields(1) = "age_cat5"
    consentAgeByTeamTable.columnFields(0) = "year"
    
    Dim donationTeamTable As pivotTableInfo
    donationTeamTable.sheetName = "CMS Donation by Team"
    donationTeamTable.tableName = "CMS Donation by Team Table"
    donationTeamTable.dataField = "cms_donation"
    donationTeamTable.rowFields(0) = "Team"
    donationTeamTable.columnFields(0) = "year"
    
    Dim donationTeamQtrTable As pivotTableInfo
    donationTeamQtrTable.sheetName = "CMS Donation by Team Quarterly"
    donationTeamQtrTable.tableName = "CMS Donation by Team Quarterly Table"
    donationTeamQtrTable.dataField = "cms_donation"
    donationTeamQtrTable.rowFields(0) = "Team"
    donationTeamQtrTable.columnFields(0) = "qtr"
    
    Dim donationRaceByTeamTable As pivotTableInfo
    donationRaceByTeamTable.sheetName = "CMS Donation by Team and Race"
    donationRaceByTeamTable.tableName = "CMS Donation by Team and Race Table"
    donationRaceByTeamTable.dataField = "cms_donation"
    donationRaceByTeamTable.rowFields(0) = "Team"
    donationRaceByTeamTable.rowFields(1) = "R"
    donationRaceByTeamTable.columnFields(0) = "year"
    
    Dim donationAgeByTeamTable As pivotTableInfo
    donationAgeByTeamTable.sheetName = "CMS Donation by Team and Age"
    donationAgeByTeamTable.tableName = "CMS Donation by Team and Age Table"
    donationAgeByTeamTable.dataField = "cms_donation"
    donationAgeByTeamTable.rowFields(0) = "Team"
    donationAgeByTeamTable.rowFields(1) = "age_cat5"
    donationAgeByTeamTable.columnFields(0) = "year"
    
    Dim consentCODByTeamTable As pivotTableInfo
    consentCODByTeamTable.sheetName = "NYODN Consent by Team and COD"
    consentCODByTeamTable.tableName = "NYODN Consent by Team and COD Table"
    consentCODByTeamTable.dataField = "nyodn_consent"
    consentCODByTeamTable.rowFields(0) = "Team"
    consentCODByTeamTable.rowFields(1) = "CauseOfDeathType"
    consentCODByTeamTable.columnFields(0) = "year"
    
    Dim donationCODByTeamTable As pivotTableInfo
    donationCODByTeamTable.sheetName = "CMS Donation by Team and COD"
    donationCODByTeamTable.tableName = "CMS Donation by Team and COD Table"
    donationCODByTeamTable.dataField = "cms_donation"
    donationCODByTeamTable.rowFields(0) = "Team"
    donationCODByTeamTable.rowFields(1) = "CauseOfDeathType"
    donationCODByTeamTable.columnFields(0) = "year"
    
    Dim pi As pivotInfo
    pi.dataSheetName = "CLEAN_REFERRALS"
    pi.tables(0) = consentTable
    pi.tables(1) = donationTable
    pi.tables(2) = ncreasonTable
    pi.tables(3) = ncreasonByTeamTable
    pi.tables(3) = outcomeTable
    pi.tables(4) = msrTable
    pi.tables(5) = raceTable
    pi.tables(6) = codTable
    pi.tables(7) = ageTable
    pi.tables(8) = consentAgeByTeamTable
    
    'Just a NOTE: Instead of creating tables above as we did, we could have also do all in one step as follows:
    ' ex. for consent table
    ' pivotInfo.tables(0).sheetName = "NYODN Consent Rates"
    ' pivotInfo.tables(0).tableName = "NYODN Consent Table"
    ' pivotInfo.tables(0).dataField = "nyodn_consent"
    ' pivotInfo.tables(0).columnFields(0) = "year"
    
    On Error GoTo printError
    
    CleanWorkbook
    
    For i = 0 To UBound(pi.tables)
        On Error GoTo failedPivot
        If Not pi.tables(i).sheetName = "" Then
            CreatePivotTable pi.dataSheetName, pi.tables(i)
        Else
            Exit For
        End If
failedPivot:
        If Not Err.Description = "" Then
            MsgBox pi.tables(i).sheetName & Err.Description
        End If
    Next i
    
printError:
    If Not Err.Description = "" Then
        MsgBox Err.Description
    End If
    
End Sub

Private Function ParseFilter(filterString As String) As filterInfo()
    
    Dim filters(5) As filterInfo
    
    Dim filterStrings() As String
    filterStrings = Split(filterString, ",")
    
    Dim filterCount
    filterCount = 0
    
    Dim fs As String
    
    For Each fs In filterStrings
    
        If filterCount = UBound(fs) Then
            Exit For
        End If
        
        Dim nameValue() As String
        nameValue = Split(Trim(fs), ":")
        
        filters(filterCount).columnName = Trim(nameValue(0))
        
        Dim values() As String
        values = Split(nameValue(1), "/")
        
        Dim valueCount
        valueCount = 0
        
        For Each filterValue In values
    
            If valueCount = UBound(filters(filterCount).values) Then
                Exit For
            End If
        
            filters(filterCount).values(valueCount) = Trim(filterValue)
            
            valueCount = valueCount + 1
            
        End
        
        filterCount = filterCount + 1
        
    Next filter
    
    ParseFilter = filters
    
End Function

Private Function ParseLine(line As String) As pivotTableInfo
    
    Dim elements() As String
    elements = Split(line, "|")

    Dim pt As pivotTableInfo
    
    pt.sheetName = Trim(elements(0))
    pt.tableName = Trim(elements(1))
    pt.dataField = Trim(elements(2))
    
    Dim columns() As String
    columns = Split(elements(3), ",")
    
    Dim count
    
    count = 0
    
    For Each col In columns
        If count > UBound(pt.columnFields) Then
            Exit For
        End If
        pt.columnFields(count) = Trim(col)
        count = count + 1
    Next col
    
    Dim rows() As String
    rows = Split(elements(4), ",")
    
    count = 0
    
    For Each rowField In rows
        If count > UBound(pt.rowFields) Then
            Exit For
        End If
        pt.rowFields(count) = Trim(rowField)
        count = count + 1
    Next rowField
    
    Dim j() As filterInfo
    j = ParseFilter(elements(5))
    
    ParseLine = pt
    
End Function

Private Function ReadInputFile(fullFilePath As String) As pivotTableInfo()
    
    Dim tables(100) As pivotTableInfo
    Dim iFileNum As Integer
    Dim line As String

    If Len(Dir$(fullFilePath)) = 0 Then
        MsgBox fullFilePath & " does not exist"
        Exit Function
    End If

    iFileNum = FreeFile()
    Open fullFilePath For Input As iFileNum

    Dim count
    
    count = 1
    
    Do While Not EOF(iFileNum)
        Line Input #iFileNum, line
        tables(count) = ParseLine(line)
        count = count + 1
    Loop

    ' close the file
    Close iFileNum

End Function

Private Sub TestFileRead()
    
    Dim filePath As String
    filePath = "C:\Users\me\jbakshi\Chotu\Misc\MakingPlotsInExcel\data.txt"
    
    Dim tables() As pivotTableInfo
    
    tables = ReadInputFile(filePath)
    
End Sub

Private Sub Serialize(pi As pivotInfo)
    Dim fullFilePath
    fullFilePath = "C:\Users\me\jbakshi\Chotu\Misc\MakingPlotsInExcel\inputdata.txt"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(fullFilePath)
    
    For i = 0 To UBound(pi.tables)
        If Not pi.tables(i).sheetName = "" Then
            Dim pt As pivotTableInfo
            
            Set pt = pi.tables(i)
            
            Dim first
            Dim number
            Dim columnString
            first = True
            
            For Each c In pt.columnFields
                If Not first Then
                    columnString = columnString & ", "
                Else
                    first = False
                End If
                
                columnString = columnString & c
            Next table
        
            Dim outputString
            
            outputString = pt.sheetName & " | " & _
                            pt.tableName & " | " & _
                            pt.dataField & " | "
                            
        Else
            Exit For
        End If
    Next i
    
    oFile.Close
    
    Set fso = Nothing
    Set oFile = Nothing

End Sub

Private Sub CreatePivotTable(dataSheetName As String, pt As pivotTableInfo)


    ' First create the sheet where the pivot table
    ' will live (if it doesn't already exists)
    
    Dim sName As String
    Dim sheetToUse As Worksheet
    
    sName = Strings.Left(pt.sheetName, 20) & " AutoGen"
    
    If Not SheetExists(sName) Then
        Worksheets.Add().Name = sName
    End If
    
    Set sheetToUse = ActiveWorkbook.Sheets(sName)
    
    sheetToUse.Select
    
    ' Delete the pivot table if it already exists
    If PivotTableExists(pt.tableName, sheetToUse) Then
        For Each table In sheetToUse.PivotTables
            table.PivotSelect "", xlDataAndLabel, True
            Selection.Delete
        Next table
    End If
    
    Dim pCache
    
    ' Add pivot cache
    Set pCache = ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, _
    SourceData:=Worksheets(dataSheetName).Range("A1").CurrentRegion)
    
    ' Add pivot table
    pCache.CreatePivotTable TableDestination:=sheetToUse.Range("A1"), _
    tableName:=pt.tableName, DefaultVersion:=xlPivotTableVersion10
    
    ' Add column fields
    With ActiveSheet.PivotTables(pt.tableName).PivotFields(pt.columnFields(0))
        .Orientation = xlColumnField
        .Position = 1
    End With
        
    ' Add data field for count
    ActiveSheet.PivotTables(pt.tableName).AddDataField ActiveSheet.PivotTables( _
        pt.tableName).PivotFields(pt.dataField), "Count", xlCount
    
    ' Add data field for percent
    ActiveSheet.PivotTables(pt.tableName).AddDataField ActiveSheet.PivotTables( _
        pt.tableName).PivotFields(pt.dataField), "Percent", xlCount
    
    ' Make the first data field count based
    With ActiveSheet.PivotTables(pt.tableName).DataPivotField
        .Orientation = xlColumnField
        .Position = 2
    End With
    
    ' Make the second data field percent based
    With ActiveSheet.PivotTables(pt.tableName).PivotFields("Percent")
        .Calculation = xlPercentOfColumn
        .NumberFormat = "0%"
    End With
    
    Dim posCounter
    
    posCounter = 1
    
    ' For each row field specified in the pivot table's
    ' definition, add a row pivot field
    For i = 0 To UBound(pt.rowFields)
        If Not pt.rowFields(i) = "" Then
            With ActiveSheet.PivotTables(pt.tableName).PivotFields(pt.rowFields(i))
                    .Orientation = xlRowField
                    .Position = posCounter
            End With
            posCounter = posCounter + 1
        Else
            Exit For
        End If
    Next i
        
    ' Add the data field as a row pivot field
    With ActiveSheet.PivotTables(pt.tableName).PivotFields(pt.dataField)
        .Orientation = xlRowField
        .Position = posCounter
    End With
    
    posCounter = 1
        
    'Add filter
    For i = 0 To UBound(pt.filters)
        If Not pt.filters(i).columnName = "" Then
            ' Add the column as a pivot field of type "xlPageField"
            With ActiveSheet.PivotTables(pt.tableName).PivotFields(pt.filters(i).columnName)
                .Orientation = xlPageField
                .Position = posCounter
                .EnableMultiplePageItems = True
            End With
            posCounter = posCounter + 1
            
            Dim filterColumn
            Set filterColumn = ActiveSheet.PivotTables(pt.tableName).PivotFields(pt.filters(i).columnName)
            
            ' Now turn on the values that are specified. Turn off
            ' other values
            For Each pItem In filterColumn.PivotItems
                If IsInArray(pt.filters(i).values, pItem.Name) Then
                    pItem.Visible = True
                Else
                    pItem.Visible = False
                End If
            Next pItem
        Else
            Exit For
        End If
    Next i

End Sub

