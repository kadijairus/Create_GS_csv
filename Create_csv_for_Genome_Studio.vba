Sub ExportCytoChipData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim firstRowIndex As Long
    Dim foundFirst As Boolean
    Dim header As String
    Dim outputFileName As String
    Dim filePath As String
    Dim rowData As String
    Dim sampleCounter As Integer
    Dim positionIndex As Integer
    Dim barcodeIndex As Integer
    Dim currentDate As String
    Dim positions As Variant
    Dim barcodes As Variant

    ' Set worksheet to the currently active sheet
    Set ws = ActiveSheet

    ' Get the current date in dd.mm.yyyy format
    currentDate = Format(Date, "dd.mm.yyyy")

    ' Find the last row with data in column L
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row

    ' Initialize tracking variables
    foundFirst = False
    firstRowIndex = 0

    ' Find the first row from bottom where column B = 1 and L is not empty
    For rowIndex = lastRow To 2 Step -1
        If ws.Cells(rowIndex, "L").Value <> "" Then
            If ws.Cells(rowIndex, "B").Value = 1 Then
                firstRowIndex = rowIndex
                foundFirst = True
                Exit For ' Stop once the first row with B = 1 is found
            End If
        End If
    Next rowIndex

    ' Exit if no valid row is found
    If firstRowIndex = 0 Then
        MsgBox "No matching row found where column B = 1 and column L is not empty.", vbExclamation
        Exit Sub
    End If

    ' Display message box with values from column C for first and last rows
    MsgBox "Esimene patsient: " & ws.Cells(firstRowIndex, "C").Value & vbCrLf & _
           "Viimane patsient: " & ws.Cells(lastRow, "C").Value, vbInformation

    ' Prepare header for the CSV file with dynamic date
    header = "[Header],,,,,,,,,,,,,,,," & vbCrLf & _
             "Investigator Name,,,,,,,,,,,,,,,," & vbCrLf & _
             "Project Name,cyto,,,,,,,,,,,,,,," & vbCrLf & _
             "Experiment Name,,,,,,,,,,,,,,,," & vbCrLf & _
             "Date," & currentDate & ",,,,,,,,,,,,," & vbCrLf & _
             "[Manifests],,,,,,,,,,,,,,,," & vbCrLf & _
             "A,GDA-8v1-0_D2,,,,,,,,,,,,,,," & vbCrLf & _
             "[Data],,,,,,,,,,,,,,,," & vbCrLf & _
             "Sample_ID,Sample_Plate,Sample_Name,Project,AMP_Plate,Sample_Well,SentrixBarcode_A,SentrixPosition_A,Scanner,Date_Scan,Replicate,Parent1,Parent2,Gender,Replicate,Parent1,Parent2"

    ' Define output file name with dynamic date
    outputFileName = "CytoChip_" & Replace(currentDate, ".", "_") & ".csv"
    filePath = ThisWorkbook.Path & "\" & outputFileName

    ' Open file for writing
    Open filePath For Output As #1
    Print #1, header

    ' Array to store SentrixPosition_A values cycling through (up to R08C01)
    positions = Array("R01C01", "R02C01", "R03C01", "R04C01", "R05C01", "R06C01", "R07C01", "R08C01")

    ' Array to simulate barcode values
    barcodes = Array("208698540015", "208698540016", "208698540075", "208698540087")

    ' Reset counters
    sampleCounter = 1
    positionIndex = 0
    barcodeIndex = 0

    ' Loop upwards from first found row with B = 1
    For rowIndex = firstRowIndex To lastRow
        If ws.Cells(rowIndex, "L").Value <> "" Then
            rowData = ws.Cells(rowIndex, "E").Value & ",cyto,,cyto,," & _
                      "A" & Format(sampleCounter, "00") & "," & _
                      barcodes(barcodeIndex) & "," & _
                      positions(positionIndex) & ",,,,,,,,,"

            Print #1, rowData

            ' Update counters for sample and position logic
            sampleCounter = sampleCounter + 1
            positionIndex = positionIndex + 1

            ' If we reach the 8th position, reset and move to next barcode
            If positionIndex > UBound(positions) Then
                positionIndex = 0
                barcodeIndex = barcodeIndex + 1
                If barcodeIndex > UBound(barcodes) Then barcodeIndex = 0
            End If

            ' Reset sampleCounter every 8 samples
            If sampleCounter > 8 Then sampleCounter = 1
        End If
    Next rowIndex

    ' Close file
    Close #1

    MsgBox "CSV file created successfully: " & filePath, vbInformation
End Sub



