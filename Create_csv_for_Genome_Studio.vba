'**********************************************************************************
' Purpose:      Create csv file with patients positions for GenomeStudio
' Input:        Patient data and Illumina chip numbers in Excel worksheet
' Output:       A .csv file with patient data combined with correct plate positions
' Version:      v2 22.07.2025
' Changes:      Added 8-block validation: 8 samples must have the same plate serial
' Author:       Kadi Jairus
'               kadi.jairus@kliinikum.ee
'*********************************************************************************

Sub ExportCytoChipData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim firstRowIndex As Long
    Dim foundFirst As Boolean
    Dim header As String
    Dim outputFileName As String
    Dim filePath As String
    Dim dlgSaveFolder As FileDialog
    Dim rowData As String
    Dim sampleCounter As Integer
    Dim positionIndex As Integer
    Dim barcodeIndex As Integer
    Dim currentDate As String
    Dim positions As Variant
    Dim barcodes As Variant
    ' Variable to hold the validated barcode (plate serial) for the current block of 8 samples
    Dim currentBarcode As String

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
                Exit For
            End If
        End If
    Next rowIndex

    ' Exit if no valid row is found
    If firstRowIndex = 0 Then
        MsgBox "Ei õnnestunud järjekorranumbriga 1 algavat patsientide blokki.", vbExclamation
        Exit Sub
    End If
    
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
    
    
    
    MsgBox "Ekspordin " & lastRow - firstRowIndex + 1 & " patsienti:" & vbCrLf & _
        "Esimene patsient: " & ws.Cells(firstRowIndex, "C").Value & vbCrLf & _
           "Viimane patsient: " & ws.Cells(lastRow, "C").Value & vbCrLf & _
           "Järgmisena palun vali kaust, kuhu soovid csv-faili salvestada"

    ' Show folder picker dialog to select the save location
    Set dlgSaveFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With dlgSaveFolder
        .Title = "Vali kaust kuhu soovid csv-faili salvestada"
        .AllowMultiSelect = False
        If .Show = -1 Then
            filePath = .SelectedItems(1) & "\" & outputFileName
        Else
            MsgBox "Kausta valimine tühistati.", vbExclamation
            Exit Sub
        End If
    End With

    ' Open file for writing
    Open filePath For Output As #1
    Print #1, header

    ' Array to store SentrixPosition_A values cycling through (up to R08C01)
    positions = Array("R01C01", "R02C01", "R03C01", "R04C01", "R05C01", "R06C01", "R07C01", "R08C01")

    ' Initialize barcode variable
    currentBarcode = ""

    ' Loop upwards from the first found row
    For rowIndex = firstRowIndex To lastRow
        If ws.Cells(rowIndex, "L").Value <> "" Then
            
            ' Determine the position within the current block of 8 (0-7)
            positionIndex = (rowIndex - firstRowIndex) Mod 8
            
            ' If it's the first sample in a new block, store its barcode as the reference.
            If positionIndex = 0 Then
                currentBarcode = ws.Cells(rowIndex, "L").Value
            End If
            
            ' *** VALIDATION STEP ***
            ' Check if the barcode in the current row matches the reference barcode for this block.
            If ws.Cells(rowIndex, "L").Value <> currentBarcode Then
                MsgBox "VIGA ANDMETES: Plaadi numbri viga." & vbCrLf & vbCrLf & _
                       "Rida " & rowIndex & " patsient " & ws.Cells(rowIndex, "E").Value & " tulbas L on vale plaadi seerianumber." & vbCrLf & vbCrLf & _
                       "Peaks olema: " & currentBarcode & vbCrLf & _
                       "Leitud kood: " & ws.Cells(rowIndex, "L").Value & vbCrLf & vbCrLf & _
                       "CSV-faili tegemine on peatatud. Palun paranda andmed ja proovi uuesti!", vbCritical, "Viga andmetes"
                Close #1 ' Close the file to avoid a partial export
                Kill filePath ' Delete the incomplete file
                Exit Sub
            End If

            ' The sample counter is the position index + 1 (1-8)
            sampleCounter = positionIndex + 1

            ' Construct the row data using the validated barcode for the entire block
            rowData = ws.Cells(rowIndex, "E").Value & ",cyto,,cyto,," & _
                      "A" & Format(sampleCounter, "00") & "," & _
                      currentBarcode & "," & _
                      positions(positionIndex) & ",,,,,,,,,"

            Print #1, rowData
        End If
    Next rowIndex

    ' Close file
    Close #1

    MsgBox "Salvestasin faili siia: " & filePath, vbInformation
End Sub
