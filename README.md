# Automation-of-PDF-generation-using-template-
Through this project, I enhanced my skills in: 📌 Excel VBA Automation 📌 PDF Generation 📌 Dynamic Template Handling 📌 Process Automation
# Prompt - Automate certificate work



I have two files: A Word document containing a certificate template with three textbox: <<CANDIDATE NAME>>, <<DATE>>, and <<CERTIFICATE ID>>.



Second File : An Excel sheet with candidate details, structured as follows:



Column A: Candidate ID

Column B: Candidate Name

Column C: Date

Column D: Certificate ID



Write an Excel VBA code for me, Where once I run the VBA code

1) First, show a pop-up to select the certificate design Word file.

2) Next, show a pop-up to select the destination folder for saving PDFs.

3) The macro should replace the textbox inside the Word document using the corresponding data from the sheet it should be dynamically add all the <<CANDIDATE NAME>>, <<DATE>>, and <<CERTIFICATE ID>> for each candidate and generate separated pdf for every candidate.

4) Generate a PDF certificate for each candidate and save it in the selected folder.

5) Each PDF file should be named using the Candidate ID (e.g., 1001\_Certificate.pdf).


Code : VBA Code to be paste in excel  

Option Explicit

Sub Generate_Certificate_PDFs()

    Dim wdApp As Object
    Dim wdDoc As Object
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Dim templatePath As String
    Dim saveFolder As String
    
    Dim candidateID As String
    Dim candidateName As String
    Dim certDate As String
    Dim certificateID As String
    
    Dim pdfFileName As String

    '========================================
    ' SELECT WORD TEMPLATE
    '========================================
    
    With Application.FileDialog(msoFileDialogFilePicker)
    
        .Title = "Select Certificate Word Template"
        .Filters.Clear
        .Filters.Add "Word Files", "*.docx"
        
        If .Show <> -1 Then Exit Sub
        
        templatePath = .SelectedItems(1)
        
    End With

    '========================================
    ' SELECT OUTPUT FOLDER
    '========================================
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .Title = "Select Folder to Save PDFs"
        
        If .Show <> -1 Then Exit Sub
        
        saveFolder = .SelectedItems(1)
        
    End With

    '========================================
    ' EXCEL SHEET
    '========================================
    
    Set ws = ThisWorkbook.Sheets(1)
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    '========================================
    ' OPEN WORD
    '========================================
    
    On Error Resume Next
    
    Set wdApp = GetObject(, "Word.Application")
    
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    
    On Error GoTo 0

    wdApp.Visible = False

    '========================================
    ' LOOP THROUGH EXCEL ROWS
    '========================================
    
    For i = 2 To lastRow
    
        candidateID = ws.Cells(i, 1).Value
        candidateName = ws.Cells(i, 2).Value
        certDate = ws.Cells(i, 3).Text
        certificateID = ws.Cells(i, 4).Value

        ' OPEN TEMPLATE
        
        Set wdDoc = wdApp.Documents.Open(templatePath)

        '====================================
        ' REPLACE PLACEHOLDERS
        '====================================
        
        ReplaceEverywhere wdDoc, "<<CANDIDATE NAME>>", candidateName
        ReplaceEverywhere wdDoc, "<<DATE>>", certDate
        ReplaceEverywhere wdDoc, "<<CERTIFICATE ID>>", certificateID

        '====================================
        ' SAVE PDF
        '====================================
        
        pdfFileName = saveFolder & "\" & candidateID & "_Certificate.pdf"

        wdDoc.ExportAsFixedFormat _
            OutputFileName:=pdfFileName, _
            ExportFormat:=17

        wdDoc.Close False

    Next i

    '========================================
    ' CLOSE WORD
    '========================================
    
    wdApp.Quit

    Set wdDoc = Nothing
    Set wdApp = Nothing

    MsgBox "All certificates generated successfully!", vbInformation

End Sub

'========================================================
' REPLACE TEXT EVERYWHERE
'========================================================

Sub ReplaceEverywhere(wdDoc As Object, findText As String, replaceText As String)

    Dim shp As Object
    Dim sec As Object
    Dim hdr As Object
    Dim rng As Object

    '----------------------------------------
    ' MAIN DOCUMENT
    '----------------------------------------
    
    With wdDoc.Content.Find
    
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Text = findText
        .Replacement.Text = replaceText
        
        .Wrap = 1
        
        .Execute Replace:=2
        
    End With

    '----------------------------------------
    ' TEXTBOXES / SHAPES
    '----------------------------------------
    
    For Each shp In wdDoc.Shapes
    
        If shp.TextFrame.HasText Then
        
            With shp.TextFrame.TextRange.Find
            
                .ClearFormatting
                .Replacement.ClearFormatting
                
                .Text = findText
                .Replacement.Text = replaceText
                
                .Wrap = 1
                
                .Execute Replace:=2
                
            End With
            
        End If
        
    Next shp

    '----------------------------------------
    ' HEADERS / FOOTERS
    '----------------------------------------
    
    For Each sec In wdDoc.Sections
    
        For Each hdr In sec.Headers
        
            Set rng = hdr.Range
            
            With rng.Find
            
                .ClearFormatting
                .Replacement.ClearFormatting
                
                .Text = findText
                .Replacement.Text = replaceText
                
                .Wrap = 1
                
                .Execute Replace:=2
                
            End With
            
        Next hdr
        
    Next sec

End Sub

