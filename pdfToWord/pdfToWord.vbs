Option Explicit
Option Private Module
'Author: Pratima Kshetry
Sub SavePDFAsWord(PDFPath As String, FileExtension As String)
      
    'In order to use the macro you must enable the Acrobat library from VBA editor:
      
    Dim objAcroApp      As Acrobat.AcroApp
    Dim objAcroAVDoc    As Acrobat.AcroAVDoc
    Dim objAcroPDDoc    As Acrobat.AcroPDDoc
    Dim objJSO          As Object
    Dim boResult        As Boolean
    Dim ExportFormat    As String
    Dim NewFilePath     As String
   
    'Check if the file exists.
    If Dir(PDFPath) = "" Then
        MsgBox "Cannot find the PDF file!" 
        Exit Sub
    End If
   
    'Check if the input file is a PDF file.
    If LCase(Right(PDFPath, 3)) <> "pdf" Then
        MsgBox "The input file is not  PDF !"
        Exit Sub
    End If
   
    'Initialize Acrobat by creating App object.
    Set objAcroApp = CreateObject("AcroExch.App")
   
    'Set AVDoc object.
    Set objAcroAVDoc = CreateObject("AcroExch.AVDoc")
   
    'Open the PDF file.
    boResult = objAcroAVDoc.Open(PDFPath, "")
       
    'Set the PDDoc object.
    Set objAcroPDDoc = objAcroAVDoc.GetPDDoc
   
    'Set the JS Object - Java Script Object.
    Set objJSO = objAcroPDDoc.GetJSObject
   
    'Check the type of conversion.
    Select Case LCase(FileExtension)
        Case "docx": ExportFormat = "com.adobe.acrobat.docx"
        Case "doc": ExportFormat = "com.adobe.acrobat.doc"
        Case Else: ExportFormat = "Wrong Input"
    End Select
    
    'Check if the format is correct and no errors.
    If ExportFormat <> "Wrong Input" And Err.Number = 0 Then
        
        'Format is correct and no errors.
        
        'Set the path of the new file. Note that Adobe instead of xls uses xml files.
        'That's why here the xls extension changes to xml.
        If LCase(FileExtension) <> "xls" Then
            NewFilePath = WorksheetFunction.Substitute(PDFPath, ".pdf", "." & LCase(FileExtension))
        Else
            NewFilePath = WorksheetFunction.Substitute(PDFPath, ".pdf", ".xml")
        End If
        
        'Save PDF file to the new format.
        boResult = objJSO.SaveAs(NewFilePath, ExportFormat)
        
        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
        
        'Close the Acrobat application.
        boResult = objAcroApp.Exit
        
       MsgBox "The PDf file:" & vbNewLine & PDFPath & vbNewLine & vbNewLine & _
        "Was saved as: " & vbNewLine & NewFilePath, vbInformation, "Conversion successfull"
         
    Else
       
        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
       
        'Close the Acrobat application.
        boResult = objAcroApp.Exit
       
        End If
       
    'Release the objects.
    Set objAcroPDDoc = Nothing
    Set objAcroAVDoc = Nothing
    Set objAcroApp = Nothing
       
End Sub