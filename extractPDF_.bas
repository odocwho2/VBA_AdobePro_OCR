Attribute VB_Name = "extractPDF_"
Sub extractPdf()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

ocrPath = ThisWorkbook.Sheets("Front sheet").Cells(4, 2).Value & "\" ' read dynamic path
fileName = Dir(ocrPath) 'read the files in the directory

k = 2 ' initialize variable K

While fileName <> ""

If Left(fileName, 3) = "OCR" Then 'make sure file contains OCR prefix
    DeliveryNumber = "" ' intialize empty variable
    SplitName = Split(fileName, ".")
    AfterSplit = SplitName(0)
    ThisWorkbook.Sheets("Data").Cells(k, 2).Value = Right(AfterSplit, Len(AfterSplit) - 6) 'read the document number and write it in excel
    
    Set objApp = CreateObject("AcroExch.App") 'set de reference for PDF App
    Set objPDDoc = CreateObject("AcroExch.PDDoc") 'set the reference for PDF Dcc
    
    If objPDDoc.Open(ocrPath & fileName) Then
    Set objjso = objPDDoc.GetJSObject
        LastPageNumber = objPDDoc.GetNumPages - 1
        Page = 0
        For CurentePage = LastPageNumber To Page Step -1 ' go through pages in reverse order
        wordsCount = objjso.GetPageNumWords(CurentePage) ' count all the words in the currente page
            For i = 0 To wordsCount ' read each word
                If objjso.getPageNthWord(CurentePage, i) = "IMP" Then ' check if current word matches our repetitive fix word
                    DeliveryNumber = objjso.getPageNthWord(CurentePage, i - 1) ' if word found then variable DeliberyNumber = current word possition -1 is the word we need to write down in Excel
                    ThisWorkbook.Sheets("Data").Cells(k, 3).Value = DeliveryNumber ' write the word in Excel
                    k = k + 1 ' increment row number
                    GoTo End_: ' go to label to skip all the words left in the document
                End If
                
            
            Next i
        Next
    Else
    ThisWorkbook.Sheets("Data").Cells(k, 4).Value = "Delivery number not found in the document !" 'write the error message if no match found after full document read.
    
    End If

     k = k + 1 ' increment row number
End_:
    objPDDoc.Close
    Set objjso = Nothing

    End If
    fileName = Dir ' load the next fileName
Wend

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

