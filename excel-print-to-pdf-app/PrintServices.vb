Public Class PrintServices

    ' Static (Shared) Subroutine
    Shared Sub XLS_print_to_PDF()

        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim objStartFolder = fso.GetAbsolutePathName(".") + "'filesToPrint\"
        Dim objFolder = fso.GetFolder(objStartFolder)
        Dim colFiles = objFolder.Files

        For Each objFile In colFiles
            If LCase(fso.GetExtensionName(objFile)) = "xlsx" Or LCase(fso.GetExtensionName(objFile)) = "xls" Then

                ' Opens the Excel file
                Dim Excel = CreateObject("Excel.Application")
                Dim ExcelDoc = Excel.Workbooks.open(objStartFolder + objFile.Name)
                Dim pdfPath = (objStartFolder + Replace(objFile.Name, "." & fso.GetExtensionName(objFile.Path), "") + ".pdf")

                ' Print all the sheets to pdf
                ExcelDoc.PrintOut(Preview:=False, ActivePrinter:="Microsoft Print to PDF", PrintToFile:=True, Collate:=False, PrToFileName:=pdfPath, IgnorePrintAreas:=False)

                ' Close Excel Document and Application
                ExcelDoc.Close
                Excel.Application.Quit
            End If
        Next
    End Sub

    Shared Sub XLS_print_first_sheet_to_PDF()

        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim objStartFolder = fso.GetAbsolutePathName(".") + "'filesToPrint\"
        Dim objFolder = fso.GetFolder(objStartFolder)
        Dim colFiles = objFolder.Files

        For Each objFile In colFiles
            If LCase(fso.GetExtensionName(objFile)) = "xlsx" Or LCase(fso.GetExtensionName(objFile)) = "xls" Then

                ' Opens the Excel file
                Dim Excel = CreateObject("Excel.Application")
                Dim ExcelDoc = Excel.Workbooks.open(objStartFolder + objFile.Name)
                Dim objWorksheet = ExcelDoc.Sheets(1)
                Dim pdfPath = (objStartFolder + Replace(objFile.Name, "." & fso.GetExtensionName(objFile.Path), "") + ".pdf")

                ' Print only the first sheet to pdf
                objWorksheet.PrintOut(Preview:=False, ActivePrinter:="Microsoft Print to PDF", PrintToFile:=True, Collate:=False, PrToFileName:=pdfPath, IgnorePrintAreas:=False)

                ' Close Excel Document and Application
                ExcelDoc.Close
                Excel.Application.Quit
            End If
        Next
    End Sub

End Class
