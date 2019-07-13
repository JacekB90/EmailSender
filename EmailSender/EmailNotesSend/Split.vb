Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office

Public Class Split

    Public Sub SplitExcel(path As String, save As String)

        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBooks As Excel.Workbooks = Nothing
        Dim xlWorkbook As Excel.Workbook = Nothing
        Dim xlWorkSheet As Excel.Worksheet = Nothing
        Dim xlWorkSheets As Excel.Sheets = Nothing

        Dim xlAppNew As Excel.Application = Nothing
        Dim xlWorkBooksNew As Excel.Workbooks = Nothing
        Dim xlWorkbookNew As Excel.Workbook = Nothing
        Dim xlWorkSheetNew As Excel.Worksheet = Nothing
        Dim xlWorkSheetsNew As Excel.Sheets = Nothing

        Dim countRows, scopeCopy, counter As Integer

        If Not System.IO.File.Exists(path) Then

            MsgBox(path & " not exists.")
            Exit Sub

        ElseIf (System.IO.Path.GetExtension(path) = ".xlsx" Or System.IO.Path.GetExtension(path) = ".xls") Then

            xlApp = New Excel.Application
            xlApp.DisplayAlerts = False
            xlWorkBooks = xlApp.Workbooks
            xlWorkbook = xlWorkBooks.Open(path)

        Else

            MsgBox(path & " wrong type of file.")
            Exit Sub

        End If

        Try
            xlWorkSheet = xlWorkbook.Worksheets(1)

            counter = 0
            countRows = xlWorkSheet.UsedRange.Rows.Count

            Do While String.IsNullOrEmpty(xlWorkSheet.Range("A2").Value) = False

                scopeCopy = 2

                For i = 1 To countRows
                    If xlWorkSheet.Range("A2").Value = xlWorkSheet.Range("A" & i + 2).Value Then

                        scopeCopy += 1

                    End If
                Next i

                xlWorkSheet.Rows("1" & ":" & scopeCopy).Copy()

                xlAppNew = New Excel.Application
                xlAppNew.DisplayAlerts = False
                xlWorkBooksNew = xlAppNew.Workbooks
                xlWorkbookNew = xlWorkBooksNew.Add()

                xlAppNew.EnableEvents = False
                xlWorkbookNew.Activate()
                xlWorkSheetNew = xlWorkbookNew.Worksheets(1)

                xlWorkSheetNew.Range("A1").PasteSpecial(Excel.XlPasteType.xlPasteValues)

                With xlWorkbookNew
                    .SaveAs(save & "\" & xlWorkSheetNew.Range("A2").Value & ".xlsx")
                    .Close()
                End With

                xlWorkSheet.Rows("2" & ":" & scopeCopy).Delete()

                counter += 1
            Loop

            xlWorkbook.Close()

            If counter = 1 Then
                MsgBox(counter & " file has been created.")
            Else
                MsgBox(counter & " files have been created.")
            End If

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try

    End Sub

End Class
