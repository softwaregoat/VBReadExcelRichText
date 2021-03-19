Imports Spire.Xls

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim workbook As Workbook = New Workbook()
        workbook.LoadFromFile("D:\Project\VisualStudio\VB\VBReadExcelRichText\ReadExcelRichText\Book1.xlsx")
        Dim sheet As Worksheet = workbook.Worksheets(0)
        RichTextBox1.Rtf = sheet.Range("A1").RichText.RtfText
    End Sub
End Class
