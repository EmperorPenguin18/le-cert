Attribute VB_Name = "Module1"

Sub GenerateCertificates()
    Dim FileName As String
    Dim PathName As String
    Dim ws As Worksheet
    Dim PersonName As String

    Set ws = ActiveWorkbook.Sheets("Sheet1")
    FileName = "namelist.csv"
    PathName = Application.ActiveWorkbook.Path
    ws.Copy
    ActiveWorkbook.SaveAs FileName:=PathName & "\" & FileName, FileFormat:=xlCSV, CreateBackup:=False

    Dim x As Variant
    x = Shell("lualatex " & PathName & "\certificate.tex", 1)

    Dim i as Integer
    i = 2
    Do While Cells(i, 8).Value <> ""
        PersonName = Cells(i, 1).Value & Cells(i, 2).Value
        x = Shell("pdftk " & PathName & "\certificate.pdf cat " & i & " output " & PersonName & ".pdf")
        SendEmail i PersonName
        i = i + 1
    Loop
End Sub

Sub SendEmail(row, PersonName)
    Dim EmailApp As Outlook.Application
    Set EmailApp = New Outlook.Application
    Dim EmailItem As Outlook.MailItem
    Set EmailItem = EmailApp.CreateItem(olMailItem)

    EmailItem.To = Cells(row, 8).Value
    EmailItem.Subject = Cells(2, 9).Value
    EmailItem.HTMLBody = Cells(2, 10).Value & vbNewLine & Cells(2, 11).Value

    EmailItem.Attachments.Add Application.ActiveWorkbook.Path & "\" & PersonName & ".pdf"

    EmailItem.Send
End Sub

Sub CreateShortcut()
    Application.OnKey "+^{A}", "GenerateCertificates"
End Sub
