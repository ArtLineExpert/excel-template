' === Обновление структуры листов из CSV ===
Public Sub UpdateStructure()
    Dim url As String
    url = "https://raw.githubusercontent.com/ArtLineExpert/excel-template/main/structure.csv"

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send

    If http.Status <> 200 Then
        MsgBox "Ошибка загрузки структуры: " & http.Status
        Exit Sub
    End If

    Dim csvContent As String
    csvContent = http.responseText

    Dim lines() As String
    lines = Split(csvContent, vbLf)

    Dim i As Long
    For i = 0 To UBound(lines)
        If Trim(lines(i)) <> "" Then
            Dim parts() As String
            parts = Split(lines(i), ",")

            If UBound(parts) >= 2 Then
                Dim sheetName As String: sheetName = parts(0)
                Dim cellAddress As String: cellAddress = parts(1)
                Dim value As String: value = parts(2)

                On Error Resume Next
                Dim ws As Worksheet
                Set ws = ThisWorkbook.Sheets(sheetName)

                If ws Is Nothing Then
                    ' Если листа нет, создаём
                    Set ws = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
                    ws.Name = sheetName
                End If

                On Error GoTo 0

                ' Если начинается с "=", вставляем как формулу
                If Left(value, 1) = "=" Then
                    ws.Range(cellAddress).Formula = value
                Else
                    ws.Range(cellAddress).Value = value
                End If
            End If
        End If
    Next i

    MsgBox "Структура успешно обновлена"
End Sub

