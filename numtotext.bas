' Полный перевод чисел в рубли с копейками
Public Function NumToText(ByVal number As Double) As String
    Dim Units, Teens, Tens, Hundreds
    Units = Array("", "ОДИН", "ДВА", "ТРИ", "ЧЕТЫРЕ", "ПЯТЬ", "ШЕСТЬ", "СЕМЬ", "ВОСЕМЬ", "ДЕВЯТЬ")
    Teens = Array("ДЕСЯТЬ", "ОДИННАДЦАТЬ", "ДВЕНАДЦАТЬ", "ТРИНАДЦАТЬ", "ЧЕТЫРНАДЦАТЬ", "ПЯТНАДЦАТЬ", "ШЕСТНАДЦАТЬ", "СЕМНАДЦАТЬ", "ВОСЕМНАДЦАТЬ", "ДЕВЯТНАДЦАТЬ")
    Tens = Array("", "", "ДВАДЦАТЬ", "ТРИДЦАТЬ", "СОРОК", "ПЯТЬДЕСЯТ", "ШЕСТЬДЕСЯТ", "СЕМЬДЕСЯТ", "ВОСЕМЬДЕСЯТ", "ДЕВЯНОСТО")
    Hundreds = Array("", "СТО", "ДВЕСТИ", "ТРИСТА", "ЧЕТЫРЕСТА", "ПЯТЬСОТ", "ШЕСТЬСОТ", "СЕМЬСОТ", "ВОСЕМЬСОТ", "ДЕВЯТЬСОТ")

    Dim rubles As Long, kopecks As Long
    rubles = Int(number)
    kopecks = Round((number - rubles) * 100, 0)

    Dim parts As String
    parts = GetPart(rubles \ 1000000, "МИЛЛИОН", "МИЛЛИОНА", "МИЛЛИОНОВ", Units, Teens, Tens, Hundreds)
    parts = parts & " " & GetPart((rubles \ 1000) Mod 1000, "ТЫСЯЧА", "ТЫСЯЧИ", "ТЫСЯЧ", Units, Teens, Tens, Hundreds, True)
    parts = parts & " " & GetPart(rubles Mod 1000, "РУБЛЬ", "РУБЛЯ", "РУБЛЕЙ", Units, Teens, Tens, Hundreds)

    parts = Trim(parts)

    ' Добавляем копейки
    parts = parts & " " & Format(kopecks, "00") & " " & GetEnding(kopecks, "КОПЕЙКА", "КОПЕЙКИ", "КОПЕЕК")

    NumToText = UCase(Trim(parts))
End Function

Private Function GetPart(ByVal num As Long, ByVal form1 As String, ByVal form2 As String, ByVal form5 As String, _
                          Units, Teens, Tens, Hundreds, Optional isFemale As Boolean = False) As String
    If num = 0 Then Exit Function

    Dim res As String
    res = res & Hundreds(num \ 100) & " "

    Dim tensUnits As Long
    tensUnits = num Mod 100

    If tensUnits >= 10 And tensUnits <= 19 Then
        res = res & Teens(tensUnits - 10) & " "
    Else
        res = res & Tens(tensUnits \ 10) & " "
        Dim unitDigit As Long
        unitDigit = tensUnits Mod 10

        If isFemale Then
            If unitDigit = 1 Then
                res = res & "ОДНА "
            ElseIf unitDigit = 2 Then
                res = res & "ДВЕ "
            Else
                res = res & Units(unitDigit) & " "
            End If
        Else
            res = res & Units(unitDigit) & " "
        End If
    End If

    res = Trim(res) & " " & GetEnding(num, form1, form2, form5)
    GetPart = Trim(res)
End Function

Private Function GetEnding(ByVal number As Long, ByVal form1 As String, ByVal form2 As String, ByVal form5 As String) As String
    Dim n As Long
    n = number Mod 100

    If n >= 11 And n <= 19 Then
        GetEnding = form5
    Else
        Select Case n Mod 10
            Case 1: GetEnding = form1
            Case 2 To 4: GetEnding = form2
            Case Else: GetEnding = form5
        End Select
    End If
End Function
