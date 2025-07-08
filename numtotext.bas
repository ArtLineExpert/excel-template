' Перевод числа в текст с правильными склонениями для рублей и копеек
Function NumToText(ByVal MyNumber As Double) As String
    Dim Units, Teens, Tens, Hundreds
    Units = Array("", "ОДИН", "ДВА", "ТРИ", "ЧЕТЫРЕ", "ПЯТЬ", "ШЕСТЬ", "СЕМЬ", "ВОСЕМЬ", "ДЕВЯТЬ")
    Teens = Array("ДЕСЯТЬ", "ОДИННАДЦАТЬ", "ДВЕНАДЦАТЬ", "ТРИНАДЦАТЬ", "ЧЕТЫРНАДЦАТЬ", "ПЯТНАДЦАТЬ", "ШЕСТНАДЦАТЬ", "СЕМНАДЦАТЬ", "ВОСЕМНАДЦАТЬ", "ДЕВЯТНАДЦАТЬ")
    Tens = Array("", "", "ДВАДЦАТЬ", "ТРИДЦАТЬ", "СОРОК", "ПЯТЬДЕСЯТ", "ШЕСТЬДЕСЯТ", "СЕМЬДЕСЯТ", "ВОСЕМЬДЕСЯТ", "ДЕВЯНОСТО")
    Hundreds = Array("", "СТО", "ДВЕСТИ", "ТРИСТА", "ЧЕТЫРЕСТА", "ПЯТЬСОТ", "ШЕСТЬСОТ", "СЕМЬСОТ", "ВОСЕМЬСОТ", "ДЕВЯТЬСОТ")

    Dim r As Long, k As Long
    r = Int(MyNumber)
    k = Round((MyNumber - r) * 100, 0)

    Dim result As String
    result = ""

    ' Сотни
    result = result & Hundreds(Int(r / 100)) & " "
    r = r Mod 100

    ' Десятки и единицы
    If r >= 10 And r <= 19 Then
        result = result & Teens(r - 10) & " "
    Else
        result = result & Tens(Int(r / 10)) & " "
        result = result & Units(r Mod 10) & " "
    End If

    result = Trim(result)

    ' Склонения РУБЛЕЙ
    Select Case Int(MyNumber) Mod 100
        Case 11 To 14
            result = result & " РУБЛЕЙ"
        Case Else
            Select Case Int(MyNumber) Mod 10
                Case 1: result = result & " РУБЛЬ"
                Case 2 To 4: result = result & " РУБЛЯ"
                Case Else: result = result & " РУБЛЕЙ"
            End Select
    End Select

    ' Копейки
    result = result & " "

    ' Склонения КОПЕЕК
    Select Case k
        Case 11 To 14
            result = result & Format(k, "00") & " КОПЕЕК"
        Case Else
            Select Case k Mod 10
                Case 1: result = result & Format(k, "00") & " КОПЕЙКА"
                Case 2 To 4: result = result & Format(k, "00") & " КОПЕЙКИ"
                Case Else: result = result & Format(k, "00") & " КОПЕЕК"
            End Select
    End Select

    NumToText = UCase(result)
End Function
