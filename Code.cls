VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ShowResult_Click()
 '- Проверка на наличие старых результатов ------------------------------------
 For i = 1 To Sheets.Count
    If (Sheets(i).Name = "Результаты") Then
        MsgBox "Найден лист с предыдущем результатом, он будет удален.", , _
                    "Уведомление"
        Application.DisplayAlerts = False
        Sheets(i).Delete
        Application.DisplayAlerts = True
    End If
 Next i
 '- Отображение заголовка -----------------------------------------------------
 Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Результаты"
 Sheets("Результаты").Cells(1, 1).Value = "ФИО"
 Sheets("Результаты").Cells(1, 2).Value = "Возраст"
 Sheets("Результаты").Cells(1, 3).Value = "ИНН"
 Sheets("Результаты").Cells(1, 4).Value = "Организация"
 Sheets("Результаты").Cells(1, 5).Value = "КПП"
 Sheets("Результаты").Cells(1, 6).Value = "Численность организации"
 Sheets("Результаты").Cells(1, 7).Value = "Дата регистрации"
 Sheets("Результаты").Cells(1, 8).Value = "Кол-во регистраций"
 '- Количество заполненных строк в каждом листе -------------------------------
 Dim RowAtListA As Single
 Dim RowAtListB As Single
 Dim RowAtListV As Single
 Dim RowAtListG As Single
 Dim RowAtListD As Single
 RowAtListA = Sheets("Список А").UsedRange.Rows.Count
 RowAtListB = Sheets("Список Б").UsedRange.Rows.Count
 RowAtListV = Sheets("Список В").UsedRange.Rows.Count
 RowAtListG = Sheets("Список Г").UsedRange.Rows.Count
 RowAtListD = Sheets("Список Д").UsedRange.Rows.Count
 '- Заполнение массивов данными с листов --------------------------------------
 Dim TempListB() As String
 Call FillOutSheet(TempListB(), "Список Б", RowAtListB)
 
 Dim TempListV() As String
 Call FillOutSheet(TempListV(), "Список В", RowAtListV)
 
 Dim TempListG() As String
 Call FillOutSheet(TempListG(), "Список Г", RowAtListG)
 
 Dim TempListD() As String
 Call FillOutSheet(TempListD(), "Список Д", RowAtListD)
 '- Обработка данных с листа "Список А" ---------------------------------------
 Dim Result() As String
 ReDim Result(RowAtListA, 11)
 For i = 1 To RowAtListA - 1
    '- Запись Фамилии Имени Отчества в массив ---------------------------------
    Dim temp() As String
    temp = Split(Sheets("Список А").Cells(i + 1, 1).Value, " ", 3)
    For j = 1 To UBound(temp) - LBound(temp) + 1
        Result(i - 1, j - 1) = temp(j - 1)
    Next j
    Erase temp
    '- Запись ИНН и КПП в массив ----------------------------------------------
    temp = Split(Sheets("Список А").Cells(i + 1, 2).Value, "-", 2)
    For j = 1 To UBound(temp) - LBound(temp) + 1
        Result(i - 1, 2 + 2 * j) = temp(j - 1)
    Next j
    Erase temp
    '- Запись даты в массив ---------------------------------------------------
    temp = Split(Sheets("Список А").Cells(i + 1, 3).Value, ".", 3)
    For j = 1 To UBound(temp) - LBound(temp) + 1
        Result(i - 1, 7 + j) = temp(j - 1)
    Next j
    Erase temp
    '- Поиск соответствия по ФИО в формате ИОФ и запись возраста в массив -----
    Dim IOF As String
    IOF = Result(i - 1, 1) & " " & Result(i - 1, 2) & " " & Result(i - 1, 0)
    Call CompareAndPush(Result(), TempListB(), IOF, 3, i)
    '- Поиск соответствия по ИНН и запись организации в массив ----------------
    Call CompareAndPush(Result(), TempListV(), Result(i - 1, 4), 5, i)
    '- Поиск соответствия по КПП и запись численности организации в массив ----
    Call CompareAndPush(Result(), TempListG(), Result(i - 1, 6), 7, i)
    '- Поиск соответствия по дате в формате гггг-мм-дд и запись кол-ва --------
    '- регистраций в массив ---------------------------------------------------
    Dim ggggmmdd As String
    ggggmmdd = Result(i - 1, 10) & "-" & Result(i - 1, 9) & "-" & _
                    Result(i - 1, 8)
    Call CompareAndPush(Result(), TempListD(), ggggmmdd, 11, i)
    '- Отображение на листе информации из массива -----------------------------
    Sheets("Результаты").Cells(i + 1, 1).Value = Result(i - 1, 0) & " " & _
                    Result(i - 1, 1) & " " & Result(i - 1, 2)
    For j = 3 To 7
    Sheets("Результаты").Cells(i + 1, j - 1).Value = Result(i - 1, j)
    Next j
    Sheets("Результаты").Cells(i + 1, 7).Value = Result(i - 1, 8) & "." & _
            Result(i - 1, 9) & "." & Result(i - 1, 10)
    Sheets("Результаты").Cells(i + 1, 8).Value = Result(i - 1, 11)
 Next i
 Erase TempListB
 Erase TempListV
 Erase TempListG
 Erase TempListD
 Erase Result
 For i = 1 To 8
    Sheets("Результаты").Cells(i + 1, i).EntireColumn.AutoFit
 Next i
End Sub
'- Данная функция заполняет массив данными первых двух колонок с заданного ----
'- листа ----------------------------------------------------------------------
Sub FillOutSheet(TempSheet() As String, NameOfSheet As String, _
                    RowAtList As Single)
    ReDim TempSheet(RowAtList, 2)
    For i = 1 To RowAtList - 1
        TempSheet(i - 1, 0) = Sheets(NameOfSheet).Cells(i + 1, 1).Value
        TempSheet(i - 1, 1) = Sheets(NameOfSheet).Cells(i + 1, 2).Value
    Next i

End Sub
'- Данная функция сравнивает значения из одного массива с заданным условием, --
'- при совпадении записывает в другой массив в заданную позицию ---------------
Sub CompareAndPush(Result() As String, TempArr() As String, _
                    ByVal Condition As String, ByVal Position As Integer, _
                    ByVal i As Integer)
    For j = 1 To UBound(TempArr) - LBound(TempArr)
        If (TempArr(j - 1, 0) = Condition) Then
            Result(i - 1, Position) = TempArr(j - 1, 1)
            Exit For
        End If
    Next j
End Sub

