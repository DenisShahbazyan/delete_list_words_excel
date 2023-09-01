Sub DelLineFromList1CellA()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim rng1 As Range, rng2 As Range
    Dim cell As Range, checkWord As Range
    Dim deleteRow As Boolean
    Dim LastRow1 As Long, LastRow2 As Long
	Dim res As Boolean

    ' Указываем названия листов
    Set ws1 = ThisWorkbook.Sheets("Лист1")
    Set ws2 = ThisWorkbook.Sheets("Лист2")
	
    ' Удаляем пустые строки на листе 2
    res = DeleteEmpty(ws2)
    
    ' Задаем диапазоны данных на листах
    LastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    Set rng1 = ws1.Range("A1:A" & LastRow1)
	LastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    Set rng2 = ws2.Range("A1:A" & LastRow2)
	
    ' Отключаем автофильтр, если уже включен
    ws1.AutoFilterMode = False
    
    ' Проходим по каждой ячейке на первом листе (с конца)
    For i = LastRow1 To 1 Step -1
        deleteRow = False
        Set cell = rng1.Cells(i, 1) ' используем ячейку на текущей итерации
        
        ' Проходим по каждому слову из второго листа
        For Each checkWord In rng2
	    ' Проверяем, является ли слово из второго листа цельным словом, а не частью другого слова
	    If IsWordPartOfAnotherWord(checkWord.Value, cell.Value) Or InStr(1, cell.Value, checkWord.Value, vbTextCompare) > 0 Then
		deleteRow = True
		Exit For ' Если нашли хотя бы одно совпадение, то можно выйти из цикла
	    End If
        Next checkWord
        
        
        ' Удаляем строку, если было найдено совпадение и слово не ¤вл¤етс¤ частью другого слова
        If deleteRow Then
            cell.EntireRow.Delete
        End If
    Next i
    
    ' Отключаем автофильтр
    ws1.AutoFilterMode = False
End Sub

'удаляем пустые строки
Function DeleteEmpty(ws2 As Worksheet) As Boolean
    Dim r As Long, lastRow As Long
    Dim rng As Range

    ' Находим последнюю строку с данными в столбце A
    lastRow = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row

    For r = 1 To lastRow
        If Application.CountA(ws2.Rows(r)) = 0 Then
            If rng Is Nothing Then
                Set rng = ws2.Rows(r)
            Else
                Set rng = Union(rng, ws2.Rows(r))
            End If
        End If
    Next r

    If Not rng Is Nothing Then
        rng.Delete
        DeleteEmpty = True ' Return True if rows were deleted
    Else
        DeleteEmpty = False ' Return False if no rows were deleted
    End If
End Function

Function IsWordPartOfAnotherWord(word As String, fullWord As String) As Boolean
    Dim wordArray() As String
    Dim i As Long
    
    ' Разбиваем полное слово на отдельные слова
    wordArray = Split(fullWord, " ")
    
    ' Проверяем каждое слово из разбитого массива
    For i = LBound(wordArray) To UBound(wordArray)
        ' Сравниваем слово с каждым элементом массива
        If InStr(1, LCase(wordArray(i)), LCase(word), vbTextCompare) > 0 Then
            IsWordPartOfAnotherWord = True
            Exit Function
        End If
    Next i
    
    ' Если ни одно слово не совпало, то возвращаем False
    IsWordPartOfAnotherWord = False
End Function
