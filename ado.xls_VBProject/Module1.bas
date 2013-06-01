Sub Example()
    Dim ADO As New ADO
    
    ADO.Query ("SELECT F1, F2 FROM [Лист1$];")
    Range("E1").CopyFromRecordset ADO.Recordset
    
    ADO.Query ("SELECT F2 FROM [Лист1$];")
    Range("F1").CopyFromRecordset ADO.Recordset
    
    ' Уничтожаем объекты Connection и Recordset
    ' (соединение закроется автоматически)
    ADO.Destroy
    
    ' Здесь снова создадутся объекты Connection и Recordset,
    ' установится соединение с источником данных,
    ' после чего выполнится запрос
    ADO.Query ("SELECT F1 FROM [Лист1$] UNION SELECT F2 FROM [Лист1$];")
    Range("G1").CopyFromRecordset ADO.Recordset
    
    ' Тут автоматически закроется соединение
    ' и уничтожатся объекты Recordset и Connection
End Sub
