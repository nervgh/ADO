Attribute VB_Name = "Module1"
Sub Example()
    Dim ADO As New ADO
    
    ADO.Query ("SELECT F1 FROM [Лист1$];")
    Range("E1").CopyFromRecordset ADO.Recordset
    
    ADO.Query ("SELECT F2 FROM [Лист1$];")
    Range("F1").CopyFromRecordset ADO.Recordset
    
    ' Закрываем соединение, чтобы не висело : )
    ADO.Disconnect
    
    ADO.Query ("SELECT F1 FROM [Лист1$] UNION SELECT F2 FROM [Лист1$];")
    Range("G1").CopyFromRecordset ADO.Recordset
    
    ' Тут автоматически закроется соединение
    ' и уничтожиться объекты Recordset и Connection
End Sub

