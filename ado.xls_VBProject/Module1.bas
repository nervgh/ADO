Attribute VB_Name = "Module1"
Sub Example()
    Dim ADO As New ADO
    
    ADO.Query ("SELECT * FROM [Лист1$A1:A9] UNION ALL SELECT * FROM [Лист1$B1:B5];")
    
    Range("E1").CopyFromRecordset ADO.Recordset
End Sub

