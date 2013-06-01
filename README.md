# ADO

## Description
Класс для работы с объектами Connection и Recordset; выполнения SQL запросов к данным эксель, текстовым файлам, базам данных и т.п.

ADO (от англ. ActiveX Data Objects — «объекты данных ActiveX») — интерфейс программирования приложений для доступа к данным, разработанный компанией Microsoft (MS Access, MS SQL Server) и основанный на технологии компонентов ActiveX. ADO позволяет представлять данные из разнообразных источников (реляционных баз данных, текстовых файлов и т. д.) в объектно-ориентированном виде.&nbsp;&nbsp;&nbsp;[wiki](http://ru.wikipedia.org/wiki/ADO)

## API
### Methods:
- **Create** - создает объекты **_Connection_** и **_Recordset_**. Вызывается при создании экземпляра класса.
- **Connect** - открывает подключение к источнику данных. Вызывается при запросе с помощью метода **_Query_**.
- **Destroy** - уничтожает объекты **_Connection_** и **_Recordset_**. Срабатывает по событию `Class_Terminate()`.
- **Disconnect** - закрывает объект **_Recordset_** и соединение с источником данных. Срабатывает в по событию `Class_Terminate()`.
- **Query** - выполняет SQL запрос. Результат запроса помещается в объект **_Recordset_**. Возвращает время, в которое был выполнен запрос.
- **ToArray** - возвращает результат запроса в виде массива.

### Properties:
- **Connection** - объект соединения с источником данных.
- **Recordset** - результат выполнения запроса.
- **DataSoure** - источник данных. Полное имя книги эксель. По умолчанию текущая книга.
- **Header** - учитывать заголовки (да/нет). По умолчанию нет. В этом случае заголовки полей назначаются автоматически F1 ... Fn. Если да, первая строка диапазона считается заголовком поля.

В случае передачи параметра `ConnectionString` в метод **_Connect_**, значение свойств **_DataSoure_** и **_Header_** не учитываются, и формирование строки подключения ложиться полностью на плечи программиста.

## FAQ
**_1. Как начать работу?_**<br />
Для того, чтобы начать работу с объектом ADO, надо его создать:
```vbscript
Dim ADO As New ADO
```

**_2. Как сделать запрос к данным текущей книги?_**<br />
В данном запросе будут выбраны все данные из столбцов A:B с Листа1 текущей книги.<br />
При этом используются настройки по умолчанию: `Header = No`, `DataSource = ThisWorkbook.FullName`.
```vbscript
Sub Example()
    Dim ADO As New ADO
        
    ADO.Query ("SELECT * FROM [Лист1$A:B]")
End Sub
```

**_3. Как сделать запрос к данным текущей книги используя имена полей / заголовки столбцов?_**
```vbscript
Sub Example()
    Dim ADO As New ADO
        
    ADO.Header = True
    ADO.Query ("SELECT FieldName FROM [Лист1$A:B]")
End Sub
```

**_4. Как сделать запрос к данным другой книги?_**
```vbscript
Sub Example()
    Dim ADO As New ADO
        
    ADO.DataSource = Workbook.FullName   ' полный путь к книге
    ADO.Query ("SELECT * FROM [Лист1$A:B]")
End Sub
```

**_5. Как сделать запрос к другим источникам данных (базе данных, текстовым файлам и т.п.)?_**<br />
В данном случае формирование строки подключения ложится целиком на плечи программиста:
```vbscript
Sub Example()
    Dim ADO As New ADO
        
    ADO.Connect ("Your connection string")
    ADO.Query ("SELECT * FROM ...")
End Sub
```

**_6. Я сделал запрос. Где результат?_**<br />
Результат выполнения запроса хранится в объекте Recordset. Достучаться до него можно так:
```vbscript
ADO.Recordset
```

**_7. Как поместить результат выполнения запроса на лист?_**
```vbscript
Sub Example()
    Dim ADO As New ADO
        
    ADO.Query ("SELECT * FROM [Лист1$A:B]")
        
    Range("A1").CopyFromRecordset ADO.Recordset    ' поместить результат выполнения запроса на лист начиная с ячейки A1
End Sub
```

**_8. Как записать результат выполнения запроса в массив?_**<br />
Например, используя родной метод `getRows()` объекта `Recordset`:
```vbscript
Sub Example()
    Dim ADO As New ADO
    Dim Arr As Variant
        
    ADO.Query ("SELECT * FROM [Лист1$A:B]")
        
    Arr = ADO.Recordset.getRows()    ' записать результат выполнения запроса в массив
End Sub
```
Но в этом случае массив будет иметь нестандартный вид. Чтобы получить обычный двумерный массив, можно воспользоваться методом ```ToArray()```:
```vbscript
Sub Example()
    Dim ADO As New ADO
    Dim Arr As Variant
        
    ADO.Query ("SELECT * FROM [Лист1$A:B]")
        
    Arr = ADO.ToArray()
End Sub
```

**_Сахар_**<br />
Метод `Query` принимает `ParamArray()`, что позволяет писать запросы наглядно и достаточно лаконично (без лишней конкатенации строк):
```vbscript
Sub Example()
    Dim ADO As New ADO
        
    ADO.Query "SELECT F1", _
              "FROM [Sheet1$A:B]", _
              "WHERE F2 > 0"
End Sub
```

## Related links
[Использование ADO с данными Excel из Visual Basic или VBA](http://support.microsoft.com/kb/257819/ru)<br />
[Использование ADO.NET для извлечения и модификации записей книги Excel с помощью Visual Basic .NET](http://support.microsoft.com/kb/316934/ru)<br />
[Передача данных из набора записей ADO в Excel средствами автоматизации](http://support.microsoft.com/kb/246335/ru)<br />
[Использование библиотеки ADO (Microsoft ActiveX Data Object)](http://www.script-coding.com/ADO.html)<br />
<a href="http://msdn.microsoft.com/ru-ru/library/windows/desktop/ms678086(v=vs.85).aspx">ADO API Reference</a><br />
[ADO Tutorial](http://www.w3schools.com/ado/default.asp)<br />