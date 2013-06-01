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

## Related links
[Использование ADO с данными Excel из Visual Basic или VBA](http://support.microsoft.com/kb/257819/ru)<br />
[Использование ADO.NET для извлечения и модификации записей книги Excel с помощью Visual Basic .NET](http://support.microsoft.com/kb/316934/ru)<br />
[Передача данных из набора записей ADO в Excel средствами автоматизации](http://support.microsoft.com/kb/246335/ru)<br />
[Использование библиотеки ADO (Microsoft ActiveX Data Object)](http://www.script-coding.com/ADO.html)<br />
<a href="http://msdn.microsoft.com/ru-ru/library/windows/desktop/ms678086(v=vs.85).aspx">ADO API Reference</a><br />
[ADO Tutorial](http://www.w3schools.com/ado/default.asp)<br />