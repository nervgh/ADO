VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------
' ADO (от англ. ActiveX Data Objects — «объекты данных ActiveX») — интерфейс программирования приложений для доступа к данным,
' разработанный компанией Microsoft (MS Access, MS SQL Server) и основанный на технологии компонентов ActiveX.
' ADO позволяет представлять данные из разнообразных источников (реляционных баз данных, текстовых файлов и т. д.)
' в объектно-ориентированном виде.
'
' Класс для создания объекта подключения (Connection) и выполенения SQL запросов в эксель.
'
' reference - http://msdn.microsoft.com/ru-ru/library/windows/desktop/ms678086(v=vs.85).aspx
' wiki - http://ru.wikipedia.org/wiki/ADO
'
' http://support.microsoft.com/kb/257819/ru
' http://support.microsoft.com/kb/316934/ru
' http://support.microsoft.com/kb/246335/ru
'
' http://www.w3schools.com/ado/default.asp
'
' @author nerv
' @version 08/04/2013, 0.3
'--------------------------------

' Методы:
'   Create     - создает объет подключения. Автоматически вызывается при инициализации.
'   Connect    - открывает соединение. Автоматически вызывается при запросе.
'   Destroy    - уничтожает объект подключения и объект записей. Вызывается автоматически при выходе из программы.
'   Disconnect - закрывает открытые записи и подключения. Вызывается автоматически при выходе из программы.
'   Query      - выполняет SQL запрос. Результат запроса помещает в объект Recordset. Возвращает время, когда был выполен запрос.

' Свойства:
'   Connection - объект соединения.
'   Recordset  - результат выполения запроса.
'   DataSoure  - источник данных. Полное имя книги эксель. По умолчанию текущая книга.
'   Header     - учитывать заголовки (да/нет). По умолчанию нет. В этом случае имана полей назначаются автоматически F1 ... Fn.
'                Если да, первая строка диапазона считается заголовком поля.

Option Explicit


Public Connection As Object
Public Recordset As Object
Public DataSource As String
Public Header As Boolean


Public Sub Create()
    Set Me.Connection = CreateObject("ADODB.Connection")
End Sub


Public Sub Destroy()
    Set Me.Recordset = Nothing
    Set Me.Connection = Nothing
End Sub


Public Sub Connect()
    Me.Connection.Open GetConnectionString()
End Sub


Public Sub Disconnect()
    If Not Me.Recordset Is Nothing Then
        If Me.Recordset.State = 1 Then
            Me.Recordset.Close
            Me.Connection.Close
        End If
    End If
End Sub


Public Function Query(ByVal QueryString As String) As Date
    If Me.Connection Is Nothing Then
        Call Me.Create
    End If
    
    If Me.Connection.State = 0 Then
        Call Me.Connect
    End If

    Set Me.Recordset = Me.Connection.Execute(QueryString)
    Query = Now
End Function


Private Sub Class_Initialize()
    Call Me.Create
End Sub


Private Sub Class_Terminate()
    Call Me.Disconnect
    Call Me.Destroy
End Sub


Private Function GetConnectionString() As String
    Dim Params As String
    Params = IIf(Val(Application.Version) < 12, _
        "Provider='Microsoft.Jet.OLEDB.4.0';Data Source=':1';Extended Properties='Excel 4.0;HDR=:2;IMEX=1';", _
        "Provider='Microsoft.ACE.OLEDB.12.0';Data Source=':1';Extended Properties='Excel 12.0;HDR=:2;IMEX=1';")
    Params = Replace(Params, ":1", IIf(Me.DataSource = "", ThisWorkbook.FullName, Me.DataSource), , 1)
    Params = Replace(Params, ":2", IIf(Me.Header, "Yes", "No"), , 1)
    GetConnectionString = Params
End Function
