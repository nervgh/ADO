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
' ADO (�� ����. ActiveX Data Objects � �������� ������ ActiveX�) � ��������� ���������������� ���������� ��� ������� � ������,
' ������������� ��������� Microsoft (MS Access, MS SQL Server) � ���������� �� ���������� ����������� ActiveX.
' ADO ��������� ������������ ������ �� ������������� ���������� (����������� ��� ������, ��������� ������ � �. �.)
' � ��������-��������������� ����.
'
' ����� ��� �������� ������� ����������� (Connection) � ����������� SQL �������� � ������.
'
' reference - http://msdn.microsoft.com/ru-ru/library/windows/desktop/ms678086(v=vs.85).aspx
' wiki - http://ru.wikipedia.org/wiki/ADO
'
' @author nerv
' @version 06/04/2013, 0.1
'--------------------------------

' ������:
'   Connect    - ������� ����� �����������. ������������� ��������� ��� �������.
'   Disconnect - ��������� �������� ������ � �����������, ���������� ��. ���������� �������������.
'   Query      - ��������� SQL ������. ��������� ������� �������� � ������ Recordset. ���������� �����, ����� ��� ������� ������.

' ��������:
'   Connection - ������ ����������
'   Recordset  - ��������� ��������� �������


Option Explicit


Public Connection As Object
Public Recordset As Object


Public Sub Connect()
    If Me.Connection Is Nothing Then
        Set Me.Connection = CreateObject("ADODB.Connection")
    End If
End Sub


Public Sub Disconnect()
    If Not Me.Recordset Is Nothing Then
        Me.Recordset.Close
    End If
    
    If Not Me.Connection Is Nothing Then
        Me.Connection.Close
    End If
    
    Set Me.Connection = Nothing
    Set Me.Recordset = Nothing
End Sub


Public Function Query(ByVal QueryString As String) As Date
    Call Me.Connect
    
    If Me.Connection.ConnectionString = "" Then
        Me.Connection.Open GetSource()
    End If
    
    Set Me.Recordset = Me.Connection.Execute(QueryString)
    Query = Now
End Function


Private Sub Class_Initialize()
    Call Me.Connect
End Sub


Private Sub Class_Terminate()
    Call Me.Disconnect
End Sub


Private Function GetSource() As String
    Dim Source As String
    Source = IIf(Val(Application.Version) < 12, _
        "Provider='Microsoft.Jet.OLEDB.4.0';Data Source=':1';Extended Properties='Excel 4.0;HDR=No;IMEX=1';", _
        "Provider='Microsoft.ACE.OLEDB.12.0';Data Source=':1';Extended Properties='Excel 12.0;HDR=No;IMEX=1';")
    Source = Replace(Source, ":1", ThisWorkbook.FullName, , 1)
    GetSource = Source
End Function