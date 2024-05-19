VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const CONNECTION_STRING As String = "Provider=MSOLEDBSQL.1;Integrated Security=SSPI;Persist Security Info=False;User ID='';Initial Catalog=VB6;Data Source=DESKTOP-5O21EOC\SQLEXPRESS;"

Private conn As ADODB.Connection

Public Sub OpenConnection()
    Set conn = New ADODB.Connection
    conn.ConnectionString = CONNECTION_STRING
    conn.Open
End Sub

Public Sub CloseConnection()
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub
Public Function ExecuteSQL(sql As String) As Boolean
    On Error GoTo ErrorHandler
    conn.Execute sql
    ExecuteSQL = True
    Exit Function

ErrorHandler:
    ExecuteSQL = False
End Function

Public Function GetDriverForVehicle(vehno As Integer) As String
    Dim driverName As String
    Dim sql As String
    sql = "SELECT Driver.DriverFirstName, Driver.DriverLastName FROM Vehicle INNER JOIN Driver ON Vehicle.DriverNo = Driver.DriverNo WHERE Vehicle.Vehno = " & vehno
    Dim rs As ADODB.Recordset
    Set rs = conn.Execute(sql)
    
    If Not rs.EOF Then
        driverName = rs.Fields("DriverFirstName").Value & " " & rs.Fields("DriverLastName").Value
    End If
    
    rs.Close
    Set rs = Nothing
    
    GetDriverForVehicle = driverName
End Function


