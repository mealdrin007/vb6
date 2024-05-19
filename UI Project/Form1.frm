VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1335
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   3
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Vehicle No."
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim obj As New Project1.Class1
Dim flexGrid As MSFlexGrid

Private Sub Form_Load()
    Set flexGrid = MSFlexGrid1
    With flexGrid
        .Rows = 2
        .Cols = 2
        .TextMatrix(0, 0) = "Vehicle No."
        .TextMatrix(0, 1) = "Driver"
        .ColWidth(0) = 1000
        .ColWidth(1) = 2000
    End With
End Sub

Private Sub Command1_Click()
    obj.OpenConnection
    
    Dim sql As String
    sql = "SELECT * FROM Vehicle"
    
    Dim success As Boolean
    success = obj.ExecuteSQL(sql)
    
    Dim vehicleNumber As Integer
    vehicleNumber = CInt(Text1.Text)
    
    Dim driverName As String
    driverName = obj.GetDriverForVehicle(vehicleNumber)
    
    Dim rowIndex As Integer
    rowIndex = flexGrid.Rows
    flexGrid.Rows = rowIndex + 1
    flexGrid.TextMatrix(rowIndex, 0) = CStr(vehicleNumber)
    flexGrid.TextMatrix(rowIndex, 1) = driverName
    
    obj.CloseConnection
End Sub

