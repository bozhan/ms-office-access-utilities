VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim dbc As DbController
Dim dbs As DbServicer
Dim dbe As DbExporter
Dim dbp As DbPermissioner
Dim lgr As DbLogger

Private Sub Class_Initialize()
  Set dbc = New DbController
  Set dbs = New DbServicer
  Set dbe = New DbExporter
  Set dbp = New DbPermissioner
  Set lgr = New DbLogger
End Sub

Private Sub Class_Terminate()
  Set dbc = Nothing
  Set dbs = Nothing
  Set dbe = Nothing
  Set dbp = Nothing
  Set lgr = Nothing
End Sub

Public Property Get Logger() As DbLogger
  If lgr Is Nothing Then Set lgr = New DbLogger
  Set Logger = lgr
End Property

Public Property Get Controller() As DbController
  If dbc Is Nothing Then Set dbc = New DbController
  Set Controller = dbc
End Property

Property Get Servicer() As DbServicer
  If dbs Is Nothing Then Set dbs = New DbServicer
  Set Servicer = dbs
End Property

Property Get Exporter() As DbExporter
  If dbe Is Nothing Then Set dbe = New DbExporter
  Set Exporter = dbe
End Property

Property Get Permissioner() As DbPermissioner
  If dbp Is Nothing Then Set dbp = New DbPermissioner
  Set Permissioner = dbp
End Property

Public Sub Refresh()
  Call Class_Initialize
End Sub
