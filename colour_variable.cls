VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colour_variable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Public variable_name As String
Public variable_description As String
Public variable_colour_win As Long
Public variable_colour_html As String

'SAVES THE COLVAR MEMBERS TO THE SYSTEM REGISTRY
Public Sub save_vars()
    SaveSetting App.ProductName, variable_name, "variable_colour_win", variable_colour_win
    SaveSetting App.ProductName, variable_name, "variable_colour_htlm", variable_colour_html
End Sub

'LOADS THE COLVAR MEMBERS FROM THE SYSTEM REGISTRY
Public Sub load_vars()
    If GetSetting(App.ProductName, variable_name, "variable_colour") <> "" Then
        variable_colour_win = Val(GetSetting(App.ProductName, variable_name, "variable_colour_win"))
        variable_colour_html = GetSetting(App.ProductName, variable_name, "variable_colour_html")
    End If
End Sub

