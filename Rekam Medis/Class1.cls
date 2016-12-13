VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub msubSetDeleteKeyComma(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 44 Then KeyAscii = 0
End Sub

Public Sub msubPesanError()
    MsgBox "Ada kesalahan dalam loading data" & vbNewLine _
        & "Hubungi administrator dan laporkan pesan berikut" & vbNewLine _
        & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub

Public Function mfuncRecFO(recordset As ADODB.recordset, Query As String) As Boolean
    On Error GoTo errLoad
    mfuncRecFO = False
    Set recordset = New ADODB.recordset
    recordset.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    mfuncRecFO = True
    Exit Function
errLoad:
    MsgBox "Ada kesalahan dalam loading data" & vbNewLine _
        & "Hubungi administrator dan laporkan pesan kesalahan berikut" & vbNewLine _
        & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Function


