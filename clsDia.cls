VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public Dia As Integer
Public frm As Object

Private Sub btn_Click()

    If frm Is Nothing Then Exit Sub
    If frm.TxtDestino Is Nothing Then Exit Sub

    frm.TxtDestino.Value = Format( _
        DateSerial(Year(frm.DataAtual), Month(frm.DataAtual), Dia), _
        "dd/mm/yyyy")

    Unload frm

End Sub


