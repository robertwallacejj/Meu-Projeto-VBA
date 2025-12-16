VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendario 
   Caption         =   "Calendário"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5055
   OleObjectBlob   =   "frmCalendario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Dias As Collection
Public DataAtual As Date
Public TxtDestino As MSForms.TextBox

' ===============================
' FORM LOAD
' ===============================
Private Sub UserForm_Initialize()

    DataAtual = Date
    CriarSemana
    MontarCalendario

End Sub

' ===============================
' BOTÕES
' ===============================
Private Sub btnPrev_Click()
    DataAtual = DateAdd("m", -1, DataAtual)
    MontarCalendario
End Sub

Private Sub btnNext_Click()
    DataAtual = DateAdd("m", 1, DataAtual)
    MontarCalendario
End Sub

' ===============================
' SEMANAS
' ===============================
Private Sub CriarSemana()

    Dim diasSemana As Variant
    Dim i As Integer
    Dim lbl As MSForms.Label

    diasSemana = Array("Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb")

    For i = Me.Controls.Count - 1 To 0 Step -1
        If Left(Me.Controls(i).Name, 9) = "lblSemana" Then
            Me.Controls.Remove Me.Controls(i).Name
        End If
    Next i

    For i = 0 To 6
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblSemana" & i)
        With lbl
            .Caption = diasSemana(i)
            .Width = 25
            .Height = 18
            .Left = 10 + i * 30
            .Top = 40
            .TextAlign = fmTextAlignCenter
            .Font.Bold = True
        End With
    Next i

End Sub

' ===============================
' MONTA CALENDÁRIO
' ===============================
Private Sub MontarCalendario()

    Dim primeiroDia As Date
    Dim inicioSemana As Integer
    Dim ultimoDia As Integer

    lblMesAno.Caption = Format(DataAtual, "mmmm yyyy")

    primeiroDia = DateSerial(Year(DataAtual), Month(DataAtual), 1)
    inicioSemana = Weekday(primeiroDia, vbSunday)
    ultimoDia = Day(DateSerial(Year(DataAtual), Month(DataAtual) + 1, 0))

    CriarDias inicioSemana, ultimoDia

End Sub

' ===============================
' CRIA OS DIAS
' ===============================
Private Sub CriarDias(inicioSemana As Integer, ultimoDia As Integer)

    Dim i As Integer
    Dim diaNum As Integer
    Dim linha As Integer, coluna As Integer
    Dim btn As MSForms.CommandButton
    Dim evt As clsDia

    For i = Me.Controls.Count - 1 To 0 Step -1
        If Left(Me.Controls(i).Name, 6) = "diaBtn" Then
            Me.Controls.Remove Me.Controls(i).Name
        End If
    Next i

    Set Dias = New Collection
    diaNum = 1

    For i = inicioSemana To inicioSemana + ultimoDia - 1

        linha = Int((i - 1) / 7)
        coluna = (i - 1) Mod 7

        Set btn = Me.Controls.Add("Forms.CommandButton.1", "diaBtn" & diaNum)

        With btn
            .Caption = diaNum
            .Width = 25
            .Height = 22
            .Left = 10 + coluna * 30
            .Top = 65 + linha * 25
        End With

        Set evt = New clsDia
        Set evt.btn = btn
        evt.Dia = diaNum
        Set evt.frm = Me

        Dias.Add evt
        diaNum = diaNum + 1
        If diaNum > ultimoDia Then Exit For

    Next i

End Sub


