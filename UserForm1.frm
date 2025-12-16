VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8850.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6195
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================
' BOTÃO DATA DE INDUÇÃO (ABRE O CALENDÁRIO)
' =====================================================
Private Sub btnDataInducao_Click()
    ' Abre o form calendário para seleção de data
    Dim frm As New frmCalendario
    Set frm.TxtDestino = Me.txtDataInducao
    frm.Show
End Sub

' =====================================================
' BOTÃO LIMPAR
' =====================================================
Private Sub btnLimpar_Click()
    LimparCampos
End Sub

' =====================================================
' BOTÃO SALVAR
' =====================================================
Private Sub btnSalvar_Click()
    On Error GoTo ErroTratado

    ' Variáveis de planilhas
    Dim wsAut As Worksheet, wsBase As Worksheet, wsRec As Worksheet
    Dim existe As Variant, autorizado As Variant
    Dim linhaOriginal As Long, ultimaLinha As Long
    Dim novoID As Long, obs As String, idOriginal As Variant
    Dim camposObrigatorios As Variant, nomeCampo As Variant
    Dim ctrl As Control

    ' Variáveis de data/hora
    Dim dataInducao As Date, dataHoraInicio As Date, dataHoraFim As Date

    ' =====================================================
    ' 1. CAMPOS OBRIGATÓRIOS
    ' =====================================================
    camposObrigatorios = Array("txtAWB", "cmbAGENTEDECARGAS", _
                               "txtDataInducao", "txtInicioInducao", "txtFimInducao", _
                               "txtSelecao", "txtLiberados", "txtDevolucao", _
                               "txtManifestado", "txtAPAC", "txtFiscalizacao")
    
    For Each nomeCampo In camposObrigatorios
        Set ctrl = Me.Controls(nomeCampo)
        If Trim(ctrl.Value) = "" Then
            MsgBox "Preencha o campo: " & Replace(nomeCampo, "txt", ""), vbExclamation
            ctrl.SetFocus
            Exit Sub
        End If
    Next nomeCampo

    ' =====================================================
    ' 2. VALIDAR DATA E HORA
    ' =====================================================
    If Not IsDate(txtDataInducao.Value) Then
        MsgBox "Selecione a data de indução.", vbExclamation
        txtDataInducao.SetFocus
        Exit Sub
    End If

    If Not IsDate(txtInicioInducao.Value) Or Not IsDate(txtFimInducao.Value) Then
        MsgBox "Horário de indução inválido.", vbExclamation
        Exit Sub
    End If

    dataInducao = CDate(txtDataInducao.Value)
    dataHoraInicio = dataInducao + TimeValue(txtInicioInducao.Value)
    dataHoraFim = dataInducao + TimeValue(txtFimInducao.Value)

    ' =====================================================
    ' 3. VALIDAR AGENTE (AUTORIZADOS)
    ' =====================================================
    Set wsAut = ThisWorkbook.Sheets("AUTORIZADOS")
    autorizado = Application.Match(cmbAGENTEDECARGAS.Value, wsAut.Columns("A"), 0)
    If IsError(autorizado) Then
        MsgBox "Agente de cargas NÃO autorizado.", vbCritical
        cmbAGENTEDECARGAS.SetFocus
        Exit Sub
    End If

    ' =====================================================
    ' 4. VALIDAR AWB NA FUP_ADUANEIRO
    ' =====================================================
    Set wsBase = ThisWorkbook.Sheets("fup_aduaneiro")
    existe = Application.Match(txtAWB.Value, wsBase.Columns("C"), 0)
    If IsError(existe) Then
        MsgBox "AWB não encontrada na FUP ADUANEIRO.", vbCritical
        txtAWB.SetFocus
        Exit Sub
    End If

    ' =====================================================
    ' 5. REGRA DO ID E OBSERVAÇÃO (FUP_RECEBIMENTO)
    ' =====================================================
    Set wsRec = ThisWorkbook.Sheets("fup_recebimento")
    existe = Application.Match(txtAWB.Value, wsRec.Columns("A"), 0)

    If IsError(existe) Then
        ' Primeiro registro -> ID = 1
        ultimaLinha = wsRec.Cells(wsRec.Rows.Count, "A").End(xlUp).Row + 1
        novoID = 1

        ' Observação opcional
        If MsgBox("Deseja adicionar uma observação?", vbYesNo + vbQuestion) = vbYes Then
            obs = InputBox("Digite a observação:")
            wsRec.Cells(ultimaLinha, "O").Value = obs
        End If
    Else
        linhaOriginal = existe
        idOriginal = wsRec.Cells(linhaOriginal, "P").Value

        If Trim(idOriginal) = "" Then
            ' Primeiro ID ainda
            ultimaLinha = linhaOriginal
            novoID = 1

            If MsgBox("Deseja adicionar uma observação?", vbYesNo + vbQuestion) = vbYes Then
                obs = InputBox("Digite a observação:")
                wsRec.Cells(ultimaLinha, "O").Value = obs
            End If
        Else
            ' Segundo registro ou mais -> observação obrigatória
            Do
                obs = InputBox("AWB já possui registro." & vbCrLf & "Informe a observação (obrigatória):")
                If Trim(obs) = "" Then MsgBox "Observação obrigatória!", vbExclamation
            Loop Until Trim(obs) <> ""

            ultimaLinha = wsRec.Cells(wsRec.Rows.Count, "A").End(xlUp).Row + 1
            novoID = wsRec.Cells(wsRec.Rows.Count, "P").End(xlUp).Value + 1
            wsRec.Cells(ultimaLinha, "O").Value = obs
        End If
    End If

    ' =====================================================
    ' 6. GRAVAR DADOS NO FUP_RECEBIMENTO
    ' =====================================================
    With wsRec
        .Cells(ultimaLinha, "A").Value = txtAWB.Value
        .Cells(ultimaLinha, "B").Value = dataHoraInicio
        .Cells(ultimaLinha, "E").Value = dataHoraFim
        .Cells(ultimaLinha, "B").NumberFormat = "dd/mm/yyyy hh:mm"
        .Cells(ultimaLinha, "E").NumberFormat = "dd/mm/yyyy hh:mm"

        .Cells(ultimaLinha, "H").Value = txtSelecao.Value
        .Cells(ultimaLinha, "I").Value = txtLiberados.Value
        .Cells(ultimaLinha, "J").Value = txtDevolucao.Value
        .Cells(ultimaLinha, "K").Value = txtManifestado.Value
        .Cells(ultimaLinha, "L").Value = txtAPAC.Value
        .Cells(ultimaLinha, "M").Value = txtFiscalizacao.Value
        .Cells(ultimaLinha, "N").Value = Now
        .Cells(ultimaLinha, "P").Value = novoID
        .Cells(ultimaLinha, "Q").Value = cmbAGENTEDECARGAS.Value  ' registra agente que lançou
    End With

    MsgBox "Registro salvo com sucesso!" & vbCrLf & "ID registrado: " & novoID, vbInformation
    LimparCampos
    Exit Sub

ErroTratado:
    MsgBox "Erro (" & Err.Number & "): " & Err.Description, vbCritical
End Sub

' =====================================================
' LOAD DO FORM
' =====================================================
Private Sub UserForm_Initialize()
    Dim ws As Worksheet, cel As Range, ultimo As Long
    Set ws = ThisWorkbook.Sheets("AUTORIZADOS")
    
    cmbAGENTEDECARGAS.Clear
    ultimo = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For Each cel In ws.Range("A2:A" & ultimo)
        If Trim(cel.Value) <> "" Then cmbAGENTEDECARGAS.AddItem cel.Value
    Next cel

    ' Melhorias visuais
    Me.BackColor = RGB(240, 248, 255)        ' Fundo do Form
    cmbAGENTEDECARGAS.BackColor = RGB(255, 255, 200)
    cmbAGENTEDECARGAS.ListRows = 10          ' aumenta número de itens visíveis
End Sub

' =====================================================
' LIMPAR CAMPOS
' =====================================================
Private Sub LimparCampos()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Then ctrl.Value = ""
    Next ctrl
    txtAWB.SetFocus
End Sub



