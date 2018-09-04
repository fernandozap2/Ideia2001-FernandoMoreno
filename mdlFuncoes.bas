Attribute VB_Name = "mdlFuncoes"
Option Explicit

Public Const MARGEM = 100
Public Const MIN_WIDTH = 4200
Public Const MIN_HEIGHT = 2200

Public Enum enIdxLabels
    IDX_CENTRO = 1
    IDX_TOPO
    IDX_BAIXO
    IDX_ESQUERDA
    IDX_DIREITA
End Enum

Public Type typeClick
    Contador            As Integer
    DataHora            As Date
    TempoUltimoClick    As Double
    Label               As String
End Type

Public arrInfoClick() As typeClick

'---------------------------------------------------------------------------------------
' Autor      :  Fernando Moreno
' Data       :  23/8/2018
' Módulo     :  mdlFuncoes
' Função     :  msgErro
' Objetivo   :  Manipular as mensagens de erro
' Retorno    :
' Manutenções:
'---------------------------------------------------------------------------------------
Public Sub MsgErro(ByVal erro As ErrObject, ByVal nomeFuncao As String)

    MsgBox "Erro! " & erro.Number & " - " & erro.Description, vbCritical, "Ideia2001:Fernando - " & nomeFuncao
End Sub
'---------------------------------------------------------------------------------------
' Autor      :  Fernando Moreno
' Data       :  3/9/2018
' Módulo     :  mdlFuncoes
' Função     :  infoArrayColetado
' Objetivo   :
' Retorno    :
' Manutenções:
'---------------------------------------------------------------------------------------
Public Function ArmazenarDadosArray(ByVal lblClick As String) As Boolean
    
    Dim nomeFuncao  As String
    Dim countArray  As Integer
    
    On Error GoTo ErrorHandler
    nomeFuncao = "infoArrayColetado"
    
    countArray = UBound(arrInfoClick) + 1
    ReDim Preserve arrInfoClick(countArray)
    
    With arrInfoClick(countArray)
        .Contador = countArray
        .DataHora = Now()
        .Label = lblClick
        If countArray > 1 Then
            .TempoUltimoClick = DateDiff("s", arrInfoClick(countArray - 1).DataHora, arrInfoClick(countArray).DataHora)
        End If
    End With
    
    Exit Function
    
ErrorHandler:
    MsgErro Err, nomeFuncao
    
End Function

'---------------------------------------------------------------------------------------
' Autor      :  Fernando Moreno
' Data       :  3/9/2018
' Módulo     :  mdlFuncoes
' Função     :  infoArrayColetado
' Objetivo   :
' Retorno    :
' Manutenções:
'---------------------------------------------------------------------------------------
Public Function InfoArrayColetado() As Boolean
    
    Dim nomeFuncao   As String
    
    On Error GoTo ErrorHandler
    nomeFuncao = "infoArrayColetado"
    
    frmInfoArrayClick.Show 1
    
    Exit Function
    
ErrorHandler:
    MsgErro Err, nomeFuncao
    
End Function
