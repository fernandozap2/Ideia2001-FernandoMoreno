VERSION 5.00
Begin VB.Form frmIdeia2001_FM 
   Caption         =   "Ideia2001:Fernando"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label"
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   375
   End
End
Attribute VB_Name = "frmIdeia2001_FM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Dim nomeFuncao      As String
    Dim numCtrls        As Integer
    Dim i               As Integer
    
    'Criando labels
    If lbl.UBound <> 0 Then
        For i = 1 To 4
            Load lbl(lbl.UBound + 1)
            lbl(lbl.UBound).Visible = True
            lbl(lbl.UBound).Alignment = 2 'Centralizado
        Next i
    End If
    
    'Inicializando variáveis
    ReDim arrInfoClick(0)
    
    Exit Sub
    
ErrorHandler:
    MsgErro Err, nomeFuncao
    
End Sub

Private Sub Form_Resize()

    Dim nomeFuncao  As String
    Dim i           As Integer
    
    'Reposicionamento
    For i = 1 To lbl.UBound
    
        Select Case i
        Case IDX_CENTRO
            lbl(i).Move (Me.ScaleWidth - lbl(i).Width) / 2, (Me.ScaleHeight - lbl(i).Height) / 2
            lbl(i).Caption = "Centro"
        Case IDX_TOPO
            lbl(i).Move (Me.ScaleWidth - lbl(i).Width) / 2, MARGEM
            lbl(i).Caption = "Topo"
        Case IDX_BAIXO
            lbl(i).Move (Me.ScaleWidth - lbl(i).Width) / 2, (Me.ScaleHeight - lbl(i).Height - MARGEM)
            lbl(i).Caption = "Baixo"
        Case IDX_ESQUERDA
            lbl(i).Move MARGEM * 2, (Me.ScaleHeight - lbl(i).Height) / 2
            lbl(i).Caption = "Esquerda"
        Case IDX_DIREITA
            lbl(i).Move (Me.ScaleWidth - lbl(i).Width - MARGEM), Val((Me.ScaleHeight - lbl(i).Height) / 2)
            lbl(i).Caption = "Direita"
        End Select
        
    Next i
    
    'Limitadores
    If Me.WindowState = 0 Then
        If Width < MIN_WIDTH Then Width = MIN_WIDTH
        If Height < MIN_HEIGHT Then Height = MIN_HEIGHT
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgErro Err, nomeFuncao
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If MsgBox("Tem certeza que deseja sair?", vbQuestion + vbYesNo, "Ideia2001:Fernando - Saindo") = vbNo Then
        Cancel = True
    Else
        End
    End If
    
End Sub

Private Sub lbl_Click(Index As Integer)

    Dim i As Integer
    
    'Mudar de cor
    For i = 1 To lbl.UBound
        lbl(i).BackStyle = 1 'Sólido
        If i = Index Then
            lbl(i).BackColor = &H80000010 'Escuro
        Else
            lbl(i).BackColor = &H8000000F 'Claro
        End If
    Next
    
    'Armazenar dados no array
    Select Case Index
    Case IDX_CENTRO
        ArmazenarDadosArray lbl(Index).Caption
        InfoArrayColetado
    Case Else
        ArmazenarDadosArray lbl(Index).Caption
    End Select

End Sub
