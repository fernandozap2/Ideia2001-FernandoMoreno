VERSION 5.00
Begin VB.Form frmInfoArrayClick 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ideia2001:Fernando - Info Cliques"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstInfoClick 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmInfoArrayClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    lstInfoClick.Clear
    
    For i = 1 To UBound(arrInfoClick)
        lstInfoClick.AddItem (arrInfoClick(i).Contador & _
                              " - " & arrInfoClick(i).DataHora & _
                              " - " & arrInfoClick(i).Label & _
                              IIf(i > 1, " - depois de " & arrInfoClick(i).TempoUltimoClick & " seg.", ""))
    Next

End Sub
