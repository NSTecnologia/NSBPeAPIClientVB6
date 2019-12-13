VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbTpConteudo 
      Height          =   315
      ItemData        =   "frmPrincipal.frx":0000
      Left            =   8160
      List            =   "frmPrincipal.frx":000D
      TabIndex        =   14
      Text            =   "json"
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar Documento para Processamento >>>>>>"
      Height          =   615
      Left            =   6600
      TabIndex        =   7
      Top             =   5160
      Width           =   3735
   End
   Begin VB.TextBox txtConteudo 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   1080
      Width           =   10215
   End
   Begin VB.TextBox txtResult 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   6120
      Width           =   10215
   End
   Begin VB.ComboBox cbTpDown 
      Height          =   315
      ItemData        =   "frmPrincipal.frx":0021
      Left            =   120
      List            =   "frmPrincipal.frx":0034
      TabIndex        =   4
      Text            =   "XP"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CheckBox checkExibir 
      Caption         =   "Exibir PDF"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox txtCNPJ 
      Height          =   315
      Left            =   5520
      TabIndex        =   2
      Top             =   360
      Width           =   4815
   End
   Begin VB.TextBox txtCaminho 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "./Notas/"
      Top             =   360
      Width           =   5295
   End
   Begin VB.TextBox txtTpAmb 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Text            =   "2"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Token:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Conteudo"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resposta do Servidor"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label13 
      Caption         =   "Tipo de Download:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ:"
      Height          =   195
      Left            =   5520
      TabIndex        =   9
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Ambiente:"
      Height          =   195
      Left            =   2400
      TabIndex        =   8
      Top             =   4800
      Width           =   1290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnviar_Click()
    On Error GoTo SAI
    Dim retorno As String
    
    If (txtCaminho.Text <> "") And (txtConteudo.Text <> "") And (cbTpConteudo.Text <> "") And (cbTpDown.Text <> "") And (txtTpAmb.Text <> "") Then
    
        retorno = emitirBPeSincrono(txtConteudo.Text, cbTpConteudo.Text, txtCNPJ.Text, cbTpDown.Text, txtTpAmb.Text, txtCaminho.Text, checkExibir.Value)
        txtResult.Text = retorno
    Else
        MsgBox ("Todos os campos devem ser preenchidos")
    End If
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleBPeAPI

End Sub
