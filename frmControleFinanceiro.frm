VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmControleFinanceiro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle Financeiro"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19755
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmControleFinanceiro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   19755
   Begin VB.Frame fra_Usuario 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dados do Lançamento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7404&
      Height          =   1980
      Left            =   45
      TabIndex        =   16
      Top             =   -15
      Width           =   16710
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   12585
         MaxLength       =   20
         TabIndex        =   5
         Text            =   "0,00"
         Top             =   1365
         Width           =   1470
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2055
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1365
         Width           =   10395
      End
      Begin VB.ComboBox cmbCategSubCateg 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3885
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   510
         Width           =   12675
      End
      Begin VB.ComboBox cmbTipo 
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   150
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   510
         Width           =   1785
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   390
         Left            =   16110
         Picture         =   "frmControleFinanceiro.frx":1486
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1365
         Width           =   420
      End
      Begin VB.TextBox txtBusca 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2055
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         Top             =   510
         Width           =   1710
      End
      Begin VB.ComboBox cmbLancamentoFuturo 
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   14160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1365
         Width           =   1860
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   405
         Left            =   150
         TabIndex        =   3
         Top             =   1365
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Descrição"
         Height          =   240
         Index           =   3
         Left            =   2055
         TabIndex        =   23
         Top             =   1110
         Width           =   960
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Valor"
         Height          =   240
         Left            =   12585
         TabIndex        =   22
         Top             =   1110
         Width           =   495
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Categoria / Sub-Categoria"
         Height          =   240
         Index           =   1
         Left            =   3885
         TabIndex        =   21
         Top             =   240
         Width           =   2610
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Transação"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   20
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data"
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   19
         Top             =   1110
         Width           =   465
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Busca"
         Height          =   240
         Index           =   4
         Left            =   2055
         TabIndex        =   18
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lançamento Futuro"
         Height          =   240
         Index           =   5
         Left            =   14160
         TabIndex        =   17
         Top             =   1110
         Width           =   1920
      End
   End
   Begin VB.Frame fra_Botoes 
      BackColor       =   &H00FFFFFF&
      Height          =   1980
      Left            =   16815
      TabIndex        =   15
      Top             =   -15
      Width           =   2925
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   975
         Picture         =   "frmControleFinanceiro.frx":1D50
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar Campos"
         Top             =   135
         Width           =   950
      End
      Begin VB.CommandButton cmdSalvar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Salvar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   975
         Picture         =   "frmControleFinanceiro.frx":28D2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salvar Registro"
         Top             =   1035
         Width           =   950
      End
      Begin VB.CommandButton cmdConsultar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consultar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   30
         Picture         =   "frmControleFinanceiro.frx":33DC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir Registro"
         Top             =   150
         Width           =   950
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   1935
         Picture         =   "frmControleFinanceiro.frx":3EE6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir Dados Selecionados"
         Top             =   135
         Width           =   950
      End
      Begin VB.CommandButton cmdBackup 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Backup"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   30
         Picture         =   "frmControleFinanceiro.frx":5F58
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fazer Backup"
         Top             =   1035
         Width           =   950
      End
      Begin VB.CommandButton cmdRestaurar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Restaurar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   1935
         Picture         =   "frmControleFinanceiro.frx":6822
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Restaurar Backup"
         Top             =   1035
         Width           =   950
      End
   End
   Begin MSComDlg.CommonDialog botaoImportacao 
      Left            =   14115
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwConsulta 
      Height          =   6870
      Left            =   0
      TabIndex        =   14
      Top             =   1980
      Width           =   19740
      _ExtentX        =   34819
      _ExtentY        =   12118
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmControleFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Private Item                  As ListItem
   
   Private VALOR_CONSULTADO      As Double
   Private VALOR_A_INCLUIR       As Double
   
   Private IL                    As Long
   
   Private LANCAMENTO_FUTURO     As Boolean
   Private INCLUINDO_ITEM        As Boolean
   
   Private ARQBKP                As String
   
   Private Const GRD_TIPO        As Integer = 1
   Private Const GRD_CATEG       As Integer = 2
   Private Const GRD_SUB_CATEG   As Integer = 3
   Private Const GRD_DATA        As Integer = 4
   Private Const GRD_DESCRICAO   As Integer = 5
   Private Const GRD_VALOR       As Integer = 6
   Private Const GRD_STATUS      As Integer = 7
   Private Const GRD_GRADE       As Integer = 8
      
   Private Reg_Envio             As String * 247
   Private TAB_TIPO              As String * 20
   Private TAB_CATEG             As String * 50
   Private TAB_SUB_CATEG         As String * 50
   Private TAB_DATA              As String * 10
   Private TAB_DESCRICAO         As String * 100
   Private TAB_VALOR             As String * 16
   Private TAB_STATUS            As String * 1
   
'  RELATÓRIO ORÇAMENTO BÁSICO
   Private DATA_BASE             As String
   
   Private RsCreditos            As New ADODB.Recordset
   Private RsDebitos             As New ADODB.Recordset
   Private RsValores             As New ADODB.Recordset
      
   Private CREDITO               As Double
   Private GASTOS_ESSENCIAIS     As Double
   Private INVESTIMENTOS_DIVIDAS As Double
   Private DESEJOS_PESSOAIS      As Double
   
Private Function SALDO_EM(Optional Data As String) As Double
   Dim Rs_Saldo   As New ADODB.Recordset
   
   CmdSql = "SELECT SUM(CASE WHEN TIPO = 'CRÉDITO' THEN VALOR ELSE VALOR * -1 END) VALOR" & vbCr
   CmdSql = CmdSql & "FROM CONTROLE_FINANCEIRO" & vbCr
   CmdSql = CmdSql & "WHERE STATUS = ''" & vbCr
   If Data <> "" Then CmdSql = CmdSql & "  AND DATA <= " & Data
   CMySql.Consulta CmdSql, Rs_Saldo
   
   SALDO_EM = IIf(IsNull(Rs_Saldo("VALOR")) = True, 0, Format(Rs_Saldo("VALOR"), "###,##0.00"))
End Function

Private Sub TABELA(Optional LimpaCampo As Boolean)
   Reg_Envio = Space(Len(Reg_Envio))

   If LimpaCampo Then
      TAB_TIPO = Space(Len(TAB_TIPO))
      TAB_CATEG = Space(Len(TAB_CATEG))
      TAB_SUB_CATEG = Space(Len(TAB_SUB_CATEG))
      TAB_DATA = Space(Len(TAB_DATA))
      TAB_DESCRICAO = Space(Len(TAB_DESCRICAO))
      TAB_VALOR = Space(Len(TAB_VALOR))
      TAB_STATUS = Space(Len(TAB_STATUS))
   End If
   
   Reg_Envio = TAB_TIPO & TAB_CATEG & TAB_SUB_CATEG & TAB_DATA & TAB_DESCRICAO & TAB_VALOR & TAB_STATUS
End Sub

Private Sub MontaLvw()
   lvwConsulta.ListItems.Clear
   lvwConsulta.ColumnHeaders.Clear
   lvwConsulta.Gridlines = True
   lvwConsulta.ColumnHeaders.Add , , "", 0
   lvwConsulta.ColumnHeaders.Add , , "Tipo", 1000, lvwColumnCenter
   lvwConsulta.ColumnHeaders.Add , , "Categoria", 2200
   lvwConsulta.ColumnHeaders.Add , , "Subcategoria", 3900
   lvwConsulta.ColumnHeaders.Add , , "Data", 1250, lvwColumnCenter
   lvwConsulta.ColumnHeaders.Add , , "Descrição", 8300
   lvwConsulta.ColumnHeaders.Add , , "Valor", 1100, lvwColumnRight
   lvwConsulta.ColumnHeaders.Add , , "Lanç Fut", 1100, lvwColumnCenter
   lvwConsulta.ColumnHeaders.Add , , "", 400, lvwColumnCenter
   lvwConsulta.View = lvwReport
End Sub

Private Sub cmbCategSubCateg_GotFocus()
   Sendkeys "{F4}"
   cmbCategSubCateg.BackColor = QBColor(14)
End Sub

Private Sub cmbCategSubCateg_LostFocus()
   cmbCategSubCateg.BackColor = QBColor(15)
End Sub

Private Sub cmbLancamentoFuturo_GotFocus()
   cmbLancamentoFuturo.BackColor = QBColor(14)
End Sub

Private Sub cmbLancamentoFuturo_LostFocus()
   cmbLancamentoFuturo.BackColor = QBColor(15)
End Sub

Private Sub cmbTipo_Click()
   txtBusca = ""
   
   If cmbTipo = "" Then
      txtBusca.Locked = True
   Else
      txtBusca.Locked = False
   End If
   
   CmdSql = "SELECT * FROM TIPO_MOVIMENTACAO" & vbCr
   CmdSql = CmdSql & "WHERE TIPO = " & PoeAspas(cmbTipo)
   CMySql.Consulta CmdSql, Rs
   
   cmbCategSubCateg.Clear
   cmbCategSubCateg.AddItem ""
      
   Do While Not Rs.EOF
      cmbCategSubCateg.AddItem Trim(Rs("CATEG")) & " - " & Trim(Rs("SUB_CATEG"))
   Rs.MoveNext
   Loop
   cmbCategSubCateg.ListIndex = 0
End Sub

Private Sub cmbTipo_GotFocus()
    cmbTipo.BackColor = QBColor(14)
End Sub

Private Sub cmbTipo_LostFocus()
   cmbTipo.BackColor = QBColor(15)
End Sub

Private Sub cmdAdd_Click()

   With lvwConsulta
      For IL = 1 To .ListItems.Count
         If .ListItems(IL).SubItems(GRD_GRADE) = "A" Then
            MsgBoxTabum Me, "SISTEMA EM MODO ALTERAÇÃO, NÃO PODE SER INCLUÍDO"
            cmbTipo.SetFocus
            Exit Sub
         End If
      Next IL
   End With
   
   If cmbTipo = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O TIPO"
      cmbTipo.SetFocus
      Exit Sub
   End If
      
   If cmbCategSubCateg = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER UMA SUBCATEGORIA"
      txtBusca.SetFocus
      Exit Sub
   End If

   If mskData.ClipText = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER DATA"
      mskData.SetFocus
      Exit Sub
   End If
     
   If Cdblx(txtValor) = 0 Then
      MsgBoxTabum Me, "FAVOR PREENCHER O VALOR"
      txtValor.SetFocus
      Exit Sub
   End If
      
   Set Item = lvwConsulta.ListItems.Add(, , "")
   Item.Bold = True
   Item.ForeColor = RGB(0, 0, 250)
   Item.SubItems(GRD_TIPO) = cmbTipo
   Item.SubItems(GRD_CATEG) = Tira_Traco(cmbCategSubCateg, 1)
   Item.SubItems(GRD_SUB_CATEG) = Tira_Traco(cmbCategSubCateg, 2)
   Item.SubItems(GRD_DATA) = mskData
   Item.SubItems(GRD_DESCRICAO) = Trim(txtDescricao)
   Item.SubItems(GRD_VALOR) = Trim(txtValor)
   Item.SubItems(GRD_STATUS) = IIf(cmbLancamentoFuturo = "SIM", "S", "")
   Item.SubItems(GRD_GRADE) = "I"
            
   lvwConsulta.ListItems(lvwConsulta.ListItems.Count).EnsureVisible
   lvwConsulta.ListItems(lvwConsulta.ListItems.Count).Selected = True
   lvwConsulta.SetFocus
         
   DoEvents
   Refresh
         
   CALCULA_ITENS_PARA_INCLUIR
         
   MsgBoxTabum Me, cmbTipo & " INCLUÍDO COM SUCESSO"
         
   cmdSalvar.Enabled = True
   cmbTipo.SetFocus
End Sub

Private Sub CALCULA_ITENS_PARA_INCLUIR()

'  TOTAL ITENS PARA INCLUIR
   VALOR_A_INCLUIR = 0
   
   With lvwConsulta
      For IL = 1 To .ListItems.Count
         Select Case .ListItems(IL).SubItems(GRD_GRADE)
         Case "I"
            If .ListItems(IL).SubItems(GRD_TIPO) = "CRÉDITO" Then
               VALOR_A_INCLUIR = VALOR_A_INCLUIR - (.ListItems(IL).SubItems(GRD_VALOR))
            Else
               VALOR_A_INCLUIR = VALOR_A_INCLUIR + (.ListItems(IL).SubItems(GRD_VALOR))
            End If
         End Select
      Next IL
   End With
         
   frmControleFinanceiro.Caption = "Orçamento Familiar - Valores para Incluir: " & Format(VALOR_A_INCLUIR, "###,##0.00")

End Sub

Private Sub cmdBackup_Click()
   Form_Load
   
   If lvwConsulta.ListItems.Count = 0 Then
      MsgBoxTabum Me, "NENHUM REGISTRO PARA SER IMPORTADO"
      Exit Sub
   End If
   
   ARQBKP = ""
   ARQBKP = "CONTROLE_FINANCEIRO_" & Format(Date, "YYYYMMDD") & "_" & Format(Time, "HHMMSS") & ".BKP"
   
   cmdBackup.Enabled = False
   Open "C:\ARQUIVOS GERAIS\PROGRAMAS\MTABUM\BACKUP\" & ARQBKP For Output As #1
      
   CmdSql = "SELECT * FROM CONTROLE_FINANCEIRO" & vbCr
   CmdSql = CmdSql & "ORDER BY DATA DESC"
   CMySql.Consulta CmdSql, Rs
   
   Do While Not Rs.EOF
      TABELA True
         TAB_TIPO = Trim(Rs("TIPO"))
         TAB_CATEG = Trim(Rs("CATEG"))
         TAB_SUB_CATEG = Trim(Rs("SUB_CATEG"))
         TAB_DATA = Trim(Rs("DATA"))
         TAB_DESCRICAO = Trim(Rs("DESCRICAO"))
         TAB_VALOR = Format((Rs("VALOR")), "###,##0.00")
         TAB_STATUS = Trim(Rs("STATUS"))
      TABELA
      Print #1, Reg_Envio
   
   Rs.MoveNext
   Loop
   
   Close #1
   MsgBoxTabum Me, "BACKUP EFETUADO COM SUCESSO EM: " & vbCr & "C:\ARQUIVOS GERAIS\PROGRAMAS\MTABUM\BACKUP\" & ARQBKP
      
   cmdBackup.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
   cmbTipo.ListIndex = 1
   cmbTipo.ListIndex = 0
   cmbCategSubCateg.ListIndex = 0
   txtBusca = ""
   mskData = "__/__/____"
   txtDescricao = ""
   txtValor = "0,00"
   cmdSalvar.Enabled = False
   Form_Load
   cmbLancamentoFuturo.ListIndex = 1
   cmbTipo.SetFocus
End Sub

Private Sub cmdConsultar_Click()
   MontaLvw
   
   cmdConsultar.Enabled = False
   cmdSalvar.Enabled = False
         
   CmdSql = "SELECT * FROM CONTROLE_FINANCEIRO" & vbCr
   CmdSql = CmdSql & "WHERE TIPO <> ''" & vbCr
   
   If cmbTipo <> "" Then CmdSql = CmdSql & "  AND TIPO = " & PoeAspas(cmbTipo) & vbCr
   If cmbCategSubCateg <> "" Then CmdSql = CmdSql & "  AND CATEG = " & PoeAspas(Tira_Traco(cmbCategSubCateg, 1)) & vbCr
   If cmbCategSubCateg <> "" Then CmdSql = CmdSql & "  AND SUB_CATEG = " & PoeAspas(Tira_Traco(cmbCategSubCateg, 2)) & vbCr
    
'  DATA
   If mskData.ClipText <> "" Then
      If mskData.ClipText <> "" And Trim(txtDescricao) <> "" And Len(Trim(txtDescricao)) <= 2 Then
         Select Case Trim(txtDescricao)
         Case "="
            CmdSql = CmdSql & "  AND DATA = " & Format(mskData, "YYYYMMDD") & vbCr
         Case ">="
            CmdSql = CmdSql & "  AND DATA >= " & Format(mskData, "YYYYMMDD") & vbCr
         Case "<="
            CmdSql = CmdSql & "  AND DATA <= " & Format(mskData, "YYYYMMDD") & vbCr
         Case Else
            CmdSql = CmdSql & "  AND DATA = " & Format(mskData, "YYYYMMDD") & vbCr
         End Select
      Else
         CmdSql = CmdSql & "  AND DATA >= " & Format(mskData, "YYYYMMDD") & vbCr
      End If
   Else
      CmdSql = CmdSql & "  AND DATA >= " & Format(Date - 730, "YYYYMMDD") & vbCr
   End If
'  FIM

   If Trim(txtDescricao) <> "" And Len(Trim(txtDescricao)) > 2 Then CmdSql = CmdSql & "  AND DESCRICAO LIKE " & PoeAspas("%" & txtDescricao & "%") & vbCr
   
'  VALOR
   If Trim(txtValor) <> "0,00" Then
      If Trim(txtDescricao) <> "" And Len(Trim(txtDescricao)) <= 2 Then
         Select Case Trim(txtDescricao)
         Case "="
            CmdSql = CmdSql & "  AND VALOR = " & Str(txtValor) & vbCr
         Case ">="
            CmdSql = CmdSql & "  AND VALOR >= " & Str(txtValor) & vbCr
         Case "<="
            CmdSql = CmdSql & "  AND VALOR <= " & Str(txtValor) & vbCr
         Case Else
            CmdSql = CmdSql & "  AND VALOR = " & Str(txtValor) & vbCr
         End Select
      Else
         CmdSql = CmdSql & "  AND VALOR = " & Str(txtValor) & vbCr
      End If
   End If
'  FIM VALOR
   
   If cmbLancamentoFuturo = "SIM" Then CmdSql = CmdSql & "  AND STATUS = 'S'" & vbCr
   If cmbLancamentoFuturo = "NÃO" Then CmdSql = CmdSql & "  AND STATUS = ''" & vbCr
   
   CmdSql = CmdSql & "ORDER BY DATA DESC,DESCRICAO"
   CMySql.Consulta CmdSql, Rs
        
   VALOR_CONSULTADO = 0
   
   Do While Not Rs.EOF
      If Trim(Rs("TIPO")) = "DÉBITO" Then
         VALOR_CONSULTADO = VALOR_CONSULTADO - Rs("VALOR")
      Else
         VALOR_CONSULTADO = VALOR_CONSULTADO + Rs("VALOR")
      End If
      
      Set Item = lvwConsulta.ListItems.Add(, , "")
      Item.Bold = True
      Item.ForeColor = RGB(0, 0, 250)
      Item.SubItems(GRD_TIPO) = Trim(Rs("TIPO"))
      Item.SubItems(GRD_CATEG) = Trim(Rs("CATEG"))
      Item.SubItems(GRD_SUB_CATEG) = Trim(Rs("SUB_CATEG"))
      Item.SubItems(GRD_DATA) = Trim(Rs("DATA"))
      Item.SubItems(GRD_DESCRICAO) = Trim(Rs("DESCRICAO"))
      Item.SubItems(GRD_VALOR) = Format(Trim(Rs("VALOR")), "###,##0.00")
      Item.SubItems(GRD_STATUS) = Trim(Rs("STATUS"))
      Item.SubItems(GRD_GRADE) = ""
   Rs.MoveNext
   Loop
   
   cmdConsultar.Enabled = True
   frmControleFinanceiro.Caption = "Orçamento Familiar - Saldo Atual: " & Format(SALDO_EM, "###,##0.00") & IIf(cmbTipo <> "" Or cmbCategSubCateg <> "" Or mskData.ClipText <> "" Or txtDescricao <> "" Or txtValor <> "0,00" Or cmbLancamentoFuturo <> "NÃO", " - Total da Consulta: " & Format(VALOR_CONSULTADO, "###,##0.00"), "")
End Sub

Private Sub cmdImprimir_Click()
   Dim I As Integer
   Dim MES_REF As String
   
   Dim DATA_INICIAL As String
   Dim DATA_FINAL   As String
   Dim SALDO        As Double
     
   cmdImprimir.Enabled = False
   Form_Load
   
   I = 0
   MES_REF = ""
   SALDO = 0
   
   CmdSql = "DELETE FROM REL_CONTROLE_FINANCEIRO"
   CMySql.Executa CmdSql, True
   
   CmdSql = "SELECT * FROM CONTROLE_FINANCEIRO"
   CMySql.Consulta CmdSql, Rs
      
   If Rs.EOF Then
      MsgBoxTabum Me, "NENHUM REGISTRO ENCONTRADO"
      cmdImprimir.Enabled = True
      Exit Sub
   End If
   
   If MsgBox("DESEJA IMPRIMIR TAMBÉM OS LANÇAMENTOS FUTUROS?", vbYesNo + vbQuestion + vbDefaultButton2, "RELATÓRIO") = vbYes Then
      LANCAMENTO_FUTURO = True
   Else
      LANCAMENTO_FUTURO = False
   End If
   
   CmdSql = "SELECT MAX(DATA)MAXIMA_DATA" & vbCr
   CmdSql = CmdSql & "FROM CONTROLE_FINANCEIRO" & vbCr
   CmdSql = CmdSql & "WHERE TIPO <> ''" & vbCr
   If LANCAMENTO_FUTURO = False Then CmdSql = CmdSql & "  AND STATUS <> 'S'" & vbCr
   If mskData.ClipText <> "" Then CmdSql = CmdSql & "  AND DATA <= " & Format(mskData, "YYYYMMDD")
   CMySql.Consulta CmdSql, Rs

   If Not Rs.EOF Then
      DATA_INICIAL = Format(FirstDay(LastDay(Rs("MAXIMA_DATA")) - 364), "YYYYMMDD")
      DATA_FINAL = Format(LastDay(Rs("MAXIMA_DATA")), "YYYYMMDD")
   Else
      cmdImprimir.Enabled = True
      Exit Sub
   End If
   
'  SALDO EM
   SALDO = SALDO_EM(Format(FirstDay(LastDay(Rs("MAXIMA_DATA")) - 364) - 1, "YYYYMMDD"))
'  FIM
   
   CmdSql = "SELECT TIPO,CATEG,SUB_CATEG,DATE_FORMAT(DATA,'%Y/%m') DATA,SUM(VALOR)VALOR" & vbCr
   CmdSql = CmdSql & "FROM CONTROLE_FINANCEIRO" & vbCr
   CmdSql = CmdSql & "WHERE DATA BETWEEN " & DATA_INICIAL & " AND " & DATA_FINAL & vbCr
   If LANCAMENTO_FUTURO = False Then CmdSql = CmdSql & "  AND STATUS <> 'S'" & vbCr
   CmdSql = CmdSql & "GROUP BY TIPO,CATEG,SUB_CATEG,DATA" & vbCr
   CmdSql = CmdSql & "ORDER BY DATA DESC"
   CMySql.Consulta CmdSql, Rs
   
   Do While Not Rs.EOF
      If MES_REF <> Month(Rs("DATA")) Then I = I + 1
      
      CmdSql = "INSERT INTO REL_CONTROLE_FINANCEIRO(TIPO,CATEG,SUB_CATEG,"
      CmdSql = CmdSql & "MES_" & Format(I, "00") & ",VALOR_" & Format(I, "00") & ")" & vbCr
      CmdSql = CmdSql & "VALUES("
      CmdSql = CmdSql & PoeAspas(Rs("TIPO")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("CATEG")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("SUB_CATEG")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("DATA")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR")) & ")"
      CMySql.Executa CmdSql, True
      
      MES_REF = Month(Rs("DATA"))
      
   Rs.MoveNext
   Loop
   
   I = 1
   MES_REF = ""
   
   CmdSql = "SELECT DISTINCT MES_01,MES_02,MES_03,MES_04,MES_05,MES_06,MES_07,MES_08,MES_09,MES_10,MES_11,MES_12" & vbCr
   CmdSql = CmdSql & "FROM REL_CONTROLE_FINANCEIRO"
   CMySql.Consulta CmdSql, Rs
   
   Do While Not Rs.EOF
      If Trim(Rs("MES_" & Format(I, "00"))) <> "" Then
         CmdSql = "UPDATE REL_CONTROLE_FINANCEIRO SET MES_" & Format(I, "00") & " = " & PoeAspas(Mid(UCase(MonthName(Mid(Rs("MES_" & Format(I, "00")), 6, 2))), 1, 3) & "/" & Mid(Rs("MES_" & Format(I, "00")), 1, 4))
         CMySql.Executa CmdSql, True
         
         I = I + 1
      End If
   Rs.MoveNext
   Loop
   
   CmdSql = "SELECT TIPO,CATEG,SUB_CATEG,MES_01,SUM(VALOR_01)VALOR_01,MES_02,SUM(VALOR_02)VALOR_02,MES_03,SUM(VALOR_03)VALOR_03,MES_04,SUM(VALOR_04)VALOR_04,MES_05,SUM(VALOR_05)VALOR_05,MES_06,SUM(VALOR_06)VALOR_06," & vbCr
   CmdSql = CmdSql & "                            MES_07,SUM(VALOR_07)VALOR_07,MES_08,SUM(VALOR_08)VALOR_08,MES_09,SUM(VALOR_09)VALOR_09,MES_10,SUM(VALOR_10)VALOR_10,MES_11,SUM(VALOR_11)VALOR_11,MES_12,SUM(VALOR_12)VALOR_12" & vbCr
   CmdSql = CmdSql & "FROM REL_CONTROLE_FINANCEIRO" & vbCr
   CmdSql = CmdSql & "GROUP BY TIPO,CATEG,SUB_CATEG,MES_01,MES_02,MES_03,MES_04,MES_05,MES_06,MES_07,MES_08,MES_09,MES_10,MES_11,MES_12" & vbCr
   CmdSql = CmdSql & "ORDER BY TIPO,CATEG,SUB_CATEG"
   CMySql.Consulta CmdSql, Rs
   
   CmdSql = "DELETE FROM REL_CONTROLE_FINANCEIRO" & vbCr
   CMySql.Executa CmdSql, True
   
   Do While Not Rs.EOF
      CmdSql = "INSERT INTO REL_CONTROLE_FINANCEIRO(TIPO,CATEG,SUB_CATEG,MES_01,VALOR_01,MES_02,VALOR_02,MES_03,VALOR_03,MES_04,VALOR_04,MES_05,VALOR_05,MES_06,VALOR_06,"
      CmdSql = CmdSql & "MES_07,VALOR_07,MES_08,VALOR_08,MES_09,VALOR_09,MES_10,VALOR_10,MES_11,VALOR_11,MES_12,VALOR_12)" & vbCr
      CmdSql = CmdSql & "VALUES("
      CmdSql = CmdSql & PoeAspas(Rs("TIPO")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("CATEG")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("SUB_CATEG")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_01")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_01")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_02")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_02")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_03")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_03")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_04")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_04")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_05")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_05")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_06")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_06")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_07")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_07")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_08")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_08")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_09")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_09")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_10")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_10")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_11")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_11")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_12")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_12")) & ")"
      CMySql.Executa CmdSql, True
   Rs.MoveNext
   Loop
     
'  CALCULANDO TOTAIS
   Dim total   As Double
   Dim j       As Integer
   Dim REG     As Integer
      
   I = 0
   total = 0
   REG = 1
   MES_REF = ""
   
   CmdSql = "DELETE FROM REL_AUXILIAR"
   CMySql.Executa CmdSql, True
   
   CmdSql = "SELECT DATE_FORMAT(DATA,'%Y/%m') DATA,SUM(CASE WHEN TIPO = 'CREDITO' THEN VALOR ELSE VALOR * -1 END) VALOR" & vbCr
   CmdSql = CmdSql & "FROM CONTROLE_FINANCEIRO" & vbCr
   CmdSql = CmdSql & "WHERE DATA BETWEEN " & DATA_INICIAL & " AND " & DATA_FINAL & vbCr
   If LANCAMENTO_FUTURO = False Then CmdSql = CmdSql & "  AND STATUS <> 'S'" & vbCr
   CmdSql = CmdSql & "GROUP BY DATA" & vbCr
   CmdSql = CmdSql & "ORDER BY DATA"
   CMySql.Consulta CmdSql, Rs
   
   Do While Not Rs.EOF
      REG = REG + 1
   Rs.MoveNext
   Loop
   
   CmdSql = "SELECT DATE_FORMAT(DATA,'%Y/%m') DATA,SUM(CASE WHEN TIPO = 'CREDITO' THEN VALOR ELSE VALOR * -1 END) VALOR" & vbCr
   CmdSql = CmdSql & "FROM CONTROLE_FINANCEIRO" & vbCr
   CmdSql = CmdSql & "WHERE DATA BETWEEN " & DATA_INICIAL & " AND " & DATA_FINAL & vbCr
   If LANCAMENTO_FUTURO = False Then CmdSql = CmdSql & "  AND STATUS <> 'S'" & vbCr
   CmdSql = CmdSql & "GROUP BY DATA" & vbCr
   CmdSql = CmdSql & "ORDER BY DATA"
   CMySql.Consulta CmdSql, Rs
      
   Do While Not Rs.EOF
      If MES_REF <> Month(Rs("DATA")) Then I = I + 1
          
      total = total + IIf(I = 1, Rs("VALOR") + SALDO, Rs("VALOR"))
          
      For j = 1 To 2
         CmdSql = "INSERT INTO REL_AUXILIAR(TIPO,MES_" & Format(REG - I, "00") & ",VALOR_" & Format(REG - I, "00") & ")" & vbCr
         CmdSql = CmdSql & "VALUES("
         
         If j = 1 Then
            CmdSql = CmdSql & PoeAspas("Saldo do Mês") & ","
            CmdSql = CmdSql & PoeAspas(Rs("DATA")) & ","
            CmdSql = CmdSql & Str(Rs("VALOR")) & ")"
         Else
            CmdSql = CmdSql & PoeAspas("Saldo Acumulado") & ","
            CmdSql = CmdSql & PoeAspas(Rs("DATA")) & ","
            CmdSql = CmdSql & Str(total) & ")"
         End If
         CMySql.Executa CmdSql, True
      Next j
      
      MES_REF = Month(Rs("DATA"))
      
   Rs.MoveNext
   Loop
  
   I = 1
   MES_REF = ""
   
   CmdSql = "SELECT DISTINCT MES_01,MES_02,MES_03,MES_04,MES_05,MES_06,MES_07,MES_08,MES_09,MES_10,MES_11,MES_12" & vbCr
   CmdSql = CmdSql & "FROM REL_AUXILIAR"
   CMySql.Consulta CmdSql, Rs
   
   Do While Not Rs.EOF
      If Trim(Rs("MES_" & Format(REG - I, "00"))) <> "" Then
         CmdSql = "UPDATE REL_AUXILIAR SET MES_" & Format(REG - I, "00") & " = " & PoeAspas(Mid(UCase(MonthName(Mid(Rs("MES_" & Format(REG - I, "00")), 6, 2))), 1, 3) & "/" & Mid(Rs("MES_" & Format(REG - I, "00")), 1, 4))
         CMySql.Executa CmdSql, True
         
         I = I + 1
      End If
   Rs.MoveNext
   Loop
     
   CmdSql = "SELECT TIPO,MES_01,SUM(VALOR_01)VALOR_01,MES_02,SUM(VALOR_02)VALOR_02,MES_03,SUM(VALOR_03)VALOR_03,MES_04,SUM(VALOR_04)VALOR_04,MES_05,SUM(VALOR_05)VALOR_05,MES_06,SUM(VALOR_06)VALOR_06," & vbCr
   CmdSql = CmdSql & "                            MES_07,SUM(VALOR_07)VALOR_07,MES_08,SUM(VALOR_08)VALOR_08,MES_09,SUM(VALOR_09)VALOR_09,MES_10,SUM(VALOR_10)VALOR_10,MES_11,SUM(VALOR_11)VALOR_11,MES_12,SUM(VALOR_12)VALOR_12" & vbCr
   CmdSql = CmdSql & "FROM REL_AUXILIAR" & vbCr
   CmdSql = CmdSql & "GROUP BY TIPO,MES_01,MES_02,MES_03,MES_04,MES_05,MES_06,MES_07,MES_08,MES_09,MES_10,MES_11,MES_12" & vbCr
   CmdSql = CmdSql & "ORDER BY TIPO DESC"
   CMySql.Consulta CmdSql, Rs
   
   CmdSql = "DELETE FROM REL_AUXILIAR" & vbCr
   CMySql.Executa CmdSql, True
   
   Do While Not Rs.EOF
      CmdSql = "INSERT INTO REL_AUXILIAR(TIPO,MES_01,VALOR_01,MES_02,VALOR_02,MES_03,VALOR_03,MES_04,VALOR_04,MES_05,VALOR_05,MES_06,VALOR_06,"
      CmdSql = CmdSql & "MES_07,VALOR_07,MES_08,VALOR_08,MES_09,VALOR_09,MES_10,VALOR_10,MES_11,VALOR_11,MES_12,VALOR_12)" & vbCr
      CmdSql = CmdSql & "VALUES("
      CmdSql = CmdSql & PoeAspas(Rs("TIPO")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_01")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_01")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_02")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_02")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_03")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_03")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_04")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_04")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_05")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_05")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_06")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_06")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_07")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_07")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_08")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_08")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_09")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_09")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_10")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_10")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_11")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_11")) & ","
      CmdSql = CmdSql & PoeAspas(Rs("MES_12")) & ","
      CmdSql = CmdSql & Str(Rs("VALOR_12")) & ")"
      CMySql.Executa CmdSql, True
   Rs.MoveNext
   Loop
'  FIM
     
'  SALDO EM
   If SALDO > 0 Then
      CmdSql = "UPDATE REL_CONTROLE_FINANCEIRO SET SALDO_EM = 'SALDO EM: " & CDate(Mid(DATA_INICIAL, 7, 2) & "/" & Mid(DATA_INICIAL, 5, 2) & "/" & Mid(DATA_INICIAL, 1, 4)) - 1 & " : " & Format(SALDO, "###,##0.00") & "'"
      CMySql.Executa CmdSql, True
   End If
'  FIM

   If mskData.ClipText = "" Then
      mskData = FirstDay(Date)
      RELATORIO_MENSAL
      mskData = "__/__/____"
   Else
      RELATORIO_MENSAL
   End If
   
   RELATORIO_CONTROLE_FINANCEIRO_BASICO
     
   With MDIPrincipal.CryRelatorio
      .DiscardSavedData = True
      .ProgressDialog = False
      .WindowControlBox = True
      .WindowControls = True
      .WindowShowSearchBtn = True
      .WindowShowPrintSetupBtn = True
      .WindowState = 2
      
      .ReportFileName = "C:\ARQUIVOS GERAIS\PROGRAMAS\MTABUM\RELSCRIPTS\Relatorios\MTABUM\REL_ORÇAMENTO_BÁSICO.rpt"
      .Destination = 0
      .Action = 1
      .PageZoom (160)
          
      .ReportFileName = "C:\ARQUIVOS GERAIS\PROGRAMAS\MTABUM\RELSCRIPTS\Relatorios\MTABUM\REL_ORÇAMENTO_ANALÍTICO.rpt"
      .Destination = 0
      .Action = 1
      .PageZoom (160)
      
      .ReportFileName = "C:\ARQUIVOS GERAIS\PROGRAMAS\MTABUM\RELSCRIPTS\Relatorios\MTABUM\REL_ORÇAMENTO.rpt"
      .Destination = 0
      .Action = 1
      .PageZoom (160)
   End With
   
   cmdImprimir.Enabled = True
End Sub

Public Function FirstDay(dtData As Date) As Date
   FirstDay = CDate("01/" & Month(dtData) & "/" & Year(dtData))
End Function

Public Function LastDay(dtData As Date) As Date
   On Error Resume Next
   
   Dim strUltDia As String
   Dim iDia As Integer
   
   iDia = 32
   Do
      iDia = iDia - 1
      strUltDia = CStr(iDia) & "/" & CStr(Month(dtData)) & "/" & CStr(Year(dtData))
   Loop While Not IsDate(strUltDia)
   
   LastDay = CDate(strUltDia)
End Function

Private Sub cmdRestaurar_Click()
   
   Dim ArquivoImportacao As String
   Dim Registro        As String
   
   botaoImportacao.FileName = ""
   botaoImportacao.DialogTitle = "Restaurar Backup"
   botaoImportacao.Filter = "Arquivos de texto(*.BKP)|*.bkp|"
   botaoImportacao.FilterIndex = 1
   botaoImportacao.ShowOpen
   
   If botaoImportacao.FileName <> "" Then
      ArquivoImportacao = botaoImportacao.FileName
      MousePointer = vbHourglass

      Dim Linha, comando As String
      Dim FileNum As Integer
      Dim I As Integer
      Dim total As Integer
            
      I = 0
      FileNum = 1
             
      Open ArquivoImportacao For Input As FileNum
      
      Do Until EOF(FileNum)
         Line Input #FileNum, Linha
         total = total + 1
      Loop
      Close #FileNum
      
      CmdSql = "DELETE FROM CONTROLE_FINANCEIRO"
      CMySql.Executa CmdSql, True
       
      Open ArquivoImportacao For Input As FileNum
     
      Do While Not EOF(FileNum)
         Line Input #FileNum, Registro
         CmdSql = "SELECT * FROM CONTROLE_FINANCEIRO" & vbCr
         CmdSql = CmdSql & "WHERE TIPO      = " & PoeAspas(Mid(Registro, 1, 20)) & vbCr
         CmdSql = CmdSql & "  AND CATEG     = " & PoeAspas(Mid(Registro, 21, 50)) & vbCr
         CmdSql = CmdSql & "  AND SUB_CATEG = " & PoeAspas(Mid(Registro, 71, 50)) & vbCr
         CmdSql = CmdSql & "  AND DATA      = " & Format(Mid(Registro, 121, 10), "YYYYMMDD") & vbCr
         CmdSql = CmdSql & "  AND DESCRICAO = " & PoeAspas(Mid(Registro, 131, 100)) & vbCr
         CmdSql = CmdSql & "  AND VALOR     = " & Str(Mid(Registro, 231, 16)) & vbCr
         CmdSql = CmdSql & "  AND STATUS    = " & PoeAspas(Mid(Registro, 247, 1)) & vbCr
         CmdSql = CmdSql & "  AND STATUS    = 'K'" 'COLOCADO PARA ANULAR ESTE SELECT, PARA INSERIR TODOS OS REGISTROS DO ARQUIVO DE BACKUP
         CMySql.Consulta CmdSql, Rs
         
         If Rs.EOF Then
            CmdSql = "INSERT INTO CONTROLE_FINANCEIRO(TIPO,CATEG,SUB_CATEG,DATA,DESCRICAO,VALOR,STATUS)" & vbCr
            CmdSql = CmdSql & "VALUES("
            CmdSql = CmdSql & PoeAspas(Mid(Registro, 1, 20)) & ","
            CmdSql = CmdSql & PoeAspas(Mid(Registro, 21, 50)) & ","
            CmdSql = CmdSql & PoeAspas(Mid(Registro, 71, 50)) & ","
            CmdSql = CmdSql & Format(Mid(Registro, 121, 10), "YYYYMMDD") & ","
            CmdSql = CmdSql & PoeAspas(Mid(Registro, 131, 100)) & ","
            CmdSql = CmdSql & Str(Mid(Registro, 231, 16)) & ","
            CmdSql = CmdSql & PoeAspas(Mid(Registro, 247, 1)) & ")"
            CMySql.Executa CmdSql, True
         End If
                   
      DoEvents
      Loop
  
   Close FileNum
         
   ArquivoImportacao = ""
   MsgBoxTabum Me, "Backup Restaurado com Sucesso!"
   MousePointer = vbDefault
   Form_Load
      
   End If
End Sub

Private Sub CmdSalvar_Click()
   cmdSalvar.Enabled = False
   
   With lvwConsulta
      For IL = 1 To .ListItems.Count
         Select Case .ListItems(IL).SubItems(GRD_GRADE)
         Case "I"
ALTERAR:
            CmdSql = "INSERT INTO CONTROLE_FINANCEIRO(TIPO,CATEG,SUB_CATEG,DATA,DESCRICAO,VALOR,STATUS)" & vbCr
            CmdSql = CmdSql & "VALUES("
            CmdSql = CmdSql & PoeAspas(.ListItems(IL).SubItems(GRD_TIPO)) & ","
            CmdSql = CmdSql & PoeAspas(.ListItems(IL).SubItems(GRD_CATEG)) & ","
            CmdSql = CmdSql & PoeAspas(.ListItems(IL).SubItems(GRD_SUB_CATEG)) & ","
            CmdSql = CmdSql & Format(.ListItems(IL).SubItems(GRD_DATA), "YYYYMMDD") & ","
            CmdSql = CmdSql & PoeAspas(.ListItems(IL).SubItems(GRD_DESCRICAO)) & ","
            CmdSql = CmdSql & Str(.ListItems(IL).SubItems(GRD_VALOR)) & ","
            CmdSql = CmdSql & PoeAspas(.ListItems(IL).SubItems(GRD_STATUS)) & ")"
            CMySql.Executa CmdSql, True
         Case "D", "A"
            CmdSql = "DELETE FROM CONTROLE_FINANCEIRO" & vbCr
            CmdSql = CmdSql & "WHERE TIPO      = " & PoeAspas(.ListItems(IL).SubItems(GRD_TIPO)) & vbCr
            CmdSql = CmdSql & "  AND CATEG     = " & PoeAspas(.ListItems(IL).SubItems(GRD_CATEG)) & vbCr
            CmdSql = CmdSql & "  AND SUB_CATEG = " & PoeAspas(.ListItems(IL).SubItems(GRD_SUB_CATEG)) & vbCr
            CmdSql = CmdSql & "  AND DATA      = " & Format(.ListItems(IL).SubItems(GRD_DATA), "YYYYMMDD") & vbCr
            CmdSql = CmdSql & "  AND DESCRICAO = " & PoeAspas(.ListItems(IL).SubItems(GRD_DESCRICAO)) & vbCr
            CmdSql = CmdSql & "  AND VALOR     = " & Str(.ListItems(IL).SubItems(GRD_VALOR)) & vbCr
            CmdSql = CmdSql & "  AND STATUS    = " & PoeAspas(IIf(cmbLancamentoFuturo = "SIM", "S", ""))
            CMySql.Executa CmdSql, True
            
            If .ListItems(IL).SubItems(GRD_GRADE) = "A" Then
               lvwConsulta.ListItems.Remove lvwConsulta.SelectedItem.Index
               Refresh
               DoEvents
               
               Set Item = lvwConsulta.ListItems.Add(, , "")
               Item.Bold = True
               Item.ForeColor = RGB(0, 0, 250)
               Item.SubItems(GRD_TIPO) = cmbTipo
               Item.SubItems(GRD_CATEG) = Tira_Traco(cmbCategSubCateg, 1)
               Item.SubItems(GRD_SUB_CATEG) = Tira_Traco(cmbCategSubCateg, 2)
               Item.SubItems(GRD_DATA) = mskData
               Item.SubItems(GRD_DESCRICAO) = Trim(txtDescricao)
               Item.SubItems(GRD_VALOR) = Trim(txtValor)
               Item.SubItems(GRD_STATUS) = IIf(cmbLancamentoFuturo = "SIM", "S", "")
               Item.SubItems(GRD_GRADE) = "I"
               
               Refresh
               DoEvents
               GoSub ALTERAR
            End If
         End Select
      Next IL
   End With
   
   cmdCancelar_Click
   cmdConsultar_Click
   cmdSalvar.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
      
   If KeyAscii = vbKeyReturn Then
      Sendkeys "{tab}"
   End If
End Sub

Private Sub Form_Load()
   MontaLvw

   cmbTipo.Clear
   cmbTipo.AddItem ""
   cmbTipo.AddItem "CRÉDITO"
   cmbTipo.AddItem "DÉBITO"
   cmbTipo.ListIndex = 0
   
   cmbLancamentoFuturo.Clear
   cmbLancamentoFuturo.AddItem ""
   cmbLancamentoFuturo.AddItem "SIM"
   cmbLancamentoFuturo.AddItem "NÃO"
   cmbLancamentoFuturo.ListIndex = 1
   
   cmdConsultar_Click
End Sub

Private Sub lvwConsulta_DblClick()
   If lvwConsulta.ListItems.Count = 0 Then Exit Sub
   INCLUINDO_ITEM = False
   
   Select Case Trim(lvwConsulta.SelectedItem.ListSubItems(GRD_GRADE))
   Case "I"
      cmbTipo = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_TIPO)
      cmbCategSubCateg = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_CATEG) & " - " & lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_SUB_CATEG)
      mskData = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_DATA)
      txtDescricao = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_DESCRICAO)
      txtValor = Format(lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_VALOR), "###,##0.00")
      cmbLancamentoFuturo = IIf(lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_STATUS) = "S", "SIM", "NÃO")
      
      lvwConsulta.ListItems.Remove lvwConsulta.SelectedItem.Index
      INCLUINDO_ITEM = True
   End Select
   
   If INCLUINDO_ITEM = False Then
      If lvwConsulta.ListItems.Count = 1 And Trim(lvwConsulta.SelectedItem.ListSubItems(GRD_GRADE)) = "" Then
         cmbTipo = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_TIPO)
         cmbCategSubCateg = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_CATEG) & " - " & lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_SUB_CATEG)
         mskData = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_DATA)
         txtDescricao = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_DESCRICAO)
         txtValor = Format(lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_VALOR), "###,##0.00")
         cmbLancamentoFuturo = IIf(lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_STATUS) = "S", "SIM", "NÃO")
         
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_GRADE) = "A"
         cmdSalvar.Enabled = True
         Exit Sub
      End If
      
      If lvwConsulta.ListItems.Count = 1 And Trim(lvwConsulta.SelectedItem.ListSubItems(GRD_GRADE)) = "A" Then
         cmbTipo.ListIndex = 0
         cmbCategSubCateg.ListIndex = 0
         mskData = "__/__/____"
         txtDescricao = ""
         txtValor = "0,00"
         cmbLancamentoFuturo.ListIndex = 0
         
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_GRADE) = ""
         cmdSalvar.Enabled = False
      End If
   End If
End Sub

Private Sub lvwConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
   If lvwConsulta.ListItems.Count = 0 Then Exit Sub
   
   Select Case KeyCode
   Case "46"
      Select Case lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_GRADE)
      Case "I"
         lvwConsulta.ListItems.Remove lvwConsulta.SelectedItem.Index
      Case ""
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_GRADE) = "D"
      Case "D"
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_GRADE) = ""
      End Select
   Case "45"
      If lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_STATUS) = "S" Then
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_STATUS) = ""
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_GRADE) = "A"
      Else
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_STATUS) = "S"
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_GRADE) = "A"
      End If
      
      cmdSalvar.Enabled = True
      
      CALCULA_ITENS_PARA_INCLUIR
      Exit Sub
   End Select
   
   cmdSalvar.Enabled = False
   
   With lvwConsulta
      For IL = 1 To .ListItems.Count
         If lvwConsulta.ListItems(IL).SubItems(GRD_GRADE) = "D" Or lvwConsulta.ListItems(IL).SubItems(GRD_GRADE) = "I" Then
            cmdSalvar.Enabled = True
            Exit Sub
         End If
      Next IL
   End With
   
End Sub

Private Sub mskData_GotFocus()
   Marca Me
   mskData.BackColor = QBColor(14)
End Sub

Private Sub mskData_LostFocus()
   mskData.BackColor = QBColor(15)
End Sub

Private Sub mskData_Validate(Cancel As Boolean)
   If mskData.ClipText <> "" Then
      If EData(mskData) = False Then
         MsgBoxTabum Me, "DATA INVÁLIDA"
         Cancel = True
         Marca Me
         mskData.SetFocus
      End If
   End If
End Sub

Private Sub txtBusca_Change()
        
   If txtBusca = "'" Then txtBusca = ""
     
   If Len(Trim(txtBusca)) > 0 Then
      CmdSql = "SELECT * FROM TIPO_MOVIMENTACAO" & vbCr
      CmdSql = CmdSql & "WHERE TIPO = " & PoeAspas(cmbTipo) & vbCr
      CmdSql = CmdSql & "  AND SUB_CATEG LIKE " & PoeAspas("%" & txtBusca & "%")
      CMySql.Consulta CmdSql, Rs
      
      cmbCategSubCateg.Clear
      cmbCategSubCateg.AddItem ""
         
      Do While Not Rs.EOF
         cmbCategSubCateg.AddItem Trim(Rs("CATEG")) & " - " & Trim(Rs("SUB_CATEG"))
      Rs.MoveNext
      Loop
      If cmbCategSubCateg.ListCount > 0 Then cmbCategSubCateg.ListIndex = 0
   Else
      CmdSql = "SELECT * FROM TIPO_MOVIMENTACAO" & vbCr
      CmdSql = CmdSql & "WHERE TIPO = " & PoeAspas(cmbTipo)
      CMySql.Consulta CmdSql, Rs
      
      cmbCategSubCateg.Clear
      cmbCategSubCateg.AddItem ""
         
      Do While Not Rs.EOF
         cmbCategSubCateg.AddItem Trim(Rs("CATEG")) & " - " & Trim(Rs("SUB_CATEG"))
      Rs.MoveNext
      Loop
      cmbCategSubCateg.ListIndex = 0
   End If
End Sub

Private Sub txtBusca_GotFocus()
   Marca Me
   txtBusca.BackColor = QBColor(14)
End Sub

Private Sub txtBusca_LostFocus()
   txtBusca.BackColor = QBColor(15)
End Sub

Private Sub txtDescricao_GotFocus()
   Marca Me
   txtDescricao.BackColor = QBColor(14)
End Sub

Private Sub txtDescricao_LostFocus()
   txtDescricao.BackColor = QBColor(15)
End Sub

Private Sub txtValor_GotFocus()
   Marca Me
   txtValor.BackColor = QBColor(14)
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 22, 44, 48 To 57
   Case vbKeyBack
   Case Else
      Beep
      KeyAscii = 0
   End Select
End Sub

Private Sub txtValor_LostFocus()
   txtValor.BackColor = QBColor(15)
   If Trim(txtValor) = "" Then txtValor = "0,00"
   txtValor = Abs(txtValor)
   txtValor = Format(txtValor, "###,##0.00")
End Sub

Private Sub txtValor_Validate(Cancel As Boolean)
   If Trim(txtValor) <> "" Then
      If IsNumeric(txtValor) = False Then
         MsgBoxTabum Me, "FAVOR DIGITAR UM VALOR VÁLIDO"
         Cancel = True
         Marca Me
         txtValor.SetFocus
      End If
   End If
End Sub

Private Sub RELATORIO_CONTROLE_FINANCEIRO_BASICO()

   CmdSql = "DELETE FROM REL_CONTROLE_FINANCEIRO_BASICO"
   CMySql.Executa CmdSql, True
      
   CmdSql = "SELECT REGRA,O.CATEG,O.SUB_CATEG,O.TIPO,ROUND(DATA / 100,0)DATA,SUM(VALOR) VALOR" & vbCr
   CmdSql = CmdSql & "FROM CONTROLE_FINANCEIRO O" & vbCr
   CmdSql = CmdSql & "        LEFT JOIN TIPO_MOVIMENTACAO C ON O.TIPO  = C.TIPO" & vbCr
   CmdSql = CmdSql & "                  AND O.CATEG = C.CATEG" & vbCr
   CmdSql = CmdSql & "          AND O.SUB_CATEG = C.SUB_CATEG" & vbCr
   CmdSql = CmdSql & "WHERE O.TIPO = 'DÉBITO'" & vbCr
   CmdSql = CmdSql & "  AND DATA >= " & Format(FirstDay(Date - 365), "YYYYMMDD") & vbCr
   CmdSql = CmdSql & "  AND STATUS <> 'D'" & vbCr
   CmdSql = CmdSql & "  AND REGRA NOT IN ('TERCEIROS')" & vbCr
   CmdSql = CmdSql & "GROUP BY 1,2,3,4,5" & vbCr
   CmdSql = CmdSql & "ORDER BY 5 DESC,1 DESC,CATEG"
   CMySql.Consulta CmdSql, RsValores

   Do While Not RsValores.EOF
      CmdSql = "SELECT REGRA,O.TIPO,ROUND(DATA / 100,0)DATA,SUM(VALOR) VALOR" & vbCr
      CmdSql = CmdSql & "FROM CONTROLE_FINANCEIRO O" & vbCr
      CmdSql = CmdSql & "        LEFT JOIN TIPO_MOVIMENTACAO C ON O.TIPO  = C.TIPO" & vbCr
      CmdSql = CmdSql & "                  AND O.CATEG = C.CATEG" & vbCr
      CmdSql = CmdSql & "          AND O.SUB_CATEG = C.SUB_CATEG" & vbCr
      CmdSql = CmdSql & "WHERE O.TIPO = 'CRÉDITO'" & vbCr
      CmdSql = CmdSql & "  AND DATA BETWEEN " & Format("01/" & Mid(RsValores("DATA"), 5, 2) & "/" & Mid(RsValores("DATA"), 1, 4), "YYYYMMDD") & " AND " & Format(LastDay("01/" & Mid(RsValores("DATA"), 5, 2) & "/" & Mid(RsValores("DATA"), 1, 4)), "YYYYMMDD") & vbCr
      CmdSql = CmdSql & "  AND STATUS <> 'D'" & vbCr
      CmdSql = CmdSql & "  AND REGRA NOT IN ('TERCEIROS')" & vbCr
      CmdSql = CmdSql & "GROUP BY 1,2,3" & vbCr
      CmdSql = CmdSql & "ORDER BY 3 DESC,1 DESC"
      CMySql.Consulta CmdSql, RsCreditos
      
      CREDITO = 0
      GASTOS_ESSENCIAIS = 0
      INVESTIMENTOS_DIVIDAS = 0
      DESEJOS_PESSOAIS = 0
      
      If RsCreditos.EOF Then
         MsgBoxTabum Me, "NÃO HÁ CRÉDITOS NESTE MÊS PARA CÁLCULOS"
         cmdImprimir.Enabled = True
         cmdImprimir.SetFocus
         Exit Sub
      End If
            
      If Not RsCreditos.EOF Then
         GASTOS_ESSENCIAIS = Round(((RsCreditos("VALOR") / 100) * 45), 2)        '45%
         INVESTIMENTOS_DIVIDAS = Round(((RsCreditos("VALOR") / 100) * 40), 2)    '40%
         DESEJOS_PESSOAIS = Round(((RsCreditos("VALOR") / 100) * 15), 2)         '15%
         CREDITO = RsCreditos("VALOR")
      End If
   
      CmdSql = "INSERT INTO REL_CONTROLE_FINANCEIRO_BASICO(REGRA,CATEG,SUBCATEG,MES,VALOR,PERC,CREDITO,ANOMES)" & vbCr
      CmdSql = CmdSql & "VALUES("
      
      Select Case RsValores("REGRA")
      Case "45% GASTOS ESSENCIAIS"
         CmdSql = CmdSql & PoeAspas(RsValores("REGRA")) & ","
      Case "40% INVETIMENTOS E DÍVIDAS"
         CmdSql = CmdSql & PoeAspas(RsValores("REGRA")) & ","
      Case "15% DESEJOS PESSOAIS"
         CmdSql = CmdSql & PoeAspas(RsValores("REGRA")) & ","
      End Select
      
      CmdSql = CmdSql & PoeAspas(RsValores("CATEG")) & ","
      CmdSql = CmdSql & PoeAspas(RsValores("SUB_CATEG")) & ","
      CmdSql = CmdSql & PoeAspas(RsValores("DATA")) & ","
      CmdSql = CmdSql & Str(RsValores("VALOR")) & ","
      
      Select Case RsValores("REGRA")
      Case "45% GASTOS ESSENCIAIS"
         CmdSql = CmdSql & Str(Round((RsValores("VALOR") / GASTOS_ESSENCIAIS) * 100, 2)) & ","
         CmdSql = CmdSql & Str(GASTOS_ESSENCIAIS) & ","
      Case "40% INVETIMENTOS E DÍVIDAS"
         CmdSql = CmdSql & Str(Round((RsValores("VALOR") / INVESTIMENTOS_DIVIDAS) * 100, 2)) & ","
         CmdSql = CmdSql & Str(INVESTIMENTOS_DIVIDAS) & ","
      Case "15% DESEJOS PESSOAIS"
         CmdSql = CmdSql & Str(Round((RsValores("VALOR") / DESEJOS_PESSOAIS) * 100, 2)) & ","
         CmdSql = CmdSql & Str(DESEJOS_PESSOAIS) & ","
      End Select
      
      CmdSql = CmdSql & PoeAspas(Mid(UCase(MonthName(Mid(RsValores("DATA"), 5, 2))), 1, 3) & "/" & Mid(RsValores("DATA"), 1, 4)) & ")"
      
      CMySql.Executa CmdSql, True
         
   RsValores.MoveNext
   Loop
 
End Sub

Private Sub RELATORIO_MENSAL()

   CmdSql = "DELETE FROM REL_CONTROLE_FINANCEIRO_SINTETICO"
   CMySql.Executa CmdSql, True

   CmdSql = "SELECT TIPO,CATEG,SUB_CATEG,DATA,DESCRICAO,VALOR" & vbCr
   CmdSql = CmdSql & "FROM CONTROLE_FINANCEIRO" & vbCr
   CmdSql = CmdSql & "WHERE DATA BETWEEN " & Format(FirstDay(CDate(mskData)), "YYYYMMDD") & " AND " & Format(LastDay(CDate(mskData)), "YYYYMMDD") & vbCr
   CmdSql = CmdSql & "ORDER BY 1,2,3,4"
   CMySql.Consulta CmdSql, Rsr

   Do While Not Rsr.EOF
      CmdSql = "INSERT INTO REL_CONTROLE_FINANCEIRO_SINTETICO(TIPO,CATEG,SUB_CATEG,DATA,DESCRICAO,ANOMES,ANOMES_DESC,VALOR)" & vbCr
      CmdSql = CmdSql & "VALUES("
      CmdSql = CmdSql & PoeAspas(Rsr("TIPO")) & ","
      CmdSql = CmdSql & PoeAspas(Rsr("CATEG")) & ","
      CmdSql = CmdSql & PoeAspas(Rsr("SUB_CATEG")) & ","
      CmdSql = CmdSql & Format(Rsr("DATA"), "YYYYMMDD") & ","
      CmdSql = CmdSql & PoeAspas(Rsr("DESCRICAO")) & ","
      CmdSql = CmdSql & PoeAspas(Mid(Format(Rsr("DATA"), "YYYYMMDD"), 1, 6)) & ","
      CmdSql = CmdSql & PoeAspas(UCase(MonthName(Mid(Mid(Format(Rsr("DATA"), "YYYYMMDD"), 1, 6), 5, 2))) & "/" & Mid(Mid(Format(Rsr("DATA"), "YYYYMMDD"), 1, 6), 1, 4)) & ","
      CmdSql = CmdSql & Str(IIf(UCase(Trim(Rsr("TIPO"))) = "CRÉDITO", Rsr("VALOR"), Rsr("VALOR") * -1)) & ")"
      CMySql.Executa CmdSql, True
   Rsr.MoveNext
   Loop

End Sub
