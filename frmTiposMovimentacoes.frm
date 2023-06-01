VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTiposMovimentacoes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Movimentações"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTiposMovimentacoes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_Botoes 
      BackColor       =   &H00FFFFFF&
      DragMode        =   1  'Automatic
      Height          =   2490
      Left            =   9165
      TabIndex        =   14
      Top             =   -15
      Width           =   2025
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
         Left            =   60
         Picture         =   "frmTiposMovimentacoes.frx":1486
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir Registro"
         Top             =   165
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
         Left            =   1005
         Picture         =   "frmTiposMovimentacoes.frx":1F90
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salvar Registro"
         Top             =   1050
         Width           =   950
      End
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
         Left            =   1005
         Picture         =   "frmTiposMovimentacoes.frx":2A9A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar Campos"
         Top             =   150
         Width           =   950
      End
   End
   Begin VB.Frame fra_Usuario 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dados do Lançamento"
      DragMode        =   1  'Automatic
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
      Height          =   2490
      Left            =   30
      TabIndex        =   9
      Top             =   -15
      Width           =   9090
      Begin VB.TextBox txtCategoria 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2055
         MaxLength       =   50
         TabIndex        =   1
         Top             =   510
         Width           =   5865
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
         Left            =   8295
         Picture         =   "frmTiposMovimentacoes.frx":361C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1935
         Width           =   420
      End
      Begin VB.TextBox txtSubcategoria 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2055
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1200
         Width           =   5865
      End
      Begin VB.ComboBox cmbRegra 
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "frmTiposMovimentacoes.frx":3EE6
         Left            =   2055
         List            =   "frmTiposMovimentacoes.frx":3EE8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1950
         Width           =   5925
      End
      Begin MSComDlg.CommonDialog botaoImportacao 
         Left            =   210
         Top             =   1110
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Subcategoria"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   2055
         TabIndex        =   13
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Categoria"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   2055
         TabIndex        =   12
         Top             =   255
         Width           =   945
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   11
         Top             =   255
         Width           =   420
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Regra"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   2055
         TabIndex        =   10
         Top             =   1680
         Width           =   555
      End
   End
   Begin MSComctlLib.ListView lvwConsulta 
      Height          =   5565
      Left            =   15
      TabIndex        =   8
      Top             =   2475
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9816
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
Attribute VB_Name = "frmTiposMovimentacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Private Item                        As ListItem
   Private IL                          As Long
   Private ARQBKP                      As String
   
   Private ALTER_GRADE                 As Boolean
   
   Private Reg_Envio                   As String * 170
   Private TAB_TIPO                    As String * 20
   Private TAB_CATEG                   As String * 50
   Private TAB_SUB_CATEG               As String * 50
   Private TAB_REGRA                   As String * 50
   
   Private Const GRD_TIPO              As Integer = 1
   Private Const GRD_CATEG             As Integer = 2
   Private Const GRD_SUB_CATEG         As Integer = 3
   Private Const GRD_REGRA             As Integer = 4
   Private Const GRD_GRADE             As Integer = 5

Private Sub MontaLvw()
   lvwConsulta.ListItems.Clear
   lvwConsulta.ColumnHeaders.Clear
   lvwConsulta.Gridlines = True
   lvwConsulta.ColumnHeaders.Add , , "", 0
   lvwConsulta.ColumnHeaders.Add , , "Tipo", 1000, lvwColumnCenter
   lvwConsulta.ColumnHeaders.Add , , "Categoria", 2500
   lvwConsulta.ColumnHeaders.Add , , "Subcategoria", 3900
   lvwConsulta.ColumnHeaders.Add , , "Regra", 3000, lvwColumnCenter
   lvwConsulta.ColumnHeaders.Add , , "", 400, lvwColumnCenter
   lvwConsulta.View = lvwReport
End Sub

Private Sub cmbRegra_LostFocus()
   cmbRegra.BackColor = QBColor(15)
End Sub

Private Sub cmbTipo_GotFocus()
   cmbTipo.BackColor = QBColor(14)
End Sub

Private Sub cmbTipo_LostFocus()
   If cmbTipo <> "" Then
      If cmbTipo = "CRÉDITO" And cmbRegra <> "RECEITAS" Then
         cmbRegra = "RECEITAS"
      Else
         If cmbTipo <> "CRÉDITO" And cmbRegra = "RECEITAS" Then
            cmbRegra = "45% GASTOS ESSENCIAIS"
         End If
      End If
   End If
   
   cmbTipo.BackColor = QBColor(15)
End Sub

Private Sub cmdAdd_Click()

   If cmbTipo = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER O TIPO"
      cmbTipo.SetFocus
      Exit Sub
   End If
      
   If txtCategoria = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER UMA CATEGORIA"
      txtCategoria.SetFocus
      Exit Sub
   End If
   
   If txtSubcategoria = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER UMA SUBCATEGORIA"
      txtSubcategoria.SetFocus
      Exit Sub
   End If
   
   If cmbRegra = "" Then
      MsgBoxTabum Me, "FAVOR PREENCHER A REGRA"
      cmbRegra.SetFocus
      Exit Sub
   End If
   
   If cmbTipo = "CRÉDITO" And cmbRegra <> "RECEITAS" And cmbRegra <> "TERCEIROS" Then
      MsgBoxTabum Me, "PARA CRÉDITO, REGRA DEVE SER RECEITAS OU TERCEIROS"
      cmbRegra.SetFocus
      Exit Sub
   End If
   
   If cmbTipo <> "CRÉDITO" And cmbRegra = "RECEITAS" Then
      MsgBoxTabum Me, "PARA DÉBITO, REGRA NÃO DEVE SER RECEITAS"
      cmbRegra.SetFocus
      Exit Sub
   End If
      
   If ALTER_GRADE = False Then
      With lvwConsulta
         For IL = 1 To .ListItems.Count
            If .ListItems(IL).Checked Then
               If .ListItems(IL).SubItems(GRD_TIPO) = Trim(cmbTipo) And .ListItems(IL).SubItems(GRD_CATEG) = Trim(txtCategoria) _
                  And .ListItems(IL).SubItems(GRD_SUB_CATEG) = Trim(txtSubcategoria) Then
                  MsgBoxTabum Me, "REFERÊNCIA JÁ EXISTE NA GRADE"
                  txtSubcategoria.SetFocus
                  Exit Sub
               End If
            End If
         Next IL
      End With
   End If
   
CONTINUA:
   
   If ALTER_GRADE = False Then
      Set Item = lvwConsulta.ListItems.Add(, , "")
      Item.Bold = True
      Item.ForeColor = RGB(0, 0, 250)
      Item.SubItems(GRD_TIPO) = cmbTipo
      Item.SubItems(GRD_CATEG) = Trim(txtCategoria)
      Item.SubItems(GRD_SUB_CATEG) = Trim(txtSubcategoria)
      Item.SubItems(GRD_REGRA) = Trim(cmbRegra)
      Item.SubItems(GRD_GRADE) = "I"
            
      lvwConsulta.ListItems(lvwConsulta.ListItems.Count).EnsureVisible
      lvwConsulta.ListItems(lvwConsulta.ListItems.Count).Selected = True
      lvwConsulta.SetFocus
   Else
      lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_TIPO) = cmbTipo
      lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_CATEG) = txtCategoria
      lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_SUB_CATEG) = txtSubcategoria
      lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_REGRA) = cmbRegra
      lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_GRADE) = "A"
      ALTER_GRADE = False
      
      cmbTipo.Enabled = True
      txtCategoria.Enabled = True
      txtSubcategoria.Enabled = True
      
      cmbTipo.ListIndex = 0
      txtCategoria = ""
      txtSubcategoria = ""
      cmbRegra.ListIndex = 0
      cmdSalvar.Enabled = True
      lvwConsulta.SetFocus
      Exit Sub
   End If
   
   DoEvents
   Refresh
         
   MsgBoxTabum Me, "REGISTROS INCLUÍDOS COM SUCESSO"
         
   cmdSalvar.Enabled = True
   cmbTipo.SetFocus
End Sub

Private Sub cmdCancelar_Click()
   cmbTipo.Enabled = True
   txtCategoria.Enabled = True
   txtSubcategoria.Enabled = True
   
   ALTER_GRADE = False
   cmbTipo.ListIndex = 0
   txtCategoria = ""
   txtSubcategoria = ""
   cmbRegra.ListIndex = 0
   cmdSalvar.Enabled = False
   Form_Load
   cmbTipo.SetFocus
End Sub

Private Sub cmdConsultar_Click()
   MontaLvw
   
   cmdConsultar.Enabled = False
   cmdSalvar.Enabled = False
         
   CmdSql = "SELECT * FROM TIPO_MOVIMENTACAO" & vbCr
   CmdSql = CmdSql & "WHERE TIPO <> ''" & vbCr
   
   If cmbTipo <> "" Then CmdSql = CmdSql & "  AND TIPO = " & PoeAspas(cmbTipo) & vbCr
   If txtCategoria <> "" Then CmdSql = CmdSql & "  AND CATEG LIKE " & PoeAspas("%" & txtCategoria & "%") & vbCr
 
   If txtSubcategoria <> "" Then CmdSql = CmdSql & "  AND SUB_CATEG LIKE " & PoeAspas("%" & txtSubcategoria & "%") & vbCr
   If cmbRegra <> "" Then CmdSql = CmdSql & "  AND REGRA LIKE " & PoeAspas("%" & cmbRegra & "%") & vbCr
   
   CmdSql = CmdSql & "ORDER BY TIPO,CATEG,SUB_CATEG,REGRA"
   CMySql.Consulta CmdSql, Rs
      
   Do While Not Rs.EOF
      Set Item = lvwConsulta.ListItems.Add(, , "")
            
      Item.Bold = True
      Item.ForeColor = RGB(0, 0, 250)
      Item.SubItems(GRD_TIPO) = Trim(Rs("TIPO"))
      Item.SubItems(GRD_CATEG) = Trim(Rs("CATEG"))
      Item.SubItems(GRD_SUB_CATEG) = Trim(Rs("SUB_CATEG"))
      Item.SubItems(GRD_REGRA) = Trim(Rs("REGRA"))
      Item.SubItems(GRD_GRADE) = ""
   Rs.MoveNext
   Loop
   
   cmdConsultar.Enabled = True
End Sub

Private Sub CmdSalvar_Click()
   cmdSalvar.Enabled = False
   
   With lvwConsulta
      For IL = 1 To .ListItems.Count
         Select Case .ListItems(IL).SubItems(GRD_GRADE)
         Case "I"
ALTERAR:
            CmdSql = "INSERT INTO TIPO_MOVIMENTACAO(TIPO,CATEG,SUB_CATEG,REGRA)" & vbCr
            CmdSql = CmdSql & "VALUES("
            CmdSql = CmdSql & PoeAspas(.ListItems(IL).SubItems(GRD_TIPO)) & ","
            CmdSql = CmdSql & PoeAspas(.ListItems(IL).SubItems(GRD_CATEG)) & ","
            CmdSql = CmdSql & PoeAspas(.ListItems(IL).SubItems(GRD_SUB_CATEG)) & ","
            CmdSql = CmdSql & PoeAspas(.ListItems(IL).SubItems(GRD_REGRA)) & ")"
            CMySql.Executa CmdSql, True
         Case "A"
               CmdSql = "UPDATE TIPO_MOVIMENTACAO SET REGRA =  " & PoeAspas(.ListItems(IL).SubItems(GRD_REGRA)) & vbCr
               CmdSql = CmdSql & "WHERE TIPO      = " & PoeAspas(.ListItems(IL).SubItems(GRD_TIPO)) & vbCr
               CmdSql = CmdSql & "  AND CATEG     = " & PoeAspas(.ListItems(IL).SubItems(GRD_CATEG)) & vbCr
               CmdSql = CmdSql & "  AND SUB_CATEG = " & PoeAspas(.ListItems(IL).SubItems(GRD_SUB_CATEG))
               CMySql.Executa CmdSql, True
         Case "D"
            CmdSql = "SELECT * FROM CONTROLE_FINANCEIRO" & vbCr
            CmdSql = CmdSql & "WHERE TIPO      = " & PoeAspas(.ListItems(IL).SubItems(GRD_TIPO)) & vbCr
            CmdSql = CmdSql & "  AND CATEG     = " & PoeAspas(.ListItems(IL).SubItems(GRD_CATEG)) & vbCr
            CmdSql = CmdSql & "  AND SUB_CATEG = " & PoeAspas(.ListItems(IL).SubItems(GRD_SUB_CATEG))
            CMySql.Consulta CmdSql, Rst

            If Rst.EOF Then
               CmdSql = "DELETE FROM TIPO_MOVIMENTACAO" & vbCr
               CmdSql = CmdSql & "WHERE TIPO      = " & PoeAspas(.ListItems(IL).SubItems(GRD_TIPO)) & vbCr
               CmdSql = CmdSql & "  AND CATEG     = " & PoeAspas(.ListItems(IL).SubItems(GRD_CATEG)) & vbCr
               CmdSql = CmdSql & "  AND SUB_CATEG = " & PoeAspas(.ListItems(IL).SubItems(GRD_SUB_CATEG)) & vbCr
               CmdSql = CmdSql & "  AND REGRA     = " & PoeAspas(.ListItems(IL).SubItems(GRD_REGRA))
               CMySql.Executa CmdSql, True
            Else
               MsgBoxTabum Me, "REGISTRO NÃO PODE SER EXCLUÍDO, EXISTEM DESPESAS GRAVADAS"
            End If
            
            If .ListItems(IL).SubItems(GRD_GRADE) = "A" Then GoSub ALTERAR
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
   
   cmbRegra.Clear
   cmbRegra.AddItem ""
   cmbRegra.AddItem "RECEITAS"
   cmbRegra.AddItem "45% GASTOS ESSENCIAIS"
   cmbRegra.AddItem "40% INVETIMENTOS E DÍVIDAS"
   cmbRegra.AddItem "15% DESEJOS PESSOAIS"
   cmbRegra.AddItem "TERCEIROS"
   cmbRegra.ListIndex = 0
      
   cmdConsultar_Click
End Sub

Private Sub lvwConsulta_DblClick()
   If lvwConsulta.ListItems.Count = 0 Then Exit Sub
 
   Select Case Trim(lvwConsulta.SelectedItem.ListSubItems(GRD_GRADE))
   Case "I"
      cmbTipo = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_TIPO)
      txtCategoria = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_CATEG)
      txtSubcategoria = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_SUB_CATEG)
      cmbRegra = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_REGRA)
      lvwConsulta.ListItems.Remove lvwConsulta.SelectedItem.Index
      cmbTipo.SetFocus
   Case ""
      ALTER_GRADE = True
      cmbTipo = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_TIPO)
      txtCategoria = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_CATEG)
      txtSubcategoria = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_SUB_CATEG)
      If lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_REGRA) <> "" Then cmbRegra = lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_REGRA)
      
      cmbTipo.Enabled = False
      txtCategoria.Enabled = False
      txtSubcategoria.Enabled = False
      
      cmbRegra.SetFocus
   End Select
   
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
      Case "D", "A"
         lvwConsulta.ListItems(lvwConsulta.SelectedItem.Index).SubItems(GRD_GRADE) = ""
      End Select
   End Select
      
   With lvwConsulta
      For IL = 1 To .ListItems.Count
         If lvwConsulta.ListItems(IL).SubItems(GRD_GRADE) = "D" Or lvwConsulta.ListItems(IL).SubItems(GRD_GRADE) = "I" Then
            cmdSalvar.Enabled = True
            Exit Sub
         End If
      Next IL
   End With
End Sub

Private Sub cmbRegra_GotFocus()
   cmbRegra.BackColor = QBColor(14)
End Sub

Private Sub txtCategoria_GotFocus()
   Marca Me
   txtCategoria.BackColor = QBColor(14)
End Sub

Private Sub txtCategoria_LostFocus()
   txtCategoria.BackColor = QBColor(15)
End Sub

Private Sub txtSubcategoria_GotFocus()
   Marca Me
   txtSubcategoria.BackColor = QBColor(14)
End Sub

Private Sub txtSubcategoria_LostFocus()
   txtSubcategoria.BackColor = QBColor(15)
End Sub
