Imports System.Drawing.Drawing2D
Imports System.Data.SqlClient
Imports System.Windows.Forms.RichTextBox
'Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Windows.Forms.DataVisualization.Charting
Public Class frm_opn

    Private Sub SairToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SairToolStripMenuItem.Click
        Try

            da = Nothing : ds = Nothing
            fechar_app("Excel")
            For intIndex As Integer = My.Application.OpenForms.Count - 1 To 0 Step -1
                If My.Application.OpenForms.Item(intIndex) IsNot Me Then
                    My.Application.OpenForms.Item(intIndex).Close()
                End If
            Next
            fechar_app("OPN")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub frm_opn_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        Try

            da = Nothing : ds = Nothing
            fechar_app("Excel")
            For intIndex As Integer = My.Application.OpenForms.Count - 1 To 0 Step -1
                If My.Application.OpenForms.Item(intIndex) IsNot Me Then
                    My.Application.OpenForms.Item(intIndex).Close()
                End If
            Next
            fechar_app("OPN")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub frm_opn_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

            ''ACESSO AO BANCO TOTVS
            'Dim linhaTexto2 As String = "Data Source=GCTOTVS01;Initial Catalog=tmPRD2;User ID=totvs;Password=totvs;Min Pool Size=5;Max Pool Size=15;Connection Reset=True;Connection Lifetime=600;Trusted_Connection=no;MultipleActiveResultSets=True"
            'TratadorDeConexao.Preparar2(linhaTexto2)
            'Dim checkconexao2 = TratadorDeConexao.Abrir2()
            'If checkconexao2 = False Then
            '    MessageBox.Show("Problema na conexão com o banco de dados", "Projetos", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    Exit Sub
            'End If
            ''ACESSO AO BANCO OPN
            'Dim linhaTexto As String = "Data Source=GCSERVER;Initial Catalog=OPN;User ID=sa;Password=050382;Min Pool Size=5;Max Pool Size=15;Connection Reset=True;Connection Lifetime=600;Trusted_Connection=no;MultipleActiveResultSets=True"
            'TratadorDeConexao.Preparar(linhaTexto)
            'Dim checkconexao = TratadorDeConexao.Abrir()
            'If checkconexao = False Then
            '    MessageBox.Show("Problema na conexão com o banco de dados", "Projetos", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    Exit Sub
            'End If
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            '''''''''''''''''''''''''''''''''''''''''''''
            Dim SQDR_opn_qtd As SqlDataReader
            Dim SQCMDqtd As New SqlCommand("Select Count(opn_codigo) as valor From OPN Where opn_status =2", Con)
            SQDR_opn_qtd = SQCMDqtd.ExecuteReader(CommandBehavior.Default)
            If SQDR_opn_qtd.Read Then
                opnTotal_valor.Text = SQDR_opn_qtd("valor")
            Else
                opnTotal_valor.Text = 0
            End If
            SQDR_opn_qtd.Close()

            Dim SQCMDqtd2 As New SqlCommand("Select Count(opn_codigo) as valor From OPN Where opn_status =3", Con)
            SQDR_opn_qtd = SQCMDqtd2.ExecuteReader(CommandBehavior.Default)
            If SQDR_opn_qtd.Read Then
                opnTotal_aceita_valor.Text = SQDR_opn_qtd("valor")
            Else
                opnTotal_aceita_valor.Text = 0
            End If
            SQDR_opn_qtd.Close()

            Dim _percet

            _percet = Val(opnTotal_aceita_valor.Text) / Val(opnTotal_valor.Text)
            aceitacao_percentual_valor.Text = _percet
            ''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQDR_opn_soma As SqlDataReader
            Dim SQCMDsoma As New SqlCommand("Select Sum(opn_valor) as valor From OPN Where opn_status =2", Con)
            SQDR_opn_soma = SQCMDsoma.ExecuteReader(CommandBehavior.Default)
            If SQDR_opn_soma.Read Then
                If Not IsDBNull(SQDR_opn_soma("valor")) Then
                    opn_valor_Total_numero.Text = FormatCurrency(SQDR_opn_soma("valor"))
                Else
                    opn_valor_Total_numero.Text = 0
                End If
            Else
                opn_valor_Total_numero.Text = 0
            End If
            SQDR_opn_soma.Close()

            Dim SQCMDsoma2 As New SqlCommand("Select Sum(opn_valor) as valor From OPN Where opn_status =3", Con)
            SQDR_opn_soma = SQCMDsoma2.ExecuteReader(CommandBehavior.Default)
            If SQDR_opn_soma.Read Then
                If Not IsDBNull(SQDR_opn_soma("valor")) Then
                    opn_valor_Total_Aceita_numero.Text = FormatCurrency(SQDR_opn_soma("valor"))
                Else
                    opn_valor_Total_Aceita_numero.Text = 0
                End If
            Else
                opn_valor_Total_Aceita_numero.Text = 0
            End If
            SQDR_opn_soma.Close()

            ''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQDR_status_opn As SqlDataReader
            Dim SQCMDCor As New SqlCommand("Select * From Status Order By codigo", Con)
            SQDR_status_opn = SQCMDCor.ExecuteReader(CommandBehavior.Default)
            Do While SQDR_status_opn.Read
                If SQDR_status_opn("Codigo") = 1 Then
                    em_aberto = Trim(SQDR_status_opn("cor"))
                ElseIf SQDR_status_opn("Codigo") = 2 Then
                    enviada = Trim(SQDR_status_opn("cor"))
                ElseIf SQDR_status_opn("Codigo") = 3 Then
                    aceita = Trim(SQDR_status_opn("cor"))
                ElseIf SQDR_status_opn("Codigo") = 4 Then
                    declinada = Trim(SQDR_status_opn("cor"))
                ElseIf SQDR_status_opn("Codigo") = 5 Then
                    cancelada_Revogada = Trim(SQDR_status_opn("cor"))
                ElseIf SQDR_status_opn("Codigo") = 6 Then
                    Suspensa = Trim(SQDR_status_opn("cor"))
                End If
            Loop
            SQDR_status_opn.Close()

            Dim SQDR_status_as As SqlDataReader
            Dim SQCMDCor_as As New SqlCommand("Select * From Status_AS Order By codigo", Con)
            SQDR_status_as = SQCMDCor_as.ExecuteReader(CommandBehavior.Default)
            Do While SQDR_status_as.Read
                If SQDR_status_as("Codigo") = 1 Then
                    aexecutar = SQDR_status_as("cor")
                ElseIf SQDR_status_as("Codigo") = 2 Then
                    execucao_pendente = SQDR_status_as("cor")
                ElseIf SQDR_status_as("Codigo") = 3 Then
                    em_execucao = SQDR_status_as("cor")
                ElseIf SQDR_status_as("Codigo") = 4 Then
                    executada = SQDR_status_as("cor")
                ElseIf SQDR_status_as("Codigo") = 5 Then
                    cancelada = SQDR_status_as("cor")
                End If
            Loop
            SQDR_status_as.Close()

            opn_consulta = False
            as_consulta = False

            cmb_codigo_opn.Items.Clear()
            benc_Combo(cmb_codigo_opn, "OPN", "Select * From OPN Order By opn_codigo", "opn_codigo")
            cmb_cliente_opn.Items.Clear()
            benc_Combo(cmb_cliente_opn, "OPN", "Select Distinct opn_cliente From OPN Where opn_cliente is not  null Order By opn_cliente", "opn_cliente")
            ToolStripStatus_cb.Items.Clear()
            benc_Combo(ToolStripStatus_cb, "Lista_OPN", "Select Distinct [Status] From Lista_OPN Where [Status] is not null Order By [Status]", "Status")
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            cmb_codigo_onp_obs.Items.Clear()
            benc_Combo(cmb_codigo_onp_obs, "OPN", "Select * From OPN Order By opn_codigo", "opn_codigo")


            cmb_codigo_as_obs.Items.Clear()
            benc_Combo(cmb_codigo_as_obs, "AS", "Select * From [AS] Order By as_codigo", "as_codigo")

            ToolStripTextBox5.BackColor = ColorTranslator.FromHtml(em_aberto)
            ToolStripTextBox6.BackColor = ColorTranslator.FromHtml(enviada)
            ToolStripTextBox7.BackColor = ColorTranslator.FromHtml(aceita)
            ToolStripTextBox8.BackColor = ColorTranslator.FromHtml(declinada)
            ToolStripTextBox10.BackColor = ColorTranslator.FromHtml(cancelada_Revogada)
            ToolStripTextBox11.BackColor = ColorTranslator.FromHtml(Suspensa)

            contar_Status(ToolStripTextBox5, "Em Aberto")
            contar_Status(ToolStripTextBox6, "Proposta Enviada")
            contar_Status(ToolStripTextBox7, "Proposta Aceita")
            contar_Status(ToolStripTextBox8, "Declinada ")
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            cmb_codigo.Items.Clear()
            benc_Combo(cmb_codigo, "AS", "Select * From [AS] Order By as_codigo", "as_codigo")

            cmb_clientes.Items.Clear()
            Dim SQDR_cliente_opn As SqlDataReader
            Dim SQCMDPorta As New SqlCommand("Select Distinct as_cliente_totvs From Lista_AS Order By as_cliente_totvs", Con)
            SQDR_cliente_opn = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
            Do While SQDR_cliente_opn.Read
                benc_Combo2(cmb_clientes, "Exibicao_Cliente", "Select A1_COD,A1_NOME From Exibicao_Cliente Where A1_COD= '" & SQDR_cliente_opn("as_cliente_totvs") & "' Order By A1_NOME", "A1_NOME")
            Loop
            SQDR_cliente_opn.Close()

            'TabPage das AS
            benc_Combo2(cmb_clientes, "Exibicao_Cliente", "Select Distinct A1_NOME From Exibicao_Cliente Order By A1_NOME", "A1_NOME")
            ToolStripTextBox1.BackColor = ColorTranslator.FromHtml(aexecutar)
            ToolStripTextBox2.BackColor = ColorTranslator.FromHtml(execucao_pendente)
            ToolStripTextBox3.BackColor = ColorTranslator.FromHtml(em_execucao)
            ToolStripTextBox4.BackColor = ColorTranslator.FromHtml(executada)
            ToolStripTextBox9.BackColor = ColorTranslator.FromHtml(cancelada)

            contar_Status_as(ToolStripTextBox1, "À Executar")
            contar_Status_as(ToolStripTextBox2, "Em Execução/Pendente")
            contar_Status_as(ToolStripTextBox3, "Em Execução")
            contar_Status_as(ToolStripTextBox4, "Executada ")
            contar_Status_as(ToolStripTextBox9, "Cancelada ")
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            'ToolStripMenu_AS_01.Checked = False
            'ToolStripMenu_AS_01.Checked = True
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            DG_01.EnableHeadersVisualStyles = False
            DG_01.ColumnHeadersDefaultCellStyle.BackColor = Color.Gainsboro   'LightSteelBlue
            DG_01.RowHeadersDefaultCellStyle.BackColor = Color.Gainsboro

            ativar_filtros_status()
            ativar_filtros_status_AS()
            ' ExibeDados(DG_01, "Select * From Lista_OPN Order By Prioridade", "Lista_OPN")
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            'DG_02.EnableHeadersVisualStyles = False
            'DG_02.ColumnHeadersDefaultCellStyle.BackColor = Color.Gainsboro   'LightSteelBlue
            'DG_02.RowHeadersDefaultCellStyle.BackColor = Color.Gainsboro

            'ExibeDadosAS(DG_02, "Select * From Lista_AS Order By [Código]", "Lista_AS")
            ''''''''''''''''''''''''''''''''''''''''''''''''''''

            If usuario_perfil = False Then
                CadastroDeASsToolStripMenuItem.Enabled = False
            Else
                CadastroDeASsToolStripMenuItem.Enabled = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Public Sub ativar_filtros_status()
        '3334-2964
        If ToolStripMenuItem3.Checked Then
            status_01 = 0
            status_01_as = 0
        Else
            status_01 = 9
            status_01_as = 9
        End If
        If ToolStripMenuItem4.Checked Then
            status_02 = 1
            status_02_as = 1
        Else
            status_02 = 9
            status_02_as = 9
        End If
        If ToolStripMenuItem5.Checked Then
            status_03 = 2
            status_03_as = 2
        Else
            status_03 = 9
            status_03_as = 9
        End If
        If ToolStripMenuItem6.Checked Then
            status_04 = 3
            status_04_as = 3
        Else
            status_04 = 9
            status_04_as = 9
        End If
        If ToolStripMenuItem7.Checked Then
            status_05 = 4
            status_05_as = 4
        Else
            status_05 = 9
            status_05_as = 9
        End If
        If SuspensaToolStripMenuItem.Checked Then
            status_06 = 5
            status_06_as = 5
        Else
            status_06 = 9
            status_06_as = 9
        End If
        If EmAbertoClienteToolStripMenuItem.Checked Then
            status_10 = 10
            status_10_as = 10
        Else
            status_10 = 9
            status_10_as = 9
        End If

        ExibeDados(DG_01, "Select * From Lista_OPN Where Prioridade = " & status_01 & " or Prioridade = " & status_02 & " or Prioridade = " & status_03 & " or Prioridade = " & status_04 & " Order By Status", "Lista_OPN")
        'ExibeDados(DG_01, "Select * From Lista_OPN Where Prioridade = " & status_01 & " or Prioridade = " & status_02 & " or Prioridade = " & status_03 & " or Prioridade = " & status_04 & " or Prioridade = " & status_05 & " or Prioridade = " & status_06 & " Order By Prioridade", "Lista_OPN")
        'ExibeDados(DG_01, "Select * From Lista_OPN Where Codigo = " & status_01 & " or Codigo = " & status_02 & " or Codigo = " & status_03 & " or Codigo = " & status_04 & " or Codigo = " & status_05 & " or Codigo = " & status_06 & " or Codigo = " & status_10 & " Order By Prioridade", "Lista_OPN")
        'DG_01.Columns(11).DefaultCellStyle.Format = "c"
    End Sub
    Public Sub ativar_filtros()
        Try
            If ToolStripMenuItem3.Checked Then
                status_01 = 0
                status_01_as = 0
            Else
                status_01 = 9
                status_01_as = 9
            End If
            If ToolStripMenuItem4.Checked Then
                status_02 = 1
                status_02_as = 1
            Else
                status_02 = 9
                status_02_as = 9
            End If
            If ToolStripMenuItem5.Checked Then
                status_03 = 2
                status_03_as = 2
            Else
                status_03 = 9
                status_03_as = 9
            End If
            If ToolStripMenuItem6.Checked Then
                status_04 = 3
                status_04_as = 3
            Else
                status_04 = 9
                status_04_as = 9
            End If
            If ToolStripMenuItem7.Checked Then
                status_05 = 4
                status_05_as = 4
            Else
                status_05 = 9
                status_05_as = 9
            End If
            If SuspensaToolStripMenuItem.Checked Then
                status_06 = 5
                status_06_as = 5
            Else
                status_06 = 9
                status_06_as = 9
            End If
            If EmAbertoClienteToolStripMenuItem.Checked Then
                status_10 = 10
                status_10_as = 10
            Else
                status_10 = 9
                status_10_as = 9
            End If

            ExibeDados(DG_01, "Select * From Lista_OPN Where Prioridade = " & status_01 & " or Prioridade = " & status_02 & " or Prioridade = " & status_03 & " or Prioridade = " & status_04 & " Order By Status", "Lista_OPN")

        Catch ex As Exception

        End Try
    End Sub
    Public Sub ativar_filtros_status_AS()

        If ToolStripMenu_AS_01.Checked Then
            status_01_as = 0
        Else
            status_01_as = 9
        End If
        If ToolStripMenu_AS_02.Checked Then
            status_02_as = 1
        Else
            status_02_as = 9
        End If
        If ToolStripMenu_AS_03.Checked Then
            status_03_as = 2
        Else
            status_03_as = 9
        End If
        If ToolStripMenu_AS_04.Checked Then
            status_04_as = 3
        Else
            status_04_as = 9
        End If
        If ToolStripMenu_AS_05.Checked Then
            status_05_as = 4
        Else
            status_05_as = 9
        End If
        'ExibeDados(DG_01, "Select * From Lista_OPN Where Prioridade = " & status_01 & " or Prioridade = " & status_02 & " or Prioridade = " & status_03 & " or Prioridade = " & status_04 & " Order By Status", "Lista_OPN")
        ExibeDados(DG_02, "Select * From Lista_AS Where Prioridade = " & status_01_as & " or Prioridade = " & status_02_as & " or Prioridade = " & status_03_as & " or Prioridade = " & status_04_as & " or Prioridade = " & status_05_as & " Order By Prioridade", "Lista_AS")
        'ExibeDadosAS(DG_02, "Select * From Lista_AS Where Prioridade = " & status_01_as & " or Prioridade = " & status_02_as & " or Prioridade = " & status_03_as & " or Prioridade = " & status_04_as & " or Prioridade = " & status_05_as & " Order By Status", "Lista_AS")
    End Sub


    'Private Sub ToolStripMenu_Proposta_Enviada_Click(sender As Object, e As EventArgs)
    '    Try
    '        If ToolStripMenu_AS_02.Checked = True Then
    '            ToolStripMenu_AS_02.Checked = False
    '            ativar_filtros_status()
    '        Else
    '            cmb_codigo.Text = "" : cmb_clientes.Text = ""
    '            ToolStripMenu_AS_02.Checked = True
    '            ativar_filtros_status()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
    '    End Try
    'End Sub

    'Private Sub ToolStripMenu_Proposta_Aceita_Click(sender As Object, e As EventArgs)
    '    Try
    '        If ToolStripMenu_AS_03.Checked = True Then
    '            ToolStripMenu_AS_03.Checked = False
    '            ativar_filtros_status()
    '        Else
    '            cmb_codigo.Text = "" : cmb_clientes.Text = ""
    '            ToolStripMenu_AS_03.Checked = True
    '            ativar_filtros_status()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
    '    End Try
    'End Sub

    'Private Sub ToolStripMenu_Declinada_Click(sender As Object, e As EventArgs)
    '    Try
    '        If ToolStripMenu_AS_04.Checked = True Then
    '            ToolStripMenu_AS_04.Checked = False
    '            ativar_filtros_status()
    '        Else
    '            cmb_codigo.Text = "" : cmb_clientes.Text = ""
    '            ToolStripMenu_AS_04.Checked = True
    '            ativar_filtros_status()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
    '    End Try
    'End Sub

    Private Sub CadastroDeOPNsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CadastroDeOPNsToolStripMenuItem.Click
        Try
            frm_opn_cadastro.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub DG_01_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG_01.CellDoubleClick
        Try
            If e.RowIndex < 0 Then Exit Sub
            If usuario_perfil = False Then Exit Sub
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            cod_edit_OPN = DG_01.Item(0, e.RowIndex).Value
            cmb_codigo_onp_obs.Text = DG_01.Item(0, e.RowIndex).Value
            frm_opn_editar.Show()
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'If Trim(DG_01.Rows(e.RowIndex).Cells(10).Value) = "Proposta Aceita" Then
            '    'frm_as_cadastro.Show()   'frm_as_cadastro
            'frm_opn_editar.Show()
            'Else
            '    'frm_opn_editar.Show()
            'End If

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim dt_01, dt_02 'As Date
            For i = 0 To DG_01.Rows.Count - 1

                Dim SQDR_obs_opn As SqlDataReader
                Dim SQCMDPorta As New SqlCommand("Select Distinct opn_codigo From obs_opn Where opn_codigo = " & Trim(DG_01.Rows(i).Cells(0).Value) & "", Con)
                SQDR_obs_opn = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
                If SQDR_obs_opn.Read Then

                    DG_01.Item(0, i).Style.Font = New Font(DG_01.Font, FontStyle.Bold)
                    DG_01.Item(0, i).Style.BackColor = Color.White

                End If
                SQDR_obs_opn.Close()

                If Trim(DG_01.Rows(i).Cells(10).Value) = "Em Aberto" Then  'And Not IsDBNull(Trim(DG.Rows(i).Cells(3).Value)) = False 

                    If DG_01.Columns(2).Name.Equals("D_limite") Then
                        Try
                            If Not IsDBNull(Trim(DG_01.Rows(i).Cells(1).Value)) Then
                                Try
                                    If IsDate(Trim(DG_01.Rows(i).Cells(1).Value)) Then

                                        dt_01 = CDate(Trim(DG_01.Rows(i).Cells(1).Value))
                                        dt_02 = CDate(DG_01.Rows(i).Cells(2).Value)
                                        Dim result As Long = DateDiff(DateInterval.Day, Date.Now, dt_02)
                                        If result <= 1.8 Then
                                            If DG_01.Item(2, i).Style.ForeColor <> Color.Red Then
                                                DG_01.Item(2, i).Style.ForeColor = Color.Red
                                                DG_01.Item(2, i).Style.BackColor = Color.Yellow
                                            ElseIf DG_01.Item(2, i).Style.ForeColor = Color.Red Then
                                                DG_01.Item(2, i).Style.ForeColor = Color.White
                                                DG_01.Item(2, i).Style.BackColor = Color.Red
                                            End If
                                        Else

                                        End If

                                    End If
                                Catch ex As Exception

                                End Try

                            End If
                        Catch ex As Exception

                        End Try

                    End If
                End If
            Next
        Catch ex As Exception
            ' MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub TabPage1_GotFocus(sender As Object, e As EventArgs) Handles TabPage1.GotFocus
        Try
            MsgBox(e.ToString)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TabPage1_MouseClick(sender As Object, e As MouseEventArgs) Handles TabPage1.MouseClick
        Try
            MsgBox(e.ToString)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TabControl1_MouseClick(sender As Object, e As MouseEventArgs) Handles TabControl1.MouseClick
        Try

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TabControl1_Selected(sender As Object, e As TabControlEventArgs) Handles TabControl1.Selected
        Try
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_status As SqlDataReader
            Dim SQDR_status_as As SqlDataReader
            '


            'ToolStripMenu_AS_05.Visible = False
            If TabControl1.SelectedIndex.ToString = 2 Then
                'ToolStripMenu_AS_01.Checked = False
                ''ativar_filtros_status_AS()
                ''ToolStripMenu_AS_01.Checked = True
                ''ativar_filtros_status_AS()


                'If ToolStripMenu_AS_01.Checked = True Then
                '    ToolStripMenu_AS_01.Checked = False
                '    ativar_filtros_status_AS()
                'Else
                '    cmb_codigo.Text = "" : cmb_clientes.Text = ""
                '    ToolStripMenu_AS_01.Checked = True
                '    ativar_filtros_status_AS()
                'End If
                ''ativar_filtros_status()


                ''Dim SQCMDPorta As New SqlCommand("Select * From Status Order By Codigo", Con)
                ''SQDR_status = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
                ''Do While SQDR_status.Read = True
                ''    ToolStripDropDownButton1.DropDownItems.Add(Trim(SQDR_status("Descricao")))
                ''Loop
                ''SQDR_status.Close()




                ''ToolStripMenu_AS_01.Text = "Em Aberto"
                ''ToolStripMenu_AS_02.Text = "Proposta Enviada"
                ''ToolStripMenu_AS_03.Text = "Proposta Aceita"
                ''ToolStripMenu_AS_04.Text = "Declinada"
            Else
                'ToolStripDropDownButton1.DropDownItems.Clear()
                'Dim SQCMDPorta As New SqlCommand("Select * From Status_AS Order By Codigo", Con)
                'SQDR_status_as = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
                'Do While SQDR_status_as.Read = True
                '    ToolStripDropDownButton1.DropDownItems.Add(Trim(SQDR_status_as("Descricao")))
                'Loop
                'SQDR_status_as.Close()
                'ToolStripMenu_AS_01.Text = "À Executar"
                'ToolStripMenu_AS_01.Text = "Em Execução/Pendente"
                'ToolStripMenu_AS_01.Text = "Em Execução"
                'ToolStripMenu_AS_01.Text = "Executada"
                'ToolStripMenu_AS_05.Text = "Cancelada"
                'ToolStripMenu_AS_05.Visible = True
                'ativar_filtros_status_AS()
            End If

            ' MsgBox(e.TabPage.Name)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripMenu_cancelada_Click(sender As Object, e As EventArgs)
        Try
            If ToolStripMenu_AS_05.Checked = True Then
                ToolStripMenu_AS_05.Checked = False
                ativar_filtros_status()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                ToolStripMenu_AS_05.Checked = True
                ativar_filtros_status()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Try
            If ToolStripMenuItem3.Checked = True Then
                ToolStripMenuItem3.Checked = False
                ativar_filtros_status()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                ToolStripMenuItem3.Checked = True
                ativar_filtros_status()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Try
            If ToolStripMenuItem4.Checked = True Then
                ToolStripMenuItem4.Checked = False
                ativar_filtros_status()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                ToolStripMenuItem4.Checked = True
                ativar_filtros_status()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        Try
            If ToolStripMenuItem5.Checked = True Then
                ToolStripMenuItem5.Checked = False
                ativar_filtros_status()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                ToolStripMenuItem5.Checked = True
                ativar_filtros_status()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        Try
            If ToolStripMenuItem6.Checked = True Then
                ToolStripMenuItem6.Checked = False
                ativar_filtros_status()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                ToolStripMenuItem6.Checked = True
                ativar_filtros_status()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripMenu_AS_01_Click(sender As Object, e As EventArgs) Handles ToolStripMenu_AS_01.Click
        Try
            If ToolStripMenu_AS_01.Checked = True Then
                ToolStripMenu_AS_01.Checked = False
                ativar_filtros_status_AS()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                ToolStripMenu_AS_01.Checked = True
                ativar_filtros_status_AS()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripMenu_AS_02_Click(sender As Object, e As EventArgs) Handles ToolStripMenu_AS_02.Click
        Try
            If ToolStripMenu_AS_02.Checked = True Then
                ToolStripMenu_AS_02.Checked = False
                ativar_filtros_status_AS()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                ToolStripMenu_AS_02.Checked = True
                ativar_filtros_status_AS()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripMenu_AS_03_Click(sender As Object, e As EventArgs) Handles ToolStripMenu_AS_03.Click
        Try
            If ToolStripMenu_AS_03.Checked = True Then
                ToolStripMenu_AS_03.Checked = False
                ativar_filtros_status_AS()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                ToolStripMenu_AS_03.Checked = True
                ativar_filtros_status_AS()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripMenu_AS_04_Click(sender As Object, e As EventArgs) Handles ToolStripMenu_AS_04.Click
        Try
            If ToolStripMenu_AS_04.Checked = True Then
                ToolStripMenu_AS_04.Checked = False
                ativar_filtros_status_AS()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                ToolStripMenu_AS_04.Checked = True
                ativar_filtros_status_AS()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripMenu_AS_05_Click(sender As Object, e As EventArgs) Handles ToolStripMenu_AS_05.Click
        Try
            If ToolStripMenu_AS_05.Checked = True Then
                ToolStripMenu_AS_05.Checked = False
                ativar_filtros_status_AS()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                ToolStripMenu_AS_05.Checked = True
                ativar_filtros_status_AS()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub CadastroDeASsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CadastroDeASsToolStripMenuItem.Click
        Try
            frm_as_cadastro.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub cmb_codigo_opn_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_codigo_opn.SelectedIndexChanged
        Try
            ToolStripMenuItem3.Checked = False
            ToolStripMenuItem4.Checked = False
            ToolStripMenuItem5.Checked = False
            ToolStripMenuItem6.Checked = False
            cmb_cliente_opn.Text = ""
            ExibeDados(DG_01, "Select * From Lista_OPN Where OPN = '" & Trim(cmb_codigo_opn.Text) & "' Order By OPN", "Lista_OPN")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub cmb_cliente_opn_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_cliente_opn.SelectedIndexChanged
        Try
            ToolStripMenuItem3.Checked = False
            ToolStripMenuItem4.Checked = False
            ToolStripMenuItem5.Checked = False
            ToolStripMenuItem6.Checked = False
            cmb_codigo_opn.Text = ""
            ExibeDados(DG_01, "Select * From Lista_OPN Where Cliente Like '%" & Trim(cmb_cliente_opn.Text) & "%' Order By Prioridade", "Lista_OPN")
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub cmb_codigo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_codigo.SelectedIndexChanged
        Try
            ToolStripMenu_AS_01.Checked = False
            ToolStripMenu_AS_02.Checked = False
            ToolStripMenu_AS_03.Checked = False
            ToolStripMenu_AS_04.Checked = False
            ToolStripMenu_AS_05.Checked = False
            'cmb_cliente_opn.Text = ""
            ExibeDados(DG_02, "Select * From Lista_AS Where [Código] = '" & Trim(cmb_codigo.Text) & "' Order By [Código]", "Lista_AS")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub cmb_clientes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_clientes.SelectedIndexChanged
        Try

            Dim Con2 As SqlConnection = TratadorDeConexao.Conexao2()
            Dim SQDR_c_totvs As SqlDataReader
            Dim str_tel As String
            Dim titulo As String = "Lauto Tecnico"
            Dim cod_c_totvs As Integer
            Dim SQCMDcc As New SqlCommand("Select * From [Exibicao_Cliente] Where  A1_NOME = '" & Trim(cmb_clientes.Text) & "'", Con2)
            SQDR_c_totvs = SQCMDcc.ExecuteReader(CommandBehavior.Default)
            If SQDR_c_totvs.Read Then
                cod_c_totvs = SQDR_c_totvs("A1_COD")
            Else
                cod_c_totvs = 0
            End If
            SQDR_c_totvs.Close()

            ToolStripMenu_AS_01.Checked = False
            ToolStripMenu_AS_02.Checked = False
            ToolStripMenu_AS_03.Checked = False
            ToolStripMenu_AS_04.Checked = False
            ToolStripMenu_AS_05.Checked = False
            'cmb_cliente_opn.Text = ""
            ExibeDados(DG_02, "Select * From Lista_AS Where [as_cliente_totvs] = " & cod_c_totvs & " Order By Prioridade", "Lista_AS")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        RichTextBox1.SelectionColor = Color.Black
        RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Italic)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Underline)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        RichTextBox1.SelectionFont = New Font(RichTextBox1.Font.Name, 15, FontStyle.Bold)
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Try

            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_arm As SqlDataReader
            Dim SQDR_max As SqlDataReader
            Dim titulo As String = "Lauto Tecnico"


            Dim SQCMDPorta As New SqlCommand("Select * From obs_opn Where opn_codigo = " & Val(cmb_codigo_onp_obs.Text) & "", Con)
            SQDR_arm = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
            If SQDR_arm.Read Then

                Dim CmdIns_posto As New SqlCommand("Update obs_opn Set observacao='" & Trim(RichTextBox1.Text) & "' Where opn_codigo= " & Trim(cmb_codigo_onp_obs.Text) & " ", Con)
                CmdIns_posto.ExecuteNonQuery() : CmdIns_posto.Dispose()

            Else

                Dim CmdIns_swit As New SqlCommand("Insert into obs_opn(opn_codigo,observacao) values (" & cmb_codigo_onp_obs.Text & ",'" & Trim(RichTextBox1.Text) & "')", Con)
                CmdIns_swit.ExecuteNonQuery() : CmdIns_swit.Dispose()

            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub cmb_codigo_onp_obs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_codigo_onp_obs.SelectedIndexChanged
        Try

            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_obs As SqlDataReader

            Dim SQCMDobs As New SqlCommand("Select * From obs_opn Where opn_codigo = " & Val(cmb_codigo_onp_obs.Text) & "", Con)
            SQDR_obs = SQCMDobs.ExecuteReader(CommandBehavior.Default)
            If SQDR_obs.Read Then

                RichTextBox1.Text = SQDR_obs("observacao")
                'Dim CmdIns_posto As New SqlCommand("Update obs_opn Set observacao='" & Trim(RichTextBox1.Text) & "' Where opn_codigo= " & Trim(cmb_codigo_onp_obs.Text) & " ", Con)
                'CmdIns_posto.ExecuteNonQuery() : CmdIns_posto.Dispose()

            Else
                RichTextBox1.Text = ""
                'Dim CmdIns_swit As New SqlCommand("Insert into obs_opn(opn_codigo,observacao) values (" & cmb_codigo_onp_obs.Text & ",'" & Trim(RichTextBox1.Text) & "')", Con)
                'CmdIns_swit.ExecuteNonQuery() : CmdIns_swit.Dispose()

            End If
            SQDR_obs.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Try
            Dim index As Integer = 0
            Dim lastIndex As Integer = RichTextBox1.Find(txt_find_obs_opn.Text, RichTextBoxFinds.Reverse)
            While index < lastIndex
                index = RichTextBox1.Find(txt_find_obs_opn.Text, index, RichTextBoxFinds.None)
                RichTextBox1.SelectionBackColor = Color.Yellow
                RichTextBox1.SelectionColor = Color.Red
                index += 1
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Try
            If usuario_perfil = False Then Exit Sub

            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim resp = MsgBox("Esta ação excluira as Observações desta OPN. Confirma esta ação?", MsgBoxStyle.YesNo, titulo_opn)
            If resp = 6 Then

                Dim CmdDel_obs_opn As New SqlCommand("Delete From obs_opn Where opn_codigo= " & Trim(cmb_codigo_onp_obs.Text) & " ", Con)
                CmdDel_obs_opn.ExecuteNonQuery() : CmdDel_obs_opn.Dispose()

                MsgBox("Ação comcluida com êxito!", MsgBoxStyle.Information, titulo_opn)
            Else
                MsgBox("Ação cancelada!", MsgBoxStyle.Information, titulo_opn)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub bnt_salba_obs_as_Click(sender As Object, e As EventArgs) Handles bnt_salba_obs_as.Click
        Try

            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_arm As SqlDataReader
            Dim SQDR_max As SqlDataReader

            Dim SQCMDPorta As New SqlCommand("Select * From obs_as Where as_codigo = " & Val(cmb_codigo_as_obs.Text) & "", Con)
            SQDR_arm = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
            If SQDR_arm.Read Then

                Dim CmdIns_posto As New SqlCommand("Update obs_as Set observacao='" & Trim(RichTextBox1.Text) & "' Where as_codigo= " & Trim(cmb_codigo_as_obs.Text) & " ", Con)
                CmdIns_posto.ExecuteNonQuery() : CmdIns_posto.Dispose()

            Else

                Dim CmdIns_swit As New SqlCommand("Insert into obs_as(as_codigo,observacao) values (" & cmb_codigo_as_obs.Text & ",'" & Trim(RichTextBox2.Text) & "')", Con)
                CmdIns_swit.ExecuteNonQuery() : CmdIns_swit.Dispose()

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub cmb_codigo_as_obs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_codigo_as_obs.SelectedIndexChanged
        Try

            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_obs As SqlDataReader

            Dim SQCMDobs As New SqlCommand("Select * From obs_as Where as_codigo = " & Val(cmb_codigo_as_obs.Text) & "", Con)
            SQDR_obs = SQCMDobs.ExecuteReader(CommandBehavior.Default)
            If SQDR_obs.Read Then

                RichTextBox2.Text = SQDR_obs("observacao")

            Else
                RichTextBox2.Text = ""

            End If
            SQDR_obs.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        Try
            Dim index As Integer = 0
            Dim lastIndex As Integer = RichTextBox2.Find(txt_find_obs_as.Text, RichTextBoxFinds.Reverse)
            While index < lastIndex
                index = RichTextBox2.Find(txt_find_obs_opn.Text, index, RichTextBoxFinds.None)
                RichTextBox2.SelectionBackColor = Color.Yellow
                RichTextBox2.SelectionColor = Color.Red
                index += 1
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        Try
            If usuario_perfil = False Then Exit Sub

            Dim Con As SqlConnection = TratadorDeConexao.Conexao()

            Dim CmdDel_obs_as As New SqlCommand("Delete From obs_as Where as_codigo= " & Trim(cmb_codigo_as_obs.Text) & " ", Con)
            CmdDel_obs_as.ExecuteNonQuery() : CmdDel_obs_as.Dispose()
            cmb_codigo_as_obs.Text = "" : RichTextBox2.Text = ""

            MsgBox("Rotina de deleção concluida!", MsgBoxStyle.Information)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub DG_01_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DG_01.ColumnHeaderMouseClick
        Try
            Dim str_status As String
            'Dim SQDR_status_opn As SqlDataReader
            For i = 0 To DG_01.NewRowIndex
                Try
                    DG_01.Columns(1).ValueType = GetType(Date)
                    str_status = Trim(DG_01.Rows(i).Cells(10).Value)
                    If str_status = "Em Aberto" Then
                        DG_01.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(em_aberto) '= Color.Yellow
                    ElseIf str_status = "Proposta Enviada" Then
                        DG_01.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(enviada)   'Color.GreenYellow
                    ElseIf str_status = "Proposta Aceita" Then
                        DG_01.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(aceita) 'Color.Green
                    ElseIf str_status = "Declinada" Then
                        DG_01.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(declinada)  'Color.Maroon
                        DG_01.Rows(i).DefaultCellStyle.ForeColor = ColorTranslator.FromHtml("#FFFFFF")
                    End If
                    'DG_01.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml("")
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "")
                End Try
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub DG_01_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DG_01.RowHeaderMouseClick
        Try

            cod_edit_OPN = DG_01.Item(0, e.RowIndex).Value

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub DG_01_UserDeletingRow(sender As Object, e As DataGridViewRowCancelEventArgs) Handles DG_01.UserDeletingRow
        Try
            If usuario_perfil = False Then Exit Sub

            If cod_edit_OPN <= 0 Then Exit Sub
            Beep()
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim Resposta As Integer
            Resposta = MessageBox.Show("Esta ação apagará a OPN selecionada e totas as AS relacionadas. Confirma essa ação", "RENATO", MessageBoxButtons.YesNo)
            If Resposta = 6 Then

                Dim CmdDel_obs_opn As New SqlCommand("Delete From OPN Where opn_codigo= " & cod_edit_OPN & " ", Con)
                CmdDel_obs_opn.ExecuteNonQuery() : CmdDel_obs_opn.Dispose()

                MsgBox("Ação concluida com êxito!", MsgBoxStyle.Information, titulo_opn)
            Else

            End If
            'ativar_filtros_status()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub DG_01_UserDeletedRow(sender As Object, e As DataGridViewRowEventArgs) Handles DG_01.UserDeletedRow
        Try
            ativar_filtros_status()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DG_02_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG_02.CellDoubleClick
        Try
            If e.RowIndex < 0 Then Exit Sub
            If usuario_perfil = False Then Exit Sub
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            as_consulta = True
            cod_edit_AS = DG_02.Item(0, e.RowIndex).Value
            cmb_codigo_as_obs.Text = DG_02.Item(0, e.RowIndex).Value
            'frm_opn_editar.Show()
            frm_as_cadastro.Show()
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'If Trim(DG_01.Rows(e.RowIndex).Cells(10).Value) = "Proposta Aceita" Then
            '    'frm_as_cadastro.Show()   'frm_as_cadastro
            'frm_opn_editar.Show()
            'Else
            '    'frm_opn_editar.Show()
            'End If

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub DG_02_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG_02.CellContentClick
        Try
            If e.RowIndex < 0 Then Exit Sub
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'as_consulta = True
            cod_edit_AS = DG_02.Item(0, e.RowIndex).Value
            cmb_codigo_as_obs.Text = DG_02.Item(0, e.RowIndex).Value
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Try
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim dt_01, dt_02 'As Date
            For i = 0 To DG_02.Rows.Count - 1

                Dim SQDR_obs_opn As SqlDataReader
                Dim SQCMDPorta As New SqlCommand("Select Distinct as_codigo From obs_as Where as_codigo = " & Trim(DG_02.Rows(i).Cells(0).Value) & "", Con)
                SQDR_obs_opn = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
                If SQDR_obs_opn.Read Then

                    DG_02.Item(0, i).Style.Font = New Font(DG_02.Font, FontStyle.Bold)
                    DG_02.Item(0, i).Style.BackColor = Color.White

                End If
                SQDR_obs_opn.Close()

                'If Trim(DG_01.Rows(i).Cells(10).Value) = "Em Aberto" Then  'And Not IsDBNull(Trim(DG.Rows(i).Cells(3).Value)) = False 

                '    If DG_01.Columns(2).Name.Equals("D_limite") Then
                '        Try
                '            If Not IsDBNull(Trim(DG_01.Rows(i).Cells(1).Value)) Then
                '                Try
                '                    If IsDate(Trim(DG_01.Rows(i).Cells(1).Value)) Then

                '                        dt_01 = CDate(Trim(DG_01.Rows(i).Cells(1).Value))
                '                        dt_02 = CDate(DG_01.Rows(i).Cells(2).Value)
                '                        Dim result As Long = DateDiff(DateInterval.Day, Date.Now, dt_02)
                '                        If result <= 1.8 Then
                '                            If DG_01.Item(2, i).Style.ForeColor <> Color.Red Then
                '                                DG_01.Item(2, i).Style.ForeColor = Color.Red
                '                                DG_01.Item(2, i).Style.BackColor = Color.Yellow
                '                            ElseIf DG_01.Item(2, i).Style.ForeColor = Color.Red Then
                '                                DG_01.Item(2, i).Style.ForeColor = Color.White
                '                                DG_01.Item(2, i).Style.BackColor = Color.Red
                '                            End If
                '                        Else

                '                        End If

                '                    End If
                '                Catch ex As Exception

                '                End Try

                '            End If
                '        Catch ex As Exception

                '        End Try

                '    End If
                'End If
            Next
        Catch ex As Exception
            ' MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub DG_01_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG_01.CellContentClick
        Try
            If e.RowIndex < 0 Then Exit Sub
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            cod_edit_OPN = DG_01.Item(0, e.RowIndex).Value
            cmb_codigo_onp_obs.Text = DG_01.Item(0, e.RowIndex).Value
            'frm_opn_editar.Show()
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'If Trim(DG_01.Rows(e.RowIndex).Cells(10).Value) = "Proposta Aceita" Then
            '    'frm_as_cadastro.Show()   'frm_as_cadastro
            'frm_opn_editar.Show()
            'Else
            '    'frm_opn_editar.Show()
            'End If

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem7.Click
        Try
            If ToolStripMenuItem7.Checked = True Then
                ToolStripMenuItem7.Checked = False
                ativar_filtros_status()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                ToolStripMenuItem7.Checked = True
                ativar_filtros_status()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub SuspensaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SuspensaToolStripMenuItem.Click
        Try
            If SuspensaToolStripMenuItem.Checked = True Then
                SuspensaToolStripMenuItem.Checked = False
                ativar_filtros_status()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                SuspensaToolStripMenuItem.Checked = True
                ativar_filtros_status()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub EmAbertoClienteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EmAbertoClienteToolStripMenuItem.Click
        Try
            If EmAbertoClienteToolStripMenuItem.Checked = True Then
                EmAbertoClienteToolStripMenuItem.Checked = False
                ativar_filtros_status()
            Else
                cmb_codigo.Text = "" : cmb_clientes.Text = ""
                EmAbertoClienteToolStripMenuItem.Checked = True
                ativar_filtros_status()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub cmb_codigo_Click(sender As Object, e As EventArgs) Handles cmb_codigo.Click

    End Sub

    Private Sub DataConvertToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataConvertToolStripMenuItem.Click
        Try
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim data_ini
            Dim a, m, d As String
            Dim newData As DateTime
            Cursor.Current = Cursors.WaitCursor
            Dim SQDR_obs_opn As SqlDataReader
            Dim SQCMDPorta As New SqlCommand("Select opn_data_abertura,opn_codigo From OPN Order By opn_data_abertura", Con)
            SQDR_obs_opn = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
            Do While SQDR_obs_opn.Read
                Try
                    data_ini = Split(Trim(SQDR_obs_opn("opn_data_abertura")), "/")
                    a = data_ini(2)
                    m = data_ini(1)
                    d = data_ini(0)
                    If IsDate(a & "-" & m & "-" & d) Then
                        newData = a & "-" & m & "-" & d
                    Else
                        'newData = ""
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                Try
                    'MsgBox(SQDR_obs_opn("opn_data_abertura"))
                    'newData = CDate(Trim(SQDR_obs_opn("opn_data_abertura")))
                    'MsgBox(newData)
                    Dim CmdIns_opn As New SqlCommand("Update OPN Set opn_data_abertura='" & newData & "' Where opn_codigo= " & SQDR_obs_opn("opn_codigo") & " ", Con)
                    CmdIns_opn.ExecuteNonQuery() : CmdIns_opn.Dispose()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

            Loop
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DG_01_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DG_01.CellFormatting
        If e.ColumnIndex = 1 Then 'your column
            'Dim d As Date
            'Try
            '    If Not IsDBNull(e.Value.ToString) Then
            '        If Date.TryParse(Trim(e.Value.ToString), d) Then
            '            e.Value = d.ToString("MM-dd-yyyy")
            '            'e.Value = CDate(d.ToString("dd-MM-yyyy"))
            '            e.FormattingApplied = True
            '        End If
            '    End If
            'Catch ex As Exception

            'End Try
        End If
    End Sub

    Private Sub DG_01_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DG_01.DataError
        Try

        Catch ex As Exception

        End Try
    End Sub
    'Public Sub ExportToExcel(ByVal dgvName As DataGridView, ByVal [option] As XlSortOn, Optional ByVal fileName As String = "")
    Public Sub ExportToExcel(ByVal dgvName As DataGridView)
        Dim objExcelApp As New Excel.Application()
        Dim objExcelBook As Excel.Workbook
        Dim objExcelSheet As Excel.Worksheet

        Dim Con2 As SqlConnection = TratadorDeConexao.Conexao2()
        Dim SQDR_cli_totvs As SqlDataReader
        Dim Con As SqlConnection = TratadorDeConexao.Conexao()
        Dim SQDR_as As SqlDataReader
        Try


            objExcelBook = objExcelApp.Workbooks.Add
            objExcelSheet = CType(objExcelBook.Worksheets(1), Excel.Worksheet)
            objExcelApp.Visible = True
            ' Ciclo nos cabeçalhos para escrever os títulos a bold/negrito
            Dim dgvColumnIndex As Int16 = 2
            objExcelSheet.Cells(1, dgvColumnIndex) = "Cliente"
            objExcelSheet.Cells(1, dgvColumnIndex).Font.Bold = True
            dgvColumnIndex = 3
            For Each col As DataGridViewColumn In dgvName.Columns
                ' MsgBox(col.HeaderText)
                If dgvColumnIndex = 13 Or dgvColumnIndex = 14 Or dgvColumnIndex = 15 Or dgvColumnIndex = 16 Then

                Else
                    objExcelSheet.Cells(1, dgvColumnIndex) = Trim(col.HeaderText)
                    objExcelSheet.Cells(1, dgvColumnIndex).Font.Bold = True
                End If

                dgvColumnIndex += 1
                'If dgvColumnIndex = 15 Then Exit For
            Next
            'Exit Sub
            ' Ciclo nas linhas/células
            Dim dgvRowIndex As Integer = 2
            Dim str_defeitor_relatados As String
            For Each row As DataGridViewRow In dgvName.Rows

                Dim dgvCellIndex As Integer = 3
                str_defeitor_relatados = ""
                For Each cell As DataGridViewCell In row.Cells
                    ''''''''''''''''''''''''''''''''''''''''''''''''''
                    Try
                        If Not IsDBNull(Trim(cell.Value)) Then
                            'If dgvColumnIndex = 13 Or dgvColumnIndex = 14 Or dgvColumnIndex = 15 Or dgvColumnIndex = 16 Then

                            'Else
                            If dgvCellIndex = 3 Then
                                'Codigo da AS
                                Dim SQCMD_as As New SqlCommand("Select * From [AS] where as_codigo = " & Trim(cell.Value) & "", Con)
                                SQDR_as = SQCMD_as.ExecuteReader(CommandBehavior.Default)
                                If SQDR_as.Read Then
                                    'Dim cli_totvs = SQDR_as("as_cliente_totvs")
                                    Dim SQCMDcc As New SqlCommand("Select * From [Exibicao_Cliente] Where A1_COD = " & Trim(SQDR_as("as_cliente_totvs")) & "", Con2)
                                    SQDR_cli_totvs = SQCMDcc.ExecuteReader(CommandBehavior.Default)
                                    If SQDR_cli_totvs.Read Then
                                        'Dim ttt = Trim(SQDR_cli_totvs("A1_NOME"))
                                        dgvCellIndex = dgvCellIndex - 1
                                        objExcelSheet.Cells(dgvRowIndex, dgvCellIndex) = "'" & Trim(SQDR_cli_totvs("A1_NOME"))
                                        dgvCellIndex = dgvCellIndex + 1
                                    End If
                                    'encontraAS(cod_edit_OPN, 0)
                                End If
                                SQDR_as.Close()
                            End If
                            If dgvCellIndex = 13 Or dgvCellIndex = 14 Or dgvCellIndex = 15 Or dgvCellIndex = 16 Then

                            ElseIf dgvCellIndex = 17 Then
                                objExcelSheet.Cells(dgvRowIndex, dgvCellIndex) = Format(cell.Value, "###,###,##0.00")
                            Else
                                objExcelSheet.Cells(dgvRowIndex, dgvCellIndex) = "'" & Trim(cell.Value)
                            End If
                            'End If
                        Else
                            objExcelSheet.Cells(dgvRowIndex, dgvCellIndex) = "'" '& Trim(cell.Value)
                        End If

                        dgvCellIndex += 1
                        'If dgvCellIndex = 15 Then Exit For
                    Catch ex As Exception
                        dgvCellIndex += 1
                    End Try
                    ''''''''''''''''''''''''''''''''''''''''''''''''''
                Next
                '''''''''''''''''''''''''''''''''''''''''''''''''''
                ''Alterado por renato em 06-10-2015
                'objExcelSheet.Cells(dgvRowIndex, dgvCellIndex + 1) = "'" & str_defeitor_relatados
                'str_defeitor_relatados = ""
                ''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''
                dgvRowIndex += 1
                'If dgvRowIndex = 17 Then Exit For
            Next

            ' Ajusta o largura das colunas automaticamente
            objExcelSheet.Columns.AutoFit()



        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)

        Finally

            objExcelSheet = Nothing
            objExcelBook = Nothing
            objExcelApp = Nothing

            ' O GC(garbage collector) recolhe a memória não usada pelo sistema. 
            ' O método Collect() força a recolha e a opção WaitForPendingFinalizers 
            ' espera até estar completo. Desta forma o EXCEL.EXE não fica no 
            ' Task Manager(gestor tarefas) ocupando memória desnecessariamente
            ' (devem ser chamados duas vezes para maior garantia)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try

    End Sub

    Private Sub frm_opn_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Try
            If e.KeyCode = Keys.F11 Then
                ExportToExcel(DG_01)
            ElseIf e.KeyCode = Keys.F4 Then
                ExportToExcel(DG_02)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ToolStripExportarExcel_Click(sender As Object, e As EventArgs) Handles ToolStripExportarExcel.Click
        ExportToExcel(DG_02)
    End Sub

    Private Sub ToolStripConsultar_Click(sender As Object, e As EventArgs) Handles ToolStripConsultar.Click
        Try
            ToolStripMenuItem3.Checked = False
            ToolStripMenuItem4.Checked = False
            ToolStripMenuItem5.Checked = False
            ToolStripMenuItem6.Checked = False
            cmb_codigo_opn.Text = ""
            'ExibeDados(DG_01, "Select * From Lista_OPN Where Cliente Like '%" & Trim(cmb_cliente_opn.Text) & "%' Order By Prioridade", "Lista_OPN")
            If cmb_cliente_opn.Text <> "" Then
                ExibeDados(DG_02, "Select * From Lista_AS Where [as_cliente_totvs] Like '%" & Trim(cmb_cliente_opn.Text) & "%' And convert(datetime, [D. de Abertua], 103) >= convert(datetime, '" & ToolStripDataInicio.Text & "', 103) AND convert(datetime, [D. de Abertua], 103) <= convert(datetime, '" & ToolStripDataFinal.Text & "', 103) Order By convert(datetime, [D. de Abertua], 103)", "Lista_AS")
            Else
                ExibeDados(DG_02, "Select * From Lista_AS Where convert(datetime, [D. de Abertua], 103) >= convert(datetime, '" & ToolStripDataInicio.Text & "', 103) AND  convert(datetime, [D. de Abertua], 103) <= convert(datetime, '" & ToolStripDataFinal.Text & "', 103) Order By convert(datetime, [D. de Abertua], 103)", "Lista_AS")
            End If
            'ExibeDados(DG_01, "Select * From Lista_OPN Where Cliente Like '%" & Trim(cmb_cliente_opn.Text) & "%' Order By Prioridade", "Lista_OPN")

            'Select * From [AS] where convert(datetime, as_data_abertura, 105) >= convert(datetime, '01-12-2014', 105) AND  convert(datetime, as_data_abertura, 105) <= convert(datetime, '31-12-2014', 105)
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub

    Private Sub ToolStripConsultarOPN_bt_Click(sender As Object, e As EventArgs) Handles ToolStripConsultarOPN_bt.Click
        Try
            '
            'ExibeDados(DG_01, "Select * From Lista_OPN Where Prioridade = " & status_01 & " or Prioridade = " & status_02 & " or Prioridade = " & status_03 & " or Prioridade = " & status_04 & " Order By Status", "Lista_OPN")
            ExibeDados(DG_01, "Select * From Lista_OPN Where [Status] = '" & ToolStripStatus_cb.Text & "'", "Lista_OPN")

            'ExibeDados(DG_01, "Select * From Lista_OPN Where [Status] Like '%" & Trim(ToolStripStatus_cb.Text) & "%' And convert(datetime, [Data], 103) >= convert(datetime, '" & ToolStripDInicio_tx.Text & "', 103) AND convert(datetime, [Data], 103) <= convert(datetime, '" & ToolStripDFinal_tx.Text & "', 103) Order By convert(datetime, [Data], 103)", "Lista_OPN")
            'Select * From Lista_OPN Where [Status] Like '%" & Trim(cmb_cliente_opn.Text) & "%' And convert(datetime, [Data], 103) >= convert(datetime, '" & ToolStripDataInicio.Text & "', 103) AND convert(datetime, [Data], 103) <= convert(datetime, '" & ToolStripDataFinal.Text & "', 103) Order By convert(datetime, [Data], 103)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ToolStripStatus_cb_Click(sender As Object, e As EventArgs) Handles ToolStripStatus_cb.Click

    End Sub

    Private Sub ToolStripStatus_cb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripStatus_cb.SelectedIndexChanged
        Try
            ExibeDados(DG_01, "Select * From Lista_OPN Where [Status] = '" & ToolStripStatus_cb.Text & "'", "Lista_OPN")
        Catch ex As Exception

        End Try
    End Sub
End Class