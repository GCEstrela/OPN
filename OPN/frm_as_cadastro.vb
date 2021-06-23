Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Net.Mail
Imports System.Data.OleDb
'Imports Microsoft.Office.Interop.Excel

Imports Excel = Microsoft.Office.Interop.Excel
Public Class frm_as_cadastro
    Private WM_NCHITTEST As Integer = &H84
    Private HTCLIENT As Integer = &H1
    Private HTCAPTION As Integer = &H2
    Private cod_cliente_totvs As Integer
    Private _descricao_as_ccusto2 As String
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Try
            MyBase.WndProc(m)
            Select Case m.Msg
                Case WM_NCHITTEST
                    If m.Result = New IntPtr(HTCLIENT) Then
                        m.Result = New IntPtr(HTCAPTION)
                    End If
            End Select
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub bnt_sair_Click(sender As Object, e As EventArgs) Handles bnt_sair.Click
        Try

            as_consulta = False
            GC.Collect() : GC.WaitForPendingFinalizers()
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub frm_as_cadastro_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        Try
            as_consulta = False
        Catch ex As Exception

        End Try
    End Sub

    Private Sub frm_as_cadastro_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'If as_consulta = False Then
            Me.Cursor = Cursors.WaitCursor
            Dim Con2 As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_max As SqlDataReader
            Dim SQDR_opn_as As SqlDataReader
            Dim titulo As String = "Sem Titulo..."

            Dim SQCMDmax As New SqlCommand("Select max(as_codigo) as cod_max From [AS]", Con2)
            SQDR_max = SQCMDmax.ExecuteReader(CommandBehavior.Default)
            If SQDR_max.Read Then
                If Not IsDBNull(SQDR_max("cod_max")) Then
                    txt_cod_AS.Text = SQDR_max("cod_max") + 1
                Else
                    txt_cod_AS.Text = "1001"
                End If
            Else
                txt_cod_AS.Text = "1001" + 1
            End If
            SQDR_max.Close()

            tx_as_01.Text = Date.Now.ToString("dd/MM/yyyy")
            '
            benc_Combo(cmb_codigo_opn, "Lista_OPN", "Select OPN From Lista_OPN Where Prioridade = " & 2 & " Order By OPN", "OPN")

            benc_Combo2(cmb_clientes, "Exibicao_Cliente", "Select Distinct A1_NOME From Exibicao_Cliente Order By A1_NOME", "A1_NOME")

            'benc_Combo2(cmb_ccusto, "Exibicao_CCusto", "Select CTT_DESC01,Registro From Exibicao_CCusto Order By CTT_DESC01", "CTT_DESC01", "Registro")

            benc_Combo(cmb_as_status, "Status_AS", "Select Distinct Descricao From Status_AS Order By Descricao", "Descricao")
            cmb_as_status.Text = "À Executar"

            'txt_01.Text = Date.Now.ToString("dd/MM/yyyy")
            Dim SQCMD_opn_as As New SqlCommand("Select * From [AS] where as_codigo = " & cod_edit_OPN & "", Con2)
            SQDR_opn_as = SQCMD_opn_as.ExecuteReader(CommandBehavior.Default)
            If SQDR_opn_as.Read Then
                cod_edit_OPN = SQDR_opn_as("opn_codigo")
                encontraAS(cod_edit_OPN, 0)
            End If
            SQDR_opn_as.Close()
            'encontraAS(cod_edit_AS, 0)

            'Else
            If as_consulta Then

                txt_cod_AS.Text = cod_edit_AS
                txt_cod_AS.Focus() ': SendKeys.Send("{ENTER}")
                'cmb_clientes.Focus()

            End If
            Me.Cursor = Cursors.Hand
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmb_ccusto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_ccusto.SelectedIndexChanged
        Try
            oldlbl_descricao.Text = ""
            Dim Con2 As SqlConnection = TratadorDeConexao.Conexao2()
            Dim SQDR_ccusto As SqlDataReader
            Dim vetor
            'vetor = Split("dfasçdfkaçfk-sadfjalsfjalk", "-")
            Dim str_cliente_selecionado = Split(Trim(cmb_ccusto.Text), "*")

            Dim SQCMDcc As New SqlCommand("Select * From [Exibicao_CCusto] Where  CTT_DESC01 = '" & Trim(str_cliente_selecionado(0)) & "' And Registro = " & str_cliente_selecionado(1) & "", Con2)
            SQDR_ccusto = SQCMDcc.ExecuteReader(CommandBehavior.Default)
            If SQDR_ccusto.Read Then

                oldlbl_descricao.Text = Trim(SQDR_ccusto("CTT_CUSTO"))
                lbl_descricao.Text = Trim(SQDR_ccusto("CTT_CUSTO"))
                '_descricao_as_ccusto2 = Trim(SQDR_ccusto("CTT_CUSTO"))
            End If
            SQDR_ccusto.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmb_clientes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_clientes.SelectedIndexChanged
        Try

            Dim Con2 As SqlConnection = TratadorDeConexao.Conexao2()
            Dim SQDR_ccusto As SqlDataReader
            Dim str_tel As String
            Dim titulo As String = "Lauto Tecnico"
            cod_cliente_totvs = 0
            Dim SQCMDcc As New SqlCommand("Select * From [Exibicao_Cliente] Where  A1_NOME = '" & Trim(cmb_clientes.Text) & "'", Con2)
            SQDR_ccusto = SQCMDcc.ExecuteReader(CommandBehavior.Default)
            If SQDR_ccusto.Read Then

                cod_cliente_totvs = SQDR_ccusto("A1_COD")

                'If Not IsDBNull(SQDR_ccusto("A1_CONTATO")) Then
                '    tx_as_02.Text = Trim(SQDR_ccusto("A1_CONTATO"))
                'Else
                '    tx_as_02.Text = ""
                'End If
                'If Len(Trim(SQDR_ccusto("A1_DDD"))) < 3 Then
                '    str_tel = "0" & Trim(SQDR_ccusto("A1_DDD")) & Trim(SQDR_ccusto("A1_TEL"))
                'Else
                '    str_tel = Trim(SQDR_ccusto("A1_DDD")) & Trim(SQDR_ccusto("A1_TEL"))
                'End If
                'tx_as_03.Text = str_tel

                'If Not IsDBNull(SQDR_ccusto("A1_EMAIL")) Then
                '    tx_as_04.Text = Trim(SQDR_ccusto("A1_EMAIL"))
                'Else
                '    tx_as_04.Text = ""
                'End If
                'If Not IsDBNull(SQDR_ccusto("A1_END")) Then
                '    tx_as_05.Text = Trim(SQDR_ccusto("A1_END"))
                'Else
                '    tx_as_05.Text = ""
                'End If

                benc_Combo2(cmb_ccusto, "Exibicao_CCusto", "Select CTT_DESC01,Registro From Exibicao_CCusto Order By CTT_DESC01", "CTT_DESC01", "Registro")
            End If
            SQDR_ccusto.Close()

            'cmb_ccusto.Text = _descricao_as_ccusto2

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txt_01_KeyDown(sender As Object, e As KeyEventArgs) Handles txt_01.KeyDown
        Try
            If e.KeyCode <> 13 Then Exit Sub

            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_opn As SqlDataReader

            cod_edit_OPN = Val(txt_01.Text)
            Dim str_staturs As String = "Laboratório"
            Dim SQCMDopn As New SqlCommand("Select * From OPN Where opn_codigo = " & cod_edit_OPN & " And opn_status = " & 3 & "", Con)
            SQDR_opn = SQCMDopn.ExecuteReader(CommandBehavior.Default)
            If SQDR_opn.Read Then

                If Not IsDBNull(SQDR_opn("opn_alterado_por")) Then
                    tx_op_01.Text = Trim(SQDR_opn("opn_alterado_por"))
                Else
                    tx_op_01.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_data_abertura")) Then
                    tx_op_02.Text = Trim(SQDR_opn("opn_data_abertura"))
                Else
                    tx_op_02.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_cliente")) Then
                    tx_op_04.Text = Trim(SQDR_opn("opn_cliente"))
                Else
                    tx_op_04.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_obs")) Then
                    tx_as_10.Text = Trim(SQDR_opn("opn_obs"))
                Else
                    tx_as_10.Text = ""
                End If

                'Descrição do Status do Cliente
                tx_op_03.Text = ""
                Dim SQDR_status As SqlDataReader
                Dim SQCMDStatus As New SqlCommand("Select * From Status Where Codigo = " & Val(SQDR_opn("opn_status")) & "", Con)
                SQDR_status = SQCMDStatus.ExecuteReader(CommandBehavior.Default)
                If SQDR_status.Read Then
                    tx_op_03.Text = Trim(SQDR_status("Descricao"))
                Else
                    tx_op_03.Text = ""
                End If
                SQDR_status.Close()
            Else
                MsgBox("OPN com status inválido!", MsgBoxStyle.Information, "OPN")
            End If
            SQDR_opn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txt_01_TextChanged(sender As Object, e As EventArgs) Handles txt_01.TextChanged

    End Sub

    Private Sub bnt_cadastro_Click(sender As Object, e As EventArgs) Handles bnt_cadastro.Click
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim _campos As String = "Z01_NUMERO,Z01_DESCRI,Z01_CLIENT,R_E_C_N_O_"
            Dim _VALORES As String = "'" & Trim(txt_cod_AS.Text).PadLeft(6, "0") & "','" & Trim(tx_as_07.Text) & "','" & cod_cliente_totvs.ToString().PadLeft(6, "0") & "','" & Trim(txt_cod_AS.Text) & "'"
            Dim _CAMPOS_VALORES As String = "Z01_DESCRI='" & Trim(tx_as_07.Text) & "',Z01_CLIENT='" & cod_cliente_totvs.ToString().PadLeft(6, "0") & "',R_E_C_N_O_='" & Trim(txt_cod_AS.Text) & "',Z01_NUMERO='" & Trim(txt_cod_AS.Text).PadLeft(6, "0") & "'"

            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim Con3 As SqlConnection = TratadorDeConexao.Conexao2()

            Dim SQDR_as As SqlDataReader
            Dim SQASTOTVS As SqlDataReader
            'Dim SQDR_max As SqlDataReader

            If Val(cod_status_as) <= 0 Then
                MsgBox("Obs: Campo Status não esta definido!")
                cmb_as_status.Focus()
                Exit Sub
            End If

            Dim SQCMDAS As New SqlCommand("Select * From [AS] Where as_codigo = " & txt_cod_AS.Text & "", Con)
            SQDR_as = SQCMDAS.ExecuteReader(CommandBehavior.Default)
            If SQDR_as.Read Then

                Dim CmdUP_as As New SqlCommand("Update [AS] Set as_data_abertura='" & Trim(tx_as_01.Text) & "',as_centro_custo='" & Trim(lbl_descricao.Text) & "',as_contrato_pedido='" & Trim(tx_as_06.Text) & "',as_inicio_contrato='" & Trim(DateTimePicker1.Text) & "',as_fim_contrato='" & Trim(DateTimePicker2.Text) & "',as_objeto='" & Trim(tx_as_07.Text) & "',as_prazo_execucao='" & Trim(tx_as_08.Text) & "',as_doc_referencia='" & Trim(tx_as_09.Text) & "',as_obs='" & Trim(tx_as_10.Text) & "',as_cliente_totvs=" & cod_cliente_totvs & ",opn_codigo=" & cod_edit_OPN & ",as_status=" & cod_status_as & ",as_descricao_ccusto2='" & Trim(cmb_ccusto.Text) & "'  Where as_codigo= " & Trim(txt_cod_AS.Text) & "", Con)
                CmdUP_as.ExecuteNonQuery() : CmdUP_as.Dispose()


                'Verifica se a AS foi criada no Totvs para atualizar ou inserir
                Dim QUERYTOTVS As New SqlCommand("Select * From [Z01010] Where Z01_NUMERO= " & Trim(txt_cod_AS.Text).PadLeft(6, "0") & "", Con3)
                SQASTOTVS = QUERYTOTVS.ExecuteReader(CommandBehavior.Default)

                If SQASTOTVS.Read Then
                    Dim CmdUP_as_TOTVS As New SqlCommand("Update [Z01010] Set " & _CAMPOS_VALORES & " Where Z01_NUMERO= " & Trim(txt_cod_AS.Text).PadLeft(6, "0") & "", Con3)
                    CmdUP_as_TOTVS.ExecuteNonQuery() : CmdUP_as_TOTVS.Dispose()
                Else
                    Dim CmdIns_as_TOTVS As New SqlCommand("Insert into Z01010 ( " & _campos & " ) values ( " & _VALORES & " )", Con3)
                    CmdIns_as_TOTVS.ExecuteNonQuery() : CmdIns_as_TOTVS.Dispose()
                End If


            Else
                    Try
                    Dim CmdIns_as As New SqlCommand("Insert into [AS](as_data_abertura,as_centro_custo,as_contrato_pedido,as_inicio_contrato,as_fim_contrato,as_objeto,as_prazo_execucao,as_doc_referencia,as_obs,as_cliente_totvs,opn_codigo,as_status,as_codigo,as_descricao_ccusto2) values ('" & Trim(tx_as_01.Text) & "','" & Trim(lbl_descricao.Text) & "','" & Trim(tx_as_06.Text) & "','" & Trim(DateTimePicker1.Text) & "','" & Trim(DateTimePicker2.Text) & "','" & Trim(tx_as_07.Text) & "','" & Trim(tx_as_08.Text) & "','" & Trim(tx_as_09.Text) & "','" & Trim(tx_as_10.Text) & "'," & cod_cliente_totvs & "," & cod_edit_OPN & "," & cod_status_as & "," & Val(txt_cod_AS.Text) & ",'" & Trim(cmb_ccusto.Text) & "')", Con)
                    CmdIns_as.ExecuteNonQuery() : CmdIns_as.Dispose()

                    Dim CmdIns_as_TOTVS As New SqlCommand("Insert into Z01010 ( " & _campos & " ) values ( " & _VALORES & " )", Con3)
                    CmdIns_as_TOTVS.ExecuteNonQuery() : CmdIns_as_TOTVS.Dispose()

                    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '        ExibeDados(frm_principal.DG, "Select * From Lista_OPN Order By Prioridade", "Lista_OPN")
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Catch ex As Exception
                    MsgBox(ex.Message)
                    SQDR_as.Close()
                    Exit Sub
                End Try

                '    Dim SQCMDEqui As New SqlCommand("Select max(opn_codigo) as cod_max From OPN", Con)
                '    SQDR_max = SQCMDEqui.ExecuteReader(CommandBehavior.Default)
                '    If SQDR_max.Read Then
                '        txt_01.Text = SQDR_max("cod_max") + 1
                '    End If
                '    SQDR_max.Close()

            End If
            SQDR_as.Close()

            MsgBox("Ação executada com êxito!", MsgBoxStyle.Information, titulo_as)

            frm_opn.ativar_filtros_status_AS()
            'Me.Focus()
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub frm_as_cadastro_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick
        MsgBox(Me.Top & "     " & Me.Left)
    End Sub
    Private Sub encontraAS(ByVal c_opn As Integer, ByVal c_as As Integer)
        Try

            If c_opn = 0 Then Exit Sub

            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_opn As SqlDataReader

            Dim str_staturs As String = "Laboratório"
            Dim SQCMDopn As New SqlCommand("Select * From OPN Where opn_codigo = " & c_opn & "", Con)
            SQDR_opn = SQCMDopn.ExecuteReader(CommandBehavior.Default)
            If SQDR_opn.Read Then
                'txt_01.Text = c_opn
                cmb_codigo_opn.Text = c_opn
                If Not IsDBNull(SQDR_opn("opn_alterado_por")) Then
                    tx_op_01.Text = Trim(SQDR_opn("opn_alterado_por"))
                Else
                    tx_op_01.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_data_abertura")) Then
                    tx_op_02.Text = Trim(SQDR_opn("opn_data_abertura"))
                Else
                    tx_op_02.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_cliente")) Then
                    tx_op_04.Text = Trim(SQDR_opn("opn_cliente"))
                Else
                    tx_op_04.Text = ""
                End If


                'Descrição do Status do Cliente
                tx_op_03.Text = ""
                Dim SQDR_status As SqlDataReader
                Dim SQCMDStatus As New SqlCommand("Select * From Status Where Codigo = " & Val(SQDR_opn("opn_status")) & "", Con)
                SQDR_status = SQCMDStatus.ExecuteReader(CommandBehavior.Default)
                If SQDR_status.Read Then
                    tx_op_03.Text = Trim(SQDR_status("Descricao"))
                Else
                    tx_op_03.Text = ""
                End If
                SQDR_status.Close()
            Else
                MsgBox("OPN com status inválido!", MsgBoxStyle.Information, "OPN")
            End If
            SQDR_opn.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub txt_cod_AS_KeyDown(sender As Object, e As KeyEventArgs) Handles txt_cod_AS.KeyDown
        Try
            If e.KeyCode <> Keys.Enter Then Exit Sub
            If Val(txt_cod_AS.Text) = 0 Then Exit Sub
            cod_edit_AS = Val(txt_cod_AS.Text)


            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_as As SqlDataReader
            Dim SQDR_cc_totvs As SqlDataReader

            'Dim str_staturs As String = "Laboratório"
            Dim SQCMDopn As New SqlCommand("Select * From [AS] Where as_codigo = " & cod_edit_AS & "", Con)
            SQDR_as = SQCMDopn.ExecuteReader(CommandBehavior.Default)
            If SQDR_as.Read Then
                'Vai buscar os dados no cliente no Totvs
                If Not IsDBNull(SQDR_as("as_cliente_totvs")) Then
                    encontra_cliente_totvs(SQDR_as("as_cliente_totvs"))
                Else
                    encontra_cliente_totvs(0)
                    tx_op_01.Text = ""
                End If

                If Not IsDBNull(SQDR_as("as_data_abertura")) Then
                    tx_op_01.Text = Trim(SQDR_as("as_data_abertura"))
                Else
                    tx_op_01.Text = ""
                End If
                If Not IsDBNull(SQDR_as("as_obs")) Then
                    tx_as_10.Text = Trim(SQDR_as("as_obs"))
                Else
                    tx_as_10.Text = ""
                End If

                cmb_ccusto.Text = "" : oldlbl_descricao.Text = ""
                If Not IsDBNull(SQDR_as("as_centro_custo")) Then
                    lbl_descricao.Text = Trim(SQDR_as("as_centro_custo"))

                    Dim Con2 As SqlConnection = TratadorDeConexao.Conexao2()
                    Dim SQCMDcc As New SqlCommand("Select * From [Exibicao_CCusto] Where  CTT_CUSTO = '" & Trim(SQDR_as("as_centro_custo")) & "'", Con2)
                    SQDR_cc_totvs = SQCMDcc.ExecuteReader(CommandBehavior.Default)
                    If SQDR_cc_totvs.Read Then
                        cmb_ccusto.Text = Trim(SQDR_cc_totvs("CTT_DESC01"))
                    End If
                    SQDR_cc_totvs.Close()

                Else
                    oldlbl_descricao.Text = ""
                    cmb_ccusto.Text = ""
                End If
                If Not IsDBNull(SQDR_as("as_contrato_pedido")) Then
                    tx_as_06.Text = Trim(SQDR_as("as_contrato_pedido"))
                Else
                    tx_as_06.Text = ""
                End If
                If Not IsDBNull(SQDR_as("as_inicio_contrato")) Then
                    DateTimePicker1.Text = Trim(SQDR_as("as_inicio_contrato"))
                Else
                    DateTimePicker1.Text = ""
                End If
                If Not IsDBNull(SQDR_as("as_fim_contrato")) Then
                    DateTimePicker2.Text = Trim(SQDR_as("as_fim_contrato"))
                Else
                    DateTimePicker1.Text = ""
                End If
                If Not IsDBNull(SQDR_as("as_objeto")) Then
                    tx_as_07.Text = Trim(SQDR_as("as_objeto"))
                Else
                    tx_as_07.Text = ""
                End If
                If Not IsDBNull(SQDR_as("as_prazo_execucao")) Then
                    tx_as_08.Text = Trim(SQDR_as("as_prazo_execucao"))
                Else
                    tx_as_08.Text = ""
                End If
                If Not IsDBNull(SQDR_as("as_doc_referencia")) Then
                    tx_as_09.Text = Trim(SQDR_as("as_doc_referencia"))
                Else
                    tx_as_09.Text = ""
                End If
                If Not IsDBNull(SQDR_as("as_descricao_ccusto2")) Then
                    _descricao_as_ccusto2 = Trim(SQDR_as("as_descricao_ccusto2"))
                Else
                    _descricao_as_ccusto2 = ""
                End If
                'If Not IsDBNull(SQDR_as("as_obs")) Then
                '    tx_as_10.Text = Trim(SQDR_as("as_obs"))
                'Else
                '    tx_as_10.Text = ""
                'End If
                encontraAS(SQDR_as("opn_codigo"), 0)

                'Descrição do Status do Cliente
                'tx_op_03.Text = ""
                Dim SQDR_status As SqlDataReader
                Dim SQCMDStatus As New SqlCommand("Select * From Lista_AS Where [Código] = " & Val(SQDR_as("as_codigo")) & "", Con)
                SQDR_status = SQCMDStatus.ExecuteReader(CommandBehavior.Default)
                If SQDR_status.Read Then
                    cmb_as_status.Text = Trim(SQDR_status("Status"))

                    'Descrição do Status do Cliente
                    Dim SQDR_status_as As SqlDataReader
                    Dim SQCMDStatus_As As New SqlCommand("Select * From Status_AS Where Codigo = " & Val(SQDR_status("Codigo")) & "", Con)
                    SQDR_status_as = SQCMDStatus_As.ExecuteReader(CommandBehavior.Default)
                    If SQDR_status_as.Read Then

                        'cod_status_edit = Val(SQDR_opn("opn_status"))
                        'cmb_status.Text = Trim(SQDR_status("Descricao"))

                        Dim _status = Trim(SQDR_status_as("Descricao"))
                        If _status = "À Executar" Then
                            txt_cod_AS.BackColor = ColorTranslator.FromHtml(aexecutar)     'Color.Yellow
                        ElseIf _status = "Em Execução/Pendente" Then
                            txt_cod_AS.BackColor = ColorTranslator.FromHtml(execucao_pendente) 'Color.GreenYellow
                        ElseIf _status = "Em Execução" Then
                            txt_cod_AS.BackColor = ColorTranslator.FromHtml(em_execucao) 'Color.Green
                        ElseIf _status = "Executada" Then
                            txt_cod_AS.BackColor = ColorTranslator.FromHtml(executada) 'Color.Maroon
                        ElseIf _status = "Cancelada" Then
                            txt_cod_AS.BackColor = ColorTranslator.FromHtml(cancelada) 'Color.Maroon
                        End If
                    Else

                        'cmb_status.Text = ""

                    End If
                    SQDR_status_as.Close()

                Else
                    cmb_as_status.Text = ""
                End If
                SQDR_status.Close()
            Else
                MsgBox("OPN com status inválido!", MsgBoxStyle.Information, "OPN")
            End If
            SQDR_as.Close()

            cmb_ccusto.Text = _descricao_as_ccusto2

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub encontra_cliente_totvs(ByVal codio_cliente_totvs As Integer)
        Try

            Dim Con2 As SqlConnection = TratadorDeConexao.Conexao2()
            Dim SQDR_ccusto As SqlDataReader
            Dim str_tel As String
            'Dim titulo As String = "Aut"
            cod_cliente_totvs = 0
            Dim SQCMDcc As New SqlCommand("Select * From [Exibicao_Cliente] Where  A1_COD = " & codio_cliente_totvs & "", Con2)
            SQDR_ccusto = SQCMDcc.ExecuteReader(CommandBehavior.Default)
            If SQDR_ccusto.Read Then

                cod_cliente_totvs = SQDR_ccusto("A1_COD")


                If Not IsDBNull(SQDR_ccusto("A1_NOME")) Then
                    cmb_clientes.Text = Trim(SQDR_ccusto("A1_NOME"))
                Else
                    cmb_clientes.Text = ""
                End If

                If Not IsDBNull(SQDR_ccusto("A1_CONTATO")) Then
                    tx_as_02.Text = Trim(SQDR_ccusto("A1_CONTATO"))
                Else
                    tx_as_02.Text = ""
                End If
                If Len(Trim(SQDR_ccusto("A1_DDD"))) < 3 Then
                    str_tel = "0" & Trim(SQDR_ccusto("A1_DDD")) & Trim(SQDR_ccusto("A1_TEL"))
                Else
                    str_tel = Trim(SQDR_ccusto("A1_DDD")) & Trim(SQDR_ccusto("A1_TEL"))
                End If
                tx_as_03.Text = str_tel

                If Not IsDBNull(SQDR_ccusto("A1_EMAIL")) Then
                    tx_as_04.Text = Trim(SQDR_ccusto("A1_EMAIL"))
                Else
                    tx_as_04.Text = ""
                End If
                If Not IsDBNull(SQDR_ccusto("A1_END")) Then
                    tx_as_05.Text = Trim(SQDR_ccusto("A1_END"))
                Else
                    tx_as_05.Text = ""
                End If
            Else
                cod_cliente_totvs = 0
                cmb_clientes.Text = ""

                tx_as_02.Text = ""
                str_tel = ""
                tx_as_03.Text = str_tel

                tx_as_04.Text = ""
                tx_as_05.Text = ""
            End If
            SQDR_ccusto.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub cmb_as_status_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_as_status.SelectedIndexChanged
        Try
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_max As SqlDataReader
            Dim titulo As String = "Lauto Tecnico"
            'cod_status = 0

            Dim SQCMDmax As New SqlCommand("Select * From Status_AS Where Descricao = '" & Trim(cmb_as_status.Text) & "'", Con)
            SQDR_max = SQCMDmax.ExecuteReader(CommandBehavior.Default)
            If SQDR_max.Read Then

                cod_status_as = SQDR_max("Codigo")

            End If
            SQDR_max.Close()
            'cod_status = 1
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim caminho_app = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim excel As Application = New Application
            Dim w As Workbook = excel.Workbooks.Open(caminho_app & "\AS.xlsx")


            excel.Workbooks(1).Worksheets(1).cells(4, 2).value = Trim(txt_cod_AS.Text)
            excel.Workbooks(1).Worksheets(1).cells(5, 2).value = Trim(tx_as_01.Text)
            excel.Workbooks(1).Worksheets(1).cells(6, 2).value = Trim(cmb_clientes.Text)
            excel.Workbooks(1).Worksheets(1).cells(7, 2).value = Trim(tx_as_02.Text)
            excel.Workbooks(1).Worksheets(1).cells(8, 2).value = Trim(tx_as_03.Text)

            excel.Workbooks(1).Worksheets(1).cells(9, 2).value = Trim(tx_as_04.Text)
            excel.Workbooks(1).Worksheets(1).cells(10, 2).value = Trim(tx_as_05.Text)
            excel.Workbooks(1).Worksheets(1).cells(11, 2).value = Trim(oldlbl_descricao.Text)
            excel.Workbooks(1).Worksheets(1).cells(12, 2).value = Trim(tx_as_06.Text)
            excel.Workbooks(1).Worksheets(1).cells(13, 2).value = "'" & Trim(DateTimePicker1.Text)
            excel.Workbooks(1).Worksheets(1).cells(14, 2).value = "'" & Trim(DateTimePicker2.Text)
            excel.Workbooks(1).Worksheets(1).cells(15, 2).value = Trim(tx_as_07.Text)
            excel.Workbooks(1).Worksheets(1).cells(4, 4).value = Trim(tx_as_08.Text)
            excel.Workbooks(1).Worksheets(1).cells(5, 4).value = Trim(tx_as_09.Text)
            excel.Workbooks(1).Worksheets(1).cells(7, 4).value = Trim(tx_as_10.Text)

            excel.Workbooks(1).Worksheets(1).cells(11, 4).value = Trim(txt_01.Text)
            excel.Workbooks(1).Worksheets(1).cells(12, 4).value = Trim(tx_op_01.Text)
            excel.Workbooks(1).Worksheets(1).cells(13, 4).value = "'" & Trim(tx_op_02.Text)
            excel.Workbooks(1).Worksheets(1).cells(14, 4).value = Trim(tx_op_03.Text)

            excel.Visible = True
            excel.Application.Workbooks.Item(1).SaveAs(caminho_app & "\AS_" & Trim(txt_cod_AS.Text) & ".xlsx")
            Me.TopMost = False
            cmb_clientes.Focus()
            Me.TopMost = True

        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmb_codigo_opn_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_codigo_opn.SelectedIndexChanged
        Try
            '
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_opn As SqlDataReader

            cod_edit_OPN = Val(cmb_codigo_opn.Text)
            txt_01.Text = Trim(cmb_codigo_opn.Text)
            'Dim str_staturs As String = "Laboratório"
            Dim SQCMDopn As New SqlCommand("Select * From OPN Where opn_codigo = " & cod_edit_OPN & " And opn_status = " & 3 & "", Con)
            SQDR_opn = SQCMDopn.ExecuteReader(CommandBehavior.Default)
            If SQDR_opn.Read Then

                If Not IsDBNull(SQDR_opn("opn_alterado_por")) Then
                    tx_op_01.Text = Trim(SQDR_opn("opn_alterado_por"))
                Else
                    tx_op_01.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_data_abertura")) Then
                    tx_op_02.Text = Trim(SQDR_opn("opn_data_abertura"))
                Else
                    tx_op_02.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_cliente")) Then
                    tx_op_04.Text = Trim(SQDR_opn("opn_cliente"))
                Else
                    tx_op_04.Text = ""
                End If
                If Len(tx_as_10.Text) <= 0 Then
                    If Not IsDBNull(SQDR_opn("opn_obs")) Then
                        tx_as_10.Text = Trim(SQDR_opn("opn_obs"))
                    Else

                        tx_as_10.Text = ""
                    End If
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''
                If Not IsDBNull(SQDR_opn("opn_contato")) Then
                    tx_as_02.Text = Trim(SQDR_opn("opn_contato"))
                Else
                    tx_as_02.Text = ""
                End If
                'If Len(Trim(SQDR_opn("A1_DDD"))) < 3 Then
                '    str_tel = "0" & Trim(SQDR_opn("A1_DDD")) & Trim(SQDR_opn("A1_TEL"))
                'Else
                '    str_tel = Trim(SQDR_opn("A1_DDD")) & Trim(SQDR_opn("A1_TEL"))
                'End If
                If Not IsDBNull(SQDR_opn("opn_telefone")) Then
                    tx_as_03.Text = Trim(SQDR_opn("opn_telefone"))
                Else
                    tx_as_03.Text = ""
                End If
                'tx_as_03.Text = str_tel

                If Not IsDBNull(SQDR_opn("opn_email")) Then
                    tx_as_04.Text = Trim(SQDR_opn("opn_email"))
                Else
                    tx_as_04.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_endereco")) Then
                    tx_as_05.Text = Trim(SQDR_opn("opn_endereco"))
                Else
                    tx_as_05.Text = ""
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''
                'Descrição do Status do Cliente
                tx_op_03.Text = ""
                Dim SQDR_status As SqlDataReader
                Dim SQCMDStatus As New SqlCommand("Select * From Status Where Codigo = " & Val(SQDR_opn("opn_status")) & "", Con)
                SQDR_status = SQCMDStatus.ExecuteReader(CommandBehavior.Default)
                If SQDR_status.Read Then
                    tx_op_03.Text = Trim(SQDR_status("Descricao"))
                Else
                    tx_op_03.Text = ""
                End If
                SQDR_status.Close()
            Else
                MsgBox("OPN com status inválido!", MsgBoxStyle.Information, "OPN")
            End If
            SQDR_opn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
        End Try
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Sub
    Private Function ValidateActiveDirectoryLogin(ByVal Domain As String, ByVal Username As String, ByVal Password As String) As Boolean
        Dim Success As Boolean = False
        ' Dim Domain As String = "GCNETWORK"

        Dim Entry As New System.DirectoryServices.DirectoryEntry("LDAP://" & Domain, Username, Password)
        Dim Searcher As New System.DirectoryServices.DirectorySearcher(Entry)
        Searcher.SearchScope = DirectoryServices.SearchScope.OneLevel
        Try
            Dim Results As System.DirectoryServices.SearchResult = Searcher.FindOne
            Success = Not (Results Is Nothing)
        Catch
            Success = False
        End Try
        Return Success
    End Function
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            If as_consulta Then

                Me.Cursor = Cursors.WaitCursor

                txt_cod_AS.Focus()
                SendKeys.Send("{ENTER}")
                Timer1.Enabled = False

                Me.Cursor = Cursors.Hand

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txt_cod_AS_TextChanged(sender As Object, e As EventArgs) Handles txt_cod_AS.TextChanged

    End Sub
End Class