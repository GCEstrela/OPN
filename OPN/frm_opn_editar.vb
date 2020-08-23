Imports System.Data.SqlClient

Public Class frm_opn_editar
    Private cod_status_edit As Integer = 0
    Dim _valorOPN
    Private Sub bnt_sair_Click(sender As Object, e As EventArgs) Handles bnt_sair.Click
        Try
            GC.Collect() : GC.WaitForPendingFinalizers()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub frm_opn_editar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            GC.Collect() : GC.WaitForPendingFinalizers()
            Me.Close()
        End If
    End Sub
    Private Sub frm_opn_editar_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_opn As SqlDataReader
            Dim mdata, dia, mes, ano As String
            txt_11.Text = Date.Now.ToString("dd/MM/yyyy")
            benc_Combo(cmb_status, "Status", "Select * From Status Order By Descricao", "Descricao")
            benc_Combo(cmb_tipo_licitacao, "Mod_licitacao", "Select * From Mod_licitacao Order By Descricao", "Descricao")

            'Dim str_staturs As String = "Laboratório"
            Dim SQCMDopn As New SqlCommand("Select * From OPN Where opn_codigo = " & cod_edit_OPN & "", Con)
            SQDR_opn = SQCMDopn.ExecuteReader(CommandBehavior.Default)
            If SQDR_opn.Read Then

            txt_cod_OPN.Text = cod_edit_OPN

            txt_cod_OPN.Focus()
                'txt_codigo.Text = cod_laudo

                If Not IsDBNull(SQDR_opn("opn_data_abertura")) Then
                    'mdata = "0" & Trim(SQDR_opn("opn_data_abertura"))
                    'mes = Mid(mdata, 1, 2)
                    'dia = Mid(mdata, 4, 2)
                    'ano = Mid(mdata, 7, 4)
                    'minha_data = FormatDateTime(CDate(SQDR_opn("opn_data_abertura")), DateFormat.ShortDate)
                    'minha_data = Date.Parse(SQDR_opn("opn_data_abertura"))
                    'txt_01.Text = dia & "/" & mes & "/" & ano     'Trim(SQDR_opn("opn_data_abertura"))
                    txt_01.Text = Trim(SQDR_opn("opn_data_abertura"))
                Else
                    txt_01.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_cliente")) Then
                    txt_02.Text = Trim(SQDR_opn("opn_cliente"))
                Else
                    txt_02.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_pedido_cliente")) Then
                    txt_03.Text = Trim(SQDR_opn("opn_pedido_cliente"))
                Else
                    txt_03.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_telefone")) Then
                    txt_04.Text = Trim(SQDR_opn("opn_telefone")) ' & "-04"
                Else
                    txt_04.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_telefone1")) Then
                    MaskedTextBox1.Text = Trim(SQDR_opn("opn_telefone1")) ' & "-04"
                Else
                    MaskedTextBox1.Text = ""
                End If


                If Not IsDBNull(SQDR_opn("opn_contato")) Then
                    txt_05.Text = Trim(SQDR_opn("opn_contato")) ' & "-05"
                Else
                    txt_05.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_endereco")) Then
                    txt_06.Text = Trim(SQDR_opn("opn_endereco")) ' & "-06"
                Else
                    txt_06.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_email")) Then
                    txt_07.Text = Trim(SQDR_opn("opn_email")) ' & "-07"
                Else
                    txt_07.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_obs")) Then
                    txt_08.Text = Trim(SQDR_opn("opn_obs")) ' & "-08"
                Else
                    txt_08.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_elaborador")) Then
                    txt_09.Text = Trim(SQDR_opn("opn_elaborador")) ' & "-09"
                Else
                    txt_09.Text = ""
                End If
                
                If Not IsDBNull(SQDR_opn("opn_proposta_envio")) Then
                    DateTimePicker2.Text = Trim(SQDR_opn("opn_proposta_envio")) ' & "-09"
                Else
                    DateTimePicker2.Text = ""
                End If

                txt_10.Text = nome_usuario_sistema

                If Not IsDBNull(SQDR_opn("opn_proposta_numero")) Then
                    txt_12.Text = Trim(SQDR_opn("opn_proposta_numero")) ' & "-09"
                Else
                    txt_12.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_data_hora_dispota")) Then
                    txt_data_hora_disputa.Text = Trim(SQDR_opn("opn_data_hora_dispota")) ' & "-09"
                Else
                    txt_data_hora_disputa.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_valor")) Then
                    maskValor_opn.Text = Trim(SQDR_opn("opn_valor")) ' & "-09"
                Else
                    maskValor_opn.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_proposta_data")) Then
                    chk_proposta_enviada.Checked = True
                    DateTimePicker1.Text = Trim(SQDR_opn("opn_proposta_data")) ' & "-09"
                Else
                    chk_proposta_enviada.Checked = False
                    DateTimePicker1.Text = ""
                End If
                If Not IsDBNull(SQDR_opn("opn_licitacao")) Then
                    Dim SQDR_licitacao As SqlDataReader
                    Dim SQCMDLicitacao As New SqlCommand("Select * From Mod_licitacao Where Codigo = " & Val(SQDR_opn("opn_licitacao")) & "", Con)
                    SQDR_licitacao = SQCMDLicitacao.ExecuteReader(CommandBehavior.Default)
                    If SQDR_licitacao.Read Then

                        If Not IsDBNull(SQDR_licitacao("descricao")) Then
                            cmb_tipo_licitacao.Text = Trim(SQDR_licitacao("descricao")) ' & "-09"
                        Else
                            cmb_tipo_licitacao.Text = ""
                        End If

                    End If
                    SQDR_licitacao.Close()

                Else
                    cmb_tipo_licitacao.Text = ""
                End If
                '
                'Descrição do Status do Cliente
                Dim SQDR_status As SqlDataReader
                Dim SQCMDStatus As New SqlCommand("Select * From Status Where Codigo = " & Val(SQDR_opn("opn_status")) & "", Con)
                SQDR_status = SQCMDStatus.ExecuteReader(CommandBehavior.Default)
                If SQDR_status.Read Then

                    cod_status_edit = Val(SQDR_opn("opn_status"))
                    cmb_status.Text = Trim(SQDR_status("Descricao"))

                    Dim _status = Trim(SQDR_status("Descricao"))
                    If _status = "Em Aberto" Then
                        txt_cod_OPN.BackColor = ColorTranslator.FromHtml(em_aberto)     'Color.Yellow
                    ElseIf _status = "Proposta Enviada" Then
                        txt_cod_OPN.BackColor = ColorTranslator.FromHtml(enviada) 'Color.GreenYellow
                    ElseIf _status = "Proposta Aceita" Then
                        txt_cod_OPN.BackColor = ColorTranslator.FromHtml(aceita) 'Color.Green
                    ElseIf _status = "Declinada" Then
                        txt_cod_OPN.BackColor = ColorTranslator.FromHtml(declinada) 'Color.Maroon
                    End If
                Else

                    cmb_status.Text = ""

                End If
                SQDR_status.Close()


                'DG.Columns(0).Frozen = True
                'DG.Columns(1).Frozen = True

                ' '''''''''''''''''''''''''''''''''''''''''''''''''''''
                If usuario_perfil = False Then
                    Button2.Enabled = False
                Else
                    Button2.Enabled = True
                End If

            End If
            SQDR_opn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub bnt_cadastro_Click(sender As Object, e As EventArgs) Handles bnt_cadastro.Click
        Try
            Cursor.Current = Cursors.WaitCursor
            '
            'If Val(cod_status) <= 0 Then
            '    MsgBox("Obs: Campo Status não esta definido!")
            '    Exit Sub
            'End If
            If Val(maskValor_opn.Text) <= 0 Then
                maskValor_opn.Text = 0
                _valorOPN = 0
            End If
            'If IsNumeric(maskValor_opn.Text) <= 0 Then
            '    maskValor_opn.Text = 0
            '    _valorOPN = 0
            'End If

            'If Len(txt_cod_OPN.Text) <= 0 Then Exit Sub

            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_arm As SqlDataReader
            'Replace(CStr(maskValor_opn.Text), ",", ".") & ", ...)"
            'Dim _valorOPN = Replace(CStr(maskValor_opn.Text), ",", ".") & ", ...)"
            'Dim _valorOPN = maskValor_opn.Text
            Dim data_limite, data_proposta As String

            Dim SQCMDPorta As New SqlCommand("Select * From OPN Where opn_codigo = " & Val(txt_cod_OPN.Text) & "", Con)
            SQDR_arm = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
            If SQDR_arm.Read Then
                'Dim dd As String = Trim(DateTimePicker2.Text)
                data_limite = Trim(DateTimePicker2.Text)
                If chk_proposta_enviada.Checked Then

                    data_proposta = Trim(DateTimePicker1.Text)
                    Dim CmdIns_posto As New SqlCommand("Update OPN Set opn_data_abertura='" & Trim(txt_01.Text) & "',opn_cliente='" & Trim(txt_02.Text) & "',opn_pedido_cliente='" & Trim(txt_03.Text) & "',opn_telefone='" & Trim(txt_04.Text) & "',opn_contato='" & Trim(txt_05.Text) & "',opn_endereco='" & Trim(txt_06.Text) & "',opn_email='" & Trim(txt_07.Text) & "',opn_obs='" & Trim(txt_08.Text) & "',opn_elaborador='" & Trim(txt_09.Text) & "',opn_status=" & cod_status & ",opn_alterado_por='" & Trim(txt_10.Text) & "',opn_proposta_envio='" & data_limite & "',opn_proposta_data='" & data_proposta & "',opn_proposta_data_alteracao='" & Trim(txt_11.Text) & "',opn_proposta_numero='" & txt_12.Text & "',opn_alterada_por='" & nome_usuario_sistema & "',opn_data_hora_dispota= '" & Trim(txt_data_hora_disputa.Text) & "',opn_valor=" & _valorOPN & ",opn_licitacao=" & cod_licitacao & " Where opn_codigo= " & Trim(txt_cod_OPN.Text) & " ", Con)
                    CmdIns_posto.ExecuteNonQuery() : CmdIns_posto.Dispose()

                Else

                    data_proposta = ""
                    Dim CmdIns_posto As New SqlCommand("Update OPN Set opn_data_abertura='" & Trim(txt_01.Text) & "',opn_cliente='" & Trim(txt_02.Text) & "',opn_pedido_cliente='" & Trim(txt_03.Text) & "',opn_telefone='" & Trim(txt_04.Text) & "',opn_contato='" & Trim(txt_05.Text) & "',opn_endereco='" & Trim(txt_06.Text) & "',opn_email='" & Trim(txt_07.Text) & "',opn_obs='" & Trim(txt_08.Text) & "',opn_elaborador='" & Trim(txt_09.Text) & "',opn_status=" & cod_status & ",opn_alterado_por='" & Trim(txt_10.Text) & "',opn_proposta_envio='" & data_limite & "',opn_proposta_data='" & data_proposta & "',opn_proposta_data_alteracao='" & Trim(txt_11.Text) & "',opn_proposta_numero='" & txt_12.Text & "',opn_alterada_por='" & nome_usuario_sistema & "',opn_data_hora_dispota= '" & Trim(txt_data_hora_disputa.Text) & "',opn_valor=" & _valorOPN & ",opn_licitacao=" & cod_licitacao & " Where opn_codigo= " & Trim(txt_cod_OPN.Text) & " ", Con)
                    CmdIns_posto.ExecuteNonQuery() : CmdIns_posto.Dispose()

                End If

                MsgBox("Ação executada com êxito!", MsgBoxStyle.Information, titulo_as)
                frm_opn.ativar_filtros_status()

                'ExibeDados(frm_opn.DG_01, "Select * From Lista_OPN Order By Prioridade", "Lista_OPN")
                'frm_principal.chk_f_01.Checked = True : frm_principal.chk_f_02.Checked = False : frm_principal.chk_f_03.Checked = False : frm_principal.chk_f_04.Checked = False
            End If
            SQDR_arm.Close()
            Cursor.Current = Cursors.Default

        Catch ex As Exception
            Cursor.Current = Cursors.Default
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmb_status_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_status.SelectedIndexChanged
        Try
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_max As SqlDataReader
            Dim titulo As String = "Lauto Tecnico"
            ' Dim cod_status = 0

            Dim SQCMDmax As New SqlCommand("Select * From Status Where Descricao = '" & Trim(cmb_status.Text) & "'", Con)
            SQDR_max = SQCMDmax.ExecuteReader(CommandBehavior.Default)
            If SQDR_max.Read Then

                cod_status = SQDR_max("Codigo")
                'If cod_status = 2 Or cod_status = 3 Then
                '    bnt_cadastro.Enabled = False
                '    'chk_proposta_enviada.Font.Bold = True
                'Else
                '    bnt_cadastro.Enabled = True
                'End If

            End If
            SQDR_max.Close()
            'cod_status = 1
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txt_cod_OPN_GotFocus(sender As Object, e As EventArgs) Handles txt_cod_OPN.GotFocus
        txt_03.Focus()
    End Sub

    Private Sub txt_cod_OPN_KeyDown(sender As Object, e As KeyEventArgs) Handles txt_cod_OPN.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Dim Con As SqlConnection = TratadorDeConexao.Conexao()
                Dim SQDR_opn As SqlDataReader

                benc_Combo(cmb_status, "Status", "Select * From Status Order By Descricao", "Descricao")

                Dim str_staturs As String = "Laboratório"
                Dim SQCMDopn As New SqlCommand("Select * From OPN Where opn_codigo = " & Val(txt_cod_OPN.Text) & "", Con)
                SQDR_opn = SQCMDopn.ExecuteReader(CommandBehavior.Default)
                If SQDR_opn.Read Then

                    'txt_cod_OPN.Text = cod_edit_OPN

                    txt_cod_OPN.Focus()
                    'txt_codigo.Text = cod_laudo
                    If Not IsDBNull(SQDR_opn("opn_data_abertura")) Then
                        txt_01.Text = Trim(SQDR_opn("opn_data_abertura"))
                    Else
                        txt_01.Text = ""
                    End If
                    If Not IsDBNull(SQDR_opn("opn_cliente")) Then
                        txt_02.Text = Trim(SQDR_opn("opn_cliente"))
                    Else
                        txt_02.Text = ""
                    End If
                    If Not IsDBNull(SQDR_opn("opn_pedido_cliente")) Then
                        txt_03.Text = Trim(SQDR_opn("opn_pedido_cliente"))
                    Else
                        txt_03.Text = ""
                    End If
                    If Not IsDBNull(SQDR_opn("opn_telefone")) Then
                        txt_04.Text = Trim(SQDR_opn("opn_telefone")) ' & "-04"
                    Else
                        txt_04.Text = ""
                    End If
                    If Not IsDBNull(SQDR_opn("opn_contato")) Then
                        txt_05.Text = Trim(SQDR_opn("opn_contato")) ' & "-05"
                    Else
                        txt_05.Text = ""
                    End If
                    If Not IsDBNull(SQDR_opn("opn_endereco")) Then
                        txt_06.Text = Trim(SQDR_opn("opn_endereco")) ' & "-06"
                    Else
                        txt_06.Text = ""
                    End If
                    If Not IsDBNull(SQDR_opn("opn_email")) Then
                        txt_07.Text = Trim(SQDR_opn("opn_email")) ' & "-07"
                    Else
                        txt_07.Text = ""
                    End If
                    If Not IsDBNull(SQDR_opn("opn_obs")) Then
                        txt_08.Text = Trim(SQDR_opn("opn_obs")) ' & "-08"
                    Else
                        txt_08.Text = ""
                    End If
                    If Not IsDBNull(SQDR_opn("opn_elaborador")) Then
                        txt_09.Text = Trim(SQDR_opn("opn_elaborador")) ' & "-09"
                    Else
                        txt_09.Text = ""
                    End If

                    'Descrição do Status do Cliente
                    Dim SQDR_status As SqlDataReader
                    Dim SQCMDStatus As New SqlCommand("Select * From Status Where Codigo = " & Val(SQDR_opn("opn_status")) & "", Con)
                    SQDR_status = SQCMDStatus.ExecuteReader(CommandBehavior.Default)
                    If SQDR_status.Read Then

                        cod_status_edit = Val(SQDR_opn("opn_status"))
                        cmb_status.Text = Trim(SQDR_status("Descricao"))

                    Else

                        cmb_status.Text = ""

                    End If
                    SQDR_status.Close()


                    'DG.Columns(0).Frozen = True
                    'DG.Columns(1).Frozen = True



                    ' '''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'TextBox1_nserie.Focus()

                    txt_11.Text = Date.Now
                End If
                SQDR_opn.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub chk_proposta_enviada_CheckedChanged(sender As Object, e As EventArgs) Handles chk_proposta_enviada.CheckedChanged
        Try
            If chk_proposta_enviada.Checked Then
                DateTimePicker1.Enabled = True
                txt_12.Enabled = True
                bnt_cadastro.Enabled = True
            Else
                DateTimePicker1.Enabled = False
                txt_12.Enabled = False
                bnt_cadastro.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txt_cod_OPN_TextChanged(sender As Object, e As EventArgs) Handles txt_cod_OPN.TextChanged

    End Sub

    Private Sub cmb_tipo_licitacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_tipo_licitacao.SelectedIndexChanged
        Try
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_max As SqlDataReader
            Dim titulo As String = "Lauto Tecnico"
            'cod_status = 0

            Dim SQCMDmax As New SqlCommand("Select * From Mod_licitacao Where Descricao = '" & Trim(cmb_tipo_licitacao.Text) & "'", Con)
            SQDR_max = SQCMDmax.ExecuteReader(CommandBehavior.Default)
            If SQDR_max.Read Then

                cod_licitacao = SQDR_max("Codigo")

            End If
            SQDR_max.Close()
            'cod_status = 1
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try

            If usuario_perfil = False Then Exit Sub

            Dim Con As SqlConnection = TratadorDeConexao.Conexao()

            Dim CmdDel_obs_opn As New SqlCommand("Delete From OPN Where opn_codigo = " & Trim(txt_cod_OPN.Text) & " ", Con)
            CmdDel_obs_opn.ExecuteNonQuery() : CmdDel_obs_opn.Dispose()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub maskValor_opn_MaskInputRejected(sender As Object, e As MaskInputRejectedEventArgs) Handles maskValor_opn.MaskInputRejected

    End Sub

    Private Sub maskValor_opn_KeyPress(sender As Object, e As KeyPressEventArgs) Handles maskValor_opn.KeyPress

        If e.KeyChar = "." Then
            e.KeyChar = ","
        ElseIf Not Char.IsNumber(e.KeyChar) And Not e.KeyChar = vbBack And Not e.KeyChar = "," Then
            e.Handled = True
        End If

    End Sub

    Private Sub maskValor_opn_Leave(sender As Object, e As EventArgs) Handles maskValor_opn.Leave
        _valorOPN = Replace(maskValor_opn.Text, ",", ".")
        maskValor_opn.Text = FormatNumber(maskValor_opn.Text, 2, , TriState.True)

    End Sub
End Class