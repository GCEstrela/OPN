Imports System.Data.SqlClient
Imports System.Security.Principal
Imports Microsoft.VisualBasic.ApplicationServices
Imports System.Text
Imports System.Collections

Public Class frm_opn_cadastro
    Private WM_NCHITTEST As Integer = &H84
    Private HTCLIENT As Integer = &H1
    Private HTCAPTION As Integer = &H2
    Protected Overrides Sub WndProc(ByRef m As Message)
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
    'Private cod_status As Integer = 0
    Private Sub bnt_sair_Click(sender As Object, e As EventArgs) Handles bnt_sair.Click
        Try
            GC.Collect() : GC.WaitForPendingFinalizers()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub bnt_cadastro_Click(sender As Object, e As EventArgs) Handles bnt_cadastro.Click
        Try
            Cursor.Current = Cursors.WaitCursor

            If Len(Trim(txt_02.Text)) <= 0 Then
                MsgBox("Campo Nome do Cliente é obrigatório!", MsgBoxStyle.Critical, titulo_opn)
                txt_02.Focus() : Exit Sub
            End If

            If cod_status <= 0 Then
                MsgBox("O campo Status é obrigatório!", MsgBoxStyle.Information, "Oportunidades")
                cmbStatus.Focus() : Exit Sub
            End If


            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_arm As SqlDataReader
            Dim SQDR_max As SqlDataReader
            Dim titulo As String = "Lauto Tecnico"

            Dim SQCMDPorta As New SqlCommand("Select * From OPN Where opn_codigo = " & Val(txt_cod_OPN.Text) & "", Con)
            SQDR_arm = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
            If SQDR_arm.Read Then

                Dim CmdIns_posto As New SqlCommand("Update OPN Set opn_data_abertura='" & Trim(txt_01.Text) & "',opn_cliente='" & Trim(txt_02.Text) & "',opn_pedido_cliente='" & Trim(txt_03.Text) & "',opn_telefone='" & Trim(mask_txt_04.Text) & "',opn_contato='" & Trim(txt_05.Text) & "',opn_endereco='" & Trim(txt_06.Text) & "',opn_email='" & Trim(txt_07.Text) & "',opn_obs='" & Trim(txt_08.Text) & "',opn_elaborador='" & Trim(txt_09.Text) & "',opn_status=" & cod_status & ",opn_proposta_envio='" & Trim(DateTimePicker1.Text) & "',opn_licitacao=" & cod_licitacao & ",opn_valor=" & maskValor_opn.Text & ",opn_telefone1='" & Trim(MaskedTextBox1.Text) & "' Where opn_codigo= " & Trim(txt_cod_OPN.Text) & " ", Con)
                CmdIns_posto.ExecuteNonQuery() : CmdIns_posto.Dispose()

                'ExibeDados(frm_principal.DG, "Select * From Lista_OPN Order By Prioridade", "Lista_OPN")
                'frm_principal.chk_f_01.Checked = True : frm_principal.chk_f_02.Checked = False : frm_principal.chk_f_03.Checked = False : frm_principal.chk_f_04.Checked = False
                ' MsgBox("Registro alterado com êxito!", MsgBoxStyle.Information, titulo)

            Else
                Try

                    Dim CmdIns_swit As New SqlCommand("Insert into OPN(opn_data_abertura,opn_cliente,opn_pedido_cliente,opn_telefone,opn_contato,opn_endereco,opn_email,opn_obs,opn_elaborador,opn_status,opn_proposta_envio,opn_licitacao,opn_valor,opn_telefone1) values ('" & Trim(txt_01.Text) & "','" & Trim(txt_02.Text) & "','" & Trim(txt_03.Text) & "','" & Trim(mask_txt_04.Text) & "','" & Trim(txt_05.Text) & "','" & Trim(txt_06.Text) & "','" & Trim(txt_07.Text) & "','" & Trim(txt_08.Text) & "','" & nome_usuario_sistema & "'," & cod_status & ",'" & DateTimePicker1.Text & "'," & cod_licitacao & "," & maskValor_opn.Text & ",'" & Trim(MaskedTextBox1.Text) & "')", Con)
                    CmdIns_swit.ExecuteNonQuery() : CmdIns_swit.Dispose()

                    'Dim CmdIns_swit As New SqlCommand("Insert into OPN(opn_data_abertura,opn_cliente,opn_pedido_cliente,opn_telefone,opn_contato,opn_endereco,opn_email,opn_obs,opn_elaborador,opn_status,opn_proposta_envio,opn_codigo,opn_licitacao) values ('" & Trim(txt_01.Text) & "','" & Trim(txt_02.Text) & "','" & Trim(txt_03.Text) & "','" & Trim(mask_txt_04.Text) & "','" & Trim(txt_05.Text) & "','" & Trim(txt_06.Text) & "','" & Trim(txt_07.Text) & "','" & Trim(txt_08.Text) & "','" & nome_usuario_sistema & "'," & cod_status & ",'" & DateTimePicker1.Text & "'," & Val(txt_cod_OPN.Text) & "," & cod_licitacao & ")", Con)
                    'CmdIns_swit.ExecuteNonQuery() : CmdIns_swit.Dispose()

                    txt_01.Text = ""
                    txt_02.Text = ""
                    txt_03.Text = ""
                    mask_txt_04.Text = ""
                    txt_05.Text = ""
                    txt_06.Text = ""
                    txt_08.Text = ""
                    txt_09.Text = ""
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'ExibeDados(frm_principal.DG, "Select * From Lista_OPN Order By Prioridade", "Lista_OPN")
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Catch ex As Exception
                    MsgBox(ex.Message)
                    SQDR_arm.Close()
                    Exit Sub
                End Try

                Dim SQCMDEqui As New SqlCommand("Select max(opn_codigo) as cod_max From OPN", Con)
                SQDR_max = SQCMDEqui.ExecuteReader(CommandBehavior.Default)
                If SQDR_max.Read Then
                    txt_cod_OPN.Text = SQDR_max("cod_max") + 1
                    txt_01.Text = Date.Now.ToString("dd/MM/yyyy")
                End If
                SQDR_max.Close()

            End If
            SQDR_arm.Close()

            frm_opn.ativar_filtros_status()

            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub frm_opn_cadastro_DoubleClick(sender As Object, e As EventArgs) Handles Me.DoubleClick

    End Sub

    Private Sub frm_opn_cadastro_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            GC.Collect() : GC.WaitForPendingFinalizers()
            Me.Close()
        End If
    End Sub

    Private Sub frm_opn_cadastro_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_max As SqlDataReader
            Dim titulo As String = "Lauto Tecnico"




            txt_cod_OPN.BackColor = ColorTranslator.FromHtml(em_aberto)
            cod_status = 1
            Dim SQCMDmax As New SqlCommand("Select max(opn_codigo) as cod_max From OPN", Con)
            SQDR_max = SQCMDmax.ExecuteReader(CommandBehavior.Default)
            If SQDR_max.Read Then
                If Not IsDBNull(SQDR_max("cod_max")) Then
                    txt_cod_OPN.Text = SQDR_max("cod_max") + 1
                Else
                    txt_cod_OPN.Text = "1000"
                End If
            Else
                txt_cod_OPN.Text = "1000"
            End If
            SQDR_max.Close()

            benc_Combo(cmbStatus, "Status", "Select * From Status Order By Descricao", "Descricao")
            benc_Combo(cmb_tipo_licitacao, "Mod_licitacao", "Select * From Mod_licitacao Order By Descricao", "Descricao")

            cmbStatus.Text = "Em Aberto" : cod_status = 1

            txt_01.Text = Date.Now.ToString("dd/MM/yyyy")
            txt_09.Text = nome_usuario_sistema

            '''''''''''''''''''''''''''''''''''
            Dim agora As DateTime = DateTime.Now
            'tx_01.Text = agora
            Dim dias As Integer = 7
            agora = agora.AddDays(dias)
            DateTimePicker1.Text = agora
            '''''''''''''''''''''''''''
            If usuario_perfil = False Then
                Button2.Enabled = False
            Else
                Button2.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmbStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbStatus.SelectedIndexChanged
        Try
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_max As SqlDataReader
            Dim titulo As String = "Lauto Tecnico"
            'cod_status = 0

            Dim SQCMDmax As New SqlCommand("Select * From Status Where Descricao = '" & Trim(cmbStatus.Text) & "'", Con)
            SQDR_max = SQCMDmax.ExecuteReader(CommandBehavior.Default)
            If SQDR_max.Read Then

                cod_status = SQDR_max("Codigo")

            End If
            SQDR_max.Close()
            'cod_status = 1
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txt_02_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_02.KeyPress
        Try

            If Char.IsLower(e.KeyChar) Then
                txt_02.SelectedText = Char.ToUpper(e.KeyChar)
                e.Handled = True
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub txt_02_LostFocus(sender As Object, e As EventArgs) Handles txt_02.LostFocus
        Try
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_max As SqlDataReader
            Dim titulo As String = "Lauto Tecnico"
            Dim nm_cliente As String = Trim(UCase(txt_02.Text))
            ' cod_status = 0

            Dim SQCMDmax As New SqlCommand("Select * From OPN Where opn_cliente = '" & nm_cliente & "'", Con)
            SQDR_max = SQCMDmax.ExecuteReader(CommandBehavior.Default)
            If SQDR_max.Read Then

                txt_02.BackColor = Color.Blue

            Else

                txt_02.BackColor = Color.White

            End If
            SQDR_max.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub txt_03_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_03.KeyPress
        Try
            If Char.IsLower(e.KeyChar) Then
                txt_03.SelectedText = Char.ToUpper(e.KeyChar)
                e.Handled = True
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub txt_05_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_05.KeyPress
        Try

            If Char.IsLower(e.KeyChar) Then
                txt_05.SelectedText = Char.ToUpper(e.KeyChar)
                e.Handled = True
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub txt_06_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_06.KeyPress
        Try

            If Char.IsLower(e.KeyChar) Then
                txt_06.SelectedText = Char.ToUpper(e.KeyChar)
                e.Handled = True
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub txt_07_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_07.KeyPress
        Try

            If Char.IsLower(e.KeyChar) Then
                txt_07.SelectedText = Char.ToUpper(e.KeyChar)
                e.Handled = True
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub txt_09_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_09.KeyPress
        Try
            If Char.IsLower(e.KeyChar) Then
                txt_09.SelectedText = Char.ToUpper(e.KeyChar)
                e.Handled = True
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub txt_08_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_08.KeyPress
        Try

            If Char.IsLower(e.KeyChar) Then
                txt_08.SelectedText = Char.ToUpper(e.KeyChar)
                e.Handled = True
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub txt_cod_OPN_GotFocus(sender As Object, e As EventArgs) Handles txt_cod_OPN.GotFocus
        txt_02.Focus()
    End Sub

    Private Sub txt_cod_OPN_TextChanged(sender As Object, e As EventArgs) Handles txt_cod_OPN.TextChanged

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Try
            'DateTimePicker1.Format = DateTimePickerFormat.Short
            'MsgBox(DateTimePicker1.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub frm_opn_cadastro_Resize(sender As Object, e As EventArgs) Handles Me.Resize

    End Sub

    Private Sub frm_opn_cadastro_ResizeBegin(sender As Object, e As EventArgs) Handles Me.ResizeBegin

    End Sub

    Private Sub frm_opn_cadastro_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd

    End Sub

    Private Sub cmb_tipo_licitacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_tipo_licitacao.SelectedIndexChanged
        Try
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_max As SqlDataReader
            Dim titulo As String = "Lauto Tecnico"
            'cod_status = 0

            Dim SQCMDmax As New SqlCommand("Select * From Status Where Descricao = '" & Trim(cmb_tipo_licitacao.Text) & "'", Con)
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

            'If usuario_perfil = False Then Exit Sub

            'Dim Con As SqlConnection = TratadorDeConexao.Conexao()

            'Dim CmdDel_obs_opn As New SqlCommand("Delete From OPN Where opn_codigo = " & Trim(txt_cod_OPN.Text) & " ", Con)
            'CmdDel_obs_opn.ExecuteNonQuery() : CmdDel_obs_opn.Dispose()

            '''''''''''''''
            If usuario_perfil = False Then Exit Sub

            If cod_edit_OPN <= 0 Then Exit Sub
            Beep()
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim Resposta As Integer
            Resposta = MessageBox.Show("Esta ação apagará a OPN selecionada e totas as AS relacionadas. Confirma essa ação", titulo_opn, MessageBoxButtons.YesNo)
            If Resposta = 6 Then

                Dim CmdDel_obs_opn As New SqlCommand("Delete From OPN Where opn_codigo= " & Trim(txt_cod_OPN.Text) & " ", Con)
                CmdDel_obs_opn.ExecuteNonQuery() : CmdDel_obs_opn.Dispose()

                MsgBox("Ação concluida com êxito!", MsgBoxStyle.Information, titulo_opn)
            Else

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub maskValor_opn_MaskInputRejected(sender As Object, e As MaskInputRejectedEventArgs) Handles maskValor_opn.MaskInputRejected

    End Sub

    Private Sub txt_01_TextChanged(sender As Object, e As EventArgs) Handles txt_01.TextChanged

    End Sub
    'Sub Masked_Key_Press(MskEdit As MaskEdBox, keyascii As Integer)

    '    ' Calculate the location of the decimal point in the mask:
    '    mask_cents_pos = InStr(MskEdit.Mask, ".")
    '    mask_dollars = mask_cents_pos - 1

    '    ' Check for period keypress:
    '    If keyascii = 46 And MskEdit.SelStart < 6 Then
    '        tlen = MskEdit.SelStart + 1       ' Store current location.
    '        MskEdit.SelStart = 0
    '        MskEdit.SelLength = tlen          ' Highlight up to the current
    '        tempo = MskEdit.SelText           ' position & save selected text.
    '        MskEdit.SelLength = mask_dollars
    '        MskEdit.SelText = ""              ' Clear to the left of decimal.
    '        MskEdit.SelStart = mask_cents_pos - tlen
    '        MskEdit.SelLength = tlen           ' Reposition caret
    '        MskEdit.SelText = tempo            ' and paste copied data.
    '        MskEdit.SelStart = mask_cents_pos  ' Position caret after cents.
    '    End If
    'End Sub

End Class