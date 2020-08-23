Imports System.Data.SqlClient
Imports System.DirectoryServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module Funcoes

    Public da As SqlDataAdapter
    Public ds As DataSet
    Public da_as As SqlDataAdapter
    Public ds_as As DataSet
    Public Function ValidateActiveDirectoryLogin(ByVal Domain As String, ByVal Username As String, ByVal Password As String) As Boolean
        Try
            Dim Success As Boolean = False

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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
        End Try
    End Function
    Public Sub fechar_app(ByVal nm_app As String)
        Try
            Dim processos As Process
            For Each processos In Process.GetProcesses
                If UCase(processos.ProcessName) = UCase(nm_app) Then
                    processos.Kill()
                End If
            Next
            'For Each processos In Process.GetProcesses
            '    If UCase(processos.ProcessName) = "LAUDO" Then
            '        processos.Kill()
            '    End If
            'Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
        End Try
    End Sub
    Public Function ExibeDados(ByRef idg As Object, ByVal selectCommand As String, ByVal nm_tabela_origem As String) As Boolean
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()

            da = New SqlDataAdapter(selectCommand, Con)

            ds = New DataSet
            da.Fill(ds, nm_tabela_origem)
            idg.DataSource = ds.Tables(nm_tabela_origem)
            If idg.Name = "DG_01" Then

                idg.Columns(0).Frozen = True
                idg.Columns(0).Width = 50

                idg.Columns(1).Frozen = True
                idg.Columns(1).Width = 93
                'idg.Columns(1).ValueType = GetType(Date)
                'idg.Columns(1).name = "Data"
                'idg.Columns(1).DefaultCellStyle.Format = "dd.MM.yyyy"
                'idg.Columns(1).DefaultCellStyle.Format = "d"
                'idg.Columns(1).DefaultCellStyle.Format = "MM-dd-yyyy"
                'idg.Columns(1).ValueType = GetType(Date)

                idg.Columns(2).Width = 100
                idg.Columns(2).name = "D_limite"
                'idg.Columns(2).ValueType = GetType(Date)
                'idg.Columns(2).DefaultCellStyle.Format = "dd/MM/yyyy"


                idg.Columns(3).Width = 100
                'idg.Columns(3).DefaultCellStyle.Format = "dd.MM.yyyy HH:mm:ss"

                idg.Columns(4).Width = 100
                idg.Columns(5).Width = 150
                idg.Columns(6).Width = 230
                idg.Columns(7).Width = 230
                idg.Columns(8).Width = 100
                idg.Columns(9).Width = 100
                idg.Columns(10).Width = 100
                idg.Columns(11).Width = 100
                idg.Columns(11).DefaultCellStyle.Format = "c"
                idg.Columns(12).Visible = False
                idg.Columns(13).Visible = False
                '''''''''''''''''''''''''''''''''''''''''''''''''
                Dim str_status As String
                'Dim SQDR_status_opn As SqlDataReader
                For i = 0 To idg.NewRowIndex
                    Try
                        str_status = Trim(idg.Rows(i).Cells(10).Value)
                        If str_status = "Em Aberto" Then
                            idg.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(em_aberto)    'Color.Yellow
                        ElseIf str_status = "Proposta Enviada" Then
                            idg.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(enviada)    'Color.GreenYellow
                        ElseIf str_status = "Proposta Aceita" Then
                            idg.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(aceita)   'Color.Green
                        ElseIf str_status = "Declinada" Then
                            idg.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(declinada)   'Color.Maroon
                            idg.Rows(i).DefaultCellStyle.Forecolor = ColorTranslator.FromHtml("#FFFFFF")
                        ElseIf str_status = "Cancelada/Revogada" Then
                            idg.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(cancelada_Revogada)
                        ElseIf str_status = "Suspensa" Then
                            idg.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(Suspensa)
                        End If

                    Catch ex As Exception
                        ' MsgBox(ex.Message, MsgBoxStyle.Critical, "")
                    End Try
                Next

                contar_Status(frm_opn.ToolStripTextBox5, "Em Aberto")
                contar_Status(frm_opn.ToolStripTextBox6, "Proposta Enviada")
                contar_Status(frm_opn.ToolStripTextBox7, "Proposta Aceita")
                contar_Status(frm_opn.ToolStripTextBox8, "Declinada")
                contar_Status(frm_opn.ToolStripTextBox10, "Cancelada/Revogada")
                contar_Status(frm_opn.ToolStripTextBox11, "Suspensa")

            ElseIf idg.Name = "DG_02" Then

                idg.Columns(0).Frozen = True
                idg.Columns(0).Width = 50
                idg.Columns(1).Frozen = True
                idg.Columns(1).Width = 93
                idg.Columns(2).Width = 100
                idg.Columns(2).name = "D_limite"
                idg.Columns(3).Width = 100

                idg.Columns(4).Width = 100
                idg.Columns(5).Width = 150
                idg.Columns(6).Width = 230
                idg.Columns(7).Width = 230
                idg.Columns(8).Width = 100
                idg.Columns(9).Width = 100
                idg.Columns(10).Width = 100
                idg.Columns(11).Visible = False
                idg.Columns(12).Visible = False
                idg.Columns(13).Visible = False
                '''''''''''''''''''''''''
                Dim str_status As String
                For i = 0 To idg.NewRowIndex
                    Try
                        str_status = Trim(idg.Rows(i).Cells(10).Value)
                        If str_status = "À Executar" Then
                            idg.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(aexecutar)    'Color.Yellow
                        ElseIf str_status = "Em Execução/Pendente" Then
                            idg.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(execucao_pendente)    'Color.Cyan
                        ElseIf str_status = "Em Execução" Then
                            idg.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(em_execucao)    ' Color.GreenYellow
                        ElseIf str_status = "Executada" Then
                            idg.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(executada)    'Color.Green
                        ElseIf str_status = "Cancelada" Then
                            idg.Rows(i).DefaultCellStyle.BackColor = ColorTranslator.FromHtml(cancelada)    'Color.Maroon
                            idg.Rows(i).DefaultCellStyle.Forecolor = ColorTranslator.FromHtml("#FFFFFF")
                        End If

                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical, "")
                    End Try
                Next

                contar_Status_as(frm_opn.ToolStripTextBox1, "À Executar")
                contar_Status_as(frm_opn.ToolStripTextBox2, "Em Execução/Pendente")
                contar_Status_as(frm_opn.ToolStripTextBox3, "Em Execução")
                contar_Status_as(frm_opn.ToolStripTextBox4, "Executada ")
            End If

            'For i As Integer = 0 To idg.Rows.Count - 1
            '    idg.Rows(i).Cells(1).Value = DateTime.Parse(idg.Rows(i).Cells(1).Value).ToString("dd-MM-yyyy")
            '    idg.Rows(i).Cells(2).Value = DateTime.Parse(idg.Rows(i).Cells(2).Value).ToString("dd-MM-yyyy")
            'Next

            'frm_principal.ToolStripTextBox1.BackColor = Color.Yellow
            'frm_principal.ToolStripTextBox2.BackColor = Color.Cyan
            'frm_principal.ToolStripTextBox3.BackColor = Color.IndianRed
            'frm_principal.ToolStripTextBox4.BackColor = Color.GreenYellow

            'contar_Status(frm_principal.ToolStripTextBox1, "Em Aberto")
            'contar_Status(frm_principal.ToolStripTextBox2, "Proposta Enviada")
            'contar_Status(frm_principal.ToolStripTextBox3, "Proposta Aceita")
            'contar_Status(frm_principal.ToolStripTextBox4, "Declinada ")

            Cursor.Current = Cursors.Default
            Return True
        Catch ex As Exception
            Return False
            Cursor.Current = Cursors.Default
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
        End Try


    End Function
    Public Function ExibeDadosAS(ByRef idg_as As Object, ByVal selectCommand As String, ByVal nm_tabela_origem As String) As Boolean
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()

            da_as = New SqlDataAdapter(selectCommand, Con)

            ds_as = New DataSet
            da_as.Fill(ds_as, nm_tabela_origem)
            idg_as.DataSource = ds_as.Tables(nm_tabela_origem)
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            idg_as.EnableHeadersVisualStyles = False
            idg_as.ColumnHeadersDefaultCellStyle.BackColor = Color.Gainsboro   'LightSteelBlue
            idg_as.RowHeadersDefaultCellStyle.BackColor = Color.Gainsboro
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            idg_as.Columns(0).Frozen = True
            idg_as.Columns(0).Width = 50
            idg_as.Columns(1).Frozen = True
            idg_as.Columns(1).Width = 93
            idg_as.Columns(2).Width = 100
            idg_as.Columns(2).name = "D_limite"
            idg_as.Columns(3).Width = 100

            idg_as.Columns(4).Width = 100
            idg_as.Columns(5).Width = 150
            idg_as.Columns(6).Width = 230
            idg_as.Columns(7).Width = 230
            idg_as.Columns(8).Width = 100
            idg_as.Columns(9).Width = 100
            idg_as.Columns(10).Width = 100
            '''''''''''''''''''''''''
            Dim str_status As String
            For i = 0 To idg_as.NewRowIndex
                Try
                    str_status = Trim(idg_as.Rows(i).Cells(10).Value)
                    If str_status = "À Executar" Then
                        idg_as.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                    ElseIf str_status = "Em Execução/Pendente" Then
                        idg_as.Rows(i).DefaultCellStyle.BackColor = Color.Cyan
                    ElseIf str_status = "Em Execução" Then
                        idg_as.Rows(i).DefaultCellStyle.BackColor = Color.IndianRed
                    ElseIf str_status = "Executada" Then
                        idg_as.Rows(i).DefaultCellStyle.BackColor = Color.GreenYellow
                    ElseIf str_status = "Cancelada" Then
                        idg_as.Rows(i).DefaultCellStyle.BackColor = Color.Maroon
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "")
                End Try
            Next
            idg_as = Nothing
            'frm_principal.ToolStripTextBox1.BackColor = Color.Yellow
            'frm_principal.ToolStripTextBox2.BackColor = Color.Cyan
            'frm_principal.ToolStripTextBox3.BackColor = Color.IndianRed
            'frm_principal.ToolStripTextBox4.BackColor = Color.GreenYellow

            'contar_Status(frm_principal.ToolStripTextBox1, "Em Aberto")
            'contar_Status(frm_principal.ToolStripTextBox2, "Proposta Enviada")
            'contar_Status(frm_principal.ToolStripTextBox3, "Proposta Aceita")
            'contar_Status(frm_principal.ToolStripTextBox4, "Declinada ")

            Cursor.Current = Cursors.Default
            idg_as = Nothing : da_as = Nothing : ds_as = Nothing
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
            Return False
            Cursor.Current = Cursors.Default

        End Try
    End Function
    Public Sub contar_Status(ByRef tx_campo As ToolStripTextBox, ByVal str_status_equipamento As String)
        Dim Con As SqlConnection = TratadorDeConexao.Conexao()
        Dim SQDR_posto As SqlDataReader

        Dim str_staturs As String = "Laboratório"
        Dim SQCMDPorta As New SqlCommand("Select COUNT(Status) as cont_Status From Lista_OPN Where Status = '" & str_status_equipamento & "'", Con)
        SQDR_posto = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
        If SQDR_posto.Read Then
            tx_campo.Text = SQDR_posto("cont_Status")
        Else
            tx_campo.Text = "0"
        End If
    End Sub
    Public Sub contar_Status_as(ByRef tx_campo As ToolStripTextBox, ByVal str_status_equipamento As String)
        Dim Con As SqlConnection = TratadorDeConexao.Conexao()
        Dim SQDR_posto As SqlDataReader

        'Dim str_staturs As String = "Laboratório"
        Dim SQCMDPorta As New SqlCommand("Select COUNT(Status) as cont_Status_as From Lista_AS Where Status = '" & str_status_equipamento & "'", Con)
        SQDR_posto = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
        If SQDR_posto.Read Then
            tx_campo.Text = SQDR_posto("cont_Status_as")
        Else
            tx_campo.Text = "0"
        End If
    End Sub
    Public Sub benc_Combo(ByRef ob_combp As Object, ByVal nm_Tabela As String, ByVal str_Selete As String, ByVal nm_Campo As String)
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim Con As SqlConnection = TratadorDeConexao.Conexao()
            Dim SQDR_ As SqlDataReader

            ob_combp.Items.Clear()
            Dim SQCMDPorta As New SqlCommand(str_Selete, Con)
            SQDR_ = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
            Do While SQDR_.Read

                ob_combp.Items.Add(Trim(SQDR_(nm_Campo)))

            Loop
            SQDR_.Close()
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            'MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
        End Try
    End Sub
    Public Sub benc_Combo2(ByRef ob_combp As Object, ByVal nm_Tabela As String, ByVal str_Selete As String, ByVal nm_Campo As String, Optional campo_recno As String = "")
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim Con2 As SqlConnection = TratadorDeConexao.Conexao2()
            Dim SQDR_ As SqlDataReader

            'ob_combp.Items.Clear()
            Dim SQCMDPorta As New SqlCommand(str_Selete, Con2)
            SQDR_ = SQCMDPorta.ExecuteReader(CommandBehavior.Default)
            Do While SQDR_.Read
                If nm_Tabela <> "Exibicao_CCusto" Then
                    ob_combp.Items.Add(Trim(SQDR_(nm_Campo)))
                Else
                    ob_combp.Items.Add(Trim(SQDR_(nm_Campo)) & " * " & SQDR_(campo_recno))
                End If
            Loop
            SQDR_.Close()
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
        End Try
    End Sub
    Public Sub txt_maiusculo(ByRef tx As TextBox, ByVal kchar As String)

        If Char.IsLower(kchar) Then
            tx.SelectedText = Char.ToUpper(kchar)

        End If

    End Sub
    Public Sub ExportToExcel(ByVal dgvName As DataGridView, ByVal [option] As XlSortOn, Optional ByVal fileName As String = "")

        Dim objExcelApp As New Excel.Application()
        Dim objExcelBook As Excel.Workbook
        Dim objExcelSheet As Excel.Worksheet

        Try


            objExcelBook = objExcelApp.Workbooks.Add
            objExcelSheet = CType(objExcelBook.Worksheets(1), Excel.Worksheet)
            objExcelApp.Visible = True
            ' Ciclo nos cabeçalhos para escrever os títulos a bold/negrito
            Dim dgvColumnIndex As Int16 = 1
            For Each col As DataGridViewColumn In dgvName.Columns
                objExcelSheet.Cells(1, dgvColumnIndex) = col.HeaderText
                objExcelSheet.Cells(1, dgvColumnIndex).Font.Bold = True
                dgvColumnIndex += 1
                'If dgvColumnIndex = 17 Then Exit For
            Next
            ' Ciclo nas linhas/células
            Dim dgvRowIndex As Integer = 2
            Dim str_defeitor_relatados As String
            'Dim codigo_atendimento_relato As String
            For Each row As DataGridViewRow In dgvName.Rows
                'For i = 0 To dgvName.VisableRows.Count()
                Dim dgvCellIndex As Integer = 1
                str_defeitor_relatados = ""
                If row.Visible = True Then
                    For Each cell As DataGridViewCell In row.Cells
                        objExcelSheet.Cells(dgvRowIndex, dgvCellIndex) = "'" & cell.Value
                        If dgvCellIndex = 16 Then
                            objExcelSheet.Cells(dgvRowIndex, dgvCellIndex) = "'" & str_defeitor_relatados
                        End If
                        dgvCellIndex += 1
                        'If dgvCellIndex = 17 Then Exit For
                    Next
                    dgvRowIndex += 1
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''''
                'If dgvRowIndex = 17 Then Exit For
            Next

            ' Ajusta o largura das colunas automaticamente
            objExcelSheet.Columns.AutoFit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "")
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

End Module
