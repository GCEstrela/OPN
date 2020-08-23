Imports System.Security
Imports System.Security.Principal.WindowsIdentity
Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory
Imports System.Text
Imports System.Data.SqlClient
Imports System.IO

Public Class frm_login
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try

            da = Nothing : ds = Nothing
            fechar_app("Excel")
            For intIndex As Integer = My.Application.OpenForms.Count - 1 To 0 Step -1
                If My.Application.OpenForms.Item(intIndex) IsNot Me Then
                    My.Application.OpenForms.Item(intIndex).Close()
                End If
            Next
            fechar_app("OPN")
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_opn)
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            nome_usuario_sistema = Trim(cmb_usuarios.Text)
            If ValidateActiveDirectoryLogin("GCNETWORK", Trim(cmb_usuarios.Text), Trim(TextBox1.Text)) Then
                Try

                    Dim Con As SqlConnection = TratadorDeConexao.Conexao()

                    Dim SQDR_user As SqlDataReader
                    Dim SQDR_user2 As SqlDataReader

                    Dim SQCMDUser As New SqlCommand("Select * From usuarios Where nome = '" & Trim(cmb_usuarios.Text) & "'", Con)
                    SQDR_user = SQCMDUser.ExecuteReader(CommandBehavior.Default)
                    If SQDR_user.Read Then
                        codigo_usuario_sistema = SQDR_user("codigo")
                        If SQDR_user("Perfil") = 1 Then
                            usuario_perfil = True
                        Else
                            usuario_perfil = False
                        End If
                    Else

                        Dim CmdIns_user As New SqlCommand("Insert into usuarios(nome,perfil) values ('" & Trim(cmb_usuarios.Text) & "'," & 2 & ")", Con)
                        CmdIns_user.ExecuteNonQuery() : CmdIns_user.Dispose()

                        Dim SQCMDUser2 As New SqlCommand("Select * From usuarios Where nome = '" & Trim(cmb_usuarios.Text) & "'", Con)
                        SQDR_user2 = SQCMDUser2.ExecuteReader(CommandBehavior.Default)
                        If SQDR_user2.Read Then
                            codigo_usuario_sistema = SQDR_user2("codigo")
                            If SQDR_user2("Perfil") = 1 Then
                                usuario_perfil = True
                            Else
                                usuario_perfil = False
                            End If
                        Else
                            codigo_usuario_sistema = "0"
                            usuario_perfil = False
                        End If
                        SQDR_user2.Close()

                    End If
                    SQDR_user.Close()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
                End Try
                Me.Hide()
                frm_opn.Show()
            Else
                TextBox1.BackColor = Color.Red
                MsgBox("Usuário e/ou senha inválido(s)..")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
        End Try
    End Sub
    'Private Sub listaActiveDirectoryLogin(ByVal Domain As String, ByVal Username As String, ByVal Password As String) 'As Boolean
    '    Try
    '        Dim Success As Boolean = False

    '        Dim Entry As New System.DirectoryServices.DirectoryEntry("LDAP://" & Domain, Username, Password)
    '        Dim Searcher As New System.DirectoryServices.DirectorySearcher(Entry)
    '        Searcher.SearchScope = DirectoryServices.SearchScope.OneLevel
    '        Try
    '            Dim Results As System.DirectoryServices.SearchResult = Searcher.FindOne
    '            Dim myResultPropColl As ResultPropertyCollection
    '            myResultPropColl = Results.Properties
    '            Dim myKey As String
    '            For Each myKey In myResultPropColl.PropertyNames
    '                cmb_usuarios.Items.Add(myResultPropColl(myKey)(0))
    '            Next
    '            'Success = Not (Results Is Nothing)
    '        Catch
    '            'Success = False
    '        End Try
    '        'Return Success
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
    '    End Try
    'End Sub

    Private Sub frm_login_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            usuario_perfil = False
            'ACESSO AO BANCO TOTVS
            'Dim linhaTexto2 As String = "Data Source=GCTOTVS01;Initial Catalog=tmPRD2;User ID=totvs;Password=totvs;Min Pool Size=5;Max Pool Size=15;Connection Reset=True;Connection Lifetime=600;Trusted_Connection=no;MultipleActiveResultSets=True"
            Dim linhaTexto2 As String = "Data Source=gctotvs02\protheus12;Initial Catalog=P12OFICIAL;User ID=operacoes;Password=operacoes;Min Pool Size=5;Max Pool Size=15;Connection Reset=True;Connection Lifetime=600;Trusted_Connection=no;MultipleActiveResultSets=True"
            TratadorDeConexao.Preparar2(linhaTexto2)
            Dim checkconexao2 = TratadorDeConexao.Abrir2()
            If checkconexao2 = False Then
                MessageBox.Show("Problema na conexão com o banco de dados", "Projetos", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'ACESSO AO BANCO OPN
            'Dim configacao As String = Application.StartupPath & "\sql.txt"
            'If IO.File.Exists(Trim(configacao)) = False Then
            '    MsgBox("Arquivo da base para criação da Folha de Dados não foi encontrado.", MsgBoxStyle.Critical, titulo_opn)
            '    Exit Sub
            'End If
            ''Sempre é a primeira linha do arquivo txt.
            'Dim writer As New StreamReader(configacao)
            'Dim linhaTexto As String = writer.ReadLine()
            Dim linhaTexto As String = "Data Source=GCSERVER;Initial Catalog=OPN;User ID=sa;Password=050382;Min Pool Size=5;Max Pool Size=15;Connection Reset=True;Connection Lifetime=600;Trusted_Connection=no;MultipleActiveResultSets=True"
            'Dim linhaTexto As String = "Data Source=GCTEC04\SQLEXPRESS;Initial Catalog=OPN;User ID=Sam;Password=sammas;Min Pool Size=5;Max Pool Size=15;Connection Reset=True;Connection Lifetime=600;Trusted_Connection=no;MultipleActiveResultSets=True"
            TratadorDeConexao.Preparar(linhaTexto)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim checkconexao = TratadorDeConexao.Abrir()
            If checkconexao = False Then
                MessageBox.Show("Problema na conexão com o banco de dados", "Projetos", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            benc_Combo(cmb_usuarios, "OPN", "Select * From usuarios Order By nome", "nome")


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
        End Try
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        Try
            If TextBox1.BackColor = Color.Red Then
                TextBox1.BackColor = Color.White
            End If

            If e.KeyCode <> Keys.Enter Then Exit Sub
            nome_usuario_sistema = Trim(cmb_usuarios.Text)
            If ValidateActiveDirectoryLogin("GCNETWORK", Trim(cmb_usuarios.Text), Trim(TextBox1.Text)) Then
                Try

                    Dim Con As SqlConnection = TratadorDeConexao.Conexao()

                    Dim SQDR_user As SqlDataReader
                    Dim SQDR_user2 As SqlDataReader

                    Dim SQCMDUser As New SqlCommand("Select * From usuarios Where nome = '" & Trim(cmb_usuarios.Text) & "'", Con)
                    SQDR_user = SQCMDUser.ExecuteReader(CommandBehavior.Default)
                    If SQDR_user.Read Then
                        codigo_usuario_sistema = SQDR_user("codigo")
                        If SQDR_user("Perfil") = 1 Then
                            usuario_perfil = True
                        Else
                            usuario_perfil = False
                        End If
                    Else

                        Dim CmdIns_user As New SqlCommand("Insert into usuarios(nome,perfil) values ('" & Trim(cmb_usuarios.Text) & "'," & 2 & ")", Con)
                        CmdIns_user.ExecuteNonQuery() : CmdIns_user.Dispose()

                        Dim SQCMDUser2 As New SqlCommand("Select * From usuarios Where nome = '" & Trim(cmb_usuarios.Text) & "'", Con)
                        SQDR_user2 = SQCMDUser2.ExecuteReader(CommandBehavior.Default)
                        If SQDR_user2.Read Then
                            codigo_usuario_sistema = SQDR_user2("codigo")
                            If SQDR_user2("Perfil") = 1 Then
                                usuario_perfil = True
                            Else
                                usuario_perfil = False
                            End If
                        Else
                            codigo_usuario_sistema = "0"
                        End If
                        SQDR_user2.Close()

                    End If
                    SQDR_user.Close()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
                End Try
                Me.Hide()
                frm_opn.Show()
            Else
                TextBox1.BackColor = Color.Red
                MsgBox("Usuário e/ou senha inválido(s)..")

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
        End Try
    End Sub

    Private Sub cmb_usuarios_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_usuarios.SelectedIndexChanged
        Try
            nome_usuario_sistema = ""
            nome_usuario_sistema = Trim(cmb_usuarios.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, titulo_as)
        End Try
    End Sub
End Class