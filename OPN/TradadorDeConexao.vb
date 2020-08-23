Imports System.Data.SqlClient
Public Class TratadorDeConexao
    Public Administrador As Boolean
    Shared SUPER As Boolean
    Public iIndex As Integer
    Shared Con As SqlConnection
    Shared Con2 As SqlConnection    'Banco TOTVS
    Public Shared Sub Preparar(ByVal StrConn As String)
        Con = New SqlConnection(StrConn)
    End Sub
    Public Shared Sub Preparar2(ByVal StrConn As String) 'Banco TOTVS
        Con2 = New SqlConnection(StrConn)
    End Sub
    Shared Function Abrir()
        Try
            Con.Open()
            'transaction = Con.BeginTransaction()
            Abrir = True
        Catch ex As Exception
            Err.Clear()
            Abrir = False
        End Try
    End Function
    Shared Function Abrir2() 'Banco TOTVS
        Try
            Con2.Open()
            'transaction = Con.BeginTransaction()
            Abrir2 = True
        Catch ex As Exception
            Err.Clear()
            Abrir2 = False
        End Try
    End Function
    Public Shared Sub Fechar()
        Con.Close()
    End Sub
    Public Shared Sub Fechar2() 'Banco TOTVS
        Con2.Close()
    End Sub
    Public Shared Function Conexao() As SqlConnection
        Conexao = Con
    End Function
    Public Shared Function Conexao2() As SqlConnection
        Conexao2 = Con2
    End Function
End Class
