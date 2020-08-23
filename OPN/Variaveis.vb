Module Variaveis
    Public em_aberto As String
    Public enviada As String
    Public aceita As String
    Public declinada As String

    Public aexecutar As String
    Public execucao_pendente As String
    Public em_execucao As String
    Public executada As String
    Public cancelada_Revogada As String
    Public cancelada As String
    Public Suspensa As String


    Public opn_consulta, as_consulta As Boolean

    Public usuario_perfil As Boolean
    Public cod_edit_OPN As Integer
    Public cod_edit_AS As Integer
    Public cod_status As Integer = 0
    Public cod_status_as As Integer = 0
    Public cod_licitacao As Integer = 0
    Public titulo_opn As String = "Oportunidade de Negócio (OPN)"
    Public titulo_as As String = "Autirização de Serviço (AS)"

    Public status_01, status_02, status_03, status_04, status_05, status_06, status_10 As Integer
    Public status_01_as, status_02_as, status_03_as, status_04_as, status_05_as, status_06_as, status_10_as As Integer
    Public codigo_usuario_sistema As Integer
    Public nome_usuario_sistema As String
End Module
