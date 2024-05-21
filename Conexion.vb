Imports System.Data.SqlClient
Public Class Conexion
    Shared cnx As New SqlConnection


    Private Shared Sub Connect()
        Try
            cnx.ConnectionString = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM;Integrated Security=True"
            cnx.Open()
        Catch ex As Exception
            MsgBox("Error al conectar a la base de datos: " & ex.Message)
        End Try
    End Sub
    Private Shared Sub Disconect()
        Try
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            End If
        Catch ex As Exception
            MsgBox("Error al desconectar de la base de datos: " & ex.Message)
        End Try
    End Sub


    'Create a method to excecute a Selection query
    Public Shared Function SelectQuery(ByVal query As String, parametros As List(Of SqlParameter)) As DataTable
        Dim dt As New DataTable
        Try
            Connect()
            Dim cmd As New SqlCommand(query, cnx)
            Dim da As New SqlDataAdapter(cmd)
            da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error al ejecutar la consulta: " & ex.Message)
        Finally
            Disconect()
        End Try
        Return dt
    End Function
End Class
