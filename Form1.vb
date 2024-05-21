

Imports System.Data.SqlClient

Public Class Form1

    ''Create a constructor
    'Public Sub New()
    '    ' This call is required by the designer.
    '    InitializeComponent()
    '    ' Add any initialization after the InitializeComponent() call.
    '    'Create a new instance of the Connection class
    '    Dim query As String = "select * from Museo"
    '    DataGridView1.DataSource = Conexion.SelectQuery(query)
    'End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CargarMuseos()
    End Sub

    Private Sub CargarMuseos()
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Dim query As String = "SELECT DISTINCT Centro_de_trabajo FROM VisitantesMuseos"

        Using connection As New SqlConnection(connectionString)
            Using adapter As New SqlDataAdapter(query, connection)
                Dim dataSet As New DataSet()

                Try
                    connection.Open()
                    adapter.Fill(dataSet, "Museos")

                    ComboBox1.DataSource = dataSet.Tables("Museos")
                    ComboBox1.DisplayMember = "Centro_de_trabajo"

                Catch ex As Exception
                    MessageBox.Show("Error al cargar los museos: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        If ComboBox1.SelectedItem IsNot Nothing Then
            Dim nombreMuseo As String = ComboBox1.Text
            CargarInfoMuseo(nombreMuseo)
        End If
    End Sub

    Private Sub CargarInfoMuseo(nombreMuseo As String)
        Dim connectionString As String = "Data Source=LESLY\MSSQLSERVER01; Initial Catalog=VisitasM; Integrated Security=True"
        Dim query As String = "SELECT * FROM VisitantesMuseos WHERE Centro_de_trabajo = @Centro_de_trabajo"

        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@Centro_de_trabajo", nombreMuseo)

                Dim adapter As New SqlDataAdapter(command)
                Dim dataSet As New DataSet()

                Try
                    connection.Open()
                    adapter.Fill(dataSet, "InfoMuseo")

                    DataGridView1.DataSource = dataSet.Tables("InfoMuseo")

                Catch ex As Exception
                    MessageBox.Show("Error al cargar la información del museo: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub
End Class


