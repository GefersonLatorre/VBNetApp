Imports System
Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml.Linq

Module Module1
    Sub Main()
        Dim config = LoadConfiguration()
        Dim connectionString = config.ConnectionString
        Dim logFilePath = config.LogFilePath

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                ' Ejecutar procedimientos almacenados aquí
                ' Ejemplo: ExecuteStoredProcedure(connection, "InsertarUsuario", ...)

                Dim results = ExecuteStoredProcedureWithResults(connection, "FiltrarUsuariosPorEdad", New SqlParameter() {
                    New SqlParameter("@Edad", 25)
                })

                File.WriteAllText(logFilePath, results)
                Console.WriteLine("Procedimientos ejecutados y resultados registrados en el archivo de log.")
            End Using
        Catch ex As Exception
            File.AppendAllText(logFilePath, $"Error: {ex.Message}" & Environment.NewLine)
            Console.WriteLine($"Error: {ex.Message}")
        End Try
    End Sub

    Function LoadConfiguration() As (ConnectionString As String, LogFilePath As String)
        Dim xml = XDocument.Load("config.xml")
        Dim connectionString = xml.Root.Element("ConnectionStrings").Element("add").Attribute("connectionString").Value
        Dim logFilePath = xml.Root.Element("LogFilePath").Value

        Return (connectionString, logFilePath)
    End Function

    Sub ExecuteStoredProcedure(connection As SqlConnection, procedureName As String, parameters As SqlParameter())
        Using command As New SqlCommand(procedureName, connection)
            command.CommandType = System.Data.CommandType.StoredProcedure
            command.Parameters.AddRange(parameters)
            command.ExecuteNonQuery()
        End Using
    End Sub

    Function ExecuteStoredProcedureWithResults(connection As SqlConnection, procedureName As String, parameters As SqlParameter()) As String
        Using command As New SqlCommand(procedureName, connection)
            command.CommandType = System.Data.CommandType.StoredProcedure
            command.Parameters.AddRange(parameters)

            Using reader = command.ExecuteReader()
                Dim result = ""
                While reader.Read()
                    result &= $"{reader("Nombre")}, {reader("Apellido")}, {reader("Edad")}, {reader("Correo")}, {reader("Hobbies")}, {reader("Activo")}" & Environment.NewLine
                End While
                Return result
            End Using
        End Using
    End Function
End Module
