Imports System.Data.SqlClient

Public Class accesodatosSQL
    Private Shared conexion As New SqlConnection
    Private Shared comando As New SqlCommand

    Public Shared Function conectar() As String
        Try
            conexion.ConnectionString = "Data Source=tcp:hads-18.database.windows.net,1433;Initial Catalog=HADS18_usuarios;Persist Security Info=True;User ID=HADS18@hads-18;Password=4QYzSiq7"
            conexion.Open()
        Catch ex As Exception
            Return "ERROR DE CONEXIÓN: " + ex.Message
        End Try
        Return "CONEXION OK"
    End Function

    Public Shared Function insertar(ByVal Nombre As String, ByVal email As String,
                                    ByVal Apellidos As String, ByVal Password As String,
                                    ByVal Pregunta As String, ByVal Respuesta As String,
                                    ByVal numconfir As String) As String
        Dim st = "insert into Usuarios (email,nombre,apellidos,pregunta,
    respuesta,pass,numconfir) 
            values ('" & email & " ','" & Nombre & "','" & Apellidos & "',
'" & Pregunta & "','" & Respuesta & "','" & Password & "','" & numconfir & "')"
        Dim numregs As Integer
        conectar()
        comando = New SqlCommand(st, conexion)
        Try
            numregs = comando.ExecuteNonQuery()
        Catch ex As Exception
            Return ex.Message
        End Try
        Return (numregs & " registro(s) insertado(s) en la BD ")
    End Function

    Public Shared Sub cerrarconexion()
        conexion.Close()
    End Sub

    Public Shared Function comprobarLogin(ByVal email As String, ByVal password As String) As String
        Dim numregs As Integer
        Dim st = "select count(*) from Usuarios where email='" & email & "' and pass='" & password & "' and confirmado=1"
        conectar()
        comando = New SqlCommand(st, conexion)
        Try
            numregs = comando.ExecuteScalar()
        Catch ex As Exception
            Return ex.Message
        End Try
        If (numregs = 1) Then
            Return "Ok"
        Else
            Return "Error"
        End If
    End Function

    Public Shared Function confirmarRegistro(ByVal email As String, ByVal num As String) As String
        Dim numregs As Integer
        Dim st = "update Usuarios set confirmado=1 where email='" & email & "' and numconfir='" & num & "'"
        conectar()
        comando = New SqlCommand(st, conexion)
        Try
            numregs = comando.ExecuteNonQuery()
        Catch ex As Exception
            Return ex.Message
        End Try
        If (numregs = 1) Then
            Return "Ok"
        Else
            Return "Error"
        End If
    End Function

    Public Shared Function existeUserConPregunta(ByVal email As String, ByVal Pregunta As String, ByVal Respuesta As String) As String
        Dim numregs As Integer
        Dim st = "select count(*) from Usuarios where email='" & email & "' and pregunta='" & Pregunta & "' and respuesta='" & Respuesta & "'"
        conectar()
        comando = New SqlCommand(st, conexion)
        Try
            numregs = comando.ExecuteScalar()
        Catch ex As Exception
            Return ex.Message
        End Try
        If (numregs = 1) Then
            Return "Ok"
        Else
            Return "Error"
        End If
    End Function

    Public Shared Function cambiaPassword(ByVal email As String, ByVal newPass As String) As String
        Dim numregs As Integer
        Dim st = "update Usuarios set pass='" & newPass & "' where email='" & email & "' and confirmado=1"
        conectar()
        comando = New SqlCommand(st, conexion)
        Try
            numregs = comando.ExecuteNonQuery()
        Catch ex As Exception
            Return ex.Message
        End Try
        If (numregs = 1) Then
            Return "Ok"
        Else
            Return "Error"
        End If
    End Function


End Class

