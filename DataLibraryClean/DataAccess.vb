
Imports System.Data.SqlClient
Imports System.IO

Public Class DataAccess
    Private Const FileConexion As String = "c:\configMonitor\conexion.txt"

    Public Function getCadenaConexion() As String
        Dim conexion As String = ""
        Dim objReader As New StreamReader(FileConexion)
        Dim sLine As String = ""
        Do
            sLine = objReader.ReadLine()
            If Not (sLine Is Nothing) Then
                conexion = sLine
            End If
        Loop Until sLine Is Nothing
        objReader.Close()
        Return conexion
    End Function

    ''' <summary>
    ''' Esta Función permite obtener el tipo de dato a validar en el campo
    ''' table es la tabla de destino
    ''' field es el campo del que obtendremos el tipo
    ''' </summary>
    Public Function GetTypeField(ByVal table As String, field As String) As String

        Using conn As New SqlConnection(getCadenaConexion)

            Dim sql As String = String.Format("SELECT {0}  FROM {1} WHERE 1=0", field, table)
            Dim cmd As New SqlClient.SqlCommand

            Dim myReader As SqlDataReader

            cmd.CommandText = sql
            cmd.Connection = conn
            cmd.Connection.Open()
            myReader = cmd.ExecuteReader(CommandBehavior.KeyInfo)
            Dim schemaTable = myReader.GetSchemaTable()

            For Each myField In schemaTable.Rows 'recorro las filas con los datos de la columna
                For Each myProperty In schemaTable.Columns 'ahora recorro las columnas que contienen las propiedades
                    If myProperty.ColumnName().ToString = "DataType" Then
                        Return myField(myProperty).ToString()
                    End If
                Next
            Next
        End Using
    End Function

    ''' <summary>
    ''' Esta Función permite obtener el valor de un campo en nuestra tabla de origen (todos son varchar)
    ''' table es la tabla de origen
    ''' field es el campo del que obtendremos el valor
    ''' keyId es el id de registro
    ''' </summary>
    Public Function getValue(ByVal table As String, ByVal field As String, ByVal KeyField As String, ByVal KeyId As Integer) As String
        Using conn As New SqlConnection(getCadenaConexion)
            Dim sql As String = String.Format("SELECT {0}  FROM {1} WHERE {2}={3}", field, table, KeyField, KeyId)
            Dim Da As New SqlDataAdapter
            Dim Cmd As New SqlCommand
            If Not conn.State = ConnectionState.Open Then
                conn.Open()
            End If
            With Cmd
                .CommandType = CommandType.Text
                .CommandText = sql
                .Connection = conn
            End With
            Da.SelectCommand = Cmd
            Dim Dt = New DataTable
            Da.Fill(Dt)
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

            If Dt.Rows(0)(field) Is DBNull.Value Then
                Return ""
            Else
                Return Dt.Rows(0)(field)
            End If

        End Using
    End Function

    ''' <summary>
    ''' Esta Función permite obtener todos los registro de la tabla source que no estén en la tabla destiny.
    ''' </summary>
    Public Function GetRowsNews(ByVal sourceTable As String, ByVal destinyTable As String, ByVal FieldKey As String) As DataTable
        Using conn As New SqlConnection(getCadenaConexion)
            Dim sql As String = String.Format("SELECT {0}  FROM {1} WHERE {0} NOT IN(select Coalesce({0},0) from {2}) ", FieldKey, sourceTable, destinyTable)
            Dim dt = GetDataTable(sql)
            Return dt
        End Using

    End Function

    Public Sub SetCleanValue(ByVal value As String, ByVal table As String, ByVal KeyField As String, ByVal KeyId As Integer, ByVal FieldUpdate As String)
        Using conn As New SqlConnection(getCadenaConexion)
            Dim updateSQL = String.Format("UPDATE {0} SET {1} ='{2}' where {3}={4}", table, FieldUpdate, value, KeyField, KeyId)
            Dim Cmd As New SqlCommand
            If Not conn.State = ConnectionState.Open Then
                conn.Open()
            End If
            With Cmd
                .CommandType = CommandType.Text
                .CommandText = updateSQL
                .Connection = conn
                .ExecuteNonQuery()
            End With
        End Using
    End Sub


    Public Sub CopyData(ByVal sp As String)
        Try
            Dim ConnString = getCadenaConexion()

            Using conn As New SqlConnection(ConnString)
                Using comm As New SqlCommand(sp)
                    comm.CommandType = CommandType.StoredProcedure
                    comm.Connection = conn
                    conn.Open()
                    comm.ExecuteNonQuery()
                End Using
                conn.Close()
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function GetDataTable(ByVal comando As String) As DataTable
        Try
            Using cnSql As New SqlConnection(getCadenaConexion)
                Dim Dt As DataTable
                Dim Da As New SqlDataAdapter
                Dim Cmd As New SqlCommand
                If Not cnSql.State = ConnectionState.Open Then
                    cnSql.Open()
                End If
                With Cmd
                    .CommandType = CommandType.Text
                    .CommandText = comando
                    .Connection = cnSql
                End With
                Da.SelectCommand = Cmd
                Dt = New DataTable
                Da.Fill(Dt)
                If cnSql.State = ConnectionState.Open Then
                    cnSql.Close()
                End If
                Return Dt
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

End Class

