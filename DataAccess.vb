Imports System.Data.SqlClient

Public Class DataAccess
    Private Connection As SqlConnection

    ''' <summary>
    '''     Empty constructor getting a conectionstring
    ''' </summary>
    Public Sub New()
        Dim Con As SqlConnection
        Con = New SqlConnection(ConfigurationManager.AppSettings("connectionString"))
        Connection = Con
    End Sub

    ''' <summary>
    '''     Constructor receiving specific connectionString name
    ''' </summary>
    ''' <param name="CONNECTION_VALUE">ConnectionString name</param>
    Public Sub New(ByVal CONNECTION_VALUE As String)
        Dim Con As SqlConnection
        Con = New SqlConnection(ConfigurationManager.AppSettings(CONNECTION_VALUE))
        CONNECTION = Con
    End Sub

    ''' <summary>
    '''     Execute query returns nothing
    ''' </summary>
    ''' <param name="Query">SQL Query</param>
    Public Sub executeQueryNotReturn(ByVal Query As String)
        Try
            Connection.Open()

            Dim Command As New SqlCommand()
            Command.Connection = Connection
            Command.CommandType = CommandType.Text
            Command.CommandText = Query

            Command.ExecuteNonQuery()

            Connection.Close()
        Catch ex As Exception
            Throw New Exception("Exception: ", ex)
        End Try
    End Sub

    ''' <summary>
    '''     Execute query returns SqlCommand
    ''' </summary>
    ''' <param name="Query">SQL Query</param>
    ''' <returns></returns>
    Public Function executeQueryReturn(ByVal Query As String) As SqlCommand
        Try
            Connection.Open()

            Dim Command As New SqlCommand()
            Command.Connection = Connection
            Command.CommandType = CommandType.Text
            Command.CommandText = Query

            Connection.Close()
            Return Command
        Catch ex As Exception
            Throw New Exception("Exception: ", ex)
        End Try
    End Function

    ''' <summary>
    '''     Execute stored procedure with params returns DataTable
    ''' </summary>
    ''' <param name="StoredProcedure">Stored procedure name</param>
    ''' <param name="Parameters">SQLParameter with params</param>
    ''' <returns></returns>
    Public Function executeSPDataTable(ByVal StoredProcedure As String, ByVal ParamArray Parameters() As SqlParameter) As DataTable
        Try
            Dim Table As New DataTable("Resultados")
            Connection.Open()

            Dim cmd As New SqlCommand
            cmd.Connection = Connection
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = StoredProcedure

            If Not Parameters Is Nothing Then
                For Each Parameter As SqlParameter In Parameters
                    cmd.Parameters.Add(Parameter)
                Next
            End If

            Dim SDA As New SqlDataAdapter(cmd)
            SDA.Fill(Table)
            Connection.Close()

            Return Table
        Catch ex As Exception
            Throw New Exception("Exception: ", ex)
        End Try
    End Function

    ''' <summary>
    '''  Execute Stored procedure with params returns DataSet
    ''' </summary>
    ''' <param name="StoredProcedure">Stored Procedure name</param>
    ''' <param name="Parameters">SqlParameter with params</param>
    ''' <returns></returns>
    Public Function executeSPDataSet(ByVal StoredProcedure As String, ByVal ParamArray Parameters() As SqlParameter) As DataSet
        Try
            Dim DataFinal As New DataSet("Resultados")
            Connection.Open()

            Dim cmd As New SqlCommand
            cmd.Connection = Connection
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = StoredProcedure

            If Not Parameters Is Nothing Then
                For Each Parameter As SqlParameter In Parameters
                    cmd.Parameters.Add(Parameter)
                Next
            End If

            Dim SDA As New SqlDataAdapter(cmd)
            SDA.Fill(DataFinal)

            Connection.Close()
            Return DataFinal
        Catch ex As Exception
            Throw New Exception("Exception: ", ex)
        End Try
    End Function

    ''' <summary>
    '''  Execute stored procedure with params using specific connectionString returns DataTable
    ''' </summary>
    ''' <param name="Conexion">ConnectionString name</param>
    ''' <param name="StoredProcedure">StoredProcedure name</param>
    ''' <param name="Parameters">SqlParameter with params</param>
    ''' <returns></returns>
    Public Function executeSPDataTable(ByVal Conexion As String, ByVal StoredProcedure As String, ByVal ParamArray Parameters() As SqlParameter) As DataTable
        Try
            Dim CON As SqlConnection
            CON = New SqlConnection(ConfigurationManager.AppSettings(Conexion))

            Dim Table As New DataTable("Resultados")
            CON.Open()

            Dim cmd As New SqlCommand
            cmd.Connection = CON
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = StoredProcedure

            If Not Parameters Is Nothing Then
                For Each Parameter As SqlParameter In Parameters
                    cmd.Parameters.Add(Parameter)
                Next
            End If

            Dim SDA As New SqlDataAdapter(cmd)
            SDA.Fill(Table)
            CON.Close()

            Return Table
        Catch ex As Exception
            Throw New Exception("Exception: ", ex)
        End Try
    End Function

    ''' <summary>
    '''  Execute stored procedure with params using specific connectionString returns DataSet
    ''' </summary>
    ''' <param name="Conexion">ConnectionString name</param>
    ''' <param name="StoredProcedure">Stored procedure name</param>
    ''' <param name="Parameters">SqlParameter with params</param>
    ''' <returns></returns>
    Public Function executeSPDataSet(ByVal Conexion As String, ByVal StoredProcedure As String, ByVal ParamArray Parameters() As SqlParameter) As DataSet
        Try
            Dim CON As SqlConnection
            CON = New SqlConnection(ConfigurationManager.AppSettings(Conexion))

            Dim DataFinal As New DataSet("Resultados")
            CON.Open()

            Dim cmd As New SqlCommand
            cmd.Connection = CON
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = StoredProcedure

            If Not Parameters Is Nothing Then
                For Each Parameter As SqlParameter In Parameters
                    cmd.Parameters.Add(Parameter)
                Next
            End If

            Dim SDA As New SqlDataAdapter(cmd)
            SDA.Fill(DataFinal)

            CON.Close()
            Return DataFinal
        Catch ex As Exception
            Throw New Exception("Exception: ", ex)
        End Try
    End Function

End Class
