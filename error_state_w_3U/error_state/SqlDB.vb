Imports System.Data.SqlClient
Imports System.Data.Common

Public Class SqlDB

    Private dbProviderFactory As DbProviderFactory
    Private _DBConnName As String
    Private _DBConnStr As String
    Private _conn As DbConnection
    Private _NotCloseConnection As Boolean = False

    Sub New()
        BuildConn()
    End Sub

    ''' <summary>
    ''' 控制 Connection是否關閉,True 不要關閉,適合大量執行sql語法
    ''' </summary>
    ''' <param name="NotClostConnection"></param>
    ''' <remarks></remarks>
    Sub New(ByVal NotClostConnection As Boolean)
        _NotCloseConnection = NotClostConnection
        BuildConn()
    End Sub

    Private Sub BuildConn()

        dbProviderFactory = DbProviderFactories.GetFactory("System.Data.SqlClient")
        '_DBConnStr = "Server=192.168.4.166;Database=package;User Id=packageuser;Password=packageuser;"
        _DBConnStr = Global.error_state.My.MySettings.Default.packageConnectionString1
        _conn = dbProviderFactory.CreateConnection()

        _conn.ConnectionString = _DBConnStr
    End Sub


    ''' <summary>
    ''' 開啟資料庫連結
    ''' </summary>
    Public Sub OpenConnection()
        If _conn.State <> ConnectionState.Open Then
            _conn.Open()
        End If

    End Sub


    ''' <summary>
    ''' 關閉資料庫連結
    ''' </summary>
    Public Sub CloseConnection()
        If _conn.State <> ConnectionState.Closed And Not _NotCloseConnection Then
            _conn.Close()
        End If
    End Sub

    ''' <summary>
    ''' 讀取資料表，傳入SQL語法
    ''' </summary>
    Public Function GetDataTable(ByVal SQLCmdStr As String, ByVal CommandType As CommandType) As DataTable
        Dim cmd As DbCommand = dbProviderFactory.CreateCommand()
        cmd.Connection = _conn
        cmd.CommandType = CommandType
        cmd.CommandText = SQLCmdStr

        Dim da As DbDataAdapter = dbProviderFactory.CreateDataAdapter()
        DA.SelectCommand = cmd

        Dim DT As DataTable = New DataTable
        OpenConnection()
        DA.Fill(DT)
        CloseConnection()
        Return DT

    End Function


    ''' <summary>
    ''' 讀取資料表，傳入SQL語法、參數、指令型號
    ''' </summary>
    Public Function GetDataTable(ByVal SQLCmdStr As String, ByVal Parameter As ArrayList, ByVal CommandType As CommandType) As DataTable
        Dim cmd As DbCommand = dbProviderFactory.CreateCommand()
        cmd.Connection = _conn
        cmd.CommandType = CommandType
        cmd.CommandText = SQLCmdStr



        For Each param In Parameter
            cmd.Parameters.Add(param)
        Next

        Dim da As DbDataAdapter = dbProviderFactory.CreateDataAdapter()
        DA.SelectCommand = cmd

        Dim DT As DataTable = New DataTable
        OpenConnection()
        DA.Fill(DT)
        CloseConnection()

        Return DT

    End Function


    ''' <summary>
    ''' 讀取資料表，傳入SQL語法、參數、指令型號
    ''' </summary>
    Public Function GetDataTable(ByVal SQLCmdStr As String, ByVal Parameter As List(Of SqlParameter), ByVal CommandType As CommandType) As DataTable
        Dim cmd As DbCommand = dbProviderFactory.CreateCommand()
        cmd.Connection = _conn
        cmd.CommandType = CommandType
        cmd.CommandText = SQLCmdStr

        For Each param In Parameter
            cmd.Parameters.Add(param)
        Next

        Dim da As DbDataAdapter = dbProviderFactory.CreateDataAdapter()
        DA.SelectCommand = cmd

        Dim DT As DataTable = New DataTable
        OpenConnection()
        DA.Fill(DT)
        CloseConnection()

        Return DT

    End Function



    ''' <summary>
    ''' 執行SQL指令，傳入SQL語法
    ''' </summary>    
    Public Function RunCmd(ByVal SQLCmdStr As String, ByVal CommandType As CommandType) As Integer
        Dim cmd As DbCommand = dbProviderFactory.CreateCommand()
        cmd.Connection = _conn
        cmd.CommandType = CommandType
        cmd.CommandText = SQLCmdStr

        OpenConnection()
        Dim Result As Integer

        Result = cmd.ExecuteNonQuery()

        CloseConnection()

        Return Result

    End Function




    ''' <summary>
    ''' 執行SQL指令，傳入SQL語法
    ''' </summary>    
    Public Function RunCmd(ByVal SQLCmdStr As String, ByVal Parameter As ArrayList, ByVal CommandType As CommandType) As Integer
        Dim cmd As DbCommand = dbProviderFactory.CreateCommand()
        cmd.Connection = _conn
        cmd.CommandType = CommandType
        cmd.CommandText = SQLCmdStr

        For Each param In Parameter
            cmd.Parameters.Add(param)
        Next

        OpenConnection()
        Dim Result As Integer

        Result = cmd.ExecuteNonQuery()

        CloseConnection()

        Return Result

    End Function


    ''' <summary>
    ''' 執行SQL指令，傳入SQL語法
    ''' </summary>    
    Public Function RunCmd(ByVal SQLCmdStr As String, ByVal Parameter As List(Of SqlParameter), ByVal CommandType As CommandType) As Integer
        Dim cmd As DbCommand = dbProviderFactory.CreateCommand()
        cmd.Connection = _conn
        cmd.CommandType = CommandType
        cmd.CommandText = SQLCmdStr

        For Each param In Parameter
            cmd.Parameters.Add(param)
        Next

        OpenConnection()
        Dim Result As Integer

        Result = cmd.ExecuteNonQuery()

        CloseConnection()

        Return Result

    End Function

End Class
