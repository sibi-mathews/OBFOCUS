
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalUltrasound
'*
'* Description: Data access layer for Table Ultrasound
'*
'* Remarks:     Uses OleDb and embedded SQL for maintaining the data.
'*-----------------------------------------------------------------------------
'*                      CHANGE HISTORY
'*   Change No:   Date:          Author:   Description:
'*   _________    ___________    ______    ____________________________________
'*      001       1/26/2005     MR        Created.                                
'* 
'******************************************************************************
Public Class dalUltrasound

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum PhysicianFields
        fldUID = 0
        fldUltrasoundPath = 1
        fldUltrasoundProgPath = 2
        fldMode = 3
        fldImageOutProgPath = 4
    End Enum


    'Used for transaction support
    Private _Transaction As SqlTransaction = Nothing


    '**************************************************************************
    '*  
    '* Name:        Transaction
    '*
    '* Description: Used for transaction support.
    '*
    '* Parameters:  If this property is set, all database operations will be
    '*              performed in the context of a database transaction.
    '*
    '**************************************************************************
    Public Property Transaction() As SqlTransaction
        Get
            Return _Transaction
        End Get
        Set(ByVal Value As SqlTransaction)
            _Transaction = Value
        End Set
    End Property 'Transaction

#End Region



#Region "Constructors"

    '**************************************************************************
    '*  
    '* Name:        New
    '*
    '* Description: Initialize a new instance of the class.
    '*
    '* Parameters:  None
    '*
    '**************************************************************************
    Public Sub New()
    End Sub 'New


    '**************************************************************************
    '*  
    '* Name:        New
    '*
    '* Description: Initialize a new instance of the class.
    '*
    '* Parameters:  Transaction - used for transaction support.
    '*
    '**************************************************************************
    Public Sub New(ByRef Transaction As SqlTransaction)
        Me.Transaction = Transaction
    End Sub 'New

#End Region



#Region "Main procedures - GetComboDual, Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Name:        GetUltrasound
    '*
    '* Description: Returns all records in the [Nurses Flow Sheet] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetUltrasound() As SqlDataReader
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasoundPathGetAll")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spUltrasoundPathGetAll")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetByKey
    '*
    '* Description: Gets all the values of a record identified by a key.
    '*
    '* Parameters:  Description - Output parameter
    '*              Picture - Output parameter
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function GetByKey(ByVal mode As String, _
                ByRef UltrasoundPath As String, _
                ByRef UltrasoundProgPath As String, _
                ByRef ImageOutProgPath As String, _
                ByRef UID As Integer, _
                ByRef IPAddress As String, _
                ByRef DataIPAddress As String) As Boolean
        Dim arParameters(6) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@mode", SqlDbType.NVarChar, 100)
        arParameters(0).Value = mode
        arParameters(1) = New SqlParameter("@UltrasoundPath", SqlDbType.NVarChar, 100)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@UltrasoundProgPath", SqlDbType.NVarChar, 100)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@ImageOutProgPath", SqlDbType.VarChar, 100)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@IPAddress", SqlDbType.NVarChar, 250)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@DataIPAddress", SqlDbType.NVarChar, 100)
        arParameters(6).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasoundGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUltrasoundGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(1).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            UltrasoundPath = ProcessNull.GetString(arParameters(1).Value)
            UltrasoundProgPath = ProcessNull.GetString(arParameters(2).Value)
            ImageOutProgPath = ProcessNull.GetString(arParameters(3).Value)
            UID = ProcessNull.GetInt32(arParameters(4).Value)
            IPAddress = ProcessNull.GetString(arParameters(5).Value)
            DataIPAddress = ProcessNull.GetString(arParameters(6).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        Update
    '*
    '* Description: Updates a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal UID As Integer, _
                            ByVal UltrasoundPath As String, _
                            ByVal UltrasoundProgPath As String, _
                            ByVal Mode As String, _
                            ByVal ImageOutProgPath As String) As Boolean

        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(0).Value = UID
        arParameters(1) = New SqlParameter("@UltrasoundPath", SqlDbType.NVarChar, 100)
        arParameters(1).Value = UltrasoundPath
        arParameters(2) = New SqlParameter("@UltrasoundProgPath", SqlDbType.NVarChar, 100)
        arParameters(2).Value = UltrasoundProgPath
        arParameters(3) = New SqlParameter("@Mode", SqlDbType.NVarChar, 50)
        arParameters(3).Value = Mode
        arParameters(4) = New SqlParameter("@ImageOutProgPath", SqlDbType.NVarChar, 100)
        arParameters(4).Value = ImageOutProgPath
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasoundPathUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUltrasoundPathUpdate", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not updated.
        If intRecordsAffected = 0 Then
            Return False
        Else
            Return True
        End If

    End Function

    '**************************************************************************
    '*  
    '* Name:        Add
    '*
    '* Description: Adds a new record to the [Ultrasound] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef uID As Integer, _
                            ByVal UltrasoundPath As String, _
                            ByVal UltrasoundProgPath As String, _
                            ByVal Mode As String, _
                            ByVal ImageOutProgPath As String) As Boolean

        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@UltrasoundPath", SqlDbType.NVarChar, 100)
        arParameters(1).Value = UltrasoundPath
        arParameters(2) = New SqlParameter("@UltrasoundProgPath", SqlDbType.NVarChar, 100)
        arParameters(2).Value = UltrasoundProgPath
        arParameters(3) = New SqlParameter("@Mode", SqlDbType.NVarChar, 50)
        arParameters(3).Value = Mode
        arParameters(4) = New SqlParameter("@ImageOutProgPath", SqlDbType.NVarChar, 100)
        arParameters(4).Value = ImageOutProgPath

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasoundPathInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUltrasoundPathInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            uID = CType(arParameters(0).Value, Integer)
            Return True
        End If

    End Function





    '**************************************************************************
    '*  
    '* Name:        Delete
    '*
    '* Description: Deletes a record from the [Ultrasound] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to delete
    '*
    '* Returns:     Boolean indicating if record was deleted or not. 
    '*              True (record found and deleted); False (otherwise).
    '*
    '**************************************************************************
    Public Function Delete(ByVal uID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(0).Value = uID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasoundPathDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUltrasoundPathDelete", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not updated.
        If intRecordsAffected = 0 Then
            Return False
        Else

            Return True
        End If

    End Function
#End Region


End Class 'dalUltrasound
