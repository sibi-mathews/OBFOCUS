
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalCalls
'*
'* Description: Data access layer for Table PatientInfo
'*
'* Remarks:     Uses OleDb and embedded SQL for maintaining the data.
'*-----------------------------------------------------------------------------
'*                      CHANGE HISTORY
'*   Change No:   Date:          Author:   Description:
'*   _________    ___________    ______    ____________________________________
'*      001       1/26/2005     MR        Created.                                
'* 
'******************************************************************************
Public Class dalCalls

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum WdiagnosisFields
        fldID = 0
        fldDateStarted = 1
        fldDateStopped = 2
        fldTradeName = 3
        fldPharmID = 4
        fldDosage = 5
        fldFrequency = 6
        fldRoute = 7
        fldDateCreated = 8
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



#Region "Main procedures - GetLab, Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Name:        GetCalls
    '*
    '* Description: Returns all records in the [Labs] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetCalls(ByVal ChartID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ChartID
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spCallsGetByKey", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spCallsGetByKey", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
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
    Public Function Update(ByVal ID As Integer, _
                       ByVal CallDate As String, _
                       ByVal SpokeTo As String, _
                       ByVal Message As String, _
                       ByVal Examiner2ID As Integer) As Boolean

        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@CallID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@CallDate", SqlDbType.SmallDateTime)
        If CallDate = "" Then
            arParameters(1).Value = DBNull.Value
        Else
            arParameters(1).Value = CallDate
        End If
        arParameters(2) = New SqlParameter("@SpokeTo", SqlDbType.NVarChar, 50)
        arParameters(2).Value = SpokeTo
        arParameters(3) = New SqlParameter("@Message", SqlDbType.NVarChar, 255)
        arParameters(3).Value = Message
        arParameters(4) = New SqlParameter("@Examiner2ID", SqlDbType.Int)
        If Examiner2ID = Nothing Then
            arParameters(4).Value = DBNull.Value
        Else
            arParameters(4).Value = Examiner2ID
        End If
       
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spCallsUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spCallsUpdate", arParameters)
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
    '* Description: Adds a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef ID As Integer, _
                        ByVal CallDate As String, _
                       ByVal SpokeTo As String, _
                       ByVal Message As String, _
                       ByVal Examiner2ID As Integer, _
                       ByVal ChartID As Integer) As Boolean


        Dim arParameters(5) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@CallID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@CallDate", SqlDbType.SmallDateTime)
        If CallDate = "" Then
            arParameters(1).Value = DBNull.Value
        Else
            arParameters(1).Value = CallDate
        End If
        arParameters(2) = New SqlParameter("@SpokeTo", SqlDbType.NVarChar, 50)
        arParameters(2).Value = SpokeTo
        arParameters(3) = New SqlParameter("@Message", SqlDbType.NVarChar, 255)
        arParameters(3).Value = Message
        arParameters(4) = New SqlParameter("@Examiner2ID", SqlDbType.Int)
        If Examiner2ID = Nothing Then
            arParameters(4).Value = DBNull.Value
        Else
            arParameters(4).Value = Examiner2ID
        End If
        arParameters(5) = New SqlParameter("@ChartID", SqlDbType.Int)
        If ChartID = Nothing Then
            arParameters(5).Value = DBNull.Value
        Else
            arParameters(5).Value = ChartID
        End If
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spCallsInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spCallsInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            ID = CType(arParameters(0).Value, Integer)
            Return True
        End If

    End Function




    '**************************************************************************
    '*  
    '* Name:        Delete
    '*
    '* Description: Deletes a record from the [PatientInfo] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to delete
    '*
    '* Returns:     Boolean indicating if record was deleted or not. 
    '*              True (record found and deleted); False (otherwise).
    '*
    '**************************************************************************
    Public Function Delete(ByVal CallID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@CallID", SqlDbType.Int)
        arParameters(0).Value = CallID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spCallsDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spCallsDelete", arParameters)
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


End Class 'dalCalls
