
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Topic:        dalReference
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
Public Class dalReference

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum ReferenceFields
        fldID = 0
        fldTopic = 1
        fldDetail = 2
        fldInternet = 3
    End Enum


    'Used for transaction support
    Private _Transaction As SqlTransaction = Nothing


    '**************************************************************************
    '*  
    '* Topic:        Transaction
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
    '* Topic:        New
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
    '* Topic:        New
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
    '* Topic:        GetByKey
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
    Public Function GetByKey(ByVal TopicID As Integer, _
                ByRef Topic As String, _
                ByRef Detail As String, _
                ByRef Internet As String) As Boolean
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.ReferenceFields.fldID) = New SqlParameter("@TopicID", SqlDbType.Int)
        arParameters(Me.ReferenceFields.fldID).Value = TopicID
        arParameters(Me.ReferenceFields.fldTopic) = New SqlParameter("@Topic", SqlDbType.NVarChar, 50)
        arParameters(Me.ReferenceFields.fldTopic).Direction = ParameterDirection.Output
        arParameters(Me.ReferenceFields.fldDetail) = New SqlParameter("@Detail", SqlDbType.VarChar, 8000)
        arParameters(Me.ReferenceFields.fldDetail).Direction = ParameterDirection.Output
        arParameters(Me.ReferenceFields.fldInternet) = New SqlParameter("@Internet", SqlDbType.VarChar, 8000)
        arParameters(Me.ReferenceFields.fldInternet).Direction = ParameterDirection.Output


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spReferenceGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spReferenceGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.ReferenceFields.fldTopic).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Topic = ProcessNull.GetString(arParameters(Me.ReferenceFields.fldTopic).Value)
            Detail = ProcessNull.GetString(arParameters(Me.ReferenceFields.fldDetail).Value)
            Internet = ProcessNull.GetString(arParameters(Me.ReferenceFields.fldInternet).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Topic:        Update
    '*
    '* Description: Updates a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was updated or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal TopicID As Integer, _
                ByVal Topic As String, _
                ByVal Detail As String, _
                   ByVal Internet As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ReferenceFields.fldID) = New SqlParameter("@TopicID", SqlDbType.Int)
        arParameters(Me.ReferenceFields.fldID).Value = TopicID
        arParameters(Me.ReferenceFields.fldTopic) = New SqlParameter("@Topic", SqlDbType.NVarChar, 50)
        arParameters(Me.ReferenceFields.fldTopic).Value = Topic
        arParameters(Me.ReferenceFields.fldDetail) = New SqlParameter("@Detail", SqlDbType.VarChar, 8000)
        arParameters(Me.ReferenceFields.fldDetail).Value = Detail
        arParameters(Me.ReferenceFields.fldInternet) = New SqlParameter("@Internet", SqlDbType.VarChar, 8000)
        arParameters(Me.ReferenceFields.fldInternet).Value = Internet
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spReferenceUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spReferenceUpdate", arParameters)
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
    '* Topic:        Add
    '*
    '* Description: Adds a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef TopicID As Integer, _
                ByVal Topic As String, _
                ByVal Detail As String, _
                ByVal Internet As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ReferenceFields.fldID) = New SqlParameter("@TopicID", SqlDbType.Int)
        arParameters(Me.ReferenceFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.ReferenceFields.fldTopic) = New SqlParameter("@Topic", SqlDbType.NVarChar, 50)
        arParameters(Me.ReferenceFields.fldTopic).Value = Topic
        arParameters(Me.ReferenceFields.fldDetail) = New SqlParameter("@Detail", SqlDbType.VarChar, 8000)
        arParameters(Me.ReferenceFields.fldDetail).Value = Detail
        arParameters(Me.ReferenceFields.fldInternet) = New SqlParameter("@Internet", SqlDbType.VarChar, 8000)
        arParameters(Me.ReferenceFields.fldInternet).Value = Internet
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spReferenceInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spReferenceInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            TopicID = CType(arParameters(0).Value, Integer)
            Return True
        End If
    End Function




    '**************************************************************************
    '*  
    '* Topic:        Delete
    '*
    '* Description: Deletes a record from the [PatientInfo] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to delete
    '*
    '* Returns:     Boolean indicating if record was deleted or not. 
    '*              True (record found and deleted); False (otherwise).
    '*
    '**************************************************************************
    Public Function Delete(ByVal ID As Integer) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@TopicID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spReferenceDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spReferenceDelete", arParameters)
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


End Class 'dalReference
