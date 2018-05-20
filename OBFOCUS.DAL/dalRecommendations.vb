
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* RecName:        dalRecommendations
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
Public Class dalRecommendations

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum RecommendationsFields
        fldID = 0
        fldRecName = 1
        fldRecommendation = 2
        fldExaminerID = 3
    End Enum


    'Used for transaction support
    Private _Transaction As SqlTransaction = Nothing


    '**************************************************************************
    '*  
    '* RecName:        Transaction
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
    '* RecName:        New
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
    '* RecName:        New
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
    '* RecName:        GetByKey
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
    Public Function GetByKey(ByVal RecommendationID As Integer, _
                ByRef RecName As String, _
                ByRef Recommendation As String, _
                ByRef ExaminerID As Integer) As Boolean
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.RecommendationsFields.fldID) = New SqlParameter("@RecommendationsID", SqlDbType.Int)
        arParameters(Me.RecommendationsFields.fldID).Value = RecommendationID
        arParameters(Me.RecommendationsFields.fldRecName) = New SqlParameter("@RecName", SqlDbType.NVarChar, 100)
        arParameters(Me.RecommendationsFields.fldRecName).Direction = ParameterDirection.Output
        arParameters(Me.RecommendationsFields.fldRecommendation) = New SqlParameter("@Recommendation", SqlDbType.VarChar, 8000)
        arParameters(Me.RecommendationsFields.fldRecommendation).Direction = ParameterDirection.Output
        arParameters(Me.RecommendationsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.RecommendationsFields.fldExaminerID).Direction = ParameterDirection.Output


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spRecommendationsGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spRecommendationsGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If ProcessNull.GetString(arParameters(Me.RecommendationsFields.fldRecName).Value) = "DataNotFound" Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            RecName = ProcessNull.GetString(arParameters(Me.RecommendationsFields.fldRecName).Value)
            Recommendation = ProcessNull.GetString(arParameters(Me.RecommendationsFields.fldRecommendation).Value)
            ExaminerID = ProcessNull.GetInt32(arParameters(Me.RecommendationsFields.fldExaminerID).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* RecName:        Update
    '*
    '* Description: Updates a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was updated or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal RecommendationID As Integer, _
                ByVal RecName As String, _
                ByVal Recommendation As String, _
                   ByVal ExaminerID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.RecommendationsFields.fldID) = New SqlParameter("@RecommendationsID", SqlDbType.Int)
        arParameters(Me.RecommendationsFields.fldID).Value = RecommendationID
        arParameters(Me.RecommendationsFields.fldRecName) = New SqlParameter("@RecName", SqlDbType.NVarChar, 100)
        arParameters(Me.RecommendationsFields.fldRecName).Value = RecName
        arParameters(Me.RecommendationsFields.fldRecommendation) = New SqlParameter("@Recommendation", SqlDbType.VarChar, 8000)
        arParameters(Me.RecommendationsFields.fldRecommendation).Value = Recommendation
        arParameters(Me.RecommendationsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.RecommendationsFields.fldExaminerID).Value = ExaminerID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spRecommendationsUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spRecommendationsUpdate", arParameters)
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
    '* RecName:        Add
    '*
    '* Description: Adds a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef RecommendationID As Integer, _
                ByVal RecName As String, _
                ByVal Recommendation As String, _
                ByVal ExaminerID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.RecommendationsFields.fldID) = New SqlParameter("@RecommendationsID", SqlDbType.Int)
        arParameters(Me.RecommendationsFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.RecommendationsFields.fldRecName) = New SqlParameter("@RecName", SqlDbType.NVarChar, 100)
        arParameters(Me.RecommendationsFields.fldRecName).Value = RecName
        arParameters(Me.RecommendationsFields.fldRecommendation) = New SqlParameter("@Recommendation", SqlDbType.VarChar, 8000)
        arParameters(Me.RecommendationsFields.fldRecommendation).Value = Recommendation
        arParameters(Me.RecommendationsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.RecommendationsFields.fldExaminerID).Value = ExaminerID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spRecommendationsInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spRecommendationsInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            RecommendationID = CType(arParameters(0).Value, Integer)
            Return True
        End If
    End Function




    '**************************************************************************
    '*  
    '* RecName:        Delete
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
        arParameters(0) = New SqlParameter("@RecommendationsID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spRecommendationsDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spRecommendationsDelete", arParameters)
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


End Class 'dalRecommendations
