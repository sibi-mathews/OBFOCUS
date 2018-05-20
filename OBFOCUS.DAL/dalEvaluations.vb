
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* EvaluationName:        dalEvaluations
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
Public Class dalEvaluations

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum EvaluationsFields
        fldID = 0
        fldEvaluationName = 1
        fldEvaluation = 2
        fldRecommendation3 = 3
        fldExaminerID = 4
    End Enum


    'Used for transaction support
    Private _Transaction As SqlTransaction = Nothing


    '**************************************************************************
    '*  
    '* EvaluationName:        Transaction
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
    '* EvaluationName:        New
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
    '* EvaluationName:        New
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
    '* EvaluationName:        GetByKey
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
    Public Function GetByKey(ByVal EvaluationID As Integer, _
                ByRef EvaluationName As String, _
                ByRef Evaluation As String, _
                ByRef Recommendation3 As String, _
                ByRef ExaminerID As Integer) As Boolean
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.EvaluationsFields.fldID) = New SqlParameter("@EvaluationID", SqlDbType.Int)
        arParameters(Me.EvaluationsFields.fldID).Value = EvaluationID
        arParameters(Me.EvaluationsFields.fldEvaluationName) = New SqlParameter("@EvaluationName", SqlDbType.NVarChar, 50)
        arParameters(Me.EvaluationsFields.fldEvaluationName).Direction = ParameterDirection.Output
        arParameters(Me.EvaluationsFields.fldEvaluation) = New SqlParameter("@Evaluation", SqlDbType.VarChar, 8000)
        arParameters(Me.EvaluationsFields.fldEvaluation).Direction = ParameterDirection.Output
        arParameters(Me.EvaluationsFields.fldRecommendation3) = New SqlParameter("@Recommendation3", SqlDbType.VarChar, 8000)
        arParameters(Me.EvaluationsFields.fldRecommendation3).Direction = ParameterDirection.Output
        arParameters(Me.EvaluationsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.EvaluationsFields.fldExaminerID).Direction = ParameterDirection.Output


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spEvaluationsGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spEvaluationsGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.EvaluationsFields.fldEvaluationName).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            EvaluationName = ProcessNull.GetString(arParameters(Me.EvaluationsFields.fldEvaluationName).Value)
            Evaluation = ProcessNull.GetString(arParameters(Me.EvaluationsFields.fldEvaluation).Value)
            Recommendation3 = ProcessNull.GetString(arParameters(Me.EvaluationsFields.fldRecommendation3).Value)
            ExaminerID = ProcessNull.GetInt32(arParameters(Me.EvaluationsFields.fldExaminerID).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* EvaluationName:        Update
    '*
    '* Description: Updates a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was updated or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal EvaluationID As Integer, _
                ByVal EvaluationName As String, _
                ByVal Evaluation As String, _
                ByVal Recommendation3 As String, _
                ByVal ExaminerID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.EvaluationsFields.fldID) = New SqlParameter("@EvaluationID", SqlDbType.Int)
        arParameters(Me.EvaluationsFields.fldID).Value = EvaluationID
        arParameters(Me.EvaluationsFields.fldEvaluationName) = New SqlParameter("@EvaluationName", SqlDbType.NVarChar, 50)
        arParameters(Me.EvaluationsFields.fldEvaluationName).Value = EvaluationName
        arParameters(Me.EvaluationsFields.fldEvaluation) = New SqlParameter("@Evaluation", SqlDbType.VarChar, 8000)
        arParameters(Me.EvaluationsFields.fldEvaluation).Value = Evaluation
        arParameters(Me.EvaluationsFields.fldRecommendation3) = New SqlParameter("@Recommendation3", SqlDbType.VarChar, 8000)
        arParameters(Me.EvaluationsFields.fldRecommendation3).Value = Recommendation3
        arParameters(Me.EvaluationsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.EvaluationsFields.fldExaminerID).Value = ExaminerID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spEvaluationsUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spEvaluationsUpdate", arParameters)
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
    '* EvaluationName:        Add
    '*
    '* Description: Adds a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef EvaluationID As Integer, _
                ByVal EvaluationName As String, _
                ByVal Evaluation As String, _
                ByVal Recommendation3 As String, _
                ByVal ExaminerID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.EvaluationsFields.fldID) = New SqlParameter("@EvaluationID", SqlDbType.Int)
        arParameters(Me.EvaluationsFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.EvaluationsFields.fldEvaluationName) = New SqlParameter("@EvaluationName", SqlDbType.NVarChar, 50)
        arParameters(Me.EvaluationsFields.fldEvaluationName).Value = EvaluationName
        arParameters(Me.EvaluationsFields.fldEvaluation) = New SqlParameter("@Evaluation", SqlDbType.VarChar, 8000)
        arParameters(Me.EvaluationsFields.fldEvaluation).Value = Evaluation
        arParameters(Me.EvaluationsFields.fldRecommendation3) = New SqlParameter("@Recommendation3", SqlDbType.VarChar, 8000)
        arParameters(Me.EvaluationsFields.fldRecommendation3).Value = Recommendation3
        arParameters(Me.EvaluationsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.EvaluationsFields.fldExaminerID).Value = ExaminerID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spEvaluationsInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spEvaluationsInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            EvaluationID = CType(arParameters(0).Value, Integer)
            Return True
        End If
    End Function




    '**************************************************************************
    '*  
    '* EvaluationName:        Delete
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
        arParameters(0) = New SqlParameter("@EvaluationID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spEvaluationsDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spEvaluationsDelete", arParameters)
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


End Class 'dalEvaluations
