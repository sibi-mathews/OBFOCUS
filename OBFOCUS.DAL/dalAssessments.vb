
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalAssessments
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
Public Class dalAssessments

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum AssessmentsFields
        fldID = 0
        fldName = 1
        fldAssessment = 2
        fldRecommendation1 = 3
        fldExaminerID = 4
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
    Public Function GetByKey(ByVal AssessmentID As Integer, _
                ByRef Name As String, _
                ByRef Assessment As String, _
                ByRef Recommendation1 As String, _
                ByRef ExaminerID As Integer) As Boolean
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.AssessmentsFields.fldID) = New SqlParameter("@AssessmentID", SqlDbType.Int)
        arParameters(Me.AssessmentsFields.fldID).Value = AssessmentID
        arParameters(Me.AssessmentsFields.fldName) = New SqlParameter("@Name", SqlDbType.NVarChar, 50)
        arParameters(Me.AssessmentsFields.fldName).Direction = ParameterDirection.Output
        arParameters(Me.AssessmentsFields.fldAssessment) = New SqlParameter("@Assessment", SqlDbType.VarChar, 8000)
        arParameters(Me.AssessmentsFields.fldAssessment).Direction = ParameterDirection.Output
        arParameters(Me.AssessmentsFields.fldRecommendation1) = New SqlParameter("@Recommendation1", SqlDbType.VarChar, 8000)
        arParameters(Me.AssessmentsFields.fldRecommendation1).Direction = ParameterDirection.Output
        arParameters(Me.AssessmentsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.AssessmentsFields.fldExaminerID).Direction = ParameterDirection.Output


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAssessmentsGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAssessmentsGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If ProcessNull.GetString(arParameters(Me.AssessmentsFields.fldName).Value) = "DataNotFound" Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Name = ProcessNull.GetString(arParameters(Me.AssessmentsFields.fldName).Value)
            Assessment = ProcessNull.GetString(arParameters(Me.AssessmentsFields.fldAssessment).Value)
            Recommendation1 = ProcessNull.GetString(arParameters(Me.AssessmentsFields.fldRecommendation1).Value)
            ExaminerID = ProcessNull.GetInt32(arParameters(Me.AssessmentsFields.fldExaminerID).Value)
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
    '* Description: Updates a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was updated or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal AssessmentID As Integer, _
                ByVal Name As String, _
                ByVal Assessment As String, _
                ByVal Recommendation1 As String, _
                ByVal ExaminerID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.AssessmentsFields.fldID) = New SqlParameter("@AssessmentID", SqlDbType.Int)
        arParameters(Me.AssessmentsFields.fldID).Value = AssessmentID
        arParameters(Me.AssessmentsFields.fldName) = New SqlParameter("@Name", SqlDbType.NVarChar, 50)
        arParameters(Me.AssessmentsFields.fldName).Value = Name
        arParameters(Me.AssessmentsFields.fldAssessment) = New SqlParameter("@Assessment", SqlDbType.VarChar, 8000)
        arParameters(Me.AssessmentsFields.fldAssessment).Value = Assessment
        arParameters(Me.AssessmentsFields.fldRecommendation1) = New SqlParameter("@Recommendation1", SqlDbType.VarChar, 8000)
        arParameters(Me.AssessmentsFields.fldRecommendation1).Value = Recommendation1
        arParameters(Me.AssessmentsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.AssessmentsFields.fldExaminerID).Value = ExaminerID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAssessmentsUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAssessmentsUpdate", arParameters)
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
    Public Function Add(ByRef AssessmentID As Integer, _
                ByVal Name As String, _
                ByVal Assessment As String, _
                ByVal Recommendation1 As String, _
                ByVal ExaminerID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.AssessmentsFields.fldID) = New SqlParameter("@AssessmentID", SqlDbType.Int)
        arParameters(Me.AssessmentsFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.AssessmentsFields.fldName) = New SqlParameter("@Name", SqlDbType.NVarChar, 50)
        arParameters(Me.AssessmentsFields.fldName).Value = Name
        arParameters(Me.AssessmentsFields.fldAssessment) = New SqlParameter("@Assessment", SqlDbType.VarChar, 8000)
        arParameters(Me.AssessmentsFields.fldAssessment).Value = Assessment
        arParameters(Me.AssessmentsFields.fldRecommendation1) = New SqlParameter("@Recommendation1", SqlDbType.VarChar, 8000)
        arParameters(Me.AssessmentsFields.fldRecommendation1).Value = Recommendation1
        arParameters(Me.AssessmentsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.AssessmentsFields.fldExaminerID).Value = ExaminerID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAssessmentsInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAssessmentsInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            AssessmentID = CType(arParameters(0).Value, Integer)
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
    Public Function Delete(ByVal ID As Integer) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@AssessmentID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAssessmentsDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAssessmentsDelete", arParameters)
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


End Class 'dalAssessments
