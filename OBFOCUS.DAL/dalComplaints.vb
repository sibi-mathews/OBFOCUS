
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* ComplaintName:        dalComplaints
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
Public Class dalComplaints

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum ComplaintsFields
        fldID = 0
        fldComplaintName = 1
        fldComplaint = 2
        fldRecommendation2 = 3
        fldExaminerID = 4
    End Enum


    'Used for transaction support
    Private _Transaction As SqlTransaction = Nothing


    '**************************************************************************
    '*  
    '* ComplaintName:        Transaction
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
    '* ComplaintName:        New
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
    '* ComplaintName:        New
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
    '* ComplaintName:        GetByKey
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
    Public Function GetByKey(ByVal ComplaintID As Integer, _
                ByRef ComplaintName As String, _
                ByRef Complaint As String, _
                ByRef Recommendation2 As String, _
                ByRef ExaminerID As Integer) As Boolean
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.ComplaintsFields.fldID) = New SqlParameter("@ComplaintID", SqlDbType.Int)
        arParameters(Me.ComplaintsFields.fldID).Value = ComplaintID
        arParameters(Me.ComplaintsFields.fldComplaintName) = New SqlParameter("@ComplaintName", SqlDbType.NVarChar, 100)
        arParameters(Me.ComplaintsFields.fldComplaintName).Direction = ParameterDirection.Output
        arParameters(Me.ComplaintsFields.fldComplaint) = New SqlParameter("@Complaint", SqlDbType.VarChar, 8000)
        arParameters(Me.ComplaintsFields.fldComplaint).Direction = ParameterDirection.Output
        arParameters(Me.ComplaintsFields.fldRecommendation2) = New SqlParameter("@Recommendation2", SqlDbType.VarChar, 8000)
        arParameters(Me.ComplaintsFields.fldRecommendation2).Direction = ParameterDirection.Output
        arParameters(Me.ComplaintsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.ComplaintsFields.fldExaminerID).Direction = ParameterDirection.Output


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spComplaintsGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spComplaintsGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If ProcessNull.GetString(arParameters(Me.ComplaintsFields.fldComplaintName).Value) = "DataNotFound" Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            ComplaintName = ProcessNull.GetString(arParameters(Me.ComplaintsFields.fldComplaintName).Value)
            Complaint = ProcessNull.GetString(arParameters(Me.ComplaintsFields.fldComplaint).Value)
            Recommendation2 = ProcessNull.GetString(arParameters(Me.ComplaintsFields.fldRecommendation2).Value)
            ExaminerID = ProcessNull.GetInt32(arParameters(Me.ComplaintsFields.fldExaminerID).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* ComplaintName:        Update
    '*
    '* Description: Updates a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was updated or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal ComplaintID As Integer, _
                ByVal ComplaintName As String, _
                ByVal Complaint As String, _
                ByVal Recommendation2 As String, _
                ByVal ExaminerID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ComplaintsFields.fldID) = New SqlParameter("@ComplaintID", SqlDbType.Int)
        arParameters(Me.ComplaintsFields.fldID).Value = ComplaintID
        arParameters(Me.ComplaintsFields.fldComplaintName) = New SqlParameter("@ComplaintName", SqlDbType.NVarChar, 100)
        arParameters(Me.ComplaintsFields.fldComplaintName).Value = ComplaintName
        arParameters(Me.ComplaintsFields.fldComplaint) = New SqlParameter("@Complaint", SqlDbType.VarChar, 8000)
        arParameters(Me.ComplaintsFields.fldComplaint).Value = Complaint
        arParameters(Me.ComplaintsFields.fldRecommendation2) = New SqlParameter("@Recommendation2", SqlDbType.VarChar, 8000)
        arParameters(Me.ComplaintsFields.fldRecommendation2).Value = Recommendation2
        arParameters(Me.ComplaintsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.ComplaintsFields.fldExaminerID).Value = ExaminerID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spComplaintsUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spComplaintsUpdate", arParameters)
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
    '* ComplaintName:        Add
    '*
    '* Description: Adds a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef ComplaintID As Integer, _
                ByVal ComplaintName As String, _
                ByVal Complaint As String, _
                ByVal Recommendation2 As String, _
                ByVal ExaminerID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ComplaintsFields.fldID) = New SqlParameter("@ComplaintID", SqlDbType.Int)
        arParameters(Me.ComplaintsFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.ComplaintsFields.fldComplaintName) = New SqlParameter("@ComplaintName", SqlDbType.NVarChar, 100)
        arParameters(Me.ComplaintsFields.fldComplaintName).Value = ComplaintName
        arParameters(Me.ComplaintsFields.fldComplaint) = New SqlParameter("@Complaint", SqlDbType.VarChar, 8000)
        arParameters(Me.ComplaintsFields.fldComplaint).Value = Complaint
        arParameters(Me.ComplaintsFields.fldRecommendation2) = New SqlParameter("@Recommendation2", SqlDbType.VarChar, 8000)
        arParameters(Me.ComplaintsFields.fldRecommendation2).Value = Recommendation2
        arParameters(Me.ComplaintsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.ComplaintsFields.fldExaminerID).Value = ExaminerID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spComplaintsInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spComplaintsInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            ComplaintID = CType(arParameters(0).Value, Integer)
            Return True
        End If
    End Function




    '**************************************************************************
    '*  
    '* ComplaintName:        Delete
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
        arParameters(0) = New SqlParameter("@ComplaintID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spComplaintsDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spComplaintsDelete", arParameters)
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


End Class 'dalComplaints
