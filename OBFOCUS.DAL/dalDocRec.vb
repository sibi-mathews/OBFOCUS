
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalDocRec
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
Public Class dalDocRec

#Region "Module level variables and enums"


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



#Region "Main procedures - Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Name:        GetDocRec
    '*
    '* Description: Returns all records in the [Labs] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetDocRec(ByVal ChartID As Integer, Optional ByVal LabOnly As Short = 0) As SqlDataReader
        Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ChartID
        arParameters(1) = New SqlParameter("@LabOnly", SqlDbType.Bit)
        arParameters(1).Value = LabOnly
        ' Letter stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spChartDocRecGetByKey", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spChartDocRecGetByKey", arParameters)
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
                       ByVal DocumentPath As String, _
                        ByVal DocumentDescrip As String, _
                       ByVal DateRecorded As String, _
                       ByVal UserID As String, _
                       ByVal DocumentReviewed As Boolean, _
                       ByVal ReviewedBy As String, _
                       ByVal ReviewedDate As String, _
                       ByVal ReviewComments As String, _
                       ByVal LabSiteID As Integer, _
                       ByVal isLab As Boolean, _
                       ByVal DocRecLabTypeID As Integer, _
                       ByVal DocRecStatID As Integer) As Boolean

        Dim arParameters(12) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@DocumentPath", SqlDbType.VarChar, 500)
        arParameters(1).Value = DocumentPath
        arParameters(2) = New SqlParameter("@DocumentDescrip", SqlDbType.VarChar, 255)
        arParameters(2).Value = DocumentDescrip
        arParameters(3) = New SqlParameter("@DateRecorded", SqlDbType.DateTime)
        If DateRecorded = "" Then
            arParameters(3).Value = DBNull.Value
        Else
            arParameters(3).Value = DateRecorded
        End If
        arParameters(4) = New SqlParameter("@UserID", SqlDbType.VarChar, 100)
        arParameters(4).Value = UserID
        arParameters(5) = New SqlParameter("@DocumentReviewed", SqlDbType.Bit)
        If DocumentReviewed = False Then
            arParameters(5).Value = 0
        Else
            arParameters(5).Value = 1
        End If
        arParameters(6) = New SqlParameter("@ReviewedBy", SqlDbType.VarChar, 100)
        arParameters(6).Value = ReviewedBy
        arParameters(7) = New SqlParameter("@ReviewedDate", SqlDbType.DateTime)
        If ReviewedDate = "" Then
            arParameters(7).Value = DBNull.Value
        Else
            arParameters(7).Value = ReviewedDate
        End If
        arParameters(8) = New SqlParameter("@ReviewComments", SqlDbType.VarChar, 255)
        arParameters(8).Value = ReviewComments
        arParameters(9) = New SqlParameter("@LabSiteID", SqlDbType.Int)
        arParameters(9).Value = LabSiteID
        arParameters(10) = New SqlParameter("@isLab", SqlDbType.Bit)
        If isLab = False Then
            arParameters(10).Value = 0
        Else
            arParameters(10).Value = 1
        End If
        arParameters(11) = New SqlParameter("@DocRecLabTypeID", SqlDbType.Int)
        arParameters(11).Value = DocRecLabTypeID
        arParameters(12) = New SqlParameter("@DocRecStatID", SqlDbType.Int)
        arParameters(12).Value = DocRecStatID
        ' Letter stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartDocRecUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartDocRecUpdate", arParameters)
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
                        ByVal DocumentPath As String, _
                        ByVal DocumentDescrip As String, _
                        ByVal ChartID As Integer, _
                        ByVal IsLab As Boolean, _
                        ByVal LabSiteID As Integer, _
                        ByVal ExamID As Integer, _
                        ByVal ExamDate As String, _
                        ByVal DocRecLabTypeID As Integer, _
                        ByVal DocumentReviewed As Boolean, _
                        ByVal ReviewedBy As String, _
                        ByVal ReviewComments As String, _
                        ByVal ExaminerID As Integer) As Boolean

        Dim arParameters(13) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@DocumentPath", SqlDbType.VarChar, 500)
        arParameters(1).Value = DocumentPath
        arParameters(2) = New SqlParameter("@DocumentDescrip", SqlDbType.VarChar, 255)
        arParameters(2).Value = DocumentDescrip
        arParameters(3) = New SqlParameter("@UserID", SqlDbType.VarChar, 50)
        arParameters(3).Value = Globals.UserName
        arParameters(4) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(4).Value = ChartID
        arParameters(5) = New SqlParameter("@IsLab", SqlDbType.Bit)
        If IsLab = False Then
            arParameters(5).Value = 0
        Else
            arParameters(5).Value = 1
        End If
        arParameters(6) = New SqlParameter("@LabSiteID", SqlDbType.Int)
        arParameters(6).Value = LabSiteID
        arParameters(7) = New SqlParameter("@ExamID", SqlDbType.Int)
        If ExamID <> 0 Then
            arParameters(7).Value = ExamID
        Else
            arParameters(7).Value = DBNull.Value
        End If
        arParameters(8) = New SqlParameter("@ExamDate", SqlDbType.DateTime)
        If Len(ExamDate) > 0 And IsDate(ExamDate) = True Then
            arParameters(8).Value = ExamDate
        Else
            arParameters(8).Value = DBNull.Value
        End If
        arParameters(9) = New SqlParameter("@DocRecLabTypeID", SqlDbType.Int)
        arParameters(9).Value = DocRecLabTypeID
        arParameters(10) = New SqlParameter("@DocumentReviewed", SqlDbType.Bit)
        arParameters(11) = New SqlParameter("@ReviewedBy", SqlDbType.VarChar, 100)
        arParameters(12) = New SqlParameter("@ReviewComments", SqlDbType.VarChar, 255)
        If DocumentReviewed = False Then
            arParameters(10).Value = 0
            arParameters(11).Value = DBNull.Value
            arParameters(12).Value = DBNull.Value
        Else
            arParameters(10).Value = 1
            arParameters(11).Value = ReviewedBy
            arParameters(12).Value = ReviewComments
        End If
        arParameters(13) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        If ExaminerID <> 0 Then
            arParameters(13).Value = ExaminerID
        Else
            arParameters(13).Value = DBNull.Value
        End If
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartDocRecInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartDocRecInsert", arParameters)
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
    Public Function Delete(ByVal ID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Letter stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartDocRecDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartDocRecDelete", arParameters)
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


End Class 'dalDocRec
