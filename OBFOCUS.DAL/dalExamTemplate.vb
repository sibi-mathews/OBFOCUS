
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalExamTemplate
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
Public Class dalExamTemplate

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



#Region "Main procedures - GetExamTemplate, Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Name:        GetExamTemplate
    '*
    '* Description: Returns all records in the [ZTables] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetExamTemplate(ByVal ExaminerID As Integer) As SqlDataReader

        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(0).Value = ExaminerID
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spExamTemplateGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spExamTemplateGet", arParameters)
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
    Public Function GetByKey(ByVal ID As Integer, _
                ByRef ExaminerID As Integer, _
                ByRef Subtype As String, _
                ByRef Appearance As String, _
                ByRef HEENT As String, _
                ByRef Neck As String, _
                ByRef Heart As String, _
                ByRef Lung As String, _
                ByRef Back As String, _
                ByRef Abdomen As String, _
                ByRef Extremities As String, _
                ByRef Pelvic As String, _
                ByRef ROS As String, _
                ByRef Neurologic As String) As Boolean
        Dim arParameters(13) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamTemplateID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@SubType", SqlDbType.NVarChar, 100)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@Appearance", SqlDbType.NVarChar, 100)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@HEENT", SqlDbType.NVarChar, 100)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@Neck", SqlDbType.NVarChar, 100)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@Heart", SqlDbType.NVarChar, 100)
        arParameters(6).Direction = ParameterDirection.Output
        arParameters(7) = New SqlParameter("@Lung", SqlDbType.NVarChar, 100)
        arParameters(7).Direction = ParameterDirection.Output
        arParameters(8) = New SqlParameter("@Back", SqlDbType.NVarChar, 100)
        arParameters(8).Direction = ParameterDirection.Output
        arParameters(9) = New SqlParameter("@Abdomen", SqlDbType.NVarChar, 100)
        arParameters(9).Direction = ParameterDirection.Output
        arParameters(10) = New SqlParameter("@Extremities", SqlDbType.NVarChar, 100)
        arParameters(10).Direction = ParameterDirection.Output
        arParameters(11) = New SqlParameter("@Pelvic", SqlDbType.NVarChar, 100)
        arParameters(11).Direction = ParameterDirection.Output
        arParameters(12) = New SqlParameter("@ROS", SqlDbType.NVarChar, 100)
        arParameters(12).Direction = ParameterDirection.Output
        arParameters(13) = New SqlParameter("@Neurologic", SqlDbType.NVarChar, 50)
        arParameters(13).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamTemplateGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamTemplateGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(1).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            ExaminerID = ProcessNull.GetInt32(arParameters(1).Value)
            Subtype = ProcessNull.GetString(arParameters(2).Value)
            Appearance = ProcessNull.GetString(arParameters(3).Value)
            HEENT = ProcessNull.GetString(arParameters(4).Value)
            Neck = ProcessNull.GetString(arParameters(5).Value)
            Heart = ProcessNull.GetString(arParameters(6).Value)
            Lung = ProcessNull.GetString(arParameters(7).Value)
            Back = ProcessNull.GetString(arParameters(8).Value)
            Abdomen = ProcessNull.GetString(arParameters(9).Value)
            Extremities = ProcessNull.GetString(arParameters(10).Value)
            Pelvic = ProcessNull.GetString(arParameters(11).Value)
            ROS = ProcessNull.GetString(arParameters(12).Value)
            Neurologic = ProcessNull.GetString(arParameters(13).Value)
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
    '* Description: Gets all the values of a record identified by a key.
    '*
    '* Parameters:  Description - Output parameter
    '*              Picture - Output parameter
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal ID As Integer, _
                ByVal ExaminerID As Integer, _
                ByVal Subtype As String, _
                ByVal Appearance As String, _
                ByVal HEENT As String, _
                ByVal Neck As String, _
                ByVal Heart As String, _
                ByVal Lung As String, _
                ByVal Back As String, _
                ByVal Abdomen As String, _
                ByVal Extremities As String, _
                ByVal Pelvic As String, _
                ByVal ROS As String, _
                ByVal Neurologic As String) As Boolean
        Dim arParameters(13) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamTemplateID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(1).Value = ExaminerID
        arParameters(2) = New SqlParameter("@SubType", SqlDbType.NVarChar, 100)
        arParameters(2).Value = Subtype
        arParameters(3) = New SqlParameter("@Appearance", SqlDbType.NVarChar, 100)
        arParameters(3).Value = Appearance
        arParameters(4) = New SqlParameter("@HEENT", SqlDbType.NVarChar, 100)
        arParameters(4).Value = HEENT
        arParameters(5) = New SqlParameter("@Neck", SqlDbType.NVarChar, 100)
        arParameters(5).Value = Neck
        arParameters(6) = New SqlParameter("@Heart", SqlDbType.NVarChar, 100)
        arParameters(6).Value = Heart
        arParameters(7) = New SqlParameter("@Lung", SqlDbType.NVarChar, 100)
        arParameters(7).Value = Lung
        arParameters(8) = New SqlParameter("@Back", SqlDbType.NVarChar, 100)
        arParameters(8).Value = Back
        arParameters(9) = New SqlParameter("@Abdomen", SqlDbType.NVarChar, 100)
        arParameters(9).Value = Abdomen
        arParameters(10) = New SqlParameter("@Extremities", SqlDbType.NVarChar, 100)
        arParameters(10).Value = Extremities
        arParameters(11) = New SqlParameter("@Pelvic", SqlDbType.NVarChar, 100)
        arParameters(11).Value = Pelvic
        arParameters(12) = New SqlParameter("@ROS", SqlDbType.NVarChar, 100)
        arParameters(12).Value = ROS
        arParameters(13) = New SqlParameter("@Neurologic", SqlDbType.NVarChar, 50)
        arParameters(13).Value = Neurologic

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamTemplateUpdate", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamTemplateUpdate", arParameters)
            End If
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


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
                ByVal ExaminerID As Integer, _
                ByVal Subtype As String, _
                ByVal Appearance As String, _
                ByVal HEENT As String, _
                ByVal Neck As String, _
                ByVal Heart As String, _
                ByVal Lung As String, _
                ByVal Back As String, _
                ByVal Abdomen As String, _
                ByVal Extremities As String, _
                ByVal Pelvic As String, _
                ByVal ROS As String, _
                ByVal Neurologic As String) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(13) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamTemplateID", SqlDbType.Int)
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(1).Value = ExaminerID
        arParameters(2) = New SqlParameter("@SubType", SqlDbType.NVarChar, 100)
        arParameters(2).Value = Subtype
        arParameters(3) = New SqlParameter("@Appearance", SqlDbType.NVarChar, 100)
        arParameters(3).Value = Appearance
        arParameters(4) = New SqlParameter("@HEENT", SqlDbType.NVarChar, 100)
        arParameters(4).Value = HEENT
        arParameters(5) = New SqlParameter("@Neck", SqlDbType.NVarChar, 100)
        arParameters(5).Value = Neck
        arParameters(6) = New SqlParameter("@Heart", SqlDbType.NVarChar, 100)
        arParameters(6).Value = Heart
        arParameters(7) = New SqlParameter("@Lung", SqlDbType.NVarChar, 100)
        arParameters(7).Value = Lung
        arParameters(8) = New SqlParameter("@Back", SqlDbType.NVarChar, 100)
        arParameters(8).Value = Back
        arParameters(9) = New SqlParameter("@Abdomen", SqlDbType.NVarChar, 100)
        arParameters(9).Value = Abdomen
        arParameters(10) = New SqlParameter("@Extremities", SqlDbType.NVarChar, 100)
        arParameters(10).Value = Extremities
        arParameters(11) = New SqlParameter("@Pelvic", SqlDbType.NVarChar, 100)
        arParameters(11).Value = Pelvic
        arParameters(12) = New SqlParameter("@ROS", SqlDbType.NVarChar, 100)
        arParameters(12).Value = ROS
        arParameters(13) = New SqlParameter("@Neurologic", SqlDbType.NVarChar, 50)
        arParameters(13).Value = Neurologic

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamTemplateInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamTemplateInsert", arParameters)
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
        arParameters(0) = New SqlParameter("@ExamTemplateID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamTemplateDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamTemplateDelete", arParameters)
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


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class 'dalExamTemplate
