
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalAnatomy
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
Public Class dalAnatomy

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum DescriptionsFields
        fldID = 0
        fldName = 1
        fldDescription = 2
        fldExaminerID = 3
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
    Public Function GetByKey(ByVal AnatomyID As Integer, _
                ByRef Name As String, _
                ByRef Description As String, _
                ByRef ExaminerID As Integer) As Boolean
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.DescriptionsFields.fldID) = New SqlParameter("@AnatomyID", SqlDbType.Int)
        arParameters(Me.DescriptionsFields.fldID).Value = AnatomyID
        arParameters(Me.DescriptionsFields.fldName) = New SqlParameter("@Name", SqlDbType.NVarChar, 100)
        arParameters(Me.DescriptionsFields.fldName).Direction = ParameterDirection.Output
        arParameters(Me.DescriptionsFields.fldDescription) = New SqlParameter("@Description", SqlDbType.VarChar, 8000)
        arParameters(Me.DescriptionsFields.fldDescription).Direction = ParameterDirection.Output
        arParameters(Me.DescriptionsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.DescriptionsFields.fldExaminerID).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAnatomyGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAnatomyGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.DescriptionsFields.fldName).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Name = ProcessNull.GetString(arParameters(Me.DescriptionsFields.fldName).Value)
            Description = ProcessNull.GetString(arParameters(Me.DescriptionsFields.fldDescription).Value)
            ExaminerID = ProcessNull.GetInt32(arParameters(Me.DescriptionsFields.fldExaminerID).Value)
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
    Public Function Update(ByVal AnatomyID As Integer, _
                ByVal Name As String, _
                ByVal Description As String, _
                   ByVal ExaminerID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.DescriptionsFields.fldID) = New SqlParameter("@AnatomyID", SqlDbType.Int)
        arParameters(Me.DescriptionsFields.fldID).Value = AnatomyID
        arParameters(Me.DescriptionsFields.fldName) = New SqlParameter("@Name", SqlDbType.NVarChar, 100)
        arParameters(Me.DescriptionsFields.fldName).Value = Name
        arParameters(Me.DescriptionsFields.fldDescription) = New SqlParameter("@Description", SqlDbType.VarChar, 8000)
        arParameters(Me.DescriptionsFields.fldDescription).Value = Description
        arParameters(Me.DescriptionsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.DescriptionsFields.fldExaminerID).Value = ExaminerID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAnatomyUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAnatomyUpdate", arParameters)
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
    Public Function Add(ByRef AnatomyID As Integer, _
                ByVal Name As String, _
                ByVal Description As String, _
                ByVal ExaminerID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.DescriptionsFields.fldID) = New SqlParameter("@AnatomyID", SqlDbType.Int)
        arParameters(Me.DescriptionsFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.DescriptionsFields.fldName) = New SqlParameter("@Name", SqlDbType.NVarChar, 100)
        arParameters(Me.DescriptionsFields.fldName).Value = Name
        arParameters(Me.DescriptionsFields.fldDescription) = New SqlParameter("@Description", SqlDbType.VarChar, 8000)
        arParameters(Me.DescriptionsFields.fldDescription).Value = Description
        arParameters(Me.DescriptionsFields.fldExaminerID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.DescriptionsFields.fldExaminerID).Value = ExaminerID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAnatomyInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAnatomyInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            AnatomyID = CType(arParameters(0).Value, Integer)
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
        arParameters(0) = New SqlParameter("@AnatomyID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAnatomyDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAnatomyDelete", arParameters)
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


End Class 'dalAnatomy
