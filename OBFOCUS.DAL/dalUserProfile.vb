
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalUserProfile
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
Public Class dalUserProfile

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 



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
    '* Name:        GetbyKey
    '*
    '* Description: GetbyKeys a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was GetbyKeyed or not. 
    '*              True (record GetbyKeyed); False (otherwise).
    '*
    '**************************************************************************
    Public Function GetbyKey(ByVal ID As String, _
                ByRef FirstName As String, _
                ByRef LastName As String, _
                ByRef Position As String, _
                ByRef DOB As String, _
                ByRef PhysicianID As Integer, _
                ByRef ExaminerID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(6) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@UserID", SqlDbType.NVarChar, 50)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@FirstName", SqlDbType.NVarChar, 50)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@LastName", SqlDbType.NVarChar, 50)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@Position", SqlDbType.NVarChar, 100)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(6).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spGetUserProfile", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spGetUserProfile", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            FirstName = ProcessNull.GetString(arParameters(1).Value)
            LastName = ProcessNull.GetString(arParameters(2).Value)
            Position = ProcessNull.GetString(arParameters(3).Value)
            DOB = ProcessNull.GetString(arParameters(4).Value)
            PhysicianID = ProcessNull.GetInt32(arParameters(5).Value)
            ExaminerID = ProcessNull.GetInt32(arParameters(6).Value)
            Return True
        End If
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
    Public Function Update(ByVal ID As String, _
               ByVal FirstName As String, _
                ByVal LastName As String, _
                ByVal Position As String, _
                ByVal DOB As String, _
                ByVal PhysicianID As Integer, _
                ByVal ExaminerID As Integer) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(6) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@UserID", SqlDbType.NVarChar, 50)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@FirstName", SqlDbType.NVarChar, 50)
        arParameters(1).Value = FirstName
        arParameters(2) = New SqlParameter("@LastName", SqlDbType.NVarChar, 50)
        arParameters(2).Value = LastName
        arParameters(3) = New SqlParameter("@Position", SqlDbType.NVarChar, 100)
        arParameters(3).Value = Position
        arParameters(4) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        If IsDate(DOB) Then
            arParameters(4).Value = DOB
        Else
            arParameters(4).Value = DBNull.Value
        End If
        arParameters(5) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        If IsNumeric(PhysicianID) Then
            arParameters(5).Value = PhysicianID
        Else
            arParameters(5).Value = DBNull.Value
        End If
        arParameters(6) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        If IsNumeric(ExaminerID) Then
            arParameters(6).Value = ExaminerID
        Else
            arParameters(6).Value = DBNull.Value
        End If
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUserProfileUpd", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUserProfileUpd", arParameters)
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
    Public Function Delete(ByVal ID As String) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@UserID", SqlDbType.NVarChar, 50)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUserProfileDel", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUserProfileDel", arParameters)
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


End Class 'dalUserProfile
