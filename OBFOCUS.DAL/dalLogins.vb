
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalLogins
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
Public Class dalLogins

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



#Region "Main procedures - GetLogins, Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Name:        GetLogins
    '*
    '* Description: Returns all records in the [syslogin] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetLogins() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spUserLoginsGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spUserLoginsGet")
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
    Public Function Update(ByVal UserLogin As String, ByVal OldPassword As String, ByVal NewPassword As String) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters

        arParameters(0) = New SqlParameter("@Name", SqlDbType.NVarChar, 256)
        arParameters(0).Value = UserLogin
        arParameters(1) = New SqlParameter("@OldPassword", SqlDbType.NVarChar, 256)
        arParameters(1).Value = OldPassword
        arParameters(2) = New SqlParameter("@NewPassword", SqlDbType.NVarChar, 256)
        arParameters(2).Value = NewPassword

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUserLoginsPWUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUserLoginsPWUpdate", arParameters)
            End If

            MsgBox("Password has been reset for " & UserLogin & ".", MsgBoxStyle.Information, "Task Complete")
        Catch exception As Exception
            Select Case exception.Message
                Case "Old (current) password incorrect for user. The password was not changed."
                    MsgBox("Old (current) password incorrect for user. The password was not changed.", MsgBoxStyle.Critical, "Task Aborted")
                Case Else
                    ExceptionManager.Publish(exception)
            End Select
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
    '* Name:        GetRole
    '*
    '* Description: Test Connection String
    '*
    '*
    '* Returns:     Boolean indicating if record was updated or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function GetRole() As Boolean
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0
        arParameters(0) = New SqlParameter("@UserName", SqlDbType.NVarChar, 256)
        arParameters(0).Value = Globals.UserName
        arParameters(1) = New SqlParameter("@Role", SqlDbType.NVarChar, 256)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(3).Direction = ParameterDirection.Output
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spRoleGet", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spRoleGet", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            Globals.UserRole = ProcessNull.GetString(arParameters(1).Value)
            Globals.LimPhysicianID = ProcessNull.GetInt32(arParameters(2).Value)
            Globals.UserExaminerID = ProcessNull.GetInt32(arParameters(3).Value)
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
    Public Function Add(ByVal UserLogin As String, ByVal NewPassword As String, ByVal Role As String) As Boolean

        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@Name", SqlDbType.NVarChar, 256)
        arParameters(0).Value = UserLogin
        arParameters(1) = New SqlParameter("@NewPassword", SqlDbType.NVarChar, 256)
        arParameters(1).Value = NewPassword
        arParameters(2) = New SqlParameter("@Role", SqlDbType.NVarChar, 256)
        arParameters(2).Value = Role

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUserLoginsAdd", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUserLoginsAdd", arParameters)
            End If
        Catch exception As Exception
            If Left(exception.Message, 29) = "User does not have permission" Then
                MessageBox.Show("User does not have permission to perform this action.  Please review or contact your administrator for more technical support.", "Task Aborted", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Else
                ExceptionManager.Publish(exception)
            End If
        End Try


        ' Return False if data was not found.
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
    Public Function Delete(ByVal UserLogin As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@Name", SqlDbType.NVarChar, 256)
        arParameters(0).Value = UserLogin

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUserLoginsDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUserLoginsDelete", arParameters)
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


End Class 'dalLogins
