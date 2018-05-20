
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalExaminer2
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
Public Class dalExaminer2

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum Examiner2Fields
        fldID = 0
        fldExaminer2Name = 1
        fldPassword = 2
        fldSuppress = 3

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
    Public Function GetByKey(ByVal Examiner2ID As Integer, _
                ByRef Examiner2Name As String, _
                ByRef Suppress As Short) As Boolean
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.Examiner2Fields.fldID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.Examiner2Fields.fldID).Value = Examiner2ID
        arParameters(Me.Examiner2Fields.fldExaminer2Name) = New SqlParameter("@Name", SqlDbType.NVarChar, 255)
        arParameters(Me.Examiner2Fields.fldExaminer2Name).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@Suppress", SqlDbType.Bit)
        arParameters(2).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminer2GetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExaminer2GetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.Examiner2Fields.fldExaminer2Name).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Examiner2Name = ProcessNull.GetString(arParameters(Me.Examiner2Fields.fldExaminer2Name).Value)
            Suppress = ProcessNull.GetInt16(arParameters(2).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        GetExaminerPassword
    '*
    '* Description: Gets all the values of a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function GetExaminerPassword(ByVal ExamID As Integer, _
                ByRef ID As Integer, _
                ByRef ExaminerName As String, _
                ByRef Password As String) As Boolean
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ExamID
        arParameters(1) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@Name", SqlDbType.NVarChar, 255)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@Password", SqlDbType.NVarChar, 255)
        arParameters(3).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminerPasswordGet", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExaminerPasswordGet", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.Examiner2Fields.fldExaminer2Name).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            ID = ProcessNull.GetInt32(arParameters(1).Value)
            ExaminerName = ProcessNull.GetString(arParameters(2).Value)
            Password = ProcessNull.GetString(arParameters(3).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        GetExaminerPasswordLim
    '*
    '* Description: Gets all the values of a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function GetExaminerPasswordLim(ByVal ID As Integer, _
                ByRef Password As String) As Boolean
        Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Password", SqlDbType.NVarChar, 255)
        arParameters(1).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminerPasswordGetLim", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExaminerPasswordGetLim", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(0).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Password = ProcessNull.GetString(arParameters(1).Value)
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
    Public Function Update(ByVal Examiner2ID As Integer, _
                ByVal Examiner2Name As String, _
                ByVal Suppress As Short) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.Examiner2Fields.fldID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.Examiner2Fields.fldID).Value = Examiner2ID
        arParameters(Me.Examiner2Fields.fldExaminer2Name) = New SqlParameter("@Name", SqlDbType.NVarChar, 255)
        arParameters(Me.Examiner2Fields.fldExaminer2Name).Value = Examiner2Name
        arParameters(2) = New SqlParameter("@Suppress", SqlDbType.Bit)
        arParameters(2).Value = Suppress
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminer2Update", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExaminer2Update", arParameters)
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
    '* Name:        UpdateSigned
    '*
    '* Description: UpdateSigneds a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was UpdateSignedd or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdateSigned(ByVal ExamID As Integer, _
                ByVal Signed As String, _
                ByVal SignedBy As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.Examiner2Fields.fldID) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(Me.Examiner2Fields.fldID).Value = ExamID
        arParameters(Me.Examiner2Fields.fldExaminer2Name) = New SqlParameter("@Signed", SqlDbType.NVarChar, 50)
        arParameters(Me.Examiner2Fields.fldExaminer2Name).Value = Signed
        arParameters(2) = New SqlParameter("@SignedBy", SqlDbType.NVarChar, 50)
        arParameters(2).Value = SignedBy

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsSignedUpd", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamsSignedUpd", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not UpdateSignedd.
        If intRecordsAffected = 0 Then
            Return False
        Else

            Return True
        End If

    End Function
    '**************************************************************************
    '*  
    '* Name:        UpdatePassword
    '*
    '* Description: UpdatePasswords a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was UpdatePasswordd or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdatePassword(ByVal ExaminerID As Integer, _
                ByVal Password As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.Examiner2Fields.fldID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.Examiner2Fields.fldID).Value = ExaminerID
        arParameters(Me.Examiner2Fields.fldExaminer2Name) = New SqlParameter("@Password", SqlDbType.NVarChar, 50)
        arParameters(Me.Examiner2Fields.fldExaminer2Name).Value = Password

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminerPasswordChange", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExaminerPasswordChange", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not UpdatePasswordd.
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
    Public Function Add(ByRef Examiner2ID As Integer, _
                ByVal Examiner2Name As String, _
                ByVal Suppress As Short) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.Examiner2Fields.fldID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.Examiner2Fields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.Examiner2Fields.fldExaminer2Name) = New SqlParameter("@Name", SqlDbType.NVarChar, 255)
        arParameters(Me.Examiner2Fields.fldExaminer2Name).Value = Examiner2Name
        arParameters(2) = New SqlParameter("@Suppress", SqlDbType.Bit)
        arParameters(2).Value = Suppress

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminer2Insert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExaminer2Insert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            Examiner2ID = CType(arParameters(0).Value, Integer)
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
        arParameters(0) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminer2Delete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExaminer2Delete", arParameters)
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


End Class 'dalExaminer2
