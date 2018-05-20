
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalCompany
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
Public Class dalCompany

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



#Region "Main procedures - GetCompany, Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Name:        GetCompany
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
    Public Function GetCompany(ByRef CompanyID As Integer, _
                ByRef CompanyName As String, _
                ByRef Address1 As String, _
                ByRef Address2 As String, _
                ByRef City As String, _
                ByRef State As String, _
                ByRef Zip As String, _
                ByRef MainPhone As String, _
                ByRef Email As String, _
                ByRef MainContact As String, _
                ByRef MainContactTitle As String, _
                ByRef MainContactPhone As String, _
                ByRef MainContactEmail As String) As Boolean

        Dim arParameters(12) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@CompanyID", SqlDbType.Int)
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@CompanyName", SqlDbType.NVarChar, 255)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@Address1", SqlDbType.NVarChar, 255)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@Address2", SqlDbType.NVarChar, 255)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@City", SqlDbType.NVarChar, 100)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@State", SqlDbType.NVarChar, 10)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@Zip", SqlDbType.NVarChar, 15)
        arParameters(6).Direction = ParameterDirection.Output
        arParameters(7) = New SqlParameter("@MainPhone", SqlDbType.NVarChar, 15)
        arParameters(7).Direction = ParameterDirection.Output
        arParameters(8) = New SqlParameter("@Email", SqlDbType.NVarChar, 100)
        arParameters(8).Direction = ParameterDirection.Output
        arParameters(9) = New SqlParameter("@MainContact", SqlDbType.NVarChar, 100)
        arParameters(9).Direction = ParameterDirection.Output
        arParameters(10) = New SqlParameter("@MainContactTitle", SqlDbType.NVarChar, 100)
        arParameters(10).Direction = ParameterDirection.Output
        arParameters(11) = New SqlParameter("@MainContactPhone", SqlDbType.NVarChar, 15)
        arParameters(11).Direction = ParameterDirection.Output
        arParameters(12) = New SqlParameter("@MainContactEmail", SqlDbType.NVarChar, 100)
        arParameters(12).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spCompanyGet", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spCompanyGet", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(0).Value Is DBNull.Value Then Return False
            ' Return True if data was found. Also populate output (ByRef) parameters.
            CompanyID = ProcessNull.GetInt32(arParameters(0).Value)
            CompanyName = ProcessNull.GetString(arParameters(1).Value)
            Address1 = ProcessNull.GetString(arParameters(2).Value)
            Address2 = ProcessNull.GetString(arParameters(3).Value)
            City = ProcessNull.GetString(arParameters(4).Value)
            State = ProcessNull.GetString(arParameters(5).Value)
            Zip = ProcessNull.GetString(arParameters(6).Value)
            MainPhone = ProcessNull.GetString(arParameters(7).Value)
            Email = ProcessNull.GetString(arParameters(8).Value)
            MainContact = ProcessNull.GetString(arParameters(9).Value)
            MainContactTitle = ProcessNull.GetString(arParameters(10).Value)
            MainContactPhone = ProcessNull.GetString(arParameters(11).Value)
            MainContactEmail = ProcessNull.GetString(arParameters(12).Value)
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
    '* Description: Adds a new record to the [Ultrasound] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef CompanyID As Integer, _
                ByVal CompanyName As String, _
                ByVal Address1 As String, _
                ByVal Address2 As String, _
                ByVal City As String, _
                ByVal State As String, _
                ByVal Zip As String, _
                ByVal MainPhone As String, _
                ByVal Email As String, _
                ByVal MainContact As String, _
                ByVal MainContactTitle As String, _
                ByVal MainContactPhone As String, _
                ByVal MainContactEmail As String) As Boolean

        Dim arParameters(12) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@CompanyID", SqlDbType.Int)
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@CompanyName", SqlDbType.NVarChar, 255)
        arParameters(1).Value = CompanyName
        arParameters(2) = New SqlParameter("@Address1", SqlDbType.NVarChar, 255)
        arParameters(2).Value = Address1
        arParameters(3) = New SqlParameter("@Address2", SqlDbType.NVarChar, 255)
        arParameters(3).Value = Address2
        arParameters(4) = New SqlParameter("@City", SqlDbType.NVarChar, 100)
        arParameters(4).Value = City
        arParameters(5) = New SqlParameter("@State", SqlDbType.NVarChar, 10)
        arParameters(5).Value = State
        arParameters(6) = New SqlParameter("@Zip", SqlDbType.NVarChar, 15)
        arParameters(6).Value = Zip
        arParameters(7) = New SqlParameter("@MainPhone", SqlDbType.NVarChar, 15)
        arParameters(7).Value = MainPhone
        arParameters(8) = New SqlParameter("@Email", SqlDbType.NVarChar, 100)
        arParameters(8).Value = Email
        arParameters(9) = New SqlParameter("@MainContact", SqlDbType.NVarChar, 100)
        arParameters(9).Value = MainContact
        arParameters(10) = New SqlParameter("@MainContactTitle", SqlDbType.NVarChar, 100)
        arParameters(10).Value = MainContactTitle
        arParameters(11) = New SqlParameter("@MainContactPhone", SqlDbType.NVarChar, 15)
        arParameters(11).Value = MainContactPhone
        arParameters(12) = New SqlParameter("@MainContactEmail", SqlDbType.NVarChar, 100)
        arParameters(12).Value = MainContactEmail
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spCompanyInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spCompanyInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            CompanyID = CType(arParameters(0).Value, Integer)
            Return True
        End If

    End Function

    '**************************************************************************
    '*  
    '* Name:        Delete
    '*
    '* Description: Deletes a record from the [Ultrasound] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to delete
    '*
    '* Returns:     Boolean indicating if record was deleted or not. 
    '*              True (record found and deleted); False (otherwise).
    '*
    '**************************************************************************
    Public Function Delete(ByVal CompanyID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@CompanyID", SqlDbType.Int)
        arParameters(0).Value = CompanyID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spCompanyDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spCompanyDelete", arParameters)
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
    '* Name:        Update
    '*
    '* Description: Updates a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(Byval CompanyID As Integer, _
                ByVal CompanyName As String, _
                ByVal Address1 As String, _
                ByVal Address2 As String, _
                ByVal City As String, _
                ByVal State As String, _
                ByVal Zip As String, _
                ByVal MainPhone As String, _
                ByVal Email As String, _
                ByVal MainContact As String, _
                ByVal MainContactTitle As String, _
                ByVal MainContactPhone As String, _
                ByVal MainContactEmail As String) As Boolean

        Dim arParameters(12) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@CompanyID", SqlDbType.Int)
        arParameters(0).Value = CompanyID
        arParameters(1) = New SqlParameter("@CompanyName", SqlDbType.NVarChar, 255)
        arParameters(1).Value = CompanyName
        arParameters(2) = New SqlParameter("@Address1", SqlDbType.NVarChar, 255)
        arParameters(2).Value = Address1
        arParameters(3) = New SqlParameter("@Address2", SqlDbType.NVarChar, 255)
        arParameters(3).Value = Address2
        arParameters(4) = New SqlParameter("@City", SqlDbType.NVarChar, 100)
        arParameters(4).Value = City
        arParameters(5) = New SqlParameter("@State", SqlDbType.NVarChar, 10)
        arParameters(5).Value = State
        arParameters(6) = New SqlParameter("@Zip", SqlDbType.NVarChar, 15)
        arParameters(6).Value = Zip
        arParameters(7) = New SqlParameter("@MainPhone", SqlDbType.NVarChar, 15)
        arParameters(7).Value = MainPhone
        arParameters(8) = New SqlParameter("@Email", SqlDbType.NVarChar, 100)
        arParameters(8).Value = Email
        arParameters(9) = New SqlParameter("@MainContact", SqlDbType.NVarChar, 100)
        arParameters(9).Value = MainContact
        arParameters(10) = New SqlParameter("@MainContactTitle", SqlDbType.NVarChar, 100)
        arParameters(10).Value = MainContactTitle
        arParameters(11) = New SqlParameter("@MainContactPhone", SqlDbType.NVarChar, 15)
        arParameters(11).Value = MainContactPhone
        arParameters(12) = New SqlParameter("@MainContactEmail", SqlDbType.NVarChar, 100)
        arParameters(12).Value = MainContactEmail
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spCompanyUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spCompanyUpdate", arParameters)
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


End Class 'dalCompany
