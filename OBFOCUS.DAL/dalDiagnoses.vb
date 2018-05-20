
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalDiagnoses
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
Public Class dalDiagnoses

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum DiagnosesFields
        fldID = 0
        fldDiagnosis = 1
        fldBillingCode = 2
        fldServiceType = 3
        fldSuppress = 4
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
    Public Function GetByKey(ByVal DiagnosisID As Integer, _
                ByRef Diagnosis As String, _
                ByRef BillingCode As String, _
                ByRef ServiceType As String, _
                ByRef Suppress As Short) As Boolean
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.DiagnosesFields.fldID) = New SqlParameter("@DiagnosisID", SqlDbType.Int)
        arParameters(Me.DiagnosesFields.fldID).Value = DiagnosisID
        arParameters(Me.DiagnosesFields.fldDiagnosis) = New SqlParameter("@Diagnosis", SqlDbType.NVarChar, 100)
        arParameters(Me.DiagnosesFields.fldDiagnosis).Direction = ParameterDirection.Output
        arParameters(Me.DiagnosesFields.fldBillingCode) = New SqlParameter("@BillingCode", SqlDbType.NVarChar, 50)
        arParameters(Me.DiagnosesFields.fldBillingCode).Direction = ParameterDirection.Output
        arParameters(Me.DiagnosesFields.fldServiceType) = New SqlParameter("@ServiceType", SqlDbType.NVarChar, 50)
        arParameters(Me.DiagnosesFields.fldServiceType).Direction = ParameterDirection.Output
        arParameters(Me.DiagnosesFields.fldSuppress) = New SqlParameter("@Suppress", SqlDbType.Bit)
        arParameters(Me.DiagnosesFields.fldSuppress).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spDiagnosesGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spDiagnosesGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.DiagnosesFields.fldDiagnosis).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Diagnosis = ProcessNull.GetString(arParameters(Me.DiagnosesFields.fldDiagnosis).Value)
            Diagnosis = Diagnosis.Trim()
            BillingCode = ProcessNull.GetString(arParameters(Me.DiagnosesFields.fldBillingCode).Value)
            BillingCode = BillingCode.Trim()
            ServiceType = ProcessNull.GetString(arParameters(Me.DiagnosesFields.fldServiceType).Value)
            ServiceType = ServiceType.Trim()
            Suppress = ProcessNull.GetInt16(arParameters(Me.DiagnosesFields.fldSuppress).Value)

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
    Public Function Update(ByVal DiagnosisID As Integer, _
                ByVal Diagnosis As String, _
                ByVal BillingCode As String, _
                ByVal ServiceType As String, _
                ByVal Suppress As Short) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.DiagnosesFields.fldID) = New SqlParameter("@DiagnosisID", SqlDbType.Int)
        arParameters(Me.DiagnosesFields.fldID).Value = DiagnosisID
        arParameters(Me.DiagnosesFields.fldDiagnosis) = New SqlParameter("@Diagnosis", SqlDbType.NVarChar, 100)
        arParameters(Me.DiagnosesFields.fldDiagnosis).Value = Diagnosis
        arParameters(Me.DiagnosesFields.fldBillingCode) = New SqlParameter("@BillingCode", SqlDbType.NVarChar, 50)
        arParameters(Me.DiagnosesFields.fldBillingCode).Value = BillingCode
        arParameters(Me.DiagnosesFields.fldServiceType) = New SqlParameter("@ServiceType", SqlDbType.NVarChar, 50)
        arParameters(Me.DiagnosesFields.fldServiceType).Value = ServiceType
        arParameters(Me.DiagnosesFields.fldSuppress) = New SqlParameter("@Suppress", SqlDbType.Bit)
        arParameters(Me.DiagnosesFields.fldSuppress).Value = Suppress

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spDiagnosesUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spDiagnosesUpdate", arParameters)
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
    Public Function Add(ByRef DiagnosisID As Integer, _
                ByVal Diagnosis As String, _
                ByVal BillingCode As String, _
                ByVal ServiceType As String, _
                ByVal Suppress As Short) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.DiagnosesFields.fldID) = New SqlParameter("@DiagnosisID", SqlDbType.Int)
        arParameters(Me.DiagnosesFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.DiagnosesFields.fldDiagnosis) = New SqlParameter("@Diagnosis", SqlDbType.NVarChar, 100)
        arParameters(Me.DiagnosesFields.fldDiagnosis).Value = Diagnosis
        arParameters(Me.DiagnosesFields.fldBillingCode) = New SqlParameter("@BillingCode", SqlDbType.NVarChar, 50)
        arParameters(Me.DiagnosesFields.fldBillingCode).Value = BillingCode
        arParameters(Me.DiagnosesFields.fldServiceType) = New SqlParameter("@ServiceType", SqlDbType.NVarChar, 50)
        arParameters(Me.DiagnosesFields.fldServiceType).Value = ServiceType
        arParameters(Me.DiagnosesFields.fldSuppress) = New SqlParameter("@Suppress", SqlDbType.Bit)
        arParameters(Me.DiagnosesFields.fldSuppress).Value = Suppress


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spDiagnosesInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spDiagnosesInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            DiagnosisID = CType(arParameters(0).Value, Integer)
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
        arParameters(Me.DiagnosesFields.fldID) = New SqlParameter("@DiagnosisID", SqlDbType.Int)
        arParameters(Me.DiagnosesFields.fldID).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spDiagnosesDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spDiagnosesDelete", arParameters)
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


End Class 'dalDiagnoses
