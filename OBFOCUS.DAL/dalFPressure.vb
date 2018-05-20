
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalFPressure
'*
'* Class: Data access layer for Table PatientInfo
'*
'* Remarks:     Uses OleDb and embedded SQL for maintaining the data.
'*-----------------------------------------------------------------------------
'*                      CHANGE HISTORY
'*   Change No:   Date:          Author:   Class:
'*   _________    ___________    ______    ____________________________________
'*      001       1/26/2005     MR        Created.                                
'* 
'******************************************************************************
Public Class dalFPressure

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum ClasssFields
        fldID = 0
        fldName = 1
        fldDescription = 2
    End Enum


    'Used for transaction support
    Private _Transaction As SqlTransaction = Nothing


    '**************************************************************************
    '*  
    '* Name:        Transaction
    '*
    '* Class: Used for transaction support.
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
    '* Class: Initialize a new instance of the class.
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
    '* Class: Initialize a new instance of the class.
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
    '* Name:        GetFPressure
    '*
    '* Description: Returns all records in the [Labs] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetFPressure() As SqlDataReader
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spFPressureGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spFPressureGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetByKey
    '*
    '* Class: Gets all the values of a record identified by a key.
    '*
    '* Parameters:  Class - Output parameter
    '*              Picture - Output parameter
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function GetByKey(ByVal FPressureID As Integer, _
                ByRef Name As String, _
                ByRef Description As String) As Boolean
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@FPressureID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Value = FPressureID
        arParameters(Me.ClasssFields.fldName) = New SqlParameter("@Name", SqlDbType.NVarChar, 100)
        arParameters(Me.ClasssFields.fldName).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldDescription) = New SqlParameter("@Description", SqlDbType.VarChar, 100)
        arParameters(Me.ClasssFields.fldDescription).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFPressureGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFPressureGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.ClasssFields.fldName).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Name = ProcessNull.GetString(arParameters(Me.ClasssFields.fldName).Value)
            Description = ProcessNull.GetString(arParameters(Me.ClasssFields.fldDescription).Value)
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
    '* Class: Updates a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was updated or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal FPressureID As Integer, _
                ByVal Name As String, _
                ByVal Description As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@FPressureID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Value = FPressureID
        arParameters(Me.ClasssFields.fldName) = New SqlParameter("@Name", SqlDbType.NVarChar, 100)
        arParameters(Me.ClasssFields.fldName).Value = Name
        arParameters(Me.ClasssFields.fldDescription) = New SqlParameter("@Description", SqlDbType.VarChar, 100)
        arParameters(Me.ClasssFields.fldDescription).Value = Description
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFPressureUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFPressureUpdate", arParameters)
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
    '* Class: Adds a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef FPressureID As Integer, _
                ByVal Name As String, _
                ByVal Description As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@FPressureID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldName) = New SqlParameter("@Name", SqlDbType.NVarChar, 100)
        arParameters(Me.ClasssFields.fldName).Value = Description '6/12/06 mbr: for now, just insert Description into Name....this field may be deleted after further testing
        arParameters(Me.ClasssFields.fldDescription) = New SqlParameter("@Description", SqlDbType.VarChar, 100)
        arParameters(Me.ClasssFields.fldDescription).Value = Description
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFPressureInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFPressureInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            FPressureID = CType(arParameters(0).Value, Integer)
            Return True
        End If
    End Function




    '**************************************************************************
    '*  
    '* Name:        Delete
    '*
    '* Class: Deletes a record from the [PatientInfo] table identified by a key.
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
        arParameters(0) = New SqlParameter("@FPressureID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFPressureDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFPressureDelete", arParameters)
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


End Class 'dalFPressure
