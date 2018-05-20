
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalUltrasonographer
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
Public Class dalUltrasonographer

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum UltrasonographerFields
        fldID = 0
        fldName = 1
        fldTItle = 2
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
    Public Function GetByKey(ByVal UltrasonographerID As Integer, _
                ByRef Name As String, _
                ByRef Title As String) As Boolean
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.UltrasonographerFields.fldID) = New SqlParameter("@UltrasonographerID", SqlDbType.Int)
        arParameters(Me.UltrasonographerFields.fldID).Value = UltrasonographerID
        arParameters(Me.UltrasonographerFields.fldName) = New SqlParameter("@Name", SqlDbType.NVarChar, 100)
        arParameters(Me.UltrasonographerFields.fldName).Direction = ParameterDirection.Output
        arParameters(Me.UltrasonographerFields.fldTItle) = New SqlParameter("@Title", SqlDbType.NVarChar, 50)
        arParameters(Me.UltrasonographerFields.fldTItle).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasonographerGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUltrasonographerGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.UltrasonographerFields.fldName).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Name = ProcessNull.GetString(arParameters(Me.UltrasonographerFields.fldName).Value)
            Title = ProcessNull.GetString(arParameters(Me.UltrasonographerFields.fldTItle).Value)

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
    Public Function Update(ByVal UltrasonographerID As Integer, _
                ByVal Name As String, _
                ByVal Title As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.UltrasonographerFields.fldID) = New SqlParameter("@UltrasonographerID", SqlDbType.Int)
        arParameters(Me.UltrasonographerFields.fldID).Value = UltrasonographerID
        arParameters(Me.UltrasonographerFields.fldName) = New SqlParameter("@Name", SqlDbType.NVarChar, 100)
        arParameters(Me.UltrasonographerFields.fldName).Value = Name
        arParameters(Me.UltrasonographerFields.fldTItle) = New SqlParameter("@Title", SqlDbType.NVarChar, 50)
        arParameters(Me.UltrasonographerFields.fldTItle).Value = Title
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasonographerUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUltrasonographerUpdate", arParameters)
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
    Public Function Add(ByRef UltrasonographerID As Integer, _
                ByVal Name As String, _
                ByVal Title As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.UltrasonographerFields.fldID) = New SqlParameter("@UltrasonographerID", SqlDbType.Int)
        arParameters(Me.UltrasonographerFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.UltrasonographerFields.fldName) = New SqlParameter("@Name", SqlDbType.NVarChar, 100)
        arParameters(Me.UltrasonographerFields.fldName).Value = Name
        arParameters(Me.UltrasonographerFields.fldTItle) = New SqlParameter("@Title", SqlDbType.NVarChar, 50)
        arParameters(Me.UltrasonographerFields.fldTItle).Value = Title
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasonographerInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUltrasonographerInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            UltrasonographerID = CType(arParameters(0).Value, Integer)
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
        arParameters(0) = New SqlParameter("@UltrasonographerID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasonographerDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUltrasonographerDelete", arParameters)
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


End Class 'dalUltrasonographer
