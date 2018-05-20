
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalPReference
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
Public Class dalPReference

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum PReferenceFields
        fldID = 0
        fldName = 1
        fldTitle = 2
        fldDetail = 3
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
    '* Name:        GetPReference
    '*
    '* Description: Returns all records in the [PReference] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetPReference(ByVal PRefName As String) As SqlDataReader
        ' Call stored procedure and return the data
        Try
            Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
            ' Set the parameters
            arParameters(0) = New SqlParameter("@PRefName", SqlDbType.NVarChar, 100)
            If PRefName = "" Then
                arParameters(0).Value = DBNull.Value
            Else
                arParameters(0).Value = PRefName
            End If
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPReferenceGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPReferenceGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function 'GetPReference
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
    Public Function GetByKey(ByVal PatientRefID As Integer, _
                ByRef Name As String, _
                ByRef Title As String, _
                ByRef Detail As String) As Boolean
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.PReferenceFields.fldID) = New SqlParameter("@PatientRefID", SqlDbType.Int)
        arParameters(Me.PReferenceFields.fldID).Value = PatientRefID
        arParameters(Me.PReferenceFields.fldName) = New SqlParameter("@PRefName", SqlDbType.NVarChar, 100)
        arParameters(Me.PReferenceFields.fldName).Direction = ParameterDirection.Output
        arParameters(Me.PReferenceFields.fldTitle) = New SqlParameter("@PRefTitle", SqlDbType.VarChar, 100)
        arParameters(Me.PReferenceFields.fldTitle).Direction = ParameterDirection.Output
        arParameters(Me.PReferenceFields.fldDetail) = New SqlParameter("@PRefDetail", SqlDbType.VarChar, 8000)
        arParameters(Me.PReferenceFields.fldDetail).Direction = ParameterDirection.Output


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPReferenceGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPReferenceGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.PReferenceFields.fldName).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Name = ProcessNull.GetString(arParameters(Me.PReferenceFields.fldName).Value)
            Title = ProcessNull.GetString(arParameters(Me.PReferenceFields.fldTitle).Value)
            Detail = ProcessNull.GetString(arParameters(Me.PReferenceFields.fldDetail).Value)
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
    Public Function Update(ByVal PatientRefID As Integer, _
                ByVal Name As String, _
                ByVal Title As String, _
                   ByVal Detail As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.PReferenceFields.fldID) = New SqlParameter("@PatientRefID", SqlDbType.Int)
        arParameters(Me.PReferenceFields.fldID).Value = PatientRefID
        arParameters(Me.PReferenceFields.fldName) = New SqlParameter("@PRefName", SqlDbType.NVarChar, 100)
        arParameters(Me.PReferenceFields.fldName).Value = Name
        arParameters(Me.PReferenceFields.fldTitle) = New SqlParameter("@PRefTitle", SqlDbType.VarChar, 100)
        arParameters(Me.PReferenceFields.fldTitle).Value = Title
        arParameters(Me.PReferenceFields.fldDetail) = New SqlParameter("@PRefDetail", SqlDbType.VarChar, 8000)
        arParameters(Me.PReferenceFields.fldDetail).Value = Detail
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPReferenceUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPReferenceUpdate", arParameters)
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
    Public Function Add(ByRef PatientRefID As Integer, _
                ByVal Name As String, _
                ByVal Title As String, _
                ByVal Detail As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.PReferenceFields.fldID) = New SqlParameter("@PatientRefID", SqlDbType.Int)
        arParameters(Me.PReferenceFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.PReferenceFields.fldName) = New SqlParameter("@PRefName", SqlDbType.NVarChar, 100)
        arParameters(Me.PReferenceFields.fldName).Value = Name
        arParameters(Me.PReferenceFields.fldTitle) = New SqlParameter("@PRefTitle", SqlDbType.VarChar, 100)
        arParameters(Me.PReferenceFields.fldTitle).Value = Title
        arParameters(Me.PReferenceFields.fldDetail) = New SqlParameter("@PRefDetail", SqlDbType.VarChar, 8000)
        arParameters(Me.PReferenceFields.fldDetail).Value = Detail
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPReferenceInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPReferenceInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            PatientRefID = CType(arParameters(0).Value, Integer)
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
        arParameters(0) = New SqlParameter("@PatientRefID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPReferenceDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPReferenceDelete", arParameters)
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


End Class 'dalPReference
