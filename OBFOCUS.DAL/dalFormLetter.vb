
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* NameForm:        dalFormLetter
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
Public Class dalFormLetter

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum ClasssFields
        fldID = 0
        fldNameForm = 1
        fldText = 2
    End Enum


    'Used for transaction support
    Private _Transaction As SqlTransaction = Nothing


    '**************************************************************************
    '*  
    '* NameForm:        Transaction
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
    '* NameForm:        New
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
    '* NameForm:        New
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
    '* NameForm:        GetByKey
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
    Public Function GetByKey(ByVal FormLetterID As Integer, _
                ByRef NameForm As String, _
                ByRef Text As String) As Boolean
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@FormLetterID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Value = FormLetterID
        arParameters(Me.ClasssFields.fldNameForm) = New SqlParameter("@NameForm", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldNameForm).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldText) = New SqlParameter("@Text", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldText).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFormLetterGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFormLetterGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.ClasssFields.fldNameForm).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            NameForm = ProcessNull.GetString(arParameters(Me.ClasssFields.fldNameForm).Value)
            Text = ProcessNull.GetString(arParameters(Me.ClasssFields.fldText).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* NameForm:        Update
    '*
    '* Class: Updates a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was updated or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal FormLetterID As Integer, _
                ByVal NameForm As String, _
                ByVal Text As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@FormLetterID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Value = FormLetterID
        arParameters(Me.ClasssFields.fldNameForm) = New SqlParameter("@NameForm", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldNameForm).Value = NameForm
        arParameters(Me.ClasssFields.fldText) = New SqlParameter("@Text", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldText).Value = Text
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFormLetterUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFormLetterUpdate", arParameters)
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
    '* NameForm:        Add
    '*
    '* Class: Adds a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef FormLetterID As Integer, _
                ByVal NameForm As String, _
                ByVal Text As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@FormLetterID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldNameForm) = New SqlParameter("@NameForm", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldNameForm).Value = NameForm
        arParameters(Me.ClasssFields.fldText) = New SqlParameter("@Text", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldText).Value = Text
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFormLetterInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFormLetterInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            FormLetterID = CType(arParameters(0).Value, Integer)
            Return True
        End If
    End Function




    '**************************************************************************
    '*  
    '* NameForm:        Delete
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
        arParameters(0) = New SqlParameter("@FormLetterID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFormLetterDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFormLetterDelete", arParameters)
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


End Class 'dalFormLetter
