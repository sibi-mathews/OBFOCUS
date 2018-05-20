
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Allergy:        dalAllergyType
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
Public Class dalAllergyType

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum ClasssFields
        fldID = 0
        fldAllergy = 1
        fldClass = 2
    End Enum


    'Used for transaction support
    Private _Transaction As SqlTransaction = Nothing


    '**************************************************************************
    '*  
    '* Allergy:        Transaction
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
    '* Allergy:        New
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
    '* Allergy:        New
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
    '* Allergy:        GetByKey
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
    Public Function GetByKey(ByVal AllergyTypeID As Integer, _
                ByRef Allergy As String, _
                ByRef ClassDescription As String) As Boolean
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@AllergyTypeID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Value = AllergyTypeID
        arParameters(Me.ClasssFields.fldAllergy) = New SqlParameter("@Allergy", SqlDbType.NVarChar, 100)
        arParameters(Me.ClasssFields.fldAllergy).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldClass) = New SqlParameter("@ClassDescription", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldClass).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAllergyTypeGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAllergyTypeGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.ClasssFields.fldAllergy).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Allergy = ProcessNull.GetString(arParameters(Me.ClasssFields.fldAllergy).Value)
            ClassDescription = ProcessNull.GetString(arParameters(Me.ClasssFields.fldClass).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
        '**************************************************************************
        '*  
        '* Allergy:        Update
        '*
        '* Class: Updates a record in the Chart table identified by a key.
        '*
        '*
        '* Returns:     Boolean indicating if record was updated or not. 
        '*              True (record found); False (otherwise).
        '*
        '**************************************************************************
    Public Function Update(ByVal AllergyTypeID As Integer, _
                ByVal Allergy As String, _
                ByVal ClassDescription As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@AllergyTypeID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Value = AllergyTypeID
        arParameters(Me.ClasssFields.fldAllergy) = New SqlParameter("@Allergy", SqlDbType.NVarChar, 100)
        arParameters(Me.ClasssFields.fldAllergy).Value = Allergy
        arParameters(Me.ClasssFields.fldClass) = New SqlParameter("@ClassDescription", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldClass).Value = ClassDescription
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAllergyTypeUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAllergyTypeUpdate", arParameters)
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
    '* Allergy:        Add
    '*
    '* Class: Adds a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef AllergyTypeID As Integer, _
                ByVal Allergy As String, _
                ByVal ClassDescription As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@AllergyTypeID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldAllergy) = New SqlParameter("@Allergy", SqlDbType.NVarChar, 100)
        arParameters(Me.ClasssFields.fldAllergy).Value = Allergy
        arParameters(Me.ClasssFields.fldClass) = New SqlParameter("@ClassDescription", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldClass).Value = ClassDescription
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAllergyTypeInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAllergyTypeInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            AllergyTypeID = CType(arParameters(0).Value, Integer)
            Return True
        End If
    End Function




    '**************************************************************************
    '*  
    '* Allergy:        Delete
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
        arParameters(0) = New SqlParameter("@AllergyTypeID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAllergyTypeDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAllergyTypeDelete", arParameters)
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


End Class 'dalAllergyType
