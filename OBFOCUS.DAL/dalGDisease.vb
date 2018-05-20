Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* GDisease:        dalGDisease
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
Public Class dalGDisease

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum ClasssFields
        fldID = 0
        fldGDisease = 1
        fldFrequency = 2
        fldInheritance = 3
    End Enum


    'Used for transaction support
    Private _Transaction As SqlTransaction = Nothing


    '**************************************************************************
    '*  
    '* GDisease:        Transaction
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
    '* GDisease:        New
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
    '* GDisease:        New
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
    '* GDisease:        GetByKey
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
    Public Function GetByKey(ByVal GDiseaseID As Integer, _
                ByRef GDisease As String, _
                ByRef Frequency As Double, _
                ByRef Inheritance As String) As Boolean
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@GDiseaseID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Value = GDiseaseID
        arParameters(Me.ClasssFields.fldGDisease) = New SqlParameter("@GDisease", SqlDbType.NVarChar, 155)
        arParameters(Me.ClasssFields.fldGDisease).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldFrequency) = New SqlParameter("@Frequency", SqlDbType.Real)
        arParameters(Me.ClasssFields.fldFrequency).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldInheritance) = New SqlParameter("@Inheritance", SqlDbType.VarChar, 50)
        arParameters(Me.ClasssFields.fldInheritance).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spGDiseaseGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spGDiseaseGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.ClasssFields.fldGDisease).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            GDisease = ProcessNull.GetString(arParameters(Me.ClasssFields.fldGDisease).Value)
            Frequency = ProcessNull.GetDouble(arParameters(Me.ClasssFields.fldFrequency).Value)
            Inheritance = ProcessNull.GetString(arParameters(Me.ClasssFields.fldInheritance).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* GDisease:        Update
    '*
    '* Class: Updates a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was updated or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal GDiseaseID As Integer, _
                ByVal GDisease As String, _
                ByVal Frequency As Double, _
                ByVal Inheritance As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@GDiseaseID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Value = GDiseaseID
        arParameters(Me.ClasssFields.fldGDisease) = New SqlParameter("@GDisease", SqlDbType.NVarChar, 155)
        arParameters(Me.ClasssFields.fldGDisease).Value = GDisease
        arParameters(Me.ClasssFields.fldFrequency) = New SqlParameter("@Frequency", SqlDbType.Real)
        arParameters(Me.ClasssFields.fldFrequency).Value = Frequency
        arParameters(Me.ClasssFields.fldInheritance) = New SqlParameter("@Inheritance", SqlDbType.VarChar, 50)
        arParameters(Me.ClasssFields.fldInheritance).Value = Inheritance
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spGDiseaseUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spGDiseaseUpdate", arParameters)
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
    '* GDisease:        Add
    '*
    '* Class: Adds a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef GDiseaseID As Integer, _
                ByVal GDisease As String, _
                ByVal Frequency As Double, _
                ByVal Inheritance As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@GDiseaseID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldGDisease) = New SqlParameter("@GDisease", SqlDbType.NVarChar, 155)
        arParameters(Me.ClasssFields.fldGDisease).Value = GDisease
        arParameters(Me.ClasssFields.fldFrequency) = New SqlParameter("@Frequency", SqlDbType.Real)
        arParameters(Me.ClasssFields.fldFrequency).Value = Frequency
        arParameters(Me.ClasssFields.fldInheritance) = New SqlParameter("@Inheritance", SqlDbType.VarChar, 50)
        arParameters(Me.ClasssFields.fldInheritance).Value = Inheritance
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spGDiseaseInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spGDiseaseInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            GDiseaseID = CType(arParameters(0).Value, Integer)
            Return True
        End If
    End Function




    '**************************************************************************
    '*  
    '* GDisease:        Delete
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
        arParameters(0) = New SqlParameter("@GDiseaseID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spGDiseaseDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spGDiseaseDelete", arParameters)
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


End Class 'dalGDisease
