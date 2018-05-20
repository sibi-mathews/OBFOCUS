
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Formulary:        dalFormulary
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
Public Class dalFormulary

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum ClasssFields
        fldID = 0
        fldGeneric = 1
        fldCategory = 2
        fldTradename = 3
        fldIndications = 4
        fldDosage = 5
        fldContraindications = 6
        fldPregnancy = 7
        fldBreastfeeding = 8
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



#Region "Main procedures - GetByKey, Add, Update & Delete"
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
    Public Function GetByKey(ByVal PharmID As Integer, _
                ByRef Generic As String, _
                ByRef Category As String, _
                ByRef Tradename As String, _
                ByRef Indications As String, _
                ByRef Dosage As String, _
                ByRef Contraindications As String, _
                ByRef Pregnancy As String, _
                ByRef Breastfeeding As String) As Boolean
        Dim arParameters(8) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@PharmID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Value = PharmID
        arParameters(Me.ClasssFields.fldGeneric) = New SqlParameter("@Generic", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldGeneric).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldCategory) = New SqlParameter("@Category", SqlDbType.VarChar, 255)
        arParameters(Me.ClasssFields.fldCategory).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldTradename) = New SqlParameter("@tradename", SqlDbType.VarChar, 100)
        arParameters(Me.ClasssFields.fldTradename).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldIndications) = New SqlParameter("@Indications", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldIndications).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldDosage) = New SqlParameter("@Dosage", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldDosage).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldContraindications) = New SqlParameter("@Contraindications", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldContraindications).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldPregnancy) = New SqlParameter("@Pregnancy", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldPregnancy).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldBreastfeeding) = New SqlParameter("@Breastfeeding", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldBreastfeeding).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFormularyGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFormularyGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.ClasssFields.fldGeneric).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Generic = ProcessNull.GetString(arParameters(Me.ClasssFields.fldGeneric).Value)
            Category = ProcessNull.GetString(arParameters(Me.ClasssFields.fldCategory).Value)
            Tradename = ProcessNull.GetString(arParameters(Me.ClasssFields.fldTradename).Value)
            Indications = ProcessNull.GetString(arParameters(Me.ClasssFields.fldIndications).Value)
            Dosage = ProcessNull.GetString(arParameters(Me.ClasssFields.fldDosage).Value)
            Contraindications = ProcessNull.GetString(arParameters(Me.ClasssFields.fldContraindications).Value)
            Pregnancy = ProcessNull.GetString(arParameters(Me.ClasssFields.fldPregnancy).Value)
            Breastfeeding = ProcessNull.GetString(arParameters(Me.ClasssFields.fldBreastfeeding).Value)

            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        GetFormularyAll
    '*
    '* Description: Returns all records in the [Nurses Flow Sheet] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetFormularyAll() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spFormularyGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spFormularyGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
  
    '**************************************************************************
    '*  
    '* Name:        GetFormularyAllByTrade
    '*
    '* Description: Returns all records in the [Nurses Flow Sheet] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetFormularyAllByTrade(ByVal TradeName As String) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@TradeName", SqlDbType.NVarChar, 100)
        arParameters(0).Value = TradeName
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spFormularyGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spFormularyGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
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
    Public Function Update(ByVal PharmID As Integer, _
                ByVal Generic As String, _
                ByVal Category As String, _
                ByVal Tradename As String, _
                ByVal Indications As String, _
                ByVal Dosage As String, _
                ByVal Contraindications As String, _
                ByVal Pregnancy As String, _
                ByVal Breastfeeding As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(8) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@PharmID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Value = PharmID
        arParameters(Me.ClasssFields.fldGeneric) = New SqlParameter("@Generic", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldGeneric).Value = Generic
        arParameters(Me.ClasssFields.fldCategory) = New SqlParameter("@Category", SqlDbType.VarChar, 255)
        arParameters(Me.ClasssFields.fldCategory).Value = Category
        arParameters(Me.ClasssFields.fldTradename) = New SqlParameter("@tradename", SqlDbType.VarChar, 100)
        arParameters(Me.ClasssFields.fldTradename).Value = Tradename
        arParameters(Me.ClasssFields.fldIndications) = New SqlParameter("@Indications", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldIndications).Value = Indications
        arParameters(Me.ClasssFields.fldDosage) = New SqlParameter("@Dosage", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldDosage).Value = Dosage
        arParameters(Me.ClasssFields.fldContraindications) = New SqlParameter("@Contraindications", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldContraindications).Value = Contraindications
        arParameters(Me.ClasssFields.fldPregnancy) = New SqlParameter("@Pregnancy", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldPregnancy).Value = Pregnancy
        arParameters(Me.ClasssFields.fldBreastfeeding) = New SqlParameter("@breastfeeding", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldBreastfeeding).Value = Breastfeeding

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFormularyUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFormularyUpdate", arParameters)
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
    Public Function Add(ByRef PharmID As Integer, _
                ByVal Generic As String, _
                ByVal Category As String, _
                ByVal Tradename As String, _
                ByVal Indications As String, _
                ByVal Dosage As String, _
                ByVal Contraindications As String, _
                ByVal Pregnancy As String, _
                ByVal Breastfeeding As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(8) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldID) = New SqlParameter("@PharmID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldGeneric) = New SqlParameter("@Generic", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldGeneric).Value = Generic
        arParameters(Me.ClasssFields.fldCategory) = New SqlParameter("@Category", SqlDbType.VarChar, 255)
        arParameters(Me.ClasssFields.fldCategory).Value = Category
        arParameters(Me.ClasssFields.fldTradename) = New SqlParameter("@tradename", SqlDbType.VarChar, 100)
        arParameters(Me.ClasssFields.fldTradename).Value = Tradename
        arParameters(Me.ClasssFields.fldIndications) = New SqlParameter("@Indications", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldIndications).Value = Indications
        arParameters(Me.ClasssFields.fldDosage) = New SqlParameter("@Dosage", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldDosage).Value = Dosage
        arParameters(Me.ClasssFields.fldContraindications) = New SqlParameter("@Contraindications", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldContraindications).Value = Contraindications
        arParameters(Me.ClasssFields.fldPregnancy) = New SqlParameter("@Pregnancy", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldPregnancy).Value = Pregnancy
        arParameters(Me.ClasssFields.fldBreastfeeding) = New SqlParameter("@Breastfeeding", SqlDbType.VarChar, 8000)
        arParameters(Me.ClasssFields.fldBreastfeeding).Value = Breastfeeding

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFormularyInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFormularyInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            PharmID = CType(arParameters(0).Value, Integer)
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
        arParameters(0) = New SqlParameter("@PharmID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFormularyDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFormularyDelete", arParameters)
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


End Class 'dalFormulary
