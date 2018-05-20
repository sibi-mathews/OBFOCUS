
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalPrescription
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
Public Class dalPrescription

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



#Region "Main procedures - GetLab, Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Name:        GetFormularyTemplateAll
    '*
    '* Description: Returns all records in the [GetFormularyTemplateAll] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetFormularyTemplateAll() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spFormularyTemplateGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spFormularyTemplateGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
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
    Public Function GetByKey(ByVal ID As Integer, _
                    ByRef PrescriptionName As String, _
                    ByRef Script As String) As Boolean

        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Formulary", SqlDbType.VarChar, 50)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@Body", SqlDbType.VarChar, 5000)
        arParameters(2).Direction = ParameterDirection.Output
        
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFormularyTemplateGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFormularyTemplateGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(0).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            PrescriptionName = ProcessNull.GetString(arParameters(1).Value)
            PrescriptionName = PrescriptionName.Trim()
            Script = ProcessNull.GetString(arParameters(2).Value)
            Script = Script.Trim()
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try

    End Function
    '**************************************************************************
    '*  
    '* Name:        GetPrescription
    '*
    '* Description: Returns all records in the [Labs] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetPrescription(ByVal ChartID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ChartID
        ' Prescription stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPrescriptionGetByKey", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPrescriptionGetByKey", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
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
    Public Function Update(ByVal ID As Integer, _
                       ByVal PrescriptionDate As String, _
                        ByVal ExaminerID As Integer, _
                       ByVal PharmacyID As Integer, _
                       ByVal Script As String, _
                       ByVal Refills As Short, _
                       ByVal Signed As String) As Boolean

        Dim arParameters(6) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PrescriptionID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@PrescriptionDate", SqlDbType.SmallDateTime)
        If PrescriptionDate = "" Then
            arParameters(1).Value = DBNull.Value
        Else
            arParameters(1).Value = PrescriptionDate
        End If
        arParameters(2) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        If ExaminerID = Nothing Then
            arParameters(2).Value = DBNull.Value
        Else
            arParameters(2).Value = ExaminerID
        End If
        arParameters(3) = New SqlParameter("@PharmacyID", SqlDbType.Int)
        arParameters(3).Value = PharmacyID
        arParameters(4) = New SqlParameter("@Script", SqlDbType.VarChar, 2000)
        arParameters(4).Value = Script
        arParameters(5) = New SqlParameter("@Refills", SqlDbType.Int)
        arParameters(5).Value = Refills
        arParameters(6) = New SqlParameter("@Signed", SqlDbType.VarChar, 10)
        arParameters(6).Value = Signed
        ' Prescription stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPrescriptionUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPrescriptionUpdate", arParameters)
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
    '* Name:        UpdateTemplate
    '*
    '* Description: Updates a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdateTemplate(ByVal ID As Integer, _
                       ByVal PrescriptionName As String, _
                        ByVal Script As String) As Boolean

        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Formulary", SqlDbType.VarChar, 50)
        arParameters(1).Value = PrescriptionName
        arParameters(2) = New SqlParameter("@Body", SqlDbType.VarChar, 5000)
        arParameters(2).Value = Script
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFormularyTemplateUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFormularyTemplateUpdate", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not UpdateTemplated.
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
    Public Function Add(ByRef ID As Integer, _
                        ByVal PrescriptionDate As String, _
                        ByVal ExaminerID As Integer, _
                       ByVal PharmacyID As Integer, _
                       ByVal Script As String, _
                       ByVal Refills As Integer, _
                       ByVal ChartID As Integer) As Boolean


        Dim arParameters(6) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PrescriptionID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@PrescriptionDate", SqlDbType.SmallDateTime)
        If PrescriptionDate = "" Then
            arParameters(1).Value = DBNull.Value
        Else
            arParameters(1).Value = PrescriptionDate
        End If
        arParameters(2) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        If ExaminerID = 0 Then
            arParameters(2).Value = DBNull.Value
        Else
            arParameters(2).Value = ExaminerID
        End If
        arParameters(3) = New SqlParameter("@PharmacyID", SqlDbType.Int)
        arParameters(3).Value = PharmacyID
        arParameters(4) = New SqlParameter("@Script", SqlDbType.VarChar, 8000)
        arParameters(4).Value = Script
        arParameters(5) = New SqlParameter("@Refills", SqlDbType.Int)
        arParameters(5).Value = Refills
        arParameters(6) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(6).Value = ChartID
        ' Prescription stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPrescriptionInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPrescriptionInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            ID = CType(arParameters(0).Value, Integer)
            Return True
        End If

    End Function

    '**************************************************************************
    '*  
    '* Name:        AddTemplate
    '*
    '* Description: Adds a new record to the [FormularyTemplate] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function AddTemplate(ByRef ID As Integer, _
                        ByVal PrescriptionName As String, _
                        ByVal Script As String) As Boolean


        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@Formulary", SqlDbType.VarChar, 50)
        arParameters(1).Value = PrescriptionName
        arParameters(2) = New SqlParameter("@Body", SqlDbType.VarChar, 5000)
        arParameters(2).Value = Script
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFormularyTemplateInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFormularyTemplateInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            ID = CType(arParameters(0).Value, Integer)
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
    Public Function Delete(ByVal PrescriptionID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PrescriptionID", SqlDbType.Int)
        arParameters(0).Value = PrescriptionID

        ' Prescription stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPrescriptionDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPrescriptionDelete", arParameters)
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
    '* Name:        DeleteTemplate
    '*
    '* Description: DeleteTemplates a record from the [FormularyTemplate] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to DeleteTemplate
    '*
    '* Returns:     Boolean indicating if record was DeleteTemplated or not. 
    '*              True (record found and DeleteTemplated); False (otherwise).
    '*
    '**************************************************************************
    Public Function DeleteTemplate(ByVal ID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Prescription stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFormularyTemplateDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFormularyTemplateDelete", arParameters)
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


End Class 'dalPrescription
