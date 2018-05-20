
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalOBHistory
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
Public Class dalOBHistory

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum OBHistoryFields
        fldID = 0
        fldDate = 1
        fldMode = 2
        fldBirthWeight = 3
        fldComplications = 4
        fldChartID = 5
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
    '* Name:        GetAll
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
    Public Function GetAll(ByVal ChartID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ChartID
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spOBHistoryGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spOBHistoryGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
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
    Public Function Update(ByVal DeliveryID As Integer, _
                ByVal OBHistoryDate As String, _
                ByVal Mode As String, _
                ByVal BirthWeight As String, _
                ByVal Complications As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.OBHistoryFields.fldID) = New SqlParameter("@DeliveryID", SqlDbType.Int)
        arParameters(Me.OBHistoryFields.fldID).Value = DeliveryID
        arParameters(Me.OBHistoryFields.fldDate) = New SqlParameter("@Date", SqlDbType.NVarChar, 50)
        arParameters(Me.OBHistoryFields.fldDate).Value = OBHistoryDate
        arParameters(Me.OBHistoryFields.fldMode) = New SqlParameter("@Mode", SqlDbType.NVarChar, 50)
        arParameters(Me.OBHistoryFields.fldMode).Value = Mode
        arParameters(Me.OBHistoryFields.fldBirthWeight) = New SqlParameter("@BirthWeight", SqlDbType.NVarChar, 50)
        arParameters(Me.OBHistoryFields.fldBirthWeight).Value = BirthWeight
        arParameters(Me.OBHistoryFields.fldComplications) = New SqlParameter("@Complications", SqlDbType.NVarChar, 255)
        arParameters(Me.OBHistoryFields.fldComplications).Value = Complications
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOBHistoryUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOBHistoryUpdate", arParameters)
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
    Public Function Add(ByRef DeliveryID As Integer, _
                ByVal OBHistoryDate As String, _
                ByVal Mode As String, _
                ByVal BirthWeight As String, _
                ByVal Complications As String, _
                ByVal ChartID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(5) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.OBHistoryFields.fldID) = New SqlParameter("@DeliveryID", SqlDbType.Int)
        arParameters(Me.OBHistoryFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.OBHistoryFields.fldDate) = New SqlParameter("@Date", SqlDbType.NVarChar, 50)
        arParameters(Me.OBHistoryFields.fldDate).Value = OBHistoryDate
        arParameters(Me.OBHistoryFields.fldMode) = New SqlParameter("@Mode", SqlDbType.NVarChar, 50)
        arParameters(Me.OBHistoryFields.fldMode).Value = Mode
        arParameters(Me.OBHistoryFields.fldBirthWeight) = New SqlParameter("@BirthWeight", SqlDbType.NVarChar, 50)
        arParameters(Me.OBHistoryFields.fldBirthWeight).Value = BirthWeight
        arParameters(Me.OBHistoryFields.fldComplications) = New SqlParameter("@Complications", SqlDbType.NVarChar, 255)
        arParameters(Me.OBHistoryFields.fldComplications).Value = Complications
        arParameters(Me.OBHistoryFields.fldChartID) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(Me.OBHistoryFields.fldChartID).Value = ChartID


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOBHistoryInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOBHistoryInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            DeliveryID = CType(arParameters(0).Value, Integer)
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
        arParameters(Me.OBHistoryFields.fldID) = New SqlParameter("@DeliveryID", SqlDbType.Int)
        arParameters(Me.OBHistoryFields.fldID).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOBHistoryDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOBHistoryDelete", arParameters)
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


End Class 'dalOBHistory
