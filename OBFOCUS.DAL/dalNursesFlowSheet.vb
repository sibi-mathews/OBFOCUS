
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalNursesFlowSheet
'*
'* Description: Data access layer for Table NursesFlowSheet
'*
'* Remarks:     Uses OleDb and embedded SQL for maintaining the data.
'*-----------------------------------------------------------------------------
'*                      CHANGE HISTORY
'*   Change No:   Date:          Author:   Description:
'*   _________    ___________    ______    ____________________________________
'*      001       1/26/2005     MR        Created.                                
'* 
'******************************************************************************
Public Class dalNursesFlowSheet

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum PhysicianFields
        fldID = 0
        fldExamDate = 1
        fldPounds = 2
        fldSBP = 3
        fldDBP = 4
        fldSBP2 = 5
        fldDBP2 = 6
        fldHearRate = 7
        fldRespRate = 8
        fldProtein = 9
        fldSugar = 10
        fldFHR = 11
        fldCounseling = 12
        fldProcedure = 13
        fldLabs = 14
        fldReturnAppt = 15
        fldComments = 16
        fldTemperature = 17
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
    '* Name:        GetNursesFlowSheet
    '*
    '* Description: Returns all records in the [Nurses Flow Sheet] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetNursesFlowSheet(ByVal ChartID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ChartID
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spNursesFlowSheetGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spNursesFlowSheetGet", arParameters)
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
                            ByVal ExamDate As String, _
                            ByVal Pounds As String, _
                            ByVal SBP As String, _
                            ByVal DBP As String, _
                            ByVal SBP2 As String, _
                            ByVal DBP2 As String, _
                            ByVal HeartRate As String, _
                            ByVal RespRate As String, _
                            ByVal Protein As String, _
                            ByVal Sugar As String, _
                            ByVal FHR As String, _
                            ByVal Counseling As String, _
                            ByVal Procedure As String, _
                            ByVal Labs As String, _
                            ByVal ReturnAppt As String, _
                            ByVal Comments As String, _
                            ByVal Temperature As String, _
                            ByVal UserID As String, _
                            ByVal Locked As Short, _
                            ByVal UpdatedBy As String, _
                            ByVal Interpreter As Short) As Boolean


        Dim arParameters(21) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@nFlowID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@ExamDate", SqlDbType.SmallDateTime)
        If ExamDate = "" Then
            arParameters(1).Value = DBNull.Value
        Else
            arParameters(1).Value = ExamDate
        End If
        arParameters(2) = New SqlParameter("@Pounds", SqlDbType.Int)
        If Pounds = Nothing Then
            arParameters(2).Value = DBNull.Value
        Else
            arParameters(2).Value = Pounds
        End If
        arParameters(3) = New SqlParameter("@SBP", SqlDbType.Int)
        If SBP = Nothing Then
            arParameters(3).Value = DBNull.Value
        Else
            arParameters(3).Value = SBP
        End If
        arParameters(4) = New SqlParameter("@DBP", SqlDbType.Int)
        If DBP = Nothing Then
            arParameters(4).Value = DBNull.Value
        Else
            arParameters(4).Value = DBP
        End If
        arParameters(5) = New SqlParameter("@SBP2", SqlDbType.Int)
        If SBP2 = Nothing Then
            arParameters(5).Value = DBNull.Value
        Else
            arParameters(5).Value = SBP2
        End If
        arParameters(6) = New SqlParameter("@DBP2", SqlDbType.Int)
        If DBP2 = Nothing Then
            arParameters(6).Value = DBNull.Value
        Else
            arParameters(6).Value = DBP2
        End If
        arParameters(7) = New SqlParameter("@HeartRate", SqlDbType.Int)
        If HeartRate = Nothing Then
            arParameters(7).Value = DBNull.Value
        Else
            arParameters(7).Value = HeartRate
        End If
        arParameters(8) = New SqlParameter("@RespRate", SqlDbType.Int)
        If RespRate = Nothing Then
            arParameters(8).Value = DBNull.Value
        Else
            arParameters(8).Value = RespRate
        End If
        arParameters(9) = New SqlParameter("@Protein", SqlDbType.NVarChar, 50)
        arParameters(9).Value = Protein
        arParameters(10) = New SqlParameter("@Sugar", SqlDbType.NVarChar, 50)
        arParameters(10).Value = Sugar
        arParameters(11) = New SqlParameter("@FHR", SqlDbType.NVarChar, 50)
        arParameters(11).Value = FHR
        arParameters(12) = New SqlParameter("@Counseling", SqlDbType.NVarChar, 50)
        arParameters(12).Value = Counseling
        arParameters(13) = New SqlParameter("@Procedure", SqlDbType.NVarChar, 50)
        arParameters(13).Value = Procedure
        arParameters(14) = New SqlParameter("@Labs", SqlDbType.NVarChar, 50)
        arParameters(14).Value = Labs
        arParameters(15) = New SqlParameter("@ReturnAppt", SqlDbType.SmallDateTime)
        If ReturnAppt = Nothing Then
            arParameters(15).Value = DBNull.Value
        Else
            arParameters(15).Value = ReturnAppt
        End If
        arParameters(16) = New SqlParameter("@Comments", SqlDbType.NVarChar, 2000)
        arParameters(16).Value = Comments
        arParameters(17) = New SqlParameter("@Temperature", SqlDbType.Real)
        If Temperature = Nothing Then
            arParameters(17).Value = DBNull.Value
        Else
            arParameters(17).Value = Temperature
        End If
        arParameters(18) = New SqlParameter("@UserID", SqlDbType.NVarChar, 50)
        arParameters(18).Value = UserID
        arParameters(19) = New SqlParameter("@Locked", SqlDbType.Bit)
        arParameters(19).Value = Locked
        arParameters(20) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(20).Value = UpdatedBy
        arParameters(21) = New SqlParameter("@Interpreter", SqlDbType.Bit)
        arParameters(21).Value = Interpreter
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spNursesFlowSheetUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spNursesFlowSheetUpdate", arParameters)
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
    '* Description: Adds a new record to the [NursesFlowSheet] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef ID As Integer, _
                            ByVal ChartID As Integer, _
                            ByVal ExamDate As String, _
                            ByVal Pounds As String, _
                            ByVal SBP As String, _
                            ByVal DBP As String, _
                            ByVal SBP2 As String, _
                            ByVal DBP2 As String, _
                            ByVal HeartRate As String, _
                            ByVal RespRate As String, _
                            ByVal Protein As String, _
                            ByVal Sugar As String, _
                            ByVal FHR As String, _
                            ByVal Counseling As String, _
                            ByVal Procedure As String, _
                            ByVal Labs As String, _
                            ByVal ReturnAppt As String, _
                            ByVal Comments As String, _
                            ByVal Temperature As String, _
                            ByVal UserID As String, _
                            ByVal Locked As Short, _
                            ByVal Interpreter As Short) As Boolean

        Dim arParameters(21) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@nFlowID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(1).Value = ID
        arParameters(2) = New SqlParameter("@ExamDate", SqlDbType.SmallDateTime)
        If ExamDate = "" Then
            arParameters(2).Value = DBNull.Value
        Else
            arParameters(2).Value = ExamDate
        End If
        arParameters(3) = New SqlParameter("@Pounds", SqlDbType.Int)
        If Pounds = Nothing Then
            arParameters(3).Value = DBNull.Value
        Else
            arParameters(3).Value = Pounds
        End If
        arParameters(4) = New SqlParameter("@SBP", SqlDbType.Int)
        If SBP = Nothing Then
            arParameters(4).Value = DBNull.Value
        Else
            arParameters(4).Value = SBP
        End If
        arParameters(5) = New SqlParameter("@DBP", SqlDbType.Int)
        If DBP = Nothing Then
            arParameters(5).Value = DBNull.Value
        Else
            arParameters(5).Value = DBP
        End If
        arParameters(6) = New SqlParameter("@SBP2", SqlDbType.Int)
        If SBP2 = Nothing Then
            arParameters(6).Value = DBNull.Value
        Else
            arParameters(6).Value = SBP2
        End If
        arParameters(7) = New SqlParameter("@DBP2", SqlDbType.Int)
        If DBP2 = Nothing Then
            arParameters(7).Value = DBNull.Value
        Else
            arParameters(7).Value = DBP2
        End If
        arParameters(8) = New SqlParameter("@HeartRate", SqlDbType.Int)
        If HeartRate = Nothing Then
            arParameters(8).Value = DBNull.Value
        Else
            arParameters(8).Value = HeartRate
        End If
        arParameters(9) = New SqlParameter("@RespRate", SqlDbType.Int)
        If RespRate = Nothing Then
            arParameters(9).Value = DBNull.Value
        Else
            arParameters(9).Value = RespRate
        End If
        arParameters(10) = New SqlParameter("@Protein", SqlDbType.NVarChar, 50)
        arParameters(10).Value = Protein
        arParameters(11) = New SqlParameter("@Sugar", SqlDbType.NVarChar, 50)
        arParameters(11).Value = Sugar
        arParameters(12) = New SqlParameter("@FHR", SqlDbType.NVarChar, 50)
        arParameters(12).Value = FHR
        arParameters(13) = New SqlParameter("@Counseling", SqlDbType.NVarChar, 50)
        arParameters(13).Value = Counseling
        arParameters(14) = New SqlParameter("@Procedure", SqlDbType.NVarChar, 50)
        arParameters(14).Value = Procedure
        arParameters(15) = New SqlParameter("@Labs", SqlDbType.NVarChar, 50)
        arParameters(15).Value = Labs
        arParameters(16) = New SqlParameter("@ReturnAppt", SqlDbType.SmallDateTime)
        If ReturnAppt = Nothing Then
            arParameters(16).Value = DBNull.Value
        Else
            arParameters(16).Value = ReturnAppt
        End If
        arParameters(17) = New SqlParameter("@Comments", SqlDbType.NVarChar, 2000)
        arParameters(17).Value = Comments
        arParameters(18) = New SqlParameter("@Temperature", SqlDbType.Real)
        If Temperature = Nothing Then
            arParameters(18).Value = DBNull.Value
        Else
            arParameters(18).Value = Temperature
        End If
        arParameters(19) = New SqlParameter("@UserID", SqlDbType.NVarChar, 50)
        arParameters(19).Value = UserID
        arParameters(20) = New SqlParameter("@Locked", SqlDbType.Bit)
        arParameters(20).Value = Locked
        arParameters(21) = New SqlParameter("@Interpreter", SqlDbType.Bit)
        arParameters(21).Value = Interpreter
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spNursesFlowSheetInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spNursesFlowSheetInsert", arParameters)
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
    '* Description: Deletes a record from the [NursesFlowSheet] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to delete
    '*
    '* Returns:     Boolean indicating if record was deleted or not. 
    '*              True (record found and deleted); False (otherwise).
    '*
    '**************************************************************************
    Public Function Delete(ByVal nFlowID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@nFlowID", SqlDbType.Int)
        arParameters(0).Value = nFlowID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spNursesFlowSheetDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spNursesFlowSheetDelete", arParameters)
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


End Class 'dalNursesFlowSheet
