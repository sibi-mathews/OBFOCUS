
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalTargetedSk
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
Public Class dalTargetedSk

#Region "Module level variables and enums"

    ' Public ENUM used to enumAD columns 
    Public Enum TargetedSkFields
        fldOBUSID = 0
        fldID = 1
        fldFetusname = 2
        fldEGA = 3
        fldHum = 4
        fldHumAge = 5
        fldUlna = 6
        fldUlnaAge = 7
        fldRad = 8
        fldRadAge = 9
        fldTib = 10
        fldTibAge = 11
        fldFib = 12
        fldFibAge = 13
        fldFem = 14
        fldFemAge = 15
        fldTC = 16
        fldTCAge = 17
        fldACM = 18
        fldExamID = 19
        fldSummary = 20
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
    Public Function GetByKey(ByVal OBUSID As Integer, _
                ByRef ID As Integer, _
                ByRef FetusName As String, _
                ByRef EGA As Single, _
                ByRef Hum As Single, _
                ByRef HumAge As Single, _
                ByRef Ulna As Single, _
                ByRef UlnaAge As Single, _
                ByRef Rad As Single, _
                ByRef RadAge As Single, _
                ByRef Tib As Single, _
                ByRef TibAge As Single, _
                ByRef Fib As Single, _
                ByRef FibAge As Single, _
                ByRef Fem As Single, _
                ByRef FemAge As Single, _
                ByRef TC As Single, _
                ByRef TCAge As Single, _
                ByRef ACM As Single, _
                ByRef ExamID As Integer, _
                ByRef Summary As String) As Boolean
        ' Set the stored procedure parameters
        Dim arParameters(20) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.TargetedSkFields.fldOBUSID) = New SqlParameter("@OBUSID", SqlDbType.Int)
        arParameters(Me.TargetedSkFields.fldOBUSID).Value = OBUSID
        arParameters(Me.TargetedSkFields.fldID) = New SqlParameter("@TargetedID", SqlDbType.Int)
        arParameters(Me.TargetedSkFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldFetusname) = New SqlParameter("@fetusname", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedSkFields.fldFetusname).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldEGA) = New SqlParameter("@ega", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldEGA).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldHum) = New SqlParameter("@hum", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldHum).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldHumAge) = New SqlParameter("@HumAge", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldHumAge).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldUlna) = New SqlParameter("@Ulna", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldUlna).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldUlnaAge) = New SqlParameter("@UlnaAge", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldUlnaAge).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldRad) = New SqlParameter("@Rad", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldRad).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldRadAge) = New SqlParameter("@RadAge", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldRadAge).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldTib) = New SqlParameter("@Tib", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldTib).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldTibAge) = New SqlParameter("@TibAge", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldTibAge).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldFib) = New SqlParameter("@Fib", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldFib).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldFibAge) = New SqlParameter("@FibAge", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldFibAge).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldFem) = New SqlParameter("@Fem", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldFem).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldFemAge) = New SqlParameter("@FemAge", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldFemAge).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldTC) = New SqlParameter("@TC", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldTC).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldTCAge) = New SqlParameter("@TCAge", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldTCAge).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldACM) = New SqlParameter("@ACM", SqlDbType.Real)
        arParameters(Me.TargetedSkFields.fldACM).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldExamID) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(Me.TargetedSkFields.fldExamID).Direction = ParameterDirection.Output
        arParameters(Me.TargetedSkFields.fldSummary) = New SqlParameter("@Summary", SqlDbType.VarChar, 8000)
        arParameters(Me.TargetedSkFields.fldSummary).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spTargetedSkGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spTargetedSkGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.TargetedSkFields.fldID).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            ID = ProcessNull.GetInt32(arParameters(Me.TargetedSkFields.fldID).Value)
            FetusName = ProcessNull.GetString(arParameters(Me.TargetedSkFields.fldFetusname).Value)
            EGA = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldEGA).Value)
            Hum = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldHum).Value)
            HumAge = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldHumAge).Value)
            Ulna = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldUlna).Value)
            UlnaAge = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldUlnaAge).Value)
            Rad = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldRad).Value)
            RadAge = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldRadAge).Value)
            Tib = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldTib).Value)
            TibAge = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldTibAge).Value)
            Fib = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldFib).Value)
            FibAge = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldFibAge).Value)
            Fem = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldFem).Value)
            FemAge = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldFemAge).Value)
            TC = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldTC).Value)
            TCAge = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldTCAge).Value)
            ACM = ProcessNull.GetDecimal(arParameters(Me.TargetedSkFields.fldACM).Value)
            ExamID = ProcessNull.GetInt32(arParameters(Me.TargetedSkFields.fldExamID).Value)
            Summary = ProcessNull.GetString(arParameters(Me.TargetedSkFields.fldSummary).Value)
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
    Public Function Update(ByVal ID As Integer, _
                ByVal FetusName As String, _
                ByVal EGA As Single, _
                ByVal Hum As Single, _
                ByVal HumAge As Single, _
                ByVal Ulna As Single, _
                ByVal UlnaAge As Single, _
                ByVal Rad As Single, _
                ByVal RadAge As Single, _
                ByVal Tib As Single, _
                ByVal TibAge As Single, _
                ByVal Fib As Single, _
                ByVal FibAge As Single, _
                ByVal Fem As Single, _
                ByVal FemAge As Single, _
                ByVal TC As Single, _
                ByVal TCAge As Single, _
                ByVal ACM As Single, _
                ByVal ExamID As Integer, _
                ByVal Summary As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(19) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@TargetedID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@fetusname", SqlDbType.NVarChar, 50)
        arParameters(1).Value = FetusName
        arParameters(2) = New SqlParameter("@ega", SqlDbType.Real)
        arParameters(2).Value = EGA
        arParameters(3) = New SqlParameter("@hum", SqlDbType.Real)
        arParameters(3).Value = Hum
        arParameters(4) = New SqlParameter("@HumAge", SqlDbType.NVarChar, 50)
        arParameters(4).Value = HumAge
        arParameters(5) = New SqlParameter("@Ulna", SqlDbType.Real)
        arParameters(5).Value = Ulna
        arParameters(6) = New SqlParameter("@UlnaAge", SqlDbType.Real)
        arParameters(6).Value = UlnaAge
        arParameters(7) = New SqlParameter("@Rad", SqlDbType.Real)
        arParameters(7).Value = Rad
        arParameters(8) = New SqlParameter("@RadAge", SqlDbType.Real)
        arParameters(8).Value = RadAge
        arParameters(9) = New SqlParameter("@Tib", SqlDbType.Real)
        arParameters(9).Value = Tib
        arParameters(10) = New SqlParameter("@TibAge", SqlDbType.Real)
        arParameters(10).Value = TibAge
        arParameters(11) = New SqlParameter("@Fib", SqlDbType.Real)
        arParameters(11).Value = Fib
        arParameters(12) = New SqlParameter("@FibAge", SqlDbType.Real)
        arParameters(12).Value = FibAge
        arParameters(13) = New SqlParameter("@Fem", SqlDbType.Real)
        arParameters(13).Value = Fem
        arParameters(14) = New SqlParameter("@FemAge", SqlDbType.Real)
        arParameters(14).Value = FemAge
        arParameters(15) = New SqlParameter("@Tc", SqlDbType.Real)
        arParameters(15).Value = TC
        arParameters(16) = New SqlParameter("@TcAge", SqlDbType.Real)
        arParameters(16).Value = TCAge
        arParameters(17) = New SqlParameter("@ACM", SqlDbType.Real)
        arParameters(17).Value = ACM
        arParameters(18) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(18).Value = ExamID
        arParameters(19) = New SqlParameter("@Summary", SqlDbType.VarChar, 8000)
        arParameters(19).Value = Summary

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spTargetedSkUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spTargetedSkUpdate", arParameters)
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


End Class 'dalTargetedSk
