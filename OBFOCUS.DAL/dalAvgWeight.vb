
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalAvgWeight
'*
'* Description: Data access layer for Table Comments
'*
'* Remarks:     Uses OleDb and embedded SQL for maintaining the data.
'*-----------------------------------------------------------------------------
'*                      CHANGE HISTORY
'*   Change No:   Date:          Author:   Description:
'*   _________    ___________    ______    ____________________________________
'*      001       5/9/2007       MR        Created.                                
'* 
'******************************************************************************
Public Class dalAvgWeight

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



#Region "Main procedures "
    '**************************************************************************
    '*  
    '* Name:        GetAvgWeight
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
    Public Function GetAvgWeight(ByVal DateFrom As String, _
                ByVal DateTo As String, _
                ByRef LessThan100 As Integer, _
                ByRef w100to125 As Integer, _
                ByRef w126to250 As Integer, _
                ByRef w251to275 As Integer, _
                ByRef w276to300 As Integer, _
                ByRef w301to325 As Integer, _
                ByRef w326to350 As Integer, _
                ByRef w351to375 As Integer, _
                ByRef w376to400 As Integer, _
                ByRef greaterthan400 As Integer) As Boolean

        Dim arParameters(11) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@FromDate", SqlDbType.DateTime)
        If Len(DateFrom) = 0 Then
            arParameters(0).Value = DBNull.Value
        ElseIf IsDate(DateFrom) Then
            arParameters(0).Value = DateFrom
        End If
        arParameters(1) = New SqlParameter("@ToDate", SqlDbType.DateTime)
        If Len(DateTo) = 0 Then
            arParameters(1).Value = DBNull.Value
        ElseIf IsDate(DateTo) Then
            arParameters(1).Value = DateTo
        End If
        arParameters(2) = New SqlParameter("@LessThan100", SqlDbType.Int)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@w100to125", SqlDbType.Int)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@w126to250", SqlDbType.Int)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@w251to275", SqlDbType.Int)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@w276to300", SqlDbType.Int)
        arParameters(6).Direction = ParameterDirection.Output
        arParameters(7) = New SqlParameter("@w301to325", SqlDbType.Int)
        arParameters(7).Direction = ParameterDirection.Output
        arParameters(8) = New SqlParameter("@w326to350", SqlDbType.Int)
        arParameters(8).Direction = ParameterDirection.Output
        arParameters(9) = New SqlParameter("@w351to375", SqlDbType.Int)
        arParameters(9).Direction = ParameterDirection.Output
        arParameters(10) = New SqlParameter("@w376to400", SqlDbType.Int)
        arParameters(10).Direction = ParameterDirection.Output
        arParameters(11) = New SqlParameter("@greaterthan400", SqlDbType.Int)
        arParameters(11).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "GetMaternalWeightByAge", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "GetMaternalWeightByAge", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(0).Value Is DBNull.Value Then Return False
            ' Return True if data was found. Also populate output (ByRef) parameters.
            LessThan100 = ProcessNull.GetInt32(arParameters(2).Value)
            w100to125 = ProcessNull.GetInt32(arParameters(3).Value)
            w126to250 = ProcessNull.GetInt32(arParameters(4).Value)
            w251to275 = ProcessNull.GetInt32(arParameters(5).Value)
            w276to300 = ProcessNull.GetInt32(arParameters(6).Value)
            w301to325 = ProcessNull.GetInt32(arParameters(7).Value)
            w326to350 = ProcessNull.GetInt32(arParameters(8).Value)
            w351to375 = ProcessNull.GetInt32(arParameters(9).Value)
            w376to400 = ProcessNull.GetInt32(arParameters(10).Value)
            greaterthan400 = ProcessNull.GetInt32(arParameters(11).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try

    End Function
#End Region


End Class 'dalAvgWeight
