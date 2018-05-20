Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Formulary:        dalIPSettings
'*
'* Class: Data access layer for Table tblIPSEttings
'*
'* Remarks:     Uses OleDb and embedded SQL for maintaining the data.
'*-----------------------------------------------------------------------------
'*                      CHANGE HISTORY
'*   Change No:   Date:          Author:   Class:
'*   _________    ___________    ______    ____________________________________
'*      001       10/3/10       MBR        Created.                                
'* 
'******************************************************************************

Public Class dalIPSettings


#Region "Module level variables and enums"


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
    '* Description:        New
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
    '* Allergy:        GetIPSettings
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
    Public Function GetIPSettings(ByVal IPAddress As String, ByVal Mode As String, _
                                    ByRef ScreenTop As Integer, _
                                    ByRef ScreenLeft As Integer, _
                                    ByRef ScreenWidth As Integer, _
                                    ByRef ScreenHeight As Integer) As Boolean

        Dim arParameters(5) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@IPAddress", SqlDbType.NVarChar, 50)
        arParameters(0).Value = IPAddress
        arParameters(1) = New SqlParameter("@Mode", SqlDbType.NVarChar, 50)
        arParameters(1).Value = Mode
        arParameters(2) = New SqlParameter("@ScreenTop", SqlDbType.Int)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@ScreenLeft", SqlDbType.Int)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@ScreenWidth", SqlDbType.Int)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@ScreenHeight", SqlDbType.Int)
        arParameters(5).Direction = ParameterDirection.Output
       
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spIPSettingsGet", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spIPSettingsGet", arParameters)
            End If


            ' Return False if data was not found.
            'If arParameters(Me.ClasssFields.fldGeneric).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            ScreenTop = ProcessNull.GetInt16(arParameters(2).Value)
            ScreenLeft = ProcessNull.GetInt16(arParameters(3).Value)
            ScreenWidth = ProcessNull.GetInt16(arParameters(4).Value)
            ScreenHeight = ProcessNull.GetInt16(arParameters(5).Value)
           

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
    Public Function Update(ByVal IPAddress As String, _
                ByVal Mode As String, _
                ByVal ScreenTop As Integer, _
                ByVal ScreenLeft As Integer, _
                ByVal ScreenWidth As Integer, _
                ByVal ScreenHeight As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(5) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@IpAddress", SqlDbType.NVarChar, 50)
        arParameters(0).Value = IPAddress
        arParameters(1) = New SqlParameter("@Mode", SqlDbType.NVarChar, 50)
        arParameters(1).Value = Mode
        arParameters(2) = New SqlParameter("@ScreenTop", SqlDbType.Int)
        arParameters(2).Value = ScreenTop
        arParameters(3) = New SqlParameter("@ScreenLeft", SqlDbType.Int)
        arParameters(3).Value = ScreenLeft
        arParameters(4) = New SqlParameter("@ScreenWidth", SqlDbType.Int)
        arParameters(4).Value = ScreenWidth
        arParameters(5) = New SqlParameter("@ScreenHeight", SqlDbType.Int)
        arParameters(5).Value = ScreenHeight
       
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spIPSettingsUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spIPSettingsUpdate", arParameters)
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


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


End Class 'dalIPSettings
