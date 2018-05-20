
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalWord
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
Public Class dalWord

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



#Region "Main procedures - GetWord, Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Name:        GetWord
    '*
    '* Description: Returns all records in the [ZTables] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetWord() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        RunSP
    '*
    '* Description: Returns all records in the [ZTables] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function RunSP(ByVal SqlString As String) As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.Text, SqlString)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.Text, SqlString)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetWordSP
    '*
    '* Description: Returns all records in the [ZTables] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetWordSP() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateSPGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateSPGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetWordSPSub
    '*
    '* Description: Returns all records in the [ZTables] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetWordSPSub(ByVal UID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(0).Value = UID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateSPSubGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateSPSubGet", arParameters)
            End If
            ' Call stored procedure and return the data
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
                ByRef TemplateName As String, _
                ByRef TemplateDescrip As String, _
                ByRef StoredProc As String, _
                ByRef StoredProcID As Integer, _
                ByRef TemplatePath As String, _
                ByRef StoredProcDescrip As String, _
                ByRef Bookmarks As String, _
                ByRef Macros As String, _
                ByRef ProtectDocForm As Short, _
                ByRef DocRecLabTypeID As Integer, _
                ByRef DocRecLabType As String) As Boolean
        Dim arParameters(11) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@TemplateName", SqlDbType.NVarChar, 50)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@TemplateDescrip", SqlDbType.NVarChar, 255)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@StoredProc", SqlDbType.NVarChar, 100)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@StoredProcID", SqlDbType.Int)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@TemplatePath", SqlDbType.NVarChar, 500)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@StoredProcDescrip", SqlDbType.NVarChar, 255)
        arParameters(6).Direction = ParameterDirection.Output
        arParameters(7) = New SqlParameter("@Bookmarks", SqlDbType.NVarChar, 1000)
        arParameters(7).Direction = ParameterDirection.Output
        arParameters(8) = New SqlParameter("@Macros", SqlDbType.NVarChar, 500)
        arParameters(8).Direction = ParameterDirection.Output
        arParameters(9) = New SqlParameter("@ProtectDocForm", SqlDbType.Bit)
        arParameters(9).Direction = ParameterDirection.Output
        arParameters(10) = New SqlParameter("@DocRecLabTypeID", SqlDbType.Int)
        arParameters(10).Direction = ParameterDirection.Output
        arParameters(11) = New SqlParameter("@DocRecLabType", SqlDbType.NVarChar, 50)
        arParameters(11).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateGetbyKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateGetbyKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(1).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            TemplateName = ProcessNull.GetString(arParameters(1).Value)
            TemplateDescrip = ProcessNull.GetString(arParameters(2).Value)
            StoredProc = ProcessNull.GetString(arParameters(3).Value)
            StoredProcID = ProcessNull.GetInt32(arParameters(4).Value)
            TemplatePath = ProcessNull.GetString(arParameters(5).Value)
            StoredProcDescrip = ProcessNull.GetString(arParameters(6).Value)
            Bookmarks = ProcessNull.GetString(arParameters(7).Value)
            Macros = ProcessNull.GetString(arParameters(8).Value)
            ProtectDocForm = ProcessNull.GetInt16(arParameters(9).Value)
            DocRecLabTypeID = ProcessNull.GetInt32(arParameters(10).Value)
            DocRecLabType = ProcessNull.GetString(arParameters(11).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        GetSPByKey
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
    Public Function GetSPByKey(ByVal StoredProcID As Integer, _
                ByRef StoredProc As String, _
                ByRef StoredProcDescrip As String, _
                ByRef Bookmarks As String) As Boolean
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@StoredProcID", SqlDbType.Int)
        arParameters(0).Value = StoredProcID
        arParameters(1) = New SqlParameter("@StoredProc", SqlDbType.NVarChar, 100)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@StoredProcDescrip", SqlDbType.NVarChar, 255)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@Bookmarks", SqlDbType.NVarChar, 1000)
        arParameters(3).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateSPGetbyKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateSPGetbyKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(1).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            StoredProc = ProcessNull.GetString(arParameters(1).Value)
            StoredProcDescrip = ProcessNull.GetString(arParameters(2).Value)
            Bookmarks = ProcessNull.GetString(arParameters(3).Value)
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
    '* Description: Gets all the values of a record identified by a key.
    '*
    '* Parameters:  Description - Output parameter
    '*              Picture - Output parameter
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal ID As Integer, _
                ByVal TemplateName As String, _
                ByVal TemplateDescrip As String, _
                ByVal StoredProcID As Integer, _
                ByVal TemplatePath As String, _
                ByVal Macros As String, _
                ByVal ProtectDocForm As Short, _
                ByVal DocRecLabTypeID As Integer) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(7) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@TemplateName", SqlDbType.NVarChar, 50)
        arParameters(0).Value = TemplateName
        arParameters(1) = New SqlParameter("@TemplateDescrip", SqlDbType.NVarChar, 255)
        arParameters(1).Value = TemplateDescrip
        arParameters(2) = New SqlParameter("@StoredProcID", SqlDbType.Int)
        arParameters(2).Value = StoredProcID
        arParameters(3) = New SqlParameter("@TemplatePath", SqlDbType.NVarChar, 500)
        arParameters(3).Value = TemplatePath
        arParameters(4) = New SqlParameter("@Macros", SqlDbType.NVarChar, 500)
        arParameters(4).Value = Macros
        arParameters(5) = New SqlParameter("@ProtectDocForm", SqlDbType.Bit)
        arParameters(5).Value = ProtectDocForm
        arParameters(6) = New SqlParameter("@DocRecLabTypeID", SqlDbType.Int)
        arParameters(6).Value = DocRecLabTypeID
        arParameters(7) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(7).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateUpd", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateUpd", arParameters)
            End If
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


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
                ByVal TemplateName As String, _
                ByVal TemplateDescrip As String, _
                ByVal StoredProcID As Integer, _
                ByVal TemplatePath As String, _
                ByVal Macros As String, _
                ByVal ProtectDocForm As Short, _
                ByVal DocRecLabTypeID As Integer) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(7) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@TemplateName", SqlDbType.NVarChar, 50)
        arParameters(0).Value = TemplateName
        arParameters(1) = New SqlParameter("@TemplateDescrip", SqlDbType.NVarChar, 255)
        arParameters(1).Value = TemplateDescrip
        arParameters(2) = New SqlParameter("@StoredProcID", SqlDbType.Int)
        arParameters(2).Value = StoredProcID
        arParameters(3) = New SqlParameter("@TemplatePath", SqlDbType.NVarChar, 500)
        arParameters(3).Value = TemplatePath
        arParameters(4) = New SqlParameter("@Macros", SqlDbType.NVarChar, 500)
        arParameters(4).Value = Macros
        arParameters(5) = New SqlParameter("@ProtectDocForm", SqlDbType.Bit)
        arParameters(5).Value = ProtectDocForm
        arParameters(6) = New SqlParameter("@DocRecLabTypeID", SqlDbType.Int)
        arParameters(6).Value = DocRecLabTypeID
        arParameters(7) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(7).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateAdd", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateAdd", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            ID = CType(arParameters(6).Value, Integer)
            Return True
        End If
    End Function

    '**************************************************************************
    '*  
    '* Name:        AddSP
    '*
    '* Description: AddSPs a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was AddSPed or not. 
    '*              True (record AddSPed); False (otherwise).
    '*
    '**************************************************************************
    Public Function AddSP(ByRef StoredProcID As Integer, _
                            ByVal StoredProc As String, _
                            ByVal StoredProcDescrip As String, _
                            ByVal Bookmarks As String) As Boolean

        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@StoredProcID", SqlDbType.Int)
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@StoredProc", SqlDbType.NVarChar, 100)
        arParameters(1).Value = StoredProc
        arParameters(2) = New SqlParameter("@StoredProcDescrip", SqlDbType.NVarChar, 255)
        arParameters(2).Value = StoredProcDescrip
        arParameters(3) = New SqlParameter("@Bookmarks", SqlDbType.NVarChar, 1000)
        arParameters(3).Value = Bookmarks
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateSpAdd", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateSpAdd", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            StoredProcID = CType(arParameters(0).Value, Integer)
            Return True
        End If

    End Function
    '**************************************************************************
    '*  
    '* Name:        AddSPSub
    '*
    '* Description: AddSPSubs a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was AddSPSubed or not. 
    '*              True (record AddSPSubed); False (otherwise).
    '*
    '**************************************************************************
    Public Function AddSPSub(ByRef ID As Integer, _
                            ByVal StoredProcID As Integer, _
                            ByVal Bookmarks As String) As Boolean

        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ID", SqlDbType.Int)
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@StoredProcID", SqlDbType.Int)
        arParameters(1).Value = StoredProcID
        arParameters(2) = New SqlParameter("@Bookmark", SqlDbType.NVarChar, 100)
        arParameters(2).Value = Bookmarks
        arParameters(3) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(3).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateSpSubAdd", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateSpSubAdd", arParameters)
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
    '* Name:        UpdateSP
    '*
    '* Description: UpdateSPs a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was UpdateSPed or not. 
    '*              True (record UpdateSPed); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdateSP(ByVal StoredProcID As Integer, _
                            ByVal StoredProc As String, _
                            ByVal StoredProcDescrip As String, _
                            ByVal Bookmarks As String) As Boolean

        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@StoredProcID", SqlDbType.Int)
        arParameters(0).Value = StoredProcID
        arParameters(1) = New SqlParameter("@StoredProc", SqlDbType.NVarChar, 100)
        arParameters(1).Value = StoredProc
        arParameters(2) = New SqlParameter("@StoredProcDescrip", SqlDbType.NVarChar, 255)
        arParameters(2).Value = StoredProcDescrip
        arParameters(3) = New SqlParameter("@Bookmarks", SqlDbType.NVarChar, 1000)
        arParameters(3).Value = Bookmarks
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateSpUpd", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateSpUpd", arParameters)
            End If
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        UpdateSPSub
    '*
    '* Description: UpdateSPSubs a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was UpdateSPSubed or not. 
    '*              True (record UpdateSPSubed); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdateSPSub(ByVal ID As Integer, _
                            ByVal StoredProcID As Integer, _
                            ByVal Bookmarks As String) As Boolean

        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@StoredProcID", SqlDbType.Int)
        arParameters(1).Value = StoredProcID
        arParameters(2) = New SqlParameter("@Bookmark", SqlDbType.NVarChar, 100)
        arParameters(2).Value = Bookmarks

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateSpSubUpd", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateSpSubUpd", arParameters)
            End If
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


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
        arParameters(0) = New SqlParameter("@UID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateDel", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateDel", arParameters)
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
    '* Name:        DeleteSP
    '*
    '* Description: DeleteSPs a record from the [PatientInfo] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to DeleteSP
    '*
    '* Returns:     Boolean indicating if record was DeleteSPd or not. 
    '*              True (record found and DeleteSPd); False (otherwise).
    '*
    '**************************************************************************
    Public Function DeleteSP(ByVal ID As Integer) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@StoredProcID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateSpDel", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateSpDel", arParameters)
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
    '* Name:        DeleteSPSub
    '*
    '* Description: DeleteSPSubs a record from the [PatientInfo] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to DeleteSPSub
    '*
    '* Returns:     Boolean indicating if record was DeleteSPSubd or not. 
    '*              True (record found and DeleteSPSubd); False (otherwise).
    '*
    '**************************************************************************
    Public Function DeleteSPSub(ByVal ID As Integer) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spWordTemplateSpSubDel", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spWordTemplateSpSubDel", arParameters)
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
End Class 'dalWord
