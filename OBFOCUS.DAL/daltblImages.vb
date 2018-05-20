
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalTblImages
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
Public Class dalTblImages

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum tblImagesFields
        fldID = 0
        fldImageCategory = 1
        fldImageDescription = 2
        fldNarrative = 3
        fldReference = 4
        fldImagePath = 5
        fldSource = 6
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
    Public Function GetByKey(ByVal ImageID As Integer, _
                ByRef ImageCategory As String, _
                ByRef ImageDescription As String, _
                ByRef Narrative As String, _
                ByRef Reference As String, _
                ByRef ImagePath As String, _
                ByRef Source As String) As Boolean

        Dim arParameters(6) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.tblImagesFields.fldID) = New SqlParameter("@ImageID", SqlDbType.Int)
        arParameters(Me.tblImagesFields.fldID).Value = ImageID
        arParameters(Me.tblImagesFields.fldImageCategory) = New SqlParameter("@ImageCategory", SqlDbType.NVarChar, 50)
        arParameters(Me.tblImagesFields.fldImageCategory).Direction = ParameterDirection.Output
        arParameters(Me.tblImagesFields.fldImageDescription) = New SqlParameter("@ImageDescription", SqlDbType.NVarChar, 50)
        arParameters(Me.tblImagesFields.fldImageDescription).Direction = ParameterDirection.Output
        arParameters(Me.tblImagesFields.fldNarrative) = New SqlParameter("@Narrative", SqlDbType.VarChar, 8000)
        arParameters(Me.tblImagesFields.fldNarrative).Direction = ParameterDirection.Output
        arParameters(Me.tblImagesFields.fldReference) = New SqlParameter("@Reference", SqlDbType.VarChar, 1000)
        arParameters(Me.tblImagesFields.fldReference).Direction = ParameterDirection.Output
        arParameters(Me.tblImagesFields.fldImagePath) = New SqlParameter("@ImagePath", SqlDbType.NVarChar, 255)
        arParameters(Me.tblImagesFields.fldImagePath).Direction = ParameterDirection.Output
        arParameters(Me.tblImagesFields.fldSource) = New SqlParameter("@Source", SqlDbType.VarChar, 1000)
        arParameters(Me.tblImagesFields.fldSource).Direction = ParameterDirection.Output


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "sptblImagesGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "sptblImagesGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.tblImagesFields.fldImageCategory).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            ImageCategory = ProcessNull.GetString(arParameters(Me.tblImagesFields.fldImageCategory).Value)
            ImageDescription = ProcessNull.GetString(arParameters(Me.tblImagesFields.fldImageDescription).Value)
            Narrative = ProcessNull.GetString(arParameters(Me.tblImagesFields.fldNarrative).Value)
            Reference = ProcessNull.GetString(arParameters(Me.tblImagesFields.fldReference).Value)
            ImagePath = ProcessNull.GetString(arParameters(Me.tblImagesFields.fldImagePath).Value)
            Source = ProcessNull.GetString(arParameters(Me.tblImagesFields.fldSource).Value)
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
    Public Function Update(ByVal ImageID As Integer, _
                ByVal ImageCategory As String, _
                ByVal ImageDescription As String, _
                ByVal Narrative As String, _
                ByVal Reference As String, _
                ByVal ImagePath As String, _
                ByVal Source As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(6) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.tblImagesFields.fldID) = New SqlParameter("@ImageID", SqlDbType.Int)
        arParameters(Me.tblImagesFields.fldID).Value = ImageID
        arParameters(Me.tblImagesFields.fldImageCategory) = New SqlParameter("@ImageCategory", SqlDbType.NVarChar, 50)
        arParameters(Me.tblImagesFields.fldImageCategory).Value = ImageCategory
        arParameters(Me.tblImagesFields.fldImageDescription) = New SqlParameter("@ImageDescription", SqlDbType.NVarChar, 50)
        arParameters(Me.tblImagesFields.fldImageDescription).Value = ImageDescription
        arParameters(Me.tblImagesFields.fldNarrative) = New SqlParameter("@Narrative", SqlDbType.VarChar, 8000)
        arParameters(Me.tblImagesFields.fldNarrative).Value = Narrative
        arParameters(Me.tblImagesFields.fldReference) = New SqlParameter("@Reference", SqlDbType.VarChar, 1000)
        arParameters(Me.tblImagesFields.fldReference).Value = Reference
        arParameters(Me.tblImagesFields.fldImagePath) = New SqlParameter("@ImagePath", SqlDbType.NVarChar, 255)
        arParameters(Me.tblImagesFields.fldImagePath).Value = ImagePath
        arParameters(Me.tblImagesFields.fldSource) = New SqlParameter("@Source", SqlDbType.VarChar, 1000)
        arParameters(Me.tblImagesFields.fldSource).Value = Source


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "sptblImagesUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "sptblImagesUpdate", arParameters)
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
    Public Function Add(ByRef ImageID As Integer, _
                ByVal ImageCategory As String, _
                ByVal ImageDescription As String, _
                ByVal Narrative As String, _
                ByVal Reference As String, _
                ByVal ImagePath As String, _
                ByVal Source As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(6) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.tblImagesFields.fldID) = New SqlParameter("@ImageID", SqlDbType.Int)
        arParameters(Me.tblImagesFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.tblImagesFields.fldImageCategory) = New SqlParameter("@ImageCategory", SqlDbType.NVarChar, 50)
        arParameters(Me.tblImagesFields.fldImageCategory).Value = ImageCategory
        arParameters(Me.tblImagesFields.fldImageDescription) = New SqlParameter("@ImageDescription", SqlDbType.NVarChar, 50)
        arParameters(Me.tblImagesFields.fldImageDescription).Value = ImageDescription
        arParameters(Me.tblImagesFields.fldNarrative) = New SqlParameter("@Narrative", SqlDbType.VarChar, 8000)
        arParameters(Me.tblImagesFields.fldNarrative).Value = Narrative
        arParameters(Me.tblImagesFields.fldReference) = New SqlParameter("@Reference", SqlDbType.VarChar, 1000)
        arParameters(Me.tblImagesFields.fldReference).Value = Reference
        arParameters(Me.tblImagesFields.fldImagePath) = New SqlParameter("@ImagePath", SqlDbType.NVarChar, 255)
        arParameters(Me.tblImagesFields.fldImagePath).Value = ImagePath
        arParameters(Me.tblImagesFields.fldSource) = New SqlParameter("@Source", SqlDbType.VarChar, 1000)
        arParameters(Me.tblImagesFields.fldSource).Value = Source

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "sptblImagesInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "sptblImagesInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            ImageID = CType(arParameters(0).Value, Integer)
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
        arParameters(0) = New SqlParameter("@ImageID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "sptblImagesDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "sptblImagesDelete", arParameters)
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


End Class 'dalTblImages
