
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalSite
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
Public Class dalSite

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum SiteFields
        fldID = 0
        fldSiteName = 1
        fldTitlePath = 2
        fldTitleText = 3
        fldSLabel = 4
        fldsubtext1 = 5
        fldsubtext2 = 6
        fldsubtext3 = 7
        fldsubtext4 = 8
        fldsubtext5 = 9
        fldsubtext6 = 10
        fldrptTitleHeader = 11
        fldrptTitleSubHeader = 12
        fldrptAltTitle = 13
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
    Public Function GetByKey(ByVal SiteID As Integer, _
                ByRef SiteName As String, _
                ByRef TitlePath As String, _
                ByRef TitleText As String, _
                ByRef SLabel As String, _
                ByRef subtext1 As String, _
                ByRef subtext2 As String, _
                ByRef subtext3 As String, _
                ByRef subtext4 As String, _
                ByRef subtext5 As String, _
                ByRef subtext6 As String, _
                ByRef rptTitleHeader As String, _
                ByRef rptTitleSubHeader As String, _
                ByRef rptAltTitle As String) As Boolean

        Dim TestNull As Object
        Dim arParameters(13) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.SiteFields.fldID) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(Me.SiteFields.fldID).Value = SiteID
        arParameters(Me.SiteFields.fldSiteName) = New SqlParameter("@SiteName", SqlDbType.NVarChar, 50)
        arParameters(Me.SiteFields.fldSiteName).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldTitlePath) = New SqlParameter("@TitlePath", SqlDbType.NVarChar, 255)
        arParameters(Me.SiteFields.fldTitlePath).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldTitleText) = New SqlParameter("@TitleText", SqlDbType.NVarChar, 100)
        arParameters(Me.SiteFields.fldTitleText).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldSLabel) = New SqlParameter("@SLabel", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldSLabel).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldsubtext1) = New SqlParameter("@subtext1", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext1).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldsubtext2) = New SqlParameter("@subtext2", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext2).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldsubtext3) = New SqlParameter("@subtext3", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext3).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldsubtext4) = New SqlParameter("@subtext4", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext4).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldsubtext5) = New SqlParameter("@subtext5", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext5).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldsubtext6) = New SqlParameter("@subtext6", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext6).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldrptTitleHeader) = New SqlParameter("@rptTitleHeader", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldrptTitleHeader).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldrptTitleSubHeader) = New SqlParameter("@rptTitleSubHeader", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldrptTitleSubHeader).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldrptAltTitle) = New SqlParameter("@rptAltTitle", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldrptAltTitle).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spSiteGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spSiteGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.SiteFields.fldSiteName).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            SiteName = ProcessNull.GetString(arParameters(Me.SiteFields.fldSiteName).Value)
            TitlePath = ProcessNull.GetString(arParameters(Me.SiteFields.fldTitlePath).Value)
            TitleText = ProcessNull.GetString(arParameters(Me.SiteFields.fldTitleText).Value)
            SLabel = ProcessNull.GetString(arParameters(Me.SiteFields.fldSLabel).Value)
            subtext1 = ProcessNull.GetString(arParameters(Me.SiteFields.fldsubtext1).Value)
            subtext2 = ProcessNull.GetString(arParameters(Me.SiteFields.fldsubtext2).Value)
            subtext3 = ProcessNull.GetString(arParameters(Me.SiteFields.fldsubtext3).Value)
            subtext4 = ProcessNull.GetString(arParameters(Me.SiteFields.fldsubtext4).Value)
            subtext5 = ProcessNull.GetString(arParameters(Me.SiteFields.fldsubtext5).Value)
            subtext6 = ProcessNull.GetString(arParameters(Me.SiteFields.fldsubtext6).Value)
            rptTitleHeader = ProcessNull.GetString(arParameters(Me.SiteFields.fldrptTitleHeader).Value)
            rptTitleSubHeader = ProcessNull.GetString(arParameters(Me.SiteFields.fldrptTitleSubHeader).Value)
            rptAltTitle = ProcessNull.GetString(arParameters(Me.SiteFields.fldrptAltTitle).Value)
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
    Public Function Update(ByVal SiteID As Integer, _
                ByVal SiteName As String, _
                ByVal TitlePath As String, _
                ByVal TitleText As String, _
                ByVal SLabel As String, _
                ByVal subtext1 As String, _
                ByVal subtext2 As String, _
                ByVal subtext3 As String, _
                ByVal subtext4 As String, _
                ByVal subtext5 As String, _
                ByVal subtext6 As String, _
                ByVal rptTitleHeader As String, _
                ByVal rptTitleSubHeader As String, _
                ByVal rptAltTitle As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(13) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.SiteFields.fldID) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(Me.SiteFields.fldID).Value = SiteID
        arParameters(Me.SiteFields.fldSiteName) = New SqlParameter("@SiteName", SqlDbType.NVarChar, 50)
        arParameters(Me.SiteFields.fldSiteName).Value = SiteName
        arParameters(Me.SiteFields.fldTitlePath) = New SqlParameter("@TitlePath", SqlDbType.NVarChar, 255)
        arParameters(Me.SiteFields.fldTitlePath).Value = TitlePath
        arParameters(Me.SiteFields.fldTitleText) = New SqlParameter("@TitleText", SqlDbType.NVarChar, 100)
        arParameters(Me.SiteFields.fldTitleText).Value = TitleText
        arParameters(Me.SiteFields.fldSLabel) = New SqlParameter("@SLabel", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldSLabel).Value = SLabel
        arParameters(Me.SiteFields.fldsubtext1) = New SqlParameter("@subtext1", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext1).Value = subtext1
        arParameters(Me.SiteFields.fldsubtext2) = New SqlParameter("@subtext2", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext2).Value = subtext2
        arParameters(Me.SiteFields.fldsubtext3) = New SqlParameter("@subtext3", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext3).Value = subtext3
        arParameters(Me.SiteFields.fldsubtext4) = New SqlParameter("@subtext4", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext4).Value = subtext4
        arParameters(Me.SiteFields.fldsubtext5) = New SqlParameter("@subtext5", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext5).Value = subtext5
        arParameters(Me.SiteFields.fldsubtext6) = New SqlParameter("@subtext6", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext6).Value = subtext6
        arParameters(Me.SiteFields.fldrptTitleHeader) = New SqlParameter("@rptTitleHeader", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldrptTitleHeader).Value = rptTitleHeader
        arParameters(Me.SiteFields.fldrptTitleSubHeader) = New SqlParameter("@rptTitleSubHeader", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldrptTitleSubHeader).Value = rptTitleSubHeader
        arParameters(Me.SiteFields.fldrptAltTitle) = New SqlParameter("@rptAltTitle", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldrptAltTitle).Value = rptAltTitle


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spSiteUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spSiteUpdate", arParameters)
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
    Public Function Add(ByRef SiteID As Integer, _
                ByVal SiteName As String, _
                ByVal TitlePath As String, _
                ByVal TitleText As String, _
                ByVal SLabel As String, _
                ByVal subtext1 As String, _
                ByVal subtext2 As String, _
                ByVal subtext3 As String, _
                ByVal subtext4 As String, _
                ByVal subtext5 As String, _
                ByVal subtext6 As String, _
                ByVal rptTitleHeader As String, _
                ByVal rptTitleSubHeader As String, _
                ByVal rptAltTitle As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(13) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.SiteFields.fldID) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(Me.SiteFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.SiteFields.fldSiteName) = New SqlParameter("@SiteName", SqlDbType.NVarChar, 50)
        arParameters(Me.SiteFields.fldSiteName).Value = SiteName
        arParameters(Me.SiteFields.fldTitlePath) = New SqlParameter("@TitlePath", SqlDbType.NVarChar, 255)
        arParameters(Me.SiteFields.fldTitlePath).Value = TitlePath
        arParameters(Me.SiteFields.fldTitleText) = New SqlParameter("@TitleText", SqlDbType.NVarChar, 100)
        arParameters(Me.SiteFields.fldTitleText).Value = TitleText
        arParameters(Me.SiteFields.fldSLabel) = New SqlParameter("@SLabel", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldSLabel).Value = SLabel
        arParameters(Me.SiteFields.fldsubtext1) = New SqlParameter("@subtext1", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext1).Value = subtext1
        arParameters(Me.SiteFields.fldsubtext2) = New SqlParameter("@subtext2", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext2).Value = subtext2
        arParameters(Me.SiteFields.fldsubtext3) = New SqlParameter("@subtext3", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext3).Value = subtext3
        arParameters(Me.SiteFields.fldsubtext4) = New SqlParameter("@subtext4", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext4).Value = subtext4
        arParameters(Me.SiteFields.fldsubtext5) = New SqlParameter("@subtext5", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext5).Value = subtext5
        arParameters(Me.SiteFields.fldsubtext6) = New SqlParameter("@subtext6", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldsubtext6).Value = rptTitleHeader
        arParameters(Me.SiteFields.fldrptTitleHeader) = New SqlParameter("@rptTitleHeader", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldrptTitleHeader).Value = rptTitleHeader
        arParameters(Me.SiteFields.fldrptTitleSubHeader) = New SqlParameter("@rptTitleSubHeader", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldrptTitleSubHeader).Value = rptTitleSubHeader
        arParameters(Me.SiteFields.fldrptAltTitle) = New SqlParameter("@rptAltTitle", SqlDbType.NVarChar, 200)
        arParameters(Me.SiteFields.fldrptAltTitle).Value = rptAltTitle

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spSiteInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spSiteInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            SiteID = CType(arParameters(0).Value, Integer)
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
        arParameters(0) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spSiteDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spSiteDelete", arParameters)
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


End Class 'dalSite
