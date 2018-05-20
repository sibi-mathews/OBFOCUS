
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalPerinatologist
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
Public Class dalPerinatologist

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum PerinatologistFields
        fldID = 0
        fldPFirstName = 1
        fldPMiddleName = 2
        fldPLastName = 3
        fldSpecialty = 4
        fldDEA = 5
        fldPhone = 6
        fldPAddress = 7
        fldPCity = 8
        fldPState = 9
        fldPZip = 10
        fldSiteID = 11
        fldTitle = 12
        fldLicense = 13
        fldBeeper = 14
        fldPerinatologist = 15
        fldSignaturePath = 16
        fldNonPhysician = 17
        fldSuppress = 18
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
    Public Function GetByKey(ByVal ExaminerID As Integer, _
                    ByRef PFirstName As String, _
                    ByRef PMiddleName As String, _
                    ByRef PLastName As String, _
                    ByRef Specialty As String, _
                    ByRef DEA As String, _
                    ByRef Phone As String, _
                    ByRef PAddress As String, _
                    ByRef PCity As String, _
                    ByRef PState As String, _
                    ByRef PZip As String, _
                    ByRef SiteID As Integer, _
                    ByRef Title As String, _
                    ByRef License As String, _
                    ByRef Beeper As String, _
                    ByRef Perinatologist As Short, _
                    ByRef SignaturePath As String, _
                    ByRef NonPhysician As Short, _
                    ByRef Suppress As Short) As Boolean

        Dim arParameters(18) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(Me.PerinatologistFields.fldID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.PerinatologistFields.fldID).Value = ExaminerID
        arParameters(Me.PerinatologistFields.fldPFirstName) = New SqlParameter("@PFirstName", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPFirstName).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldPMiddleName) = New SqlParameter("@PMiddleName", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPMiddleName).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldPLastName) = New SqlParameter("@PLastName", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPLastName).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldSpecialty) = New SqlParameter("@Specialty", SqlDbType.NVarChar, 255)
        arParameters(Me.PerinatologistFields.fldSpecialty).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldDEA) = New SqlParameter("@DEA", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldDEA).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldPhone) = New SqlParameter("@HomePhone", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPhone).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldPAddress) = New SqlParameter("@PAddress", SqlDbType.NVarChar, 255)
        arParameters(Me.PerinatologistFields.fldPAddress).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldPCity) = New SqlParameter("@PCity", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPCity).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldPState) = New SqlParameter("@PStateorProvince", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPState).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldPZip) = New SqlParameter("@PPostalCode", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPZip).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldSiteID) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(Me.PerinatologistFields.fldSiteID).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldTitle) = New SqlParameter("@Title", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldTitle).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldLicense) = New SqlParameter("@License", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldLicense).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldBeeper) = New SqlParameter("@Beeper", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldBeeper).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldPerinatologist) = New SqlParameter("@Perinatologist", SqlDbType.SmallInt)
        arParameters(Me.PerinatologistFields.fldPerinatologist).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldSignaturePath) = New SqlParameter("@SignaturePath", SqlDbType.NVarChar, 255)
        arParameters(Me.PerinatologistFields.fldSignaturePath).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldNonPhysician) = New SqlParameter("@NonPhysician", SqlDbType.Bit)
        arParameters(Me.PerinatologistFields.fldNonPhysician).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldSuppress) = New SqlParameter("@Suppress", SqlDbType.Bit)
        arParameters(Me.PerinatologistFields.fldSuppress).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminerGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExaminerGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(0).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            PFirstName = ProcessNull.GetString(arParameters(1).Value)
            PFirstName = PFirstName.Trim()
            PMiddleName = ProcessNull.GetString(arParameters(2).Value)
            PMiddleName = PMiddleName.Trim()
            PLastName = ProcessNull.GetString(arParameters(3).Value)
            PLastName = PLastName.Trim()
            Specialty = ProcessNull.GetString(arParameters(4).Value)
            Specialty = Specialty.Trim()
            DEA = ProcessNull.GetString(arParameters(5).Value)
            DEA = DEA.Trim()
            Phone = ProcessNull.GetString(arParameters(6).Value)
            Phone = Phone.Trim()
            PAddress = ProcessNull.GetString(arParameters(7).Value)
            PAddress = PAddress.Trim()
            PCity = ProcessNull.GetString(arParameters(8).Value)
            PCity = PCity.Trim()
            PState = ProcessNull.GetString(arParameters(9).Value)
            PState = PState.Trim()
            PZip = ProcessNull.GetString(arParameters(10).Value)
            PZip = PZip.Trim()
            SiteID = ProcessNull.GetInt32(arParameters(11).Value)
            Title = ProcessNull.GetString(arParameters(12).Value)
            Title = Title.Trim()
            License = ProcessNull.GetString(arParameters(13).Value)
            License = License.Trim()
            Beeper = ProcessNull.GetString(arParameters(14).Value)
            Beeper = Beeper.Trim()
            Perinatologist = ProcessNull.GetInt16(arParameters(15).Value)
            SignaturePath = ProcessNull.GetString(arParameters(16).Value)
            SignaturePath = SignaturePath.Trim()
            NonPhysician = ProcessNull.GetInt16(arParameters(17).Value)
            Suppress = ProcessNull.GetInt16(arParameters(18).Value)
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
    Public Function Update(ByVal ExaminerID As Integer, _
               ByVal PFirstName As String, _
               ByVal PMiddleName As String, _
               ByVal PLastName As String, _
               ByVal Specialty As String, _
               ByVal DEA As String, _
               ByVal Phone As String, _
               ByVal PAddress As String, _
               ByVal PCity As String, _
               ByVal PState As String, _
               ByVal PZip As String, _
               ByVal SiteID As Integer, _
               ByVal Title As String, _
               ByVal License As String, _
               ByVal Beeper As String, _
               ByVal Perinatologist As Short, _
               ByVal SignaturePath As String, _
               ByVal NonPhysician As Short, _
               ByVal Suppress As Short) As Boolean


        Dim intRecordsAffected As Integer = 0
        Dim arParameters(18) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(Me.PerinatologistFields.fldID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.PerinatologistFields.fldID).Value = ExaminerID
        arParameters(Me.PerinatologistFields.fldPFirstName) = New SqlParameter("@PFirstName", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPFirstName).Value = PFirstName
        arParameters(Me.PerinatologistFields.fldPMiddleName) = New SqlParameter("@PMiddleName", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPMiddleName).Value = PMiddleName
        arParameters(Me.PerinatologistFields.fldPLastName) = New SqlParameter("@PLastName", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPLastName).Value = PLastName
        arParameters(Me.PerinatologistFields.fldSpecialty) = New SqlParameter("@Specialty", SqlDbType.NVarChar, 255)
        arParameters(Me.PerinatologistFields.fldSpecialty).Value = Specialty
        arParameters(Me.PerinatologistFields.fldDEA) = New SqlParameter("@DEA", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldDEA).Value = DEA
        arParameters(Me.PerinatologistFields.fldPhone) = New SqlParameter("@HomePhone", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPhone).Value = Phone
        arParameters(Me.PerinatologistFields.fldPAddress) = New SqlParameter("@PAddress", SqlDbType.NVarChar, 255)
        arParameters(Me.PerinatologistFields.fldPAddress).Value = PAddress
        arParameters(Me.PerinatologistFields.fldPCity) = New SqlParameter("@PCity", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPCity).Value = PCity
        arParameters(Me.PerinatologistFields.fldPState) = New SqlParameter("@PStateorProvince", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPState).Value = PState
        arParameters(Me.PerinatologistFields.fldPZip) = New SqlParameter("@PPostalCode", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPZip).Value = PZip
        arParameters(Me.PerinatologistFields.fldSiteID) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(Me.PerinatologistFields.fldSiteID).Value = SiteID
        arParameters(Me.PerinatologistFields.fldTitle) = New SqlParameter("@Title", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldTitle).Value = Title
        arParameters(Me.PerinatologistFields.fldLicense) = New SqlParameter("@License", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldLicense).Value = License
        arParameters(Me.PerinatologistFields.fldBeeper) = New SqlParameter("@Beeper", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldBeeper).Value = Beeper
        arParameters(Me.PerinatologistFields.fldPerinatologist) = New SqlParameter("@Perinatologist", SqlDbType.SmallInt)
        arParameters(Me.PerinatologistFields.fldPerinatologist).Value = Perinatologist
        arParameters(Me.PerinatologistFields.fldSignaturePath) = New SqlParameter("@SignaturePath", SqlDbType.NVarChar, 255)
        arParameters(Me.PerinatologistFields.fldSignaturePath).Value = SignaturePath
        arParameters(Me.PerinatologistFields.fldNonPhysician) = New SqlParameter("@NonPhysician", SqlDbType.Bit)
        arParameters(Me.PerinatologistFields.fldNonPhysician).Value = NonPhysician
        arParameters(Me.PerinatologistFields.fldSuppress) = New SqlParameter("@Suppress", SqlDbType.Bit)
        arParameters(Me.PerinatologistFields.fldSuppress).Value = Suppress

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminerUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExaminerUpdate", arParameters)
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
    Public Function Add(ByRef ExaminerID As Integer, _
               ByVal PFirstName As String, _
               ByVal PMiddleName As String, _
               ByVal PLastName As String, _
               ByVal Specialty As String, _
               ByVal DEA As String, _
               ByVal Phone As String, _
               ByVal PAddress As String, _
               ByVal PCity As String, _
               ByVal PState As String, _
               ByVal PZip As String, _
               ByVal SiteID As Integer, _
               ByVal Title As String, _
               ByVal License As String, _
               ByVal Beeper As String, _
               ByVal Perinatologist As Short, _
               ByVal SignaturePath As String, _
               ByVal NonPhysician As Short, _
               ByVal Suppress As Short) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(18) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(Me.PerinatologistFields.fldID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.PerinatologistFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.PerinatologistFields.fldPFirstName) = New SqlParameter("@PFirstName", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPFirstName).Value = PFirstName
        arParameters(Me.PerinatologistFields.fldPMiddleName) = New SqlParameter("@PMiddleName", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPMiddleName).Value = PMiddleName
        arParameters(Me.PerinatologistFields.fldPLastName) = New SqlParameter("@PLastName", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPLastName).Value = PLastName
        arParameters(Me.PerinatologistFields.fldSpecialty) = New SqlParameter("@Specialty", SqlDbType.NVarChar, 255)
        arParameters(Me.PerinatologistFields.fldSpecialty).Value = Specialty
        arParameters(Me.PerinatologistFields.fldDEA) = New SqlParameter("@DEA", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldDEA).Value = DEA
        arParameters(Me.PerinatologistFields.fldPhone) = New SqlParameter("@HomePhone", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPhone).Value = Phone
        arParameters(Me.PerinatologistFields.fldPAddress) = New SqlParameter("@PAddress", SqlDbType.NVarChar, 255)
        arParameters(Me.PerinatologistFields.fldPAddress).Value = PAddress
        arParameters(Me.PerinatologistFields.fldPCity) = New SqlParameter("@PCity", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPCity).Value = PCity
        arParameters(Me.PerinatologistFields.fldPState) = New SqlParameter("@PStateorProvince", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPState).Value = PState
        arParameters(Me.PerinatologistFields.fldPZip) = New SqlParameter("@PPostalCode", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldPZip).Value = PZip
        arParameters(Me.PerinatologistFields.fldSiteID) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(Me.PerinatologistFields.fldSiteID).Value = SiteID
        arParameters(Me.PerinatologistFields.fldTitle) = New SqlParameter("@Title", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldTitle).Value = Title
        arParameters(Me.PerinatologistFields.fldLicense) = New SqlParameter("@License", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldLicense).Value = License
        arParameters(Me.PerinatologistFields.fldBeeper) = New SqlParameter("@Beeper", SqlDbType.NVarChar, 50)
        arParameters(Me.PerinatologistFields.fldBeeper).Value = Beeper
        arParameters(Me.PerinatologistFields.fldPerinatologist) = New SqlParameter("@Perinatologist", SqlDbType.SmallInt)
        arParameters(Me.PerinatologistFields.fldPerinatologist).Value = Perinatologist
        arParameters(Me.PerinatologistFields.fldSignaturePath) = New SqlParameter("@SignaturePath", SqlDbType.NVarChar, 255)
        arParameters(Me.PerinatologistFields.fldSignaturePath).Value = SignaturePath
        arParameters(Me.PerinatologistFields.fldNonPhysician) = New SqlParameter("@NonPhysician", SqlDbType.Bit)
        arParameters(Me.PerinatologistFields.fldNonPhysician).Value = NonPhysician
        arParameters(Me.PerinatologistFields.fldSuppress) = New SqlParameter("@Suppress", SqlDbType.Bit)
        arParameters(Me.PerinatologistFields.fldSuppress).Value = Suppress

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminerInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExaminerInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            ExaminerID = CType(arParameters(0).Value, Integer)
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
        arParameters(Me.PerinatologistFields.fldID) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(Me.PerinatologistFields.fldID).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminerDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExaminerDelete", arParameters)
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


End Class 'dalPerinatologist
