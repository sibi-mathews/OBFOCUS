
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalBilling
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
Public Class dalBilling

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum BillingFields
        fldChartID = 0
        fldID = 1
        fldAddress1 = 2
        fldPHone = 3
        fldCity = 4
        fldState = 5
        fldZip = 6
        fldEmPerson = 7
        fldEmPHone = 8
        fldEmployedAs = 9
        fldEmployer = 10
        fldEAddress = 11
        fldECity = 12
        fldEState = 13
        fldEZip = 14
        fldEPhone = 15
        fldPrimaryInsurance = 16
        fldPIZip = 17
        fldIDNo = 18
        fldPIAddress = 19
        fldPICity = 20
        fldPIState = 21
        fldPlaceOfBirth = 22
        fldRace = 23
        fldLanguage = 24
        fldDriverLic = 25
        fldSocialSecurity = 26
        fldReligion = 27
        fldCellPhone = 28
        fldMaritalStatus = 29
        fldAKA = 30
        fldSpecialNeeds = 31
        fldEmail = 32
        fldPager = 33
        fldEExt = 34
        fldDOB = 35
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
    Public Function GetByKey(ByVal ChartID As Integer, _
                ByRef ID As Integer, _
                ByRef Address1 As String, _
                ByRef Phone As String, _
                ByRef City As String, _
                ByRef State As String, _
                ByRef Zip As String, _
                ByRef EmPerson As String, _
                ByRef EmPHone As String, _
                ByRef EmployedAs As String, _
                ByRef Employer As String, _
                ByRef EAddress As String, _
                ByRef ECity As String, _
                ByRef EState As String, _
                ByRef EZip As String, _
                ByRef EPhone As String, _
                ByRef PrimaryInsurance As String, _
                ByRef PIZip As String, _
                ByRef IDNo As String, _
                ByRef PIAddress As String, _
                ByRef PICity As String, _
                ByRef PIState As String, _
                ByRef PlaceOfBirth As String, _
                ByRef Race As String, _
                ByRef Language As String, _
                ByRef DriverLic As String, _
                ByRef SocialSecurity As String, _
                ByRef Religion As String, _
                ByRef CellPhone As String, _
                ByRef MaritalStatus As String, _
                ByRef AKA As String, _
                ByRef SpecialNeeds As String, _
                ByRef Email As String, _
                ByRef Pager As String, _
                ByRef EExt As String, _
                ByRef DOB As String) As Boolean
        ' Set the stored procedure parameters
        Dim arParameters(35) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.BillingFields.fldChartID) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(Me.BillingFields.fldChartID).Value = ChartID
        arParameters(Me.BillingFields.fldID) = New SqlParameter("@BillingID", SqlDbType.Int)
        arParameters(Me.BillingFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldAddress1) = New SqlParameter("@Address1", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldAddress1).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldPHone) = New SqlParameter("@PHone", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldPHone).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldCity) = New SqlParameter("@City", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldCity).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldState) = New SqlParameter("@State", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldState).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldZip) = New SqlParameter("@Zip", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldZip).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldEmPerson) = New SqlParameter("@EMPerson", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldEmPerson).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldEmPHone) = New SqlParameter("@EMPhone", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldEmPHone).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldEmployedAs) = New SqlParameter("@EmployedAs", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldEmployedAs).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldEmployer) = New SqlParameter("@Employer", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldEmployer).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldEAddress) = New SqlParameter("@EAddress", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldEAddress).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldECity) = New SqlParameter("@ECity", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldECity).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldEState) = New SqlParameter("@Estate", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldEState).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldEZip) = New SqlParameter("@EZip", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldEZip).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldEPhone) = New SqlParameter("@EPHone", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldEPhone).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldPrimaryInsurance) = New SqlParameter("@PrimaryInsurance", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldPrimaryInsurance).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldPIZip) = New SqlParameter("@PIZip", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldPIZip).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldIDNo) = New SqlParameter("@IDNo", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldIDNo).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldPIAddress) = New SqlParameter("@PIAddress", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldPIAddress).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldPICity) = New SqlParameter("@PICity", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldPICity).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldPIState) = New SqlParameter("@PIState", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldPIState).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldPlaceOfBirth) = New SqlParameter("@PlaceOfBirth", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldPlaceOfBirth).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldRace) = New SqlParameter("@Race", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldRace).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldLanguage) = New SqlParameter("@Language", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldLanguage).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldDriverLic) = New SqlParameter("@DriverLic", SqlDbType.NVarChar, 255)
        arParameters(Me.BillingFields.fldDriverLic).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldSocialSecurity) = New SqlParameter("@SocialSecurity", SqlDbType.NVarChar, 50)
        arParameters(Me.BillingFields.fldSocialSecurity).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldReligion) = New SqlParameter("@Religion", SqlDbType.NVarChar, 255)
        arParameters(Me.BillingFields.fldReligion).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldCellPhone) = New SqlParameter("@CellPhone", SqlDbType.NVarChar, 255)
        arParameters(Me.BillingFields.fldCellPhone).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldMaritalStatus) = New SqlParameter("@MaritalStatus", SqlDbType.NVarChar, 255)
        arParameters(Me.BillingFields.fldMaritalStatus).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldAKA) = New SqlParameter("@AKA", SqlDbType.NVarChar, 255)
        arParameters(Me.BillingFields.fldAKA).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldSpecialNeeds) = New SqlParameter("@SpecialNeeds", SqlDbType.NVarChar, 255)
        arParameters(Me.BillingFields.fldSpecialNeeds).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldEmail) = New SqlParameter("@Email", SqlDbType.NVarChar, 255)
        arParameters(Me.BillingFields.fldEmail).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldPager) = New SqlParameter("@Pager", SqlDbType.NVarChar, 255)
        arParameters(Me.BillingFields.fldPager).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldEExt) = New SqlParameter("@EExt", SqlDbType.NVarChar, 255)
        arParameters(Me.BillingFields.fldEExt).Direction = ParameterDirection.Output
        arParameters(Me.BillingFields.fldDOB) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        arParameters(Me.BillingFields.fldDOB).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spBillingGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spBillingGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.BillingFields.fldID).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            ID = ProcessNull.GetInt32(arParameters(Me.BillingFields.fldID).Value)
            Address1 = ProcessNull.GetString(arParameters(Me.BillingFields.fldAddress1).Value)
            Phone = ProcessNull.GetString(arParameters(Me.BillingFields.fldPHone).Value)
            City = ProcessNull.GetString(arParameters(Me.BillingFields.fldCity).Value)
            State = ProcessNull.GetString(arParameters(Me.BillingFields.fldState).Value)
            Zip = ProcessNull.GetString(arParameters(Me.BillingFields.fldZip).Value)
            EmPerson = ProcessNull.GetString(arParameters(Me.BillingFields.fldEmPerson).Value)
            EmPHone = ProcessNull.GetString(arParameters(Me.BillingFields.fldEmPHone).Value)
            EmployedAs = ProcessNull.GetString(arParameters(Me.BillingFields.fldEmployedAs).Value)
            Employer = ProcessNull.GetString(arParameters(Me.BillingFields.fldEmployer).Value)
            EAddress = ProcessNull.GetString(arParameters(Me.BillingFields.fldEAddress).Value)
            EPhone = ProcessNull.GetString(arParameters(Me.BillingFields.fldEPhone).Value)
            ECity = ProcessNull.GetString(arParameters(Me.BillingFields.fldECity).Value)
            EState = ProcessNull.GetString(arParameters(Me.BillingFields.fldEState).Value)
            EZip = ProcessNull.GetString(arParameters(Me.BillingFields.fldEZip).Value)
            IDNo = ProcessNull.GetString(arParameters(Me.BillingFields.fldIDNo).Value)
            PrimaryInsurance = ProcessNull.GetString(arParameters(Me.BillingFields.fldPrimaryInsurance).Value)
            PIAddress = ProcessNull.GetString(arParameters(Me.BillingFields.fldPIAddress).Value)
            PICity = ProcessNull.GetString(arParameters(Me.BillingFields.fldPICity).Value)
            PIState = ProcessNull.GetString(arParameters(Me.BillingFields.fldPIState).Value)
            PIZip = ProcessNull.GetString(arParameters(Me.BillingFields.fldPIZip).Value)
            PlaceOfBirth = ProcessNull.GetString(arParameters(Me.BillingFields.fldPlaceOfBirth).Value)
            Race = ProcessNull.GetString(arParameters(Me.BillingFields.fldRace).Value)
            Language = ProcessNull.GetString(arParameters(Me.BillingFields.fldLanguage).Value)
            DriverLic = ProcessNull.GetString(arParameters(Me.BillingFields.fldDriverLic).Value)
            SocialSecurity = ProcessNull.GetString(arParameters(Me.BillingFields.fldSocialSecurity).Value)
            Religion = ProcessNull.GetString(arParameters(Me.BillingFields.fldReligion).Value)
            CellPhone = ProcessNull.GetString(arParameters(Me.BillingFields.fldCellPhone).Value)
            MaritalStatus = ProcessNull.GetString(arParameters(Me.BillingFields.fldMaritalStatus).Value)
            AKA = ProcessNull.GetString(arParameters(Me.BillingFields.fldAKA).Value)
            SpecialNeeds = ProcessNull.GetString(arParameters(Me.BillingFields.fldSpecialNeeds).Value)
            Email = ProcessNull.GetString(arParameters(Me.BillingFields.fldEmail).Value)
            Pager = ProcessNull.GetString(arParameters(Me.BillingFields.fldPager).Value)
            EExt = ProcessNull.GetString(arParameters(Me.BillingFields.fldEExt).Value)
            DOB = ProcessNull.GetString(arParameters(Me.BillingFields.fldDOB).Value)
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
                        ByVal ChartID As Integer, _
                        ByVal Address1 As String, _
                        ByVal Phone As String) As Boolean

        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@BillingID", SqlDbType.Int)
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(1).Value = ChartID
        arParameters(2) = New SqlParameter("@Address1", SqlDbType.NVarChar, 50)
        arParameters(2).Value = Address1
        arParameters(3) = New SqlParameter("@Phone", SqlDbType.NVarChar, 50)
        arParameters(3).Value = Phone

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spBillingInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spBillingInsert", arParameters)
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
                ByVal Address1 As String, _
                ByVal Phone As String, _
                ByVal City As String, _
                ByVal State As String, _
                ByVal Zip As String, _
                ByVal EmPerson As String, _
                ByVal EmPHone As String, _
                ByVal EmployedAs As String, _
                ByVal Employer As String, _
                ByVal EAddress As String, _
                ByVal ECity As String, _
                ByVal EState As String, _
                ByVal EZip As String, _
                ByVal EPhone As String, _
                ByVal PrimaryInsurance As String, _
                ByVal PIZip As String, _
                ByVal IDNo As String, _
                ByVal PIAddress As String, _
                ByVal PICity As String, _
                ByVal PIState As String, _
                ByVal PlaceOfBirth As String, _
                ByVal Race As String, _
                ByVal Language As String, _
                ByVal DriverLic As String, _
                ByVal SocialSecurity As String, _
                ByVal Religion As String, _
                ByVal CellPhone As String, _
                ByVal MaritalStatus As String, _
                ByVal AKA As String, _
                ByVal SpecialNeeds As String, _
                ByVal Email As String, _
                ByVal Pager As String, _
                ByVal EExt As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(33) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@BillingID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Address1", SqlDbType.NVarChar, 50)
        arParameters(1).Value = Address1
        arParameters(2) = New SqlParameter("@PHone", SqlDbType.NVarChar, 50)
        arParameters(2).Value = Phone
        arParameters(3) = New SqlParameter("@City", SqlDbType.NVarChar, 50)
        arParameters(3).Value = City
        arParameters(4) = New SqlParameter("@State", SqlDbType.NVarChar, 50)
        arParameters(4).Value = State
        arParameters(5) = New SqlParameter("@Zip", SqlDbType.NVarChar, 50)
        arParameters(5).Value = Zip
        arParameters(6) = New SqlParameter("@EMPerson", SqlDbType.NVarChar, 50)
        arParameters(6).Value = EmPerson
        arParameters(7) = New SqlParameter("@EMPhone", SqlDbType.NVarChar, 50)
        arParameters(7).Value = EmPHone
        arParameters(8) = New SqlParameter("@EmployedAs", SqlDbType.NVarChar, 50)
        arParameters(8).Value = EmployedAs
        arParameters(9) = New SqlParameter("@Employer", SqlDbType.NVarChar, 50)
        arParameters(9).Value = Employer
        arParameters(10) = New SqlParameter("@EAddress", SqlDbType.NVarChar, 50)
        arParameters(10).Value = EAddress
        arParameters(11) = New SqlParameter("@ECity", SqlDbType.NVarChar, 50)
        arParameters(11).Value = ECity
        arParameters(12) = New SqlParameter("@Estate", SqlDbType.NVarChar, 50)
        arParameters(12).Value = EState
        arParameters(13) = New SqlParameter("@EZip", SqlDbType.NVarChar, 50)
        arParameters(13).Value = EZip
        arParameters(14) = New SqlParameter("@EPHone", SqlDbType.NVarChar, 50)
        arParameters(14).Value = EPhone
        arParameters(15) = New SqlParameter("@PrimaryInsurance", SqlDbType.NVarChar, 50)
        arParameters(15).Value = PrimaryInsurance
        arParameters(16) = New SqlParameter("@PIZip", SqlDbType.NVarChar, 50)
        arParameters(16).Value = PIZip
        arParameters(17) = New SqlParameter("@IDNo", SqlDbType.NVarChar, 50)
        arParameters(17).Value = IDNo
        arParameters(18) = New SqlParameter("@PIAddress", SqlDbType.NVarChar, 50)
        arParameters(18).Value = PIAddress
        arParameters(19) = New SqlParameter("@PICity", SqlDbType.NVarChar, 50)
        arParameters(19).Value = PICity
        arParameters(20) = New SqlParameter("@PIState", SqlDbType.NVarChar, 50)
        arParameters(20).Value = PIState
        arParameters(21) = New SqlParameter("@PlaceOfBirth", SqlDbType.NVarChar, 255)
        arParameters(21).Value = PlaceOfBirth
        arParameters(22) = New SqlParameter("@Race", SqlDbType.NVarChar, 50)
        arParameters(22).Value = Race
        arParameters(23) = New SqlParameter("@Language", SqlDbType.NVarChar, 50)
        arParameters(23).Value = Language
        arParameters(24) = New SqlParameter("@DriverLic", SqlDbType.NVarChar, 255)
        arParameters(24).Value = DriverLic
        arParameters(25) = New SqlParameter("@SocialSecurity", SqlDbType.NVarChar, 50)
        arParameters(25).Value = SocialSecurity
        arParameters(26) = New SqlParameter("@Religion", SqlDbType.NVarChar, 255)
        arParameters(26).Value = Religion
        arParameters(27) = New SqlParameter("@CellPhone", SqlDbType.NVarChar, 255)
        arParameters(27).Value = CellPhone
        arParameters(28) = New SqlParameter("@MaritalStatus", SqlDbType.NVarChar, 255)
        arParameters(28).Value = MaritalStatus
        arParameters(29) = New SqlParameter("@AKA", SqlDbType.NVarChar, 255)
        arParameters(29).Value = AKA
        arParameters(30) = New SqlParameter("@SpecialNeeds", SqlDbType.NVarChar, 255)
        arParameters(30).Value = SpecialNeeds
        arParameters(31) = New SqlParameter("@Email", SqlDbType.NVarChar, 255)
        arParameters(31).Value = Email
        arParameters(32) = New SqlParameter("@Pager", SqlDbType.NVarChar, 255)
        arParameters(32).Value = Pager
        arParameters(33) = New SqlParameter("@EExt", SqlDbType.NVarChar, 255)
        arParameters(33).Value = EExt

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spBillingUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spBillingUpdate", arParameters)
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


End Class 'dalBilling
