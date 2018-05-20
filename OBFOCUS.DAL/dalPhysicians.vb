
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalPhysicians
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
Public Class dalPhysicians

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum PhysicianFields
        fldID = 0
        fldSiteID = 1
        fldSalute = 2
        fldLastName = 3
        fldFirstName = 4
        fldTitle = 5
        fldAddress = 6
        fldCity = 7
        fldState = 8
        fldCountry = 9
        fldPostalCode = 10
        fldPhoneNumber = 11
        fldFaxNumber = 12
        fldSuppress = 13
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
    '* Name:        GetPhysicians
    '*
    '* Description: Returns all records in the [Physicians] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetPhysicians(ByVal SiteID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@SiteID", SqlDbType.Int)
        If SiteID = 0 Then
            arParameters(0).Value = DBNull.Value
        Else
            arParameters(0).Value = SiteID
        End If

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPhysicianGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPhysicianGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetAll
    '*
    '* Description: Returns all records in the [Physicians] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAll(ByVal SiteID As Integer, ByVal PLastName As String, ByVal PhysicianID As Integer) As SqlDataReader
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@SiteID", SqlDbType.Int)
        If Len(SiteID) = 0 Or SiteID = 0 Then
            arParameters(0).Value = DBNull.Value
        Else
            arParameters(0).Value = SiteID
        End If
        arParameters(1) = New SqlParameter("@PLastName", SqlDbType.VarChar, 50)
        If Len(PLastName) = 0 Then
            arParameters(1).Value = DBNull.Value
        Else
            arParameters(1).Value = PLastName
        End If
        arParameters(2) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        If Len(PhysicianID) = 0 Or PhysicianID = 0 Then
            arParameters(2).Value = DBNull.Value
        Else
            arParameters(2).Value = PhysicianID
        End If
      
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPhysicianGetAll", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPhysicianGetAll", arParameters)
            End If
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
    Public Function GetByKey(ByVal PhysicianID As Integer, ByRef SiteID As Integer, ByRef Salute As String, _
            ByRef LastName As String, ByRef FirstName As String, ByRef Title As String, _
            ByRef Address As String, ByRef city As String, ByRef State As String, _
            ByRef Country As String, ByRef PostalCode As String, ByRef PhoneNumber As String, _
            ByRef FaxNumber As String, ByRef Suppress As Short) As Boolean

        Dim TestNull As Object
        Dim arParameters(13) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.PhysicianFields.fldID) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(Me.PhysicianFields.fldID).Value = PhysicianID
        arParameters(Me.PhysicianFields.fldSiteID) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(Me.PhysicianFields.fldSiteID).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldSalute) = New SqlParameter("@Salute", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldSalute).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldLastName) = New SqlParameter("@PlastName", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldLastName).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldFirstName) = New SqlParameter("@PFirstName", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldFirstName).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldTitle) = New SqlParameter("@Title", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldTitle).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldAddress) = New SqlParameter("@PAddress", SqlDbType.VarChar, 255)
        arParameters(Me.PhysicianFields.fldAddress).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldCity) = New SqlParameter("@PCity", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldCity).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldState) = New SqlParameter("@PState", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldState).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldCountry) = New SqlParameter("@PCountry", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldCountry).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldPostalCode) = New SqlParameter("@PPostalCode", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldPostalCode).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldPhoneNumber) = New SqlParameter("@PPhoneNumber", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldPhoneNumber).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldFaxNumber) = New SqlParameter("@PFaxNumber", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldFaxNumber).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldSuppress) = New SqlParameter("@Suppress", SqlDbType.Bit)
        arParameters(Me.PhysicianFields.fldSuppress).Direction = ParameterDirection.Output


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPhysicianGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPhysicianGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.PhysicianFields.fldLastName).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            SiteID = ProcessNull.GetInt32(arParameters(Me.PhysicianFields.fldSiteID).Value)
            Salute = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldSalute).Value)
            Salute = Salute.Trim()
            LastName = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldLastName).Value)
            LastName = LastName.Trim()
            FirstName = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldFirstName).Value)
            FirstName = FirstName.Trim()
            Title = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldTitle).Value)
            Title = Title.Trim()
            Address = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldAddress).Value)
            Address = Address.Trim()
            city = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldCity).Value)
            city = city.Trim()
            State = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldState).Value)
            State = State.Trim()
            Country = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldCountry).Value)
            Country = Country.Trim()
            PostalCode = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldPostalCode).Value)
            PostalCode = PostalCode.Trim()
            PhoneNumber = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldPhoneNumber).Value)
            PhoneNumber = PhoneNumber.Trim()
            FaxNumber = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldFaxNumber).Value)
            FaxNumber = FaxNumber.Trim()
            Suppress = ProcessNull.GetInt16(arParameters(Me.PhysicianFields.fldSuppress).Value)
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
    Public Function Update(ByVal PhysicianID As Integer, ByVal SiteID As Integer, ByVal Salute As String, _
            ByVal LastName As String, ByVal FirstName As String, ByVal Title As String, _
            ByVal Address As String, ByVal city As String, ByVal State As String, _
            ByVal Country As String, ByVal PostalCode As String, ByVal PhoneNumber As String, _
            ByVal FaxNumber As String, ByVal Suppress As Short) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(13) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.PhysicianFields.fldID) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(Me.PhysicianFields.fldID).Value = PhysicianID
        arParameters(Me.PhysicianFields.fldSiteID) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(Me.PhysicianFields.fldSiteID).Value = SiteID
        arParameters(Me.PhysicianFields.fldSalute) = New SqlParameter("@Salute", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldSalute).Value = Salute
        arParameters(Me.PhysicianFields.fldLastName) = New SqlParameter("@PlastName", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldLastName).Value = LastName
        arParameters(Me.PhysicianFields.fldFirstName) = New SqlParameter("@PFirstName", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldFirstName).Value = FirstName
        arParameters(Me.PhysicianFields.fldTitle) = New SqlParameter("@Title", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldTitle).Value = Title
        arParameters(Me.PhysicianFields.fldAddress) = New SqlParameter("@PAddress", SqlDbType.VarChar, 255)
        arParameters(Me.PhysicianFields.fldAddress).Value = Address
        arParameters(Me.PhysicianFields.fldCity) = New SqlParameter("@PCity", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldCity).Value = city
        arParameters(Me.PhysicianFields.fldState) = New SqlParameter("@PState", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldState).Value = State
        arParameters(Me.PhysicianFields.fldCountry) = New SqlParameter("@PCountry", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldCountry).Value = Country
        arParameters(Me.PhysicianFields.fldPostalCode) = New SqlParameter("@PPostalCode", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldPostalCode).Value = PostalCode
        arParameters(Me.PhysicianFields.fldPhoneNumber) = New SqlParameter("@PPhoneNumber", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldPhoneNumber).Value = PhoneNumber
        arParameters(Me.PhysicianFields.fldFaxNumber) = New SqlParameter("@PFaxNumber", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldFaxNumber).Value = FaxNumber
        arParameters(Me.PhysicianFields.fldSuppress) = New SqlParameter("@Suppress", SqlDbType.Bit)
        arParameters(Me.PhysicianFields.fldSuppress).Value = Suppress


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPhysicianUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPhysicianUpdate", arParameters)
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
    '* Parameters:  ID - Returns AutoNumber Column
    '*              ListedDate - Input parameter
    '*              APN - Input parameter
    '*              TG - Input parameter
    '*              StreetNumber1 - Input parameter
    '*              StreetAddressDir1 - Input parameter
    '*              StreetAddress1 - Input parameter
    '*              StreetCity1 - Input parameter
    '*              StreetState1 - Input parameter
    '*              StreetZip1 - Input parameter
    '*              StreetNumber2 - Input parameter
    '*              StreetAddressDir2 - Input parameter
    '*              StreetAddress2 - Input parameter
    '*              StreetCity2 - Input parameter
    '*              StreetState2 - Input parameter
    '*              StreetZip2 - Input parameter
    '*              Use - Input parameter
    '*              SqFt - Input parameter
    '*              TaxValue - Input parameter
    '*              YearBuilt - Input parameter
    '*              Lot - Input parameter
    '*              TaxYear - Input parameter
    '*              Rooms - Input parameter
    '*              Zoning - Input parameter
    '*              PurchaseDate - Input parameter
    '*              TrustDeedAmt1 - Input parameter
    '*              TrustDeedDate1 - Input parameter
    '*              TrustDeedSpec1 - Input parameter
    '*              TrustDeedAmt2 - Input parameter
    '*              TrustDeedDate2 - Input parameter
    '*              TrustDeedSpec2 - Input parameter
    '*              TrustDeedAmt3 - Input parameter
    '*              TrustDeedDate3 - Input parameter
    '*              TrustDeedSpec3 - Input parameter
    '*              TrustDeedAmt4 - Input parameter
    '*              TrustDeedDate4 - Input parameter
    '*              TrustDeedSpec4 - Input parameter
    '*              Trustor - Input parameter
    '*              Owner - Input parameter
    '*              Trustee - Input parameter
    '*              TrusteePhone - Input parameter
    '*              TrusteeSaleNumber - Input parameter
    '*              Beneficiary - Input parameter
    '*              BeneficiaryPhone - Input parameter
    '*              SaleDate - Input parameter
    '*              SaleTime - Input parameter
    '*              MinimumBid - Input parameter
    '*              Site - Input parameter
    '*              Notes - Input parameter
    '*              LoanNumber - Input parameter
    '*              NOD - Input parameter
    '*              NTS - Input parameter
    '*              TDID - Input parameter
    '*              TitlePerson - Input parameter
    '*              AppraisedDate - Input parameter
    '*              County - Input parameter
    '*              Occupancy - Input parameter
    '*              Taxes - Input parameter
    '*              Legal - Input parameter
    '*              Appraiser - Input parameter
    '*
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef PhysicianID As Integer, ByVal SiteID As Integer, ByVal Salute As String, _
            ByVal LastName As String, ByVal FirstName As String, ByVal Title As String, _
            ByVal Address As String, ByVal city As String, ByVal State As String, _
            ByVal Country As String, ByVal PostalCode As String, ByVal PhoneNumber As String, _
            ByVal FaxNumber As String, ByVal Suppress As Short) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(13) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.PhysicianFields.fldID) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(Me.PhysicianFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.PhysicianFields.fldSiteID) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(Me.PhysicianFields.fldSiteID).Value = SiteID
        arParameters(Me.PhysicianFields.fldSalute) = New SqlParameter("@Salute", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldSalute).Value = Salute
        arParameters(Me.PhysicianFields.fldLastName) = New SqlParameter("@PlastName", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldLastName).Value = LastName
        arParameters(Me.PhysicianFields.fldFirstName) = New SqlParameter("@PFirstName", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldFirstName).Value = FirstName
        arParameters(Me.PhysicianFields.fldTitle) = New SqlParameter("@Title", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldTitle).Value = Title
        arParameters(Me.PhysicianFields.fldAddress) = New SqlParameter("@PAddress", SqlDbType.VarChar, 255)
        arParameters(Me.PhysicianFields.fldAddress).Value = Address
        arParameters(Me.PhysicianFields.fldCity) = New SqlParameter("@PCity", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldCity).Value = city
        arParameters(Me.PhysicianFields.fldState) = New SqlParameter("@PState", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldState).Value = State
        arParameters(Me.PhysicianFields.fldCountry) = New SqlParameter("@PCountry", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldCountry).Value = Country
        arParameters(Me.PhysicianFields.fldPostalCode) = New SqlParameter("@PPostalCode", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldPostalCode).Value = PostalCode
        arParameters(Me.PhysicianFields.fldPhoneNumber) = New SqlParameter("@PPhoneNumber", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldPhoneNumber).Value = PhoneNumber
        arParameters(Me.PhysicianFields.fldFaxNumber) = New SqlParameter("@PFaxNumber", SqlDbType.VarChar, 50)
        arParameters(Me.PhysicianFields.fldFaxNumber).Value = FaxNumber
        arParameters(Me.PhysicianFields.fldSuppress) = New SqlParameter("@Suppress", SqlDbType.Bit)
        arParameters(Me.PhysicianFields.fldSuppress).Value = Suppress


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPhysicianInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPhysicianInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            PhysicianID = CType(arParameters(0).Value, Integer)
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
        arParameters(0) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPhysicianDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPhysicianDelete", arParameters)
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
    '* Name:        Copy
    '*
    '* Description: Copys a record from the [PatientInfo] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to Copy
    '*
    '* Returns:     Boolean indicating if record was Copyd or not. 
    '*              True (record found and Copyd); False (otherwise).
    '*
    '**************************************************************************
    Public Function Copy(ByVal ID As Integer) As Boolean

        Dim strSQL As String = ""
        Dim intRecordsAffected As Integer = 0

        '' Build SQL string
        'strSQL = strSQL & "INSERT INTO [Archive] "
        'strSQL = strSQL & "  ( [ListedDate] "
        'strSQL = strSQL & "  , [APN] "
        'strSQL = strSQL & "  , [TG] "
        'strSQL = strSQL & "  , [StreetNumber1] "
        'strSQL = strSQL & "  , [StreetAddressDir1] "
        'strSQL = strSQL & "  , [StreetAddress1] "
        'strSQL = strSQL & "  , [StreetCity1] "
        'strSQL = strSQL & "  , [StreetState1] "
        'strSQL = strSQL & "  , [StreetZip1] "
        'strSQL = strSQL & "  , [StreetNumber2] "
        'strSQL = strSQL & "  , [StreetAddressDir2] "
        'strSQL = strSQL & "  , [StreetAddress2] "
        'strSQL = strSQL & "  , [StreetCity2] "
        'strSQL = strSQL & "  , [StreetState2] "
        'strSQL = strSQL & "  , [StreetZip2] "
        'strSQL = strSQL & "  , [Use] "
        'strSQL = strSQL & "  , [SqFt] "
        'strSQL = strSQL & "  , [TaxValue] "
        'strSQL = strSQL & "  , [YearBuilt] "
        'strSQL = strSQL & "  , [Lot] "
        'strSQL = strSQL & "  , [TaxYear] "
        'strSQL = strSQL & "  , [Rooms] "
        'strSQL = strSQL & "  , [Zoning] "
        'strSQL = strSQL & "  , [PurchaseDate] "
        'strSQL = strSQL & "  , [TrustDeedAmt1] "
        'strSQL = strSQL & "  , [TrustDeedDate1] "
        'strSQL = strSQL & "  , [TrustDeedSpec1] "
        'strSQL = strSQL & "  , [TrustDeedAmt2] "
        'strSQL = strSQL & "  , [TrustDeedDate2] "
        'strSQL = strSQL & "  , [TrustDeedSpec2] "
        'strSQL = strSQL & "  , [TrustDeedAmt3] "
        'strSQL = strSQL & "  , [TrustDeedDate3] "
        'strSQL = strSQL & "  , [TrustDeedSpec3] "
        'strSQL = strSQL & "  , [TrustDeedAmt4] "
        'strSQL = strSQL & "  , [TrustDeedDate4] "
        'strSQL = strSQL & "  , [TrustDeedSpec4] "
        'strSQL = strSQL & "  , [Trustor] "
        'strSQL = strSQL & "  , [Owner] "
        'strSQL = strSQL & "  , [Trustee] "
        'strSQL = strSQL & "  , [TrusteePhone] "
        'strSQL = strSQL & "  , [TrusteeSaleNumber] "
        'strSQL = strSQL & "  , [Beneficiary] "
        'strSQL = strSQL & "  , [BeneficiaryPhone] "
        'strSQL = strSQL & "  , [SaleDate] "
        'strSQL = strSQL & "  , [SaleTime] "
        'strSQL = strSQL & "  , [MinimumBid] "
        'strSQL = strSQL & "  , [Site] "
        'strSQL = strSQL & "  , [Notes] "
        'strSQL = strSQL & "  , [LoanNumber] "
        'strSQL = strSQL & "  , [NOD] "
        'strSQL = strSQL & "  , [NTS] "
        'strSQL = strSQL & "  , [TDID] "
        'strSQL = strSQL & "  , [TitlePerson] "
        'strSQL = strSQL & "  , [AppraisedDate] "
        'strSQL = strSQL & "  , [County] "
        'strSQL = strSQL & "  , [Occupancy] "
        'strSQL = strSQL & "  , [Taxes] "
        'strSQL = strSQL & "  , [Legal] "
        'strSQL = strSQL & "  , [Appraiser] "
        'strSQL = strSQL & "  , [TrustDeedAmt5] "
        'strSQL = strSQL & "  , [TrustDeedDate5] "
        'strSQL = strSQL & "  , [TrustDeedSpec5] "
        'strSQL = strSQL & "  , [TrustDeedAmt6] "
        'strSQL = strSQL & "  , [TrustDeedDate6] "
        'strSQL = strSQL & "  , [TrustDeedSpec6] "
        'strSQL = strSQL & ") SELECT "
        'strSQL = strSQL & "  [ListedDate] "
        'strSQL = strSQL & "  , [APN] "
        'strSQL = strSQL & "  , [TG] "
        'strSQL = strSQL & "  , [StreetNumber1] "
        'strSQL = strSQL & "  , [StreetAddressDir1] "
        'strSQL = strSQL & "  , [StreetAddress1] "
        'strSQL = strSQL & "  , [StreetCity1] "
        'strSQL = strSQL & "  , [StreetState1] "
        'strSQL = strSQL & "  , [StreetZip1] "
        'strSQL = strSQL & "  , [StreetNumber2] "
        'strSQL = strSQL & "  , [StreetAddressDir2] "
        'strSQL = strSQL & "  , [StreetAddress2] "
        'strSQL = strSQL & "  , [StreetCity2] "
        'strSQL = strSQL & "  , [StreetState2] "
        'strSQL = strSQL & "  , [StreetZip2] "
        'strSQL = strSQL & "  , [Use] "
        'strSQL = strSQL & "  , [SqFt] "
        'strSQL = strSQL & "  , [TaxValue] "
        'strSQL = strSQL & "  , [YearBuilt] "
        'strSQL = strSQL & "  , [Lot] "
        'strSQL = strSQL & "  , [TaxYear] "
        'strSQL = strSQL & "  , [Rooms] "
        'strSQL = strSQL & "  , [Zoning] "
        'strSQL = strSQL & "  , [PurchaseDate] "
        'strSQL = strSQL & "  , [TrustDeedAmt1] "
        'strSQL = strSQL & "  , [TrustDeedDate1] "
        'strSQL = strSQL & "  , [TrustDeedSpec1] "
        'strSQL = strSQL & "  , [TrustDeedAmt2] "
        'strSQL = strSQL & "  , [TrustDeedDate2] "
        'strSQL = strSQL & "  , [TrustDeedSpec2] "
        'strSQL = strSQL & "  , [TrustDeedAmt3] "
        'strSQL = strSQL & "  , [TrustDeedDate3] "
        'strSQL = strSQL & "  , [TrustDeedSpec3] "
        'strSQL = strSQL & "  , [TrustDeedAmt4] "
        'strSQL = strSQL & "  , [TrustDeedDate4] "
        'strSQL = strSQL & "  , [TrustDeedSpec4] "
        'strSQL = strSQL & "  , [Trustor] "
        'strSQL = strSQL & "  , [Owner] "
        'strSQL = strSQL & "  , [Trustee] "
        'strSQL = strSQL & "  , [TrusteePhone] "
        'strSQL = strSQL & "  , [TrusteeSaleNumber] "
        'strSQL = strSQL & "  , [Beneficiary] "
        'strSQL = strSQL & "  , [BeneficiaryPhone] "
        'strSQL = strSQL & "  , [SaleDate] "
        'strSQL = strSQL & "  , [SaleTime] "
        'strSQL = strSQL & "  , [MinimumBid] "
        'strSQL = strSQL & "  , [Site] "
        'strSQL = strSQL & "  , [Notes] "
        'strSQL = strSQL & "  , [LoanNumber] "
        'strSQL = strSQL & "  , [NOD] "
        'strSQL = strSQL & "  , [NTS] "
        'strSQL = strSQL & "  , [TDID] "
        'strSQL = strSQL & "  , [TitlePerson] "
        'strSQL = strSQL & "  , [AppraisedDate] "
        'strSQL = strSQL & "  , [County] "
        'strSQL = strSQL & "  , [Occupancy] "
        'strSQL = strSQL & "  , [Taxes] "
        'strSQL = strSQL & "  , [Legal] "
        'strSQL = strSQL & "  , [Appraiser] "
        'strSQL = strSQL & "  , [TrustDeedAmt5] "
        'strSQL = strSQL & "  , [TrustDeedDate5] "
        'strSQL = strSQL & "  , [TrustDeedSpec5] "
        'strSQL = strSQL & "  , [TrustDeedAmt6] "
        'strSQL = strSQL & "  , [TrustDeedDate6] "
        'strSQL = strSQL & "  , [TrustDeedSpec6] "
        'strSQL = strSQL & " FROM [PatientInfo] "
        'strSQL = strSQL & " WHERE [ID] = " & SqlHelper.SQLString(ID)


        ' Execute the SQL
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.Text, strSQL)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.Text, strSQL)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try

        ' Return boolean indicating if record was Copied.
        Return (intRecordsAffected <> 0)

    End Function

    '**************************************************************************
    '*  
    '* Name:        CopytoMain
    '*
    '* Description: CopytoMains a record from the [PatientInfo] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to CopytoMain
    '*
    '* Returns:     Boolean indicating if record was CopytoMaind or not. 
    '*              True (record found and CopytoMaind); False (otherwise).
    '*
    '**************************************************************************
    Public Function CopytoMain(ByVal ID As Integer) As Boolean

        Dim strSQL As String = ""
        Dim intRecordsAffected As Integer = 0

        '' Build SQL string
        'strSQL = strSQL & "INSERT INTO [PatientInfo] "
        'strSQL = strSQL & "  ( [ListedDate] "
        'strSQL = strSQL & "  , [APN] "
        'strSQL = strSQL & "  , [TG] "
        'strSQL = strSQL & "  , [StreetNumber1] "
        'strSQL = strSQL & "  , [StreetAddressDir1] "
        'strSQL = strSQL & "  , [StreetAddress1] "
        'strSQL = strSQL & "  , [StreetCity1] "
        'strSQL = strSQL & "  , [StreetState1] "
        'strSQL = strSQL & "  , [StreetZip1] "
        'strSQL = strSQL & "  , [StreetNumber2] "
        'strSQL = strSQL & "  , [StreetAddressDir2] "
        'strSQL = strSQL & "  , [StreetAddress2] "
        'strSQL = strSQL & "  , [StreetCity2] "
        'strSQL = strSQL & "  , [StreetState2] "
        'strSQL = strSQL & "  , [StreetZip2] "
        'strSQL = strSQL & "  , [Use] "
        'strSQL = strSQL & "  , [SqFt] "
        'strSQL = strSQL & "  , [TaxValue] "
        'strSQL = strSQL & "  , [YearBuilt] "
        'strSQL = strSQL & "  , [Lot] "
        'strSQL = strSQL & "  , [TaxYear] "
        'strSQL = strSQL & "  , [Rooms] "
        'strSQL = strSQL & "  , [Zoning] "
        'strSQL = strSQL & "  , [PurchaseDate] "
        'strSQL = strSQL & "  , [TrustDeedAmt1] "
        'strSQL = strSQL & "  , [TrustDeedDate1] "
        'strSQL = strSQL & "  , [TrustDeedSpec1] "
        'strSQL = strSQL & "  , [TrustDeedAmt2] "
        'strSQL = strSQL & "  , [TrustDeedDate2] "
        'strSQL = strSQL & "  , [TrustDeedSpec2] "
        'strSQL = strSQL & "  , [TrustDeedAmt3] "
        'strSQL = strSQL & "  , [TrustDeedDate3] "
        'strSQL = strSQL & "  , [TrustDeedSpec3] "
        'strSQL = strSQL & "  , [TrustDeedAmt4] "
        'strSQL = strSQL & "  , [TrustDeedDate4] "
        'strSQL = strSQL & "  , [TrustDeedSpec4] "
        'strSQL = strSQL & "  , [Trustor] "
        'strSQL = strSQL & "  , [Owner] "
        'strSQL = strSQL & "  , [Trustee] "
        'strSQL = strSQL & "  , [TrusteePhone] "
        'strSQL = strSQL & "  , [TrusteeSaleNumber] "
        'strSQL = strSQL & "  , [Beneficiary] "
        'strSQL = strSQL & "  , [BeneficiaryPhone] "
        'strSQL = strSQL & "  , [SaleDate] "
        'strSQL = strSQL & "  , [SaleTime] "
        'strSQL = strSQL & "  , [MinimumBid] "
        'strSQL = strSQL & "  , [Site] "
        'strSQL = strSQL & "  , [Notes] "
        'strSQL = strSQL & "  , [LoanNumber] "
        'strSQL = strSQL & "  , [NOD] "
        'strSQL = strSQL & "  , [NTS] "
        'strSQL = strSQL & "  , [TDID] "
        'strSQL = strSQL & "  , [TitlePerson] "
        'strSQL = strSQL & "  , [AppraisedDate] "
        'strSQL = strSQL & "  , [County] "
        'strSQL = strSQL & "  , [Occupancy] "
        'strSQL = strSQL & "  , [Taxes] "
        'strSQL = strSQL & "  , [Legal] "
        'strSQL = strSQL & "  , [Appraiser] "
        'strSQL = strSQL & "  , [TrustDeedAmt5] "
        'strSQL = strSQL & "  , [TrustDeedDate5] "
        'strSQL = strSQL & "  , [TrustDeedSpec5] "
        'strSQL = strSQL & "  , [TrustDeedAmt6] "
        'strSQL = strSQL & "  , [TrustDeedDate6] "
        'strSQL = strSQL & "  , [TrustDeedSpec6] "
        'strSQL = strSQL & ") SELECT "
        'strSQL = strSQL & "  [ListedDate] "
        'strSQL = strSQL & "  , [APN] "
        'strSQL = strSQL & "  , [TG] "
        'strSQL = strSQL & "  , [StreetNumber1] "
        'strSQL = strSQL & "  , [StreetAddressDir1] "
        'strSQL = strSQL & "  , [StreetAddress1] "
        'strSQL = strSQL & "  , [StreetCity1] "
        'strSQL = strSQL & "  , [StreetState1] "
        'strSQL = strSQL & "  , [StreetZip1] "
        'strSQL = strSQL & "  , [StreetNumber2] "
        'strSQL = strSQL & "  , [StreetAddressDir2] "
        'strSQL = strSQL & "  , [StreetAddress2] "
        'strSQL = strSQL & "  , [StreetCity2] "
        'strSQL = strSQL & "  , [StreetState2] "
        'strSQL = strSQL & "  , [StreetZip2] "
        'strSQL = strSQL & "  , [Use] "
        'strSQL = strSQL & "  , [SqFt] "
        'strSQL = strSQL & "  , [TaxValue] "
        'strSQL = strSQL & "  , [YearBuilt] "
        'strSQL = strSQL & "  , [Lot] "
        'strSQL = strSQL & "  , [TaxYear] "
        'strSQL = strSQL & "  , [Rooms] "
        'strSQL = strSQL & "  , [Zoning] "
        'strSQL = strSQL & "  , [PurchaseDate] "
        'strSQL = strSQL & "  , [TrustDeedAmt1] "
        'strSQL = strSQL & "  , [TrustDeedDate1] "
        'strSQL = strSQL & "  , [TrustDeedSpec1] "
        'strSQL = strSQL & "  , [TrustDeedAmt2] "
        'strSQL = strSQL & "  , [TrustDeedDate2] "
        'strSQL = strSQL & "  , [TrustDeedSpec2] "
        'strSQL = strSQL & "  , [TrustDeedAmt3] "
        'strSQL = strSQL & "  , [TrustDeedDate3] "
        'strSQL = strSQL & "  , [TrustDeedSpec3] "
        'strSQL = strSQL & "  , [TrustDeedAmt4] "
        'strSQL = strSQL & "  , [TrustDeedDate4] "
        'strSQL = strSQL & "  , [TrustDeedSpec4] "
        'strSQL = strSQL & "  , [Trustor] "
        'strSQL = strSQL & "  , [Owner] "
        'strSQL = strSQL & "  , [Trustee] "
        'strSQL = strSQL & "  , [TrusteePhone] "
        'strSQL = strSQL & "  , [TrusteeSaleNumber] "
        'strSQL = strSQL & "  , [Beneficiary] "
        'strSQL = strSQL & "  , [BeneficiaryPhone] "
        'strSQL = strSQL & "  , [SaleDate] "
        'strSQL = strSQL & "  , [SaleTime] "
        'strSQL = strSQL & "  , [MinimumBid] "
        'strSQL = strSQL & "  , [Site] "
        'strSQL = strSQL & "  , [Notes] "
        'strSQL = strSQL & "  , [LoanNumber] "
        'strSQL = strSQL & "  , [NOD] "
        'strSQL = strSQL & "  , [NTS] "
        'strSQL = strSQL & "  , [TDID] "
        'strSQL = strSQL & "  , [TitlePerson] "
        'strSQL = strSQL & "  , [AppraisedDate] "
        'strSQL = strSQL & "  , [County] "
        'strSQL = strSQL & "  , [Occupancy] "
        'strSQL = strSQL & "  , [Taxes] "
        'strSQL = strSQL & "  , [Legal] "
        'strSQL = strSQL & "  , [Appraiser] "
        'strSQL = strSQL & "  , [TrustDeedAmt5] "
        'strSQL = strSQL & "  , [TrustDeedDate5] "
        'strSQL = strSQL & "  , [TrustDeedSpec5] "
        'strSQL = strSQL & "  , [TrustDeedAmt6] "
        'strSQL = strSQL & "  , [TrustDeedDate6] "
        'strSQL = strSQL & "  , [TrustDeedSpec6] "
        'strSQL = strSQL & " FROM [PatientInfo] "
        'strSQL = strSQL & " WHERE [ID] = " & SqlHelper.SQLString(ID)


        ' Execute the SQL
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.Text, strSQL)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.Text, strSQL)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try

        ' Return boolean indicating if record was Copied.
        Return (intRecordsAffected <> 0)

    End Function
#End Region


End Class 'dalPhysicians
