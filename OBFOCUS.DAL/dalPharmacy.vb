
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalPharmacy
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
Public Class dalPharmacy

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum PharmacyFields
        fldID = 0
        fldPharmacyName = 1
        fldStreetAddress = 2
        fldCity = 3
        fldState = 4
        fldZip = 5
        fldBusinessPhone = 6
        fldFax = 7
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
    Public Function GetByKey(ByVal PharmacyID As Integer, _
                ByRef PharmacyName As String, _
                ByRef StreetAddress As String, _
                ByRef City As String, _
                ByRef State As String, _
                ByRef Zip As String, _
                ByRef BusinessPhone As String, _
                ByRef Fax As String) As Boolean
        Dim arParameters(7) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.PharmacyFields.fldID) = New SqlParameter("@PharmacyID", SqlDbType.Int)
        arParameters(Me.PharmacyFields.fldID).Value = PharmacyID
        arParameters(Me.PharmacyFields.fldPharmacyName) = New SqlParameter("@PharmacyName", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldPharmacyName).Direction = ParameterDirection.Output
        arParameters(Me.PharmacyFields.fldStreetAddress) = New SqlParameter("@StreetAddress", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldStreetAddress).Direction = ParameterDirection.Output
        arParameters(Me.PharmacyFields.fldCity) = New SqlParameter("@City", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldCity).Direction = ParameterDirection.Output
        arParameters(Me.PharmacyFields.fldState) = New SqlParameter("@State", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldState).Direction = ParameterDirection.Output
        arParameters(Me.PharmacyFields.fldZip) = New SqlParameter("@Zip", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldZip).Direction = ParameterDirection.Output
        arParameters(Me.PharmacyFields.fldBusinessPhone) = New SqlParameter("@BusinessPhone", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldBusinessPhone).Direction = ParameterDirection.Output
        arParameters(Me.PharmacyFields.fldFax) = New SqlParameter("@Fax", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldFax).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPharmacyGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPharmacyGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.PharmacyFields.fldPharmacyName).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            PharmacyName = ProcessNull.GetString(arParameters(Me.PharmacyFields.fldPharmacyName).Value)
            PharmacyName = PharmacyName.Trim()
            StreetAddress = ProcessNull.GetString(arParameters(Me.PharmacyFields.fldStreetAddress).Value)
            StreetAddress = StreetAddress.Trim()
            City = ProcessNull.GetString(arParameters(Me.PharmacyFields.fldCity).Value)
            City = City.Trim()
            State = ProcessNull.GetString(arParameters(Me.PharmacyFields.fldState).Value)
            State = State.Trim()
            Zip = ProcessNull.GetString(arParameters(Me.PharmacyFields.fldZip).Value)
            Zip = Zip.Trim()
            BusinessPhone = ProcessNull.GetString(arParameters(Me.PharmacyFields.fldBusinessPhone).Value)
            BusinessPhone = BusinessPhone.Trim()
            Fax = ProcessNull.GetString(arParameters(Me.PharmacyFields.fldFax).Value)
            Fax = Fax.Trim()
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
    Public Function Update(ByVal PharmacyID As Integer, _
                ByVal PharmacyName As String, _
                ByVal StreetAddress As String, _
                ByVal City As String, _
                ByVal State As String, _
                ByVal Zip As String, _
                ByVal BusinessPhone As String, _
                ByVal Fax As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(7) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.PharmacyFields.fldID) = New SqlParameter("@PharmacyID", SqlDbType.Int)
        arParameters(Me.PharmacyFields.fldID).Value = PharmacyID
        arParameters(Me.PharmacyFields.fldPharmacyName) = New SqlParameter("@PharmacyName", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldPharmacyName).Value = PharmacyName
        arParameters(Me.PharmacyFields.fldStreetAddress) = New SqlParameter("@StreetAddress", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldStreetAddress).Value = StreetAddress
        arParameters(Me.PharmacyFields.fldCity) = New SqlParameter("@City", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldCity).Value = City
        arParameters(Me.PharmacyFields.fldState) = New SqlParameter("@State", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldState).Value = State
        arParameters(Me.PharmacyFields.fldZip) = New SqlParameter("@Zip", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldZip).Value = Zip
        arParameters(Me.PharmacyFields.fldBusinessPhone) = New SqlParameter("@BusinessPhone", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldBusinessPhone).Value = BusinessPhone
        arParameters(Me.PharmacyFields.fldFax) = New SqlParameter("@Fax", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldFax).Value = Fax



        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPharmacyUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPharmacyUpdate", arParameters)
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
    Public Function Add(ByRef PharmacyID As Integer, _
                ByVal PharmacyName As String, _
                ByVal StreetAddress As String, _
                ByVal City As String, _
                ByVal State As String, _
                ByVal Zip As String, _
                ByVal BusinessPhone As String, _
                ByVal Fax As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(7) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.PharmacyFields.fldID) = New SqlParameter("@PharmacyID", SqlDbType.Int)
        arParameters(Me.PharmacyFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.PharmacyFields.fldPharmacyName) = New SqlParameter("@PharmacyName", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldPharmacyName).Value = PharmacyName
        arParameters(Me.PharmacyFields.fldStreetAddress) = New SqlParameter("@StreetAddress", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldStreetAddress).Value = StreetAddress
        arParameters(Me.PharmacyFields.fldCity) = New SqlParameter("@City", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldCity).Value = City
        arParameters(Me.PharmacyFields.fldState) = New SqlParameter("@State", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldState).Value = State
        arParameters(Me.PharmacyFields.fldZip) = New SqlParameter("@Zip", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldZip).Value = Zip
        arParameters(Me.PharmacyFields.fldBusinessPhone) = New SqlParameter("@BusinessPhone", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldBusinessPhone).Value = BusinessPhone
        arParameters(Me.PharmacyFields.fldFax) = New SqlParameter("@Fax", SqlDbType.NVarChar, 50)
        arParameters(Me.PharmacyFields.fldFax).Value = Fax



        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPharmacyInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPharmacyInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            PharmacyID = CType(arParameters(0).Value, Integer)
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
        arParameters(Me.PharmacyFields.fldID) = New SqlParameter("@PharmacyID", SqlDbType.Int)
        arParameters(Me.PharmacyFields.fldID).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPharmacyDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPharmacyDelete", arParameters)
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


End Class 'dalPharmacy
