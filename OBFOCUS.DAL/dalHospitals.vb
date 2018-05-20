
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalHospitals
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
Public Class dalHospitals

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum HospitalFields
        fldID = 0
        fldNameHospital = 1
        fldHAddress = 2
        fldHCity = 3
        fldHState = 4
        fldHZip = 5
        fldHPhone = 6
        fldHFax = 7
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
    Public Function GetByKey(ByVal delHospitalID As Integer, _
                ByRef NameHospital As String, _
                ByRef HAddress As String, _
                ByRef HCity As String, _
                ByRef HState As String, _
                ByRef HZip As String, _
                ByRef HPhone As String, _
                ByRef HFax As String) As Boolean
        Dim arParameters(7) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.HospitalFields.fldID) = New SqlParameter("@delHospitalID", SqlDbType.Int)
        arParameters(Me.HospitalFields.fldID).Value = delHospitalID
        arParameters(Me.HospitalFields.fldNameHospital) = New SqlParameter("@NameHospital", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldNameHospital).Direction = ParameterDirection.Output
        arParameters(Me.HospitalFields.fldHAddress) = New SqlParameter("@HAddress", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHAddress).Direction = ParameterDirection.Output
        arParameters(Me.HospitalFields.fldHCity) = New SqlParameter("@HCity", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHCity).Direction = ParameterDirection.Output
        arParameters(Me.HospitalFields.fldHState) = New SqlParameter("@HState", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHState).Direction = ParameterDirection.Output
        arParameters(Me.HospitalFields.fldHZip) = New SqlParameter("@HZip", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHZip).Direction = ParameterDirection.Output
        arParameters(Me.HospitalFields.fldHPhone) = New SqlParameter("@HPhone", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHPhone).Direction = ParameterDirection.Output
        arParameters(Me.HospitalFields.fldHFax) = New SqlParameter("@HFax", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHFax).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spHospitalGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spHospitalGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.HospitalFields.fldNameHospital).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            NameHospital = ProcessNull.GetString(arParameters(Me.HospitalFields.fldNameHospital).Value)
            NameHospital = NameHospital.Trim()
            HAddress = ProcessNull.GetString(arParameters(Me.HospitalFields.fldHAddress).Value)
            HAddress = HAddress.Trim()
            HCity = ProcessNull.GetString(arParameters(Me.HospitalFields.fldHCity).Value)
            HCity = HCity.Trim()
            HState = ProcessNull.GetString(arParameters(Me.HospitalFields.fldHState).Value)
            HState = HState.Trim()
            HZip = ProcessNull.GetString(arParameters(Me.HospitalFields.fldHZip).Value)
            HZip = HZip.Trim()
            HPhone = ProcessNull.GetString(arParameters(Me.HospitalFields.fldHPhone).Value)
            HPhone = HPhone.Trim()
            HFax = ProcessNull.GetString(arParameters(Me.HospitalFields.fldHFax).Value)
            HFax = HFax.Trim()
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        GetByExamID
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
    Public Function GetByExamID(ByVal ExamID As Integer, _
                ByRef NameHospital As String, _
                ByRef SiteID As Integer, _
                ByRef HFax As String) As Boolean
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ExamID
        arParameters(1) = New SqlParameter("@NameHospital", SqlDbType.NVarChar, 50)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@HFax", SqlDbType.NVarChar, 50)
        arParameters(3).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spHospitalGetByExamID", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spHospitalGetByExamID", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.HospitalFields.fldNameHospital).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            NameHospital = ProcessNull.GetString(arParameters(1).Value)
            NameHospital = NameHospital.Trim()
            SiteID = ProcessNull.GetInt32(arParameters(2).Value)
            HFax = ProcessNull.GetString(arParameters(3).Value)
            HFax = HFax.Trim()
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
    Public Function Update(ByVal DelHospitalID As Integer, _
                ByVal NameHospital As String, _
                ByVal HAddress As String, _
                ByVal HCity As String, _
                ByVal HState As String, _
                ByVal HZip As String, _
                ByVal HPhone As String, _
                ByVal HFax As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(7) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.HospitalFields.fldID) = New SqlParameter("@delHospitalID", SqlDbType.Int)
        arParameters(Me.HospitalFields.fldID).Value = DelHospitalID
        arParameters(Me.HospitalFields.fldNameHospital) = New SqlParameter("@NameHospital", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldNameHospital).Value = NameHospital
        arParameters(Me.HospitalFields.fldHAddress) = New SqlParameter("@HAddress", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHAddress).Value = HAddress
        arParameters(Me.HospitalFields.fldHCity) = New SqlParameter("@HCity", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHCity).Value = HCity
        arParameters(Me.HospitalFields.fldHState) = New SqlParameter("@HState", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHState).Value = HState
        arParameters(Me.HospitalFields.fldHZip) = New SqlParameter("@HZip", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHZip).Value = HZip
        arParameters(Me.HospitalFields.fldHPhone) = New SqlParameter("@HPhone", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHPhone).Value = HPhone
        arParameters(Me.HospitalFields.fldHFax) = New SqlParameter("@HFax", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHFax).Value = HFax



        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spHospitalUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spHospitalUpdate", arParameters)
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
    Public Function Add(ByRef DelHospitalID As Integer, _
                ByVal NameHospital As String, _
                ByVal HAddress As String, _
                ByVal HCity As String, _
                ByVal HState As String, _
                ByVal HZip As String, _
                ByVal HPhone As String, _
                ByVal HFax As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(7) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.HospitalFields.fldID) = New SqlParameter("@DelHospitalID", SqlDbType.Int)
        arParameters(Me.HospitalFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.HospitalFields.fldNameHospital) = New SqlParameter("@NameHospital", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldNameHospital).Value = NameHospital
        arParameters(Me.HospitalFields.fldHAddress) = New SqlParameter("@HAddress", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHAddress).Value = HAddress
        arParameters(Me.HospitalFields.fldHCity) = New SqlParameter("@HCity", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHCity).Value = HCity
        arParameters(Me.HospitalFields.fldHState) = New SqlParameter("@HState", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHState).Value = HState
        arParameters(Me.HospitalFields.fldHZip) = New SqlParameter("@HZip", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHZip).Value = HZip
        arParameters(Me.HospitalFields.fldHPhone) = New SqlParameter("@HPhone", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHPhone).Value = HPhone
        arParameters(Me.HospitalFields.fldHFax) = New SqlParameter("@HFax", SqlDbType.NVarChar, 50)
        arParameters(Me.HospitalFields.fldHFax).Value = HFax



        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spHospitalInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spHospitalInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            DelHospitalID = CType(arParameters(0).Value, Integer)
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
        arParameters(Me.HospitalFields.fldID) = New SqlParameter("@delHospitalID", SqlDbType.Int)
        arParameters(Me.HospitalFields.fldID).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spHospitalDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spHospitalDelete", arParameters)
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


End Class 'dalHospitals
