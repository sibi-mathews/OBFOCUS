
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalAllergies
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
Public Class dalAllergies

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum WdiagnosisFields
        fldID = 0
        fldAllergy = 1
        fldAllergyTypeID = 2
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



#Region "Main procedures - GetLab, Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Name:        GetAllergies
    '*
    '* Description: Returns all records in the [Labs] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAllergies(ByVal ChartID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ChartID
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spAllergiesGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spAllergiesGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        Update
    '*
    '* Description: Updates a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal ID As Integer, _
                           ByVal AllergyTypeID As Integer, _
                           ByVal UpdatedBy As String) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@AllergiesID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@AllergyTypeID", SqlDbType.Int)
        If AllergyTypeID = Nothing Then
            arParameters(1).Value = DBNull.Value
        Else
            arParameters(1).Value = AllergyTypeID
        End If
        arParameters(2) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(2).Value = UpdatedBy
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAllergiesUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAllergiesUpdate", arParameters)
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
    Public Function Add(ByRef ID As Integer, _
                        ByVal AllergyTypeID As Integer, _
                        ByVal UserID As String) As Boolean

        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@AllergiesID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(1).Value = ID
        arParameters(2) = New SqlParameter("@AllergyTypeID", SqlDbType.Int)
        If AllergyTypeID = Nothing Then
            arParameters(2).Value = DBNull.Value
        Else
            arParameters(2).Value = AllergyTypeID
        End If
        arParameters(3) = New SqlParameter("@UserID", SqlDbType.NVarChar, 50)
        arParameters(3).Value = UserID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAllergiesInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAllergiesInsert", arParameters)
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
    Public Function Delete(ByVal AllergiesID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@AllergiesID", SqlDbType.Int)
        arParameters(0).Value = AllergiesID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAllergiesDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAllergiesDelete", arParameters)
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


End Class 'dalAllergies
