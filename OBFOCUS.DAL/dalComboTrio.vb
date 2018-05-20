
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalComboTrio
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
Public Class dalComboTrio

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum ComboDualFields
        fldDescription = 0
        fldOther = 1
        fldID = 2
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



#Region "Main procedures - Add, Update & Delete"
    '* ************************************************************
    '* Name:        GetProcedures
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetProcedures() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spProceduresGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spProceduresGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '* ************************************************************
    '* Name:        GetAllergyType
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAllergyType() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spAllergyTypeGetWClass")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spAllergyTypeGetWClass")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '* ************************************************************
    '* Name:        GetAlcohol
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAlcohol() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spAlcoholGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spAlcoholGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '* ************************************************************
    '* Name:        GetT21Risk
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetT21Risk() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spT21RiskGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spT21RiskGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '* ************************************************************
    '* Name:        GetFormLetter
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetFormLetter() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spFormLetterGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spFormLetterGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '* ************************************************************
    '* Name:        GetAnatomy
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAnatomy(ByVal ExaminerID As Integer) As SqlDataReader
        ' Call stored procedure and return the data
        Try
            Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
            ' Set the parameters
            arParameters(0) = New SqlParameter("@ExaminerID", SqlDbType.VarChar, 100)
            arParameters(0).Value = ExaminerID
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spAnatomyGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spAnatomyGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetAdnexa
    '*
    '* Description: Returns all records in the [Adnexa] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAdnexa(ByVal ExaminerID As Integer) As SqlDataReader
        ' Call stored procedure and return the data
        Try
            Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters
            ' Set the parameters
            arParameters(0) = New SqlParameter("@ExaminerID", SqlDbType.Int)
            If ExaminerID = Nothing Then
                arParameters(0).Value = DBNull.Value
            Else
                arParameters(0).Value = ExaminerID
            End If
            arParameters(1) = New SqlParameter("@ModeParam", SqlDbType.Bit)
            arParameters(1).Value = 0

            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spAdnexaGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spAdnexaGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function 'GetAdnexa
    '* ************************************************************
    '* Name:        GetGDisease
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetGDisease() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spGDiseaseGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spGDiseaseGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetBillingCode
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetBillingCode() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spBillingCodeGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spBillingCodeGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetAdnexaAll
    '*
    '* Description: Returns all records in the [Adnexa] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAdnexaAll() As SqlDataReader
        ' Call stored procedure and return the data
        Try
            Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters
            ' Set the parameters
            arParameters(0) = New SqlParameter("@ExaminerID", SqlDbType.Int)
            arParameters(0).Value = DBNull.Value
            arParameters(1) = New SqlParameter("@ModeParam", SqlDbType.Bit)
            arParameters(1).Value = 1

            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spAdnexaGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spAdnexaGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function 'GetAdnexaAll
    '**************************************************************************
    '*  
    '* Name:        GetEvaluation
    '*
    '* Description: Returns all records in the [Evaluation] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetEvaluation(ByVal ExaminerID As Integer) As SqlDataReader
        ' Call stored procedure and return the data
        Try
            Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
            ' Set the parameters
            arParameters(0) = New SqlParameter("@ExaminerID", SqlDbType.Int)
            If ExaminerID = Nothing Then
                arParameters(0).Value = DBNull.Value
            Else
                arParameters(0).Value = ExaminerID
            End If

            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spEvaluationGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spEvaluationGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function 'GetEvaluation
    '**************************************************************************
    '*  
    '* Name:        GetRecommendation
    '*
    '* Description: Returns all records in the [Recommendation] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetRecommendation(ByVal ExaminerID As Integer) As SqlDataReader
        ' Call stored procedure and return the data
        Try
            Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
            ' Set the parameters
            arParameters(0) = New SqlParameter("@ExaminerID", SqlDbType.Int)
            If ExaminerID = Nothing Then
                arParameters(0).Value = DBNull.Value
            Else
                arParameters(0).Value = ExaminerID
            End If

            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spRecommendationGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spRecommendationGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function 'GetRecommendation
    '**************************************************************************
    '*  
    '* Name:        GetFormulary
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetFormulary() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spFormularyGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spFormularyGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetUltrasonographer
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetUltrasonographer() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasonographerGetAll")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spUltrasonographerGetAll")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
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
    Public Function Update(ByVal ID As Integer, _
                           ByVal MedicalRecord As Integer, _
                           ByVal DOB As Date, _
                           ByVal Gravida As String, _
                           ByVal Para As String, _
                           ByVal SAB As String, _
                           ByVal TOP As String, _
                           ByVal Term As String, _
                           ByVal Living As String, _
                           ByVal Race As String, _
                           ByVal LMP As Date, _
                           ByVal EarlyUS As Date, _
                           ByVal UseEDCBy As String, _
                           ByVal EDC As Date, _
                           ByVal DateCreated As Date, _
                           ByVal PatientID As Integer) As Boolean

        Dim arParameters(15) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@MedicalRecord", SqlDbType.Int)
        If MedicalRecord = Nothing Then
            arParameters(1).Value = DBNull.Value
        Else
            arParameters(1).Value = MedicalRecord
        End If
        arParameters(2) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        If DOB = Nothing Then
            arParameters(2).Value = DBNull.Value
        Else
            arParameters(2).Value = DOB
        End If
        arParameters(3) = New SqlParameter("@Race", SqlDbType.NVarChar, 50)
        arParameters(3).Value = Race
        arParameters(4) = New SqlParameter("@Gravida", SqlDbType.Int)
        If Trim(Gravida) = "" Then
            arParameters(4).Value = DBNull.Value
        Else
            arParameters(4).Value = CType(Gravida, Integer)
        End If
        arParameters(5) = New SqlParameter("@Para", SqlDbType.Int)
        If Trim(Para) = "" Then
            arParameters(5).Value = DBNull.Value
        Else
            arParameters(5).Value = CType(Para, Integer)
        End If
        arParameters(6) = New SqlParameter("@SAB", SqlDbType.Int)
        If Trim(SAB) = "" Then
            arParameters(6).Value = DBNull.Value
        Else
            arParameters(6).Value = CType(SAB, Integer)
        End If
        arParameters(7) = New SqlParameter("@TOP", SqlDbType.Int)
        If Trim(TOP) = "" Then
            arParameters(7).Value = DBNull.Value
        Else
            arParameters(7).Value = CType(TOP, Integer)
        End If
        arParameters(8) = New SqlParameter("@Term", SqlDbType.Int)
        If Trim(Term) = "" Then
            arParameters(8).Value = DBNull.Value
        Else
            arParameters(8).Value = CType(Term, Integer)
        End If
        arParameters(9) = New SqlParameter("@Living", SqlDbType.Int)
        If Trim(Living) = "" Then
            arParameters(9).Value = DBNull.Value
        Else
            arParameters(9).Value = CType(Living, Integer)
        End If
        arParameters(10) = New SqlParameter("@LMP", SqlDbType.SmallDateTime)
        If LMP = Nothing Then
            arParameters(10).Value = DBNull.Value
        Else
            arParameters(10).Value = LMP
        End If
        arParameters(11) = New SqlParameter("@EarlyUS", SqlDbType.SmallDateTime)
        If EarlyUS = Nothing Then
            arParameters(11).Value = DBNull.Value
        Else
            arParameters(11).Value = EarlyUS
        End If
        arParameters(12) = New SqlParameter("@EDC", SqlDbType.SmallDateTime)
        If EDC = Nothing Then
            arParameters(12).Value = DBNull.Value
        Else
            arParameters(12).Value = EDC
        End If
        arParameters(13) = New SqlParameter("@UseEDCBy", SqlDbType.NVarChar, 50)
        arParameters(13).Value = UseEDCBy
        arParameters(14) = New SqlParameter("@DateCreated", SqlDbType.SmallDateTime)
        If DateCreated = Nothing Then
            arParameters(14).Value = DBNull.Value
        Else
            arParameters(14).Value = DateCreated
        End If
        arParameters(15) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(15).Value = PatientID


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartUpdate", arParameters)
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
    Public Function Add(ByRef ID As Integer, _
                        ByVal ListedDate As Date, _
                        ByVal APN As String, _
                        ByVal TG As String, _
                        ByVal StreetNumber1 As Short, _
                        ByVal StreetAddressDir1 As String, _
                        ByVal StreetAddress1 As String, _
                        ByVal StreetCity1 As String, _
                        ByVal StreetState1 As String, _
                        ByVal StreetZip1 As String, _
                        ByVal StreetNumber2 As Short, _
                        ByVal StreetAddressDir2 As String, _
                        ByVal StreetAddress2 As String, _
                        ByVal StreetCity2 As String, _
                        ByVal StreetState2 As String, _
                        ByVal StreetZip2 As String, _
                        ByVal Use As String, _
                        ByVal SqFt As Integer, _
                        ByVal TaxValue As Integer, _
                        ByVal YearBuilt As Integer, _
                        ByVal Lot As Short, _
                        ByVal TaxYear As Short, _
                        ByVal Rooms As String, _
                        ByVal Zoning As String, _
                        ByVal PurchaseDate As Date, _
                        ByVal TrustDeedAmt1 As Decimal, _
                        ByVal TrustDeedDate1 As Date, _
                        ByVal TrustDeedSpec1 As String, _
                        ByVal TrustDeedAmt2 As Decimal, _
                        ByVal TrustDeedDate2 As Date, _
                        ByVal TrustDeedSpec2 As String, _
                        ByVal TrustDeedAmt3 As Decimal, _
                        ByVal TrustDeedDate3 As Date, _
                        ByVal TrustDeedSpec3 As String, _
                        ByVal TrustDeedAmt4 As Decimal, _
                        ByVal TrustDeedDate4 As Date, _
                        ByVal TrustDeedSpec4 As String, _
                        ByVal Trustor As String, _
                        ByVal Owner As String, _
                        ByVal Trustee As String, _
                        ByVal TrusteePhone As String, _
                        ByVal TrusteeSaleNumber As String, _
                        ByVal Beneficiary As String, _
                        ByVal BeneficiaryPhone As String, _
                        ByVal SaleDate As Date, _
                        ByVal SaleTime As String, _
                        ByVal MinimumBid As Decimal, _
                        ByVal Site As String, _
                        ByVal Notes As String, _
                        ByVal LoanNumber As String, _
                        ByVal NOD As String, _
                        ByVal NTS As String, _
                        ByVal TDID As String, _
                        ByVal TitlePerson As String, _
                        ByVal AppraisedDate As Date, _
                        ByVal County As String, _
                        ByVal Occupancy As String, _
                        ByVal Taxes As String, _
                        ByVal Legal As String, _
                        ByVal Appraiser As String, _
                        ByVal TrustDeedAmt5 As Decimal, _
                        ByVal TrustDeedDate5 As Date, _
                        ByVal TrustDeedSpec5 As String, _
                        ByVal TrustDeedAmt6 As Decimal, _
                        ByVal TrustDeedDate6 As Date, _
                        ByVal TrustDeedSpec6 As String) As Boolean

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
        'strSQL = strSQL & ") VALUES ("
        'strSQL = strSQL & "   " & SqlHelper.SQLString(ListedDate)
        'If APN = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(APN)
        'End If

        'If TG = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TG)
        'End If

        'If StreetNumber1 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(StreetNumber1)
        'End If

        'If StreetAddressDir1 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(StreetAddressDir1)
        'End If

        'If StreetAddress1 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(StreetAddress1)
        'End If

        'If StreetCity1 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(StreetCity1)
        'End If

        'If StreetState1 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(StreetState1)
        'End If

        'If StreetZip1 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(StreetZip1)
        'End If

        'If StreetNumber2 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(StreetNumber2)
        'End If

        'If StreetAddressDir2 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(StreetAddressDir2)
        'End If

        'If StreetAddress2 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(StreetAddress2)
        'End If

        'If StreetCity2 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(StreetCity2)
        'End If

        'If StreetState2 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(StreetState2)
        'End If

        'If StreetZip2 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(StreetZip2)
        'End If

        'If Use = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Use)
        'End If

        'If SqFt = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(SqFt)
        'End If

        'If TaxValue = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TaxValue)
        'End If

        'If YearBuilt = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(YearBuilt)
        'End If

        'If Lot = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Lot)
        'End If

        'If TaxYear = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TaxYear)
        'End If

        'If Rooms = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Rooms)
        'End If

        'If Zoning = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Zoning)
        'End If

        'If PurchaseDate = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(PurchaseDate)
        'End If

        'If TrustDeedAmt1 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedAmt1)
        'End If

        'If TrustDeedDate1 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedDate1)
        'End If

        'If TrustDeedSpec1 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedSpec1)
        'End If

        'If TrustDeedAmt2 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedAmt2)
        'End If

        'If TrustDeedDate2 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedDate2)
        'End If

        'If TrustDeedSpec2 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedSpec2)
        'End If

        'If TrustDeedAmt3 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedAmt3)
        'End If

        'If TrustDeedDate3 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedDate3)
        'End If

        'If TrustDeedSpec3 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedSpec3)
        'End If

        'If TrustDeedAmt4 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedAmt4)
        'End If

        'If TrustDeedDate4 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedDate4)
        'End If

        'If TrustDeedSpec4 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedSpec4)
        'End If

        'If Trustor = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Trustor)
        'End If

        'If Owner = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Owner)
        'End If

        'If Trustee = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Trustee)
        'End If

        'If TrusteePhone = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrusteePhone)
        'End If

        'If TrusteeSaleNumber = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrusteeSaleNumber)
        'End If

        'If Beneficiary = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Beneficiary)
        'End If

        'If BeneficiaryPhone = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(BeneficiaryPhone)
        'End If

        'If SaleDate = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(SaleDate)
        'End If

        'If SaleTime = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(SaleTime)
        'End If

        'If MinimumBid = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(MinimumBid)
        'End If

        'If Site = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Site)
        'End If

        'If Notes = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Notes)
        'End If

        'If LoanNumber = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(LoanNumber)
        'End If

        'If NOD = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(NOD)
        'End If

        'If NTS = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(NTS)
        'End If

        'If TDID = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TDID)
        'End If

        'If TitlePerson = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TitlePerson)
        'End If

        'If AppraisedDate = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(AppraisedDate)
        'End If

        'If County = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(County)
        'End If

        'If Occupancy = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Occupancy)
        'End If

        'If Taxes = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Taxes)
        'End If

        'If Legal = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Legal)
        'End If

        'If Legal = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(Appraiser)
        'End If
        'If TrustDeedAmt5 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedAmt5)
        'End If

        'If TrustDeedDate5 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedDate5)
        'End If

        'If TrustDeedSpec5 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedSpec5)
        'End If
        'If TrustDeedAmt6 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedAmt6)
        'End If

        'If TrustDeedDate6 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedDate6)
        'End If

        'If TrustDeedSpec6 = Nothing Then
        '    strSQL = strSQL & "  ,Null"
        'Else
        '    strSQL = strSQL & "  ," & SqlHelper.SQLString(TrustDeedSpec6)
        'End If

        'strSQL = strSQL & ")"

        Try
            ' Execute the SQL and get the identity value created.
            If Me.Transaction Is Nothing Then
                ''intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.Text, strSQL, ID)
            Else
                '  intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.Text, strSQL, ID)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return boolean indicating if record was added.
        Return (intRecordsAffected <> 0)

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

        Dim strSQL As String = ""
        Dim intRecordsAffected As Integer = 0

        ' Build SQL string 
        'strSQL = strSQL & "DELETE FROM [PatientInfo] "
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

        ' Return boolean indicating if record was deleted.
        Return (intRecordsAffected <> 0)

    End Function



#End Region


End Class 'dalComboTrio
