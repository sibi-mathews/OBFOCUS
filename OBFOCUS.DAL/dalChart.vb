
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalChart
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
Public Class dalChart

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum ChartFields
        fldChartID = 0
        fldPatientID = 1
        fldMedicalRecord = 2
        fldLastName = 3
        fldFirst = 4
        fldDOB = 5
        fldRace = 6
        fldGravida = 7
        fldPara = 8
        fldSAB = 9
        fldTop = 10
        fldTerm = 11
        fldLiving = 12
        fldEDC = 13
        fldLMP = 14
        fldEarlyUs = 15
        fldUseEDCBy = 16
        fldRefDX = 17
        fldPhysicianID = 18
        fldDelHospitalID = 19
        fldSiteID = 20
        fldDateCreated = 21
        fldType = 22
        fldRH = 23
        fldAntiBody = 24
        fldPreWeight = 25
        fldHeight = 26
        fldPLastName = 27
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



#Region "Main procedures - GetAll, GetByKey, Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Name:        GetAll
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '* Parameters: bDefault,Index,bAsc
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAll(ByVal bDefault As Boolean, _
    ByVal Index As Integer, ByVal bAsc As Boolean, _
    ByVal Lastname As String, ByVal PatientID As Object, _
    ByVal DOB As Object) As SqlDataReader
        Dim arParameters(6) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@bDefault", SqlDbType.Bit)
        arParameters(0).Value = IIf(bDefault = True, 1, 0)
        arParameters(1) = New SqlParameter("@Index", SqlDbType.Int)
        arParameters(1).Value = Index
        arParameters(2) = New SqlParameter("@bAsc", SqlDbType.Bit)
        arParameters(2).Value = IIf(bAsc = True, 1, 0)
        arParameters(3) = New SqlParameter("@LastName", SqlDbType.NVarChar, 50)
        If Len(Lastname) = 0 Then
            arParameters(3).Value = DBNull.Value
        Else
            arParameters(3).Value = Lastname
        End If
        arParameters(4) = New SqlParameter("@PatientID", SqlDbType.Int)
        If Len(PatientID) = 0 Then
            arParameters(4).Value = DBNull.Value
        ElseIf CType(PatientID, Integer) <> 0 Then
            arParameters(4).Value = PatientID
        End If
        arParameters(5) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        If Len(DOB) = 0 Then
            arParameters(5).Value = DBNull.Value
        Else
            arParameters(5).Value = DOB
        End If
        arParameters(6) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        If Globals.LimPhysicianID = 0 Then
            arParameters(6).Value = DBNull.Value
        Else
            arParameters(6).Value = Globals.LimPhysicianID
        End If
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoGetAll", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoGetAll", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetChartAuditTrail
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '* Parameters: bDefault,Index,bAsc
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetChartAuditTrail(ByVal ChartID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ChartID
       

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spChart_AuditTrailGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spChart_AuditTrailGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetChart_LockAll
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '* Parameters: bDefault,Index,bAsc
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetChart_LockAll() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spChart_LockGetAll")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spChart_LockGetAll")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetReference
    '*
    '* Description: Returns all records in the [Reference] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetReference(ByVal head As String, _
                                ByVal Face As String, _
                                ByVal Thorax As String, _
                                ByVal Spine As String, _
                                ByVal Limbs As String, _
                                ByVal Hands As String, _
                                ByVal Feet As String, _
                                ByVal Cardiac As String, _
                                ByVal Kidney As String) As SqlDataReader
        ' Call stored procedure and return the data
        Try
            Dim arParameters(8) As SqlParameter         ' Array to hold stored procedure parameters
            ' Set the parameters
            arParameters(0) = New SqlParameter("@head", SqlDbType.NVarChar, 50)
            If head = "" Then
                arParameters(0).Value = DBNull.Value
            Else
                arParameters(0).Value = head
            End If
            arParameters(1) = New SqlParameter("@Face", SqlDbType.NVarChar, 51)
            If Face = "" Then
                arParameters(1).Value = DBNull.Value
            Else
                arParameters(1).Value = Face
            End If
            arParameters(2) = New SqlParameter("@Thorax", SqlDbType.NVarChar, 50)
            If Thorax = "" Then
                arParameters(2).Value = DBNull.Value
            Else
                arParameters(2).Value = Thorax
            End If
            arParameters(3) = New SqlParameter("@Spine", SqlDbType.NVarChar, 50)
            If Spine = "" Then
                arParameters(3).Value = DBNull.Value
            Else
                arParameters(3).Value = Spine
            End If
            arParameters(4) = New SqlParameter("@Limbs", SqlDbType.NVarChar, 50)
            If Limbs = "" Then
                arParameters(4).Value = DBNull.Value
            Else
                arParameters(4).Value = Limbs
            End If
            arParameters(5) = New SqlParameter("@Hands", SqlDbType.NVarChar, 50)
            If Hands = "" Then
                arParameters(5).Value = DBNull.Value
            Else
                arParameters(5).Value = Hands
            End If
            arParameters(6) = New SqlParameter("@Feet", SqlDbType.NVarChar, 50)
            If Feet = "" Then
                arParameters(6).Value = DBNull.Value
            Else
                arParameters(6).Value = Feet
            End If
            arParameters(7) = New SqlParameter("@Cardiac", SqlDbType.NVarChar, 50)
            If Cardiac = "" Then
                arParameters(7).Value = DBNull.Value
            Else
                arParameters(7).Value = Cardiac
            End If
            arParameters(8) = New SqlParameter("@Kidney", SqlDbType.NVarChar, 50)
            If Kidney = "" Then
                arParameters(8).Value = DBNull.Value
            Else
                arParameters(8).Value = Kidney
            End If
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spReferenceGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spReferenceGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function 'GetReference
    '**************************************************************************
    '*  
    '* Name:        GetAllByCriteria
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '* Parameters: bDefault,Index,bAsc
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAllByCriteria(ByVal bDefault As Boolean, ByVal Index As Integer, ByVal bAsc As Boolean, ByVal Lastname As String, ByVal PatientID As Object, ByVal DOB As Object, ByVal PhysicianID As Object) As SqlDataReader
        Dim arParameters(6) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@bDefault", SqlDbType.Bit)
        arParameters(0).Value = IIf(bDefault = True, 1, 0)
        arParameters(1) = New SqlParameter("@Index", SqlDbType.Int)
        arParameters(1).Value = Index
        arParameters(2) = New SqlParameter("@bAsc", SqlDbType.Bit)
        arParameters(2).Value = IIf(bAsc = True, 1, 0)
        arParameters(3) = New SqlParameter("@LastName", SqlDbType.NVarChar, 50)
        If Len(Lastname) = 0 Then
            arParameters(3).Value = DBNull.Value
        Else
            arParameters(3).Value = Lastname
        End If
        arParameters(4) = New SqlParameter("@PatientID", SqlDbType.Int)
        If Len(PatientID) = 0 Then
            arParameters(4).Value = DBNull.Value
        ElseIf CType(PatientID, Integer) <> 0 Then
            arParameters(4).Value = PatientID
        End If
        arParameters(5) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        If Len(DOB) = 0 Then
            arParameters(5).Value = DBNull.Value
        ElseIf IsDate(DOB) Then
            arParameters(5).Value = DOB
        End If
        arParameters(6) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        If Len(PhysicianID) = 0 Then
            arParameters(6).Value = DBNull.Value
        ElseIf CType(PhysicianID, Integer) <> 0 Then
            arParameters(6).Value = PhysicianID
        End If
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoGetAllByCriteria", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoGetAllByCriteria", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetAllByExamination
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '* Parameters: bDefault,Index,bAsc
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAllByExamination(ByVal bDefault As Boolean, ByVal Index As Integer, _
        ByVal bAsc As Boolean, ByVal Lastname As String, ByVal PatientID As Object, _
        ByVal ExamDateFr As Object, ByVal ExamDateTo As Object, _
        ByVal ExaminerID As Object, _
        Optional ByVal GeneticCounseling As Boolean = False, _
        Optional ByVal GetSigned As Short = -1) As SqlDataReader
        Dim arParameters(9) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@bDefault", SqlDbType.Bit)
        arParameters(0).Value = IIf(bDefault = True, 1, 0)
        arParameters(1) = New SqlParameter("@Index", SqlDbType.Int)
        arParameters(1).Value = Index
        arParameters(2) = New SqlParameter("@bAsc", SqlDbType.Bit)
        arParameters(2).Value = IIf(bAsc = True, 1, 0)
        arParameters(3) = New SqlParameter("@LastName", SqlDbType.NVarChar, 50)
        If Len(Lastname) = 0 Then
            arParameters(3).Value = DBNull.Value
        Else
            arParameters(3).Value = Lastname
        End If
        arParameters(4) = New SqlParameter("@PatientID", SqlDbType.Int)
        If Len(PatientID) = 0 Then
            arParameters(4).Value = DBNull.Value
        ElseIf CType(PatientID, Integer) <> 0 Then
            arParameters(4).Value = PatientID
        End If
        arParameters(5) = New SqlParameter("@ExamDateFr", SqlDbType.SmallDateTime)
        If Len(ExamDateFr) = 0 Then
            arParameters(5).Value = DBNull.Value
        ElseIf IsDate(ExamDateFr) Then
            arParameters(5).Value = ExamDateFr
        End If
        arParameters(6) = New SqlParameter("@ExamDateTo", SqlDbType.SmallDateTime)
        If Len(ExamDateTo) = 0 Then
            arParameters(6).Value = DBNull.Value
        ElseIf IsDate(ExamDateTo) Then
            arParameters(6).Value = ExamDateTo
        End If
        arParameters(7) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        If Len(ExaminerID) = 0 Then
            arParameters(7).Value = DBNull.Value
        ElseIf CType(ExaminerID, Integer) <> 0 Then
            arParameters(7).Value = ExaminerID
        Else
            arParameters(7).Value = DBNull.Value
        End If
        arParameters(8) = New SqlParameter("@GeneticCounseling", SqlDbType.Bit)
        If GeneticCounseling = False Then
            arParameters(8).Value = DBNull.Value
        Else
            arParameters(8).Value = GeneticCounseling
        End If
        arParameters(9) = New SqlParameter("@Signed", SqlDbType.VarChar, 50)
        If GetSigned = 1 Then
            arParameters(9).Value = "Yes"
        ElseIf GetSigned = 0 Then
            arParameters(9).Value = "No"
        Else
            arParameters(9).Value = DBNull.Value
        End If
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoGetAllByExamination", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoGetAllByExamination", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetAllFlaggedChard
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '* Parameters: bDefault,Index,bAsc
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAllFlaggedChart(ByVal bDefault As Boolean, ByVal Index As Integer, _
        ByVal bAsc As Boolean, ByVal ChartID As Integer, _
        ByVal DocRecStatID As Integer) As SqlDataReader
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@bDefault", SqlDbType.Bit)
        arParameters(0).Value = IIf(bDefault = True, 1, 0)
        arParameters(1) = New SqlParameter("@Index", SqlDbType.Int)
        arParameters(1).Value = Index
        arParameters(2) = New SqlParameter("@bAsc", SqlDbType.Bit)
        arParameters(2).Value = IIf(bAsc = True, 1, 0)
        arParameters(3) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(3).Value = IIf(ChartID = 0, DBNull.Value, ChartID)
        arParameters(4) = New SqlParameter("@DocRecStatID", SqlDbType.Int)
        arParameters(4).Value = IIf(DocRecStatID = 0, DBNull.Value, DocRecStatID)

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoGetAllFlaggedChartRev", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoGetAllFlaggedChartRev", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetAllFlaggedChard
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '* Parameters: bDefault,Index,bAsc
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetDocChart(ByVal bDefault As Boolean, ByVal Index As Integer, _
        ByVal bAsc As Boolean, ByVal LabSiteID As Integer, ByVal ChartID As Integer, _
        ByVal DocRecLabTypeID As Integer, ByVal DocRecStatID As Integer, ByVal bLabsOnly As Boolean, _
        ByVal bSuppressRevWNoComments As Boolean, ByVal bRevOnly As Boolean, _
        ByVal ExamDateFr As String, ByVal ExamDateTo As String, ByVal ExaminerID As Integer) As SqlDataReader
        Dim arParameters(12) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@bDefault", SqlDbType.Bit)
        arParameters(0).Value = IIf(bDefault = True, 1, 0)
        arParameters(1) = New SqlParameter("@Index", SqlDbType.Int)
        arParameters(1).Value = Index
        arParameters(2) = New SqlParameter("@bAsc", SqlDbType.Bit)
        arParameters(2).Value = IIf(bAsc = True, 1, 0)
        arParameters(3) = New SqlParameter("@LabSiteID", SqlDbType.Int)
        arParameters(3).Value = IIf(LabSiteID = 0, DBNull.Value, LabSiteID)
        arParameters(4) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(4).Value = IIf(ChartID = 0, DBNull.Value, ChartID)
        arParameters(5) = New SqlParameter("@DocRecLabTypeID", SqlDbType.Int)
        arParameters(5).Value = IIf(DocRecLabTypeID = 0, DBNull.Value, DocRecLabTypeID)
        arParameters(6) = New SqlParameter("@DocRecStatID", SqlDbType.Int)
        arParameters(6).Value = IIf(DocRecStatID = 0, DBNull.Value, DocRecStatID)
        arParameters(7) = New SqlParameter("@bLabsOnly", SqlDbType.Bit)
        arParameters(7).Value = IIf(bLabsOnly = True, 1, DBNull.Value)
        arParameters(8) = New SqlParameter("@bSuppressRevWNoComments", SqlDbType.Bit)
        arParameters(8).Value = IIf(bSuppressRevWNoComments = True, 1, DBNull.Value)
        arParameters(9) = New SqlParameter("@bRevOnly", SqlDbType.Bit)
        arParameters(9).Value = IIf(bRevOnly = True, 1, DBNull.Value)
        arParameters(10) = New SqlParameter("@ExamDateFr", SqlDbType.SmallDateTime)
        If Len(ExamDateFr) = 0 Then
            arParameters(10).Value = DBNull.Value
        ElseIf IsDate(ExamDateFr) Then
            arParameters(10).Value = ExamDateFr
        End If
        arParameters(11) = New SqlParameter("@ExamDateTo", SqlDbType.SmallDateTime)
        If Len(ExamDateTo) = 0 Then
            arParameters(11).Value = DBNull.Value
        ElseIf IsDate(ExamDateTo) Then
            arParameters(11).Value = ExamDateTo
        End If
        arParameters(12) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        If ExaminerID = 0 Then
            arParameters(12).Value = DBNull.Value
        Else
            arParameters(12).Value = ExaminerID
        End If
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoGetAllDocChart", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoGetAllDocChart", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetAllByDiagnosis
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '* Parameters: bDefault,Index,bAsc
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAllByDiagnosis(ByVal bDefault As Boolean, ByVal Index As Integer, ByVal bAsc As Boolean, ByVal DiagnosisID As Object) As SqlDataReader
        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@bDefault", SqlDbType.Bit)
        arParameters(0).Value = IIf(bDefault = True, 1, 0)
        arParameters(1) = New SqlParameter("@Index", SqlDbType.Int)
        arParameters(1).Value = Index
        arParameters(2) = New SqlParameter("@bAsc", SqlDbType.Bit)
        arParameters(2).Value = IIf(bAsc = True, 1, 0)
        arParameters(3) = New SqlParameter("@DiagnosisID", SqlDbType.Int)

        If Len(DiagnosisID) = 0 Then
            arParameters(3).Value = DBNull.Value
        ElseIf CType(DiagnosisID, Integer) <> 0 Then
            arParameters(3).Value = DiagnosisID
        End If

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoGetAllByDiagnosis", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoGetAllByDiagnosis", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function

    '**************************************************************************
    '*  
    '* Name:        GetPatientLast
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetPatientLast() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoGetLastName")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoGetLastName")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetPatientLastByPhysicianID
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetPatientLastByPhysicianID(ByVal PhysicianID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(0).Value = PhysicianID
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoGetLastName", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoGetLastName", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetPatientID
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetPatientID() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spChartGetPatientIdAll")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spChartGetPatientIdAll")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function

    '**************************************************************************
    '*  
    '* Name:        GetChartID
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetChartID() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spChartGetChartIDAll")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spChartGetChartIDAll")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function

    '**************************************************************************
    '*  
    '* Name:        GetExaminers
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetExaminers() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminerGetAll")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spExaminerGetAll")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetExaminers
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetExaminers(ByVal ShowAll As Short, Optional ByVal ActiveOnly As Short = 1) As SqlDataReader

        ' Call stored procedure and return the data
        Try
            Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters
            ' Set the parameters
            arParameters(0) = New SqlParameter("@ShowAll", SqlDbType.Bit)
            arParameters(0).Value = ShowAll
            arParameters(1) = New SqlParameter("@ActiveOnly", SqlDbType.Bit)
            arParameters(1).Value = ActiveOnly

            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminerGetAll", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spExaminerGetAll", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function

    '**************************************************************************
    '*  
    '* Name:        GetExaminers2
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetExaminers2(Optional ByVal ActiveOnly As Short = 1) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ActiveOnly", SqlDbType.Bit)
        arParameters(0).Value = ActiveOnly
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spExaminer2GetAll", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spExaminer2GetAll", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetHospitals
    '*
    '* Description: Returns all records in the [Hospitals] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetHospitals() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spHospitalGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spHospitalGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetPharmacy
    '*
    '* Description: Returns all records in the [Pharmacy] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetPharmacy() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPharmacyGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPharmacyGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetSite
    '*
    '* Description: Returns all records in the [Site] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetSite() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spSiteGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spSiteGet")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetWorkingDiagnoses
    '*
    '* Description: Returns all records in the [WorkingDiagnoses] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetWorkingDiagnoses(ByVal Diagnosis As String) As SqlDataReader

        ' Call stored procedure and return the data
        Try
            Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
            ' Set the parameters
            arParameters(0) = New SqlParameter("@Diagnosis", SqlDbType.VarChar, 100)
            arParameters(0).Value = Diagnosis
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spDiagnosesGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spDiagnosesGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function

    '**************************************************************************
    '*  
    '* Name:        GetDefault
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetDefault() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spDefaultsGetAll")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spDefaultsGetAll")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function

    '**************************************************************************
    '*  
    '* Name:        GetSites
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetSites() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spSiteGetAll")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spSiteGetAll")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        TestConnection
    '*
    '* Description: Test Connection String
    '*
    '*
    '* Returns:     Boolean indicating if record was updated or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function TestConnection() As Boolean
        Dim intRecordsAffected As Integer = 0
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.Text, "select @@version")
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.Text, "select @@version")
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Globals.bAppContinue = False
        End Try
        ' Return boolean indicating if record was updated.
        Return (intRecordsAffected <> 0)
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetByKey
    '*
    '* Description: Gets all the values of a record in the [PatientInfo]and Chart tables
    '*              identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function GetByKey(ByVal ID As Integer, ByRef PatientID As Integer, ByRef MedicalRecord As String, _
    ByRef PatientLast As String, ByRef PatientFirst As String, ByRef DOB As Date, _
    ByRef Race As String, ByRef Gravida As Integer, ByRef Para As Integer, ByRef SAB As String, ByRef TOP As String, _
    ByRef TERM As String, ByRef Living As String, ByRef EDC As Date, ByRef LMP As Date, ByRef EarlyUS As Date, _
    ByRef UseEDCBy As String, ByRef RefDX As Integer, ByRef PhysicianID As Integer, ByRef DelHospitalID As Integer, _
    ByRef SiteID As Integer, ByRef DateCreated As Date, ByRef Type As String, ByRef RH As String, ByRef AntiBody As String, _
    ByRef PreWeight As Integer, ByRef Height As Integer, ByRef SocialSecurity As String, ByRef Allergies As String, ByRef ExamNumber As Integer, _
     ByRef UserID As String, _
     ByRef UpdatedBy As String, _
     ByRef UpdatedDate As Date, _
    ByRef LastOpenedBy As String, _
     ByRef LastOpenedDate As Date, ByRef TSUpdate As Integer, ByRef Workstation As String, _
     ByRef DefaultExaminerID As Integer) As Boolean

        Dim TestNull As Object
        Dim arParameters(39) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ChartFields.fldChartID) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(Me.ChartFields.fldChartID).Value = ID
        arParameters(Me.ChartFields.fldPatientID) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(Me.ChartFields.fldPatientID).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldMedicalRecord) = New SqlParameter("@MedicalRecord", SqlDbType.NVarChar, 50)
        arParameters(Me.ChartFields.fldMedicalRecord).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldLastName) = New SqlParameter("@PatientLast", SqlDbType.NVarChar, 50)
        arParameters(Me.ChartFields.fldLastName).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldFirst) = New SqlParameter("@PatientFirst", SqlDbType.NVarChar, 50)
        arParameters(Me.ChartFields.fldFirst).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldDOB) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        arParameters(Me.ChartFields.fldDOB).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldRace) = New SqlParameter("@Race", SqlDbType.NVarChar, 50)
        arParameters(Me.ChartFields.fldRace).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldGravida) = New SqlParameter("@Gravida", SqlDbType.Int)
        arParameters(Me.ChartFields.fldGravida).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldPara) = New SqlParameter("@Para", SqlDbType.Int)
        arParameters(Me.ChartFields.fldPara).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldSAB) = New SqlParameter("@SAB", SqlDbType.Int)
        arParameters(Me.ChartFields.fldSAB).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldTop) = New SqlParameter("@Top", SqlDbType.Int)
        arParameters(Me.ChartFields.fldTop).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldTerm) = New SqlParameter("@Term", SqlDbType.Int)
        arParameters(Me.ChartFields.fldTerm).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldLiving) = New SqlParameter("@Living", SqlDbType.Int)
        arParameters(Me.ChartFields.fldLiving).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldEDC) = New SqlParameter("@EDC", SqlDbType.SmallDateTime)
        arParameters(Me.ChartFields.fldEDC).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldLMP) = New SqlParameter("@LMP", SqlDbType.SmallDateTime)
        arParameters(Me.ChartFields.fldLMP).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldEarlyUs) = New SqlParameter("@EarlyUS", SqlDbType.SmallDateTime)
        arParameters(Me.ChartFields.fldEarlyUs).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldUseEDCBy) = New SqlParameter("@UseEDCBy", SqlDbType.NVarChar, 50)
        arParameters(Me.ChartFields.fldUseEDCBy).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldRefDX) = New SqlParameter("@RefDX", SqlDbType.Int)
        arParameters(Me.ChartFields.fldRefDX).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldPhysicianID) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(Me.ChartFields.fldPhysicianID).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldDelHospitalID) = New SqlParameter("@DelHospitalID", SqlDbType.Int)
        arParameters(Me.ChartFields.fldDelHospitalID).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldSiteID) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(Me.ChartFields.fldSiteID).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldDateCreated) = New SqlParameter("@DateCreated", SqlDbType.SmallDateTime)
        arParameters(Me.ChartFields.fldDateCreated).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldType) = New SqlParameter("@Type", SqlDbType.VarChar, 50)
        arParameters(Me.ChartFields.fldType).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldRH) = New SqlParameter("@RH", SqlDbType.VarChar, 50)
        arParameters(Me.ChartFields.fldRH).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldAntiBody) = New SqlParameter("@AntiBody", SqlDbType.VarChar, 50)
        arParameters(Me.ChartFields.fldAntiBody).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldPreWeight) = New SqlParameter("@PreWeight", SqlDbType.Int)
        arParameters(Me.ChartFields.fldPreWeight).Direction = ParameterDirection.Output
        arParameters(Me.ChartFields.fldHeight) = New SqlParameter("@Height", SqlDbType.Int)
        arParameters(Me.ChartFields.fldHeight).Direction = ParameterDirection.Output
        arParameters(27) = New SqlParameter("@SocialSecurity", SqlDbType.NVarChar, 50)
        arParameters(27).Direction = ParameterDirection.Output
        arParameters(28) = New SqlParameter("@Allergies", SqlDbType.NVarChar, 255)
        arParameters(28).Direction = ParameterDirection.Output
        arParameters(29) = New SqlParameter("@ExamNumber", SqlDbType.Int)
        arParameters(29).Direction = ParameterDirection.Output
        arParameters(30) = New SqlParameter("@UserID", SqlDbType.NVarChar, 50)
        arParameters(30).Direction = ParameterDirection.Output
        arParameters(31) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(31).Direction = ParameterDirection.Output
        arParameters(32) = New SqlParameter("@UpdatedDate", SqlDbType.SmallDateTime)
        arParameters(32).Direction = ParameterDirection.Output
        arParameters(33) = New SqlParameter("@LastOpenedBy", SqlDbType.NVarChar, 50)
        arParameters(33).Direction = ParameterDirection.Output
        arParameters(34) = New SqlParameter("@LastOpenedDate", SqlDbType.SmallDateTime)
        arParameters(34).Direction = ParameterDirection.Output
        arParameters(35) = New SqlParameter("@TSUpdate", SqlDbType.BigInt)
        arParameters(35).Direction = ParameterDirection.Output
        arParameters(36) = New SqlParameter("@InUserID", SqlDbType.NVarChar, 50)
        arParameters(36).Value = Globals.UserName
        arParameters(37) = New SqlParameter("@Workstation", SqlDbType.NVarChar, 100)
        arParameters(37).Value = Environment.MachineName.ToString
        arParameters(38) = New SqlParameter("@role", SqlDbType.NVarChar, 256)
        arParameters(38).Value = UCase(Globals.UserRole)
        arParameters(39) = New SqlParameter("@DefaultExaminerID", SqlDbType.Int)
        arParameters(39).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.ChartFields.fldPatientID).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            PatientID = ProcessNull.GetInt32(arParameters(Me.ChartFields.fldPatientID).Value)
            MedicalRecord = ProcessNull.GetString(arParameters(Me.ChartFields.fldMedicalRecord).Value)
            PatientLast = ProcessNull.GetString(arParameters(Me.ChartFields.fldLastName).Value)
            PatientLast = PatientLast.Trim()
            PatientFirst = ProcessNull.GetString(arParameters(Me.ChartFields.fldFirst).Value)
            PatientFirst = PatientFirst.Trim()
            DOB = ProcessNull.GetDate(arParameters(Me.ChartFields.fldDOB).Value)
            Race = ProcessNull.GetString(arParameters(Me.ChartFields.fldRace).Value)
            Race = Race.Trim()
            Gravida = ProcessNull.GetInt32(arParameters(Me.ChartFields.fldGravida).Value)
            Para = ProcessNull.GetInt32(arParameters(Me.ChartFields.fldPara).Value)
            SAB = ProcessNull.GetString(arParameters(Me.ChartFields.fldSAB).Value)
            TOP = ProcessNull.GetString(arParameters(Me.ChartFields.fldTop).Value)
            TERM = ProcessNull.GetString(arParameters(Me.ChartFields.fldTerm).Value)
            Living = ProcessNull.GetString(arParameters(Me.ChartFields.fldLiving).Value)
            EDC = ProcessNull.GetDate(arParameters(Me.ChartFields.fldEDC).Value)
            LMP = ProcessNull.GetDate(arParameters(Me.ChartFields.fldLMP).Value)
            EarlyUS = ProcessNull.GetDate(arParameters(Me.ChartFields.fldEarlyUs).Value)
            UseEDCBy = ProcessNull.GetString(arParameters(Me.ChartFields.fldUseEDCBy).Value)
            UseEDCBy = UseEDCBy.Trim()
            RefDX = ProcessNull.GetInt32(arParameters(Me.ChartFields.fldRefDX).Value)
            PhysicianID = ProcessNull.GetInt32(arParameters(Me.ChartFields.fldPhysicianID).Value)
            DelHospitalID = ProcessNull.GetInt32(arParameters(Me.ChartFields.fldDelHospitalID).Value)
            SiteID = ProcessNull.GetInt32(arParameters(Me.ChartFields.fldSiteID).Value)
            DateCreated = ProcessNull.GetDate(arParameters(Me.ChartFields.fldDateCreated).Value)
            Type = ProcessNull.GetString(arParameters(Me.ChartFields.fldType).Value)
            RH = ProcessNull.GetString(arParameters(Me.ChartFields.fldRH).Value)
            AntiBody = ProcessNull.GetString(arParameters(Me.ChartFields.fldAntiBody).Value)
            PreWeight = ProcessNull.GetInt32(arParameters(Me.ChartFields.fldPreWeight).Value)
            Height = ProcessNull.GetInt32(arParameters(Me.ChartFields.fldHeight).Value)
            SocialSecurity = ProcessNull.GetString(arParameters(27).Value)
            Allergies = ProcessNull.GetString(arParameters(28).Value)
            ExamNumber = ProcessNull.GetInt32(arParameters(29).Value)
            UserID = ProcessNull.GetString(arParameters(30).Value)
            UpdatedBy = ProcessNull.GetString(arParameters(31).Value)
            UpdatedDate = ProcessNull.GetDate(arParameters(32).Value)
            LastOpenedBy = ProcessNull.GetString(arParameters(33).Value)
            LastOpenedDate = ProcessNull.GetDate(arParameters(34).Value)
            TSUpdate = CType(arParameters(35).Value, Integer)
            DefaultExaminerID = CType(arParameters(39).Value, Integer)
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        GetChart_LockByKey
    '*
    '* Description: Gets all the values of a record in the [PatientInfo]and Chart tables
    '*              identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function GetChart_LockByKey(ByRef ID As Integer, _
                                ByRef OpenedBy As String, _
                                ByRef OpenedDate As Date, _
                                ByRef Workstation As String) As Boolean

        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ChartFields.fldChartID) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(Me.ChartFields.fldChartID).Value = ID
        arParameters(1) = New SqlParameter("@OpenedBy", SqlDbType.NVarChar, 50)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@OpenedDate", SqlDbType.SmallDateTime)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@WorkStation", SqlDbType.NVarChar, 100)
        arParameters(3).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChart_LockGet", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChart_LockGet", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(1).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            OpenedBy = ProcessNull.GetString(arParameters(1).Value)
            OpenedDate = ProcessNull.GetDate(arParameters(2).Value)
            Workstation = ProcessNull.GetString(arParameters(3).Value)
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        GetIntakeByKey
    '*
    '* Description: Gets all the values of a record in the [PatientInfo]and Chart tables
    '*              identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function GetIntakeByKey(ByVal ID As Integer, _
        ByRef PatientLast As String, _
        ByRef PatientFirst As String, _
        ByRef DOB As Date, _
        ByRef Type As String, _
        ByRef RH As String, _
        ByRef AntiBody As String, _
        ByRef Gravida As Integer, _
        ByRef Para As Integer, _
        ByRef SAB As String, _
        ByRef TOP As String, _
        ByRef Term As String, _
        ByRef Living As String, _
        ByRef LMP As Date, _
        ByRef EDC As Date, _
        ByRef RefDx As Integer, _
        ByRef DelHospitalID As Integer, _
        ByRef PhysicianID As Integer, _
        ByRef Normal As Short, _
        ByRef NormalComments As String, _
        ByRef Bleeding As Short, _
        ByRef BleedingComments As String, _
        ByRef Cramping As Short, _
        ByRef CrampingComments As String, _
        ByRef Excess As Short, _
        ByRef ExcessComments As String, _
        ByRef Radiation As Short, _
        ByRef RadiationComments As String, _
        ByRef Chemicals As Short, _
        ByRef ChemicalsComments As String, _
        ByRef Smoking As Short, _
        ByRef SmokingComments As String, _
        ByRef Alcohol As Short, _
        ByRef AlcoholComments As String, _
        ByRef Drugs As Short, _
        ByRef DrugsComments As String, _
        ByRef Fever As Short, _
        ByRef FeverComments As String, _
        ByRef MedicalHx As Short, _
        ByRef MedicalHistory As String, _
        ByRef SurgicalHx As Short, _
        ByRef SurgicalHistory As String, _
        ByRef GynHx As Short, _
        ByRef GynHistory As String, _
        ByRef FamilyHx As Short, _
        ByRef FamilyHistory As String, _
        ByRef SocialHistory As String, _
        ByRef Transfusion As String, _
        ByRef BirthDefectsMat As Short, _
        ByRef DefectsMatco As String, _
        ByRef BirthDefectsPat As Short, _
        ByRef DefectsPatco As String, _
        ByRef Medications As String) As Boolean

        Dim arParameters(52) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@PatientLast", SqlDbType.NVarChar, 50)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@PatientFirst", SqlDbType.NVarChar, 50)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@Type", SqlDbType.NVarChar, 50)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@RH", SqlDbType.NVarChar, 50)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@AntiBody", SqlDbType.NVarChar, 50)
        arParameters(6).Direction = ParameterDirection.Output
        arParameters(7) = New SqlParameter("@Gravida", SqlDbType.Int)
        arParameters(7).Direction = ParameterDirection.Output
        arParameters(8) = New SqlParameter("@Para", SqlDbType.Int)
        arParameters(8).Direction = ParameterDirection.Output
        arParameters(9) = New SqlParameter("@SAB", SqlDbType.NVarChar, 50)
        arParameters(9).Direction = ParameterDirection.Output
        arParameters(10) = New SqlParameter("@TOP", SqlDbType.NVarChar, 50)
        arParameters(10).Direction = ParameterDirection.Output
        arParameters(11) = New SqlParameter("@Term", SqlDbType.NVarChar, 50)
        arParameters(11).Direction = ParameterDirection.Output
        arParameters(12) = New SqlParameter("@Living", SqlDbType.NVarChar, 50)
        arParameters(12).Direction = ParameterDirection.Output
        arParameters(13) = New SqlParameter("@LMP", SqlDbType.SmallDateTime)
        arParameters(13).Direction = ParameterDirection.Output
        arParameters(14) = New SqlParameter("@EDC", SqlDbType.SmallDateTime)
        arParameters(14).Direction = ParameterDirection.Output
        arParameters(15) = New SqlParameter("@RefDx", SqlDbType.Int)
        arParameters(15).Direction = ParameterDirection.Output
        arParameters(16) = New SqlParameter("@DelHospitalID", SqlDbType.Int)
        arParameters(16).Direction = ParameterDirection.Output
        arParameters(17) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(17).Direction = ParameterDirection.Output
        arParameters(18) = New SqlParameter("@Normal", SqlDbType.SmallInt)
        arParameters(18).Direction = ParameterDirection.Output
        arParameters(19) = New SqlParameter("@NormalComments", SqlDbType.NVarChar, 50)
        arParameters(19).Direction = ParameterDirection.Output
        arParameters(20) = New SqlParameter("@Bleeding", SqlDbType.SmallInt)
        arParameters(20).Direction = ParameterDirection.Output
        arParameters(21) = New SqlParameter("@BleedingComments", SqlDbType.NVarChar, 50)
        arParameters(21).Direction = ParameterDirection.Output
        arParameters(22) = New SqlParameter("@Cramping", SqlDbType.SmallInt)
        arParameters(22).Direction = ParameterDirection.Output
        arParameters(23) = New SqlParameter("@CrampingComments", SqlDbType.NVarChar, 50)
        arParameters(23).Direction = ParameterDirection.Output
        arParameters(24) = New SqlParameter("@Excess", SqlDbType.SmallInt)
        arParameters(24).Direction = ParameterDirection.Output
        arParameters(25) = New SqlParameter("@ExcessComments", SqlDbType.NVarChar, 50)
        arParameters(25).Direction = ParameterDirection.Output
        arParameters(26) = New SqlParameter("@Radiation", SqlDbType.SmallInt)
        arParameters(26).Direction = ParameterDirection.Output
        arParameters(27) = New SqlParameter("@RadiationComments", SqlDbType.NVarChar, 50)
        arParameters(27).Direction = ParameterDirection.Output
        arParameters(28) = New SqlParameter("@Chemicals", SqlDbType.SmallInt)
        arParameters(28).Direction = ParameterDirection.Output
        arParameters(29) = New SqlParameter("@ChemicalsComments", SqlDbType.NVarChar, 50)
        arParameters(29).Direction = ParameterDirection.Output
        arParameters(30) = New SqlParameter("@Smoking", SqlDbType.SmallInt)
        arParameters(30).Direction = ParameterDirection.Output
        arParameters(31) = New SqlParameter("@SmokingComments", SqlDbType.NVarChar, 50)
        arParameters(31).Direction = ParameterDirection.Output
        arParameters(32) = New SqlParameter("@Alcohol", SqlDbType.SmallInt)
        arParameters(32).Direction = ParameterDirection.Output
        arParameters(33) = New SqlParameter("@AlcoholComments", SqlDbType.NVarChar, 50)
        arParameters(33).Direction = ParameterDirection.Output
        arParameters(34) = New SqlParameter("@Drugs", SqlDbType.SmallInt)
        arParameters(34).Direction = ParameterDirection.Output
        arParameters(35) = New SqlParameter("@DrugsComments", SqlDbType.NVarChar, 50)
        arParameters(35).Direction = ParameterDirection.Output
        arParameters(36) = New SqlParameter("@Fever", SqlDbType.SmallInt)
        arParameters(36).Direction = ParameterDirection.Output
        arParameters(37) = New SqlParameter("@FeverComments", SqlDbType.NVarChar, 50)
        arParameters(37).Direction = ParameterDirection.Output
        arParameters(38) = New SqlParameter("@MedicalHx", SqlDbType.SmallInt)
        arParameters(38).Direction = ParameterDirection.Output
        arParameters(39) = New SqlParameter("@MedicalHistory", SqlDbType.NVarChar, 50)
        arParameters(39).Direction = ParameterDirection.Output
        arParameters(40) = New SqlParameter("@SurgicalHx", SqlDbType.SmallInt)
        arParameters(40).Direction = ParameterDirection.Output
        arParameters(41) = New SqlParameter("@SurgicalHistory", SqlDbType.NVarChar, 50)
        arParameters(41).Direction = ParameterDirection.Output
        arParameters(42) = New SqlParameter("@GynHx", SqlDbType.SmallInt)
        arParameters(42).Direction = ParameterDirection.Output
        arParameters(43) = New SqlParameter("@GynHistory", SqlDbType.NVarChar, 50)
        arParameters(43).Direction = ParameterDirection.Output
        arParameters(44) = New SqlParameter("@FamilyHx", SqlDbType.SmallInt)
        arParameters(44).Direction = ParameterDirection.Output
        arParameters(45) = New SqlParameter("@FamilyHistory", SqlDbType.NVarChar, 50)
        arParameters(45).Direction = ParameterDirection.Output
        arParameters(46) = New SqlParameter("@SocialHistory", SqlDbType.NVarChar, 50)
        arParameters(46).Direction = ParameterDirection.Output
        arParameters(47) = New SqlParameter("@Transfusion", SqlDbType.NVarChar, 50)
        arParameters(47).Direction = ParameterDirection.Output
        arParameters(48) = New SqlParameter("@BirthDefectsMat", SqlDbType.SmallInt)
        arParameters(48).Direction = ParameterDirection.Output
        arParameters(49) = New SqlParameter("@DefectsMatco", SqlDbType.NVarChar, 50)
        arParameters(49).Direction = ParameterDirection.Output
        arParameters(50) = New SqlParameter("@BirthDefectsPat", SqlDbType.SmallInt)
        arParameters(50).Direction = ParameterDirection.Output
        arParameters(51) = New SqlParameter("@DefectsPatco", SqlDbType.NVarChar, 50)
        arParameters(51).Direction = ParameterDirection.Output
        arParameters(52) = New SqlParameter("@Medications", SqlDbType.NVarChar, 255)
        arParameters(52).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartIntakeGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartIntakeGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(0).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            PatientLast = ProcessNull.GetString(arParameters(1).Value)
            PatientFirst = ProcessNull.GetString(arParameters(2).Value)
            DOB = ProcessNull.GetDateTime(arParameters(3).Value)
            Type = ProcessNull.GetString(arParameters(4).Value)
            RH = ProcessNull.GetString(arParameters(5).Value)
            AntiBody = ProcessNull.GetString(arParameters(6).Value)
            Gravida = ProcessNull.GetInt32(arParameters(7).Value)
            Para = ProcessNull.GetInt32(arParameters(8).Value)
            SAB = ProcessNull.GetString(arParameters(9).Value)
            TOP = ProcessNull.GetString(arParameters(10).Value)
            Term = ProcessNull.GetString(arParameters(11).Value)
            Living = ProcessNull.GetString(arParameters(12).Value)
            LMP = ProcessNull.GetDateTime(arParameters(13).Value)
            EDC = ProcessNull.GetDateTime(arParameters(14).Value)
            RefDx = ProcessNull.GetInt32(arParameters(15).Value)
            DelHospitalID = ProcessNull.GetInt32(arParameters(16).Value)
            PhysicianID = ProcessNull.GetInt32(arParameters(17).Value)
            Normal = ProcessNull.GetInt16(arParameters(18).Value)
            NormalComments = ProcessNull.GetString(arParameters(19).Value)
            Bleeding = ProcessNull.GetInt16(arParameters(20).Value)
            BleedingComments = ProcessNull.GetString(arParameters(21).Value)
            Cramping = ProcessNull.GetInt16(arParameters(22).Value)
            CrampingComments = ProcessNull.GetString(arParameters(23).Value)
            Excess = ProcessNull.GetInt16(arParameters(24).Value)
            ExcessComments = ProcessNull.GetString(arParameters(25).Value)
            Radiation = ProcessNull.GetInt16(arParameters(26).Value)
            RadiationComments = ProcessNull.GetString(arParameters(27).Value)
            Chemicals = ProcessNull.GetInt16(arParameters(28).Value)
            ChemicalsComments = ProcessNull.GetString(arParameters(29).Value)
            Smoking = ProcessNull.GetInt16(arParameters(30).Value)
            SmokingComments = ProcessNull.GetString(arParameters(31).Value)
            Alcohol = ProcessNull.GetInt16(arParameters(32).Value)
            AlcoholComments = ProcessNull.GetString(arParameters(33).Value)
            Drugs = ProcessNull.GetInt16(arParameters(34).Value)
            DrugsComments = ProcessNull.GetString(arParameters(35).Value)
            Fever = ProcessNull.GetInt16(arParameters(36).Value)
            FeverComments = ProcessNull.GetString(arParameters(37).Value)
            MedicalHx = ProcessNull.GetInt16(arParameters(38).Value)
            MedicalHistory = ProcessNull.GetString(arParameters(39).Value)
            SurgicalHx = ProcessNull.GetInt16(arParameters(40).Value)
            SurgicalHistory = ProcessNull.GetString(arParameters(41).Value)
            GynHx = ProcessNull.GetInt16(arParameters(42).Value)
            GynHistory = ProcessNull.GetString(arParameters(43).Value)
            FamilyHx = ProcessNull.GetInt16(arParameters(44).Value)
            FamilyHistory = ProcessNull.GetString(arParameters(45).Value)
            SocialHistory = ProcessNull.GetString(arParameters(46).Value)
            Transfusion = ProcessNull.GetString(arParameters(47).Value)
            BirthDefectsMat = ProcessNull.GetInt16(arParameters(48).Value)
            DefectsMatco = ProcessNull.GetString(arParameters(49).Value)
            BirthDefectsPat = ProcessNull.GetInt16(arParameters(50).Value)
            DefectsPatco = ProcessNull.GetString(arParameters(51).Value)
            Medications = ProcessNull.GetString(arParameters(52).Value)
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        CheckPatient
    '*
    '* Description: Gets all the values of a record in the [PatientInfo]and Chart tables
    '*              identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function CheckPatient(ByVal PatientLast As String, ByVal PatientFirst As String, _
                            ByVal StreetAddress1 As String, ByVal DOB As Date, _
                            ByVal PatientAutoNum As Short, _
                            ByRef PatientID As Integer) As Boolean

        Dim TestNull As Object
        Dim arParameters(6) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PatientLast", SqlDbType.NVarChar, 50)
        arParameters(0).Value = PatientLast
        arParameters(1) = New SqlParameter("@PatientFirst", SqlDbType.NVarChar, 50)
        arParameters(1).Value = PatientFirst
        arParameters(2) = New SqlParameter("@DOB", SqlDbType.NVarChar, 50)
        arParameters(2).Value = DOB
        arParameters(3) = New SqlParameter("@StreetAddress1", SqlDbType.NVarChar, 50)
        arParameters(3).Value = StreetAddress1
        arParameters(4) = New SqlParameter("@PatientAutoNum", SqlDbType.Bit)
        arParameters(4).Value = PatientAutoNum
        arParameters(5) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(5).Value = PatientID
        arParameters(6) = New SqlParameter("@GetPatientID", SqlDbType.Int)
        arParameters(6).Direction = ParameterDirection.Output


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spCheckPatient", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spCheckPatient", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(0).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            PatientID = ProcessNull.GetInt32(arParameters(6).Value)

            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        CheckPatientID
    '*
    '* Description: Gets all the values of a record in the [PatientInfo]and Chart tables
    '*              identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function CheckPatientID(ByRef PatientID As Integer) As Boolean

        Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(0).Value = PatientID
        arParameters(1) = New SqlParameter("@PatientIDOut", SqlDbType.Int)
        arParameters(1).Direction = ParameterDirection.Output


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spCheckPatientID", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spCheckPatientID", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(1).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            PatientID = ProcessNull.GetInt32(arParameters(1).Value)

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
    Public Function Update(ByVal ID As Integer, _
                           ByVal MedicalRecord As String, _
                           ByVal DOB As Date, _
                           ByVal Gravida As Integer, _
                           ByVal Para As Integer, _
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
                           ByVal PatientID As Integer, _
                           ByVal SiteID As Integer, _
                           ByVal DelHospitalID As Integer, _
                           ByVal PhysicianID As Integer, _
                           ByVal RefDX As Integer, _
                           ByVal ExamNumber As Integer, _
                           ByVal UpdatedBy As String, _
                           ByVal UpdatedDate As Date, _
                           ByVal LastOpenedBy As String, _
                           ByVal LastOpenedDate As Date, _
                           ByVal TSUpdate As Integer) As Boolean

        Dim arParameters(25) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@MedicalRecord", SqlDbType.NVarChar, 50)
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
        If Gravida = 0 Then
            arParameters(4).Value = DBNull.Value
        Else
            arParameters(4).Value = CType(Gravida, Integer)
        End If
        arParameters(5) = New SqlParameter("@Para", SqlDbType.Int)
        If Para = 0 Then
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
        arParameters(16) = New SqlParameter("@SiteID", SqlDbType.Int)
        If SiteID = Nothing Then
            arParameters(16).Value = DBNull.Value
        Else
            arParameters(16).Value = SiteID
        End If
        arParameters(17) = New SqlParameter("@DelHospitalID", SqlDbType.Int)
        If DelHospitalID = Nothing Then
            arParameters(17).Value = DBNull.Value
        Else
            arParameters(17).Value = DelHospitalID
        End If
        arParameters(18) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        If PhysicianID = Nothing Then
            arParameters(18).Value = DBNull.Value
        Else
            arParameters(18).Value = PhysicianID
        End If
        arParameters(19) = New SqlParameter("@RefDX", SqlDbType.Int)
        If RefDX = Nothing Then
            arParameters(19).Value = DBNull.Value
        Else
            arParameters(19).Value = RefDX
        End If
        arParameters(20) = New SqlParameter("@ExamNumber", SqlDbType.Int)
        If ExamNumber = Nothing Then
            arParameters(20).Value = DBNull.Value
        Else
            arParameters(20).Value = ExamNumber
        End If
        arParameters(21) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(21).Value = UpdatedBy
        arParameters(22) = New SqlParameter("@UpdatedDate", SqlDbType.SmallDateTime)
        If UpdatedDate = Nothing Then
            arParameters(22).Value = DBNull.Value
        Else
            arParameters(22).Value = UpdatedDate
        End If
        arParameters(23) = New SqlParameter("@LastOpenedBy", SqlDbType.NVarChar, 50)
        arParameters(23).Value = LastOpenedBy
        arParameters(24) = New SqlParameter("@LastOpenedDate", SqlDbType.SmallDateTime)
        If LastOpenedDate = Nothing Then
            arParameters(24).Value = DBNull.Value
        Else
            arParameters(24).Value = LastOpenedDate
        End If
        arParameters(25) = New SqlParameter("@TSUpdate", SqlDbType.BigInt)
        arParameters(25).Value = TSUpdate
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartUpdate", arParameters)
            End If
        Catch exception As Exception
            If Len(exception.Message) >= 39 Then
                If Left(exception.Message, 39) = "Warning. Cannot find record. Parameters" Then
                    'MessageBox.Show("NO UPDATES ARE DONE ON THIS RECORD DUE TO CONCURRENCY ISSUE!  This record was updated by another user since you have requested a record fetch.", "Task Aborted - Concurrency Problem", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                ExceptionManager.Publish(exception)
            End If
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
                           ByVal BloodType As String, _
                           ByVal RH As String, _
                           ByVal AntiBody As String, _
                           ByVal UpdatedBy As String) As Boolean

        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@BloodType", SqlDbType.NVarChar, 50)
        arParameters(1).Value = BloodType
        arParameters(2) = New SqlParameter("@RH", SqlDbType.NVarChar, 50)
        arParameters(2).Value = RH
        arParameters(3) = New SqlParameter("@Antibody", SqlDbType.NVarChar, 50)
        arParameters(3).Value = AntiBody
        arParameters(4) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(4).Value = UpdatedBy

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartUpdateLimited", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartUpdateLimited", arParameters)
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
    '* Name:        UpdateDefaults
    '*
    '* Description: UpdateDefaultss a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was UpdateDefaultsd or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdateDefaults(ByVal SiteID As Integer, _
                           ByVal ExaminerID As Integer) As Boolean

        Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(0).Value = SiteID
        arParameters(1) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(1).Value = ExaminerID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spDefaultsUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spDefaultsUpdate", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not UpdateDefaultsd.
        If intRecordsAffected = 0 Then
            Return False
        Else

            Return True
        End If

    End Function
    '**************************************************************************
    '*  
    '* Name:        ChartAuditTrail
    '*
    '* Description: ChartAuditTrails a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was ChartAuditTraild or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function ChartAuditTrail(ByVal ChartID As Integer, _
                           ByVal InUserID As String, _
                           ByVal WorkStation As String, _
                           ByVal bReadOnly As Short) As Boolean

        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ChartID
        arParameters(1) = New SqlParameter("@InUserID", SqlDbType.NVarChar, 50)
        arParameters(1).Value = InUserID
        arParameters(2) = New SqlParameter("@ReadOnly", SqlDbType.Bit)
        arParameters(2).Value = bReadOnly
        arParameters(3) = New SqlParameter("@Workstation", SqlDbType.NVarChar, 100)
        arParameters(3).Value = WorkStation

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartOpenAudit", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartOpenAudit", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not ChartAuditTraild.
        If intRecordsAffected = 0 Then
            Return False
        Else
            Return True
        End If

    End Function
    '**************************************************************************
    '*  
    '* Name:        UpdateIntake
    '*
    '* Description: UpdateIntakes a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was UpdateIntaked or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdateIntake(ByVal ID As Integer, _
        ByVal PatientLast As String, _
        ByVal PatientFirst As String, _
        ByVal DOB As Date, _
        ByVal Type As String, _
        ByVal RH As String, _
        ByVal AntiBody As String, _
        ByVal Gravida As Integer, _
        ByVal Para As Integer, _
        ByVal SAB As String, _
        ByVal TOP As String, _
        ByVal Term As String, _
        ByVal Living As String, _
        ByVal LMP As Date, _
        ByVal EDC As Date, _
        ByVal RefDx As Integer, _
        ByVal DelHospitalID As Integer, _
        ByVal PhysicianID As Integer, _
        ByVal Normal As Short, _
        ByVal NormalComments As String, _
        ByVal Bleeding As Short, _
        ByVal BleedingComments As String, _
        ByVal Cramping As Short, _
        ByVal CrampingComments As String, _
        ByVal Excess As Short, _
        ByVal ExcessComments As String, _
        ByVal Radiation As Short, _
        ByVal RadiationComments As String, _
        ByVal Chemicals As Short, _
        ByVal ChemicalsComments As String, _
        ByVal Smoking As Short, _
        ByVal SmokingComments As String, _
        ByVal Alcohol As Short, _
        ByVal AlcoholComments As String, _
        ByVal Drugs As Short, _
        ByVal DrugsComments As String, _
        ByVal Fever As Short, _
        ByVal FeverComments As String, _
        ByVal MedicalHx As Short, _
        ByVal MedicalHistory As String, _
        ByVal SurgicalHx As Short, _
        ByVal SurgicalHistory As String, _
        ByVal GynHx As Short, _
        ByVal GynHistory As String, _
        ByVal FamilyHx As Short, _
        ByVal FamilyHistory As String, _
        ByVal SocialHistory As String, _
        ByVal Transfusion As String, _
        ByVal BirthDefectsMat As Short, _
        ByVal DefectsMatco As String, _
        ByVal BirthDefectsPat As Short, _
        ByVal DefectsPatco As String, _
        ByVal PatientID As Integer, _
        ByVal Medications As String) As Boolean

        Dim arParameters(53) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@PatientLast", SqlDbType.NVarChar, 50)
        arParameters(1).Value = PatientLast
        arParameters(2) = New SqlParameter("@PatientFirst", SqlDbType.NVarChar, 50)
        arParameters(2).Value = PatientFirst
        arParameters(3) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        arParameters(3).Value = DOB
        arParameters(4) = New SqlParameter("@Type", SqlDbType.NVarChar, 50)
        arParameters(4).Value = Type
        arParameters(5) = New SqlParameter("@RH", SqlDbType.NVarChar, 50)
        arParameters(5).Value = RH
        arParameters(6) = New SqlParameter("@AntiBody", SqlDbType.NVarChar, 50)
        arParameters(6).Value = AntiBody
        arParameters(7) = New SqlParameter("@Gravida", SqlDbType.Int)
        arParameters(7).Value = Gravida
        arParameters(8) = New SqlParameter("@Para", SqlDbType.Int)
        arParameters(8).Value = Para
        arParameters(9) = New SqlParameter("@SAB", SqlDbType.NVarChar, 50)
        arParameters(9).Value = SAB
        arParameters(10) = New SqlParameter("@TOP", SqlDbType.NVarChar, 50)
        arParameters(10).Value = TOP
        arParameters(11) = New SqlParameter("@Term", SqlDbType.NVarChar, 50)
        arParameters(11).Value = Term
        arParameters(12) = New SqlParameter("@Living", SqlDbType.NVarChar, 50)
        arParameters(12).Value = Living
        arParameters(13) = New SqlParameter("@LMP", SqlDbType.SmallDateTime)
        If LMP.ToShortDateString > "1/1/1900" Then
            arParameters(13).Value = LMP
        Else
            arParameters(13).Value = DBNull.Value
        End If
        arParameters(14) = New SqlParameter("@EDC", SqlDbType.SmallDateTime)
        If EDC.ToShortDateString > "1/1/1900" Then
            arParameters(14).Value = EDC
        Else
            arParameters(14).Value = DBNull.Value
        End If
        arParameters(15) = New SqlParameter("@RefDx", SqlDbType.Int)
        arParameters(15).Value = RefDx
        arParameters(16) = New SqlParameter("@DelHospitalID", SqlDbType.Int)
        arParameters(16).Value = DelHospitalID
        arParameters(17) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(17).Value = PhysicianID
        arParameters(18) = New SqlParameter("@Normal", SqlDbType.SmallInt)
        arParameters(18).Value = Normal
        arParameters(19) = New SqlParameter("@NormalComments", SqlDbType.NVarChar, 50)
        arParameters(19).Value = NormalComments
        arParameters(20) = New SqlParameter("@Bleeding", SqlDbType.SmallInt)
        arParameters(20).Value = Bleeding
        arParameters(21) = New SqlParameter("@BleedingComments", SqlDbType.NVarChar, 50)
        arParameters(21).Value = BleedingComments
        arParameters(22) = New SqlParameter("@Cramping", SqlDbType.SmallInt)
        arParameters(22).Value = Cramping
        arParameters(23) = New SqlParameter("@CrampingComments", SqlDbType.NVarChar, 50)
        arParameters(23).Value = CrampingComments
        arParameters(24) = New SqlParameter("@Excess", SqlDbType.SmallInt)
        arParameters(24).Value = Excess
        arParameters(25) = New SqlParameter("@ExcessComments", SqlDbType.NVarChar, 50)
        arParameters(25).Value = ExcessComments
        arParameters(26) = New SqlParameter("@Radiation", SqlDbType.SmallInt)
        arParameters(26).Value = Radiation
        arParameters(27) = New SqlParameter("@RadiationComments", SqlDbType.NVarChar, 50)
        arParameters(27).Value = RadiationComments
        arParameters(28) = New SqlParameter("@Chemicals", SqlDbType.SmallInt)
        arParameters(28).Value = Chemicals
        arParameters(29) = New SqlParameter("@ChemicalsComments", SqlDbType.NVarChar, 50)
        arParameters(29).Value = ChemicalsComments
        arParameters(30) = New SqlParameter("@Smoking", SqlDbType.SmallInt)
        arParameters(30).Value = Smoking
        arParameters(31) = New SqlParameter("@SmokingComments", SqlDbType.NVarChar, 50)
        arParameters(31).Value = SmokingComments
        arParameters(32) = New SqlParameter("@Alcohol", SqlDbType.SmallInt)
        arParameters(32).Value = Alcohol
        arParameters(33) = New SqlParameter("@AlcoholComments", SqlDbType.NVarChar, 50)
        arParameters(33).Value = AlcoholComments
        arParameters(34) = New SqlParameter("@Drugs", SqlDbType.SmallInt)
        arParameters(34).Value = Drugs
        arParameters(35) = New SqlParameter("@DrugsComments", SqlDbType.NVarChar, 50)
        arParameters(35).Value = DrugsComments
        arParameters(36) = New SqlParameter("@Fever", SqlDbType.SmallInt)
        arParameters(36).Value = Fever
        arParameters(37) = New SqlParameter("@FeverComments", SqlDbType.NVarChar, 50)
        arParameters(37).Value = FeverComments
        arParameters(38) = New SqlParameter("@MedicalHx", SqlDbType.SmallInt)
        arParameters(38).Value = MedicalHx
        arParameters(39) = New SqlParameter("@MedicalHistory", SqlDbType.NVarChar, 50)
        arParameters(39).Value = MedicalHistory
        arParameters(40) = New SqlParameter("@SurgicalHx", SqlDbType.SmallInt)
        arParameters(40).Value = SurgicalHx
        arParameters(41) = New SqlParameter("@SurgicalHistory", SqlDbType.NVarChar, 50)
        arParameters(41).Value = SurgicalHistory
        arParameters(42) = New SqlParameter("@GynHx", SqlDbType.SmallInt)
        arParameters(42).Value = GynHx
        arParameters(43) = New SqlParameter("@GynHistory", SqlDbType.NVarChar, 50)
        arParameters(43).Value = GynHistory
        arParameters(44) = New SqlParameter("@FamilyHx", SqlDbType.SmallInt)
        arParameters(44).Value = FamilyHx
        arParameters(45) = New SqlParameter("@FamilyHistory", SqlDbType.NVarChar, 50)
        arParameters(45).Value = FamilyHistory
        arParameters(46) = New SqlParameter("@SocialHistory", SqlDbType.NVarChar, 50)
        arParameters(46).Value = SocialHistory
        arParameters(47) = New SqlParameter("@Transfusion", SqlDbType.NVarChar, 50)
        arParameters(47).Value = Transfusion
        arParameters(48) = New SqlParameter("@BirthDefectsMat", SqlDbType.SmallInt)
        arParameters(48).Value = BirthDefectsMat
        arParameters(49) = New SqlParameter("@DefectsMatco", SqlDbType.NVarChar, 50)
        arParameters(49).Value = DefectsMatco
        arParameters(50) = New SqlParameter("@BirthDefectsPat", SqlDbType.SmallInt)
        arParameters(50).Value = BirthDefectsPat
        arParameters(51) = New SqlParameter("@DefectsPatco", SqlDbType.NVarChar, 50)
        arParameters(51).Value = DefectsPatco
        arParameters(52) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(52).Value = PatientID
        arParameters(53) = New SqlParameter("@Medications", SqlDbType.NVarChar, 255)
        arParameters(53).Value = Medications
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartIntakeUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartIntakeUpdate", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not UpdateIntaked.
        If intRecordsAffected = 0 Then
            Return False
        Else

            Return True
        End If

    End Function
    '**************************************************************************
    '*  
    '* Name:        TransposePatientID
    '*
    '* Description: TransposePatientIDs a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was TransposePatientIDd or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function TransposePatientID(ByVal ID1 As Integer, _
        ByVal ID2 As Integer) As Boolean

        Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PatientID1", SqlDbType.Int)
        arParameters(0).Value = ID1
        arParameters(1) = New SqlParameter("@PatientID2", SqlDbType.Int)
        arParameters(1).Value = ID2
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientTranspose", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPatientTranspose", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not TransposePatientIDd.
        If intRecordsAffected = 0 Then
            Return False
        Else
            Return True
        End If

    End Function
    '**************************************************************************
    '*  
    '* Name:        MergeChartID
    '*
    '* Description: MergeChartIDs a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was MergeChartIDd or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function MergeChartID(ByVal ID1 As Integer, _
        ByVal ID2 As Integer) As Boolean

        Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ChartID1", SqlDbType.Int)
        arParameters(0).Value = ID1
        arParameters(1) = New SqlParameter("@ChartID2", SqlDbType.Int)
        arParameters(1).Value = ID2
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartMergeRecs", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartMergeRecs", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not MergeChartIDd.
        If intRecordsAffected = 0 Then
            Return False
        Else
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
                           ByVal PreWeight As Integer, _
                           ByVal Height As Integer, _
                           ByVal UpdatedBy As String) As Boolean

        Dim arParameters(3) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Preweight", SqlDbType.Int)
        If PreWeight = Nothing Then
            arParameters(1).Value = DBNull.Value
        Else
            arParameters(1).Value = PreWeight
        End If
        arParameters(2) = New SqlParameter("@Height", SqlDbType.Int)
        If Height = Nothing Then
            arParameters(2).Value = DBNull.Value
        Else
            arParameters(2).Value = Height
        End If
        arParameters(3) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(3).Value = UpdatedBy
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartUpdateHeight", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartUpdateHeight", arParameters)
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
    '* Returns:     Boolean indicating if record was added or not. 
    '*              True (record added); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByRef ChartID As Integer, _
                    ByVal PatientID As Integer, _
                    ByVal PatientLast As String, _
                    ByVal PatientFirst As String, _
                    ByVal DOB As Date, _
                    ByVal PhysicianID As Integer, _
                    ByVal Gravida As Integer, _
                    ByVal Para As Integer, _
                    ByVal LMP As Date, _
                    ByVal SiteID As Integer, _
                    ByVal PatientAutoNum As Short, _
                    ByVal DelHospitalID As Integer, _
                    ByVal UserID As String, _
                    ByVal DefaultExaminerID As Integer) As Boolean

        Dim arParameters(13) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(1).Value = PatientID
        arParameters(2) = New SqlParameter("@PatientLast", SqlDbType.NVarChar, 50)
        arParameters(2).Value = PatientLast
        arParameters(3) = New SqlParameter("@PatientFirst", SqlDbType.NVarChar, 50)
        arParameters(3).Value = PatientFirst
        arParameters(4) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        arParameters(4).Value = DOB
        arParameters(5) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(5).Value = IIf(PhysicianID = 0, DBNull.Value, PhysicianID)
        arParameters(6) = New SqlParameter("@Gravida", SqlDbType.Int)
        arParameters(6).Value = Gravida
        arParameters(7) = New SqlParameter("@Para", SqlDbType.Int)
        arParameters(7).Value = Para
        arParameters(8) = New SqlParameter("@LMP", SqlDbType.SmallDateTime)
        If LMP = Nothing Then
            arParameters(8).Value = DBNull.Value
        Else
            arParameters(8).Value = LMP
        End If
        arParameters(9) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(9).Value = IIf(SiteID = 0, DBNull.Value, SiteID)
        arParameters(10) = New SqlParameter("@PatientAutoNum", SqlDbType.Bit)
        arParameters(10).Value = PatientAutoNum
        arParameters(11) = New SqlParameter("@DelHospitalID", SqlDbType.Int)
        arParameters(11).Value = DelHospitalID
        arParameters(12) = New SqlParameter("@UserID", SqlDbType.NVarChar, 50)
        arParameters(12).Value = UserID
        arParameters(13) = New SqlParameter("@DefaultExaminerID", SqlDbType.Int)
        arParameters(13).Value = DefaultExaminerID


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            ChartID = CType(arParameters(0).Value, Integer)
            Return True
        End If

    End Function
    '**************************************************************************
    '*  
    '* Name:        AddPatient
    '*
    '* Description: AddPatients a new record to the [PatientInfo] table.
    '*
    '* Returns:     Boolean indicating if record was AddPatiented or not. 
    '*              True (record AddPatiented); False (otherwise).
    '*
    '**************************************************************************
    Public Function AddPatient(ByRef PatientID As Integer, _
                    ByVal MedicalRecord As String, _
                    ByVal PatientLast As String, _
                    ByVal PatientFirst As String, _
                    ByVal SocialSecurity As String, _
                    ByVal DOB As Date, _
                    ByVal Race As String, _
                    ByVal Language As String, _
                    ByVal Type As String, _
                    ByVal RH As String, _
                    ByVal DateCreated As Date, _
                    ByVal PatientAutoNum As Short) As Boolean



        Dim arParameters(12) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(0).Value = PatientID
        arParameters(1) = New SqlParameter("@MedicalRecord", SqlDbType.NVarChar, 50)
        arParameters(1).Value = PatientID
        arParameters(2) = New SqlParameter("@PatientLast", SqlDbType.NVarChar, 50)
        arParameters(2).Value = PatientLast
        arParameters(3) = New SqlParameter("@PatientFirst", SqlDbType.NVarChar, 50)
        arParameters(3).Value = PatientFirst
        arParameters(4) = New SqlParameter("@SocialSecurity", SqlDbType.NVarChar, 50)
        arParameters(4).Value = SocialSecurity
        arParameters(5) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        If DOB = Nothing Then
            arParameters(5).Value = DBNull.Value
        Else
            arParameters(5).Value = DOB
        End If
        arParameters(6) = New SqlParameter("@Race", SqlDbType.NVarChar, 50)
        arParameters(6).Value = Race
        arParameters(7) = New SqlParameter("@Language", SqlDbType.NVarChar, 50)
        arParameters(7).Value = Language
        arParameters(8) = New SqlParameter("@Type", SqlDbType.NVarChar, 50)
        arParameters(8).Value = Type
        arParameters(9) = New SqlParameter("@RH", SqlDbType.NVarChar, 50)
        arParameters(9).Value = RH
        arParameters(10) = New SqlParameter("@DateCreated", SqlDbType.SmallDateTime)
        arParameters(10).Value = Now
        arParameters(11) = New SqlParameter("@PatientAutonum", SqlDbType.Bit)
        arParameters(11).Value = PatientAutoNum
        arParameters(12) = New SqlParameter("@NewPatientID", SqlDbType.Int)
        arParameters(12).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            PatientID = CType(arParameters(12).Value, Integer)
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
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChartDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChartDelete", arParameters)
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
    '* Name:        DeleteChartLock
    '*
    '* Description: Deletes a record from the [PatientInfo] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to delete
    '*
    '* Returns:     Boolean indicating if record was deleted or not. 
    '*              True (record found and deleted); False (otherwise).
    '*
    '**************************************************************************
    Public Function DeleteChartLock(ByVal ID As Integer) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spChart_LockDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spChart_LockDelete", arParameters)
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


End Class 'dalChart
