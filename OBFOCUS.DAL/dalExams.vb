
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalExams
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
Public Class dalExams

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum WdiagnosisFields
        fldID = 0
        fldExamDate = 1
        fldIndication = 2
        fldSigned = 3
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
    '* Name:        GetExams
    '*
    '* Description: Returns all records in the [Labs] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetExams(ByVal ChartID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ChartID
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spExamsGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        SearchAssessment
    '*
    '* Description: Returns all records in the [Labs] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function SearchAssessment(ByVal Term1 As String, ByVal Term2 As String) As SqlDataReader
        Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@Term1", SqlDbType.VarChar, 50)
        arParameters(0).Value = Term1
        arParameters(1) = New SqlParameter("@Term2", SqlDbType.VarChar, 50)
        arParameters(1).Value = Term2
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spSearchAssessmentGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spSearchAssessmentGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
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
    Public Function GetByKey(ByVal ID As Integer, _
                            ByRef Service As Integer, _
                            ByRef ExaminerID As Integer, _
                            ByRef SiteID As Integer, _
                            ByRef Examiner2ID As Integer, _
                            ByRef UltrasonographerID As Integer, _
                            ByRef Encounter As String, _
                            ByRef PhysicianID As Integer, _
                            ByRef EarlyUS As String, _
                            ByRef LMP As String, _
                            ByRef UseEDCBy As String, _
                            ByRef Early As Short, _
                            ByRef NOMS As Short, _
                            ByRef TapeID As String, _
                            ByRef EDC As String, _
                            ByRef Intro As String, _
                            ByRef ProcedureDone As Short, _
                            ByRef PN As Short, _
                            ByRef GYN As Short, _
                            ByRef PatientLast As String, _
                            ByRef PatientFirst As String, _
                            ByRef RH As String, _
                            ByRef Type As String, _
                            ByRef PatientID As Integer, _
                            ByRef ChartID As Integer, _
                            ByRef Signed As String, _
                            ByRef ExamDate As String, _
                            ByRef Adnexa As String, _
                            ByRef Evaluation As String, _
                            ByRef Uterus As String, _
                            ByRef Appearance As String, _
                            ByRef Skin As String, _
                            ByRef HEENT As String, _
                            ByRef Neck As String, _
                            ByRef Heart As String, _
                            ByRef Lung As String, _
                            ByRef Chest As String, _
                            ByRef Abdomen As String, _
                            ByRef Back As String, _
                            ByRef Pelvic As String, _
                            ByRef Rectal As String, _
                            ByRef Extremities As String, _
                            ByRef Neurologic As String, _
                            ByRef ROS As String, _
                            ByRef VitalsOutput As String, _
                            ByRef Complaints As String, _
                            ByRef Assessment As String, _
                            ByRef Recommendation As String, _
                            ByRef LabOutput As String, _
                            ByRef Singleton As String, _
                            ByRef PatientHeight As Integer, _
                            ByRef PreWeight As Integer, _
                            ByRef ReportTypeID As Integer, _
                            ByRef BloodType As String, _
                            ByRef CervixLength As Double, _
                            ByRef InternalOS As String, _
                            ByRef FPressure As String, _
                            ByRef SuppressIntro As Short, _
                            ByRef RoomNumber As String) As Boolean

        Dim arParameters(58) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Service", SqlDbType.Int)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@Examiner2ID", SqlDbType.Int)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@UltrasonographerID", SqlDbType.Int)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@Encounter", SqlDbType.NVarChar, 200)
        arParameters(6).Direction = ParameterDirection.Output
        arParameters(7) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(7).Direction = ParameterDirection.Output
        arParameters(8) = New SqlParameter("@EarlyUS", SqlDbType.SmallDateTime)
        arParameters(8).Direction = ParameterDirection.Output
        arParameters(9) = New SqlParameter("@LMP", SqlDbType.SmallDateTime)
        arParameters(9).Direction = ParameterDirection.Output
        arParameters(10) = New SqlParameter("@UseEDCBy", SqlDbType.NVarChar, 50)
        arParameters(10).Direction = ParameterDirection.Output
        arParameters(11) = New SqlParameter("@Early", SqlDbType.SmallInt)
        arParameters(11).Direction = ParameterDirection.Output
        arParameters(12) = New SqlParameter("@NOMS", SqlDbType.SmallInt)
        arParameters(12).Direction = ParameterDirection.Output
        arParameters(13) = New SqlParameter("@TapeID", SqlDbType.NVarChar, 50)
        arParameters(13).Direction = ParameterDirection.Output
        arParameters(14) = New SqlParameter("@EDC", SqlDbType.SmallDateTime)
        arParameters(14).Direction = ParameterDirection.Output
        arParameters(15) = New SqlParameter("@Intro", SqlDbType.NVarChar, 50)
        arParameters(15).Direction = ParameterDirection.Output
        arParameters(16) = New SqlParameter("@ProcedureDone", SqlDbType.SmallInt)
        arParameters(16).Direction = ParameterDirection.Output
        arParameters(17) = New SqlParameter("@PN", SqlDbType.SmallInt)
        arParameters(17).Direction = ParameterDirection.Output
        arParameters(18) = New SqlParameter("@Gyn", SqlDbType.SmallInt)
        arParameters(18).Direction = ParameterDirection.Output
        arParameters(19) = New SqlParameter("@PatientLast", SqlDbType.NVarChar, 50)
        arParameters(19).Direction = ParameterDirection.Output
        arParameters(20) = New SqlParameter("@PatientFirst", SqlDbType.NVarChar, 50)
        arParameters(20).Direction = ParameterDirection.Output
        arParameters(21) = New SqlParameter("@RH", SqlDbType.NVarChar, 50)
        arParameters(21).Direction = ParameterDirection.Output
        arParameters(22) = New SqlParameter("@Type", SqlDbType.NVarChar, 50)
        arParameters(22).Direction = ParameterDirection.Output
        arParameters(23) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(23).Direction = ParameterDirection.Output
        arParameters(24) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(24).Direction = ParameterDirection.Output
        arParameters(25) = New SqlParameter("@Signed", SqlDbType.NVarChar, 50)
        arParameters(25).Direction = ParameterDirection.Output
        arParameters(26) = New SqlParameter("@ExamDate", SqlDbType.SmallDateTime)
        arParameters(26).Direction = ParameterDirection.Output
        arParameters(27) = New SqlParameter("@Adnexa", SqlDbType.VarChar, 8000)
        arParameters(27).Direction = ParameterDirection.Output
        arParameters(28) = New SqlParameter("@Evaluation", SqlDbType.VarChar, 8000)
        arParameters(28).Direction = ParameterDirection.Output
        arParameters(29) = New SqlParameter("@Uterus", SqlDbType.VarChar, 8000)
        arParameters(29).Direction = ParameterDirection.Output
        arParameters(30) = New SqlParameter("@Appearance", SqlDbType.NVarChar, 100)
        arParameters(30).Direction = ParameterDirection.Output
        arParameters(31) = New SqlParameter("@Skin", SqlDbType.NVarChar, 200)
        arParameters(31).Direction = ParameterDirection.Output
        arParameters(32) = New SqlParameter("@HEENT", SqlDbType.NVarChar, 100)
        arParameters(32).Direction = ParameterDirection.Output
        arParameters(33) = New SqlParameter("@Neck", SqlDbType.NVarChar, 100)
        arParameters(33).Direction = ParameterDirection.Output
        arParameters(34) = New SqlParameter("@Heart", SqlDbType.NVarChar, 100)
        arParameters(34).Direction = ParameterDirection.Output
        arParameters(35) = New SqlParameter("@Lung", SqlDbType.NVarChar, 100)
        arParameters(35).Direction = ParameterDirection.Output
        arParameters(36) = New SqlParameter("@Chest", SqlDbType.NVarChar, 100)
        arParameters(36).Direction = ParameterDirection.Output
        arParameters(37) = New SqlParameter("@Abdomen", SqlDbType.NVarChar, 100)
        arParameters(37).Direction = ParameterDirection.Output
        arParameters(38) = New SqlParameter("@Back", SqlDbType.NVarChar, 100)
        arParameters(38).Direction = ParameterDirection.Output
        arParameters(39) = New SqlParameter("@Pelvic", SqlDbType.NVarChar, 100)
        arParameters(39).Direction = ParameterDirection.Output
        arParameters(40) = New SqlParameter("@Rectal", SqlDbType.NVarChar, 100)
        arParameters(40).Direction = ParameterDirection.Output
        arParameters(41) = New SqlParameter("@Extremities", SqlDbType.NVarChar, 100)
        arParameters(41).Direction = ParameterDirection.Output
        arParameters(42) = New SqlParameter("@Neurologic", SqlDbType.NVarChar, 100)
        arParameters(42).Direction = ParameterDirection.Output
        arParameters(43) = New SqlParameter("@ROS", SqlDbType.VarChar, 8000)
        arParameters(43).Direction = ParameterDirection.Output
        arParameters(44) = New SqlParameter("@VitalsOutput", SqlDbType.VarChar, 8000)
        arParameters(44).Direction = ParameterDirection.Output
        arParameters(45) = New SqlParameter("@Complaints", SqlDbType.VarChar, 8000)
        arParameters(45).Direction = ParameterDirection.Output
        arParameters(46) = New SqlParameter("@Assessment", SqlDbType.VarChar, 8000)
        arParameters(46).Direction = ParameterDirection.Output
        arParameters(47) = New SqlParameter("@Recommendation", SqlDbType.VarChar, 8000)
        arParameters(47).Direction = ParameterDirection.Output
        arParameters(48) = New SqlParameter("@LabOutput", SqlDbType.VarChar, 8000)
        arParameters(48).Direction = ParameterDirection.Output
        arParameters(49) = New SqlParameter("@Singleton", SqlDbType.VarChar, 8000)
        arParameters(49).Direction = ParameterDirection.Output
        arParameters(50) = New SqlParameter("@Height", SqlDbType.Int)
        arParameters(50).Direction = ParameterDirection.Output
        arParameters(51) = New SqlParameter("@PreWeight", SqlDbType.Int)
        arParameters(51).Direction = ParameterDirection.Output
        arParameters(52) = New SqlParameter("@ReportTypeID", SqlDbType.Int)
        arParameters(52).Direction = ParameterDirection.Output
        arParameters(53) = New SqlParameter("@BloodType", SqlDbType.NVarChar, 50)
        arParameters(53).Direction = ParameterDirection.Output
        arParameters(54) = New SqlParameter("@CervixLength", SqlDbType.Real)
        arParameters(54).Direction = ParameterDirection.Output
        arParameters(55) = New SqlParameter("@InternalOS", SqlDbType.NVarChar, 100)
        arParameters(55).Direction = ParameterDirection.Output
        arParameters(56) = New SqlParameter("@FPressure", SqlDbType.NVarChar, 100)
        arParameters(56).Direction = ParameterDirection.Output
        arParameters(57) = New SqlParameter("@SuppressIntro", SqlDbType.SmallInt)
        arParameters(57).Direction = ParameterDirection.Output
        arParameters(58) = New SqlParameter("@RoomNumber", SqlDbType.NVarChar, 15)
        arParameters(58).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamsGetByKey", arParameters)
            End If

            ' Return False if data was not found.
            If arParameters(0).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Service = ProcessNull.GetInt32(arParameters(1).Value)
            ExaminerID = ProcessNull.GetInt32(arParameters(2).Value)
            SiteID = ProcessNull.GetInt32(arParameters(3).Value)
            Examiner2ID = ProcessNull.GetInt32(arParameters(4).Value)
            UltrasonographerID = ProcessNull.GetInt32(arParameters(5).Value)
            Encounter = ProcessNull.GetString(arParameters(6).Value)
            Encounter = Encounter.Trim()
            PhysicianID = ProcessNull.GetInt32(arParameters(7).Value)
            EarlyUS = ProcessNull.GetString(arParameters(8).Value)
            LMP = ProcessNull.GetString(arParameters(9).Value)
            UseEDCBy = ProcessNull.GetString(arParameters(10).Value)
            Early = ProcessNull.GetInt16(arParameters(11).Value)
            NOMS = ProcessNull.GetInt16(arParameters(12).Value)
            TapeID = ProcessNull.GetString(arParameters(13).Value)
            TapeID = TapeID.Trim()
            EDC = ProcessNull.GetString(arParameters(14).Value)
            Intro = ProcessNull.GetString(arParameters(15).Value)
            Intro = Intro.Trim()
            ProcedureDone = ProcessNull.GetInt16(arParameters(16).Value)
            PN = ProcessNull.GetInt16(arParameters(17).Value)
            GYN = ProcessNull.GetInt16(arParameters(18).Value)
            PatientLast = ProcessNull.GetString(arParameters(19).Value)
            PatientLast = PatientLast.Trim()
            PatientFirst = ProcessNull.GetString(arParameters(20).Value)
            PatientFirst = PatientFirst.Trim()
            RH = ProcessNull.GetString(arParameters(21).Value)
            RH = RH.Trim()
            Type = ProcessNull.GetString(arParameters(22).Value)
            Type = Type.Trim()
            PatientID = ProcessNull.GetInt32(arParameters(23).Value)
            ChartID = ProcessNull.GetInt32(arParameters(24).Value)
            Signed = ProcessNull.GetString(arParameters(25).Value)
            Signed = Signed.Trim()
            ExamDate = ProcessNull.GetString(arParameters(26).Value)
            Adnexa = ProcessNull.GetString(arParameters(27).Value)
            Evaluation = ProcessNull.GetString(arParameters(28).Value)
            Uterus = ProcessNull.GetString(arParameters(29).Value)
            Appearance = ProcessNull.GetString(arParameters(30).Value)
            Skin = ProcessNull.GetString(arParameters(31).Value)
            HEENT = ProcessNull.GetString(arParameters(32).Value)
            Neck = ProcessNull.GetString(arParameters(33).Value)
            Heart = ProcessNull.GetString(arParameters(34).Value)
            Lung = ProcessNull.GetString(arParameters(35).Value)
            Chest = ProcessNull.GetString(arParameters(36).Value)
            Abdomen = ProcessNull.GetString(arParameters(37).Value)
            Back = ProcessNull.GetString(arParameters(38).Value)
            Pelvic = ProcessNull.GetString(arParameters(39).Value)
            Rectal = ProcessNull.GetString(arParameters(40).Value)
            Extremities = ProcessNull.GetString(arParameters(41).Value)
            Neurologic = ProcessNull.GetString(arParameters(42).Value)
            ROS = ProcessNull.GetString(arParameters(43).Value)
            VitalsOutput = ProcessNull.GetString(arParameters(44).Value)
            Complaints = ProcessNull.GetString(arParameters(45).Value)
            Assessment = ProcessNull.GetString(arParameters(46).Value)
            Recommendation = ProcessNull.GetString(arParameters(47).Value)
            LabOutput = ProcessNull.GetString(arParameters(48).Value)
            Singleton = ProcessNull.GetString(arParameters(49).Value)
            PatientHeight = ProcessNull.GetInt32(arParameters(50).Value)
            PreWeight = ProcessNull.GetInt32(arParameters(51).Value)
            ReportTypeID = ProcessNull.GetInt32(arParameters(52).Value)
            BloodType = ProcessNull.GetString(arParameters(53).Value)
            CervixLength = ProcessNull.GetDouble(arParameters(54).Value)
            CervixLength = Math.Round(CervixLength, 2)
            InternalOS = ProcessNull.GetString(arParameters(55).Value)
            FPressure = ProcessNull.GetString(arParameters(56).Value)
            SuppressIntro = ProcessNull.GetInt16(arParameters(57).Value)
            RoomNumber = ProcessNull.GetString(arParameters(58).Value)
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
    '* Description: Updates a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal ID As Integer, _
                            ByVal Service As Integer, _
                            ByVal ExaminerID As Integer, _
                            ByVal SiteID As Integer, _
                            ByVal Examiner2ID As Integer, _
                            ByVal UltrasonographerID As Integer, _
                            ByVal Encounter As String, _
                            ByVal PhysicianID As Integer, _
                            ByVal EarlyUS As String, _
                            ByVal LMP As String, _
                            ByVal UseEDCBy As String, _
                            ByVal Early As Short, _
                            ByVal NOMS As Short, _
                            ByVal TapeID As String, _
                            ByVal EDC As String, _
                            ByVal Intro As String, _
                            ByVal ProcedureDone As Short, _
                            ByVal PN As Short, _
                            ByVal GYN As Short, _
                            ByVal ChartID As Integer, _
                            ByVal Signed As String, _
                            ByVal Indications As String, _
                            ByVal Singleton As String, _
                            ByVal ReportTypeID As Integer, _
                            ByVal UpdatedBy As String, _
                            ByVal SuppressIntro As Short, _
                            ByVal RoomNumber As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(26) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Service", SqlDbType.Int)
        arParameters(1).Value = IIf(Service = 0, System.DBNull.Value, Service)
        arParameters(2) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(2).Value = IIf(ExaminerID = 0, System.DBNull.Value, ExaminerID)
        arParameters(3) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(3).Value = IIf(SiteID = 0, System.DBNull.Value, SiteID)
        arParameters(4) = New SqlParameter("@Examiner2ID", SqlDbType.Int)
        arParameters(4).Value = IIf(Examiner2ID = 0, System.DBNull.Value, Examiner2ID)
        arParameters(5) = New SqlParameter("@UltrasonographerID", SqlDbType.Int)
        arParameters(5).Value = IIf(UltrasonographerID = 0, System.DBNull.Value, UltrasonographerID)
        arParameters(6) = New SqlParameter("@Encounter", SqlDbType.NVarChar, 200)
        arParameters(6).Value = Encounter
        arParameters(7) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(7).Value = IIf(PhysicianID = 0, System.DBNull.Value, PhysicianID)
        arParameters(8) = New SqlParameter("@EarlyUS", SqlDbType.SmallDateTime)
        If Trim(EarlyUS) = "" Then
            arParameters(8).Value = DBNull.Value
        Else
            arParameters(8).Value = EarlyUS
        End If
        arParameters(9) = New SqlParameter("@LMP", SqlDbType.SmallDateTime)
        If Trim(LMP) = "" Then
            arParameters(9).Value = DBNull.Value
        Else
            arParameters(9).Value = LMP
        End If
        arParameters(10) = New SqlParameter("@UseEDCBy", SqlDbType.NVarChar, 50)
        arParameters(10).Value = UseEDCBy
        arParameters(11) = New SqlParameter("@Early", SqlDbType.SmallInt)
        arParameters(11).Value = Early
        arParameters(12) = New SqlParameter("@NOMS", SqlDbType.SmallInt)
        arParameters(12).Value = NOMS
        arParameters(13) = New SqlParameter("@TapeID", SqlDbType.NVarChar, 50)
        arParameters(13).Value = TapeID
        arParameters(14) = New SqlParameter("@EDC", SqlDbType.SmallDateTime)
        If Trim(EDC) = "" Then
            arParameters(14).Value = DBNull.Value
        Else
            arParameters(14).Value = EDC
        End If
        arParameters(15) = New SqlParameter("@Intro", SqlDbType.NVarChar, 50)
        arParameters(15).Value = Intro
        arParameters(16) = New SqlParameter("@ProcedureDone", SqlDbType.SmallInt)
        arParameters(16).Value = ProcedureDone
        arParameters(17) = New SqlParameter("@PN", SqlDbType.SmallInt)
        arParameters(17).Value = PN
        arParameters(18) = New SqlParameter("@Gyn", SqlDbType.SmallInt)
        arParameters(18).Value = GYN
        arParameters(19) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(19).Value = ChartID
        arParameters(20) = New SqlParameter("@Signed", SqlDbType.NVarChar, 50)
        arParameters(20).Value = Signed
        arParameters(21) = New SqlParameter("@Indications", SqlDbType.NVarChar, 255)
        arParameters(21).Value = Indications
        arParameters(22) = New SqlParameter("@Singleton", SqlDbType.VarChar, 8000)
        arParameters(22).Value = Singleton
        arParameters(23) = New SqlParameter("@ReportTypeID", SqlDbType.Int)
        arParameters(23).Value = ReportTypeID
        arParameters(24) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(24).Value = UpdatedBy
        arParameters(25) = New SqlParameter("@SuppressIntro", SqlDbType.SmallInt)
        arParameters(25).Value = SuppressIntro
        arParameters(26) = New SqlParameter("@RoomNumber", SqlDbType.NVarChar, 15)
        arParameters(26).Value = RoomNumber
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamsUpdate", arParameters)
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
    '* Name:        UpdateExamDate
    '*
    '* Description: UpdateExamDates a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdateExamDate(ByVal ID As Integer, _
                           ByVal ExamDate As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@ExamDate", SqlDbType.SmallDateTime)
        If Trim(ExamDate) = "" Then
            arParameters(1).Value = DBNull.Value
        Else
            arParameters(1).Value = ExamDate
        End If
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamDateUpd", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamDateUpd", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not UpdateExamDated.
        If intRecordsAffected = 0 Then
            Return False
        Else
            Return True
        End If

    End Function

    '**************************************************************************
    '*  
    '* Name:        UpdateAdnexa
    '*
    '* Description: Updates a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdateAdnexa(ByVal ID As Integer, _
                            ByVal Adnexa As String, _
                            ByVal Uterus As String, _
                            ByVal Evaluation As String, _
                            ByVal CervixLength As Double, _
                            ByVal InternalOS As String, _
                            ByVal FPressure As String, _
                            ByVal UpdatedBy As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(7) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Adnexa", SqlDbType.VarChar, 8000)
        arParameters(1).Value = Adnexa
        arParameters(2) = New SqlParameter("@Uterus", SqlDbType.VarChar, 8000)
        arParameters(2).Value = Uterus
        arParameters(3) = New SqlParameter("@Evaluation", SqlDbType.VarChar, 8000)
        arParameters(3).Value = Evaluation
        arParameters(4) = New SqlParameter("@CervixLength", SqlDbType.Real)
        arParameters(4).Value = CervixLength
        arParameters(5) = New SqlParameter("@InternalOS", SqlDbType.NVarChar, 100)
        arParameters(5).Value = InternalOS
        arParameters(6) = New SqlParameter("@FPressure", SqlDbType.NVarChar, 100)
        arParameters(6).Value = FPressure
        arParameters(7) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(7).Value = UpdatedBy

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsAdnexaUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamsAdnexaUpdate", arParameters)
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

    End Function 'UpdateAdnexa
    '**************************************************************************
    '*  
    '* Name:        UpdateAssessment
    '*
    '* Description: Updates a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdateAssessment(ByVal ID As Integer, _
                            ByVal Complaints As String, _
                            ByVal Assessment As String, _
                            ByVal Recommendation As String, _
                            ByVal UpdatedBy As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Complaints", SqlDbType.VarChar, 8000)
        arParameters(1).Value = Complaints
        arParameters(2) = New SqlParameter("@Assessment", SqlDbType.VarChar, 8000)
        arParameters(2).Value = Assessment
        arParameters(3) = New SqlParameter("@Recommendation", SqlDbType.VarChar, 8000)
        arParameters(3).Value = Recommendation
        arParameters(4) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(4).Value = UpdatedBy
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsAssessmentUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamsAssessmentUpdate", arParameters)
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

    End Function 'UpdateAssessment
    '**************************************************************************
    '*  
    '* Name:        UpdateMedLab
    '*
    '* Description: Updates a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdateMedLab(ByVal ID As Integer, _
                            ByVal PatientID As Integer, _
                            ByVal RH As String, _
                            ByVal Type As String, _
                            ByVal LabOutput As String, _
                            ByVal UpdatedBy As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(5) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(1).Value = PatientID
        arParameters(2) = New SqlParameter("@RH", SqlDbType.NVarChar, 50)
        arParameters(2).Value = RH
        arParameters(3) = New SqlParameter("@Type", SqlDbType.NVarChar, 50)
        arParameters(3).Value = Type
        arParameters(4) = New SqlParameter("@LabOutput", SqlDbType.VarChar, 8000)
        arParameters(4).Value = LabOutput
        arParameters(5) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(5).Value = UpdatedBy
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsMedLabUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamsMedLabUpdate", arParameters)
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

    End Function 'UpdateMedLab
    '**************************************************************************
    '*  
    '* Name:        UpdateExamination
    '*
    '* Description: Updates a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdateExamination(ByVal ID As Integer, _
                            ByVal Appearance As String, _
                            ByVal Skin As String, _
                            ByVal HEENT As String, _
                            ByVal Neck As String, _
                            ByVal Heart As String, _
                            ByVal Lung As String, _
                            ByVal Chest As String, _
                            ByVal Abdomen As String, _
                            ByVal Back As String, _
                            ByVal Pelvic As String, _
                            ByVal Rectal As String, _
                            ByVal Extremities As String, _
                            ByVal Neurologic As String, _
                            ByVal ROS As String, _
                            ByVal VitalsOutput As String, _
                            ByVal UpdatedBy As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(16) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Appearance", SqlDbType.NVarChar, 100)
        arParameters(1).Value = Appearance
        arParameters(2) = New SqlParameter("@Skin", SqlDbType.NVarChar, 200)
        arParameters(2).Value = Skin
        arParameters(3) = New SqlParameter("@HEENT", SqlDbType.NVarChar, 100)
        arParameters(3).Value = HEENT
        arParameters(4) = New SqlParameter("@Neck", SqlDbType.NVarChar, 100)
        arParameters(4).Value = Neck
        arParameters(5) = New SqlParameter("@Heart", SqlDbType.NVarChar, 100)
        arParameters(5).Value = Heart
        arParameters(6) = New SqlParameter("@Lung", SqlDbType.NVarChar, 100)
        arParameters(6).Value = Lung
        arParameters(7) = New SqlParameter("@Chest", SqlDbType.NVarChar, 100)
        arParameters(7).Value = Chest
        arParameters(8) = New SqlParameter("@Abdomen", SqlDbType.NVarChar, 100)
        arParameters(8).Value = Abdomen
        arParameters(9) = New SqlParameter("@Back", SqlDbType.NVarChar, 100)
        arParameters(9).Value = Back
        arParameters(10) = New SqlParameter("@Pelvic", SqlDbType.NVarChar, 100)
        arParameters(10).Value = Pelvic
        arParameters(11) = New SqlParameter("@Rectal", SqlDbType.NVarChar, 100)
        arParameters(11).Value = Rectal
        arParameters(12) = New SqlParameter("@Extremities", SqlDbType.NVarChar, 100)
        arParameters(12).Value = Extremities
        arParameters(13) = New SqlParameter("@Neurologic", SqlDbType.NVarChar, 100)
        arParameters(13).Value = Neurologic
        arParameters(14) = New SqlParameter("@ROS", SqlDbType.VarChar, 8000)
        arParameters(14).Value = ROS
        arParameters(15) = New SqlParameter("@VitalsOutput", SqlDbType.VarChar, 8000)
        arParameters(15).Value = VitalsOutput
        arParameters(16) = New SqlParameter("@UpdatedBy", SqlDbType.VarChar, 50)
        arParameters(16).Value = UpdatedBy
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsExamUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamsExamUpdate", arParameters)
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

    End Function 'UpdateExamination
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
                        ByVal ExaminerID As Integer, _
                        ByVal SiteID As Integer, _
                        ByVal PhysicianID As Integer, _
                        ByVal EarlyUS As String, _
                        ByVal LMP As String, _
                        ByVal UseEDCBy As String, _
                        ByVal EDC As String, _
                        ByVal ChartID As Integer, _
                        ByVal Indications As String, _
                        ByVal UserID As String, _
                        ByRef AddIndication As Boolean, _
                        ByVal CopyMode As Boolean) As Boolean

        Dim arParameters(12) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(1).Value = ExaminerID
        arParameters(2) = New SqlParameter("@SiteID", SqlDbType.Int)
        arParameters(2).Value = IIf(SiteID = 0, DBNull.Value, SiteID)
        arParameters(3) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        arParameters(3).Value = IIf(PhysicianID = 0, DBNull.Value, PhysicianID)
        arParameters(4) = New SqlParameter("@EarlyUS", SqlDbType.SmallDateTime)
        If Trim(EarlyUS) = "" Then
            arParameters(4).Value = DBNull.Value
        Else
            arParameters(4).Value = EarlyUS
        End If
        arParameters(5) = New SqlParameter("@LMP", SqlDbType.SmallDateTime)
        If Trim(LMP) = "" Then
            arParameters(5).Value = DBNull.Value
        Else
            arParameters(5).Value = LMP
        End If
        arParameters(6) = New SqlParameter("@UseEDCBy", SqlDbType.NVarChar, 50)
        arParameters(6).Value = UseEDCBy
        arParameters(7) = New SqlParameter("@EDC", SqlDbType.SmallDateTime)
        If Trim(EDC) = "" Then
            arParameters(7).Value = DBNull.Value
        Else
            arParameters(7).Value = EDC
        End If
        arParameters(8) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(8).Value = ChartID
        arParameters(9) = New SqlParameter("@Indications", SqlDbType.NVarChar, 255)
        arParameters(9).Value = Indications
        arParameters(10) = New SqlParameter("@UserID", SqlDbType.NVarChar, 50)
        arParameters(10).Value = UserID
        arParameters(11) = New SqlParameter("@AddIndication", SqlDbType.Bit)
        arParameters(11).Direction = ParameterDirection.Output
        arParameters(12) = New SqlParameter("@CopyMode", SqlDbType.Bit)
        arParameters(12).Value = CopyMode
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamsInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            ID = CType(arParameters(0).Value, Integer)
            AddIndication = CType(arParameters(11).Value, Boolean)
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
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamsDelete", arParameters)
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
    '* Name:        UpdFetusCount
    '*
    '* Description: UpdFetusCounts a record from the [PatientInfo] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to UpdFetusCount
    '*
    '* Returns:     Boolean indicating if record was UpdFetusCountd or not. 
    '*              True (record found and UpdFetusCountd); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdFetusCount(ByVal ExamID As Integer, ByVal FetusCount As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ExamID
        arParameters(1) = New SqlParameter("@FetusCount", SqlDbType.Int)
        arParameters(1).Value = FetusCount

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "sp_Exam_OrderUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "sp_Exam_OrderUpdate", arParameters)
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
    '* Name:        GetExamAuditTrail
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '* Parameters: bDefault,Index,bAsc
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetExamAuditTrail(ByVal ExamID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ExamID


        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spExams_AuditTrailGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spExams_AuditTrailGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetExamPrintAuditTrail
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '* Parameters: bDefault,Index,bAsc
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetExamPrintAuditTrail(ByVal ExamID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ExamID


        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spExams_PrintAuditTrailGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spExams_PrintAuditTrailGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetExamFaxAuditTrail
    '*
    '* Description: Returns all records in the [PatientInfo] table according
    '*              to specified criteria.
    '*
    '* Parameters: bDefault,Index,bAsc
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetExamFaxAuditTrail(ByVal ExamID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ExamID


        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spExams_FaxAuditTrailGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spExams_FaxAuditTrailGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        ExamAuditTrail
    '*
    '* Description: ExamAuditTrails a record in the Exam table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was ExamAuditTraild or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function ExamAuditTrail(ByVal ExamID As Integer, _
                           ByVal InUserID As String, _
                           ByVal WorkStation As String, _
                           ByVal bReadOnly As Short, _
                           ByVal Signed As String) As Boolean

        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ExamID
        arParameters(1) = New SqlParameter("@InUserID", SqlDbType.NVarChar, 50)
        arParameters(1).Value = InUserID
        arParameters(2) = New SqlParameter("@ReadOnly", SqlDbType.Bit)
        arParameters(2).Value = bReadOnly
        arParameters(3) = New SqlParameter("@Workstation", SqlDbType.NVarChar, 100)
        arParameters(3).Value = WorkStation
        arParameters(4) = New SqlParameter("@Signed", SqlDbType.Bit)
        If Signed = "Yes" Then
            arParameters(4).Value = 1
        Else
            arParameters(4).Value = 0
        End If

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsOpenAudit", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamsOpenAudit", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not ExamAuditTraild.
        If intRecordsAffected = 0 Then
            Return False
        Else
            Return True
        End If

    End Function

    '**************************************************************************
    '*  
    '* Name:        ExamPrintAuditTrail
    '*
    '* Description: ExamPrintAuditTrails a record in the Exam table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was ExamPrintAuditTraild or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function ExamPrintAuditTrail(ByVal ExamID As Integer, _
                           ByVal InUserID As String, _
                           ByVal WorkStation As String, _
                           ByVal bReadOnly As Short, _
                           ByVal Signed As String) As Boolean

        Dim arParameters(4) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ExamID
        arParameters(1) = New SqlParameter("@InUserID", SqlDbType.NVarChar, 50)
        arParameters(1).Value = InUserID
        arParameters(2) = New SqlParameter("@ReadOnly", SqlDbType.Bit)
        arParameters(2).Value = bReadOnly
        arParameters(3) = New SqlParameter("@Workstation", SqlDbType.NVarChar, 100)
        arParameters(3).Value = WorkStation
        arParameters(4) = New SqlParameter("@Signed", SqlDbType.NVarChar, 255)
        arParameters(4).Value = Signed


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsPrintAudit", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamsPrintAudit", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not ExamPrintAuditTraild.
        If intRecordsAffected = 0 Then
            Return False
        Else
            Return True
        End If

    End Function
    '**************************************************************************
    '*  
    '* Name:        ExamFaxAuditTrail
    '*
    '* Description: ExamFaxAuditTrails a record in the Exam table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was ExamFaxAuditTraild or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function ExamFaxAuditTrail(ByVal ExamID As Integer, _
                           ByVal InUserID As String, _
                           ByVal WorkStation As String, _
                           ByVal bReadOnly As Short, _
                           ByVal Signed As String, _
                           ByVal Recipient As String, _
                           ByVal RecipientFax As String) As Boolean

        Dim arParameters(6) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ExamID
        arParameters(1) = New SqlParameter("@InUserID", SqlDbType.NVarChar, 50)
        arParameters(1).Value = InUserID
        arParameters(2) = New SqlParameter("@ReadOnly", SqlDbType.Bit)
        arParameters(2).Value = bReadOnly
        arParameters(3) = New SqlParameter("@Workstation", SqlDbType.NVarChar, 100)
        arParameters(3).Value = WorkStation
        arParameters(4) = New SqlParameter("@Signed", SqlDbType.NVarChar, 255)
        arParameters(4).Value = Signed
        arParameters(5) = New SqlParameter("@Recipient", SqlDbType.NVarChar, 50)
        arParameters(5).Value = Recipient
        arParameters(6) = New SqlParameter("@RecipientFax", SqlDbType.NVarChar, 50)
        arParameters(6).Value = RecipientFax


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spExamsFaxAudit", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spExamsFaxAudit", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not ExamFaxAuditTraild.
        If intRecordsAffected = 0 Then
            Return False
        Else
            Return True
        End If

    End Function
#End Region


End Class 'dalExams
