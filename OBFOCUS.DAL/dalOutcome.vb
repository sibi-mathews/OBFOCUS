
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalOutcome
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
Public Class dalOutcome

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum OutcomeFields

        fldChartID = 0
        fldPatientLast = 1
        fldPatientFirst = 2
        fldDOB = 3
        fldTrackA = 4
        fldEDC = 5
        fldGravida = 6
        fldPara = 7
        fldLanguage = 8
        fldOutcomeID = 9
        fldLetterSent = 10
        fldDateReviewed = 11
        fldODiagnosis = 12
        fldUltrasound = 13
        fldProcedures = 14
        fldAntepartum = 15
        fldMode = 16
        fldWeightInfantA = 17
        fldApgar1A = 18
        fldApgar1B = 19
        fldApgar1C = 20
        fldApgar5A = 21
        fldApgar5B = 22
        fldApgar5C = 23
        fldWeightInfantB = 24
        fldWeightInfantC = 25
        fldNICU = 26
        fldDateDel = 27
        fldRegNurse = 28
        fldPedNeo = 29
        fldNeoDX = 30
        fldNeoCourse = 31
        fldAddInfo = 32
        fldPatientID = 33
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
        ByRef ChartID As Integer, _
        ByRef PatientFirst As String, _
        ByRef PatientLast As String, _
        ByRef TrackA As Int16, _
        ByRef DOB As Date, _
        ByRef EDC As Date, _
        ByRef Gravida As Integer, _
        ByRef Para As Integer, _
        ByRef Language As String, _
        ByRef OutcomeID As Integer, _
        ByRef LetterSent As Date, _
        ByRef DateReviewed As Date, _
        ByRef ODiagnosis As String, _
        ByRef Ultrasound As String, _
        ByRef Procedures As String, _
        ByRef Antepartum As String, _
        ByRef Mode As String, _
        ByRef NICU As String, _
        ByRef DateDel As Date, _
        ByRef RegNurse As String, _
        ByRef PedNeo As Integer, _
        ByRef NeoDX As String, _
        ByRef NeoCourse As String, _
        ByRef AddInfo As String, _
        ByRef WeightInfantA As Double, _
        ByRef WeightInfantB As Double, _
        ByRef WeightInfantC As Double, _
        ByRef Apgar1A As Integer, _
        ByRef Apgar1B As Integer, _
        ByRef Apgar1C As Integer, _
        ByRef Apgar5A As Integer, _
        ByRef Apgar5B As Integer, _
        ByRef Apgar5C As Integer) As Boolean

        Dim arParameters(32) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        'arParameters(Me.OutcomeFields.fldPatientID) = New SqlParameter("@PatientID", SqlDbType.Int)
        'arParameters(Me.OutcomeFields.fldPatientID).Value = ID
        arParameters(Me.OutcomeFields.fldChartID) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldChartID).Value = ChartID
        arParameters(Me.OutcomeFields.fldPatientLast) = New SqlParameter("@PatientLast", SqlDbType.NVarChar, 50)
        arParameters(Me.OutcomeFields.fldPatientLast).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldPatientFirst) = New SqlParameter("@PatientFirst", SqlDbType.NVarChar, 50)
        arParameters(Me.OutcomeFields.fldPatientFirst).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldDOB) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        arParameters(Me.OutcomeFields.fldDOB).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldTrackA) = New SqlParameter("@TrackA", SqlDbType.SmallInt)
        arParameters(Me.OutcomeFields.fldTrackA).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldEDC) = New SqlParameter("@EDC", SqlDbType.SmallDateTime)
        arParameters(Me.OutcomeFields.fldEDC).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldGravida) = New SqlParameter("@Gravida", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldGravida).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldPara) = New SqlParameter("@Para", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldPara).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldLanguage) = New SqlParameter("@Language", SqlDbType.NVarChar, 50)
        arParameters(Me.OutcomeFields.fldLanguage).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldOutcomeID) = New SqlParameter("@OutcomeID", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldOutcomeID).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldLetterSent) = New SqlParameter("@LetterSent", SqlDbType.SmallDateTime)
        arParameters(Me.OutcomeFields.fldLetterSent).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldDateReviewed) = New SqlParameter("@DateReviewed", SqlDbType.SmallDateTime)
        arParameters(Me.OutcomeFields.fldDateReviewed).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldODiagnosis) = New SqlParameter("@ODiagnosis", SqlDbType.NVarChar, 255)
        arParameters(Me.OutcomeFields.fldODiagnosis).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldUltrasound) = New SqlParameter("@Ultrasound", SqlDbType.VarChar, 8000)
        arParameters(Me.OutcomeFields.fldUltrasound).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldProcedures) = New SqlParameter("@Procedures", SqlDbType.VarChar, 8000)
        arParameters(Me.OutcomeFields.fldProcedures).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldAntepartum) = New SqlParameter("@Antepartum", SqlDbType.NVarChar, 255)
        arParameters(Me.OutcomeFields.fldAntepartum).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldMode) = New SqlParameter("@Mode", SqlDbType.NVarChar, 50)
        arParameters(Me.OutcomeFields.fldMode).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldWeightInfantA) = New SqlParameter("@WeightInfantA", SqlDbType.Real)
        arParameters(Me.OutcomeFields.fldWeightInfantA).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldApgar1A) = New SqlParameter("@Apgar1A", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldApgar1A).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldApgar1B) = New SqlParameter("@Apgar1B", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldApgar1B).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldApgar1C) = New SqlParameter("@Apgar1C", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldApgar1C).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldApgar5A) = New SqlParameter("@Apgar5A", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldApgar5A).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldApgar5B) = New SqlParameter("@Apgar5B", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldApgar5B).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldApgar5C) = New SqlParameter("@Apgar5C", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldApgar5C).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldWeightInfantB) = New SqlParameter("@WeightInfantB", SqlDbType.Real)
        arParameters(Me.OutcomeFields.fldWeightInfantB).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldWeightInfantC) = New SqlParameter("@WeightInfantC", SqlDbType.Real)
        arParameters(Me.OutcomeFields.fldWeightInfantC).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldNICU) = New SqlParameter("@NICU", SqlDbType.NVarChar, 100)
        arParameters(Me.OutcomeFields.fldNICU).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldDateDel) = New SqlParameter("@DateDel", SqlDbType.SmallDateTime)
        arParameters(Me.OutcomeFields.fldDateDel).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldRegNurse) = New SqlParameter("@RegNurse", SqlDbType.NVarChar, 80)
        arParameters(Me.OutcomeFields.fldRegNurse).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldPedNeo) = New SqlParameter("@PedNeo", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldPedNeo).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldNeoDX) = New SqlParameter("@NeoDX", SqlDbType.NVarChar, 255)
        arParameters(Me.OutcomeFields.fldNeoDX).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldNeoCourse) = New SqlParameter("@NeoCourse", SqlDbType.VarChar, 8000)
        arParameters(Me.OutcomeFields.fldNeoCourse).Direction = ParameterDirection.Output
        arParameters(Me.OutcomeFields.fldAddInfo) = New SqlParameter("@Addinfo", SqlDbType.VarChar, 8000)
        arParameters(Me.OutcomeFields.fldAddInfo).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOutcomeGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOutcomeGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.OutcomeFields.fldChartID).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            ChartID = ProcessNull.GetInt32(arParameters(Me.OutcomeFields.fldChartID).Value)
            PatientLast = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldPatientLast).Value)
            PatientLast = PatientLast.Trim()
            PatientFirst = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldPatientFirst).Value)
            PatientFirst = PatientFirst.Trim()
            DOB = ProcessNull.GetDate(arParameters(Me.OutcomeFields.fldDOB).Value)
            EDC = ProcessNull.GetDate(arParameters(Me.OutcomeFields.fldEDC).Value)
            TrackA = ProcessNull.GetInt16(arParameters(Me.OutcomeFields.fldTrackA).Value)
            Gravida = ProcessNull.GetInt32(arParameters(Me.OutcomeFields.fldGravida).Value)
            Para = ProcessNull.GetInt32(arParameters(Me.OutcomeFields.fldPara).Value)
            Language = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldLanguage).Value)
            OutcomeID = ProcessNull.GetInt32(arParameters(Me.OutcomeFields.fldOutcomeID).Value)
            LetterSent = ProcessNull.GetDate(arParameters(Me.OutcomeFields.fldLetterSent).Value)
            DateReviewed = ProcessNull.GetDate(arParameters(Me.OutcomeFields.fldDateReviewed).Value)
            ODiagnosis = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldODiagnosis).Value)
            Ultrasound = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldUltrasound).Value)
            Procedures = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldProcedures).Value)
            Antepartum = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldAntepartum).Value)
            Mode = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldMode).Value)
            NICU = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldNICU).Value)
            DateDel = ProcessNull.GetDate(arParameters(Me.OutcomeFields.fldDateDel).Value)
            RegNurse = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldRegNurse).Value)
            PedNeo = ProcessNull.GetInt32(arParameters(Me.OutcomeFields.fldPedNeo).Value)
            NeoDX = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldNeoDX).Value)
            NeoCourse = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldNeoCourse).Value)
            AddInfo = ProcessNull.GetString(arParameters(Me.OutcomeFields.fldAddInfo).Value)
            Apgar1A = ProcessNull.GetInt32(arParameters(Me.OutcomeFields.fldApgar1A).Value)
            Apgar1B = ProcessNull.GetInt32(arParameters(Me.OutcomeFields.fldApgar1B).Value)
            Apgar1C = ProcessNull.GetInt32(arParameters(Me.OutcomeFields.fldApgar1C).Value)
            Apgar5A = ProcessNull.GetInt32(arParameters(Me.OutcomeFields.fldApgar5A).Value)
            Apgar5B = ProcessNull.GetInt32(arParameters(Me.OutcomeFields.fldApgar5B).Value)
            Apgar5C = ProcessNull.GetInt32(arParameters(Me.OutcomeFields.fldApgar5C).Value)
            WeightInfantA = ProcessNull.GetDouble(arParameters(Me.OutcomeFields.fldWeightInfantA).Value)
            WeightInfantB = ProcessNull.GetDouble(arParameters(Me.OutcomeFields.fldWeightInfantB).Value)
            WeightInfantC = ProcessNull.GetDouble(arParameters(Me.OutcomeFields.fldWeightInfantC).Value)
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
        ByVal ChartID As Integer, _
        ByVal PatientFirst As String, _
        ByVal PatientLast As String, _
        ByVal TrackA As Int16, _
        ByVal DOB As Date, _
        ByVal EDC As Date, _
        ByVal Gravida As Integer, _
        ByVal Para As Integer, _
        ByVal Language As String, _
        ByVal OutcomeID As Integer, _
        ByVal LetterSent As Date, _
        ByVal DateReviewed As Date, _
        ByVal ODiagnosis As String, _
        ByVal Ultrasound As String, _
        ByVal Procedures As String, _
        ByVal Antepartum As String, _
        ByVal Mode As String, _
        ByVal NICU As String, _
        ByVal DateDel As Date, _
        ByVal RegNurse As String, _
        ByVal PedNeo As Integer, _
        ByVal NeoDX As String, _
        ByVal NeoCourse As String, _
        ByVal AddInfo As String, _
        ByVal WeightInfantA As Double, _
        ByVal WeightInfantB As Double, _
        ByVal WeightInfantC As Double, _
        ByVal Apgar1A As Integer, _
        ByVal Apgar1B As Integer, _
        ByVal Apgar1C As Integer, _
        ByVal Apgar5A As Integer, _
        ByVal Apgar5B As Integer, _
        ByVal Apgar5C As Integer) As Boolean

        Dim arParameters(33) As SqlParameter
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(Me.OutcomeFields.fldChartID) = New SqlParameter("@ChartID", SqlDbType.Int)
        If ChartID = Nothing Then
            arParameters(Me.OutcomeFields.fldChartID).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldChartID).Value = ChartID
        End If
        arParameters(Me.OutcomeFields.fldPatientLast) = New SqlParameter("@PatientLast", SqlDbType.NVarChar, 50)
        arParameters(Me.OutcomeFields.fldPatientLast).Value = PatientLast
        arParameters(Me.OutcomeFields.fldPatientFirst) = New SqlParameter("@PatientFirst", SqlDbType.NVarChar, 50)
        arParameters(Me.OutcomeFields.fldPatientFirst).Value = PatientFirst
        arParameters(Me.OutcomeFields.fldDOB) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        If DOB = Nothing Then
            arParameters(Me.OutcomeFields.fldDOB).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldDOB).Value = DOB
        End If
        arParameters(Me.OutcomeFields.fldTrackA) = New SqlParameter("@TrackA", SqlDbType.SmallInt)
        If TrackA = Nothing Then
            arParameters(Me.OutcomeFields.fldTrackA).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldTrackA).Value = TrackA
        End If
        arParameters(Me.OutcomeFields.fldEDC) = New SqlParameter("@EDC", SqlDbType.SmallDateTime)
        If EDC = Nothing Then
            arParameters(Me.OutcomeFields.fldEDC).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldEDC).Value = EDC
        End If
        arParameters(Me.OutcomeFields.fldGravida) = New SqlParameter("@Gravida", SqlDbType.Int)
        If Gravida = Nothing Then
            arParameters(Me.OutcomeFields.fldGravida).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldGravida).Value = Gravida
        End If
        arParameters(Me.OutcomeFields.fldPara) = New SqlParameter("@Para", SqlDbType.Int)
        If Para = Nothing Then
            arParameters(Me.OutcomeFields.fldPara).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldPara).Value = Para
        End If
        arParameters(Me.OutcomeFields.fldLanguage) = New SqlParameter("@Language", SqlDbType.NVarChar, 50)
        arParameters(Me.OutcomeFields.fldLanguage).Value = Language
        arParameters(Me.OutcomeFields.fldOutcomeID) = New SqlParameter("@OutcomeID", SqlDbType.Int)
        If OutcomeID = Nothing Then
            arParameters(Me.OutcomeFields.fldOutcomeID).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldOutcomeID).Value = OutcomeID
        End If
        arParameters(Me.OutcomeFields.fldLetterSent) = New SqlParameter("@LetterSent", SqlDbType.SmallDateTime)
        If LetterSent = Nothing Then
            arParameters(Me.OutcomeFields.fldLetterSent).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldLetterSent).Value = LetterSent
        End If
        arParameters(Me.OutcomeFields.fldDateReviewed) = New SqlParameter("@DateReviewed", SqlDbType.SmallDateTime)
        If DateReviewed = Nothing Then
            arParameters(Me.OutcomeFields.fldDateReviewed).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldDateReviewed).Value = DateReviewed
        End If
        arParameters(Me.OutcomeFields.fldODiagnosis) = New SqlParameter("@ODiagnosis", SqlDbType.NVarChar, 255)
        arParameters(Me.OutcomeFields.fldODiagnosis).Value = ODiagnosis
        arParameters(Me.OutcomeFields.fldUltrasound) = New SqlParameter("@Ultrasound", SqlDbType.VarChar, 8000)
        arParameters(Me.OutcomeFields.fldUltrasound).Value = Ultrasound
        arParameters(Me.OutcomeFields.fldProcedures) = New SqlParameter("@Procedures", SqlDbType.VarChar, 8000)
        arParameters(Me.OutcomeFields.fldProcedures).Value = Procedures
        arParameters(Me.OutcomeFields.fldAntepartum) = New SqlParameter("@Antepartum", SqlDbType.NVarChar, 255)
        arParameters(Me.OutcomeFields.fldAntepartum).Value = Antepartum
        arParameters(Me.OutcomeFields.fldMode) = New SqlParameter("@Mode", SqlDbType.NVarChar, 50)
        arParameters(Me.OutcomeFields.fldMode).Value = Mode
        arParameters(Me.OutcomeFields.fldWeightInfantA) = New SqlParameter("@WeightInfantA", SqlDbType.Real)
        If WeightInfantA = Nothing Then
            arParameters(Me.OutcomeFields.fldWeightInfantA).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldWeightInfantA).Value = WeightInfantA
        End If
        arParameters(Me.OutcomeFields.fldApgar1A) = New SqlParameter("@Apgar1A", SqlDbType.Int)
        If Apgar1A = Nothing Then
            arParameters(Me.OutcomeFields.fldApgar1A).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldApgar1A).Value = Apgar1A
        End If
        arParameters(Me.OutcomeFields.fldApgar1B) = New SqlParameter("@Apgar1B", SqlDbType.Int)
        If Apgar1B = Nothing Then
            arParameters(Me.OutcomeFields.fldApgar1B).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldApgar1B).Value = Apgar1B
        End If
        arParameters(Me.OutcomeFields.fldApgar1C) = New SqlParameter("@Apgar1C", SqlDbType.Int)
        If Apgar1C = Nothing Then
            arParameters(Me.OutcomeFields.fldApgar1C).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldApgar1C).Value = Apgar1C
        End If
        arParameters(Me.OutcomeFields.fldApgar5A) = New SqlParameter("@Apgar5A", SqlDbType.Int)
        If Apgar5A = Nothing Then
            arParameters(Me.OutcomeFields.fldApgar5A).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldApgar5A).Value = Apgar5A
        End If
        arParameters(Me.OutcomeFields.fldApgar5B) = New SqlParameter("@Apgar5B", SqlDbType.Int)
        If Apgar5B = Nothing Then
            arParameters(Me.OutcomeFields.fldApgar5B).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldApgar5B).Value = Apgar5B
        End If
        arParameters(Me.OutcomeFields.fldApgar5C) = New SqlParameter("@Apgar5C", SqlDbType.Int)
        If Apgar5C = Nothing Then
            arParameters(Me.OutcomeFields.fldApgar5C).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldApgar5C).Value = Apgar5C
        End If
        arParameters(Me.OutcomeFields.fldWeightInfantB) = New SqlParameter("@WeightInfantB", SqlDbType.Real)
        If WeightInfantB = Nothing Then
            arParameters(Me.OutcomeFields.fldWeightInfantB).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldWeightInfantB).Value = WeightInfantB
        End If
        arParameters(Me.OutcomeFields.fldWeightInfantC) = New SqlParameter("@WeightInfantC", SqlDbType.Real)
        If WeightInfantC = Nothing Then
            arParameters(Me.OutcomeFields.fldWeightInfantC).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldWeightInfantC).Value = WeightInfantC
        End If
        arParameters(Me.OutcomeFields.fldNICU) = New SqlParameter("@NICU", SqlDbType.NVarChar, 100)
        arParameters(Me.OutcomeFields.fldNICU).Value = NICU
        arParameters(Me.OutcomeFields.fldDateDel) = New SqlParameter("@DateDel", SqlDbType.SmallDateTime)
        If DateDel = Nothing Then
            arParameters(Me.OutcomeFields.fldDateDel).Value = DBNull.Value
        Else
            arParameters(Me.OutcomeFields.fldDateDel).Value = DateDel
        End If
        arParameters(Me.OutcomeFields.fldRegNurse) = New SqlParameter("@RegNurse", SqlDbType.NVarChar, 80)
        arParameters(Me.OutcomeFields.fldRegNurse).Value = RegNurse
        arParameters(Me.OutcomeFields.fldPedNeo) = New SqlParameter("@PedNeo", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldPedNeo).Value = PedNeo
        arParameters(Me.OutcomeFields.fldNeoDX) = New SqlParameter("@NeoDX", SqlDbType.NVarChar, 255)
        arParameters(Me.OutcomeFields.fldNeoDX).Value = NeoDX
        arParameters(Me.OutcomeFields.fldNeoCourse) = New SqlParameter("@NeoCourse", SqlDbType.VarChar, 8000)
        arParameters(Me.OutcomeFields.fldNeoCourse).Value = NeoCourse
        arParameters(Me.OutcomeFields.fldAddInfo) = New SqlParameter("@Addinfo", SqlDbType.VarChar, 8000)
        arParameters(Me.OutcomeFields.fldAddInfo).Value = AddInfo
        arParameters(Me.OutcomeFields.fldPatientID) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(Me.OutcomeFields.fldPatientID).Value = ID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOutcomeUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOutcomeUpdate", arParameters)
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


End Class 'dalOutcome
