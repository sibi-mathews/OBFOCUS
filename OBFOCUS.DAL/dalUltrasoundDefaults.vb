
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalUltrasoundDefaults
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
Public Class dalUltrasoundDefaults

#Region "Module level variables and enums"

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



#Region "Main procedures - GetUltrasoundDefaults, Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Name:        GetUltrasoundDefaults
    '*
    '* Description: Returns all records in the [ZTables] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetUltrasoundDefaults(ByVal ExaminerID As Integer) As SqlDataReader

        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(0).Value = ExaminerID
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasoundDefaultsGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spUltrasoundDefaultsGet", arParameters)
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
    Public Function GetByKey(ByVal ID As Integer, _
                            ByRef Cardiaccomments As String, _
                            ByRef Situscomments As String, _
                            ByRef Chamber4Comments As String, _
                            ByRef Chamber5Comments As String, _
                            ByRef Aorticarchcomments As String, _
                            ByRef Pulmonaryartcomments As String, _
                            ByRef Spinecomments As String, _
                            ByRef LKidneycomments As String, _
                            ByRef RKidneycomments As String, _
                            ByRef Bladdercomments As String, _
                            ByRef Stomachcomments As String, _
                            ByRef Diaphragmcomments As String, _
                            ByRef UExtremitiescomments As String, _
                            ByRef LExtremitiescomments As String, _
                            ByRef Movementcomments As String, _
                            ByRef Cordinsertioncomments As String, _
                            ByRef Vessel3Comments As String, _
                            ByRef Facecomments As String, _
                            ByRef Intracranialcomments As String, _
                            ByRef Cerebellumcomments As String, _
                            ByRef Latventcomments As String, _
                            ByRef Cisternacomments As String, _
                            ByRef Nuchalcomments As String, _
                            ByRef PLACENTA As String, _
                            ByRef Previa As String, _
                            ByRef Gender As String, _
                            ByRef GSac As String, _
                            ByRef FetalPole As String, _
                            ByRef Ysac As String, _
                            ByRef AFI As String, _
                            ByRef Adnexa As String, _
                            ByRef Uterus As String, _
                            ByRef Procedures As String, _
                            ByRef Ultrasound As Short, _
                            ByRef GeneticCounsel As Short, _
                            ByRef Amniocentesis As Short, _
                            ByRef ExAFP As Short, _
                            ByRef Maternalchrome As Short, _
                            ByRef OtherProcedure As Short, _
                            ByRef Gauge As Integer, _
                            ByRef AFRemoved As Integer, _
                            ByRef NInsertions As Integer, _
                            ByRef AFColor As String, _
                            ByRef Transplacental As String, _
                            ByRef Complications As String, _
                            ByRef LAmniocentesis As Short, _
                            ByRef Rhogam As String, _
                            ByRef Cardiac As Short, _
                            ByRef Situs As Short, _
                            ByRef Chamber4 As Short, _
                            ByRef Chamber5 As Short, _
                            ByRef Aorticarch As Short, _
                            ByRef Pulmonaryart As Short, _
                            ByRef Spine As Short, _
                            ByRef LKidney As Short, _
                            ByRef RKidney As Short, _
                            ByRef Bladder As Short, _
                            ByRef Stomach As Short, _
                            ByRef Diaphragm As Short, _
                            ByRef UExtremities As Short, _
                            ByRef LExtremities As Short, _
                            ByRef Movement As Short, _
                            ByRef Cordinsertion As Short, _
                            ByRef Vessel3 As Short, _
                            ByRef Face As Short, _
                            ByRef Intracranial As Short, _
                            ByRef Cerebellum As Short, _
                            ByRef Latvent As Short, _
                            ByRef Cisterna As Short, _
                            ByRef Nuchal As Short, _
                            ByRef DefaultName As String, _
                            ByRef ExaminerID As Integer) As Boolean

        Dim arParameters(71) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@USDefaultID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Cardiaccomments", SqlDbType.NVarChar, 50)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@Situscomments", SqlDbType.NVarChar, 50)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@Chamber4Comments", SqlDbType.NVarChar, 50)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@Chamber5Comments", SqlDbType.NVarChar, 50)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@Aorticarchcomments", SqlDbType.NVarChar, 50)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@Pulmonaryartcomments", SqlDbType.NVarChar, 50)
        arParameters(6).Direction = ParameterDirection.Output
        arParameters(7) = New SqlParameter("@Spinecomments", SqlDbType.NVarChar, 50)
        arParameters(7).Direction = ParameterDirection.Output
        arParameters(8) = New SqlParameter("@LKidneycomments", SqlDbType.NVarChar, 50)
        arParameters(8).Direction = ParameterDirection.Output
        arParameters(9) = New SqlParameter("@RKidneycomments", SqlDbType.NVarChar, 50)
        arParameters(9).Direction = ParameterDirection.Output
        arParameters(10) = New SqlParameter("@Bladdercomments", SqlDbType.NVarChar, 50)
        arParameters(10).Direction = ParameterDirection.Output
        arParameters(11) = New SqlParameter("@Stomachcomments", SqlDbType.NVarChar, 50)
        arParameters(11).Direction = ParameterDirection.Output
        arParameters(12) = New SqlParameter("@UExtremitiescomments", SqlDbType.NVarChar, 50)
        arParameters(12).Direction = ParameterDirection.Output
        arParameters(13) = New SqlParameter("@LExtremitiescomments", SqlDbType.NVarChar, 50)
        arParameters(13).Direction = ParameterDirection.Output
        arParameters(14) = New SqlParameter("@Movementcomments", SqlDbType.NVarChar, 50)
        arParameters(14).Direction = ParameterDirection.Output
        arParameters(15) = New SqlParameter("@Cordinsertioncomments", SqlDbType.NVarChar, 50)
        arParameters(15).Direction = ParameterDirection.Output
        arParameters(16) = New SqlParameter("@Vessel3Comments", SqlDbType.NVarChar, 50)
        arParameters(16).Direction = ParameterDirection.Output
        arParameters(17) = New SqlParameter("@Facecomments", SqlDbType.NVarChar, 50)
        arParameters(17).Direction = ParameterDirection.Output
        arParameters(18) = New SqlParameter("@Intracranialcomments", SqlDbType.NVarChar, 50)
        arParameters(18).Direction = ParameterDirection.Output
        arParameters(19) = New SqlParameter("@Cerebellumcomments", SqlDbType.NVarChar, 50)
        arParameters(19).Direction = ParameterDirection.Output
        arParameters(20) = New SqlParameter("@Latventcomments", SqlDbType.NVarChar, 50)
        arParameters(20).Direction = ParameterDirection.Output
        arParameters(21) = New SqlParameter("@Cisternacomments", SqlDbType.NVarChar, 50)
        arParameters(21).Direction = ParameterDirection.Output
        arParameters(22) = New SqlParameter("@Nuchalcomments", SqlDbType.NVarChar, 50)
        arParameters(22).Direction = ParameterDirection.Output
        arParameters(23) = New SqlParameter("@PLACENTA", SqlDbType.NVarChar, 50)
        arParameters(23).Direction = ParameterDirection.Output
        arParameters(24) = New SqlParameter("@Previa", SqlDbType.NVarChar, 50)
        arParameters(24).Direction = ParameterDirection.Output
        arParameters(25) = New SqlParameter("@GSac", SqlDbType.NVarChar, 50)
        arParameters(25).Direction = ParameterDirection.Output
        arParameters(26) = New SqlParameter("@FetalPole", SqlDbType.NVarChar, 50)
        arParameters(26).Direction = ParameterDirection.Output
        arParameters(27) = New SqlParameter("@Ysac", SqlDbType.NVarChar, 50)
        arParameters(27).Direction = ParameterDirection.Output
        arParameters(28) = New SqlParameter("@AFI", SqlDbType.NVarChar, 50)
        arParameters(28).Direction = ParameterDirection.Output
        arParameters(29) = New SqlParameter("@Adnexa", SqlDbType.NVarChar, 255)
        arParameters(29).Direction = ParameterDirection.Output
        arParameters(30) = New SqlParameter("@Uterus", SqlDbType.NVarChar, 255)
        arParameters(30).Direction = ParameterDirection.Output
        arParameters(31) = New SqlParameter("@Procedures", SqlDbType.NVarChar, 4000)
        arParameters(31).Direction = ParameterDirection.Output
        arParameters(32) = New SqlParameter("@Ultrasound", SqlDbType.SmallInt)
        arParameters(32).Direction = ParameterDirection.Output
        arParameters(33) = New SqlParameter("@GeneticCounsel", SqlDbType.SmallInt)
        arParameters(33).Direction = ParameterDirection.Output
        arParameters(34) = New SqlParameter("@Amniocentesis", SqlDbType.SmallInt)
        arParameters(34).Direction = ParameterDirection.Output
        arParameters(35) = New SqlParameter("@ExAFP", SqlDbType.SmallInt)
        arParameters(35).Direction = ParameterDirection.Output
        arParameters(36) = New SqlParameter("@Maternalchrome", SqlDbType.SmallInt)
        arParameters(36).Direction = ParameterDirection.Output
        arParameters(37) = New SqlParameter("@OtherProcedure", SqlDbType.SmallInt)
        arParameters(37).Direction = ParameterDirection.Output
        arParameters(38) = New SqlParameter("@Gauge", SqlDbType.Int)
        arParameters(38).Direction = ParameterDirection.Output
        arParameters(39) = New SqlParameter("@AFRemoved", SqlDbType.Int)
        arParameters(39).Direction = ParameterDirection.Output
        arParameters(40) = New SqlParameter("@NInsertions", SqlDbType.Int)
        arParameters(40).Direction = ParameterDirection.Output
        arParameters(41) = New SqlParameter("@AFColor", SqlDbType.NVarChar, 50)
        arParameters(41).Direction = ParameterDirection.Output
        arParameters(42) = New SqlParameter("@Transplacental", SqlDbType.NVarChar, 50)
        arParameters(42).Direction = ParameterDirection.Output
        arParameters(43) = New SqlParameter("@Complications", SqlDbType.NVarChar, 50)
        arParameters(43).Direction = ParameterDirection.Output
        arParameters(44) = New SqlParameter("@LAmniocentesis", SqlDbType.SmallInt)
        arParameters(44).Direction = ParameterDirection.Output
        arParameters(45) = New SqlParameter("@Rhogam", SqlDbType.NVarChar, 50)
        arParameters(45).Direction = ParameterDirection.Output
        arParameters(46) = New SqlParameter("@Cardiac", SqlDbType.SmallInt)
        arParameters(46).Direction = ParameterDirection.Output
        arParameters(47) = New SqlParameter("@Situs", SqlDbType.SmallInt)
        arParameters(47).Direction = ParameterDirection.Output
        arParameters(48) = New SqlParameter("@Chamber4", SqlDbType.SmallInt)
        arParameters(48).Direction = ParameterDirection.Output
        arParameters(49) = New SqlParameter("@Chamber5", SqlDbType.SmallInt)
        arParameters(49).Direction = ParameterDirection.Output
        arParameters(50) = New SqlParameter("@Aorticarch", SqlDbType.SmallInt)
        arParameters(50).Direction = ParameterDirection.Output
        arParameters(51) = New SqlParameter("@Pulmonaryart", SqlDbType.SmallInt)
        arParameters(51).Direction = ParameterDirection.Output
        arParameters(52) = New SqlParameter("@Spine", SqlDbType.SmallInt)
        arParameters(52).Direction = ParameterDirection.Output
        arParameters(53) = New SqlParameter("@LKidney", SqlDbType.SmallInt)
        arParameters(53).Direction = ParameterDirection.Output
        arParameters(54) = New SqlParameter("@RKidney", SqlDbType.SmallInt)
        arParameters(54).Direction = ParameterDirection.Output
        arParameters(55) = New SqlParameter("@Bladder", SqlDbType.SmallInt)
        arParameters(55).Direction = ParameterDirection.Output
        arParameters(56) = New SqlParameter("@Stomach", SqlDbType.SmallInt)
        arParameters(56).Direction = ParameterDirection.Output
        arParameters(57) = New SqlParameter("@Diaphragm", SqlDbType.SmallInt)
        arParameters(57).Direction = ParameterDirection.Output
        arParameters(58) = New SqlParameter("@UExtremities", SqlDbType.SmallInt)
        arParameters(58).Direction = ParameterDirection.Output
        arParameters(59) = New SqlParameter("@LExtremities", SqlDbType.SmallInt)
        arParameters(59).Direction = ParameterDirection.Output
        arParameters(60) = New SqlParameter("@Movement", SqlDbType.SmallInt)
        arParameters(60).Direction = ParameterDirection.Output
        arParameters(61) = New SqlParameter("@Cordinsertion", SqlDbType.SmallInt)
        arParameters(61).Direction = ParameterDirection.Output
        arParameters(62) = New SqlParameter("@Vessel3", SqlDbType.SmallInt)
        arParameters(62).Direction = ParameterDirection.Output
        arParameters(63) = New SqlParameter("@Face", SqlDbType.SmallInt)
        arParameters(63).Direction = ParameterDirection.Output
        arParameters(64) = New SqlParameter("@Intracranial", SqlDbType.SmallInt)
        arParameters(64).Direction = ParameterDirection.Output
        arParameters(65) = New SqlParameter("@Cerebellum", SqlDbType.SmallInt)
        arParameters(65).Direction = ParameterDirection.Output
        arParameters(66) = New SqlParameter("@Latvent", SqlDbType.SmallInt)
        arParameters(66).Direction = ParameterDirection.Output
        arParameters(67) = New SqlParameter("@Cisterna", SqlDbType.SmallInt)
        arParameters(67).Direction = ParameterDirection.Output
        arParameters(68) = New SqlParameter("@Nuchal", SqlDbType.SmallInt)
        arParameters(68).Direction = ParameterDirection.Output
        arParameters(69) = New SqlParameter("@USDefaultName", SqlDbType.NVarChar, 50)
        arParameters(69).Direction = ParameterDirection.Output
        arParameters(70) = New SqlParameter("@DiaphragmComments", SqlDbType.NVarChar, 50)
        arParameters(70).Direction = ParameterDirection.Output
        arParameters(71) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(71).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasoundDefaultsGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUltrasoundDefaultsGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(1).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Cardiaccomments = ProcessNull.GetString(arParameters(1).Value)
            Situscomments = ProcessNull.GetString(arParameters(2).Value)
            Chamber4Comments = ProcessNull.GetString(arParameters(3).Value)
            Chamber5Comments = ProcessNull.GetString(arParameters(4).Value)
            Aorticarchcomments = ProcessNull.GetString(arParameters(5).Value)
            Pulmonaryartcomments = ProcessNull.GetString(arParameters(6).Value)
            Spinecomments = ProcessNull.GetString(arParameters(7).Value)
            LKidneycomments = ProcessNull.GetString(arParameters(8).Value)
            RKidneycomments = ProcessNull.GetString(arParameters(9).Value)
            Bladdercomments = ProcessNull.GetString(arParameters(10).Value)
            Stomachcomments = ProcessNull.GetString(arParameters(11).Value)
            UExtremitiescomments = ProcessNull.GetString(arParameters(12).Value)
            LExtremitiescomments = ProcessNull.GetString(arParameters(13).Value)
            Movementcomments = ProcessNull.GetString(arParameters(14).Value)
            Cordinsertioncomments = ProcessNull.GetString(arParameters(15).Value)
            Vessel3Comments = ProcessNull.GetString(arParameters(16).Value)
            Facecomments = ProcessNull.GetString(arParameters(17).Value)
            Intracranialcomments = ProcessNull.GetString(arParameters(18).Value)
            Cerebellumcomments = ProcessNull.GetString(arParameters(19).Value)
            Latventcomments = ProcessNull.GetString(arParameters(20).Value)
            Cisternacomments = ProcessNull.GetString(arParameters(21).Value)
            Nuchalcomments = ProcessNull.GetString(arParameters(22).Value)
            PLACENTA = ProcessNull.GetString(arParameters(23).Value)
            Previa = ProcessNull.GetString(arParameters(24).Value)
            GSac = ProcessNull.GetString(arParameters(25).Value)
            FetalPole = ProcessNull.GetString(arParameters(26).Value)
            Ysac = ProcessNull.GetString(arParameters(27).Value)
            AFI = ProcessNull.GetString(arParameters(28).Value)
            Adnexa = ProcessNull.GetString(arParameters(29).Value)
            Uterus = ProcessNull.GetString(arParameters(30).Value)
            Procedures = ProcessNull.GetString(arParameters(31).Value)
            Ultrasound = ProcessNull.GetInt16(arParameters(32).Value)
            GeneticCounsel = ProcessNull.GetInt16(arParameters(33).Value)
            Amniocentesis = ProcessNull.GetInt16(arParameters(34).Value)
            ExAFP = ProcessNull.GetInt16(arParameters(35).Value)
            Maternalchrome = ProcessNull.GetInt16(arParameters(36).Value)
            OtherProcedure = ProcessNull.GetInt16(arParameters(37).Value)
            Gauge = ProcessNull.GetInt32(arParameters(38).Value)
            AFRemoved = ProcessNull.GetInt32(arParameters(39).Value)
            NInsertions = ProcessNull.GetInt32(arParameters(40).Value)
            AFColor = ProcessNull.GetString(arParameters(41).Value)
            Transplacental = ProcessNull.GetString(arParameters(42).Value)
            Complications = ProcessNull.GetString(arParameters(43).Value)
            LAmniocentesis = ProcessNull.GetInt16(arParameters(44).Value)
            Rhogam = ProcessNull.GetString(arParameters(45).Value)
            Cardiac = ProcessNull.GetInt16(arParameters(46).Value)
            Situs = ProcessNull.GetInt16(arParameters(47).Value)
            Chamber4 = ProcessNull.GetInt16(arParameters(48).Value)
            Chamber5 = ProcessNull.GetInt16(arParameters(49).Value)
            Aorticarch = ProcessNull.GetInt16(arParameters(50).Value)
            Pulmonaryart = ProcessNull.GetInt16(arParameters(51).Value)
            Spine = ProcessNull.GetInt16(arParameters(52).Value)
            LKidney = ProcessNull.GetInt16(arParameters(53).Value)
            RKidney = ProcessNull.GetInt16(arParameters(54).Value)
            Bladder = ProcessNull.GetInt16(arParameters(55).Value)
            Stomach = ProcessNull.GetInt16(arParameters(56).Value)
            Diaphragm = ProcessNull.GetInt16(arParameters(57).Value)
            UExtremities = ProcessNull.GetInt16(arParameters(58).Value)
            LExtremities = ProcessNull.GetInt16(arParameters(59).Value)
            Movement = ProcessNull.GetInt16(arParameters(60).Value)
            Cordinsertion = ProcessNull.GetInt16(arParameters(61).Value)
            Vessel3 = ProcessNull.GetInt16(arParameters(62).Value)
            Face = ProcessNull.GetInt16(arParameters(63).Value)
            Intracranial = ProcessNull.GetInt16(arParameters(64).Value)
            Cerebellum = ProcessNull.GetInt16(arParameters(65).Value)
            Latvent = ProcessNull.GetInt16(arParameters(66).Value)
            Cisterna = ProcessNull.GetInt16(arParameters(67).Value)
            Nuchal = ProcessNull.GetInt16(arParameters(68).Value)
            DefaultName = ProcessNull.GetString(arParameters(69).Value)
            Diaphragmcomments = ProcessNull.GetString(arParameters(70).Value)
            ExaminerID = ProcessNull.GetInt32(arParameters(71).Value)
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
    '* Description: Gets all the values of a record identified by a key.
    '*
    '* Parameters:  Description - Output parameter
    '*              Picture - Output parameter
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal ID As Integer, _
                            ByVal Cardiaccomments As String, _
                            ByVal Situscomments As String, _
                            ByVal Chamber4Comments As String, _
                            ByVal Chamber5Comments As String, _
                            ByVal Aorticarchcomments As String, _
                            ByVal Pulmonaryartcomments As String, _
                            ByVal Spinecomments As String, _
                            ByVal LKidneycomments As String, _
                            ByVal RKidneycomments As String, _
                            ByVal Bladdercomments As String, _
                            ByVal Stomachcomments As String, _
                            ByVal Diaphragmcomments As String, _
                            ByVal UExtremitiescomments As String, _
                            ByVal LExtremitiescomments As String, _
                            ByVal Movementcomments As String, _
                            ByVal Cordinsertioncomments As String, _
                            ByVal Vessel3Comments As String, _
                            ByVal Facecomments As String, _
                            ByVal Intracranialcomments As String, _
                            ByVal Cerebellumcomments As String, _
                            ByVal Latventcomments As String, _
                            ByVal Cisternacomments As String, _
                            ByVal Nuchalcomments As String, _
                            ByVal PLACENTA As String, _
                            ByVal Previa As String, _
                            ByVal Gender As String, _
                            ByVal GSac As String, _
                            ByVal FetalPole As String, _
                            ByVal Ysac As String, _
                            ByVal AFI As String, _
                            ByVal Adnexa As String, _
                            ByVal Uterus As String, _
                            ByVal Procedures As String, _
                            ByVal Ultrasound As Short, _
                            ByVal GeneticCounsel As Short, _
                            ByVal Amniocentesis As Short, _
                            ByVal ExAFP As Short, _
                            ByVal Maternalchrome As Short, _
                            ByVal OtherProcedure As Short, _
                            ByVal Gauge As Integer, _
                            ByVal AFRemoved As Integer, _
                            ByVal NInsertions As Integer, _
                            ByVal AFColor As String, _
                            ByVal Transplacental As String, _
                            ByVal Complications As String, _
                            ByVal LAmniocentesis As Short, _
                            ByVal Rhogam As String, _
                            ByVal Cardiac As Short, _
                            ByVal Situs As Short, _
                            ByVal Chamber4 As Short, _
                            ByVal Chamber5 As Short, _
                            ByVal Aorticarch As Short, _
                            ByVal Pulmonaryart As Short, _
                            ByVal Spine As Short, _
                            ByVal LKidney As Short, _
                            ByVal RKidney As Short, _
                            ByVal Bladder As Short, _
                            ByVal Stomach As Short, _
                            ByVal Diaphragm As Short, _
                            ByVal UExtremities As Short, _
                            ByVal LExtremities As Short, _
                            ByVal Movement As Short, _
                            ByVal Cordinsertion As Short, _
                            ByVal Vessel3 As Short, _
                            ByVal Face As Short, _
                            ByVal Intracranial As Short, _
                            ByVal Cerebellum As Short, _
                            ByVal Latvent As Short, _
                            ByVal Cisterna As Short, _
                            ByVal Nuchal As Short, _
                            ByVal DefaultName As String, _
                            ByVal ExaminerID As Integer) As Boolean
        Dim arParameters(71) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@USDefaultID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Cardiaccomments", SqlDbType.NVarChar, 50)
        arParameters(1).Value = Cardiaccomments
        arParameters(2) = New SqlParameter("@Situscomments", SqlDbType.NVarChar, 50)
        arParameters(2).Value = Situscomments
        arParameters(3) = New SqlParameter("@Chamber4Comments", SqlDbType.NVarChar, 50)
        arParameters(3).Value = Chamber4Comments
        arParameters(4) = New SqlParameter("@Chamber5Comments", SqlDbType.NVarChar, 50)
        arParameters(4).Value = Chamber5Comments
        arParameters(5) = New SqlParameter("@Aorticarchcomments", SqlDbType.NVarChar, 50)
        arParameters(5).Value = Aorticarchcomments
        arParameters(6) = New SqlParameter("@Pulmonaryartcomments", SqlDbType.NVarChar, 50)
        arParameters(6).Value = Pulmonaryartcomments
        arParameters(7) = New SqlParameter("@Spinecomments", SqlDbType.NVarChar, 50)
        arParameters(7).Value = Spinecomments
        arParameters(8) = New SqlParameter("@LKidneycomments", SqlDbType.NVarChar, 50)
        arParameters(8).Value = LKidneycomments
        arParameters(9) = New SqlParameter("@RKidneycomments", SqlDbType.NVarChar, 50)
        arParameters(9).Value = RKidneycomments
        arParameters(10) = New SqlParameter("@Bladdercomments", SqlDbType.NVarChar, 50)
        arParameters(10).Value = Bladdercomments
        arParameters(11) = New SqlParameter("@Stomachcomments", SqlDbType.NVarChar, 50)
        arParameters(11).Value = Stomachcomments
        arParameters(12) = New SqlParameter("@UExtremitiescomments", SqlDbType.NVarChar, 50)
        arParameters(12).Value = UExtremitiescomments
        arParameters(13) = New SqlParameter("@LExtremitiescomments", SqlDbType.NVarChar, 50)
        arParameters(13).Value = LExtremitiescomments
        arParameters(14) = New SqlParameter("@Movementcomments", SqlDbType.NVarChar, 50)
        arParameters(14).Value = Movementcomments
        arParameters(15) = New SqlParameter("@Cordinsertioncomments", SqlDbType.NVarChar, 50)
        arParameters(15).Value = Cordinsertioncomments
        arParameters(16) = New SqlParameter("@Vessel3Comments", SqlDbType.NVarChar, 50)
        arParameters(16).Value = Vessel3Comments
        arParameters(17) = New SqlParameter("@Facecomments", SqlDbType.NVarChar, 50)
        arParameters(17).Value = Facecomments
        arParameters(18) = New SqlParameter("@Intracranialcomments", SqlDbType.NVarChar, 50)
        arParameters(18).Value = Intracranialcomments
        arParameters(19) = New SqlParameter("@Cerebellumcomments", SqlDbType.NVarChar, 50)
        arParameters(19).Value = Cerebellumcomments
        arParameters(20) = New SqlParameter("@Latventcomments", SqlDbType.NVarChar, 50)
        arParameters(20).Value = Latventcomments
        arParameters(21) = New SqlParameter("@Cisternacomments", SqlDbType.NVarChar, 50)
        arParameters(21).Value = Cisternacomments
        arParameters(22) = New SqlParameter("@Nuchalcomments", SqlDbType.NVarChar, 50)
        arParameters(22).Value = Nuchalcomments
        arParameters(23) = New SqlParameter("@PLACENTA", SqlDbType.NVarChar, 50)
        arParameters(23).Value = PLACENTA
        arParameters(24) = New SqlParameter("@Previa", SqlDbType.NVarChar, 50)
        arParameters(24).Value = Previa
        arParameters(25) = New SqlParameter("@GSac", SqlDbType.NVarChar, 50)
        arParameters(25).Value = GSac
        arParameters(26) = New SqlParameter("@FetalPole", SqlDbType.NVarChar, 50)
        arParameters(26).Value = FetalPole
        arParameters(27) = New SqlParameter("@Ysac", SqlDbType.NVarChar, 50)
        arParameters(27).Value = Ysac
        arParameters(28) = New SqlParameter("@AFI", SqlDbType.NVarChar, 50)
        arParameters(28).Value = AFI
        arParameters(29) = New SqlParameter("@Adnexa", SqlDbType.NVarChar, 255)
        arParameters(29).Value = Adnexa
        arParameters(30) = New SqlParameter("@Uterus", SqlDbType.NVarChar, 255)
        arParameters(30).Value = Uterus
        arParameters(31) = New SqlParameter("@Procedures", SqlDbType.NVarChar, 4000)
        arParameters(31).Value = Procedures
        arParameters(32) = New SqlParameter("@Ultrasound", SqlDbType.SmallInt)
        arParameters(32).Value = Ultrasound
        arParameters(33) = New SqlParameter("@GeneticCounsel", SqlDbType.SmallInt)
        arParameters(33).Value = GeneticCounsel
        arParameters(34) = New SqlParameter("@Amniocentesis", SqlDbType.SmallInt)
        arParameters(34).Value = Amniocentesis
        arParameters(35) = New SqlParameter("@ExAFP", SqlDbType.SmallInt)
        arParameters(35).Value = ExAFP
        arParameters(36) = New SqlParameter("@Maternalchrome", SqlDbType.SmallInt)
        arParameters(36).Value = Maternalchrome
        arParameters(37) = New SqlParameter("@OtherProcedure", SqlDbType.SmallInt)
        arParameters(37).Value = OtherProcedure
        arParameters(38) = New SqlParameter("@Gauge", SqlDbType.Int)
        arParameters(38).Value = Gauge
        arParameters(39) = New SqlParameter("@AFRemoved", SqlDbType.Int)
        arParameters(39).Value = AFRemoved
        arParameters(40) = New SqlParameter("@NInsertions", SqlDbType.Int)
        arParameters(40).Value = NInsertions
        arParameters(41) = New SqlParameter("@AFColor", SqlDbType.NVarChar, 50)
        arParameters(41).Value = AFColor
        arParameters(42) = New SqlParameter("@Transplacental", SqlDbType.NVarChar, 50)
        arParameters(42).Value = Transplacental
        arParameters(43) = New SqlParameter("@Complications", SqlDbType.NVarChar, 50)
        arParameters(43).Value = Complications
        arParameters(44) = New SqlParameter("@LAmniocentesis", SqlDbType.SmallInt)
        arParameters(44).Value = LAmniocentesis
        arParameters(45) = New SqlParameter("@Rhogam", SqlDbType.NVarChar, 50)
        arParameters(45).Value = Rhogam
        arParameters(46) = New SqlParameter("@Cardiac", SqlDbType.SmallInt)
        arParameters(46).Value = Cardiac
        arParameters(47) = New SqlParameter("@Situs", SqlDbType.SmallInt)
        arParameters(47).Value = Situs
        arParameters(48) = New SqlParameter("@Chamber4", SqlDbType.SmallInt)
        arParameters(48).Value = Chamber4
        arParameters(49) = New SqlParameter("@Chamber5", SqlDbType.SmallInt)
        arParameters(49).Value = Chamber5
        arParameters(50) = New SqlParameter("@Aorticarch", SqlDbType.SmallInt)
        arParameters(50).Value = Aorticarch
        arParameters(51) = New SqlParameter("@Pulmonaryart", SqlDbType.SmallInt)
        arParameters(51).Value = Pulmonaryart
        arParameters(52) = New SqlParameter("@Spine", SqlDbType.SmallInt)
        arParameters(52).Value = Spine
        arParameters(53) = New SqlParameter("@LKidney", SqlDbType.SmallInt)
        arParameters(53).Value = LKidney
        arParameters(54) = New SqlParameter("@RKidney", SqlDbType.SmallInt)
        arParameters(54).Value = RKidney
        arParameters(55) = New SqlParameter("@Bladder", SqlDbType.SmallInt)
        arParameters(55).Value = Bladder
        arParameters(56) = New SqlParameter("@Stomach", SqlDbType.SmallInt)
        arParameters(56).Value = Stomach
        arParameters(57) = New SqlParameter("@Diaphragm", SqlDbType.SmallInt)
        arParameters(57).Value = Diaphragm
        arParameters(58) = New SqlParameter("@UExtremities", SqlDbType.SmallInt)
        arParameters(58).Value = UExtremities
        arParameters(59) = New SqlParameter("@LExtremities", SqlDbType.SmallInt)
        arParameters(59).Value = LExtremities
        arParameters(60) = New SqlParameter("@Movement", SqlDbType.SmallInt)
        arParameters(60).Value = Movement
        arParameters(61) = New SqlParameter("@Cordinsertion", SqlDbType.SmallInt)
        arParameters(61).Value = Cordinsertion
        arParameters(62) = New SqlParameter("@Vessel3", SqlDbType.SmallInt)
        arParameters(62).Value = Vessel3
        arParameters(63) = New SqlParameter("@Face", SqlDbType.SmallInt)
        arParameters(63).Value = Face
        arParameters(64) = New SqlParameter("@Intracranial", SqlDbType.SmallInt)
        arParameters(64).Value = Intracranial
        arParameters(65) = New SqlParameter("@Cerebellum", SqlDbType.SmallInt)
        arParameters(65).Value = Cerebellum
        arParameters(66) = New SqlParameter("@Latvent", SqlDbType.SmallInt)
        arParameters(66).Value = Latvent
        arParameters(67) = New SqlParameter("@Cisterna", SqlDbType.SmallInt)
        arParameters(67).Value = Cisterna
        arParameters(68) = New SqlParameter("@Nuchal", SqlDbType.SmallInt)
        arParameters(68).Value = Nuchal
        arParameters(69) = New SqlParameter("@USDefaultName", SqlDbType.NVarChar, 50)
        arParameters(69).Value = DefaultName
        arParameters(70) = New SqlParameter("@DiaphragmComments", SqlDbType.NVarChar, 50)
        arParameters(70).Value = Diaphragmcomments
        arParameters(71) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(71).Value = ExaminerID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasoundDefaultsUpdate", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUltrasoundDefaultsUpdate", arParameters)
            End If
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
                            ByVal Cardiaccomments As String, _
                            ByVal Situscomments As String, _
                            ByVal Chamber4Comments As String, _
                            ByVal Chamber5Comments As String, _
                            ByVal Aorticarchcomments As String, _
                            ByVal Pulmonaryartcomments As String, _
                            ByVal Spinecomments As String, _
                            ByVal LKidneycomments As String, _
                            ByVal RKidneycomments As String, _
                            ByVal Bladdercomments As String, _
                            ByVal Stomachcomments As String, _
                            ByVal Diaphragmcomments As String, _
                            ByVal UExtremitiescomments As String, _
                            ByVal LExtremitiescomments As String, _
                            ByVal Movementcomments As String, _
                            ByVal Cordinsertioncomments As String, _
                            ByVal Vessel3Comments As String, _
                            ByVal Facecomments As String, _
                            ByVal Intracranialcomments As String, _
                            ByVal Cerebellumcomments As String, _
                            ByVal Latventcomments As String, _
                            ByVal Cisternacomments As String, _
                            ByVal Nuchalcomments As String, _
                            ByVal PLACENTA As String, _
                            ByVal Previa As String, _
                            ByVal Gender As String, _
                            ByVal GSac As String, _
                            ByVal FetalPole As String, _
                            ByVal Ysac As String, _
                            ByVal AFI As String, _
                            ByVal Adnexa As String, _
                            ByVal Uterus As String, _
                            ByVal Procedures As String, _
                            ByVal Ultrasound As Short, _
                            ByVal GeneticCounsel As Short, _
                            ByVal Amniocentesis As Short, _
                            ByVal ExAFP As Short, _
                            ByVal Maternalchrome As Short, _
                            ByVal OtherProcedure As Short, _
                            ByVal Gauge As Integer, _
                            ByVal AFRemoved As Integer, _
                            ByVal NInsertions As Integer, _
                            ByVal AFColor As String, _
                            ByVal Transplacental As String, _
                            ByVal Complications As String, _
                            ByVal LAmniocentesis As Short, _
                            ByVal Rhogam As String, _
                            ByVal Cardiac As Short, _
                            ByVal Situs As Short, _
                            ByVal Chamber4 As Short, _
                            ByVal Chamber5 As Short, _
                            ByVal Aorticarch As Short, _
                            ByVal Pulmonaryart As Short, _
                            ByVal Spine As Short, _
                            ByVal LKidney As Short, _
                            ByVal RKidney As Short, _
                            ByVal Bladder As Short, _
                            ByVal Stomach As Short, _
                            ByVal Diaphragm As Short, _
                            ByVal UExtremities As Short, _
                            ByVal LExtremities As Short, _
                            ByVal Movement As Short, _
                            ByVal Cordinsertion As Short, _
                            ByVal Vessel3 As Short, _
                            ByVal Face As Short, _
                            ByVal Intracranial As Short, _
                            ByVal Cerebellum As Short, _
                            ByVal Latvent As Short, _
                            ByVal Cisterna As Short, _
                            ByVal Nuchal As Short, _
                            ByVal DefaultName As String, _
                            ByVal ExaminerID As Integer) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(71) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@USDefaultID", SqlDbType.Int)
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@Cardiaccomments", SqlDbType.NVarChar, 50)
        arParameters(1).Value = Cardiaccomments
        arParameters(2) = New SqlParameter("@Situscomments", SqlDbType.NVarChar, 50)
        arParameters(2).Value = Situscomments
        arParameters(3) = New SqlParameter("@Chamber4Comments", SqlDbType.NVarChar, 50)
        arParameters(3).Value = Chamber4Comments
        arParameters(4) = New SqlParameter("@Chamber5Comments", SqlDbType.NVarChar, 50)
        arParameters(4).Value = Chamber5Comments
        arParameters(5) = New SqlParameter("@Aorticarchcomments", SqlDbType.NVarChar, 50)
        arParameters(5).Value = Aorticarchcomments
        arParameters(6) = New SqlParameter("@Pulmonaryartcomments", SqlDbType.NVarChar, 50)
        arParameters(6).Value = Pulmonaryartcomments
        arParameters(7) = New SqlParameter("@Spinecomments", SqlDbType.NVarChar, 50)
        arParameters(7).Value = Spinecomments
        arParameters(8) = New SqlParameter("@LKidneycomments", SqlDbType.NVarChar, 50)
        arParameters(8).Value = LKidneycomments
        arParameters(9) = New SqlParameter("@RKidneycomments", SqlDbType.NVarChar, 50)
        arParameters(9).Value = RKidneycomments
        arParameters(10) = New SqlParameter("@Bladdercomments", SqlDbType.NVarChar, 50)
        arParameters(10).Value = Bladdercomments
        arParameters(11) = New SqlParameter("@Stomachcomments", SqlDbType.NVarChar, 50)
        arParameters(11).Value = Stomachcomments
        arParameters(12) = New SqlParameter("@UExtremitiescomments", SqlDbType.NVarChar, 50)
        arParameters(12).Value = UExtremitiescomments
        arParameters(13) = New SqlParameter("@LExtremitiescomments", SqlDbType.NVarChar, 50)
        arParameters(13).Value = LExtremitiescomments
        arParameters(14) = New SqlParameter("@Movementcomments", SqlDbType.NVarChar, 50)
        arParameters(14).Value = Movementcomments
        arParameters(15) = New SqlParameter("@Cordinsertioncomments", SqlDbType.NVarChar, 50)
        arParameters(15).Value = Cordinsertioncomments
        arParameters(16) = New SqlParameter("@Vessel3Comments", SqlDbType.NVarChar, 50)
        arParameters(16).Value = Vessel3Comments
        arParameters(17) = New SqlParameter("@Facecomments", SqlDbType.NVarChar, 50)
        arParameters(17).Value = Facecomments
        arParameters(18) = New SqlParameter("@Intracranialcomments", SqlDbType.NVarChar, 50)
        arParameters(18).Value = Intracranialcomments
        arParameters(19) = New SqlParameter("@Cerebellumcomments", SqlDbType.NVarChar, 50)
        arParameters(19).Value = Cerebellumcomments
        arParameters(20) = New SqlParameter("@Latventcomments", SqlDbType.NVarChar, 50)
        arParameters(20).Value = Latventcomments
        arParameters(21) = New SqlParameter("@Cisternacomments", SqlDbType.NVarChar, 50)
        arParameters(21).Value = Cisternacomments
        arParameters(22) = New SqlParameter("@Nuchalcomments", SqlDbType.NVarChar, 50)
        arParameters(22).Value = Nuchalcomments
        arParameters(23) = New SqlParameter("@PLACENTA", SqlDbType.NVarChar, 50)
        arParameters(23).Value = PLACENTA
        arParameters(24) = New SqlParameter("@Previa", SqlDbType.NVarChar, 50)
        arParameters(24).Value = Previa
        arParameters(25) = New SqlParameter("@GSac", SqlDbType.NVarChar, 50)
        arParameters(25).Value = GSac
        arParameters(26) = New SqlParameter("@FetalPole", SqlDbType.NVarChar, 50)
        arParameters(26).Value = FetalPole
        arParameters(27) = New SqlParameter("@Ysac", SqlDbType.NVarChar, 50)
        arParameters(27).Value = Ysac
        arParameters(28) = New SqlParameter("@AFI", SqlDbType.NVarChar, 50)
        arParameters(28).Value = AFI
        arParameters(29) = New SqlParameter("@Adnexa", SqlDbType.NVarChar, 255)
        arParameters(29).Value = Adnexa
        arParameters(30) = New SqlParameter("@Uterus", SqlDbType.NVarChar, 255)
        arParameters(30).Value = Uterus
        arParameters(31) = New SqlParameter("@Procedures", SqlDbType.NVarChar, 8000)
        arParameters(31).Value = Procedures
        arParameters(32) = New SqlParameter("@Ultrasound", SqlDbType.SmallInt)
        arParameters(32).Value = Ultrasound
        arParameters(33) = New SqlParameter("@GeneticCounsel", SqlDbType.SmallInt)
        arParameters(33).Value = GeneticCounsel
        arParameters(34) = New SqlParameter("@Amniocentesis", SqlDbType.SmallInt)
        arParameters(34).Value = Amniocentesis
        arParameters(35) = New SqlParameter("@ExAFP", SqlDbType.SmallInt)
        arParameters(35).Value = ExAFP
        arParameters(36) = New SqlParameter("@Maternalchrome", SqlDbType.SmallInt)
        arParameters(36).Value = Maternalchrome
        arParameters(37) = New SqlParameter("@OtherProcedure", SqlDbType.SmallInt)
        arParameters(37).Value = OtherProcedure
        arParameters(38) = New SqlParameter("@Gauge", SqlDbType.Int)
        arParameters(38).Value = Gauge
        arParameters(39) = New SqlParameter("@AFRemoved", SqlDbType.Int)
        arParameters(39).Value = AFRemoved
        arParameters(40) = New SqlParameter("@NInsertions", SqlDbType.Int)
        arParameters(40).Value = NInsertions
        arParameters(41) = New SqlParameter("@AFColor", SqlDbType.NVarChar, 50)
        arParameters(41).Value = AFColor
        arParameters(42) = New SqlParameter("@Transplacental", SqlDbType.NVarChar, 50)
        arParameters(42).Value = Transplacental
        arParameters(43) = New SqlParameter("@Complications", SqlDbType.NVarChar, 50)
        arParameters(43).Value = Complications
        arParameters(44) = New SqlParameter("@LAmniocentesis", SqlDbType.SmallInt)
        arParameters(44).Value = LAmniocentesis
        arParameters(45) = New SqlParameter("@Rhogam", SqlDbType.NVarChar, 50)
        arParameters(45).Value = Rhogam
        arParameters(46) = New SqlParameter("@Cardiac", SqlDbType.SmallInt)
        arParameters(46).Value = Cardiac
        arParameters(47) = New SqlParameter("@Situs", SqlDbType.SmallInt)
        arParameters(47).Value = Situs
        arParameters(48) = New SqlParameter("@Chamber4", SqlDbType.SmallInt)
        arParameters(48).Value = Chamber4
        arParameters(49) = New SqlParameter("@Chamber5", SqlDbType.SmallInt)
        arParameters(49).Value = Chamber5
        arParameters(50) = New SqlParameter("@Aorticarch", SqlDbType.SmallInt)
        arParameters(50).Value = Aorticarch
        arParameters(51) = New SqlParameter("@Pulmonaryart", SqlDbType.SmallInt)
        arParameters(51).Value = Pulmonaryart
        arParameters(52) = New SqlParameter("@Spine", SqlDbType.SmallInt)
        arParameters(52).Value = Spine
        arParameters(53) = New SqlParameter("@LKidney", SqlDbType.SmallInt)
        arParameters(53).Value = LKidney
        arParameters(54) = New SqlParameter("@RKidney", SqlDbType.SmallInt)
        arParameters(54).Value = RKidney
        arParameters(55) = New SqlParameter("@Bladder", SqlDbType.SmallInt)
        arParameters(55).Value = Bladder
        arParameters(56) = New SqlParameter("@Stomach", SqlDbType.SmallInt)
        arParameters(56).Value = Stomach
        arParameters(57) = New SqlParameter("@Diaphragm", SqlDbType.SmallInt)
        arParameters(57).Value = Diaphragm
        arParameters(58) = New SqlParameter("@UExtremities", SqlDbType.SmallInt)
        arParameters(58).Value = UExtremities
        arParameters(59) = New SqlParameter("@LExtremities", SqlDbType.SmallInt)
        arParameters(59).Value = LExtremities
        arParameters(60) = New SqlParameter("@Movement", SqlDbType.SmallInt)
        arParameters(60).Value = Movement
        arParameters(61) = New SqlParameter("@Cordinsertion", SqlDbType.SmallInt)
        arParameters(61).Value = Cordinsertion
        arParameters(62) = New SqlParameter("@Vessel3", SqlDbType.SmallInt)
        arParameters(62).Value = Vessel3
        arParameters(63) = New SqlParameter("@Face", SqlDbType.SmallInt)
        arParameters(63).Value = Face
        arParameters(64) = New SqlParameter("@Intracranial", SqlDbType.SmallInt)
        arParameters(64).Value = Intracranial
        arParameters(65) = New SqlParameter("@Cerebellum", SqlDbType.SmallInt)
        arParameters(65).Value = Cerebellum
        arParameters(66) = New SqlParameter("@Latvent", SqlDbType.SmallInt)
        arParameters(66).Value = Latvent
        arParameters(67) = New SqlParameter("@Cisterna", SqlDbType.SmallInt)
        arParameters(67).Value = Cisterna
        arParameters(68) = New SqlParameter("@Nuchal", SqlDbType.SmallInt)
        arParameters(68).Value = Nuchal
        arParameters(69) = New SqlParameter("@USDefaultName", SqlDbType.NVarChar, 50)
        arParameters(69).Value = DefaultName
        arParameters(70) = New SqlParameter("@DiaphragmComments", SqlDbType.NVarChar, 50)
        arParameters(70).Value = Diaphragmcomments
        arParameters(71) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(71).Value = ExaminerID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasoundDefaultsInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUltrasoundDefaultsInsert", arParameters)
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
    Public Function Delete(ByVal ID As Integer) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@USDefaultID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spUltrasoundDefaultsDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spUltrasoundDefaultsDelete", arParameters)
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


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class 'dalUltrasoundDefaults
