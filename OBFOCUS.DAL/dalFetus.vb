
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalFetus
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
Public Class dalFetus

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 

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
    '* Name:        InsertUltrasound
    '*
    '* Description: InsertUltrasounds a new record to the [Ultrasound] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was InsertUltrasounded or not. 
    '*              True (record InsertUltrasounded); False (otherwise).
    '*
    '**************************************************************************
    Public Function InsertUltrasound(ByVal chartid As Integer, _
                                    ByVal FetusName As String, _
                                    ByVal ExamID As Integer, _
                                    ByVal USGAw As Integer, _
                                    ByVal USGAd As Integer, _
                                    ByVal AFI As String, _
                                    ByVal BPDM As String, _
                                    ByVal ACM As String, _
                                    ByVal HCM As String, _
                                    ByVal CRL As String, _
                                    ByVal SAC As String, _
                                    ByVal FLM As String, _
                                    ByVal EFW As String, _
                                    ByVal PSV As Double, _
                                    ByVal PeakGradient As Double, _
                                    ByVal EDV As Double, _
                                    ByVal MeanVelocity As Double, _
                                    ByVal RI As Double, _
                                    ByVal SD As Double, _
                                    ByVal PI As Double, _
                                    ByVal MCAPSV As Double, _
                                    ByVal MCAPeakGradient As Double, _
                                    ByVal MCAEDV As Double, _
                                    ByVal MCAMeanVelocity As Double, _
                                    ByVal MCARI As Double, _
                                    ByVal MCASD As Double, _
                                    ByVal MCAPI As Double) As Boolean

        Dim arParameters(26) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = chartid
        arParameters(1) = New SqlParameter("@FetusName", SqlDbType.VarChar, 50)
        arParameters(1).Value = FetusName
        arParameters(2) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(2).Value = ExamID
        arParameters(3) = New SqlParameter("@USGAw", SqlDbType.Int)
        arParameters(3).Value = USGAw
        arParameters(4) = New SqlParameter("@USGAd", SqlDbType.Int)
        arParameters(4).Value = USGAd
        arParameters(5) = New SqlParameter("@AFI", SqlDbType.NVarChar, 50)
        arParameters(5).Value = AFI
        arParameters(6) = New SqlParameter("@BPDM", SqlDbType.Real)
        If Not IsNumeric(BPDM) Then
            arParameters(6).Value = DBNull.Value
        Else
            arParameters(6).Value = BPDM
        End If
        arParameters(7) = New SqlParameter("@ACM", SqlDbType.Real)
        If Not IsNumeric(ACM) Then
            arParameters(7).Value = DBNull.Value
        Else
            arParameters(7).Value = ACM
        End If
        arParameters(8) = New SqlParameter("@HCM", SqlDbType.Real)
        If Not IsNumeric(HCM) Then
            arParameters(8).Value = DBNull.Value
        Else
            arParameters(8).Value = HCM
        End If
        arParameters(9) = New SqlParameter("@CRL", SqlDbType.Real)
        If Not IsNumeric(CRL) Then
            arParameters(9).Value = DBNull.Value
        Else
            arParameters(9).Value = CRL
        End If
        arParameters(10) = New SqlParameter("@SAC", SqlDbType.Int)
        If Not IsNumeric(SAC) Then
            arParameters(10).Value = DBNull.Value
        Else
            arParameters(10).Value = SAC
        End If
        arParameters(11) = New SqlParameter("@FLM", SqlDbType.Real)
        If Not IsNumeric(FLM) Then
            arParameters(11).Value = DBNull.Value
        Else
            arParameters(11).Value = FLM
        End If
        arParameters(12) = New SqlParameter("@EFW", SqlDbType.Int)
        If Not IsNumeric(EFW) Then
            arParameters(12).Value = DBNull.Value
        Else
            arParameters(12).Value = EFW
        End If
        arParameters(13) = New SqlParameter("@PSV", SqlDbType.Real)
        arParameters(13).Value = PSV
        arParameters(14) = New SqlParameter("@PeakGradient", SqlDbType.Real)
        arParameters(14).Value = PeakGradient
        arParameters(15) = New SqlParameter("@EDV", SqlDbType.Real)
        arParameters(15).Value = EDV
        arParameters(16) = New SqlParameter("@MeanVelocity", SqlDbType.Real)
        arParameters(16).Value = MeanVelocity
        arParameters(17) = New SqlParameter("@RI", SqlDbType.Real)
        arParameters(17).Value = RI
        arParameters(18) = New SqlParameter("@SD", SqlDbType.Real)
        arParameters(18).Value = SD
        arParameters(19) = New SqlParameter("@PI", SqlDbType.Real)
        arParameters(19).Value = PI
        arParameters(20) = New SqlParameter("@MCAPSV", SqlDbType.Real)
        arParameters(20).Value = MCAPSV
        arParameters(21) = New SqlParameter("@MCAPeakGradient", SqlDbType.Real)
        arParameters(21).Value = MCAPeakGradient
        arParameters(22) = New SqlParameter("@MCAEDV", SqlDbType.Real)
        arParameters(22).Value = MCAEDV
        arParameters(23) = New SqlParameter("@MCAMeanVelocity", SqlDbType.Real)
        arParameters(23).Value = MCAMeanVelocity
        arParameters(24) = New SqlParameter("@MCARI", SqlDbType.Real)
        arParameters(24).Value = MCARI
        arParameters(25) = New SqlParameter("@MCASD", SqlDbType.Real)
        arParameters(25).Value = MCASD
        arParameters(26) = New SqlParameter("@MCAPI", SqlDbType.Real)
        arParameters(26).Value = MCAPI


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "sp_UltrasoundInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "sp_UltrasoundInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try


        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            Return True
        End If

    End Function
    '**************************************************************************
    '*  
    '* Name:        GetFetus
    '*
    '* Description: Returns all records in the [Labs] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetFetus(ByVal ExamID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ExamID
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spFetusGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spFetusGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function
    '**************************************************************************
    '*  
    '* Name:        GetPlacenta
    '*
    '* Description: Returns all records in the [Labs] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetPlacenta(ByVal ExamID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(0).Value = ExamID
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPlacentaGetByKey", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPlacentaGetByKey", arParameters)
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
    Public Function GetByKey(ByVal OBUSID As Integer, _
            ByRef FetusName As String, _
            ByRef Comments As String, _
         ByRef SAC As String, _
         ByRef SACW As String, _
         ByRef CRL As String, _
         ByRef CRLW As String, _
         ByRef Side As String, _
         ByRef Position As String, _
         ByRef BPDM As String, _
         ByRef BPDW As String, _
         ByRef HCM As String, _
         ByRef HCW As String, _
         ByRef ACM As String, _
         ByRef ACW As String, _
         ByRef FLM As String, _
         ByRef FLW As String, _
         ByRef CI As String, _
         ByRef SD As Double, _
         ByRef MCAPSV As Double, _
         ByRef MCARI As Double, _
         ByRef HCAC As String, _
         ByRef USGAw As Integer, _
         ByRef USGAd As Integer, _
         ByRef EFW As String, _
         ByRef CARDIACMOTION As Short, _
         ByRef CARDIACMOTIONab As Short, _
         ByRef Cardiaccomments As String, _
         ByRef Situs As Short, _
         ByRef Situsab As Short, _
         ByRef Situscomments As String, _
         ByRef Chamber4 As Short, _
         ByRef Chamber4ab As Short, _
         ByRef Chamber4comments As String, _
         ByRef Chamber5 As Short, _
         ByRef Chamber5ab As Short, _
         ByRef Chamber5comments As String, _
         ByRef Aorticarch As Short, _
         ByRef Aorticarchab As Short, _
         ByRef Aorticarchcomments As String, _
         ByRef Pulmonartart As Short, _
         ByRef Pulmonaryartab As Short, _
         ByRef Pulmonaryartcomments As String, _
         ByRef SPINE As Short, _
         ByRef SPINEab As Short, _
         ByRef Spinecomments As String, _
         ByRef LKidney As Short, _
         ByRef LKidneyab As Short, _
         ByRef LKidneycomments As String, _
         ByRef RKidney As Short, _
         ByRef RKidneyab As Short, _
         ByRef RKidneycomments As String, _
         ByRef Bladder As Short, _
         ByRef Bladderab As Short, _
         ByRef Bladdercomments As String, _
         ByRef STOMACH As Short, _
         ByRef STOMACHab As Short, _
         ByRef Stomachcomments As String, _
         ByRef Diaphragm As Short, _
         ByRef Diaphragmab As Short, _
         ByRef Diaphragmcomments As String, _
         ByRef PLACENTA As String, _
         ByRef Previa As String, _
         ByRef UEXTREMITIES As Short, _
         ByRef UEXTREMITIESab As Short, _
         ByRef UExtremitiescomments As String, _
         ByRef LEXTREMITIES As Short, _
         ByRef LEXTREMITIESab As Short, _
         ByRef LExtremitiescomments As String, _
         ByRef Movement As Short, _
         ByRef Movementab As Short, _
         ByRef Movementcomments As String, _
         ByRef CORDINSERTION As Short, _
         ByRef CORDINSERTIONab As Short, _
         ByRef Cordinsertioncomments As String, _
         ByRef Vessel3 As Short, _
         ByRef Vessel3ab As Short, _
         ByRef Vessel3comments As String, _
         ByRef FACE As Short, _
         ByRef FACEab As Short, _
         ByRef Facecomments As String, _
         ByRef INTRACRANIALANATOMY As Short, _
         ByRef INTRACRANIALANATOMYab As Short, _
         ByRef Intracranialcomments As String, _
         ByRef Cerebellum As Short, _
         ByRef Cerebellumab As Short, _
         ByRef Cerebellumcomments As String, _
         ByRef Latvent As Short, _
         ByRef Latventab As Short, _
         ByRef Latventcomments As String, _
         ByRef Cisterna As Short, _
         ByRef Cisternaab As Short, _
         ByRef Cisternacomments As String, _
         ByRef Nuchal As Short, _
         ByRef Nuchalab As Short, _
         ByRef Nuchalcomments As String, _
         ByRef NT As Double, _
         ByRef Gender As String, _
         ByRef AFI As String, _
         ByRef BPDP As String, _
         ByRef ACP As String, _
         ByRef HCP As String, _
         ByRef FLP As String, _
         ByRef Procedures As String, _
         ByRef OD450 As Double, _
         ByRef Ultrasound As Short, _
         ByRef GeneticCounsel As Short, _
         ByRef Amniocentesis As Short, _
         ByRef ExAFP As Short, _
         ByRef Maternalchrome As Short, _
         ByRef OtherProcedure As Short, _
         ByRef Gauge As Integer, _
         ByRef AFRemoved As Integer, _
         ByRef PSV As Double, _
        ByRef PeakGradient As Double, _
        ByRef EDV As Double, _
        ByRef MeanVelocity As Double, _
        ByRef RI As Double, _
        ByRef PI As Double, _
        ByRef MCAPeakGradient As Double, _
        ByRef MCAEDV As Double, _
        ByRef MCAMeanVelocity As Double, _
        ByRef MCASD As Double, _
        ByRef MCAPI As Double, _
        ByRef LAmniocentesis As Short, _
        ByRef NInsertions As Integer, _
        ByRef AFColor As String, _
        ByRef Transplacental As String, _
        ByRef Complications As String, _
        ByRef Rhogam As String, _
        ByRef HeartRate As Double, _
        ByRef gSac As String, _
        ByRef FetalPole As String, _
        ByRef ySac As String, _
        ByRef Neck As Short, _
        ByRef NeckAb As Short, _
        ByRef Diap As Short, _
        ByRef DiapAb As Short) As Boolean

        Dim arParameters(137) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@OBUSID", SqlDbType.Int)
        arParameters(0).Value = OBUSID
        arParameters(1) = New SqlParameter("@FetusName", SqlDbType.NVarChar, 50)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@Comments", SqlDbType.VarChar, 8000)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@SAC", SqlDbType.Int)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@SACW", SqlDbType.Real)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@CRL", SqlDbType.Real)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@CRLW", SqlDbType.Real)
        arParameters(6).Direction = ParameterDirection.Output
        arParameters(7) = New SqlParameter("@Side", SqlDbType.NVarChar, 50)
        arParameters(7).Direction = ParameterDirection.Output
        arParameters(8) = New SqlParameter("@Position", SqlDbType.NVarChar, 50)
        arParameters(8).Direction = ParameterDirection.Output
        arParameters(9) = New SqlParameter("@BPDM", SqlDbType.Real)
        arParameters(9).Direction = ParameterDirection.Output
        arParameters(10) = New SqlParameter("@BPDW", SqlDbType.Real)
        arParameters(10).Direction = ParameterDirection.Output
        arParameters(11) = New SqlParameter("@HCM", SqlDbType.Real)
        arParameters(11).Direction = ParameterDirection.Output
        arParameters(12) = New SqlParameter("@HCW", SqlDbType.Real)
        arParameters(12).Direction = ParameterDirection.Output
        arParameters(13) = New SqlParameter("@ACM", SqlDbType.Real)
        arParameters(13).Direction = ParameterDirection.Output
        arParameters(14) = New SqlParameter("@ACW", SqlDbType.Real)
        arParameters(14).Direction = ParameterDirection.Output
        arParameters(15) = New SqlParameter("@FLM", SqlDbType.Real)
        arParameters(15).Direction = ParameterDirection.Output
        arParameters(16) = New SqlParameter("@FLW", SqlDbType.Real)
        arParameters(16).Direction = ParameterDirection.Output
        arParameters(17) = New SqlParameter("@CI", SqlDbType.Real)
        arParameters(17).Direction = ParameterDirection.Output
        arParameters(18) = New SqlParameter("@SD", SqlDbType.Real)
        arParameters(18).Direction = ParameterDirection.Output
        arParameters(19) = New SqlParameter("@MCAPSV", SqlDbType.Real)
        arParameters(19).Direction = ParameterDirection.Output
        arParameters(20) = New SqlParameter("@MCARI", SqlDbType.Real)
        arParameters(20).Direction = ParameterDirection.Output
        arParameters(21) = New SqlParameter("@HCAC", SqlDbType.Real)
        arParameters(21).Direction = ParameterDirection.Output
        arParameters(22) = New SqlParameter("@USGAw", SqlDbType.Int)
        arParameters(22).Direction = ParameterDirection.Output
        arParameters(23) = New SqlParameter("@USGAd", SqlDbType.Int)
        arParameters(23).Direction = ParameterDirection.Output
        arParameters(24) = New SqlParameter("@EFW", SqlDbType.Int)
        arParameters(24).Direction = ParameterDirection.Output
        arParameters(25) = New SqlParameter("@CARDIACMOTION", SqlDbType.SmallInt)
        arParameters(25).Direction = ParameterDirection.Output
        arParameters(26) = New SqlParameter("@CARDIACMOTIONab", SqlDbType.SmallInt)
        arParameters(26).Direction = ParameterDirection.Output
        arParameters(27) = New SqlParameter("@Cardiaccomments", SqlDbType.NVarChar, 50)
        arParameters(27).Direction = ParameterDirection.Output
        arParameters(28) = New SqlParameter("@Situs", SqlDbType.SmallInt)
        arParameters(28).Direction = ParameterDirection.Output
        arParameters(29) = New SqlParameter("@Situsab", SqlDbType.SmallInt)
        arParameters(29).Direction = ParameterDirection.Output
        arParameters(30) = New SqlParameter("@Situscomments", SqlDbType.NVarChar, 50)
        arParameters(30).Direction = ParameterDirection.Output
        arParameters(31) = New SqlParameter("@4Chamber", SqlDbType.SmallInt)
        arParameters(31).Direction = ParameterDirection.Output
        arParameters(32) = New SqlParameter("@4Chamberab", SqlDbType.SmallInt)
        arParameters(32).Direction = ParameterDirection.Output
        arParameters(33) = New SqlParameter("@4Chambercomments", SqlDbType.NVarChar, 50)
        arParameters(33).Direction = ParameterDirection.Output
        arParameters(34) = New SqlParameter("@5Chamber", SqlDbType.SmallInt)
        arParameters(34).Direction = ParameterDirection.Output
        arParameters(35) = New SqlParameter("@5Chamberab", SqlDbType.SmallInt)
        arParameters(35).Direction = ParameterDirection.Output
        arParameters(36) = New SqlParameter("@5Chambercomments", SqlDbType.NVarChar, 50)
        arParameters(36).Direction = ParameterDirection.Output
        arParameters(37) = New SqlParameter("@Aorticarch", SqlDbType.SmallInt)
        arParameters(37).Direction = ParameterDirection.Output
        arParameters(38) = New SqlParameter("@Aorticarchab", SqlDbType.SmallInt)
        arParameters(38).Direction = ParameterDirection.Output
        arParameters(39) = New SqlParameter("@Aorticarchcomments", SqlDbType.NVarChar, 50)
        arParameters(39).Direction = ParameterDirection.Output
        arParameters(40) = New SqlParameter("@Pulmonartart", SqlDbType.SmallInt)
        arParameters(40).Direction = ParameterDirection.Output
        arParameters(41) = New SqlParameter("@Pulmonaryartab", SqlDbType.SmallInt)
        arParameters(41).Direction = ParameterDirection.Output
        arParameters(42) = New SqlParameter("@Pulmonaryartcomments", SqlDbType.NVarChar, 50)
        arParameters(42).Direction = ParameterDirection.Output
        arParameters(43) = New SqlParameter("@SPINE", SqlDbType.SmallInt)
        arParameters(43).Direction = ParameterDirection.Output
        arParameters(44) = New SqlParameter("@SPINEab", SqlDbType.SmallInt)
        arParameters(44).Direction = ParameterDirection.Output
        arParameters(45) = New SqlParameter("@SPINEcomments", SqlDbType.NVarChar, 50)
        arParameters(45).Direction = ParameterDirection.Output
        arParameters(46) = New SqlParameter("@LKidney", SqlDbType.SmallInt)
        arParameters(46).Direction = ParameterDirection.Output
        arParameters(47) = New SqlParameter("@LKidneyab", SqlDbType.SmallInt)
        arParameters(47).Direction = ParameterDirection.Output
        arParameters(48) = New SqlParameter("@LKidneycomments", SqlDbType.NVarChar, 50)
        arParameters(48).Direction = ParameterDirection.Output
        arParameters(49) = New SqlParameter("@RKidney", SqlDbType.SmallInt)
        arParameters(49).Direction = ParameterDirection.Output
        arParameters(50) = New SqlParameter("@RKidneyab", SqlDbType.SmallInt)
        arParameters(50).Direction = ParameterDirection.Output
        arParameters(51) = New SqlParameter("@RKidneyComments", SqlDbType.NVarChar, 50)
        arParameters(51).Direction = ParameterDirection.Output
        arParameters(52) = New SqlParameter("@Bladder", SqlDbType.SmallInt)
        arParameters(52).Direction = ParameterDirection.Output
        arParameters(53) = New SqlParameter("@Bladderab", SqlDbType.SmallInt)
        arParameters(53).Direction = ParameterDirection.Output
        arParameters(54) = New SqlParameter("@BladderComments", SqlDbType.NVarChar, 50)
        arParameters(54).Direction = ParameterDirection.Output
        arParameters(55) = New SqlParameter("@STOMACH", SqlDbType.SmallInt)
        arParameters(55).Direction = ParameterDirection.Output
        arParameters(56) = New SqlParameter("@STOMACHab", SqlDbType.SmallInt)
        arParameters(56).Direction = ParameterDirection.Output
        arParameters(57) = New SqlParameter("@STOMACHComments", SqlDbType.NVarChar, 50)
        arParameters(57).Direction = ParameterDirection.Output
        arParameters(58) = New SqlParameter("@Diaphragm", SqlDbType.SmallInt)
        arParameters(58).Direction = ParameterDirection.Output
        arParameters(59) = New SqlParameter("@Diaphragmab", SqlDbType.SmallInt)
        arParameters(59).Direction = ParameterDirection.Output
        arParameters(60) = New SqlParameter("@DiaphragmComments", SqlDbType.NVarChar, 50)
        arParameters(60).Direction = ParameterDirection.Output
        arParameters(61) = New SqlParameter("@PLACENTA", SqlDbType.NVarChar, 50)
        arParameters(61).Direction = ParameterDirection.Output
        arParameters(62) = New SqlParameter("@Previa", SqlDbType.NVarChar, 50)
        arParameters(62).Direction = ParameterDirection.Output
        arParameters(63) = New SqlParameter("@UEXTREMITIES", SqlDbType.SmallInt)
        arParameters(63).Direction = ParameterDirection.Output
        arParameters(64) = New SqlParameter("@UEXTREMITIESab", SqlDbType.SmallInt)
        arParameters(64).Direction = ParameterDirection.Output
        arParameters(65) = New SqlParameter("@UExtremitiescomments", SqlDbType.NVarChar, 50)
        arParameters(65).Direction = ParameterDirection.Output
        arParameters(66) = New SqlParameter("@LEXTREMITIES", SqlDbType.SmallInt)
        arParameters(66).Direction = ParameterDirection.Output
        arParameters(67) = New SqlParameter("@LEXTREMITIESab", SqlDbType.SmallInt)
        arParameters(67).Direction = ParameterDirection.Output
        arParameters(68) = New SqlParameter("@LEXTREMITIEScomments", SqlDbType.NVarChar, 50)
        arParameters(68).Direction = ParameterDirection.Output
        arParameters(69) = New SqlParameter("@Movement", SqlDbType.SmallInt)
        arParameters(69).Direction = ParameterDirection.Output
        arParameters(70) = New SqlParameter("@Movementab", SqlDbType.SmallInt)
        arParameters(70).Direction = ParameterDirection.Output
        arParameters(71) = New SqlParameter("@Movementcomments", SqlDbType.NVarChar, 50)
        arParameters(71).Direction = ParameterDirection.Output
        arParameters(72) = New SqlParameter("@CORDINSERTION", SqlDbType.SmallInt)
        arParameters(72).Direction = ParameterDirection.Output
        arParameters(73) = New SqlParameter("@CORDINSERTIONab", SqlDbType.SmallInt)
        arParameters(73).Direction = ParameterDirection.Output
        arParameters(74) = New SqlParameter("@CORDINSERTIONcomments", SqlDbType.NVarChar, 50)
        arParameters(74).Direction = ParameterDirection.Output
        arParameters(75) = New SqlParameter("@3Vessel", SqlDbType.SmallInt)
        arParameters(75).Direction = ParameterDirection.Output
        arParameters(76) = New SqlParameter("@3Vesselab", SqlDbType.SmallInt)
        arParameters(76).Direction = ParameterDirection.Output
        arParameters(77) = New SqlParameter("@3Vesselcomments", SqlDbType.NVarChar, 50)
        arParameters(77).Direction = ParameterDirection.Output
        arParameters(78) = New SqlParameter("@FACE", SqlDbType.SmallInt)
        arParameters(78).Direction = ParameterDirection.Output
        arParameters(79) = New SqlParameter("@FACEab", SqlDbType.SmallInt)
        arParameters(79).Direction = ParameterDirection.Output
        arParameters(80) = New SqlParameter("@FACEcomments", SqlDbType.NVarChar, 50)
        arParameters(80).Direction = ParameterDirection.Output
        arParameters(81) = New SqlParameter("@INTRACRANIALANATOMY", SqlDbType.SmallInt)
        arParameters(81).Direction = ParameterDirection.Output
        arParameters(82) = New SqlParameter("@INTRACRANIALANATOMYab", SqlDbType.SmallInt)
        arParameters(82).Direction = ParameterDirection.Output
        arParameters(83) = New SqlParameter("@INTRACRANIALcomments", SqlDbType.NVarChar, 50)
        arParameters(83).Direction = ParameterDirection.Output
        arParameters(84) = New SqlParameter("@Cerebellum", SqlDbType.SmallInt)
        arParameters(84).Direction = ParameterDirection.Output
        arParameters(85) = New SqlParameter("@Cerebellumab", SqlDbType.SmallInt)
        arParameters(85).Direction = ParameterDirection.Output
        arParameters(86) = New SqlParameter("@Cerebellumcomments", SqlDbType.NVarChar, 50)
        arParameters(86).Direction = ParameterDirection.Output
        arParameters(87) = New SqlParameter("@Latvent", SqlDbType.SmallInt)
        arParameters(87).Direction = ParameterDirection.Output
        arParameters(88) = New SqlParameter("@Latventab", SqlDbType.SmallInt)
        arParameters(88).Direction = ParameterDirection.Output
        arParameters(89) = New SqlParameter("@Latventcomments", SqlDbType.NVarChar, 50)
        arParameters(89).Direction = ParameterDirection.Output
        arParameters(90) = New SqlParameter("@Cisterna", SqlDbType.SmallInt)
        arParameters(90).Direction = ParameterDirection.Output
        arParameters(91) = New SqlParameter("@Cisternaab", SqlDbType.SmallInt)
        arParameters(91).Direction = ParameterDirection.Output
        arParameters(92) = New SqlParameter("@Cisternacomments", SqlDbType.NVarChar, 50)
        arParameters(92).Direction = ParameterDirection.Output
        arParameters(93) = New SqlParameter("@Nuchal", SqlDbType.SmallInt)
        arParameters(93).Direction = ParameterDirection.Output
        arParameters(94) = New SqlParameter("@Nuchalab", SqlDbType.SmallInt)
        arParameters(94).Direction = ParameterDirection.Output
        arParameters(95) = New SqlParameter("@Nuchalcomments", SqlDbType.NVarChar, 50)
        arParameters(95).Direction = ParameterDirection.Output
        arParameters(96) = New SqlParameter("@NT", SqlDbType.Real)
        arParameters(96).Direction = ParameterDirection.Output
        arParameters(97) = New SqlParameter("@Gender", SqlDbType.NVarChar, 50)
        arParameters(97).Direction = ParameterDirection.Output
        arParameters(98) = New SqlParameter("@AFI", SqlDbType.NVarChar, 50)
        arParameters(98).Direction = ParameterDirection.Output
        arParameters(99) = New SqlParameter("@BPDP", SqlDbType.Int)
        arParameters(99).Direction = ParameterDirection.Output
        arParameters(100) = New SqlParameter("@ACP", SqlDbType.Int)
        arParameters(100).Direction = ParameterDirection.Output
        arParameters(101) = New SqlParameter("@HCP", SqlDbType.Int)
        arParameters(101).Direction = ParameterDirection.Output
        arParameters(102) = New SqlParameter("@FLP", SqlDbType.Int)
        arParameters(102).Direction = ParameterDirection.Output
        arParameters(103) = New SqlParameter("@Procedures", SqlDbType.VarChar, 8000)
        arParameters(103).Direction = ParameterDirection.Output
        arParameters(104) = New SqlParameter("@OD450", SqlDbType.Real)
        arParameters(104).Direction = ParameterDirection.Output
        arParameters(105) = New SqlParameter("@Ultrasound", SqlDbType.SmallInt)
        arParameters(105).Direction = ParameterDirection.Output
        arParameters(106) = New SqlParameter("@GeneticCounsel", SqlDbType.SmallInt)
        arParameters(106).Direction = ParameterDirection.Output
        arParameters(107) = New SqlParameter("@Amniocentesis", SqlDbType.SmallInt)
        arParameters(107).Direction = ParameterDirection.Output
        arParameters(108) = New SqlParameter("@ExAFP", SqlDbType.SmallInt)
        arParameters(108).Direction = ParameterDirection.Output
        arParameters(109) = New SqlParameter("@Maternalchrome", SqlDbType.SmallInt)
        arParameters(109).Direction = ParameterDirection.Output
        arParameters(110) = New SqlParameter("@OtherProcedure", SqlDbType.SmallInt)
        arParameters(110).Direction = ParameterDirection.Output
        arParameters(111) = New SqlParameter("@Gauge", SqlDbType.Int)
        arParameters(111).Direction = ParameterDirection.Output
        arParameters(112) = New SqlParameter("@AFRemoved", SqlDbType.Int)
        arParameters(112).Direction = ParameterDirection.Output
        arParameters(113) = New SqlParameter("@PSV", SqlDbType.Real)
        arParameters(113).Direction = ParameterDirection.Output
        arParameters(114) = New SqlParameter("@PeakGradient", SqlDbType.Real)
        arParameters(114).Direction = ParameterDirection.Output
        arParameters(115) = New SqlParameter("@EDV", SqlDbType.Real)
        arParameters(115).Direction = ParameterDirection.Output
        arParameters(116) = New SqlParameter("@MeanVelocity", SqlDbType.Real)
        arParameters(116).Direction = ParameterDirection.Output
        arParameters(117) = New SqlParameter("@RI", SqlDbType.Real)
        arParameters(117).Direction = ParameterDirection.Output
        arParameters(118) = New SqlParameter("@PI", SqlDbType.Real)
        arParameters(118).Direction = ParameterDirection.Output
        arParameters(119) = New SqlParameter("@MCAPeakGradient", SqlDbType.Real)
        arParameters(119).Direction = ParameterDirection.Output
        arParameters(120) = New SqlParameter("@MCAEDV", SqlDbType.Real)
        arParameters(120).Direction = ParameterDirection.Output
        arParameters(121) = New SqlParameter("@MCAMeanVelocity", SqlDbType.Real)
        arParameters(121).Direction = ParameterDirection.Output
        arParameters(122) = New SqlParameter("@MCASD", SqlDbType.Real)
        arParameters(122).Direction = ParameterDirection.Output
        arParameters(123) = New SqlParameter("@MCAPI", SqlDbType.Real)
        arParameters(123).Direction = ParameterDirection.Output
        arParameters(124) = New SqlParameter("@LAmniocentesis", SqlDbType.SmallInt)
        arParameters(124).Direction = ParameterDirection.Output
        arParameters(125) = New SqlParameter("@NInsertions", SqlDbType.Int)
        arParameters(125).Direction = ParameterDirection.Output
        arParameters(126) = New SqlParameter("@AFColor", SqlDbType.NVarChar, 50)
        arParameters(126).Direction = ParameterDirection.Output
        arParameters(127) = New SqlParameter("@transplacental", SqlDbType.NVarChar, 50)
        arParameters(127).Direction = ParameterDirection.Output
        arParameters(128) = New SqlParameter("@complications", SqlDbType.NVarChar, 50)
        arParameters(128).Direction = ParameterDirection.Output
        arParameters(129) = New SqlParameter("@rhogam", SqlDbType.NVarChar, 50)
        arParameters(129).Direction = ParameterDirection.Output
        arParameters(130) = New SqlParameter("@HeartRate", SqlDbType.Real)
        arParameters(130).Direction = ParameterDirection.Output
        arParameters(131) = New SqlParameter("@gSac", SqlDbType.NVarChar, 100)
        arParameters(131).Direction = ParameterDirection.Output
        arParameters(132) = New SqlParameter("@FetalPole", SqlDbType.NVarChar, 100)
        arParameters(132).Direction = ParameterDirection.Output
        arParameters(133) = New SqlParameter("@ySac", SqlDbType.NVarChar, 100)
        arParameters(133).Direction = ParameterDirection.Output
        arParameters(134) = New SqlParameter("@Neck", SqlDbType.SmallInt)
        arParameters(134).Direction = ParameterDirection.Output
        arParameters(135) = New SqlParameter("@NeckAb", SqlDbType.SmallInt)
        arParameters(135).Direction = ParameterDirection.Output
        arParameters(136) = New SqlParameter("@Diap", SqlDbType.SmallInt)
        arParameters(136).Direction = ParameterDirection.Output
        arParameters(137) = New SqlParameter("@DiapAb", SqlDbType.SmallInt)
        arParameters(137).Direction = ParameterDirection.Output


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFetusGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFetusGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(1).Value Is DBNull.Value Then Return False
            FetusName = ProcessNull.GetString(arParameters(1).Value)
            Comments = ProcessNull.GetString(arParameters(2).Value)
            If IsNumeric(arParameters(3).Value) Then
                SAC = CType(ProcessNull.GetInt32(arParameters(3).Value), String)
            Else
                SAC = ProcessNull.GetString(arParameters(3).Value)
            End If
            If IsNumeric(arParameters(4).Value) Then
                SACW = CType(Math.Round(ProcessNull.GetDouble(arParameters(4).Value), 1), String)
            Else
                SACW = ProcessNull.GetString(arParameters(4).Value)
            End If
            If IsNumeric(arParameters(5).Value) Then
                CRL = CType(Math.Round(ProcessNull.GetDouble(arParameters(5).Value), 1), String)
            Else
                CRL = ProcessNull.GetString(arParameters(5).Value)
            End If
            If IsNumeric(arParameters(6).Value) Then
                CRLW = CType(Math.Round(ProcessNull.GetDouble(arParameters(6).Value), 1), String)
            Else
                CRLW = ProcessNull.GetString(arParameters(6).Value)
            End If
            Side = ProcessNull.GetString(arParameters(7).Value)
            Position = ProcessNull.GetString(arParameters(8).Value)
            If IsNumeric(arParameters(9).Value) Then
                BPDM = CType(Math.Round(ProcessNull.GetDouble(arParameters(9).Value), 1), String)
            Else
                BPDM = ProcessNull.GetString(arParameters(9).Value)
            End If
            If IsNumeric(arParameters(10).Value) Then
                BPDW = CType(Math.Round(ProcessNull.GetDouble(arParameters(10).Value), 1), String)
            Else
                BPDW = ProcessNull.GetString(arParameters(10).Value)
            End If
            If IsNumeric(arParameters(11).Value) Then
                HCM = CType(Math.Round(ProcessNull.GetDouble(arParameters(11).Value)), String)
            Else
                HCM = ProcessNull.GetString(arParameters(11).Value)
            End If
            If IsNumeric(arParameters(12).Value) Then
                HCW = CType(Math.Round(ProcessNull.GetDouble(arParameters(12).Value), 1), String)
            Else
                HCW = ProcessNull.GetString(arParameters(12).Value)
            End If
            If IsNumeric(arParameters(13).Value) Then
                ACM = CType(Math.Round(ProcessNull.GetDouble(arParameters(13).Value), 1), String)
            Else
                ACM = ProcessNull.GetString(arParameters(13).Value)
            End If
            If IsNumeric(arParameters(14).Value) Then
                ACW = CType(Math.Round(ProcessNull.GetDouble(arParameters(14).Value), 1), String)
            Else
                ACW = ProcessNull.GetString(arParameters(14).Value)
            End If
            If IsNumeric(arParameters(15).Value) Then
                FLM = CType(Math.Round(ProcessNull.GetDouble(arParameters(15).Value), 2), String)
            Else
                FLM = ProcessNull.GetString(arParameters(15).Value)
            End If
            If IsNumeric(arParameters(16).Value) Then
                FLW = CType(Math.Round(ProcessNull.GetDouble(arParameters(16).Value), 1), String)
            Else
                FLW = ProcessNull.GetString(arParameters(16).Value)
            End If
            If IsNumeric(arParameters(17).Value) Then
                CI = CType(Math.Round(ProcessNull.GetDouble(arParameters(17).Value), 1), String)
            Else
                CI = ProcessNull.GetString(arParameters(17).Value)
            End If
            SD = ProcessNull.GetDouble(arParameters(18).Value)
            SD = Math.Round(SD, 2)
            MCAPSV = ProcessNull.GetDouble(arParameters(19).Value)
            MCAPSV = Math.Round(MCAPSV, 2)
            MCARI = ProcessNull.GetDouble(arParameters(20).Value)
            MCARI = Math.Round(MCARI, 2)
            If IsNumeric(arParameters(21).Value) Then
                HCAC = CType(Math.Round(ProcessNull.GetDouble(arParameters(21).Value), 1), String)
            Else
                HCAC = ProcessNull.GetString(arParameters(21).Value)
            End If
            USGAw = ProcessNull.GetInt32(arParameters(22).Value)
            USGAd = ProcessNull.GetInt32(arParameters(23).Value)
            EFW = ProcessNull.GetString(arParameters(24).Value)
            CARDIACMOTION = ProcessNull.GetInt16(arParameters(25).Value)
            CARDIACMOTIONab = ProcessNull.GetInt16(arParameters(26).Value)
            Cardiaccomments = ProcessNull.GetString(arParameters(27).Value)
            Situs = ProcessNull.GetInt16(arParameters(28).Value)
            Situsab = ProcessNull.GetInt16(arParameters(29).Value)
            Situscomments = ProcessNull.GetString(arParameters(30).Value)
            Chamber4 = ProcessNull.GetInt16(arParameters(31).Value)
            Chamber4ab = ProcessNull.GetInt16(arParameters(32).Value)
            Chamber4comments = ProcessNull.GetString(arParameters(33).Value)
            Chamber5 = ProcessNull.GetInt16(arParameters(34).Value)
            Chamber5ab = ProcessNull.GetInt16(arParameters(35).Value)
            Chamber5comments = ProcessNull.GetString(arParameters(36).Value)
            Aorticarch = ProcessNull.GetInt16(arParameters(37).Value)
            Aorticarchab = ProcessNull.GetInt16(arParameters(38).Value)
            Aorticarchcomments = ProcessNull.GetString(arParameters(39).Value)
            Pulmonartart = ProcessNull.GetInt16(arParameters(40).Value)
            Pulmonaryartab = ProcessNull.GetInt16(arParameters(41).Value)
            Pulmonaryartcomments = ProcessNull.GetString(arParameters(42).Value)
            SPINE = ProcessNull.GetInt16(arParameters(43).Value)
            SPINEab = ProcessNull.GetInt16(arParameters(44).Value)
            Spinecomments = ProcessNull.GetString(arParameters(45).Value)
            LKidney = ProcessNull.GetInt16(arParameters(46).Value)
            LKidneyab = ProcessNull.GetInt16(arParameters(47).Value)
            LKidneycomments = ProcessNull.GetString(arParameters(48).Value)
            RKidney = ProcessNull.GetInt16(arParameters(49).Value)
            RKidneyab = ProcessNull.GetInt16(arParameters(50).Value)
            RKidneycomments = ProcessNull.GetString(arParameters(51).Value)
            Bladder = ProcessNull.GetInt16(arParameters(52).Value)
            Bladderab = ProcessNull.GetInt16(arParameters(53).Value)
            Bladdercomments = ProcessNull.GetString(arParameters(54).Value)
            STOMACH = ProcessNull.GetInt16(arParameters(55).Value)
            STOMACHab = ProcessNull.GetInt16(arParameters(56).Value)
            Stomachcomments = ProcessNull.GetString(arParameters(57).Value)
            Diaphragm = ProcessNull.GetInt16(arParameters(58).Value)
            Diaphragmab = ProcessNull.GetInt16(arParameters(59).Value)
            Diaphragmcomments = ProcessNull.GetString(arParameters(60).Value)
            PLACENTA = ProcessNull.GetString(arParameters(61).Value)
            Previa = ProcessNull.GetString(arParameters(62).Value)
            UEXTREMITIES = ProcessNull.GetInt16(arParameters(63).Value)
            UEXTREMITIESab = ProcessNull.GetInt16(arParameters(64).Value)
            UExtremitiescomments = ProcessNull.GetString(arParameters(65).Value)
            LEXTREMITIES = ProcessNull.GetInt16(arParameters(66).Value)
            LEXTREMITIESab = ProcessNull.GetInt16(arParameters(67).Value)
            LExtremitiescomments = ProcessNull.GetString(arParameters(68).Value)
            Movement = ProcessNull.GetInt16(arParameters(69).Value)
            Movementab = ProcessNull.GetInt16(arParameters(70).Value)
            Movementcomments = ProcessNull.GetString(arParameters(71).Value)
            CORDINSERTION = ProcessNull.GetInt16(arParameters(72).Value)
            CORDINSERTIONab = ProcessNull.GetInt16(arParameters(73).Value)
            Cordinsertioncomments = ProcessNull.GetString(arParameters(74).Value)
            Vessel3 = ProcessNull.GetInt16(arParameters(75).Value)
            Vessel3ab = ProcessNull.GetInt16(arParameters(76).Value)
            Vessel3comments = ProcessNull.GetString(arParameters(77).Value)
            FACE = ProcessNull.GetInt16(arParameters(78).Value)
            FACEab = ProcessNull.GetInt16(arParameters(79).Value)
            Facecomments = ProcessNull.GetString(arParameters(80).Value)
            INTRACRANIALANATOMY = ProcessNull.GetInt16(arParameters(81).Value)
            INTRACRANIALANATOMYab = ProcessNull.GetInt16(arParameters(82).Value)
            Intracranialcomments = ProcessNull.GetString(arParameters(83).Value)
            Cerebellum = ProcessNull.GetInt16(arParameters(84).Value)
            Cerebellumab = ProcessNull.GetInt16(arParameters(85).Value)
            Cerebellumcomments = ProcessNull.GetString(arParameters(86).Value)
            Latvent = ProcessNull.GetInt16(arParameters(87).Value)
            Latventab = ProcessNull.GetInt16(arParameters(88).Value)
            Latventcomments = ProcessNull.GetString(arParameters(89).Value)
            Cisterna = ProcessNull.GetInt16(arParameters(90).Value)
            Cisternaab = ProcessNull.GetInt16(arParameters(91).Value)
            Cisternacomments = ProcessNull.GetString(arParameters(92).Value)
            Nuchal = ProcessNull.GetInt16(arParameters(93).Value)
            Nuchalab = ProcessNull.GetInt16(arParameters(94).Value)
            Nuchalcomments = ProcessNull.GetString(arParameters(95).Value)
            NT = ProcessNull.GetDecimal(arParameters(96).Value)
            Gender = ProcessNull.GetString(arParameters(97).Value)
            AFI = ProcessNull.GetString(arParameters(98).Value)
            BPDP = ProcessNull.GetString(arParameters(99).Value)
            ACP = ProcessNull.GetString(arParameters(100).Value)
            HCP = ProcessNull.GetString(arParameters(101).Value)
            FLP = ProcessNull.GetString(arParameters(102).Value)
            Procedures = ProcessNull.GetString(arParameters(103).Value)
            OD450 = ProcessNull.GetDouble(arParameters(104).Value)
            Ultrasound = ProcessNull.GetInt16(arParameters(105).Value)
            GeneticCounsel = ProcessNull.GetInt16(arParameters(106).Value)
            Amniocentesis = ProcessNull.GetInt16(arParameters(107).Value)
            ExAFP = ProcessNull.GetInt16(arParameters(108).Value)
            Maternalchrome = ProcessNull.GetInt16(arParameters(109).Value)
            OtherProcedure = ProcessNull.GetInt16(arParameters(110).Value)
            Gauge = ProcessNull.GetInt32(arParameters(111).Value)
            AFRemoved = ProcessNull.GetInt32(arParameters(112).Value)
            PSV = ProcessNull.GetDouble(arParameters(113).Value)
            PSV = Math.Round(PSV, 2)
            PeakGradient = ProcessNull.GetDouble(arParameters(114).Value)
            PeakGradient = Math.Round(PeakGradient, 2)
            EDV = ProcessNull.GetDouble(arParameters(115).Value)
            EDV = Math.Round(EDV, 2)
            MeanVelocity = ProcessNull.GetDouble(arParameters(116).Value)
            MeanVelocity = Math.Round(MeanVelocity, 2)
            RI = ProcessNull.GetDouble(arParameters(117).Value)
            RI = Math.Round(RI, 2)
            PI = ProcessNull.GetDouble(arParameters(118).Value)
            PI = Math.Round(PI, 2)
            MCAPeakGradient = ProcessNull.GetDouble(arParameters(119).Value)
            MCAPeakGradient = Math.Round(MCAPeakGradient, 2)
            MCAEDV = ProcessNull.GetDouble(arParameters(120).Value)
            MCAEDV = Math.Round(MCAEDV, 2)
            MCAMeanVelocity = ProcessNull.GetDouble(arParameters(121).Value)
            MCAMeanVelocity = Math.Round(MCAMeanVelocity, 2)
            MCASD = ProcessNull.GetDouble(arParameters(122).Value)
            MCASD = Math.Round(MCASD, 2)
            MCAPI = ProcessNull.GetDouble(arParameters(123).Value)
            MCAPI = Math.Round(MCAPI, 2)
            LAmniocentesis = ProcessNull.GetInt16(arParameters(124).Value)
            NInsertions = ProcessNull.GetInt32(arParameters(125).Value)
            AFColor = ProcessNull.GetString(arParameters(126).Value)
            Transplacental = ProcessNull.GetString(arParameters(127).Value)
            Complications = ProcessNull.GetString(arParameters(128).Value)
            Rhogam = ProcessNull.GetString(arParameters(129).Value)
            HeartRate = ProcessNull.GetDouble(arParameters(130).Value)
            gSac = ProcessNull.GetString(arParameters(131).Value)
            FetalPole = ProcessNull.GetString(arParameters(132).Value)
            ySac = ProcessNull.GetString(arParameters(133).Value)
            Neck = ProcessNull.GetInt16(arParameters(134).Value)
            NeckAb = ProcessNull.GetInt16(arParameters(135).Value)
            Diap = ProcessNull.GetInt16(arParameters(136).Value)
            DiapAb = ProcessNull.GetInt16(arParameters(137).Value)
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
    Public Function Update(ByVal OBUSID As Integer, _
            ByVal FetusName As String, _
            ByVal Comments As String, _
         ByVal SAC As String, _
         ByVal SACW As String, _
         ByVal CRL As String, _
         ByVal CRLW As String, _
         ByVal Side As String, _
         ByVal Position As String, _
         ByVal BPDM As String, _
         ByVal BPDW As String, _
         ByVal HCM As String, _
         ByVal HCW As String, _
         ByVal ACM As String, _
         ByVal ACW As String, _
         ByVal FLM As String, _
         ByVal FLW As String, _
         ByVal CI As String, _
         ByVal SD As Double, _
         ByVal MCAPSV As Double, _
         ByVal MCARI As Double, _
         ByVal HCAC As String, _
         ByVal USGAw As Integer, _
         ByVal USGAd As Integer, _
         ByVal EFW As String, _
         ByVal CARDIACMOTION As Short, _
         ByVal CARDIACMOTIONab As Short, _
         ByVal Cardiaccomments As String, _
         ByVal Situs As Short, _
         ByVal Situsab As Short, _
         ByVal Situscomments As String, _
         ByVal Chamber4 As Short, _
         ByVal Chamber4ab As Short, _
         ByVal Chamber4comments As String, _
         ByVal Chamber5 As Short, _
         ByVal Chamber5ab As Short, _
         ByVal Chamber5comments As String, _
         ByVal Aorticarch As Short, _
         ByVal Aorticarchab As Short, _
         ByVal Aorticarchcomments As String, _
         ByVal Pulmonartart As Short, _
         ByVal Pulmonaryartab As Short, _
         ByVal Pulmonaryartcomments As String, _
         ByVal SPINE As Short, _
         ByVal SPINEab As Short, _
         ByVal Spinecomments As String, _
         ByVal LKidney As Short, _
         ByVal LKidneyab As Short, _
         ByVal LKidneycomments As String, _
         ByVal RKidney As Short, _
         ByVal RKidneyab As Short, _
         ByVal RKidneycomments As String, _
         ByVal Bladder As Short, _
         ByVal Bladderab As Short, _
         ByVal Bladdercomments As String, _
         ByVal STOMACH As Short, _
         ByVal STOMACHab As Short, _
         ByVal Stomachcomments As String, _
         ByVal Diaphragm As Short, _
         ByVal Diaphragmab As Short, _
         ByVal Diaphragmcomments As String, _
         ByVal PLACENTA As String, _
         ByVal Previa As String, _
         ByVal UEXTREMITIES As Short, _
         ByVal UEXTREMITIESab As Short, _
         ByVal UExtremitiescomments As String, _
         ByVal LEXTREMITIES As Short, _
         ByVal LEXTREMITIESab As Short, _
         ByVal LExtremitiescomments As String, _
         ByVal Movement As Short, _
         ByVal Movementab As Short, _
         ByVal Movementcomments As String, _
         ByVal CORDINSERTION As Short, _
         ByVal CORDINSERTIONab As Short, _
         ByVal Cordinsertioncomments As String, _
         ByVal Vessel3 As Short, _
         ByVal Vessel3ab As Short, _
         ByVal Vessel3comments As String, _
         ByVal FACE As Short, _
         ByVal FACEab As Short, _
         ByVal Facecomments As String, _
         ByVal INTRACRANIALANATOMY As Short, _
         ByVal INTRACRANIALANATOMYab As Short, _
         ByVal Intracranialcomments As String, _
         ByVal Cerebellum As Short, _
         ByVal Cerebellumab As Short, _
         ByVal Cerebellumcomments As String, _
         ByVal Latvent As Short, _
         ByVal Latventab As Short, _
         ByVal Latventcomments As String, _
         ByVal Cisterna As Short, _
         ByVal Cisternaab As Short, _
         ByVal Cisternacomments As String, _
         ByVal Nuchal As Short, _
         ByVal Nuchalab As Short, _
         ByVal Nuchalcomments As String, _
         ByVal NT As Double, _
         ByVal Gender As String, _
         ByVal AFI As String, _
         ByVal BPDP As String, _
         ByVal ACP As String, _
         ByVal HCP As String, _
         ByVal FLP As String, _
         ByVal Procedures As String, _
         ByVal OD450 As Double, _
         ByVal Ultrasound As Short, _
         ByVal GeneticCounsel As Short, _
         ByVal Amniocentesis As Short, _
         ByVal ExAFP As Short, _
         ByVal Maternalchrome As Short, _
         ByVal OtherProcedure As Short, _
         ByVal Gauge As Integer, _
         ByVal AFRemoved As Integer, _
         ByVal Order As String, _
        ByVal PSV As Double, _
        ByVal PeakGradient As Double, _
        ByVal EDV As Double, _
        ByVal MeanVelocity As Double, _
        ByVal RI As Double, _
        ByVal PI As Double, _
        ByVal MCAPeakGradient As Double, _
        ByVal MCAEDV As Double, _
        ByVal MCAMeanVelocity As Double, _
        ByVal MCASD As Double, _
        ByVal MCAPI As Double, _
        ByVal LAmniocentesis As Short, _
        ByVal NInsertions As Integer, _
        ByVal AFColor As String, _
        ByVal Transplacental As String, _
        ByVal Complications As String, _
        ByVal Rhogam As String, _
        ByVal HeartRate As Double, _
        ByVal gSac As String, _
        ByVal FetalPole As String, _
        ByVal ySac As String, _
        ByVal UpdatedBy As String, _
        ByVal Neck As Short, _
        ByVal NeckAb As Short, _
        ByVal Diap As Short, _
        ByVal DiapAb As Short) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(139) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@OBUSID", SqlDbType.Int)
        arParameters(0).Value = OBUSID
        arParameters(1) = New SqlParameter("@FetusName", SqlDbType.NVarChar, 50)
        arParameters(1).Value = FetusName
        arParameters(2) = New SqlParameter("@Comments", SqlDbType.VarChar, 8000)
        arParameters(2).Value = Comments
        arParameters(3) = New SqlParameter("@SAC", SqlDbType.Int)
        If Not IsNumeric(SAC) Then
            arParameters(3).Value = DBNull.Value
        Else
            arParameters(3).Value = SAC
        End If
        arParameters(4) = New SqlParameter("@SACW", SqlDbType.Real)
        If Not IsNumeric(SACW) Then
            arParameters(4).Value = DBNull.Value
        Else
            arParameters(4).Value = SACW
        End If
        arParameters(5) = New SqlParameter("@CRL", SqlDbType.Real)
        If Not IsNumeric(CRL) Then
            arParameters(5).Value = DBNull.Value
        Else
            arParameters(5).Value = CRL
        End If
        arParameters(6) = New SqlParameter("@CRLW", SqlDbType.Real)
        If Not IsNumeric(CRLW) Then
            arParameters(6).Value = DBNull.Value
        Else
            arParameters(6).Value = CRLW
        End If
        arParameters(7) = New SqlParameter("@Side", SqlDbType.NVarChar, 50)
        arParameters(7).Value = Side
        arParameters(8) = New SqlParameter("@Position", SqlDbType.NVarChar, 50)
        arParameters(8).Value = Position
        arParameters(9) = New SqlParameter("@BPDM", SqlDbType.Real)
        If Not IsNumeric(BPDM) Then
            arParameters(9).Value = DBNull.Value
        Else
            arParameters(9).Value = BPDM
        End If
        arParameters(10) = New SqlParameter("@BPDW", SqlDbType.Real)
        If Not IsNumeric(BPDW) Then
            arParameters(10).Value = DBNull.Value
        Else
            arParameters(10).Value = BPDW
        End If
        arParameters(11) = New SqlParameter("@HCM", SqlDbType.Real)
        If Not IsNumeric(HCM) Then
            arParameters(11).Value = DBNull.Value
        Else
            arParameters(11).Value = HCM
        End If
        arParameters(12) = New SqlParameter("@HCW", SqlDbType.Real)
        If Not IsNumeric(HCW) Then
            arParameters(12).Value = DBNull.Value
        Else
            arParameters(12).Value = HCW
        End If
        arParameters(13) = New SqlParameter("@ACM", SqlDbType.Real)
        If Not IsNumeric(ACM) Then
            arParameters(13).Value = DBNull.Value
        Else
            arParameters(13).Value = ACM
        End If
        arParameters(14) = New SqlParameter("@ACW", SqlDbType.Real)
        If Not IsNumeric(ACW) Then
            arParameters(14).Value = DBNull.Value
        Else
            arParameters(14).Value = ACW
        End If
        arParameters(15) = New SqlParameter("@FLM", SqlDbType.Real)
        If Not IsNumeric(FLM) Then
            arParameters(15).Value = DBNull.Value
        Else
            arParameters(15).Value = FLM
        End If
        arParameters(16) = New SqlParameter("@FLW", SqlDbType.Real)
        If Not IsNumeric(FLW) Then
            arParameters(16).Value = DBNull.Value
        Else
            arParameters(16).Value = FLW
        End If
        arParameters(17) = New SqlParameter("@CI", SqlDbType.Real)
        If Not IsNumeric(CI) Then
            arParameters(17).Value = DBNull.Value
        Else
            arParameters(17).Value = CI
        End If
        arParameters(18) = New SqlParameter("@SD", SqlDbType.Real)
        arParameters(18).Value = SD
        arParameters(19) = New SqlParameter("@MCAPSV", SqlDbType.Real)
        arParameters(19).Value = MCAPSV
        arParameters(20) = New SqlParameter("@MCARI", SqlDbType.Real)
        arParameters(20).Value = MCARI
        arParameters(21) = New SqlParameter("@HCAC", SqlDbType.Real)
        If Not IsNumeric(HCAC) Then
            arParameters(21).Value = DBNull.Value
        Else
            arParameters(21).Value = HCAC
        End If
        arParameters(22) = New SqlParameter("@USGAw", SqlDbType.Int)
        arParameters(22).Value = USGAw
        arParameters(23) = New SqlParameter("@USGAd", SqlDbType.Int)
        arParameters(23).Value = USGAd
        arParameters(24) = New SqlParameter("@EFW", SqlDbType.Int)
        If Not IsNumeric(EFW) Then
            arParameters(24).Value = DBNull.Value
        Else
            arParameters(24).Value = EFW
        End If
        arParameters(25) = New SqlParameter("@CARDIACMOTION", SqlDbType.SmallInt)
        arParameters(25).Value = CARDIACMOTION
        arParameters(26) = New SqlParameter("@CARDIACMOTIONab", SqlDbType.SmallInt)
        arParameters(26).Value = CARDIACMOTIONab
        arParameters(27) = New SqlParameter("@Cardiaccomments", SqlDbType.NVarChar, 50)
        arParameters(27).Value = Cardiaccomments
        arParameters(28) = New SqlParameter("@Situs", SqlDbType.SmallInt)
        arParameters(28).Value = Situs
        arParameters(29) = New SqlParameter("@Situsab", SqlDbType.SmallInt)
        arParameters(29).Value = Situsab
        arParameters(30) = New SqlParameter("@Situscomments", SqlDbType.NVarChar, 50)
        arParameters(30).Value = Situscomments
        arParameters(31) = New SqlParameter("@4Chamber", SqlDbType.SmallInt)
        arParameters(31).Value = Chamber4
        arParameters(32) = New SqlParameter("@4Chamberab", SqlDbType.SmallInt)
        arParameters(32).Value = Chamber4ab
        arParameters(33) = New SqlParameter("@4Chambercomments", SqlDbType.NVarChar, 50)
        arParameters(33).Value = Chamber4comments
        arParameters(34) = New SqlParameter("@5Chamber", SqlDbType.SmallInt)
        arParameters(34).Value = Chamber5
        arParameters(35) = New SqlParameter("@5Chamberab", SqlDbType.SmallInt)
        arParameters(35).Value = Chamber5ab
        arParameters(36) = New SqlParameter("@5Chambercomments", SqlDbType.NVarChar, 50)
        arParameters(36).Value = Chamber5comments
        arParameters(37) = New SqlParameter("@Aorticarch", SqlDbType.SmallInt)
        arParameters(37).Value = Aorticarch
        arParameters(38) = New SqlParameter("@Aorticarchab", SqlDbType.SmallInt)
        arParameters(38).Value = Aorticarchab
        arParameters(39) = New SqlParameter("@Aorticarchcomments", SqlDbType.NVarChar, 50)
        arParameters(39).Value = Aorticarchcomments
        arParameters(40) = New SqlParameter("@Pulmonartart", SqlDbType.SmallInt)
        arParameters(40).Value = Pulmonartart
        arParameters(41) = New SqlParameter("@Pulmonaryartab", SqlDbType.SmallInt)
        arParameters(41).Value = Pulmonaryartab
        arParameters(42) = New SqlParameter("@Pulmonaryartcomments", SqlDbType.NVarChar, 50)
        arParameters(42).Value = Pulmonaryartcomments
        arParameters(43) = New SqlParameter("@SPINE", SqlDbType.SmallInt)
        arParameters(43).Value = SPINE
        arParameters(44) = New SqlParameter("@SPINEab", SqlDbType.SmallInt)
        arParameters(44).Value = SPINEab
        arParameters(45) = New SqlParameter("@SPINEcomments", SqlDbType.NVarChar, 50)
        arParameters(45).Value = Spinecomments
        arParameters(46) = New SqlParameter("@LKidney", SqlDbType.SmallInt)
        arParameters(46).Value = LKidney
        arParameters(47) = New SqlParameter("@LKidneyab", SqlDbType.SmallInt)
        arParameters(47).Value = LKidneyab
        arParameters(48) = New SqlParameter("@LKidneycomments", SqlDbType.NVarChar, 50)
        arParameters(48).Value = LKidneycomments
        arParameters(49) = New SqlParameter("@RKidney", SqlDbType.SmallInt)
        arParameters(49).Value = RKidney
        arParameters(50) = New SqlParameter("@RKidneyab", SqlDbType.SmallInt)
        arParameters(50).Value = RKidneyab
        arParameters(51) = New SqlParameter("@RKidneyComments", SqlDbType.NVarChar, 50)
        arParameters(51).Value = RKidneycomments
        arParameters(52) = New SqlParameter("@Bladder", SqlDbType.SmallInt)
        arParameters(52).Value = Bladder
        arParameters(53) = New SqlParameter("@Bladderab", SqlDbType.SmallInt)
        arParameters(53).Value = Bladderab
        arParameters(54) = New SqlParameter("@BladderComments", SqlDbType.NVarChar, 50)
        arParameters(54).Value = Bladdercomments
        arParameters(55) = New SqlParameter("@STOMACH", SqlDbType.SmallInt)
        arParameters(55).Value = STOMACH
        arParameters(56) = New SqlParameter("@STOMACHab", SqlDbType.SmallInt)
        arParameters(56).Value = STOMACHab
        arParameters(57) = New SqlParameter("@STOMACHComments", SqlDbType.NVarChar, 50)
        arParameters(57).Value = Stomachcomments
        arParameters(58) = New SqlParameter("@Diaphragm", SqlDbType.SmallInt)
        arParameters(58).Value = Diaphragm
        arParameters(59) = New SqlParameter("@Diaphragmab", SqlDbType.SmallInt)
        arParameters(59).Value = Diaphragmab
        arParameters(60) = New SqlParameter("@DiaphragmComments", SqlDbType.NVarChar, 50)
        arParameters(60).Value = Diaphragmcomments
        arParameters(61) = New SqlParameter("@PLACENTA", SqlDbType.NVarChar, 50)
        arParameters(61).Value = PLACENTA
        arParameters(62) = New SqlParameter("@Previa", SqlDbType.NVarChar, 50)
        arParameters(62).Value = Previa
        arParameters(63) = New SqlParameter("@UEXTREMITIES", SqlDbType.SmallInt)
        arParameters(63).Value = UEXTREMITIES
        arParameters(64) = New SqlParameter("@UEXTREMITIESab", SqlDbType.SmallInt)
        arParameters(64).Value = UEXTREMITIESab
        arParameters(65) = New SqlParameter("@UExtremitiescomments", SqlDbType.NVarChar, 50)
        arParameters(65).Value = UExtremitiescomments
        arParameters(66) = New SqlParameter("@LEXTREMITIES", SqlDbType.SmallInt)
        arParameters(66).Value = LEXTREMITIES
        arParameters(67) = New SqlParameter("@LEXTREMITIESab", SqlDbType.SmallInt)
        arParameters(67).Value = LEXTREMITIESab
        arParameters(68) = New SqlParameter("@LEXTREMITIEScomments", SqlDbType.NVarChar, 50)
        arParameters(68).Value = LExtremitiescomments
        arParameters(69) = New SqlParameter("@Movement", SqlDbType.SmallInt)
        arParameters(69).Value = Movement
        arParameters(70) = New SqlParameter("@Movementab", SqlDbType.SmallInt)
        arParameters(70).Value = Movementab
        arParameters(71) = New SqlParameter("@Movementcomments", SqlDbType.NVarChar, 50)
        arParameters(71).Value = Movementcomments
        arParameters(72) = New SqlParameter("@CORDINSERTION", SqlDbType.SmallInt)
        arParameters(72).Value = CORDINSERTION
        arParameters(73) = New SqlParameter("@CORDINSERTIONab", SqlDbType.SmallInt)
        arParameters(73).Value = CORDINSERTIONab
        arParameters(74) = New SqlParameter("@CORDINSERTIONcomments", SqlDbType.NVarChar, 50)
        arParameters(74).Value = Cordinsertioncomments
        arParameters(75) = New SqlParameter("@3Vessel", SqlDbType.SmallInt)
        arParameters(75).Value = Vessel3
        arParameters(76) = New SqlParameter("@3Vesselab", SqlDbType.SmallInt)
        arParameters(76).Value = Vessel3ab
        arParameters(77) = New SqlParameter("@3Vesselcomments", SqlDbType.NVarChar, 50)
        arParameters(77).Value = Vessel3comments
        arParameters(78) = New SqlParameter("@FACE", SqlDbType.SmallInt)
        arParameters(78).Value = FACE
        arParameters(79) = New SqlParameter("@FACEab", SqlDbType.SmallInt)
        arParameters(79).Value = FACEab
        arParameters(80) = New SqlParameter("@FACEcomments", SqlDbType.NVarChar, 50)
        arParameters(80).Value = Facecomments
        arParameters(81) = New SqlParameter("@INTRACRANIALANATOMY", SqlDbType.SmallInt)
        arParameters(81).Value = INTRACRANIALANATOMY
        arParameters(82) = New SqlParameter("@INTRACRANIALANATOMYab", SqlDbType.SmallInt)
        arParameters(82).Value = INTRACRANIALANATOMYab
        arParameters(83) = New SqlParameter("@INTRACRANIALcomments", SqlDbType.NVarChar, 50)
        arParameters(83).Value = Intracranialcomments
        arParameters(84) = New SqlParameter("@Cerebellum", SqlDbType.SmallInt)
        arParameters(84).Value = Cerebellum
        arParameters(85) = New SqlParameter("@Cerebellumab", SqlDbType.SmallInt)
        arParameters(85).Value = Cerebellumab
        arParameters(86) = New SqlParameter("@Cerebellumcomments", SqlDbType.NVarChar, 50)
        arParameters(86).Value = Cerebellumcomments
        arParameters(87) = New SqlParameter("@Latvent", SqlDbType.SmallInt)
        arParameters(87).Value = Latvent
        arParameters(88) = New SqlParameter("@Latventab", SqlDbType.SmallInt)
        arParameters(88).Value = Latventab
        arParameters(89) = New SqlParameter("@Latventcomments", SqlDbType.NVarChar, 50)
        arParameters(89).Value = Latventcomments
        arParameters(90) = New SqlParameter("@Cisterna", SqlDbType.SmallInt)
        arParameters(90).Value = Cisterna
        arParameters(91) = New SqlParameter("@Cisternaab", SqlDbType.SmallInt)
        arParameters(91).Value = Cisternaab
        arParameters(92) = New SqlParameter("@Cisternacomments", SqlDbType.NVarChar, 50)
        arParameters(92).Value = Cisternacomments
        arParameters(93) = New SqlParameter("@Nuchal", SqlDbType.SmallInt)
        arParameters(93).Value = Nuchal
        arParameters(94) = New SqlParameter("@Nuchalab", SqlDbType.SmallInt)
        arParameters(94).Value = Nuchalab
        arParameters(95) = New SqlParameter("@Nuchalcomments", SqlDbType.NVarChar, 50)
        arParameters(95).Value = Nuchalcomments
        arParameters(96) = New SqlParameter("@NT", SqlDbType.Real)
        arParameters(96).Value = NT
        arParameters(97) = New SqlParameter("@Gender", SqlDbType.NVarChar, 50)
        arParameters(97).Value = Gender
        arParameters(98) = New SqlParameter("@AFI", SqlDbType.NVarChar, 50)
        arParameters(98).Value = AFI
        arParameters(99) = New SqlParameter("@BPDP", SqlDbType.Int)
        If Not IsNumeric(BPDP) Then
            arParameters(99).Value = DBNull.Value
        Else
            arParameters(99).Value = BPDP
        End If
        arParameters(100) = New SqlParameter("@ACP", SqlDbType.Int)
        If Not IsNumeric(ACP) Then
            arParameters(100).Value = DBNull.Value
        Else
            arParameters(100).Value = ACP
        End If
        arParameters(101) = New SqlParameter("@HCP", SqlDbType.Int)
        If Not IsNumeric(HCP) Then
            arParameters(101).Value = DBNull.Value
        Else
            arParameters(101).Value = HCP
        End If
        arParameters(102) = New SqlParameter("@FLP", SqlDbType.Int)
        If Not IsNumeric(FLP) Then
            arParameters(102).Value = DBNull.Value
        Else
            arParameters(102).Value = FLP
        End If
        arParameters(103) = New SqlParameter("@Procedures", SqlDbType.VarChar, 8000)
        arParameters(103).Value = Procedures
        arParameters(104) = New SqlParameter("@OD450", SqlDbType.Real)
        arParameters(104).Value = OD450
        arParameters(105) = New SqlParameter("@Ultrasound", SqlDbType.SmallInt)
        arParameters(105).Value = Ultrasound
        arParameters(106) = New SqlParameter("@GeneticCounsel", SqlDbType.SmallInt)
        arParameters(106).Value = GeneticCounsel
        arParameters(107) = New SqlParameter("@Amniocentesis", SqlDbType.SmallInt)
        arParameters(107).Value = Amniocentesis
        arParameters(108) = New SqlParameter("@ExAFP", SqlDbType.SmallInt)
        arParameters(108).Value = ExAFP
        arParameters(109) = New SqlParameter("@Maternalchrome", SqlDbType.SmallInt)
        arParameters(109).Value = Maternalchrome
        arParameters(110) = New SqlParameter("@OtherProcedure", SqlDbType.SmallInt)
        arParameters(110).Value = OtherProcedure
        arParameters(111) = New SqlParameter("@Gauge", SqlDbType.Int)
        arParameters(111).Value = Gauge
        arParameters(112) = New SqlParameter("@AFRemoved", SqlDbType.Int)
        arParameters(112).Value = AFRemoved
        arParameters(113) = New SqlParameter("@Order", SqlDbType.NVarChar, 50)
        arParameters(113).Value = Order
        arParameters(114) = New SqlParameter("@PSV", SqlDbType.Real)
        arParameters(114).Value = PSV
        arParameters(115) = New SqlParameter("@PeakGradient", SqlDbType.Real)
        arParameters(115).Value = PeakGradient
        arParameters(116) = New SqlParameter("@EDV", SqlDbType.Real)
        arParameters(116).Value = EDV
        arParameters(117) = New SqlParameter("@MeanVelocity", SqlDbType.Real)
        arParameters(117).Value = MeanVelocity
        arParameters(118) = New SqlParameter("@RI", SqlDbType.Real)
        arParameters(118).Value = RI
        arParameters(119) = New SqlParameter("@PI", SqlDbType.Real)
        arParameters(119).Value = PI
        arParameters(120) = New SqlParameter("@MCAPeakGradient", SqlDbType.Real)
        arParameters(120).Value = MCAPeakGradient
        arParameters(121) = New SqlParameter("@MCAEDV", SqlDbType.Real)
        arParameters(121).Value = MCAEDV
        arParameters(122) = New SqlParameter("@MCAMeanVelocity", SqlDbType.Real)
        arParameters(122).Value = MCAMeanVelocity
        arParameters(123) = New SqlParameter("@MCASD", SqlDbType.Real)
        arParameters(123).Value = MCASD
        arParameters(124) = New SqlParameter("@MCAPI", SqlDbType.Real)
        arParameters(124).Value = MCAPI
        arParameters(125) = New SqlParameter("@LAmniocentesis", SqlDbType.SmallInt)
        arParameters(125).Value = LAmniocentesis
        arParameters(126) = New SqlParameter("@NInsertions", SqlDbType.Int)
        arParameters(126).Value = NInsertions
        arParameters(127) = New SqlParameter("@AFColor", SqlDbType.NVarChar, 50)
        arParameters(127).Value = AFColor
        arParameters(128) = New SqlParameter("@transplacental", SqlDbType.NVarChar, 50)
        arParameters(128).Value = Transplacental
        arParameters(129) = New SqlParameter("@complications", SqlDbType.NVarChar, 50)
        arParameters(129).Value = Complications
        arParameters(130) = New SqlParameter("@rhogam", SqlDbType.NVarChar, 50)
        arParameters(130).Value = Rhogam
        arParameters(131) = New SqlParameter("@HeartRate", SqlDbType.Real)
        arParameters(131).Value = HeartRate
        arParameters(132) = New SqlParameter("@gSac", SqlDbType.NVarChar, 100)
        arParameters(132).Value = gSac
        arParameters(133) = New SqlParameter("@FetalPole", SqlDbType.NVarChar, 100)
        arParameters(133).Value = FetalPole
        arParameters(134) = New SqlParameter("@ySac", SqlDbType.NVarChar, 100)
        arParameters(134).Value = ySac
        arParameters(135) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(135).Value = UpdatedBy
        arParameters(136) = New SqlParameter("@Neck", SqlDbType.SmallInt)
        arParameters(136).Value = Neck
        arParameters(137) = New SqlParameter("@Neckab", SqlDbType.SmallInt)
        arParameters(137).Value = NeckAb
        arParameters(138) = New SqlParameter("@Diap", SqlDbType.SmallInt)
        arParameters(138).Value = Diap
        arParameters(139) = New SqlParameter("@Diapab", SqlDbType.SmallInt)
        arParameters(139).Value = DiapAb

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFetusUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFetusUpdate", arParameters)
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
    Public Function Add(ByRef OBUSID As Integer, _
            ByVal ChartID As Integer, _
            ByVal ExamID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@OBUSID", SqlDbType.Int)
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(1).Value = ChartID
        arParameters(2) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(2).Value = ExamID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFetusInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFetusInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not found.
        If intRecordsAffected = 0 Then
            Return False
        Else
            OBUSID = CType(arParameters(0).Value, Integer)
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
        arParameters(0) = New SqlParameter("@OBUSID", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spFetusDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spFetusDelete", arParameters)
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


End Class 'dalFetus
