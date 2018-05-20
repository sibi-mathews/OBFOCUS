
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Formulary:        dalProgramSettings
'*
'* Class: Data access layer for Table PatientInfo
'*
'* Remarks:     Uses OleDb and embedded SQL for maintaining the data.
'*-----------------------------------------------------------------------------
'*                      CHANGE HISTORY
'*   Change No:   Date:          Author:   Class:
'*   _________    ___________    ______    ____________________________________
'*      001       1/26/2005     MR        Created.                                
'* 
'******************************************************************************
Public Class dalProgramSettings

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum ClasssFields
        fldPatientAutoNum = 0
        fldConnectionString = 1
        fldIPAddress = 2
        fldDelHospitalID = 3
        fldShowAnatomy = 4
        fldDefaultAssessment = 5
        fldDefaultRecommendation = 6
        fldReportTypeID = 7
        fldShowFetalLength = 8
        fldCheckUltrasound = 9
        fldDataIPAddress = 10
        fldShowAbnormalChkBx = 11
        fldSourceDocumentPath = 12
        fldDestDocumentPath = 13
        fldCustomDictPath = 14
        fldSmokingComments = 15
        fldAlcoholComments = 16
        fldDrugsComments = 17
        fldMedicalHistory = 18
        fldSurgicalHistory = 19
        fldGynHistory = 20
        fldFamilyHistory = 21
        fldSocialHistory = 22
        fldTransfusions = 23
        fldMedications = 24
        fldImpImgAsJpg = 25
        fldShowEDCByUltrasound = 26
        fldWordTemplatePath = 27
        fldUseWinfax = 28
        fldCastelleUID = 29
        fldCastellePW = 30
        fldCastelleIPAddress = 31
    End Enum


    'Used for transaction support
    Private _Transaction As SqlTransaction = Nothing


    '**************************************************************************
    '*  
    '* Allergy:        Transaction
    '*
    '* Class: Used for transaction support.
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
    '* Description:        New
    '*
    '* Class: Initialize a new instance of the class.
    '*
    '* Parameters:  None
    '*
    '**************************************************************************
    Public Sub New()
    End Sub 'New


    '**************************************************************************
    '*  
    '* Allergy:        New
    '*
    '* Class: Initialize a new instance of the class.
    '*
    '* Parameters:  Transaction - used for transaction support.
    '*
    '**************************************************************************
    Public Sub New(ByRef Transaction As SqlTransaction)
        Me.Transaction = Transaction
    End Sub 'New

#End Region



#Region "Main procedures - GetByKey, Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Allergy:        GetProgramSettings
    '*
    '* Class: Gets all the values of a record identified by a key.
    '*
    '* Parameters:  Class - Output parameter
    '*              Picture - Output parameter
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function GetProgramSettings(ByRef PatientAutoNum As Short, ByRef IPAddress As String, _
                                    ByRef ConnectionString As String, ByRef DelHospitalID As Integer, _
                                    ByRef ShowAnatomy As Short, _
                                    ByRef DefaultAssessment As String, _
                                    ByRef DefaultRecommendation As String, _
                                    ByRef ReportTypeID As Integer, _
                                    ByRef ShowFetalLength As Short, _
                                    ByRef CheckUltrasound As Short, _
                                    ByRef DataIPAddress As String, _
                                    ByRef ShowAbnormalChkBx As Short, _
                                    ByRef Source_DocumentPath As String, _
                                    ByRef Dest_DocumentPath As String, _
                                    ByRef CustomDictPath As String, _
                                    ByRef SmokingComments As String, _
                                    ByRef AlcoholComments As String, _
                                    ByRef DrugsComments As String, _
                                    ByRef MedicalHistory As String, _
                                    ByRef SurgicalHistory As String, _
                                    ByRef GynHistory As String, _
                                    ByRef FamilyHistory As String, _
                                    ByRef SocialHistory As String, _
                                    ByRef Transfusions As String, _
                                    ByRef Medications As String, _
                                    ByRef ImpImgAsJpg As Short, _
                                    ByRef ShowEDCByUltrasound As Short, _
                                    ByRef WordTemplatePath As String, _
                                    ByRef UseWinfax As Short, _
                                    ByRef CastelleUID As String, _
                                    ByRef CastellePW As String, _
                                    ByRef CastelleIPAddress As String) As Boolean

        Dim arParameters(31) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldPatientAutoNum) = New SqlParameter("@PatientAutoNum", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldPatientAutoNum).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldIPAddress) = New SqlParameter("@IPAddress", SqlDbType.NVarChar, 250)
        arParameters(Me.ClasssFields.fldIPAddress).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldConnectionString) = New SqlParameter("@ConnectionString", SqlDbType.VarChar, 250)
        arParameters(Me.ClasssFields.fldConnectionString).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldDelHospitalID) = New SqlParameter("@DelHospitalID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldDelHospitalID).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldShowAnatomy) = New SqlParameter("@ShowAnatomy", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldShowAnatomy).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldDefaultAssessment) = New SqlParameter("@DefaultAssessment", SqlDbType.VarChar, 1000)
        arParameters(Me.ClasssFields.fldDefaultAssessment).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldDefaultRecommendation) = New SqlParameter("@DefaultRecommendation", SqlDbType.VarChar, 1000)
        arParameters(Me.ClasssFields.fldDefaultRecommendation).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldReportTypeID) = New SqlParameter("@ReportTypeID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldReportTypeID).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldShowFetalLength) = New SqlParameter("@ShowFetalLength", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldShowFetalLength).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldCheckUltrasound) = New SqlParameter("@CheckUltrasound", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldCheckUltrasound).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldDataIPAddress) = New SqlParameter("@DataIPAddress", SqlDbType.NVarChar, 100)
        arParameters(Me.ClasssFields.fldDataIPAddress).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldShowAbnormalChkBx) = New SqlParameter("@ShowAbnormalChkBx", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldShowAbnormalChkBx).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldSourceDocumentPath) = New SqlParameter("@Source_DocumentPath", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldSourceDocumentPath).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldDestDocumentPath) = New SqlParameter("@Dest_DocumentPath", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldDestDocumentPath).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldCustomDictPath) = New SqlParameter("@CustomDictPath", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldCustomDictPath).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldSmokingComments) = New SqlParameter("@SmokingComments", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldSmokingComments).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldAlcoholComments) = New SqlParameter("@AlcoholComments", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldAlcoholComments).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldDrugsComments) = New SqlParameter("@DrugsComments", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldDrugsComments).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldMedicalHistory) = New SqlParameter("@MedicalHistory", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldMedicalHistory).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldSurgicalHistory) = New SqlParameter("@SurgicalHistory", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldSurgicalHistory).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldGynHistory) = New SqlParameter("@GynHistory", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldGynHistory).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldFamilyHistory) = New SqlParameter("@FamilyHistory", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldFamilyHistory).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldSocialHistory) = New SqlParameter("@SocialHistory", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldSocialHistory).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldTransfusions) = New SqlParameter("@Transfusions", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldTransfusions).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldMedications) = New SqlParameter("@Medications", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldMedications).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldImpImgAsJpg) = New SqlParameter("@ImpImgAsJpg", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldImpImgAsJpg).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldShowEDCByUltrasound) = New SqlParameter("@ShowEDCByUltrasound", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldShowEDCByUltrasound).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldWordTemplatePath) = New SqlParameter("@WordTemplatePath", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldWordTemplatePath).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldUseWinfax) = New SqlParameter("@UseWinfax", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldUseWinfax).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldCastelleUID) = New SqlParameter("@CastelleUID", SqlDbType.NVarChar, 25)
        arParameters(Me.ClasssFields.fldCastelleUID).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldCastellePW) = New SqlParameter("@CastellePW", SqlDbType.NVarChar, 25)
        arParameters(Me.ClasssFields.fldCastellePW).Direction = ParameterDirection.Output
        arParameters(Me.ClasssFields.fldCastelleIPAddress) = New SqlParameter("@CastelleIPAddress", SqlDbType.NVarChar, 25)
        arParameters(Me.ClasssFields.fldCastelleIPAddress).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spProgramSettingsGet", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spProgramSettingsGet", arParameters)
            End If


            ' Return False if data was not found.
            'If arParameters(Me.ClasssFields.fldGeneric).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            PatientAutoNum = ProcessNull.GetInt16(arParameters(Me.ClasssFields.fldPatientAutoNum).Value)
            ConnectionString = ProcessNull.GetString(arParameters(Me.ClasssFields.fldConnectionString).Value)
            IPAddress = ProcessNull.GetString(arParameters(Me.ClasssFields.fldIPAddress).Value)
            DelHospitalID = ProcessNull.GetInt32(arParameters(Me.ClasssFields.fldDelHospitalID).Value)
            ShowAnatomy = ProcessNull.GetInt16(arParameters(Me.ClasssFields.fldShowAnatomy).Value)
            DefaultAssessment = ProcessNull.GetString(arParameters(Me.ClasssFields.fldDefaultAssessment).Value)
            DefaultRecommendation = ProcessNull.GetString(arParameters(Me.ClasssFields.fldDefaultRecommendation).Value)
            ReportTypeID = ProcessNull.GetInt32(arParameters(Me.ClasssFields.fldReportTypeID).Value)
            ShowFetalLength = ProcessNull.GetInt16(arParameters(Me.ClasssFields.fldShowFetalLength).Value)
            CheckUltrasound = ProcessNull.GetInt16(arParameters(Me.ClasssFields.fldCheckUltrasound).Value)
            DataIPAddress = ProcessNull.GetString(arParameters(Me.ClasssFields.fldDataIPAddress).Value)
            ShowAbnormalChkBx = ProcessNull.GetInt16(arParameters(Me.ClasssFields.fldShowAbnormalChkBx).Value)
            Source_DocumentPath = ProcessNull.GetString(arParameters(Me.ClasssFields.fldSourceDocumentPath).Value)
            Dest_DocumentPath = ProcessNull.GetString(arParameters(Me.ClasssFields.fldDestDocumentPath).Value)
            CustomDictPath = ProcessNull.GetString(arParameters(Me.ClasssFields.fldCustomDictPath).Value)
            SmokingComments = ProcessNull.GetString(arParameters(Me.ClasssFields.fldSmokingComments).Value)
            AlcoholComments = ProcessNull.GetString(arParameters(Me.ClasssFields.fldAlcoholComments).Value)
            DrugsComments = ProcessNull.GetString(arParameters(Me.ClasssFields.fldDrugsComments).Value)
            MedicalHistory = ProcessNull.GetString(arParameters(Me.ClasssFields.fldMedicalHistory).Value)
            SurgicalHistory = ProcessNull.GetString(arParameters(Me.ClasssFields.fldSurgicalHistory).Value)
            GynHistory = ProcessNull.GetString(arParameters(Me.ClasssFields.fldGynHistory).Value)
            FamilyHistory = ProcessNull.GetString(arParameters(Me.ClasssFields.fldFamilyHistory).Value)
            SocialHistory = ProcessNull.GetString(arParameters(Me.ClasssFields.fldSocialHistory).Value)
            Transfusions = ProcessNull.GetString(arParameters(Me.ClasssFields.fldTransfusions).Value)
            Medications = ProcessNull.GetString(arParameters(Me.ClasssFields.fldMedications).Value)
            ImpImgAsJpg = ProcessNull.GetInt16(arParameters(Me.ClasssFields.fldImpImgAsJpg).Value)
            ShowEDCByUltrasound = ProcessNull.GetInt16(arParameters(Me.ClasssFields.fldShowEDCByUltrasound).Value)
            WordTemplatePath = ProcessNull.GetString(arParameters(Me.ClasssFields.fldWordTemplatePath).Value)
            UseWinfax = ProcessNull.GetInt16(arParameters(Me.ClasssFields.fldUseWinfax).Value)
            CastelleUID = ProcessNull.GetString(arParameters(Me.ClasssFields.fldCastelleUID).Value)
            CastellePW = ProcessNull.GetString(arParameters(Me.ClasssFields.fldCastellePW).Value)
            CastelleIPAddress = ProcessNull.GetString(arParameters(Me.ClasssFields.fldCastelleIPAddress).Value)

            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Allergy:        Update
    '*
    '* Class: Updates a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was updated or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal PatientAutoNum As Short, _
                ByVal IPAddress As String, _
                ByVal ConnectionString As String, _
                ByVal DelHospitalID As Integer, _
                ByVal ShowAnatomy As Short, _
                ByVal DefaultAssessment As String, _
                ByVal DefaultRecommendation As String, _
                ByVal ReportTypeID As Integer, _
                ByVal ShowFetalLength As Short, _
                ByVal CheckUltrasound As Short, _
                ByVal DataIPAddress As String, _
                ByVal ShowAbnormalChkBx As Short, _
                ByVal Source_DocumentPath As String, _
                ByVal Dest_DocumentPath As String, _
                ByVal CustomDictPath As String, _
                ByVal SmokingComments As String, _
                ByVal AlcoholComments As String, _
                ByVal DrugsComments As String, _
                ByVal MedicalHistory As String, _
                ByVal SurgicalHistory As String, _
                ByVal GynHistory As String, _
                ByVal FamilyHistory As String, _
                ByVal SocialHistory As String, _
                ByVal Transfusions As String, _
                ByVal Medications As String, _
                ByVal ImpImgAsJpg As Short, _
                ByVal ShowEDCByUltrasound As Short, _
                ByVal WordTemplatePath As String, _
                ByVal UseWinfax As Short, _
                ByVal CastelleUID As String, _
                ByVal CastellePW As String, _
                ByVal CastelleIPAddress As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(31) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(Me.ClasssFields.fldPatientAutoNum) = New SqlParameter("@PatientAutoNum", SqlDbType.SmallInt)
        arParameters(Me.ClasssFields.fldPatientAutoNum).Value = PatientAutoNum
        arParameters(Me.ClasssFields.fldIPAddress) = New SqlParameter("@IPAddress", SqlDbType.NVarChar, 250)
        arParameters(Me.ClasssFields.fldIPAddress).Value = IPAddress
        arParameters(Me.ClasssFields.fldConnectionString) = New SqlParameter("@ConnectionString", SqlDbType.VarChar, 250)
        arParameters(Me.ClasssFields.fldConnectionString).Value = ConnectionString
        arParameters(Me.ClasssFields.fldDelHospitalID) = New SqlParameter("@DelHospitalID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldDelHospitalID).Value = DelHospitalID
        arParameters(Me.ClasssFields.fldShowAnatomy) = New SqlParameter("@ShowAnatomy", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldShowAnatomy).Value = ShowAnatomy
        arParameters(Me.ClasssFields.fldDefaultAssessment) = New SqlParameter("@DefaultAssessment", SqlDbType.VarChar, 1000)
        arParameters(Me.ClasssFields.fldDefaultAssessment).Value = DefaultAssessment
        arParameters(Me.ClasssFields.fldDefaultRecommendation) = New SqlParameter("@DefaultRecommendation", SqlDbType.VarChar, 1000)
        arParameters(Me.ClasssFields.fldDefaultRecommendation).Value = DefaultRecommendation
        arParameters(Me.ClasssFields.fldReportTypeID) = New SqlParameter("@ReportTypeID", SqlDbType.Int)
        arParameters(Me.ClasssFields.fldReportTypeID).Value = ReportTypeID
        arParameters(Me.ClasssFields.fldShowFetalLength) = New SqlParameter("@ShowFetalLength", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldShowFetalLength).Value = ShowFetalLength
        arParameters(Me.ClasssFields.fldCheckUltrasound) = New SqlParameter("@CheckUltrasound", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldCheckUltrasound).Value = CheckUltrasound
        arParameters(Me.ClasssFields.fldDataIPAddress) = New SqlParameter("@DataIPAddress", SqlDbType.NVarChar, 100)
        arParameters(Me.ClasssFields.fldDataIPAddress).Value = DataIPAddress
        arParameters(Me.ClasssFields.fldShowAbnormalChkBx) = New SqlParameter("@ShowAbnormalChkBx", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldShowAbnormalChkBx).Value = ShowAbnormalChkBx
        arParameters(Me.ClasssFields.fldSourceDocumentPath) = New SqlParameter("@Source_DocumentPath", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldSourceDocumentPath).Value = Source_DocumentPath
        arParameters(Me.ClasssFields.fldDestDocumentPath) = New SqlParameter("@Dest_DocumentPath", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldDestDocumentPath).Value = Dest_DocumentPath
        arParameters(Me.ClasssFields.fldCustomDictPath) = New SqlParameter("@CustomDictPath", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldCustomDictPath).Value = CustomDictPath
        arParameters(Me.ClasssFields.fldSmokingComments) = New SqlParameter("@SmokingComments", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldSmokingComments).Value = SmokingComments
        arParameters(Me.ClasssFields.fldAlcoholComments) = New SqlParameter("@AlcoholComments", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldAlcoholComments).Value = AlcoholComments
        arParameters(Me.ClasssFields.fldDrugsComments) = New SqlParameter("@DrugsComments", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldDrugsComments).Value = DrugsComments
        arParameters(Me.ClasssFields.fldMedicalHistory) = New SqlParameter("@MedicalHistory", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldMedicalHistory).Value = MedicalHistory
        arParameters(Me.ClasssFields.fldSurgicalHistory) = New SqlParameter("@SurgicalHistory", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldSurgicalHistory).Value = SurgicalHistory
        arParameters(Me.ClasssFields.fldGynHistory) = New SqlParameter("@GynHistory", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldGynHistory).Value = GynHistory
        arParameters(Me.ClasssFields.fldFamilyHistory) = New SqlParameter("@FamilyHistory", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldFamilyHistory).Value = FamilyHistory
        arParameters(Me.ClasssFields.fldSocialHistory) = New SqlParameter("@SocialHistory", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldSocialHistory).Value = SocialHistory
        arParameters(Me.ClasssFields.fldTransfusions) = New SqlParameter("@Transfusions", SqlDbType.NVarChar, 50)
        arParameters(Me.ClasssFields.fldTransfusions).Value = Transfusions
        arParameters(Me.ClasssFields.fldMedications) = New SqlParameter("@Medications", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldMedications).Value = Medications
        arParameters(Me.ClasssFields.fldImpImgAsJpg) = New SqlParameter("@ImpImgAsJpg", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldImpImgAsJpg).Value = ImpImgAsJpg
        arParameters(Me.ClasssFields.fldShowEDCByUltrasound) = New SqlParameter("@ShowEDCByUltrasound", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldShowEDCByUltrasound).Value = ShowEDCByUltrasound
        arParameters(Me.ClasssFields.fldWordTemplatePath) = New SqlParameter("@WordTemplatePath", SqlDbType.NVarChar, 255)
        arParameters(Me.ClasssFields.fldWordTemplatePath).Value = WordTemplatePath
        arParameters(Me.ClasssFields.fldUseWinfax) = New SqlParameter("@UseWinfax", SqlDbType.Bit)
        arParameters(Me.ClasssFields.fldUseWinfax).Value = UseWinfax
        arParameters(Me.ClasssFields.fldCastelleUID) = New SqlParameter("@CastelleUID", SqlDbType.NVarChar, 25)
        arParameters(Me.ClasssFields.fldCastelleUID).Value = CastelleUID
        arParameters(Me.ClasssFields.fldCastellePW) = New SqlParameter("@CastellePW", SqlDbType.NVarChar, 25)
        arParameters(Me.ClasssFields.fldCastellePW).Value = CastellePW
        arParameters(Me.ClasssFields.fldCastelleIPAddress) = New SqlParameter("@CastelleIPAddress", SqlDbType.NVarChar, 25)
        arParameters(Me.ClasssFields.fldCastelleIPAddress).Value = CastelleIPAddress

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spProgramSettingsUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spProgramSettingsUpdate", arParameters)
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
End Class 'dalProgramSettings
