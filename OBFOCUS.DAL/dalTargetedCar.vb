
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalTargetedCar
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
Public Class dalTargetedCar

#Region "Module level variables and enums"

    ' Public ENUM used to enumAD columns 
    Public Enum TargetedCarFields
        fldOBUSID = 0
        fldID = 1
        fldFetusname = 2
        fldBPD = 3
        fldEGA = 4
        fldAxis = 5
        fldRate = 6
        fldSeptum = 7
        fldRVIDDIA = 8
        fldLVIDDIA = 9
        fldRVIDLVIDDIA = 10
        fldBIVIDDIA = 11
        fldLAID = 12
        fldPAD = 13
        fldAD = 14
        fldMMVEL = 15
        fldMTVEL = 16
        fldMPVEL = 17
        fldMAVEL = 18
        fldExamID = 19
        fldDescription = 20
        fldCTRatio = 21
        fldVSymmetry = 22
        fldASymmetry = 23
        fldIVSeptum = 24
        fldIASeptum = 25
        fldMitral = 26
        fldLVentricle = 27
        fldRVentricle = 28
        fldTricuspid = 29
        fldForamen = 30
        fldAorticvalve = 31
        fldAorticOutflow = 32
        fldDuctus = 33
        fldAorticDoppler = 34
        fldPulmonaryDoppler = 35
        fldPericardialeff = 36
        fldOther = 37
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
    ''**************************************************************************
    ''*  
    ''* Name:        GetByKey
    ''*
    ''* Description: Gets all the values of a record identified by a key.
    ''*
    ''* Parameters:  Description - Output parameter
    ''*              Picture - Output parameter
    ''*
    ''* Returns:     Boolean indicating if record was found or not. 
    ''*              True (record found); False (otherwise).
    ''*
    ''**************************************************************************
    'Public Function GetByKey(ByVal OBUSID As Integer, _
    '            ByRef ID As Integer, _
    '            ByRef FetusName As String, _
    '            ByRef BPD As Single, _
    '            ByRef EGA As Single, _
    '            ByRef Axis As String, _
    '            ByRef Rate As Single, _
    '            ByRef SepTum As Single, _
    '            ByRef RVIDDIA As Single, _
    '            ByRef LVIDDIA As Single, _
    '            ByRef RVIDLVIDDIA As Single, _
    '            ByRef BIVIDDIA As Single, _
    '            ByRef LAID As Single, _
    '            ByRef PAD As Single, _
    '            ByRef AD As Single, _
    '            ByRef MMVEL As Single, _
    '            ByRef MTVEL As Single, _
    '            ByRef MPVEL As Single, _
    '            ByRef MAVEL As Single, _
    '            ByRef ExamID As Integer) As Boolean
    '    ' Set the stored procedure parameters
    '    Dim arParameters(19) As SqlParameter         ' Array to hold stored procedure parameters

    '    ' Set the stored procedure parameters
    '    arParameters(Me.TargetedCarFields.fldOBUSID) = New SqlParameter("@OBUSID", SqlDbType.Int)
    '    arParameters(Me.TargetedCarFields.fldOBUSID).Value = OBUSID
    '    arParameters(Me.TargetedCarFields.fldID) = New SqlParameter("@CardioID", SqlDbType.Int)
    '    arParameters(Me.TargetedCarFields.fldID).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldFetusname) = New SqlParameter("@fetusname", SqlDbType.NVarChar, 50)
    '    arParameters(Me.TargetedCarFields.fldFetusname).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldBPD) = New SqlParameter("@bpd", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldBPD).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldEGA) = New SqlParameter("@ega", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldEGA).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldAxis) = New SqlParameter("@axis", SqlDbType.NVarChar, 50)
    '    arParameters(Me.TargetedCarFields.fldAxis).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldRate) = New SqlParameter("@rate", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldRate).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldSeptum) = New SqlParameter("@septum", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldSeptum).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldRVIDDIA) = New SqlParameter("@rviddia", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldRVIDDIA).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldLVIDDIA) = New SqlParameter("@lviddia", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldLVIDDIA).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldRVIDLVIDDIA) = New SqlParameter("@RVIDLVIDDIA", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldRVIDLVIDDIA).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldBIVIDDIA) = New SqlParameter("@BIVIDDIA", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldBIVIDDIA).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldLAID) = New SqlParameter("@laid", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldLAID).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldPAD) = New SqlParameter("@pad", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldPAD).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldAD) = New SqlParameter("@ad", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldAD).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldMMVEL) = New SqlParameter("@mmvel", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldMMVEL).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldMTVEL) = New SqlParameter("@mtvel", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldMTVEL).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldMPVEL) = New SqlParameter("@mpvel", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldMPVEL).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldMAVEL) = New SqlParameter("@mavel", SqlDbType.Real)
    '    arParameters(Me.TargetedCarFields.fldMAVEL).Direction = ParameterDirection.Output
    '    arParameters(Me.TargetedCarFields.fldExamID) = New SqlParameter("@ExamID", SqlDbType.Int)
    '    arParameters(Me.TargetedCarFields.fldExamID).Direction = ParameterDirection.Output

    '    ' Call stored procedure
    '    Try
    '        If Me.Transaction Is Nothing Then
    '            SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spTargetedCarGetByKey", arParameters)
    '        Else
    '            SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spTargetedCarGetByKey", arParameters)
    '        End If


    '        ' Return False if data was not found.
    '        If arParameters(Me.TargetedCarFields.fldID).Value Is DBNull.Value Then Return False

    '        ' Return True if data was found. Also populate output (ByRef) parameters.
    '        ID = ProcessNull.GetInt32(arParameters(Me.TargetedCarFields.fldID).Value)
    '        FetusName = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldFetusname).Value)
    '        BPD = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldBPD).Value)
    '        EGA = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldEGA).Value)
    '        Axis = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldAxis).Value)
    '        Rate = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldRate).Value)
    '        SepTum = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldSeptum).Value)
    '        RVIDDIA = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldRVIDDIA).Value)
    '        LVIDDIA = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldLVIDDIA).Value)
    '        RVIDLVIDDIA = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldRVIDLVIDDIA).Value)
    '        BIVIDDIA = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldBIVIDDIA).Value)
    '        LAID = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldLAID).Value)
    '        PAD = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldPAD).Value)
    '        AD = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldAD).Value)
    '        MMVEL = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldMMVEL).Value)
    '        MTVEL = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldMTVEL).Value)
    '        MPVEL = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldMPVEL).Value)
    '        MAVEL = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldMAVEL).Value)
    '        ExamID = ProcessNull.GetInt32(arParameters(Me.TargetedCarFields.fldExamID).Value)
    '        Return True

    '    Catch ex As Exception
    '        ExceptionManager.Publish(ex)
    '        Return False
    '    End Try


    'End Function
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
                ByRef ID As Integer, _
                ByRef FetusName As String, _
                ByRef BPD As Single, _
                ByRef EGA As Single, _
                ByRef Axis As String, _
                ByRef Rate As Single, _
                ByRef SepTum As Single, _
                ByRef RVIDDIA As Single, _
                ByRef LVIDDIA As Single, _
                ByRef RVIDLVIDDIA As Single, _
                ByRef BIVIDDIA As Single, _
                ByRef LAID As Single, _
                ByRef PAD As Single, _
                ByRef AD As Single, _
                ByRef MMVEL As Single, _
                ByRef MTVEL As Single, _
                ByRef MPVEL As Single, _
                ByRef MAVEL As Single, _
                ByRef ExamID As Integer, _
                ByRef Description As String, _
                ByRef CTRatio As Single, _
                ByRef VSymmetry As String, _
                ByRef ASymmetry As String, _
                ByRef IVSeptum As String, _
                ByRef IASeptum As String, _
                ByRef Mitral As String, _
                ByRef LVentricle As String, _
                ByRef RVentricle As String, _
                ByRef Tricuspid As String, _
                ByRef Foramen As String, _
                ByRef Aorticvalve As String, _
                ByRef AorticOutflow As String, _
                ByRef Ductus As String, _
                ByRef AorticDoppler As String, _
                ByRef PulmonaryDoppler As String, _
                ByRef Pericardialeff As String, _
                ByRef Other As String) As Boolean
        ' Set the stored procedure parameters
        Dim arParameters(37) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(Me.TargetedCarFields.fldOBUSID) = New SqlParameter("@OBUSID", SqlDbType.Int)
        arParameters(Me.TargetedCarFields.fldOBUSID).Value = OBUSID
        arParameters(Me.TargetedCarFields.fldID) = New SqlParameter("@CardioID", SqlDbType.Int)
        arParameters(Me.TargetedCarFields.fldID).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldFetusname) = New SqlParameter("@fetusname", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldFetusname).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldBPD) = New SqlParameter("@bpd", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldBPD).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldEGA) = New SqlParameter("@ega", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldEGA).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldAxis) = New SqlParameter("@axis", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldAxis).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldRate) = New SqlParameter("@rate", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldRate).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldSeptum) = New SqlParameter("@septum", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldSeptum).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldRVIDDIA) = New SqlParameter("@rviddia", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldRVIDDIA).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldLVIDDIA) = New SqlParameter("@lviddia", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldLVIDDIA).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldRVIDLVIDDIA) = New SqlParameter("@RVIDLVIDDIA", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldRVIDLVIDDIA).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldBIVIDDIA) = New SqlParameter("@BIVIDDIA", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldBIVIDDIA).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldLAID) = New SqlParameter("@laid", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldLAID).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldPAD) = New SqlParameter("@pad", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldPAD).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldAD) = New SqlParameter("@ad", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldAD).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldMMVEL) = New SqlParameter("@mmvel", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldMMVEL).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldMTVEL) = New SqlParameter("@mtvel", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldMTVEL).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldMPVEL) = New SqlParameter("@mpvel", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldMPVEL).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldMAVEL) = New SqlParameter("@mavel", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldMAVEL).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldExamID) = New SqlParameter("@ExamID", SqlDbType.Int)
        arParameters(Me.TargetedCarFields.fldExamID).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldDescription) = New SqlParameter("@Description", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldDescription).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldCTRatio) = New SqlParameter("@CTRatio", SqlDbType.Real)
        arParameters(Me.TargetedCarFields.fldCTRatio).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldVSymmetry) = New SqlParameter("@VSymmetry", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldVSymmetry).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldASymmetry) = New SqlParameter("@ASymmetry", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldASymmetry).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldIVSeptum) = New SqlParameter("@IVSeptum", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldIVSeptum).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldIASeptum) = New SqlParameter("@IASeptum", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldIASeptum).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldMitral) = New SqlParameter("@Mitral", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldMitral).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldLVentricle) = New SqlParameter("@LVentricle", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldLVentricle).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldRVentricle) = New SqlParameter("@RVentricle", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldRVentricle).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldTricuspid) = New SqlParameter("@Tricuspid", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldTricuspid).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldForamen) = New SqlParameter("@Foramen", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldForamen).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldAorticvalve) = New SqlParameter("@Aorticvalve", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldAorticvalve).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldAorticOutflow) = New SqlParameter("@AorticOutflow", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldAorticOutflow).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldDuctus) = New SqlParameter("@Ductus", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldDuctus).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldAorticDoppler) = New SqlParameter("@AorticDoppler", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldAorticDoppler).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldPulmonaryDoppler) = New SqlParameter("@PulmonaryDoppler", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldPulmonaryDoppler).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldPericardialeff) = New SqlParameter("@Pericardialeff", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldPericardialeff).Direction = ParameterDirection.Output
        arParameters(Me.TargetedCarFields.fldOther) = New SqlParameter("@Other", SqlDbType.NVarChar, 50)
        arParameters(Me.TargetedCarFields.fldOther).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spTargetedCarGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spTargetedCarGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(Me.TargetedCarFields.fldID).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            ID = ProcessNull.GetInt32(arParameters(Me.TargetedCarFields.fldID).Value)
            FetusName = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldFetusname).Value)
            BPD = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldBPD).Value)
            EGA = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldEGA).Value)
            Axis = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldAxis).Value)
            Rate = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldRate).Value)
            SepTum = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldSeptum).Value)
            RVIDDIA = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldRVIDDIA).Value)
            LVIDDIA = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldLVIDDIA).Value)
            RVIDLVIDDIA = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldRVIDLVIDDIA).Value)
            BIVIDDIA = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldBIVIDDIA).Value)
            LAID = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldLAID).Value)
            PAD = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldPAD).Value)
            AD = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldAD).Value)
            MMVEL = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldMMVEL).Value)
            MTVEL = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldMTVEL).Value)
            MPVEL = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldMPVEL).Value)
            MAVEL = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldMAVEL).Value)
            ExamID = ProcessNull.GetInt32(arParameters(Me.TargetedCarFields.fldExamID).Value)
            Description = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldDescription).Value)
            CTRatio = ProcessNull.GetDecimal(arParameters(Me.TargetedCarFields.fldCTRatio).Value)
            VSymmetry = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldVSymmetry).Value)
            ASymmetry = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldASymmetry).Value)
            IVSeptum = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldIVSeptum).Value)
            IASeptum = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldIASeptum).Value)
            Mitral = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldMitral).Value)
            LVentricle = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldLVentricle).Value)
            RVentricle = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldRVentricle).Value)
            Tricuspid = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldTricuspid).Value)
            Foramen = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldForamen).Value)
            Aorticvalve = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldAorticvalve).Value)
            AorticOutflow = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldAorticOutflow).Value)
            Ductus = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldDuctus).Value)
            AorticDoppler = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldAorticDoppler).Value)
            PulmonaryDoppler = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldPulmonaryDoppler).Value)
            Pericardialeff = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldPericardialeff).Value)
            Other = ProcessNull.GetString(arParameters(Me.TargetedCarFields.fldOther).Value)
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
                ByVal FetusName As String, _
                ByVal BPD As Single, _
                ByVal EGA As Single, _
                ByVal Axis As String, _
                ByVal Rate As Single, _
                ByVal SepTum As Single, _
                ByVal RVIDDIA As Single, _
                ByVal LVIDDIA As Single, _
                ByVal RVIDLVIDDIA As Single, _
                ByVal BIVIDDIA As Single, _
                ByVal LAID As Single, _
                ByVal PAD As Single, _
                ByVal AD As Single, _
                ByVal MMVEL As Single, _
                ByVal MTVEL As Single, _
                ByVal MPVEL As Single, _
                ByVal MAVEL As Single, _
                ByVal Description As String, _
                ByVal CTRatio As Single, _
                ByVal VSymmetry As String, _
                ByVal ASymmetry As String, _
                ByVal IVSeptum As String, _
                ByVal IASeptum As String, _
                ByVal Mitral As String, _
                ByVal LVentricle As String, _
                ByVal RVentricle As String, _
                ByVal Tricuspid As String, _
                ByVal Foramen As String, _
                ByVal Aorticvalve As String, _
                ByVal AorticOutflow As String, _
                ByVal Ductus As String, _
                ByVal AorticDoppler As String, _
                ByVal PulmonaryDoppler As String, _
                ByVal Pericardialeff As String, _
                ByVal Other As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(35) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@CardioID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@fetusname", SqlDbType.NVarChar, 50)
        arParameters(1).Value = FetusName
        arParameters(2) = New SqlParameter("@bpd", SqlDbType.Real)
        arParameters(2).Value = BPD
        arParameters(3) = New SqlParameter("@ega", SqlDbType.Real)
        arParameters(3).Value = EGA
        arParameters(4) = New SqlParameter("@axis", SqlDbType.NVarChar, 50)
        arParameters(4).Value = Axis
        arParameters(5) = New SqlParameter("@rate", SqlDbType.Real)
        arParameters(5).Value = Rate
        arParameters(6) = New SqlParameter("@septum", SqlDbType.Real)
        arParameters(6).Value = SepTum
        arParameters(7) = New SqlParameter("@RVIDDIA", SqlDbType.Real)
        arParameters(7).Value = RVIDDIA
        arParameters(8) = New SqlParameter("@LVIDDIA", SqlDbType.Real)
        arParameters(8).Value = LVIDDIA
        arParameters(9) = New SqlParameter("@RVIDLVIDDIA", SqlDbType.Real)
        arParameters(9).Value = RVIDLVIDDIA
        arParameters(10) = New SqlParameter("@BIVIDDIA", SqlDbType.Real)
        arParameters(10).Value = BIVIDDIA
        arParameters(11) = New SqlParameter("@LAID", SqlDbType.Real)
        arParameters(11).Value = LAID
        arParameters(12) = New SqlParameter("@PAD", SqlDbType.Real)
        arParameters(12).Value = PAD
        arParameters(13) = New SqlParameter("@AD", SqlDbType.Real)
        arParameters(13).Value = AD
        arParameters(14) = New SqlParameter("@MMVEL", SqlDbType.Real)
        arParameters(14).Value = MMVEL
        arParameters(15) = New SqlParameter("@MTVEL", SqlDbType.Real)
        arParameters(15).Value = MTVEL
        arParameters(16) = New SqlParameter("@MPVEL", SqlDbType.Real)
        arParameters(16).Value = MPVEL
        arParameters(17) = New SqlParameter("@MAVEL", SqlDbType.Real)
        arParameters(17).Value = MAVEL
        arParameters(18) = New SqlParameter("@Description", SqlDbType.NVarChar, 50)
        arParameters(18).Value = Description
        arParameters(19) = New SqlParameter("@CTRatio", SqlDbType.Real)
        arParameters(19).Value = CTRatio
        arParameters(20) = New SqlParameter("@VSymmetry", SqlDbType.NVarChar, 50)
        arParameters(20).Value = VSymmetry
        arParameters(21) = New SqlParameter("@ASymmetry", SqlDbType.NVarChar, 50)
        arParameters(21).Value = ASymmetry
        arParameters(22) = New SqlParameter("@IVSeptum", SqlDbType.NVarChar, 50)
        arParameters(22).Value = IVSeptum
        arParameters(23) = New SqlParameter("@IASeptum", SqlDbType.NVarChar, 50)
        arParameters(23).Value = IASeptum
        arParameters(24) = New SqlParameter("@Mitral", SqlDbType.NVarChar, 50)
        arParameters(24).Value = Mitral
        arParameters(25) = New SqlParameter("@LVentricle", SqlDbType.NVarChar, 50)
        arParameters(25).Value = LVentricle
        arParameters(26) = New SqlParameter("@RVentricle", SqlDbType.NVarChar, 50)
        arParameters(26).Value = RVentricle
        arParameters(27) = New SqlParameter("@Tricuspid", SqlDbType.NVarChar, 50)
        arParameters(27).Value = Tricuspid
        arParameters(28) = New SqlParameter("@Foramen", SqlDbType.NVarChar, 50)
        arParameters(28).Value = Foramen
        arParameters(29) = New SqlParameter("@Aorticvalve", SqlDbType.NVarChar, 50)
        arParameters(29).Value = Aorticvalve
        arParameters(30) = New SqlParameter("@AorticOutflow", SqlDbType.NVarChar, 50)
        arParameters(30).Value = AorticOutflow
        arParameters(31) = New SqlParameter("@Ductus", SqlDbType.NVarChar, 50)
        arParameters(31).Value = Ductus
        arParameters(32) = New SqlParameter("@AorticDoppler", SqlDbType.NVarChar, 50)
        arParameters(32).Value = AorticDoppler
        arParameters(33) = New SqlParameter("@PulmonaryDoppler", SqlDbType.NVarChar, 50)
        arParameters(33).Value = PulmonaryDoppler
        arParameters(34) = New SqlParameter("@Pericardialeff", SqlDbType.NVarChar, 50)
        arParameters(34).Value = Pericardialeff
        arParameters(35) = New SqlParameter("@Other", SqlDbType.NVarChar, 50)
        arParameters(35).Value = Other
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spTargetedCarUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spTargetedCarUpdate", arParameters)
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


End Class 'dalTargetedCar
