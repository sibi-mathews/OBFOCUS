
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalAntenatal
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
Public Class dalAntenatal

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum PhysicianFields
        fldID = 0
        fldTestDate = 1
        fldWeeks = 2
        fldSeen = 3
        fldSBP = 4
        fldDBP = 5
        fldUrineP = 6
        fldUrineLuc = 7
        fldFHRBL = 8
        fldNSTResulta = 9
        fldAFI = 10
        fldBPP = 11
        fldPlacentaGrade = 12
        fldFAS = 13
        fldRN = 14
        fldDoctor = 15
        fldDays = 16
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
    '* Name:        GetAntenatal
    '*
    '* Description: Returns all records in the [Nurses Flow Sheet] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAntenatal(ByVal ChartID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ChartID
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spAntenataltGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spAntenataltGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
        End Try
    End Function

    '**************************************************************************
    '*  
    '* Name:        GetAntenatalwParam
    '*
    '* Description: Returns all records in the [Nurses Flow Sheet] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetAntenatalwParam(ByVal FromDate As Object, ByVal ToDate As Object) As SqlDataReader
        Dim arParameters(1) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@FromDate", SqlDbType.SmallDateTime)
        If Len(FromDate) = 0 Then
            arParameters(0).Value = DBNull.Value
        ElseIf IsDate(FromDate) Then
            arParameters(0).Value = FromDate
        End If
        arParameters(1) = New SqlParameter("@ToDate", SqlDbType.SmallDateTime)
        If Len(ToDate) = 0 Then
            arParameters(1).Value = DBNull.Value
        ElseIf IsDate(ToDate) Then
            arParameters(1).Value = ToDate
        End If
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spAntenatalGetwParam", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spAntenatalGetwParam", arParameters)
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
    Public Function GetByKey(ByVal PhysicianID As Integer, ByRef SiteID As Integer, ByRef Salute As String, _
            ByRef LastName As String, ByRef FirstName As String, ByRef Title As String, _
            ByRef Address As String, ByRef city As String, ByRef State As String, _
            ByRef Country As String, ByRef PostalCode As String, ByRef PhoneNumber As String, _
            ByRef FaxNumber As String, ByRef Suppress As Short) As Boolean

        'Dim TestNull As Object
        'Dim arParameters(13) As SqlParameter         ' Array to hold stored procedure parameters


        '' Set the stored procedure parameters
        'arParameters(Me.PhysicianFields.fldID) = New SqlParameter("@PhysicianID", SqlDbType.Int)
        'arParameters(Me.PhysicianFields.fldID).Value = PhysicianID
        'arParameters(Me.PhysicianFields.fldSiteID) = New SqlParameter("@SiteID", SqlDbType.Int)
        'arParameters(Me.PhysicianFields.fldSiteID).Direction = ParameterDirection.Output
        'arParameters(Me.PhysicianFields.fldSalute) = New SqlParameter("@Salute", SqlDbType.VarChar, 50)
        'arParameters(Me.PhysicianFields.fldSalute).Direction = ParameterDirection.Output
        'arParameters(Me.PhysicianFields.fldLastName) = New SqlParameter("@PlastName", SqlDbType.VarChar, 50)
        'arParameters(Me.PhysicianFields.fldLastName).Direction = ParameterDirection.Output
        'arParameters(Me.PhysicianFields.fldFirstName) = New SqlParameter("@PFirstName", SqlDbType.VarChar, 50)
        'arParameters(Me.PhysicianFields.fldFirstName).Direction = ParameterDirection.Output
        'arParameters(Me.PhysicianFields.fldTitle) = New SqlParameter("@Title", SqlDbType.VarChar, 50)
        'arParameters(Me.PhysicianFields.fldTitle).Direction = ParameterDirection.Output
        'arParameters(Me.PhysicianFields.fldAddress) = New SqlParameter("@PAddress", SqlDbType.VarChar, 255)
        'arParameters(Me.PhysicianFields.fldAddress).Direction = ParameterDirection.Output
        'arParameters(Me.PhysicianFields.fldCity) = New SqlParameter("@PCity", SqlDbType.VarChar, 50)
        'arParameters(Me.PhysicianFields.fldCity).Direction = ParameterDirection.Output
        'arParameters(Me.PhysicianFields.fldState) = New SqlParameter("@PState", SqlDbType.VarChar, 50)
        'arParameters(Me.PhysicianFields.fldState).Direction = ParameterDirection.Output
        'arParameters(Me.PhysicianFields.fldCountry) = New SqlParameter("@PCountry", SqlDbType.VarChar, 50)
        'arParameters(Me.PhysicianFields.fldCountry).Direction = ParameterDirection.Output
        'arParameters(Me.PhysicianFields.fldPostalCode) = New SqlParameter("@PPostalCode", SqlDbType.VarChar, 50)
        'arParameters(Me.PhysicianFields.fldPostalCode).Direction = ParameterDirection.Output
        'arParameters(Me.PhysicianFields.fldPhoneNumber) = New SqlParameter("@PPhoneNumber", SqlDbType.VarChar, 50)
        'arParameters(Me.PhysicianFields.fldPhoneNumber).Direction = ParameterDirection.Output
        'arParameters(Me.PhysicianFields.fldFaxNumber) = New SqlParameter("@PFaxNumber", SqlDbType.VarChar, 50)
        'arParameters(Me.PhysicianFields.fldFaxNumber).Direction = ParameterDirection.Output
        'arParameters(Me.PhysicianFields.fldSuppress) = New SqlParameter("@Suppress", SqlDbType.Bit)
        'arParameters(Me.PhysicianFields.fldSuppress).Direction = ParameterDirection.Output


        ' Call stored procedure
        Try
            'If Me.Transaction Is Nothing Then
            '    SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPhysicianGetByKey", arParameters)
            'Else
            '    SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPhysicianGetByKey", arParameters)
            'End If


            ' Return False if data was not found.
            'If arParameters(Me.PhysicianFields.fldLastName).Value Is DBNull.Value Then Return False

            '' Return True if data was found. Also populate output (ByRef) parameters.
            'SiteID = ProcessNull.GetInt32(arParameters(Me.PhysicianFields.fldSiteID).Value)
            'Salute = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldSalute).Value)
            'Salute = Salute.Trim()
            'LastName = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldLastName).Value)
            'LastName = LastName.Trim()
            'FirstName = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldFirstName).Value)
            'FirstName = FirstName.Trim()
            'Title = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldTitle).Value)
            'Title = Title.Trim()
            'Address = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldAddress).Value)
            'Address = Address.Trim()
            'city = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldCity).Value)
            'city = city.Trim()
            'State = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldState).Value)
            'State = State.Trim()
            'Country = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldCountry).Value)
            'Country = Country.Trim()
            'PostalCode = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldPostalCode).Value)
            'PostalCode = PostalCode.Trim()
            'PhoneNumber = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldPhoneNumber).Value)
            'PhoneNumber = PhoneNumber.Trim()
            'FaxNumber = ProcessNull.GetString(arParameters(Me.PhysicianFields.fldFaxNumber).Value)
            'FaxNumber = FaxNumber.Trim()
            'Suppress = ProcessNull.GetInt16(arParameters(Me.PhysicianFields.fldSuppress).Value)
            'Return True

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
                            ByVal TestDate As String, _
                            ByVal Weeks As Integer, _
                            ByVal Seen As Integer, _
                            ByVal SBP As Integer, _
                            ByVal DBP As Integer, _
                            ByVal Urinep As String, _
                            ByVal UrineLuc As String, _
                            ByVal FHRBL As String, _
                            ByVal NSTResults As String, _
                            ByVal AFI As Integer, _
                            ByVal BPP As String, _
                            ByVal PlacentaGrade As String, _
                            ByVal FAS As String, _
                            ByVal RN As String, _
                            ByVal Doctor As String, _
                            ByVal UserID As String, _
                            ByVal Locked As Short, _
                            ByVal Days As Integer, _
                            ByVal Comments As String, _
                            ByVal TestNumber As Integer, _
                            ByVal FR As String, _
                            ByVal Indications As String, _
                            ByVal UpdatedBy As String) As Boolean


        Dim arParameters(23) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@IDTest", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(1) = New SqlParameter("@Date", SqlDbType.SmallDateTime)
        If TestDate Is Nothing Or TestDate = "" Then
            arParameters(1).Value = DBNull.Value
        Else
            arParameters(1).Value = TestDate
        End If
        arParameters(2) = New SqlParameter("@Weeks", SqlDbType.Int)
        If Weeks = Nothing Then
            arParameters(2).Value = DBNull.Value
        Else
            arParameters(2).Value = Weeks
        End If
        arParameters(3) = New SqlParameter("@Seen", SqlDbType.SmallInt)
        If Seen = Nothing Then
            arParameters(3).Value = DBNull.Value
        Else
            arParameters(3).Value = Seen
        End If
        arParameters(4) = New SqlParameter("@SBP", SqlDbType.Int)
        If SBP = Nothing Then
            arParameters(4).Value = DBNull.Value
        Else
            arParameters(4).Value = SBP
        End If
        arParameters(5) = New SqlParameter("@DBP", SqlDbType.Int)
        If DBP = Nothing Then
            arParameters(5).Value = DBNull.Value
        Else
            arParameters(5).Value = DBP
        End If
        arParameters(6) = New SqlParameter("@UrineP", SqlDbType.NVarChar, 50)
        arParameters(6).Value = Urinep
        arParameters(7) = New SqlParameter("@UrineLuc", SqlDbType.NVarChar, 50)
        arParameters(7).Value = UrineLuc
        arParameters(8) = New SqlParameter("@FHRBL", SqlDbType.NVarChar, 50)
        arParameters(8).Value = FHRBL
        arParameters(9) = New SqlParameter("@NSTResults", SqlDbType.NVarChar, 50)
        arParameters(9).Value = NSTResults
        arParameters(10) = New SqlParameter("@AFI", SqlDbType.Int)
        If AFI = Nothing Then
            arParameters(10).Value = DBNull.Value
        Else
            arParameters(10).Value = AFI
        End If
        arParameters(11) = New SqlParameter("@BPP", SqlDbType.NVarChar, 50)
        arParameters(11).Value = BPP
        arParameters(12) = New SqlParameter("@PlacentaGrade", SqlDbType.NVarChar, 50)
        arParameters(12).Value = PlacentaGrade
        arParameters(13) = New SqlParameter("@FAS", SqlDbType.NVarChar, 50)
        arParameters(13).Value = FAS
        arParameters(14) = New SqlParameter("@RN", SqlDbType.NVarChar, 50)
        arParameters(14).Value = RN
        arParameters(15) = New SqlParameter("@Doctor", SqlDbType.NVarChar, 50)
        arParameters(15).Value = Doctor
        arParameters(16) = New SqlParameter("@UserID", SqlDbType.NVarChar, 50)
        arParameters(16).Value = UserID
        arParameters(17) = New SqlParameter("@Locked", SqlDbType.Bit)
        arParameters(17).Value = Locked
        arParameters(18) = New SqlParameter("@Days", SqlDbType.Int)
        If Weeks = Nothing Then
            arParameters(18).Value = DBNull.Value
        Else
            arParameters(18).Value = Days
        End If
        arParameters(19) = New SqlParameter("@Comments", SqlDbType.NVarChar, 2000)
        arParameters(19).Value = Comments
        arParameters(20) = New SqlParameter("@TestNumber", SqlDbType.Int)
        If TestNumber = Nothing Then
            arParameters(20).Value = DBNull.Value
        Else
            arParameters(20).Value = TestNumber
        End If
        arParameters(21) = New SqlParameter("@FR", SqlDbType.NVarChar, 50)
        arParameters(21).Value = FR
        arParameters(22) = New SqlParameter("@Indications", SqlDbType.NVarChar, 100)
        arParameters(22).Value = Indications
        arParameters(23) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(23).Value = UpdatedBy
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAntenatalUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAntenatalUpdate", arParameters)
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
                            ByVal ChartID As Integer, _
                            ByVal TestDate As String, _
                            ByVal Weeks As Integer, _
                            ByVal Seen As Integer, _
                            ByVal SBP As Integer, _
                            ByVal DBP As Integer, _
                            ByVal Urinep As String, _
                            ByVal UrineLuc As String, _
                            ByVal FHRBL As String, _
                            ByVal NSTResults As String, _
                            ByVal AFI As Integer, _
                            ByVal BPP As String, _
                            ByVal PlacentaGrade As String, _
                            ByVal FAS As String, _
                            ByVal RN As String, _
                            ByVal Doctor As String, _
                            ByVal UserID As String, _
                            ByVal Locked As Short, _
                            ByVal Days As Integer, _
                            ByVal Comments As String, _
                            ByVal TestNumber As Integer, _
                            ByVal FR As String, _
                            ByVal Indications As String, _
                            ByVal CreatedBy As String) As Boolean

        Dim arParameters(24) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@IDTest", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(1).Value = ID
        arParameters(2) = New SqlParameter("@Date", SqlDbType.SmallDateTime)
        If TestDate Is Nothing Or TestDate = "" Then
            arParameters(2).Value = DBNull.Value
        Else
            arParameters(2).Value = TestDate
        End If
        arParameters(3) = New SqlParameter("@Weeks", SqlDbType.Int)
        If Weeks = Nothing Then
            arParameters(3).Value = DBNull.Value
        Else
            arParameters(3).Value = Weeks
        End If
        arParameters(4) = New SqlParameter("@Seen", SqlDbType.SmallInt)
        If Seen = Nothing Then
            arParameters(4).Value = DBNull.Value
        Else
            arParameters(4).Value = Seen
        End If
        arParameters(5) = New SqlParameter("@SBP", SqlDbType.Int)
        If SBP = Nothing Then
            arParameters(5).Value = DBNull.Value
        Else
            arParameters(5).Value = SBP
        End If
        arParameters(6) = New SqlParameter("@DBP", SqlDbType.Int)
        If DBP = Nothing Then
            arParameters(6).Value = DBNull.Value
        Else
            arParameters(6).Value = DBP
        End If
        arParameters(7) = New SqlParameter("@UrineP", SqlDbType.NVarChar, 50)
        arParameters(7).Value = Urinep
        arParameters(8) = New SqlParameter("@UrineLuc", SqlDbType.NVarChar, 50)
        arParameters(8).Value = UrineLuc
        arParameters(9) = New SqlParameter("@FHRBL", SqlDbType.NVarChar, 50)
        arParameters(9).Value = FHRBL
        arParameters(10) = New SqlParameter("@NSTResults", SqlDbType.NVarChar, 50)
        arParameters(10).Value = NSTResults
        arParameters(11) = New SqlParameter("@AFI", SqlDbType.Int)
        If AFI = Nothing Then
            arParameters(11).Value = DBNull.Value
        Else
            arParameters(11).Value = AFI
        End If
        arParameters(12) = New SqlParameter("@BPP", SqlDbType.NVarChar, 50)
        arParameters(12).Value = BPP
        arParameters(13) = New SqlParameter("@PlacentaGrade", SqlDbType.NVarChar, 50)
        arParameters(13).Value = PlacentaGrade
        arParameters(14) = New SqlParameter("@FAS", SqlDbType.NVarChar, 50)
        arParameters(14).Value = FAS
        arParameters(15) = New SqlParameter("@RN", SqlDbType.NVarChar, 50)
        arParameters(15).Value = RN
        arParameters(16) = New SqlParameter("@Doctor", SqlDbType.NVarChar, 50)
        arParameters(16).Value = Doctor
        arParameters(17) = New SqlParameter("@UserID", SqlDbType.NVarChar, 50)
        arParameters(17).Value = UserID
        arParameters(18) = New SqlParameter("@Locked", SqlDbType.Bit)
        arParameters(18).Value = Locked
        arParameters(19) = New SqlParameter("@Days", SqlDbType.Int)
        If Weeks = Nothing Then
            arParameters(19).Value = DBNull.Value
        Else
            arParameters(19).Value = Days
        End If
        arParameters(20) = New SqlParameter("@Comments", SqlDbType.NVarChar, 2000)
        arParameters(20).Value = Comments
        arParameters(21) = New SqlParameter("@TestNumber", SqlDbType.Int)
        If TestNumber = Nothing Then
            arParameters(21).Value = DBNull.Value
        Else
            arParameters(21).Value = TestNumber
        End If
        arParameters(22) = New SqlParameter("@FR", SqlDbType.NVarChar, 50)
        arParameters(22).Value = FR
        arParameters(23) = New SqlParameter("@Indications", SqlDbType.NVarChar, 100)
        arParameters(23).Value = Indications
        arParameters(24) = New SqlParameter("@CreatedBy", SqlDbType.NVarChar, 100)
        arParameters(24).Value = CreatedBy
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAntenatalInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAntenatalInsert", arParameters)
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
        arParameters(0) = New SqlParameter("@IDTest", SqlDbType.Int)
        arParameters(0).Value = ID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spAntenatalDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spAntenatalDelete", arParameters)
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


End Class 'dalAntenatal
