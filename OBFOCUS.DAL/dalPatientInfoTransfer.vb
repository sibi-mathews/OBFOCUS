
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalPatientInfoTransfer
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
Public Class dalPatientInfoTransfer

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



#Region "Main procedures - Get, Add, Update & Delete"
    '**************************************************************************
    '*  
    '* Name:        GetPatientInfoTransfer
    '*
    '* Description: Returns all records in the [PatientInfoTransfers] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetPatientInfoTransfer() As SqlDataReader

        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoTransferGet")
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoTransferGet")
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
    Public Function GetByKey(ByRef PatientID As Integer) As Boolean
        ' Set the stored procedure parameters
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(0).Value = PatientID
        arParameters(0).Direction = ParameterDirection.InputOutput
       
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoTransferGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoTransferGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(0).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            PatientID = ProcessNull.GetInt32(arParameters(0).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        GetCommOptIn
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
    Public Function GetCommOptIn(ByVal PatientID As Integer, ByVal Mode As String) As Boolean
        ' Set the stored procedure parameters
        Dim arParameters(2) As SqlParameter         ' Array to hold stored procedure parameters
        Dim OptInValue As Boolean = Nothing

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(0).Value = PatientID
        arParameters(1) = New SqlParameter("@Mode", SqlDbType.NVarChar, 10)
        arParameters(1).Value = Mode
        arParameters(2) = New SqlParameter("@OptInValue", SqlDbType.Bit)
        arParameters(2).Direction = ParameterDirection.Output
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientnfoCommGet", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPatientnfoCommGet", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(2).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            OptInValue = ProcessNull.GetBoolean(arParameters(2).Value)
            Return OptInValue

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        GetPatientInfoByKey
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
    Public Function GetPatientInfoByKey(ByVal PatientID As Integer, _
                                        ByRef MedicalRecord As String, _
                                        ByRef PatientLast As String, _
                                        ByRef PatientFirst As String, _
                                        ByRef SocialSecurity As String, _
                                        ByRef DOB As String, _
                                        ByRef Race As String, _
                                        ByRef Language As String, _
                                        ByRef Type As String, _
                                        ByRef RH As String, _
                                        ByRef PlaceOfBirth As String, _
                                        ByRef DriverLic As String, _
                                        ByRef Religion As String, _
                                        ByRef FaxOptIn As Short, _
                                        ByRef EmailoptIn As Short, _
                                        ByRef MailOptIn As Short, _
                                        ByRef OtherOptIn As Short) As Boolean
        ' Set the stored procedure parameters
        Dim arParameters(16) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(0).Value = PatientID
        arParameters(1) = New SqlParameter("@medicalrecord", SqlDbType.NVarChar, 50)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@patientlast", SqlDbType.NVarChar, 50)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@patientfirst", SqlDbType.NVarChar, 50)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@socialsecurity", SqlDbType.NVarChar, 50)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@dob", SqlDbType.SmallDateTime)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@race", SqlDbType.NVarChar, 50)
        arParameters(6).Direction = ParameterDirection.Output
        arParameters(7) = New SqlParameter("@language", SqlDbType.NVarChar, 50)
        arParameters(7).Direction = ParameterDirection.Output
        arParameters(8) = New SqlParameter("@type", SqlDbType.NVarChar, 50)
        arParameters(8).Direction = ParameterDirection.Output
        arParameters(9) = New SqlParameter("@rh", SqlDbType.NVarChar, 50)
        arParameters(9).Direction = ParameterDirection.Output
        arParameters(10) = New SqlParameter("@placeofbirth", SqlDbType.NVarChar, 255)
        arParameters(10).Direction = ParameterDirection.Output
        arParameters(11) = New SqlParameter("@driverlic", SqlDbType.NVarChar, 255)
        arParameters(11).Direction = ParameterDirection.Output
        arParameters(12) = New SqlParameter("@religion", SqlDbType.NVarChar, 255)
        arParameters(12).Direction = ParameterDirection.Output
        arParameters(13) = New SqlParameter("@FaxOptIn", SqlDbType.Bit)
        arParameters(13).Direction = ParameterDirection.Output
        arParameters(14) = New SqlParameter("@EmailOptIn", SqlDbType.Bit)
        arParameters(14).Direction = ParameterDirection.Output
        arParameters(15) = New SqlParameter("@MailOptIn", SqlDbType.Bit)
        arParameters(15).Direction = ParameterDirection.Output
        arParameters(16) = New SqlParameter("@OtherOptIn", SqlDbType.Bit)
        arParameters(16).Direction = ParameterDirection.Output



        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(0).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            MedicalRecord = ProcessNull.GetString(arParameters(1).Value)
            PatientLast = ProcessNull.GetString(arParameters(2).Value)
            PatientFirst = ProcessNull.GetString(arParameters(3).Value)
            SocialSecurity = ProcessNull.GetString(arParameters(4).Value)
            DOB = ProcessNull.GetString(arParameters(5).Value)
            Race = ProcessNull.GetString(arParameters(6).Value)
            Language = ProcessNull.GetString(arParameters(7).Value)
            Type = ProcessNull.GetString(arParameters(8).Value)
            RH = ProcessNull.GetString(arParameters(9).Value)
            PlaceOfBirth = ProcessNull.GetString(arParameters(10).Value)
            DriverLic = ProcessNull.GetString(arParameters(11).Value)
            Religion = ProcessNull.GetString(arParameters(12).Value)
            FaxOptIn = ProcessNull.GetInt16(arParameters(13).Value)
            EmailoptIn = ProcessNull.GetInt16(arParameters(14).Value)
            MailOptIn = ProcessNull.GetInt16(arParameters(15).Value)
            OtherOptIn = ProcessNull.GetInt16(arParameters(16).Value)
            Return True

        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Allergy:        Delete
    '*
    '* Class: Deletes a record from the [PatientInfo] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to delete
    '*
    '* Returns:     Boolean indicating if record was deleted or not. 
    '*              True (record found and deleted); False (otherwise).
    '*
    '**************************************************************************
    Public Function Delete(ByVal PatientID As Integer) As Boolean
        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters


        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(0).Value = PatientID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoTransferDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoTransferDelete", arParameters)
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
    '* Name:        Update
    '*
    '* Description: Updates a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was updated or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Update(ByVal PatientID As Integer, _
                            ByVal PatientLast As String, _
                            ByVal PatientFirst As String, _
                            ByVal SocialSecurity As String, _
                            ByVal DOB As String, _
                            ByVal PlaceOfBirth As String, _
                            ByVal Race As String, _
                            ByVal Language As String, _
                            ByVal DriverLic As String, _
                            ByVal Religion As String, _
                            ByVal MaritalStatus As String, _
                            ByVal AKA As String, _
                            ByVal SpecialNeeds As String, _
                            ByVal Address1 As String, _
                            ByVal City As String, _
                            ByVal State As String, _
                            ByVal Zip As String, _
                            ByVal Phone As String, _
                            ByVal CellPhone As String, _
                            ByVal Pager As String, _
                            ByVal EmployedAs As String, _
                            ByVal Employer As String, _
                            ByVal EAddress As String, _
                            ByVal ECity As String, _
                            ByVal EState As String, _
                            ByVal EZip As String, _
                            ByVal EPhone As String, _
                            ByVal Email As String, _
                            ByVal EExt As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(28) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(0).Value = PatientID
        arParameters(1) = New SqlParameter("@PatientLast", SqlDbType.NVarChar, 255)
        arParameters(1).Value = PatientLast
        arParameters(2) = New SqlParameter("@PatientFirst", SqlDbType.NVarChar, 255)
        arParameters(2).Value = PatientFirst
        arParameters(3) = New SqlParameter("@SocialSecurity", SqlDbType.NVarChar, 255)
        arParameters(3).Value = SocialSecurity
        arParameters(4) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        arParameters(4).Value = DOB
        arParameters(5) = New SqlParameter("@PlaceOfBirth", SqlDbType.NVarChar, 255)
        arParameters(5).Value = PlaceOfBirth
        arParameters(6) = New SqlParameter("@Race", SqlDbType.NVarChar, 255)
        arParameters(6).Value = Race
        arParameters(7) = New SqlParameter("@Language", SqlDbType.NVarChar, 255)
        arParameters(7).Value = Language
        arParameters(8) = New SqlParameter("@DriverLic", SqlDbType.NVarChar, 255)
        arParameters(8).Value = DriverLic
        arParameters(9) = New SqlParameter("@Religion", SqlDbType.NVarChar, 255)
        arParameters(9).Value = Religion
        arParameters(10) = New SqlParameter("@MaritalStatus", SqlDbType.NVarChar, 255)
        arParameters(10).Value = MaritalStatus
        arParameters(11) = New SqlParameter("@AKA", SqlDbType.NVarChar, 255)
        arParameters(11).Value = AKA
        arParameters(12) = New SqlParameter("@SpecialNeeds", SqlDbType.NVarChar, 255)
        arParameters(12).Value = SpecialNeeds
        arParameters(13) = New SqlParameter("@Address1", SqlDbType.NVarChar, 255)
        arParameters(13).Value = Address1
        arParameters(14) = New SqlParameter("@City", SqlDbType.NVarChar, 255)
        arParameters(14).Value = City
        arParameters(15) = New SqlParameter("@State", SqlDbType.NVarChar, 255)
        arParameters(15).Value = State
        arParameters(16) = New SqlParameter("@Zip", SqlDbType.NVarChar, 255)
        arParameters(16).Value = Zip
        arParameters(17) = New SqlParameter("@Phone", SqlDbType.NVarChar, 255)
        arParameters(17).Value = Phone
        arParameters(18) = New SqlParameter("@CellPhone", SqlDbType.NVarChar, 255)
        arParameters(18).Value = CellPhone
        arParameters(19) = New SqlParameter("@Pager", SqlDbType.NVarChar, 255)
        arParameters(19).Value = Pager
        arParameters(20) = New SqlParameter("@EmployedAs", SqlDbType.NVarChar, 255)
        arParameters(20).Value = EmployedAs
        arParameters(21) = New SqlParameter("@Employer", SqlDbType.NVarChar, 255)
        arParameters(21).Value = Employer
        arParameters(22) = New SqlParameter("@EAddress", SqlDbType.NVarChar, 255)
        arParameters(22).Value = EAddress
        arParameters(23) = New SqlParameter("@ECity", SqlDbType.NVarChar, 255)
        arParameters(23).Value = ECity
        arParameters(24) = New SqlParameter("@EState", SqlDbType.NVarChar, 255)
        arParameters(24).Value = EState
        arParameters(25) = New SqlParameter("@EZip", SqlDbType.NVarChar, 255)
        arParameters(25).Value = EZip
        arParameters(26) = New SqlParameter("@EPhone", SqlDbType.NVarChar, 255)
        arParameters(26).Value = EPhone
        arParameters(27) = New SqlParameter("@Email", SqlDbType.NVarChar, 255)
        arParameters(27).Value = Email
        arParameters(28) = New SqlParameter("@EExt", SqlDbType.NVarChar, 255)
        arParameters(28).Value = EExt


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoTransferUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoTransferUpdate", arParameters)
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
    '* Name:        UpdatePatientInfo
    '*
    '* Description: Updates a record in the PatientInfo table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was UpdatePatientInfod or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdatePatientInfo(ByVal PatientID As Integer, _
                            ByVal MedicalRecord As String, _
                            ByVal PatientLast As String, _
                            ByVal PatientFirst As String, _
                            ByVal SocialSecurity As String, _
                            ByVal DOB As String, _
                            ByVal PlaceOfBirth As String, _
                            ByVal Race As String, _
                            ByVal Language As String, _
                            ByVal Type As String, _
                            ByVal RH As String, _
                            ByVal DriverLic As String, _
                            ByVal Religion As String, _
                            ByVal OldPatientID As Integer, _
                            ByVal PatientAutoNum As Integer, _
                            ByVal UpdatePatientID As Integer, _
                            ByVal FaxOptIn As Short, _
                            ByVal EmailOptIn As Short, _
                            ByVal MailOptIn As Short, _
                            ByVal OtherOptIn As Short) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(20) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(0).Value = PatientID
        arParameters(1) = New SqlParameter("@MedicalRecord", SqlDbType.NVarChar, 50)
        arParameters(1).Value = MedicalRecord
        arParameters(2) = New SqlParameter("@PatientLast", SqlDbType.NVarChar, 50)
        arParameters(2).Value = PatientLast
        arParameters(3) = New SqlParameter("@PatientFirst", SqlDbType.NVarChar, 50)
        arParameters(3).Value = PatientFirst
        arParameters(4) = New SqlParameter("@SocialSecurity", SqlDbType.NVarChar, 50)
        arParameters(4).Value = SocialSecurity
        arParameters(5) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        arParameters(5).Value = DOB
        arParameters(6) = New SqlParameter("@Race", SqlDbType.NVarChar, 50)
        arParameters(6).Value = Race
        arParameters(7) = New SqlParameter("@Language", SqlDbType.NVarChar, 50)
        arParameters(7).Value = Language
        arParameters(8) = New SqlParameter("@Type", SqlDbType.NVarChar, 50)
        arParameters(8).Value = Type
        arParameters(9) = New SqlParameter("@RH", SqlDbType.NVarChar, 50)
        arParameters(9).Value = RH
        arParameters(10) = New SqlParameter("@DateCreated", SqlDbType.SmallDateTime)
        arParameters(10).Value = Now()
        arParameters(11) = New SqlParameter("@PlaceOfBirth", SqlDbType.NVarChar, 255)
        arParameters(11).Value = PlaceOfBirth
        arParameters(12) = New SqlParameter("@DriverLic", SqlDbType.NVarChar, 255)
        arParameters(12).Value = DriverLic
        arParameters(13) = New SqlParameter("@Religion", SqlDbType.NVarChar, 255)
        arParameters(13).Value = Religion
        arParameters(14) = New SqlParameter("@OldPatientID", SqlDbType.Int)
        arParameters(14).Value = OldPatientID
        arParameters(15) = New SqlParameter("@PatientAutoNum", SqlDbType.Bit)
        arParameters(15).Value = PatientAutoNum
        arParameters(16) = New SqlParameter("@UpdatePatientID", SqlDbType.Bit)
        arParameters(16).Value = UpdatePatientID
        arParameters(17) = New SqlParameter("@FaxOptIn", SqlDbType.Bit)
        arParameters(17).Value = FaxOptIn
        arParameters(18) = New SqlParameter("@EmailOptIn", SqlDbType.Bit)
        arParameters(18).Value = EmailOptIn
        arParameters(19) = New SqlParameter("@mailOptIn", SqlDbType.Bit)
        arParameters(19).Value = MailOptIn
        arParameters(20) = New SqlParameter("@OtherOptIn", SqlDbType.Bit)
        arParameters(20).Value = OtherOptIn
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoUpdate", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not UpdatePatientInfod.
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
    '* Description: Adds a record in the Chart table identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was Addd or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function Add(ByVal PatientID As Integer, _
                            ByVal PatientLast As String, _
                            ByVal PatientFirst As String, _
                            ByVal SocialSecurity As String, _
                            ByVal DOB As String, _
                            ByVal PlaceOfBirth As String, _
                            ByVal Race As String, _
                            ByVal Language As String, _
                            ByVal DriverLic As String, _
                            ByVal Religion As String, _
                            ByVal MaritalStatus As String, _
                            ByVal AKA As String, _
                            ByVal SpecialNeeds As String, _
                            ByVal Address1 As String, _
                            ByVal City As String, _
                            ByVal State As String, _
                            ByVal Zip As String, _
                            ByVal Phone As String, _
                            ByVal CellPhone As String, _
                            ByVal Pager As String, _
                            ByVal EmployedAs As String, _
                            ByVal Employer As String, _
                            ByVal EAddress As String, _
                            ByVal ECity As String, _
                            ByVal EState As String, _
                            ByVal EZip As String, _
                            ByVal EPhone As String, _
                            ByVal Email As String, _
                            ByVal EExt As String) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(28) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@PatientID", SqlDbType.Int)
        arParameters(0).Value = PatientID
        arParameters(1) = New SqlParameter("@PatientLast", SqlDbType.NVarChar, 255)
        arParameters(1).Value = PatientLast
        arParameters(2) = New SqlParameter("@PatientFirst", SqlDbType.NVarChar, 255)
        arParameters(2).Value = PatientFirst
        arParameters(3) = New SqlParameter("@SocialSecurity", SqlDbType.NVarChar, 255)
        arParameters(3).Value = SocialSecurity
        arParameters(4) = New SqlParameter("@DOB", SqlDbType.SmallDateTime)
        arParameters(4).Value = DOB
        arParameters(5) = New SqlParameter("@PlaceOfBirth", SqlDbType.NVarChar, 255)
        arParameters(5).Value = PlaceOfBirth
        arParameters(6) = New SqlParameter("@Race", SqlDbType.NVarChar, 255)
        arParameters(6).Value = Race
        arParameters(7) = New SqlParameter("@Language", SqlDbType.NVarChar, 255)
        arParameters(7).Value = Language
        arParameters(8) = New SqlParameter("@DriverLic", SqlDbType.NVarChar, 255)
        arParameters(8).Value = DriverLic
        arParameters(9) = New SqlParameter("@Religion", SqlDbType.NVarChar, 255)
        arParameters(9).Value = Religion
        arParameters(10) = New SqlParameter("@MaritalStatus", SqlDbType.NVarChar, 255)
        arParameters(10).Value = MaritalStatus
        arParameters(11) = New SqlParameter("@AKA", SqlDbType.NVarChar, 255)
        arParameters(11).Value = AKA
        arParameters(12) = New SqlParameter("@SpecialNeeds", SqlDbType.NVarChar, 255)
        arParameters(12).Value = SpecialNeeds
        arParameters(13) = New SqlParameter("@Address1", SqlDbType.NVarChar, 255)
        arParameters(13).Value = Address1
        arParameters(14) = New SqlParameter("@City", SqlDbType.NVarChar, 255)
        arParameters(14).Value = City
        arParameters(15) = New SqlParameter("@State", SqlDbType.NVarChar, 255)
        arParameters(15).Value = State
        arParameters(16) = New SqlParameter("@Zip", SqlDbType.NVarChar, 255)
        arParameters(16).Value = Zip
        arParameters(17) = New SqlParameter("@Phone", SqlDbType.NVarChar, 255)
        arParameters(17).Value = Phone
        arParameters(18) = New SqlParameter("@CellPhone", SqlDbType.NVarChar, 255)
        arParameters(18).Value = CellPhone
        arParameters(19) = New SqlParameter("@Pager", SqlDbType.NVarChar, 255)
        arParameters(19).Value = Pager
        arParameters(20) = New SqlParameter("@EmployedAs", SqlDbType.NVarChar, 255)
        arParameters(20).Value = EmployedAs
        arParameters(21) = New SqlParameter("@Employer", SqlDbType.NVarChar, 255)
        arParameters(21).Value = Employer
        arParameters(22) = New SqlParameter("@EAddress", SqlDbType.NVarChar, 255)
        arParameters(22).Value = EAddress
        arParameters(23) = New SqlParameter("@ECity", SqlDbType.NVarChar, 255)
        arParameters(23).Value = ECity
        arParameters(24) = New SqlParameter("@EState", SqlDbType.NVarChar, 255)
        arParameters(24).Value = EState
        arParameters(25) = New SqlParameter("@EZip", SqlDbType.NVarChar, 255)
        arParameters(25).Value = EZip
        arParameters(26) = New SqlParameter("@EPhone", SqlDbType.NVarChar, 255)
        arParameters(26).Value = EPhone
        arParameters(27) = New SqlParameter("@Email", SqlDbType.NVarChar, 255)
        arParameters(27).Value = Email
        arParameters(28) = New SqlParameter("@EExt", SqlDbType.NVarChar, 255)
        arParameters(28).Value = EExt


        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spPatientInfoTransferInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spPatientInfoTransferInsert", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not Addd.
        If intRecordsAffected = 0 Then
            Return False
        Else

            Return True
        End If

    End Function



#End Region


End Class 'dalPatientInfoTransfer
