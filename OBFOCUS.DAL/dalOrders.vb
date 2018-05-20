
Option Explicit On 
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
'******************************************************************************
'*
'* Name:        dalOrders
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
Public Class dalOrders

#Region "Module level variables and enums"

    ' Public ENUM used to enumerate columns 
    Public Enum WdiagnosisFields
        fldID = 0
        fldOrderDate = 1
        fldOrderDescription = 2
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
    '* Name:        GetOrders
    '*
    '* Description: Returns all records in the [Labs] table according
    '*              to specified criteria.
    '*
    '*
    '* Returns:     DataReader containing the specified data. 
    '*
    '**************************************************************************
    Public Function GetOrders(ByVal ChartID As Integer) As SqlDataReader
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the parameters
        arParameters(0) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(0).Value = ChartID
        ' Call stored procedure and return the data
        Try
            If Me.Transaction Is Nothing Then
                Return SqlHelper.ExecuteReader(Globals.ConnectionString, CommandType.StoredProcedure, "spOrdersGet", arParameters)
            Else
                Return SqlHelper.ExecuteReader(Me.Transaction, CommandType.StoredProcedure, "spOrdersGet", arParameters)
            End If
        Catch ex As Exception
            ExceptionManager.Publish(ex)
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
    Public Function Update(ByVal OrderID As Integer, _
                        ByVal OrderDate As String, _
                        ByVal Description As String, _
                        ByVal ExaminerID As Integer, _
                        ByVal Weight As Integer, _
                        ByVal Allergies As String, _
                        ByVal Admitto As String, _
                        ByVal Vitals As String, _
                        ByVal GlucoseCheck As String, _
                        ByVal Activity As String, _
                        ByVal Nursing As String, _
                        ByVal Diet As String, _
                        ByVal IVType As String, _
                        ByVal Fetus As String, _
                        ByVal Meds As String, _
                        ByVal InsulinDrip As Short, _
                        ByVal InsulinScale As Short, _
                        ByVal HeparinIV As Short, _
                        ByVal LoadingDose As Integer, _
                        ByVal Labs As String, _
                        ByVal Tests As String, _
                        ByVal CallFor As String, _
                        ByVal Additional As String, _
                        ByVal ChartAllergies As String, _
                        ByVal ChartID As Integer, _
                        ByVal DelHospitalID As Integer, _
                        ByVal UpdatedBy As String) As Boolean

        Dim arParameters(26) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@OrderID", SqlDbType.Int)
        arParameters(0).Value = OrderID
        arParameters(1) = New SqlParameter("@OrderDate", SqlDbType.SmallDateTime)
        If OrderDate = "" Then
            arParameters(1).Value = System.DBNull.Value
        Else
            arParameters(1).Value = OrderDate
        End If
        arParameters(2) = New SqlParameter("@Description", SqlDbType.NVarChar, 200)
        arParameters(2).Value = Description
        arParameters(3) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(3).Value = ExaminerID
        arParameters(4) = New SqlParameter("@Weight", SqlDbType.Int)
        arParameters(4).Value = Weight
        arParameters(5) = New SqlParameter("@Allergies", SqlDbType.NVarChar, 150)
        arParameters(5).Value = Allergies
        arParameters(6) = New SqlParameter("@Admitto", SqlDbType.NVarChar, 50)
        arParameters(6).Value = Admitto
        arParameters(7) = New SqlParameter("@Vitals", SqlDbType.NVarChar, 255)
        arParameters(7).Value = Vitals
        arParameters(8) = New SqlParameter("@GlucoseCheck", SqlDbType.NVarChar, 200)
        arParameters(8).Value = GlucoseCheck
        arParameters(9) = New SqlParameter("@Activity", SqlDbType.NVarChar, 255)
        arParameters(9).Value = Activity
        arParameters(10) = New SqlParameter("@Nursing", SqlDbType.NVarChar, 255)
        arParameters(10).Value = Nursing
        arParameters(11) = New SqlParameter("@Diet", SqlDbType.NVarChar, 255)
        arParameters(11).Value = Diet
        arParameters(12) = New SqlParameter("@IVType", SqlDbType.NVarChar, 255)
        arParameters(12).Value = IVType
        arParameters(13) = New SqlParameter("@Fetus", SqlDbType.NVarChar, 255)
        arParameters(13).Value = Fetus
        arParameters(14) = New SqlParameter("@Meds", SqlDbType.NVarChar, 50)
        arParameters(14).Value = Meds
        arParameters(15) = New SqlParameter("@InsulinDrip", SqlDbType.SmallInt)
        arParameters(15).Value = InsulinDrip
        arParameters(16) = New SqlParameter("@InsulinScale", SqlDbType.SmallInt)
        arParameters(16).Value = InsulinScale
        arParameters(17) = New SqlParameter("@HeparinIV", SqlDbType.SmallInt)
        arParameters(17).Value = HeparinIV
        arParameters(18) = New SqlParameter("@LoadingDose", SqlDbType.Int)
        arParameters(18).Value = LoadingDose
        arParameters(19) = New SqlParameter("@Labs", SqlDbType.VarChar, 8000)
        arParameters(19).Value = Labs
        arParameters(20) = New SqlParameter("@Tests", SqlDbType.VarChar, 8000)
        arParameters(20).Value = Tests
        arParameters(21) = New SqlParameter("@CallFor", SqlDbType.VarChar, 8000)
        arParameters(21).Value = CallFor
        arParameters(22) = New SqlParameter("@Additional", SqlDbType.VarChar, 8000)
        arParameters(22).Value = Additional
        arParameters(23) = New SqlParameter("@ChartAllergies", SqlDbType.NVarChar, 255)
        arParameters(23).Value = ChartAllergies
        arParameters(24) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(24).Value = ChartID
        arParameters(25) = New SqlParameter("@DelHospitalID", SqlDbType.Int)
        arParameters(25).Value = DelHospitalID
        arParameters(26) = New SqlParameter("@UpdatedBy", SqlDbType.NVarChar, 50)
        arParameters(26).Value = UpdatedBy

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOrdersUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOrdersUpdate", arParameters)
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
    '* Name:        UpdateOrderSet
    '*
    '* Description: UpdateOrderSets a record identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function UpdateOrderSet(ByVal OrderID As Integer, _
                        ByVal Description As String, _
                        ByVal Admitto As String, _
                        ByVal Vitals As String, _
                        ByVal GlucoseCheck As String, _
                        ByVal Activity As String, _
                        ByVal Nursing As String, _
                        ByVal Diet As String, _
                        ByVal IVType As String, _
                        ByVal Fetus As String, _
                        ByVal Meds As String, _
                        ByVal InsulinDrip As Short, _
                        ByVal InsulinScale As Short, _
                        ByVal HeparinIV As Short, _
                        ByVal LoadingDose As Integer, _
                        ByVal Labs As String, _
                        ByVal Tests As String, _
                        ByVal CallFor As String, _
                        ByVal Additional As String) As Boolean

        Dim arParameters(18) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@OrderID", SqlDbType.Int)
        arParameters(0).Value = OrderID
        arParameters(1) = New SqlParameter("@Description", SqlDbType.NVarChar, 200)
        arParameters(1).Value = Description
        arParameters(2) = New SqlParameter("@Admitto", SqlDbType.NVarChar, 50)
        arParameters(2).Value = Admitto
        arParameters(3) = New SqlParameter("@Vitals", SqlDbType.NVarChar, 255)
        arParameters(3).Value = Vitals
        arParameters(4) = New SqlParameter("@GlucoseCheck", SqlDbType.NVarChar, 200)
        arParameters(4).Value = GlucoseCheck
        arParameters(5) = New SqlParameter("@Activity", SqlDbType.NVarChar, 255)
        arParameters(5).Value = Activity
        arParameters(6) = New SqlParameter("@Nursing", SqlDbType.NVarChar, 255)
        arParameters(6).Value = Nursing
        arParameters(7) = New SqlParameter("@Diet", SqlDbType.NVarChar, 255)
        arParameters(7).Value = Diet
        arParameters(8) = New SqlParameter("@IVType", SqlDbType.NVarChar, 255)
        arParameters(8).Value = IVType
        arParameters(9) = New SqlParameter("@Fetus", SqlDbType.NVarChar, 255)
        arParameters(9).Value = Fetus
        arParameters(10) = New SqlParameter("@Meds", SqlDbType.NVarChar, 50)
        arParameters(10).Value = Meds
        arParameters(11) = New SqlParameter("@InsulinDrip", SqlDbType.SmallInt)
        arParameters(11).Value = InsulinDrip
        arParameters(12) = New SqlParameter("@InsulinScale", SqlDbType.SmallInt)
        arParameters(12).Value = InsulinScale
        arParameters(13) = New SqlParameter("@HeparinIV", SqlDbType.SmallInt)
        arParameters(13).Value = HeparinIV
        arParameters(14) = New SqlParameter("@LoadingDose", SqlDbType.Int)
        arParameters(14).Value = LoadingDose
        arParameters(15) = New SqlParameter("@Labs", SqlDbType.VarChar, 8000)
        arParameters(15).Value = Labs
        arParameters(16) = New SqlParameter("@Tests", SqlDbType.VarChar, 8000)
        arParameters(16).Value = Tests
        arParameters(17) = New SqlParameter("@CallFor", SqlDbType.VarChar, 8000)
        arParameters(17).Value = CallFor
        arParameters(18) = New SqlParameter("@Additional", SqlDbType.VarChar, 8000)
        arParameters(18).Value = Additional

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOrderSetUpdate", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOrderSetUpdate", arParameters)
            End If
        Catch exception As Exception
            ExceptionManager.Publish(exception)
        End Try

        ' Return False if data was not UpdateOrderSetd.
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
                        ByVal OrderDate As String, _
                        ByVal Description As String, _
                        ByVal ExaminerID As Integer, _
                        ByVal Weight As Integer, _
                        ByVal Allergies As String, _
                        ByVal Admitto As String, _
                        ByVal Vitals As String, _
                        ByVal GlucoseCheck As String, _
                        ByVal Activity As String, _
                        ByVal Nursing As String, _
                        ByVal Diet As String, _
                        ByVal IVType As String, _
                        ByVal Fetus As String, _
                        ByVal Meds As String, _
                        ByVal InsulinDrip As Short, _
                        ByVal InsulinScale As Short, _
                        ByVal HeparinIV As Short, _
                        ByVal LoadingDose As Integer, _
                        ByVal Labs As String, _
                        ByVal Tests As String, _
                        ByVal CallFor As String, _
                        ByVal Additional As String, _
                        ByVal ChartAllergies As String, _
                        ByVal ChartID As Integer, _
                        ByVal DelHospitalID As Integer, _
                        ByVal UserID As String) As Boolean

        Dim arParameters(26) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@OrderID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@OrderDate", SqlDbType.SmallDateTime)
        If OrderDate = "" Then
            arParameters(1).Value = System.DBNull.Value
        Else
            arParameters(1).Value = OrderDate
        End If
        arParameters(2) = New SqlParameter("@Description", SqlDbType.NVarChar, 200)
        arParameters(2).Value = Description
        arParameters(3) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(3).Value = ExaminerID
        arParameters(4) = New SqlParameter("@Weight", SqlDbType.Int)
        arParameters(4).Value = Weight
        arParameters(5) = New SqlParameter("@Allergies", SqlDbType.NVarChar, 150)
        arParameters(5).Value = Allergies
        arParameters(6) = New SqlParameter("@Admitto", SqlDbType.NVarChar, 50)
        arParameters(6).Value = Admitto
        arParameters(7) = New SqlParameter("@Vitals", SqlDbType.NVarChar, 255)
        arParameters(7).Value = Vitals
        arParameters(8) = New SqlParameter("@GlucoseCheck", SqlDbType.NVarChar, 200)
        arParameters(8).Value = GlucoseCheck
        arParameters(9) = New SqlParameter("@Activity", SqlDbType.NVarChar, 255)
        arParameters(9).Value = Activity
        arParameters(10) = New SqlParameter("@Nursing", SqlDbType.NVarChar, 255)
        arParameters(10).Value = Nursing
        arParameters(11) = New SqlParameter("@Diet", SqlDbType.NVarChar, 255)
        arParameters(11).Value = Diet
        arParameters(12) = New SqlParameter("@IVType", SqlDbType.NVarChar, 255)
        arParameters(12).Value = IVType
        arParameters(13) = New SqlParameter("@Fetus", SqlDbType.NVarChar, 255)
        arParameters(13).Value = Fetus
        arParameters(14) = New SqlParameter("@Meds", SqlDbType.NVarChar, 50)
        arParameters(14).Value = Meds
        arParameters(15) = New SqlParameter("@InsulinDrip", SqlDbType.SmallInt)
        arParameters(15).Value = InsulinDrip
        arParameters(16) = New SqlParameter("@InsulinScale", SqlDbType.SmallInt)
        arParameters(16).Value = InsulinScale
        arParameters(17) = New SqlParameter("@HeparinIV", SqlDbType.SmallInt)
        arParameters(17).Value = HeparinIV
        arParameters(18) = New SqlParameter("@LoadingDose", SqlDbType.Int)
        arParameters(18).Value = LoadingDose
        arParameters(19) = New SqlParameter("@Labs", SqlDbType.VarChar, 8000)
        arParameters(19).Value = Labs
        arParameters(20) = New SqlParameter("@Tests", SqlDbType.VarChar, 8000)
        arParameters(20).Value = Tests
        arParameters(21) = New SqlParameter("@CallFor", SqlDbType.VarChar, 8000)
        arParameters(21).Value = CallFor
        arParameters(22) = New SqlParameter("@Additional", SqlDbType.VarChar, 8000)
        arParameters(22).Value = Additional
        arParameters(23) = New SqlParameter("@ChartAllergies", SqlDbType.NVarChar, 255)
        arParameters(23).Value = ChartAllergies
        arParameters(24) = New SqlParameter("@ChartID", SqlDbType.Int)
        arParameters(24).Value = ChartID
        arParameters(25) = New SqlParameter("@DelHospitalID", SqlDbType.Int)
        arParameters(25).Value = DelHospitalID
        arParameters(26) = New SqlParameter("@UserID", SqlDbType.NVarChar, 50)
        arParameters(26).Value = UserID
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOrdersInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOrdersInsert", arParameters)
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
    '* Name:        AddOrderSet
    '*
    '* Description: AddOrderSets a new record to the [PatientInfo] table.
    '*
    '*
    '* Returns:     Boolean indicating if record was AddOrderSeted or not. 
    '*              True (record AddOrderSeted); False (otherwise).
    '*
    '**************************************************************************
    Public Function AddOrderSet(ByRef ID As Integer, _
                        ByVal Description As String, _
                        ByVal Admitto As String, _
                        ByVal Vitals As String, _
                        ByVal GlucoseCheck As String, _
                        ByVal Activity As String, _
                        ByVal Nursing As String, _
                        ByVal Diet As String, _
                        ByVal IVType As String, _
                        ByVal Fetus As String, _
                        ByVal Meds As String, _
                        ByVal InsulinDrip As Short, _
                        ByVal InsulinScale As Short, _
                        ByVal HeparinIV As Short, _
                        ByVal LoadingDose As Integer, _
                        ByVal Labs As String, _
                        ByVal Tests As String, _
                        ByVal CallFor As String, _
                        ByVal Additional As String) As Boolean

        Dim arParameters(18) As SqlParameter         ' Array to hold stored procedure parameters
        Dim intRecordsAffected As Integer = 0

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@OrderID", SqlDbType.Int)
        arParameters(0).Value = ID
        arParameters(0).Direction = ParameterDirection.Output
        arParameters(1) = New SqlParameter("@Description", SqlDbType.NVarChar, 200)
        arParameters(1).Value = Description
        arParameters(2) = New SqlParameter("@Admitto", SqlDbType.NVarChar, 50)
        arParameters(2).Value = Admitto
        arParameters(3) = New SqlParameter("@Vitals", SqlDbType.NVarChar, 255)
        arParameters(3).Value = Vitals
        arParameters(4) = New SqlParameter("@GlucoseCheck", SqlDbType.NVarChar, 200)
        arParameters(4).Value = GlucoseCheck
        arParameters(5) = New SqlParameter("@Activity", SqlDbType.NVarChar, 255)
        arParameters(5).Value = Activity
        arParameters(6) = New SqlParameter("@Nursing", SqlDbType.NVarChar, 255)
        arParameters(6).Value = Nursing
        arParameters(7) = New SqlParameter("@Diet", SqlDbType.NVarChar, 255)
        arParameters(7).Value = Diet
        arParameters(8) = New SqlParameter("@IVType", SqlDbType.NVarChar, 255)
        arParameters(8).Value = IVType
        arParameters(9) = New SqlParameter("@Fetus", SqlDbType.NVarChar, 255)
        arParameters(9).Value = Fetus
        arParameters(10) = New SqlParameter("@Meds", SqlDbType.NVarChar, 50)
        arParameters(10).Value = Meds
        arParameters(11) = New SqlParameter("@InsulinDrip", SqlDbType.SmallInt)
        arParameters(11).Value = InsulinDrip
        arParameters(12) = New SqlParameter("@InsulinScale", SqlDbType.SmallInt)
        arParameters(12).Value = InsulinScale
        arParameters(13) = New SqlParameter("@HeparinIV", SqlDbType.SmallInt)
        arParameters(13).Value = HeparinIV
        arParameters(14) = New SqlParameter("@LoadingDose", SqlDbType.Int)
        arParameters(14).Value = LoadingDose
        arParameters(15) = New SqlParameter("@Labs", SqlDbType.VarChar, 8000)
        arParameters(15).Value = Labs
        arParameters(16) = New SqlParameter("@Tests", SqlDbType.VarChar, 8000)
        arParameters(16).Value = Tests
        arParameters(17) = New SqlParameter("@CallFor", SqlDbType.VarChar, 8000)
        arParameters(17).Value = CallFor
        arParameters(18) = New SqlParameter("@Additional", SqlDbType.VarChar, 8000)
        arParameters(18).Value = Additional
        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOrderSetInsert", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOrderSetInsert", arParameters)
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
    Public Function Delete(ByVal OrderID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@OrderID", SqlDbType.Int)
        arParameters(0).Value = OrderID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOrdersDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOrdersDelete", arParameters)
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
    '* Name:        DeleteOrderSet
    '*
    '* Description: DeleteOrderSets a record from the [PatientInfo] table identified by a key.
    '*
    '* Parameters:  ID - Key of record that we want to DeleteOrderSet
    '*
    '* Returns:     Boolean indicating if record was DeleteOrderSetd or not. 
    '*              True (record found and DeleteOrderSetd); False (otherwise).
    '*
    '**************************************************************************
    Public Function DeleteOrderSet(ByVal OrderID As Integer) As Boolean

        Dim intRecordsAffected As Integer = 0
        Dim arParameters(0) As SqlParameter         ' Array to hold stored procedure parameters
        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@OrderID", SqlDbType.Int)
        arParameters(0).Value = OrderID

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOrderSetDelete", arParameters)
            Else
                intRecordsAffected = SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOrderSetDelete", arParameters)
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
    Public Function GetByKey(ByVal OrderID As Integer, _
                        ByRef OrderDate As String, _
                        ByRef Description As String, _
                        ByRef ExaminerID As Integer, _
                        ByRef Weight As Integer, _
                        ByRef Allergies As String, _
                        ByRef Admitto As String, _
                        ByRef Vitals As String, _
                        ByRef GlucoseCheck As String, _
                        ByRef Activity As String, _
                        ByRef Nursing As String, _
                        ByRef Diet As String, _
                        ByRef IVType As String, _
                        ByRef Fetus As String, _
                        ByRef Meds As String, _
                        ByRef InsulinDrip As Short, _
                        ByRef InsulinScale As Short, _
                        ByRef HeparinIV As Short, _
                        ByRef LoadingDose As Integer, _
                        ByRef Labs As String, _
                        ByRef Tests As String, _
                        ByRef CallFor As String, _
                        ByRef Additional As String) As Boolean

        Dim arParameters(22) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@OrderID", SqlDbType.Int)
        arParameters(0).Value = OrderID
        arParameters(1) = New SqlParameter("@OrderDate", SqlDbType.SmallDateTime)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@Description", SqlDbType.NVarChar, 200)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@ExaminerID", SqlDbType.Int)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@Weight", SqlDbType.Int)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@Allergies", SqlDbType.NVarChar, 150)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@Admitto", SqlDbType.NVarChar, 50)
        arParameters(6).Direction = ParameterDirection.Output
        arParameters(7) = New SqlParameter("@Vitals", SqlDbType.NVarChar, 255)
        arParameters(7).Direction = ParameterDirection.Output
        arParameters(8) = New SqlParameter("@GlucoseCheck", SqlDbType.NVarChar, 200)
        arParameters(8).Direction = ParameterDirection.Output
        arParameters(9) = New SqlParameter("@Activity", SqlDbType.NVarChar, 255)
        arParameters(9).Direction = ParameterDirection.Output
        arParameters(10) = New SqlParameter("@Nursing", SqlDbType.NVarChar, 255)
        arParameters(10).Direction = ParameterDirection.Output
        arParameters(11) = New SqlParameter("@Diet", SqlDbType.NVarChar, 255)
        arParameters(11).Direction = ParameterDirection.Output
        arParameters(12) = New SqlParameter("@IVType", SqlDbType.NVarChar, 255)
        arParameters(12).Direction = ParameterDirection.Output
        arParameters(13) = New SqlParameter("@Fetus", SqlDbType.NVarChar, 255)
        arParameters(13).Direction = ParameterDirection.Output
        arParameters(14) = New SqlParameter("@Meds", SqlDbType.NVarChar, 50)
        arParameters(14).Direction = ParameterDirection.Output
        arParameters(15) = New SqlParameter("@InsulinDrip", SqlDbType.SmallInt)
        arParameters(15).Direction = ParameterDirection.Output
        arParameters(16) = New SqlParameter("@InsulinScale", SqlDbType.SmallInt)
        arParameters(16).Direction = ParameterDirection.Output
        arParameters(17) = New SqlParameter("@HeparinIV", SqlDbType.SmallInt)
        arParameters(17).Direction = ParameterDirection.Output
        arParameters(18) = New SqlParameter("@LoadingDose", SqlDbType.Int)
        arParameters(18).Direction = ParameterDirection.Output
        arParameters(19) = New SqlParameter("@Labs", SqlDbType.VarChar, 8000)
        arParameters(19).Direction = ParameterDirection.Output
        arParameters(20) = New SqlParameter("@Tests", SqlDbType.VarChar, 8000)
        arParameters(20).Direction = ParameterDirection.Output
        arParameters(21) = New SqlParameter("@CallFor", SqlDbType.VarChar, 8000)
        arParameters(21).Direction = ParameterDirection.Output
        arParameters(22) = New SqlParameter("@Additional", SqlDbType.VarChar, 8000)
        arParameters(22).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOrdersGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOrdersGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(0).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            OrderDate = ProcessNull.GetString(arParameters(1).Value)
            Description = ProcessNull.GetString(arParameters(2).Value)
            ExaminerID = ProcessNull.GetInt32(arParameters(3).Value)
            Weight = ProcessNull.GetInt32(arParameters(4).Value)
            Allergies = ProcessNull.GetString(arParameters(5).Value)
            Admitto = ProcessNull.GetString(arParameters(6).Value)
            Vitals = ProcessNull.GetString(arParameters(7).Value)
            GlucoseCheck = ProcessNull.GetString(arParameters(8).Value)
            Activity = ProcessNull.GetString(arParameters(9).Value)
            Nursing = ProcessNull.GetString(arParameters(10).Value)
            Diet = ProcessNull.GetString(arParameters(11).Value)
            IVType = ProcessNull.GetString(arParameters(12).Value)
            Fetus = ProcessNull.GetString(arParameters(13).Value)
            Meds = ProcessNull.GetString(arParameters(14).Value)
            InsulinDrip = ProcessNull.GetInt16(arParameters(15).Value)
            InsulinScale = ProcessNull.GetInt16(arParameters(16).Value)
            HeparinIV = ProcessNull.GetInt16(arParameters(17).Value)
            LoadingDose = ProcessNull.GetInt32(arParameters(18).Value)
            Labs = ProcessNull.GetString(arParameters(19).Value)
            Tests = ProcessNull.GetString(arParameters(20).Value)
            CallFor = ProcessNull.GetString(arParameters(21).Value)
            Additional = ProcessNull.GetString(arParameters(22).Value)
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function
    '**************************************************************************
    '*  
    '* Name:        GetOrderSetByKey
    '*
    '* Description: Gets all the values of a record in the [PatientInfo]and Chart tables
    '*              identified by a key.
    '*
    '*
    '* Returns:     Boolean indicating if record was found or not. 
    '*              True (record found); False (otherwise).
    '*
    '**************************************************************************
    Public Function GetOrderSetByKey(ByVal OrderID As Integer, _
                        ByRef Description As String, _
                        ByRef Admitto As String, _
                        ByRef Vitals As String, _
                        ByRef GlucoseCheck As String, _
                        ByRef Activity As String, _
                        ByRef Nursing As String, _
                        ByRef Diet As String, _
                        ByRef IVType As String, _
                        ByRef Fetus As String, _
                        ByRef Meds As String, _
                        ByRef InsulinDrip As Short, _
                        ByRef InsulinScale As Short, _
                        ByRef HeparinIV As Short, _
                        ByRef LoadingDose As Integer, _
                        ByRef Labs As String, _
                        ByRef Tests As String, _
                        ByRef CallFor As String, _
                        ByRef Additional As String) As Boolean

        Dim arParameters(18) As SqlParameter         ' Array to hold stored procedure parameters

        ' Set the stored procedure parameters
        arParameters(0) = New SqlParameter("@OrderID", SqlDbType.Int)
        arParameters(0).Value = OrderID
        arParameters(1) = New SqlParameter("@Description", SqlDbType.NVarChar, 200)
        arParameters(1).Direction = ParameterDirection.Output
        arParameters(2) = New SqlParameter("@Admitto", SqlDbType.NVarChar, 50)
        arParameters(2).Direction = ParameterDirection.Output
        arParameters(3) = New SqlParameter("@Vitals", SqlDbType.NVarChar, 255)
        arParameters(3).Direction = ParameterDirection.Output
        arParameters(4) = New SqlParameter("@GlucoseCheck", SqlDbType.NVarChar, 200)
        arParameters(4).Direction = ParameterDirection.Output
        arParameters(5) = New SqlParameter("@Activity", SqlDbType.NVarChar, 255)
        arParameters(5).Direction = ParameterDirection.Output
        arParameters(6) = New SqlParameter("@Nursing", SqlDbType.NVarChar, 255)
        arParameters(6).Direction = ParameterDirection.Output
        arParameters(7) = New SqlParameter("@Diet", SqlDbType.NVarChar, 255)
        arParameters(7).Direction = ParameterDirection.Output
        arParameters(8) = New SqlParameter("@IVType", SqlDbType.NVarChar, 255)
        arParameters(8).Direction = ParameterDirection.Output
        arParameters(9) = New SqlParameter("@Fetus", SqlDbType.NVarChar, 255)
        arParameters(9).Direction = ParameterDirection.Output
        arParameters(10) = New SqlParameter("@Meds", SqlDbType.NVarChar, 50)
        arParameters(10).Direction = ParameterDirection.Output
        arParameters(11) = New SqlParameter("@InsulinDrip", SqlDbType.SmallInt)
        arParameters(11).Direction = ParameterDirection.Output
        arParameters(12) = New SqlParameter("@InsulinScale", SqlDbType.SmallInt)
        arParameters(12).Direction = ParameterDirection.Output
        arParameters(13) = New SqlParameter("@HeparinIV", SqlDbType.SmallInt)
        arParameters(13).Direction = ParameterDirection.Output
        arParameters(14) = New SqlParameter("@LoadingDose", SqlDbType.Int)
        arParameters(14).Direction = ParameterDirection.Output
        arParameters(15) = New SqlParameter("@Labs", SqlDbType.VarChar, 8000)
        arParameters(15).Direction = ParameterDirection.Output
        arParameters(16) = New SqlParameter("@Tests", SqlDbType.VarChar, 8000)
        arParameters(16).Direction = ParameterDirection.Output
        arParameters(17) = New SqlParameter("@CallFor", SqlDbType.VarChar, 8000)
        arParameters(17).Direction = ParameterDirection.Output
        arParameters(18) = New SqlParameter("@Additional", SqlDbType.VarChar, 8000)
        arParameters(18).Direction = ParameterDirection.Output

        ' Call stored procedure
        Try
            If Me.Transaction Is Nothing Then
                SqlHelper.ExecuteNonQuery(Globals.ConnectionString, CommandType.StoredProcedure, "spOrderSetGetByKey", arParameters)
            Else
                SqlHelper.ExecuteNonQuery(Me.Transaction, CommandType.StoredProcedure, "spOrderSetGetByKey", arParameters)
            End If


            ' Return False if data was not found.
            If arParameters(0).Value Is DBNull.Value Then Return False

            ' Return True if data was found. Also populate output (ByRef) parameters.
            Description = ProcessNull.GetString(arParameters(1).Value)
            Admitto = ProcessNull.GetString(arParameters(2).Value)
            Vitals = ProcessNull.GetString(arParameters(3).Value)
            GlucoseCheck = ProcessNull.GetString(arParameters(4).Value)
            Activity = ProcessNull.GetString(arParameters(5).Value)
            Nursing = ProcessNull.GetString(arParameters(6).Value)
            Diet = ProcessNull.GetString(arParameters(7).Value)
            IVType = ProcessNull.GetString(arParameters(8).Value)
            Fetus = ProcessNull.GetString(arParameters(9).Value)
            Meds = ProcessNull.GetString(arParameters(10).Value)
            InsulinDrip = ProcessNull.GetInt16(arParameters(11).Value)
            InsulinScale = ProcessNull.GetInt16(arParameters(12).Value)
            HeparinIV = ProcessNull.GetInt16(arParameters(13).Value)
            LoadingDose = ProcessNull.GetInt32(arParameters(14).Value)
            Labs = ProcessNull.GetString(arParameters(15).Value)
            Tests = ProcessNull.GetString(arParameters(16).Value)
            CallFor = ProcessNull.GetString(arParameters(17).Value)
            Additional = ProcessNull.GetString(arParameters(18).Value)
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try


    End Function

#End Region


End Class 'dalOrders
