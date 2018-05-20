Option Explicit On 
Option Strict On
Imports System.Data.SqlClient

'******************************************************************************
'*
'* Name:        Globals.vb
'*
'* Description: Static class which provides functions to access the database
'*              connection string. It also contains a function to process
'*              keys for tables which use composite primary keys.
'*
'*-----------------------------------------------------------------------------
'*                      CHANGE HISTORY
'*   Change No:   Date:          Author:   Description:
'*   _________    ___________    ______    ____________________________________
'*      001       18/03/2002     Custom Created.
'* 
'******************************************************************************
Public NotInheritable Class Globals

    Private Shared mConnectionString As String = ""
    Private Shared mRptConnectionString As String = ""
    Private Shared mCrystalRptPath As String = ""
    Private Shared mstrAppPath As String = ""
    Private Shared mUserName As String = ""
    Private Shared mPassword As String = ""
    Private Shared mOldUserName As String = ""
    Private Shared mOldPassword As String = ""
    Private Shared mbClose As Boolean = False
    Private Shared mbCancel As Boolean = False
    Private Shared mDefaultExaminerID As Integer = Nothing
    Private Shared mDefaultSiteID As Integer = Nothing
    Private Shared mChartChanged As Boolean = False
    Private Shared mNFSChanged As Boolean = False
    Private Shared mAssessmentChanged As Boolean = False
    Private Shared mWDChanged As Boolean = False
    Private Shared mMedChanged As Boolean = False
    Private Shared mLabChanged As Boolean = False
    Private Shared mKEY_SEPERATOR As String = "ABCD-1232-2423-4242"
    Private Shared mLMP As String = ""
    Private Shared mUseEDCBy As String = ""
    Private Shared mEarlyUS As String = ""
    Private Shared mEDC As String = ""
    Private Shared mTempStr1 As String = ""
    Private Shared mTempStr2 As String = ""
    Private Shared mTempStr3 As String = ""
    Private Shared mTempStr4 As String = ""
    Private Shared mTempStr5 As String = ""
    Private Shared mTempStr6 As String = ""
    Private Shared mTempStr7 As String = ""
    Private Shared mTempStr8 As String = ""
    Private Shared mTempStr9 As String = ""
    Private Shared mTempStr10 As String = ""
    Private Shared mTempStr11 As String = ""
    Private Shared mbUpdate As Boolean = False
    Private Shared mChartID As Integer = Nothing
    Private Shared mExamID As Integer = Nothing
    Private Shared mPatientID As Integer = Nothing
    Private Shared mIntakeChanged As Boolean = False
    Private Shared mbExamChanged As Boolean = False
    Private Shared mbChartChanged As Boolean = False
    Private Shared mAppContinue As Boolean = True
    Private Shared mbResetConnectionstring As Boolean = False
    Private Shared mExamMode As Boolean = False
    Private Shared mRHChanged As Boolean = False
    Private Shared mCustomDictPath As String = ""
    Private Shared mDest_DocumentPath As String = ""
    Private Shared mSource_DocumentPath As String = ""
    Private Shared mExamDate As String = ""
    Private Shared mShowEDCByUltrasound As Boolean = False
    Private Shared mChartLocked As Boolean = False
    Private Shared mUserRole As String = ""
    Private Shared mLimPhysicianID As Integer = 0
    Private Shared mWordTemplatePath As String = ""
    Private Shared mBloodType As String = ""
    Private Shared mRH As String = ""
    Private Shared mAntibody As String = ""
    Private Shared mUseWinfax As Boolean = False
    Private Shared mCastelleUID As String = ""
    Private Shared mCastellePW As String = ""
    Private Shared mCastelleIPAddress As String = ""
    Private Shared mUserExaminerID As Integer = Nothing
    Public Event UserNameChanged As EventHandler
    Private Shared mGetTimeLastUpd As Date = Now()
    Private Shared mReviewedStamp As String = Application.StartupPath + "/ReviewedStamp.gif"
    Private Shared mMaternalAge As Integer = Nothing
    Private Shared mDOB As Date = Nothing
    Private Shared mStrReceive As String = ""
    Public Shared frmSerialPort As frmPortListener = Nothing



    '***********************************************************b***************
    '*  
    '* Name:        New
    '*
    '* Description: Since this class provides only static methods, make the
    '*              Default constructor private to prevent instances from being
    '*              created with "new Globals()".
    '*
    '**************************************************************************
    Private Sub New()
    End Sub 'New


    '**************************************************************************
    '*  
    '* Name:        CrystalRptPath
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared ReadOnly Property IPAddress() As String
        Get
            Dim pcAddress As System.Net.IPAddress
            Dim pcIPAddress As String
            With System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName())
                pcAddress = New System.Net.IPAddress(.AddressList(0).Address)
                pcIPAddress = pcAddress.ToString
            End With
            IPAddress = pcIPAddress
            Return IPAddress
        End Get
    End Property 'IPAddress

    '**************************************************************************
    '*  
    '* Name:        ConnectionString
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared ReadOnly Property ConnectionString() As String
        Get
            If Len(mConnectionString) = 0 Or mbResetConnectionstring = True Then
                Dim objIniFile As New ProcessIni(Application.StartupPath & "\OBFOCUS.ini")
                Dim strData As String = objIniFile.GetString("Login", "ConnectionString", " ")
                Dim strDataSource As String = objIniFile.GetString("Login", "Data Source", " ")
                mConnectionString = strData & "Data Source=" & strDataSource & ";User ID=" & UserName & ";Password=" & Password & ";"
                mbResetConnectionstring = False
            End If
            Return mConnectionString
        End Get
    End Property 'ConnectionString
    '**************************************************************************
    '*  
    '* Name:        CheckConnectionString
    '*
    '* Description: Database connection string. check if ini has connectionstring
    '*              else, let user enter connectionstring
    '*
    '**************************************************************************
    Public Shared Sub CheckConnectionString()
        If Len(mConnectionString) = 0 Then
            Dim getDataSource As String
            Dim objIniFile As New ProcessIni(Application.StartupPath & "\OBFOCUS.ini")
            Dim strData As String = objIniFile.GetString("Login", "Data Source", " ")
            If Len(strData) = 0 Then
                getDataSource = InputBox("Server name needs to be recorded.  If you are unsure about your server name, " _
                        & vbCrLf & "please contact your network administrator." & vbCrLf & vbCrLf & _
                        "Please enter server name now.", "Program Initialization")
                If Len(getDataSource) = 0 Then
                    MsgBox("You have not entered a server name.  You will not be able to access your database!", MsgBoxStyle.Critical, "Invalid Entry")
                Else
                    objIniFile.WriteString("Login", "Data Source", getDataSource)
                End If
            End If
            'now let's check if reportsettings section is also missing DataSource
            strData = objIniFile.GetString("Settings", "Server", " ")
            If Len(strData) = 0 Then
                If Len(getDataSource) = 0 Then
                    getDataSource = InputBox("Server name needs to be recorded.  If you are unsure about your server name, " _
                            & vbCrLf & "please contact your network administrator." & vbCrLf & vbCrLf & _
                            "Please enter server name now.", "Program Initialization")
                End If
                If Len(getDataSource) = 0 Then
                    MsgBox("You have not entered a server name for your reports.  You will not be able to access your database!", MsgBoxStyle.Critical, "Invalid Entry")
                Else
                    objIniFile.WriteString("Settings", "Server", getDataSource)
                End If
            End If
        End If
    End Sub 'CheckConnectionString
    '**************************************************************************
    '*  
    '* Name:        RptConnectionString
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared ReadOnly Property RptConnectionString() As String
        Get
            'TODO Change application to read connection string from registry, rather then hardcode it as above.
            If Len(mRptConnectionString) = 0 Then
                Dim objIniFile As New ProcessIni(Application.StartupPath & "\OBFOCUS.ini")
                Dim strData As String = objIniFile.GetString("Settings", "RptConnectionString", " ")
                Dim strDataSource As String = objIniFile.GetString("Settings", "Server", " ")
                mRptConnectionString = strData & "Server=" & strDataSource & ";User ID=" & UserName & ";Password=" & Password & ";"
            End If
            Return mRptConnectionString
        End Get
    End Property 'RptConnectionString

    '**************************************************************************
    '*  
    '* Name:        CrystalRptPath
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared ReadOnly Property CrystalRptPath() As String
        Get
            'TODO Change application to read connection string from registry, rather then hardcode it as above.
            If Len(mCrystalRptPath) = 0 Then
                Dim objIniFile As New ProcessIni(Application.StartupPath & "\OBFOCUS.ini")
                Dim strData As String = objIniFile.GetString("Settings", "CrystalRptPath", " ")
                mCrystalRptPath = strData
            End If
            Return mCrystalRptPath
        End Get
    End Property 'CrystalRptPath
    '**************************************************************************
    '*  
    '* Name:        AppPath
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared ReadOnly Property AppPath() As String
        Get
            'TODO Change application to read connection string from registry, rather then hardcode it as above.
            If Len(mstrAppPath) = 0 Then
                Dim objIniFile As New ProcessIni(Application.StartupPath & "\OBFOCUS.ini")
                Dim strData As String = objIniFile.GetString("Settings", "strAppPath", " ")
                mstrAppPath = strData
            End If
            Return mstrAppPath
        End Get
    End Property 'AppPath
    '**************************************************************************
    '*  
    '* Name:        bAppContinue
    '*
    '* Description: This boolean field is used to determine whether a user can
    '*              continue on with the app.  If he/she has the proper login/pw
    '*              then AppContinue is set to true.
    '*
    '**************************************************************************
    Public Shared Property bAppContinue() As Boolean
        Get
            Return mAppContinue
        End Get
        Set(ByVal Value As Boolean)
            If mAppContinue = Value Then Exit Property
            mAppContinue = Value
        End Set
    End Property 'AppContinue
    '**************************************************************************
    '*  
    '* Name:        GetTimeLastUpd
    '*
    '* Description: This boolean field is used to determine whether a user can
    '*              continue on with the app.  If he/she has the proper login/pw
    '*              then AppContinue is set to true.
    '*
    '**************************************************************************
    Public Shared Property GetTimeLastUpd() As Date
        Get
            Return mGetTimeLastUpd
        End Get
        Set(ByVal Value As Date)
            If mGetTimeLastUpd = Value Then Exit Property
            mGetTimeLastUpd = Value
        End Set
    End Property 'GetTimeLastUpd
    '**************************************************************************
    '*  
    '* Name:        UserName
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property UserName() As String
        Get
            Return mUserName
        End Get
        Set(ByVal Value As String)
            If mUserName = Value Then Exit Property
            mUserName = Value
        End Set
    End Property 'UserName
    '**************************************************************************
    '*  
    '* Name:        MaternalAge
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property MaternalAge() As Integer
        Get
            Return mMaternalAge
        End Get
        Set(ByVal Value As Integer)
            If mMaternalAge = Value Then Exit Property
            mMaternalAge = Value
        End Set
    End Property 'MaternalAge
    '**************************************************************************
    '*  
    '* Name:        LMP
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property LMP() As String
        Get
            Return mLMP
        End Get
        Set(ByVal Value As String)
            If mLMP = Value Then Exit Property
            mLMP = Value
        End Set
    End Property 'LMP
    '**************************************************************************
    '*  
    '* Name:        ReviewedStamp
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property ReviewedStamp() As String
        Get
            Return mReviewedStamp
        End Get
        Set(ByVal Value As String)
            If mReviewedStamp = Value Then Exit Property
            mReviewedStamp = Value
        End Set
    End Property 'ReviewedStamp
    '**************************************************************************
    '*  
    '* Name:        StrReceive
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property StrReceive() As String
        Get
            Return mStrReceive
        End Get
        Set(ByVal Value As String)
            If mStrReceive = Value Then Exit Property
            mStrReceive = Value
        End Set
    End Property 'StrReceive
    '**************************************************************************
    '*  
    '* Name:        UseEDCBy
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property UseEDCBy() As String
        Get
            Return mUseEDCBy
        End Get
        Set(ByVal Value As String)
            If mUseEDCBy = Value Then Exit Property
            mUseEDCBy = Value
        End Set
    End Property 'UseEDCBy
    '**************************************************************************
    '*  
    '* Name:        EarlyUS
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property EarlyUS() As String
        Get
            Return mEarlyUS
        End Get
        Set(ByVal Value As String)
            If mEarlyUS = Value Then Exit Property
            mEarlyUS = Value
        End Set
    End Property 'EarlyUS
    '**************************************************************************
    '*  
    '* Name:        EDC
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property EDC() As String
        Get
            Return mEDC
        End Get
        Set(ByVal Value As String)
            If mEDC = Value Then Exit Property
            mEDC = Value
        End Set
    End Property 'EDC
    '**************************************************************************
    '*  
    '* Name:        DefaultExaminerID
    '*
    '* Description: Database connection integer. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property DefaultExaminerID() As Integer
        Get
            Return mDefaultExaminerID
        End Get
        Set(ByVal Value As Integer)
            If mDefaultExaminerID = Value Then Exit Property
            mDefaultExaminerID = Value
        End Set
    End Property 'DefaultExaminerID
    '**************************************************************************
    '*  
    '* Name:        DefaultSiteID
    '*
    '* Description: Database connection integer. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property DefaultSiteID() As Integer
        Get
            Return mDefaultSiteID
        End Get
        Set(ByVal Value As Integer)
            If mDefaultSiteID = Value Then Exit Property
            mDefaultSiteID = Value
        End Set
    End Property 'DefaultSiteID
    '**************************************************************************
    '*  
    '* Name:        ChartID
    '*
    '* Description: Database connection integer. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property ChartID() As Integer
        Get
            Return mChartID
        End Get
        Set(ByVal Value As Integer)
            If mChartID = Value Then Exit Property
            mChartID = Value
        End Set
    End Property 'ChartID
    '**************************************************************************
    '*  
    '* Name:        DOB
    '*
    '* Description: Database connection Date. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property DOB() As Date
        Get
            Return mDOB
        End Get
        Set(ByVal Value As Date)
            If mDOB = Value Then Exit Property
            mDOB = Value
        End Set
    End Property 'DOB
    '**************************************************************************
    '*  
    '* Name:        PatientID
    '*
    '* Description: Database connection integer. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property PatientID() As Integer
        Get
            Return mPatientID
        End Get
        Set(ByVal Value As Integer)
            If mPatientID = Value Then Exit Property
            mPatientID = Value
        End Set
    End Property 'PatientID
    '**************************************************************************
    '*  
    '* Name:        ExamID
    '*
    '* Description: Database connection integer. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property ExamID() As Integer
        Get
            Return mExamID
        End Get
        Set(ByVal Value As Integer)
            If mExamID = Value Then Exit Property
            mExamID = Value
        End Set
    End Property 'ExamID

    '**************************************************************************
    '*  
    '* Name:        Password
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property Password() As String
        Get
            Return mPassword
        End Get
        Set(ByVal Value As String)
            If mPassword = Value Then Exit Property
            mPassword = Value
        End Set
    End Property 'Password
    '**************************************************************************
    '*  
    '* Name:        TempStr1
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property TempStr1() As String
        Get
            Return mTempStr1
        End Get
        Set(ByVal Value As String)
            If mTempStr1 = Value Then Exit Property
            mTempStr1 = Value
        End Set
    End Property 'TempStr1
    '**************************************************************************
    '*  
    '* Name:        TempStr2
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property TempStr2() As String
        Get
            Return mTempStr2
        End Get
        Set(ByVal Value As String)
            If mTempStr2 = Value Then Exit Property
            mTempStr2 = Value
        End Set
    End Property 'TempStr2
    '**************************************************************************
    '*  
    '* Name:        TempStr3
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property TempStr3() As String
        Get
            Return mTempStr3
        End Get
        Set(ByVal Value As String)
            If mTempStr3 = Value Then Exit Property
            mTempStr3 = Value
        End Set
    End Property 'TempStr3
    '**************************************************************************
    '*  
    '* Name:        TempStr4
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property TempStr4() As String
        Get
            Return mTempStr4
        End Get
        Set(ByVal Value As String)
            If mTempStr4 = Value Then Exit Property
            mTempStr4 = Value
        End Set
    End Property 'TempStr4
    '**************************************************************************
    '*  
    '* Name:        TempStr5
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property TempStr5() As String
        Get
            Return mTempStr5
        End Get
        Set(ByVal Value As String)
            If mTempStr5 = Value Then Exit Property
            mTempStr5 = Value
        End Set
    End Property 'TempStr5
    '**************************************************************************
    '*  
    '* Name:        TempStr6
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property TempStr6() As String
        Get
            Return mTempStr6
        End Get
        Set(ByVal Value As String)
            If mTempStr6 = Value Then Exit Property
            mTempStr6 = Value
        End Set
    End Property 'TempStr6
    '**************************************************************************
    '*  
    '* Name:        TempStr7
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property TempStr7() As String
        Get
            Return mTempStr7
        End Get
        Set(ByVal Value As String)
            If mTempStr7 = Value Then Exit Property
            mTempStr7 = Value
        End Set
    End Property 'TempStr7
    '**************************************************************************
    '*  
    '* Name:        TempStr8
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property TempStr8() As String
        Get
            Return mTempStr8
        End Get
        Set(ByVal Value As String)
            If mTempStr8 = Value Then Exit Property
            mTempStr8 = Value
        End Set
    End Property 'TempStr8
    '**************************************************************************
    '*  
    '* Name:        TempStr9
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property TempStr9() As String
        Get
            Return mTempStr9
        End Get
        Set(ByVal Value As String)
            If mTempStr9 = Value Then Exit Property
            mTempStr9 = Value
        End Set
    End Property 'TempStr9
    '**************************************************************************
    '*  
    '* Name:        TempStr10
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property TempStr10() As String
        Get
            Return mTempStr10
        End Get
        Set(ByVal Value As String)
            If mTempStr10 = Value Then Exit Property
            mTempStr10 = Value
        End Set
    End Property 'TempStr10
    '**************************************************************************
    '*  
    '* Name:        TempStr11
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property TempStr11() As String
        Get
            Return mTempStr11
        End Get
        Set(ByVal Value As String)
            If mTempStr11 = Value Then Exit Property
            mTempStr11 = Value
        End Set
    End Property 'TempStr11
    '**************************************************************************
    '*  
    '* Name:        BClose
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property BClose() As Boolean
        Get
            Return mbClose
        End Get
        Set(ByVal Value As Boolean)
            If mbClose = Value Then Exit Property
            mbClose = Value
        End Set
    End Property 'BClose
    '**************************************************************************
    '*  
    '* Name:        bChartChanged
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property bChartChanged() As Boolean
        Get
            Return mbChartChanged
        End Get
        Set(ByVal Value As Boolean)
            If mbChartChanged = Value Then Exit Property
            mbChartChanged = Value
        End Set
    End Property 'bChartChanged
    '**************************************************************************
    '*  
    '* Name:        ShowEDCByUltrasound
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property ShowEDCByUltrasound() As Boolean
        Get
            Return mShowEDCByUltrasound
        End Get
        Set(ByVal Value As Boolean)
            If mShowEDCByUltrasound = Value Then Exit Property
            mShowEDCByUltrasound = Value
        End Set
    End Property 'ShowEDCByUltrasound
    '**************************************************************************
    '*  
    '* Name:        BUpdate
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property BUpdate() As Boolean
        Get
            Return mbUpdate
        End Get
        Set(ByVal Value As Boolean)
            If mbUpdate = Value Then Exit Property
            mbUpdate = Value
        End Set
    End Property 'BUpdate
    '**************************************************************************
    '*  
    '* Name:        ChartLocked
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property ChartLocked() As Boolean
        Get
            Return mChartLocked
        End Get
        Set(ByVal Value As Boolean)
            If mChartLocked = Value Then Exit Property
            mChartLocked = Value
        End Set
    End Property 'ChartLocked
    '**************************************************************************
    '*  
    '* Name:        bResetConnectionString
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property bResetConnectionString() As Boolean
        Get
            Return mbResetConnectionstring
        End Get
        Set(ByVal Value As Boolean)
            If mbResetConnectionstring = Value Then Exit Property
            mbResetConnectionstring = Value
        End Set
    End Property 'bResetConnectionString
    '**************************************************************************
    '*  
    '* Name:        BExamChanged
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property BExamChanged() As Boolean
        Get
            Return mbExamChanged
        End Get
        Set(ByVal Value As Boolean)
            If mbExamChanged = Value Then Exit Property
            mbExamChanged = Value
        End Set
    End Property 'BExamChanged
    '**************************************************************************
    '*  
    '* Name:        WDChanged
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property WDChanged() As Boolean
        Get
            Return mWDChanged
        End Get
        Set(ByVal Value As Boolean)
            If mWDChanged = Value Then Exit Property
            mWDChanged = Value
        End Set
    End Property 'WDChanged
    '**************************************************************************
    '*  
    '* Name:        MedChanged
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property MedChanged() As Boolean
        Get
            Return mMedChanged
        End Get
        Set(ByVal Value As Boolean)
            If mMedChanged = Value Then Exit Property
            mMedChanged = Value
        End Set
    End Property 'MedChanged
    '**************************************************************************
    '*  
    '* Name:        LabChanged
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property LabChanged() As Boolean
        Get
            Return mLabChanged
        End Get
        Set(ByVal Value As Boolean)
            If mLabChanged = Value Then Exit Property
            mLabChanged = Value
        End Set
    End Property 'LabChanged
    '**************************************************************************
    '*  
    '* Name:        RHChanged
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property RHChanged() As Boolean
        Get
            Return mRHChanged
        End Get
        Set(ByVal Value As Boolean)
            If mRHChanged = Value Then Exit Property
            mRHChanged = Value
        End Set
    End Property 'RHChanged
    '**************************************************************************
    '*  
    '* Name:        IntakeChanged
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property IntakeChanged() As Boolean
        Get
            Return mIntakeChanged
        End Get
        Set(ByVal Value As Boolean)
            If mIntakeChanged = Value Then Exit Property
            mIntakeChanged = Value
        End Set
    End Property 'IntakeChanged
    '**************************************************************************
    '*  
    '* Name:        ChartChanged
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property ChartChanged() As Boolean
        Get
            Return mChartChanged
        End Get
        Set(ByVal Value As Boolean)
            If mChartChanged = Value Then Exit Property
            mChartChanged = Value
        End Set
    End Property 'ChartChanged
    '**************************************************************************
    '*  
    '* Name:        NFSChanged
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property NFSChanged() As Boolean
        Get
            Return mNFSChanged
        End Get
        Set(ByVal Value As Boolean)
            If mNFSChanged = Value Then Exit Property
            mNFSChanged = Value
        End Set
    End Property 'NFSChanged
    '**************************************************************************
    '*  
    '* Name:        AssessmentChanged
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property AssessmentChanged() As Boolean
        Get
            Return mAssessmentChanged
        End Get
        Set(ByVal Value As Boolean)
            If mAssessmentChanged = Value Then Exit Property
            mAssessmentChanged = Value
        End Set
    End Property 'AssessmentChanged
    '**************************************************************************
    '*  
    '* Name:        bCancel
    '*
    '* Description: This will tell the system whether to cancel an event
    '*              (e.g. Close event) or not.
    '*
    '**************************************************************************
    Public Shared Property bCancel() As Boolean
        Get
            Return mbCancel
        End Get
        Set(ByVal Value As Boolean)
            If mbCancel = Value Then Exit Property
            mbCancel = Value
        End Set
    End Property 'bCancel
    '**************************************************************************
    '*  
    '* Name:        ExamMode
    '*
    '* Description: This will tell the system whether users clicked onto Close
    '*              button.  If they did, no need to run RefreshData
    '*
    '**************************************************************************
    Public Shared Property ExamMode() As Boolean
        Get
            Return mExamMode
        End Get
        Set(ByVal Value As Boolean)
            If mExamMode = Value Then Exit Property
            mExamMode = Value
        End Set
    End Property 'ExamMode
    '**************************************************************************
    '*  
    '* Name:        CustomDictPath
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property CustomDictPath() As String
        Get
            Return mCustomDictPath
        End Get
        Set(ByVal Value As String)
            If mCustomDictPath = Value Then Exit Property
            mCustomDictPath = Value
        End Set
    End Property 'CustomDictPath
    '**************************************************************************
    '*  
    '* Name:        LimPhysicianID
    '*
    '* Description: Database connection integer. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property LimPhysicianID() As Integer
        Get
            Return mLimPhysicianID
        End Get
        Set(ByVal Value As Integer)
            If mLimPhysicianID = Value Then Exit Property
            mLimPhysicianID = Value
        End Set
    End Property 'LimPhysicianID
    '**************************************************************************
    '*  
    '* Name:        ExamDate
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property ExamDate() As String
        Get
            Return mExamDate
        End Get
        Set(ByVal Value As String)
            If mExamDate = Value Then Exit Property
            mExamDate = Value
        End Set
    End Property 'ExamDate
    '**************************************************************************
    '*  
    '* Name:        Dest_DocumentPath
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property Dest_DocumentPath() As String
        Get
            Return mDest_DocumentPath
        End Get
        Set(ByVal Value As String)
            If mDest_DocumentPath = Value Then Exit Property
            mDest_DocumentPath = Value
        End Set
    End Property 'Dest_DocumentPath
    '**************************************************************************
    '*  
    '* Name:        Source_DocumentPath
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property Source_DocumentPath() As String
        Get
            Return mSource_DocumentPath
        End Get
        Set(ByVal Value As String)
            If mSource_DocumentPath = Value Then Exit Property
            mSource_DocumentPath = Value
        End Set
    End Property 'Source_DocumentPath
    '**************************************************************************
    '*  
    '* Name:        WordTemplatePath
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property WordTemplatePath() As String
        Get
            Return mWordTemplatePath
        End Get
        Set(ByVal Value As String)
            If mWordTemplatePath = Value Then Exit Property
            mWordTemplatePath = Value
        End Set
    End Property 'WordTemplatePath
    '**************************************************************************
    '*  
    '* Name:        OldUserName
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property OldUserName() As String
        Get
            Return mOldUserName
        End Get
        Set(ByVal Value As String)
            If mOldUserName = Value Then Exit Property
            mOldUserName = Value
        End Set
    End Property 'OldUserName
    '**************************************************************************
    '*  
    '* Name:        OldPassword
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property OldPassword() As String
        Get
            Return mOldPassword
        End Get
        Set(ByVal Value As String)
            If mOldPassword = Value Then Exit Property
            mOldPassword = Value
        End Set
    End Property 'OldPassword
    '**************************************************************************
    '*  
    '* Name:        UserRole
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property UserRole() As String
        Get
            Return mUserRole
        End Get
        Set(ByVal Value As String)
            If mUserRole = Value Then Exit Property
            mUserRole = Value
        End Set
    End Property 'UserRole
    '**************************************************************************
    '*  
    '* Name:        RH
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property RH() As String
        Get
            Return mRH
        End Get
        Set(ByVal Value As String)
            If mRH = Value Then Exit Property
            mRH = Value
        End Set
    End Property 'RH
    '**************************************************************************
    '*  
    '* Name:        BloodType
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property BloodType() As String
        Get
            Return mBloodType
        End Get
        Set(ByVal Value As String)
            If mBloodType = Value Then Exit Property
            mBloodType = Value
        End Set
    End Property 'BloodType
    '**************************************************************************
    '*  
    '* Name:        Antibody
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property Antibody() As String
        Get
            Return mAntibody
        End Get
        Set(ByVal Value As String)
            If mAntibody = Value Then Exit Property
            mAntibody = Value
        End Set
    End Property 'Antibody
    '**************************************************************************
    '*  
    '* Name:        UseWinfax
    '*
    '* Description: This boolean field is used to determine whether a user can
    '*              continue on with the app.  If he/she has the proper login/pw
    '*              then UseWinfax is set to true.
    '*
    '**************************************************************************
    Public Shared Property UseWinfax() As Boolean
        Get
            Return mUseWinfax
        End Get
        Set(ByVal Value As Boolean)
            If mUseWinfax = Value Then Exit Property
            mUseWinfax = Value
        End Set
    End Property 'UseWinfax
    '**************************************************************************
    '*  
    '* Name:        CastelleUID
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property CastelleUID() As String
        Get
            Return mCastelleUID
        End Get
        Set(ByVal Value As String)
            If mCastelleUID = Value Then Exit Property
            mCastelleUID = Value
        End Set
    End Property 'CastelleUID
    '**************************************************************************
    '*  
    '* Name:        CastellePW
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property CastellePW() As String
        Get
            Return mCastellePW
        End Get
        Set(ByVal Value As String)
            If mCastellePW = Value Then Exit Property
            mCastellePW = Value
        End Set
    End Property 'CastellePW
    '**************************************************************************
    '*  
    '* Name:        CastelleIPAddress
    '*
    '* Description: Database connection string. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property CastelleIPAddress() As String
        Get
            Return mCastelleIPAddress
        End Get
        Set(ByVal Value As String)
            If mCastelleIPAddress = Value Then Exit Property
            mCastelleIPAddress = Value
        End Set
    End Property 'CastelleIPAddress
    '**************************************************************************
    '*  
    '* Name:        UserExaminerID
    '*
    '* Description: Database connection integer. This could be stored and read 
    '*              from the registry or an application INI file.
    '*
    '**************************************************************************
    Public Shared Property UserExaminerID() As Integer
        Get
            Return mUserExaminerID
        End Get
        Set(ByVal Value As Integer)
            If mUserExaminerID = Value Then Exit Property
            mUserExaminerID = Value
        End Set
    End Property 'UserExaminerID
    '**************************************************************************
    '*  
    '* Name:        SeperatorKey
    '*
    '* Description: The seperator key is used to act as a delimitter between the 
    '*              constituent parts of a composite key.
    '*
    '**************************************************************************
    Public Shared ReadOnly Property KeySeperator() As String
        Get
            Return mKEY_SEPERATOR
        End Get
    End Property 'KeySeperator


    '**************************************************************************
    '*  
    '* Name:        ParseKey
    '*
    '* Description: Parses a key that is made up of several primary keys that 
    '*              seperated by a KEY_SEPERATOR. The primary keys are copied to 
    '*              an array.
    '*
    '* Parameters:  Key         - String which contains the primary keys, seperated 
    '*                            by a sequence of characters
    '*              PrimaryKeys - array that will be populated with the individual
    '*                            elements of the key
    '*
    '* Example:     Key - key1SEPERATORkey2SEPERATOR ..... keyn
    '*              PrimaryKeys(0)   - key1
    '*              PrimaryKeys(1)   - key2
    '*              PrimaryKeys(n-1) - keyn
    '* 
    '* Remarks:     This is used to help make the programming of tables which use
    '*              composite keys easier.
    '**************************************************************************
    Public Shared Sub ParseKey(ByVal Key As String, ByVal PrimaryKeys() As String)

        Dim intPos As Integer = -1
        Dim i As Integer = 0

        intPos = Key.IndexOf(KeySeperator)

        While intPos <> -1
            PrimaryKeys(i) = Key.Substring(0, intPos)
            Key = Key.Remove(0, (PrimaryKeys(i).Length + KeySeperator.Length))
            i = i + 1
            intPos = Key.IndexOf(KeySeperator)
        End While
        PrimaryKeys(i) = Key

    End Sub 'ParseKey


    '**************************************************************************
    '*  
    '* Name:        XP
    '*
    '* Description: Returns boolean indicating if the programme is running on
    '*              XP. If it is, the grid will be assigned the XP style.
    '**************************************************************************
    Public Shared Function XP() As Boolean
        Dim VersionDetect As New Custom.Windows.Forms.VersionDetector
        Return (VersionDetect.WindowsName.IndexOf("XP") > 0)
    End Function 'XP


    '**************************************************************************
    '*  
    '* Name:        InitVars
    '*
    '* Description: Initialize variables in the Globals class
    '**************************************************************************
    Public Shared Sub InitVars()
        mbClose = False
        mbCancel = False
    End Sub 'InitVars

#Region "Database Transaction Support"

    Private Shared _Transaction As SqlTransaction = Nothing

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
    Public Shared Property Transaction() As SqlTransaction
        Get
            Return _Transaction
        End Get
        Set(ByVal Value As SqlTransaction)
            _Transaction = Value
        End Set
    End Property 'Transaction


    '**************************************************************************
    '*  
    '* Name:        BeginTransaction
    '*
    '* Description: Used for transaction support.
    '*
    '* Returns:     Boolean indicating if operation was sucessful or not.
    '*
    '* Remarks:     Opens a transaction and sets the Transaction object to it.
    '*              All subsequent calls to the DAL will use a transaction.
    '*
    '**************************************************************************
    Public Shared Function BeginTransaction() As Boolean
        Try
            Dim cn As New SqlConnection(Globals.ConnectionString)
            cn.Open()
            Transaction = cn.BeginTransaction()
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try
    End Function 'BeginTransaction


    '**************************************************************************
    '*  
    '* Name:        CommitTransaction
    '*
    '* Description: Commit transaction to database.
    '*
    '* Returns:     Boolean indicating if operation was sucessful or not.
    '*
    '* Remarks:     Closes the transaction and sets it to Nothing so that all
    '*              subsequent database operations are performed outside a 
    '*              transaction.
    '*
    '**************************************************************************
    Public Shared Function CommitTransaction() As Boolean
        Try
            If Transaction Is Nothing Then
                Throw New Exception("Programming error. Cannot commit transaction because the Transaction object is nothing.")
            End If
            Transaction.Commit()
            Transaction = Nothing
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try
    End Function 'CommitTransaction


    '**************************************************************************
    '*  
    '* Name:        RollbackTransaction
    '*
    '* Description: Rollback transaction from database.
    '*
    '* Returns:     Boolean indicating if operation was sucessful or not.
    '*
    '* Remarks:     Closes the transaction and sets it to Nothing so that all
    '*              subsequent database operations are performed outside a 
    '*              transaction.
    '*
    '**************************************************************************
    Public Shared Function RollbackTransaction() As Boolean
        Try
            If Transaction Is Nothing Then
                Throw New Exception("Programming error. Cannot commit transaction because the Transaction object is nothing.")
            End If
            Transaction.Rollback()
            Transaction = Nothing
            Return True
        Catch ex As Exception
            ExceptionManager.Publish(ex)
            Return False
        End Try
    End Function 'RollbackTransaction

#End Region

#Region "Shared Methods"
    Public Shared Function Round2C(ByVal x As Object) As Object
        If x Is System.DBNull.Value Then
            Round2C = System.DBNull.Value
        Else
            Round2C = System.Math.Round((Int(CType(x, Double) * 100 + 0.5) / 100), 2)
        End If
    End Function

    Public Shared Function Num2FracA(ByVal x As Object, ByVal Denominator As Long) As String
        Dim Temp As String, Fixed As Double, Numerator As Long
        If Not IsNumeric(x) Then
            Num2FracA = CType(x, String)
            Exit Function
        End If
        If CType(CType(x, String).Substring(InStr(CType(x, String), "."), Len(CType(x, String)) - InStr(CType(x, String), ".")), Double) = 0 Then
            Num2FracA = CType(x, String)
            Exit Function
        End If
        x = CType(x, Double)
        x = System.Math.Abs(CType(x, Double))
        Fixed = CType(x, Integer)
        Numerator = CType((CType(x, Double) - Fixed) * Denominator + 0.5, Integer)
        If Numerator = Denominator Then
            Fixed = Fixed + 1
            Numerator = 0
        End If
        If Fixed > 0 Then
            Temp = CType(Fixed, String)
        End If
        If Numerator > 0 Then
            Temp = Temp & " " & Numerator & "/" & Denominator
        End If
        Num2FracA = Temp
    End Function
    Public Shared Function PostfixNum(ByVal vNumIn As Object) As String
        Dim psNumList(4) As String
        Dim psNumString As String
        Dim piNumEval As Integer

        psNumList(0) = "th"
        psNumList(1) = "st"
        psNumList(2) = "nd"
        psNumList(3) = "rd"
        psNumList(4) = ""

        psNumString = CStr(vNumIn)
        piNumEval = CType(Microsoft.VisualBasic.Right(psNumString, 1), Integer)

        Select Case piNumEval
            Case 0, Is > 3
                piNumEval = 0
            Case 1 To 3
                piNumEval = piNumEval
            Case Else
                piNumEval = 4
        End Select

        PostfixNum = psNumString & psNumList(piNumEval)

    End Function

#End Region
End Class 'Globals










