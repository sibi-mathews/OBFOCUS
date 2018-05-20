Option Explicit On 
Option Strict On
'******************************************************************************
'*
'* Name:        ExceptionManager.vb
'*
'* Description: Logs and processes all application exceptions.
'*
'* Remarks:     This class has the same structure as the Microsoft Exception Management
'*              Application Block for .NET. For more information, see
'*              http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnbda/html/daag.asp
'*
'*              By using this generic structure, the Custom Exception Manager
'*              can be replaced with the one developed by Microsoft. In time,
'*              we will replace this class with the one from Microsoft.
'*
'* Remarks:     This class was developed as part of VS.NET beta 2. Microsoft 
'*              are planning a version of this class around April 2002 which 
'*              will be released for the VS.NET RTM version.
'*
'*-----------------------------------------------------------------------------
'*                      CHANGE HISTORY
'*   Change No:   Date:          Author:   Description:
'*   _________    ___________    ______    ____________________________________
'*      001       14-Mar-2002    Custom Created.
'*      002       07-Aug-2002    Custom Added message for no stored proc.                                
'* 
'******************************************************************************
Public NotInheritable Class ExceptionManager


    '**************************************************************************
    '*  
    '* Name:        New
    '*
    '* Description: Private constructor to restrict an instance of this class 
    '*              from being created with "new ExceptionManager()".
    '*
    '**************************************************************************
    Private Sub New()
    End Sub 'New


    '**************************************************************************
    '*  
    '* Name:        Publish
    '*
    '* Description: Static method to publish the exception information.
    '*
    '* Parameters:  exception: The exception object whose information should be published. 
    '*
    '* Remarks:     At the moment, this proc only displays the exception to the
    '*              user. When the Micrsoft Exception Management Block will be 
    '*              used, this method will be able to log the exception to 
    '*              various places (e.g. database, log file).
    '**************************************************************************
    Public Overloads Shared Function Publish(ByVal exception As Exception) As String
        'Change 002: ensure users have generated stored procs and installed them on the database.
        'TODO: Please remove this check. It is only included for users new to Custom.NET.
        Dim errorMessage As String = String.Empty
        If exception.Message.ToUpper().StartsWith("COULD NOT FIND STORED PROCEDURE") Then
            errorMessage = exception.Message & vbCrLf & vbCrLf & "Please make sure that you have generated the SQL scriOBFOCUS and that you have installed them on the database." & vbCrLf & vbCrLf & "For more information, please consult Tutorial 3.0 of the Custom.NET Tutorial document."
        ElseIf exception.Message.ToUpper() = "TIMESTAMP ERROR" Then
            errorMessage = "The changes you made were not sucessfully saved because another user updated the record while you were performing the operation." & vbCrLf & vbCrLf & "Please close the screen, refresh the data and try the operation again."
        ElseIf exception.Message.ToUpper().IndexOf("YEAR, MONTH") > 0 And exception.StackTrace.ToUpper().IndexOf("IBM.DATA.DB2") > 0 Then
            errorMessage = "There is a mismatch between the server date format and the date format for DB2. Please see http://www7b.software.ibm.com/dmdd/library/techarticle/0211yip/0211yip3.html for more information. Please note that this is not a problem with the generated code."
        ElseIf exception.Message.ToUpper().StartsWith("LOGIN FAILED FOR USER") Then
            errorMessage = "Invalid Login/Password"
        ElseIf exception.Message.ToUpper().StartsWith("CANNOT OPEN DATABASE REQUESTED IN LOGIN") Then
            errorMessage = "Invalid Login/Password"
        Else
            errorMessage = exception.Source & vbCrLf & vbCrLf & exception.Message & vbCrLf & vbCrLf & exception.StackTrace
        End If

        Return errorMessage
    End Function 'Publish


End Class 'ExceptionManager


