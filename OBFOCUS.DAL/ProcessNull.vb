Option Explicit On 
Option Strict On
'******************************************************************************
'*
'* Name:        ProcessNull
'*
'* Description: Static class which provides functions to help process and handle
'*              NULL values from a database.
'*
'* Remarks:     The class includes methods which convert NULL values to 
'*              appropriate simple data types. For example:
'*                NULL  --> <blank> (string)
'*                NULL  --> 0       (integer)
'*                NULL  --> 0       (decimal)
'*              In some cases, this conversion may not be appropriate.
'* 
'*-----------------------------------------------------------------------------
'*                      CHANGE HISTORY
'*   Change No:   Date:          Author:   Description:
'*   _________    ___________    ______    ____________________________________
'*      001       18/03/2002     Custom Created.
'*      002       21/10/2003     Custom Added GetInt64.
'* 
'******************************************************************************
Public NotInheritable Class ProcessNull


#Region "Constructor"

    '**************************************************************************
    '*  
    '* Name:        New
    '*
    '* Description: Since this class provides only static methods, make the
    '*              default constructor private to prevent instances from being
    '*              created with "new ProcessNull()".
    '*
    '**************************************************************************
    Private Sub New()
    End Sub 'New

#End Region



#Region "Static Methods"

    Public Shared Function GetString(ByVal Value As Object) As String
        If IsDBNull(Value) Then
            Return ""
        Else
            Return CType(Value, String).TrimEnd()
        End If
    End Function 'GetString


    Public Shared Function GetInt16(ByVal Value As Object) As Int16
        If IsDBNull(Value) Then
            Return 0
        Else
            Return CType(Value, Int16)
        End If
    End Function 'GetInt16


    Public Shared Function GetInt32(ByVal Value As Object) As Int32
        If IsDBNull(Value) Then
            Return 0
        Else
            Return CType(Value, Int32)
        End If
    End Function 'GetInt32

    Public Shared Function GetInt64(ByVal Value As Object) As Int64
        If IsDBNull(Value) Then
            Return 0
        Else
            Return CType(Value, Int64)
        End If
    End Function 'GetInt64

    Public Shared Function GetDecimal(ByVal Value As Object) As Decimal
        If IsDBNull(Value) Then
            Return 0
        Else
            Return CType(Value, Decimal)
        End If
    End Function 'GetDecimal


    Public Shared Function GetDouble(ByVal Value As Object) As Double
        If IsDBNull(Value) Then
            Return 0
        Else
            Return CType(Value, Double)
        End If
    End Function 'GetDouble


    Public Shared Function GetFloat(ByVal Value As Object) As Single
        If IsDBNull(Value) Then
            Return 0
        Else
            Return CType(Value, Single)
        End If
    End Function 'GetFloat


    Public Shared Function GetDateTime(ByVal Value As Object) As DateTime
        If IsDBNull(Value) Then
            Return Nothing
        Else
            Return CType(Value, DateTime)
        End If
    End Function 'GetDateTime
    Public Shared Function GetDate(ByVal Value As Object) As Date
        If IsDBNull(Value) Then
            Return Nothing
        Else
            Return CType(Value, Date)
        End If
    End Function 'GetDate

    Public Shared Function GetSqlValue(ByVal Value As Object) As Object
        If IsDBNull(Value) Then
            Return Nothing
        Else
            Return Value
        End If
    End Function 'GetSqlValue


    Public Shared Function GetByte(ByVal Value As Object) As Byte
        If IsDBNull(Value) Then
            Return 0
        Else
            Return CType(Value, Byte)
        End If
    End Function 'GetByte

    Public Shared Function GetBytes(ByVal Value As Object) As Byte()
        If IsDBNull(Value) Then
            Return Nothing
        Else
            'Commented out code demonstrates reading an image from a database. However, the CType function seems to work just as well.
            'Dim myPicture(dr.GetBytes(dalCategories.CategoriesFields.fldPicture, 0, Nothing, 0, Integer.MaxValue) - 1) As Byte
            'dr.GetBytes(dalCategories.CategoriesFields.fldPicture, 0, myPicture, 0, myPicture.Length)
            'return myPicture
            Return CType(Value, Byte())
        End If
    End Function 'GetBytes

    Public Shared Function GetBoolean(ByVal Value As Object) As Boolean
        If IsDBNull(Value) Then
            Return False
        Else
            Return CType(Value, Boolean)
        End If
    End Function 'GetBoolean

    Public Shared Function GetGUID(ByVal Value As Object) As Guid
        If IsDBNull(Value) Then
            Return Nothing
        Else
            Return CType(Value, Guid)
        End If
    End Function 'GetGUID

#End Region


End Class 'ProcessNull




