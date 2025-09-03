Attribute VB_Name = "StringHelpers"

Option Explicit

'*******************************************************************************
' Module: StringHelpers
' Kind: Standard Module
' Purpose: Provides string manipulation and validation utilities
' Author: Giovanni Di Toro
' Date: 28-08-2019
' Updated:
'         28-08-2025 - Enhanced performance and added new utility functions
'         01-09-2025 - Added Brazilian phone and CEP formatting functions
'*******************************************************************************

'*** ALWAYS START FUNCTIONS AND SUBS WITH THE PREFIX "str" ***

'***************************************************************************
'Purpose: Matches two(2) strings in their lowercase and trimmed forms.
'Inputs:
'   first_str:(String) First string to be evaluated.
'   second_str:(String) Second string to be evaluated and matched with the first.
'Outputs:(Boolean) True if the strings match or false if they don't.
'***************************************************************************

Function strMatch(first_str As String, second_str As String) As Boolean
    Dim Result As Boolean
    If Trim(LCase(first_str)) = Trim(LCase(second_str)) Then
        Result = True
    Else
        Result = False
    End If
    strMatch = Result
End Function

'***************************************************************************
'Purpose: Converts Strings of dates to Date type
'Inputs
'   data:(String) Date as string to be converted.
'Outputs: (Date) Converted date string to Date format
'***************************************************************************

Function strDate(data As String) As Date
    If data <> vbNullString Then
        strDate = Format(data, "dd-mm-yyyy")
        strDate = CDate(strDate)
    End If
End Function

'***************************************************************************
'Purpose: Converts string values to strings only containing numbers.
'Inputs
'   strIn:(String) String input to be converted
'   (Optional) punctuation:(Boolean) Takes punctuation in to account (True) or not (False).
'Outputs:(String) String that only contains numbers
'***************************************************************************

Function strNum(strIn As String, Optional punctuation As Boolean = False) As String
    If strIn <> vbNullString Then
        Dim objRegex
        Set objRegex = CreateObject("vbscript.regexp")
        With objRegex
             .Global = True
            If punctuation = True Then
                .pattern = "[^(?= ^\d)(?=^.)(?=^,)]+"
            Else
                .pattern = "[^\d]+"
            End If
            strNum = .Replace(strIn, vbNullString)
        End With
    Else
        strNum = 0
    End If
End Function

'***************************************************************************
'Purpose: Checks if  the String input is any value equivalent to "0" or empty string.
'Inputs
'   sData:(String) Value to be evaluated.
'Outputs:(Boolean) True for empty value or False for non-empty value.
'***************************************************************************

Function strIsNull(sData As String) As Boolean
    If sData = vbNullString Or sData = "00:00:00" Or sData = "0" Or sData = " " Then
        strIsNull = True
    Else
        strIsNull = False
    End If
End Function

'***************************************************************************
'Purpose: Returns either the value of the String input or a value of vbNullString.
'Inputs
'   sData:(String) Value to be evaluated.
'Outputs:(String) Value of the input data, if it's "0" equivalent it's transformed to vbNullString
'***************************************************************************

Function strRealEmpty(sData As String) As String
    If strIsNull(sData) Then
        strRealEmpty = ""
    Else
        strRealEmpty = sData
    End If
End Function

'***************************************************************************
'Purpose: Converts a string value to Brazilian currency format (consolidated from GlobalsGeral)
'Inputs
'   val:(String) Value to be converted
'   iDecPlaces:(Integer) Number of decimal places (optional, default 2)
'Outputs:(String) Formatted currency string
'***************************************************************************
Function strToCurrency(val As String, Optional iDecPlaces As Integer = 2) As String
    Dim logVal As String
    logVal = val
    Dim sAvn() As String, sDec As String
    Dim iPos As Integer

    iPos = Len(val) - 3
    val = strNum(val, True)

    If iPos < 1 Then iPos = 1

    If InStr(iPos, val, ".") > 0 Then
        val = Mid(val, 1, iPos - 1) & Replace(val, ".", ",", (Len(val) - 3))
    End If

    If InStr(1, val, ".") > 0 Then
        val = Replace(val, ".", "")
    End If

    val = Replace(val, ",", ".")
    sAvn = Split(val, ".")
    Application.ThousandsSeparator = ","
    sDec = "00"

    Dim decCalc As Single
    If arrayLength(sAvn) > 1 Then
        decCalc = Round(CSng("0," & sAvn(1)), iDecPlaces)
        If decCalc > 0 Then
            If InStr(1, CStr(decCalc), ",") > 0 Then
                sDec = Split(decCalc, ",")(1)
            ElseIf decCalc = 1 Then
                sDec = 0
                sAvn(0) = CStr(CSng(sAvn(0)) + 1)
            End If
        End If
    End If

    ' Ensure proper decimal formatting
    If Len(sDec) = 1 Then sDec = sDec & "0"

    strToCurrency = Format(sAvn(0), "#,##0") & "," & sDec
End Function

'***************************************************************************
'Purpose: Removes all non-numeric characters from a string
'Inputs
'   inputString:(String) String to clean
'Outputs: (String) String containing only numeric characters
'***************************************************************************
Public Function RemoveNonNumeric(inputString As String) As String
    Dim result As String
    Dim i As Long
    Dim char As String

    result = ""
    For i = 1 To Len(inputString)
        char = Mid(inputString, i, 1)
        If IsNumeric(char) Then
            result = result & char
        End If
    Next i

    RemoveNonNumeric = result
End Function

'***************************************************************************
' Function: FormatPhoneNumber
' Purpose: Formats Brazilian phone numbers to standard display format
' Inputs:
'   - phone (String): Raw phone number (with or without formatting)
' Outputs:
'   - String: Formatted phone number (XX) XXXXX-XXXX or (XX) XXXX-XXXX
' Notes: Handles both 10-digit (landline) and 11-digit (mobile) numbers
'***************************************************************************
Public Function FormatPhoneNumber(ByVal phone As String) As String
    Const PROC_NAME As String = "FormatPhoneNumber"

    Dim cleanPhone As String
    cleanPhone = strNumeric(phone)

    Select Case Len(cleanPhone)
        Case 11
            ' Mobile format: (XX) XXXXX-XXXX
            FormatPhoneNumber = "(" & Left(cleanPhone, 2) & ") " & Mid(cleanPhone, 3, 5) & "-" & Right(cleanPhone, 4)
        Case 10
            ' Landline format: (XX) XXXX-XXXX
            FormatPhoneNumber = "(" & Left(cleanPhone, 2) & ") " & Mid(cleanPhone, 3, 4) & "-" & Right(cleanPhone, 4)
        Case Else
            ' Invalid length - return original
            FormatPhoneNumber = phone
    End Select
End Function

'***************************************************************************
' Function: FormatCEP
' Purpose: Formats Brazilian postal code (CEP) to standard display format
' Inputs:
'   - cep (String): Raw CEP (with or without formatting)
' Outputs:
'   - String: Formatted CEP XXXXX-XXX
' Notes: Validates 8-digit CEP format
'***************************************************************************
Public Function FormatCEP(ByVal cep As String) As String
    Const PROC_NAME As String = "FormatCEP"

    Dim cleanCEP As String
    cleanCEP = strNumeric(cep)

    If Len(cleanCEP) = 8 Then
        ' Standard format: XXXXX-XXX
        FormatCEP = Left(cleanCEP, 5) & "-" & Right(cleanCEP, 3)
    Else
        ' Invalid length - return original
        FormatCEP = cep
    End If
End Function
