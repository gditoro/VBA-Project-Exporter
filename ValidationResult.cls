VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'*******************************************************************************
' Name: ValidationResult
' Kind: Class Module
' Purpose: Encapsulates validation results with error messages
' Author: Giovanni Di Toro
' Date: 28-08-2025
'*******************************************************************************

Private m_isValid As Boolean
Private m_errors As Collection

Private Sub Class_Initialize()
    Set m_errors = New Collection
    m_isValid = True
End Sub

Public Property Get IsValid() As Boolean
    IsValid = m_isValid
End Property

Public Property Let IsValid(value As Boolean)
    m_isValid = value
End Property

Public Sub AddError(errorMessage As String)
    m_errors.Add errorMessage
    m_isValid = False
End Sub

Public Function GetErrors() As Collection
    Set GetErrors = m_errors
End Function

Public Function GetErrorsAsString() As String
    Dim result As String
    Dim i As Long

    For i = 1 To m_errors.Count
        If i > 1 Then result = result & vbNewLine
        result = result & "• " & m_errors(i)
    Next i

    GetErrorsAsString = result
End Function

Public Function HasErrors() As Boolean
    HasErrors = (m_errors.Count > 0)
End Function

Public Function ErrorCount() As Long
    ErrorCount = m_errors.Count
End Function