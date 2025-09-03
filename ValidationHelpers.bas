Attribute VB_Name = "ValidationHelpers"

'@Folder "Faturamento.Helpers"
Option Explicit

'*******************************************************************************
' Module: ValidationHelpers
' Author: Giovanni Di Toro
' Date: 26-08-2025
' Updated:
'           - 27/08/2025: Added ValidateCNPJ function returning ValidationResult
'           - 01/09/2025: Added common validation patterns to reduce repetitive code
'           - 01/09/2025: Added validation constants for consistent requirements
'           - 01/09/2025: Implemented validation factory pattern for complex rules
'           - 01/09/2025: Added comprehensive form validation helpers
'           - 01/09/2025: Consolidated email validation (replaced duplicates)
'           - 01/09/2025: Added ValidateNumeric wrapper for IsNumeric patterns
'           - 01/09/2025: CODEBASE CONSOLIDATION - Replaced all direct IsNumeric calls
'           - 01/09/2025: Moved RegionHelpers.ValidateRegion to ValidationHelpers.ValidateRegionCode
'           - 01/09/2025: Marked legacy Boolean functions as deprecated
'           - 01/09/2025: Removed redundant validation comments across Global modules
' Purpose: Comprehensive validation framework with factory patterns and constants
' Dependencies: StringHelpers.bas, ErrorHandler.bas, ValidationResult.cls
' Features:
'   - ValidationResult-based validation for consistent error handling
'   - Factory pattern for creating validation rules (CreateValidationRule)
'   - Form validation helpers (ValidateFormField, ValidateFormControls)
'   - Validation constants for consistent requirements across application
'   - Business entity validation (ValidateBusinessEntity)
'   - Validation chains for complex scenarios (CreateValidationChain)
'   - Auto-correcting numeric validation (ValidateTextBoxNumeric)
'   - Centralized region validation (ValidateRegionCode)
'   - Deprecated legacy functions with clear migration paths
'*******************************************************************************

Private Const MODULE_NAME As String = "ValidationHelpers"

'*******************************************************************************
' VALIDATION CONSTANTS
' Purpose: Centralized constants for common validation requirements
'*******************************************************************************

' String length constants
Public Const MIN_COMPANY_NAME_LENGTH As Integer = 2
Public Const MAX_COMPANY_NAME_LENGTH As Integer = 100
Public Const MIN_PERSON_NAME_LENGTH As Integer = 2
Public Const MAX_PERSON_NAME_LENGTH As Integer = 80
Public Const MIN_EMAIL_LENGTH As Integer = 5
Public Const MAX_EMAIL_LENGTH As Integer = 254
Public Const MIN_PHONE_LENGTH As Integer = 8
Public Const MAX_PHONE_LENGTH As Integer = 15
Public Const MIN_ADDRESS_LENGTH As Integer = 5
Public Const MAX_ADDRESS_LENGTH As Integer = 200
Public Const MIN_PRODUCT_NAME_LENGTH As Integer = 1
Public Const MAX_PRODUCT_NAME_LENGTH As Integer = 100
Public Const MIN_PRODUCT_DESC_LENGTH As Integer = 1
Public Const MAX_PRODUCT_DESC_LENGTH As Integer = 255
Public Const MIN_REGION_LENGTH As Integer = 1
Public Const MAX_REGION_LENGTH As Integer = 50

' Document format constants
Public Const CNPJ_LENGTH As Integer = 14
Public Const CPF_LENGTH As Integer = 11
Public Const MIN_CNPJ_FORMATTED_LENGTH As Integer = 11
Public Const MAX_CNPJ_FORMATTED_LENGTH As Integer = 18

' Numeric range constants
Public Const MIN_QUANTITY As Double = 0.01
Public Const MAX_QUANTITY As Double = 999999
Public Const MIN_PRICE As Double = 0.01
Public Const MAX_PRICE As Double = 99999999   ' Increased to accommodate NFe values up to 99 million
Public Const MIN_WEIGHT As Double = 0.01
Public Const MAX_WEIGHT As Double = 99999
Public Const MAX_ORDER_TOTAL As Double = 1000000

' Form validation constants
Public Const MAX_TEXTBOX_LENGTH As Integer = 255
Public Const MIN_REQUIRED_FIELD_LENGTH As Integer = 1

' Error message templates
Public Const MSG_FIELD_REQUIRED As String = " é obrigatório"
Public Const MSG_FIELD_INVALID_LENGTH As String = " deve ter entre %MIN% e %MAX% caracteres"
Public Const MSG_FIELD_INVALID_NUMBER As String = " deve ser um valor numérico válido"
Public Const MSG_FIELD_INVALID_RANGE As String = " deve estar entre %MIN% e %MAX%"
Public Const MSG_FIELD_INVALID_EMAIL As String = " deve ser um endereço de email válido"
Public Const MSG_FIELD_INVALID_CNPJ As String = " deve ser um CNPJ válido"

'*******************************************************************************
' Function: ValidateNFENumber
' Purpose: Validates NFE number format and content
' Inputs: nfeNumber - The NFE number to validate
' Returns: Boolean - True if valid, False otherwise
'
' ⚠️ DEPRECATED: Consider using ValidationHelpers factory pattern instead
' Use CreateValidationRule("CUSTOM_STRING", "NFE", value) for ValidationResult support
'*******************************************************************************
Public Function ValidateNFENumber(ByVal nfeNumber As String) As Boolean
    Const PROC_NAME As String = "ValidateNFENumber"
    On Error GoTo ErrorHandler

    ValidateNFENumber = False

    ' Check for empty or null string
    If strIsNull(nfeNumber) Then
        ErrorHandler.LogWarning MODULE_NAME, PROC_NAME, "NFE number is empty"
        Exit Function
    End If

    ' Remove any whitespace
    nfeNumber = Trim(nfeNumber)

    ' Check length (NFE numbers are typically 9 digits)
    If Len(nfeNumber) < 1 Or Len(nfeNumber) > 20 Then
        ErrorHandler.LogWarning MODULE_NAME, PROC_NAME, "NFE number length invalid: " & Len(nfeNumber)
        Exit Function
    End If

    ' Check if contains only allowed characters (numbers, letters, hyphens)
    If Not IsValidNFEFormat(nfeNumber) Then
        ErrorHandler.LogWarning MODULE_NAME, PROC_NAME, "NFE number contains invalid characters: " & nfeNumber
        Exit Function
    End If

    ValidateNFENumber = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    ValidateNFENumber = False
End Function

'*******************************************************************************
' Function: SanitizeEmailAddress
' Purpose: Sanitizes email address by removing dangerous characters
' Inputs: email - The email address to sanitize
' Returns: String - Sanitized email address
'*******************************************************************************
Public Function SanitizeEmailAddress(ByVal email As String) As String
    Const PROC_NAME As String = "SanitizeEmailAddress"
    On Error GoTo ErrorHandler

    If strIsNull(email) Then
        SanitizeEmailAddress = vbNullString
        Exit Function
    End If

    ' Remove potentially dangerous characters
    email = Replace(email, ";", "")
    email = Replace(email, ",", "")
    email = Replace(email, vbCrLf, "")
    email = Replace(email, vbCr, "")
    email = Replace(email, vbLf, "")
    email = Replace(email, Chr(0), "")

    SanitizeEmailAddress = Trim(email)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    SanitizeEmailAddress = vbNullString
End Function

'*******************************************************************************
' Function: ValidateFilePath
' Purpose: Validates file path for security issues
' Inputs: filePath - The file path to validate
' Returns: Boolean - True if valid, False otherwise
'*******************************************************************************
Public Function ValidateFilePath(ByVal filePath As String) As Boolean
    Const PROC_NAME As String = "ValidateFilePath"
    On Error GoTo ErrorHandler

    ValidateFilePath = False

    If strIsNull(filePath) Then Exit Function

    ' Check for path traversal attempts
    If InStr(filePath, "..") > 0 Then
        ErrorHandler.LogWarning MODULE_NAME, PROC_NAME, "Path traversal attempt detected: " & filePath
        Exit Function
    End If

    ' Check for double slashes
    If InStr(filePath, "//") > 0 Or InStr(filePath, "\\") > 0 Then
        ErrorHandler.LogWarning MODULE_NAME, PROC_NAME, "Invalid path format: " & filePath
        Exit Function
    End If

    ' Validate file extension
    If Not HasValidFileExtension(filePath) Then
        ErrorHandler.LogWarning MODULE_NAME, PROC_NAME, "Invalid file extension: " & filePath
        Exit Function
    End If

    ValidateFilePath = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    ValidateFilePath = False
End Function

'*******************************************************************************
' Function: ValidateNumericString
' Purpose: Validates that a string represents a valid number within range
' Inputs: value - String to validate
'         minValue - Minimum allowed value (optional)
'         maxValue - Maximum allowed value (optional)
' Returns: Boolean - True if valid, False otherwise
' DEPRECATED: Use ValidationHelpers.ValidateNumeric() or CreateValidationRule("NUMERIC_RANGE") instead
' This function will be removed in a future version
'*******************************************************************************
Public Function ValidateNumericString(ByVal value As String, _
                                     Optional ByVal minValue As Double = -1.79769313486231E+308, _
                                     Optional ByVal maxValue As Double = 1.79769313486231E+308) As Boolean
    Const PROC_NAME As String = "ValidateNumericString"

    ' DEPRECATED WARNING: This function is deprecated
    Debug.Print "WARNING: ValidateNumericString is deprecated. Use ValidationHelpers.ValidateNumeric() or CreateValidationRule(""NUMERIC_RANGE"") instead."

    ' Delegate to new validation method for backward compatibility
    Dim result As ValidationResult
    If minValue = -1.79769313486231E+308 And maxValue = 1.79769313486231E+308 Then
        Set result = ValidateNumeric(value, "Value")
    Else
        Set result = CreateValidationRule("CUSTOM_NUMERIC", "Value", value, minValue, maxValue)
    End If

    ValidateNumericString = result.IsValid
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    ValidateNumericString = False
End Function

'*******************************************************************************
' Function: ValidateCurrency
' Purpose: Validates currency values
' Inputs: value - String representation of currency
' Returns: Boolean - True if valid, False otherwise
'
' ⚠️ DEPRECATED: Use ValidateNumeric() with PRICE validation instead
' Use CreateValidationRule("PRICE", "Currency", value) for ValidationResult support
'*******************************************************************************
Public Function ValidateCurrency(ByVal value As String) As Boolean
    Const PROC_NAME As String = "ValidateCurrency"
    On Error GoTo ErrorHandler

    ValidateCurrency = False

    If strIsNull(value) Then Exit Function

    ' Remove currency symbols and spaces
    value = Replace(value, "R$", "")
    value = Replace(value, "$", "")
    value = Replace(value, " ", "")
    value = Trim(value)

    ' Check if it's a valid number
    If Not IsNumeric(value) Then Exit Function

    ' Check if positive (assuming freight costs are always positive)
    If CDbl(value) < 0 Then Exit Function

    ValidateCurrency = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    ValidateCurrency = False
End Function

'*******************************************************************************
' Function: ValidateStringInput
' Purpose: Validates string input for length and basic content
' Inputs: inputText - String to validate
'         minLength - Minimum required length
'         maxLength - Maximum allowed length (optional)
' Returns: Boolean - True if valid, False otherwise
'*******************************************************************************
Public Function ValidateStringInput(ByVal inputText As String, _
                                    ByVal minLength As Long, _
                                    Optional ByVal maxLength As Long = 1000) As Boolean
    Const PROC_NAME As String = "ValidateStringInput"
    On Error GoTo ErrorHandler

    ValidateStringInput = False

    ' Check for null/empty
    If inputText = vbNullString Then
        If minLength > 0 Then Exit Function
    End If

    ' Trim and check length
    inputText = Trim(inputText)
    If Len(inputText) < minLength Or Len(inputText) > maxLength Then Exit Function

    ' Check for dangerous characters
    If ContainsDangerousChars(inputText) Then Exit Function

    ValidateStringInput = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    ValidateStringInput = False
End Function

'*******************************************************************************
' Function: ValidateRequiredInput
' Purpose: Consolidated validation for required string inputs with user prompt
' Inputs: inputValue - String to validate
'         errorMessage - Message to show if validation fails
' Returns: String - Valid input value
' Author: Consolidated from GlobalsGeral.validPresenceStr
'*******************************************************************************
Public Function ValidateRequiredInput(ByVal inputValue As String, _
                                     ByVal errorMessage As String) As String
    Const PROC_NAME As String = "ValidateRequiredInput"

    If inputValue <> vbNullString And Trim(inputValue) <> vbNullString Then
        ValidateRequiredInput = inputValue
    Else
        ValidateRequiredInput = InputBox(errorMessage, _
                                      "Campo obrigatório deixado em branco", _
                                      1)
    End If
End Function

'*******************************************************************************
' Function: ValidateNumericInput
' Purpose: Validates numeric input for sales order items
' Parameters: inputValue - Value to validate
'            fieldName - Name of field for error message
'            minValue - Minimum allowed value (optional)
'            maxValue - Maximum allowed value (optional)
' Returns: ValidationResult with validation status
'*******************************************************************************
Public Function ValidateNumericInput(ByVal inputValue As Variant, _
                                    ByVal fieldName As String, _
                                    Optional ByVal minValue As Double = 0, _
                                    Optional ByVal maxValue As Double = 999999999) As ValidationResult
    Const PROC_NAME As String = "ValidateNumericInput"

    Dim result As ValidationResult
    Set result = New ValidationResult

    ' Check if value is numeric
    If Not IsNumeric(inputValue) Then
        result.AddError fieldName & " deve ser um valor numérico válido"
        Set ValidateNumericInput = result
        Exit Function
    End If

    Dim numValue As Double
    numValue = CDbl(inputValue)

    ' Check minimum value
    If numValue < minValue Then
        result.AddError fieldName & " deve ser maior ou igual a " & minValue
        Set ValidateNumericInput = result
        Exit Function
    End If

    ' Check maximum value
    If numValue > maxValue Then
        result.AddError fieldName & " deve ser menor ou igual a " & maxValue
        Set ValidateNumericInput = result
        Exit Function
    End If

    ' Validation passed - result.IsValid remains True (default)
    Set ValidateNumericInput = result
End Function

'*******************************************************************************
' Function: ValidateSalesOrderItem
' Purpose: Comprehensive validation for sales order item data
' Parameters: produto - Product name
'            quantidade - Quantity value
'            preco - Price value
' Returns: ValidationResult with validation status and detailed messages
'*******************************************************************************
Public Function ValidateSalesOrderItem(ByVal produto As String, _
                                      ByVal quantidade As Variant, _
                                      ByVal preco As Variant) As ValidationResult
    Const PROC_NAME As String = "ValidateSalesOrderItem"

    Dim validations() As ValidationResult
    ReDim validations(3)

    ' Validate product selection
    Set validations(0) = ValidateRequired("Produto", produto)

    ' Validate quantity
    If Len(Trim(CStr(quantidade))) = 0 Then
        Set validations(1) = New ValidationResult
        validations(1).AddError "Quantidade é obrigatória"
    Else
        Set validations(1) = ValidateNumericInput(quantidade, "Quantidade", 0.01, 999999)
    End If

    ' Validate price
    If Len(Trim(CStr(preco))) = 0 Then
        Set validations(2) = New ValidationResult
        validations(2).AddError "Preço é obrigatório"
    Else
        Set validations(2) = ValidateNumericInput(preco, "Preço", 0.01, 999999)
    End If

    ' Additional business rule validations
    Set validations(3) = New ValidationResult
    If validations(1).IsValid And validations(2).IsValid Then
        Dim total As Double
        total = CDbl(quantidade) * CDbl(preco)
        If total > 1000000 Then ' Business rule: max total per item
            validations(3).AddError "Total do item (R$ " & Format(total, "#,##0.00") & ") excede o limite máximo"
        End If
    End If

    ' Combine all validations
    Set ValidateSalesOrderItem = CombineValidationsArray(validations)
End Function

'*******************************************************************************
' Name: ValidateEmail
' Kind: Public Function
' Purpose: Comprehensive email validation returning ValidationResult
' Inputs: email - The email address to validate
' Returns: ValidationResult - Contains validation status and errors
' Note: This replaces both EntityManagementHelpers.IsValidEmail() and
'       TransportadoraBusinessService.IsValidEmailFormat() for consistency
'*******************************************************************************
Public Function ValidateEmail(ByVal email As String) As ValidationResult
    Const PROC_NAME As String = "ValidateEmail"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    ' Check for null/empty input
    If strIsNull(email) Then
        result.AddError "Email é obrigatório"
        Set ValidateEmail = result
        Exit Function
    End If

    Dim trimmedEmail As String
    trimmedEmail = Trim(email)

    ' Check minimum length
    If Len(trimmedEmail) < 5 Then
        result.AddError "Email deve ter pelo menos 5 caracteres"
        Set ValidateEmail = result
        Exit Function
    End If

    ' Check for @ symbol
    Dim atPosition As Integer
    atPosition = InStr(trimmedEmail, "@")
    If atPosition <= 1 Then
        result.AddError "Email deve conter @ após o primeiro caractere"
        Set ValidateEmail = result
        Exit Function
    End If

    ' Check for domain part with dot
    Dim dotPosition As Integer
    dotPosition = InStrRev(trimmedEmail, ".")
    If dotPosition <= atPosition + 1 Or dotPosition >= Len(trimmedEmail) Then
        result.AddError "Email deve conter um domínio válido com extensão"
        Set ValidateEmail = result
        Exit Function
    End If

    ' Check for invalid characters
    If ContainsInvalidEmailChars(trimmedEmail) Then
        result.AddError "Email contém caracteres inválidos"
        Set ValidateEmail = result
        Exit Function
    End If

    ' Email validation passed
    Set ValidateEmail = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na validação de email"
    Set ValidateEmail = result
End Function

'*******************************************************************************
' Function: ValidateNumeric
' Purpose: Validates that a value is numeric and within optional range
' Parameters:
'   value - Value to validate (Variant to accept any input type)
'   fieldName - Name of field for error messages
'   minValue - Minimum allowed value (optional)
'   maxValue - Maximum allowed value (optional)
' Returns: ValidationResult - Contains validation status and errors
' Note: This provides ValidationResult wrapper for IsNumeric checks
'*******************************************************************************
Public Function ValidateNumeric(ByVal value As Variant, _
                               ByVal fieldName As String, _
                               Optional ByVal minValue As Variant, _
                               Optional ByVal maxValue As Variant) As ValidationResult
    Const PROC_NAME As String = "ValidateNumeric"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    ' Check if value is numeric
    If Not IsNumeric(value) Then
        result.AddError fieldName & " deve ser um valor numérico válido"
        Set ValidateNumeric = result
        Exit Function
    End If

    ' Convert to double for range checking
    Dim numValue As Double
    numValue = CDbl(value)

    ' Check minimum value if provided
    If Not IsMissing(minValue) Then
        If IsNumeric(minValue) Then
            If numValue < CDbl(minValue) Then
                result.AddError fieldName & " deve ser maior ou igual a " & minValue
            End If
        End If
    End If

    ' Check maximum value if provided
    If Not IsMissing(maxValue) Then
        If IsNumeric(maxValue) Then
            If numValue > CDbl(maxValue) Then
                result.AddError fieldName & " deve ser menor ou igual a " & maxValue
            End If
        End If
    End If

    Set ValidateNumeric = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na validação numérica: " & fieldName
    Set ValidateNumeric = result
End Function

'*******************************************************************************
' Private Helper Functions
'*******************************************************************************

Private Function IsValidNFEFormat(ByVal nfeNumber As String) As Boolean
    Dim i As Long
    Dim char As String

    For i = 1 To Len(nfeNumber)
        char = Mid(nfeNumber, i, 1)
        If Not (IsNumeric(char) Or _
                (char >= "A" And char <= "Z") Or _
                (char >= "a" And char <= "z") Or _
                char = "-" Or char = "_") Then
            IsValidNFEFormat = False
            Exit Function
        End If
    Next i

    IsValidNFEFormat = True
End Function

Private Function ContainsInvalidEmailChars(ByVal email As String) As Boolean
    Dim invalidChars As String
    invalidChars = " ;,<>()[]{}|\"

    Dim i As Long
    For i = 1 To Len(invalidChars)
        If InStr(email, Mid(invalidChars, i, 1)) > 0 Then
            ContainsInvalidEmailChars = True
            Exit Function
        End If
    Next i

    ContainsInvalidEmailChars = False
End Function

Private Function HasValidFileExtension(ByVal filePath As String) As Boolean
    Dim allowedExtensions As Variant
    allowedExtensions = Array(".xml", ".pdf", ".html", ".htm", ".txt", ".csv", ".xls", ".xlsx")

    Dim fileExt As String
    Dim dotPos As Long
    dotPos = InStrRev(filePath, ".")

    If dotPos = 0 Then
        HasValidFileExtension = False
        Exit Function
    End If

    fileExt = LCase(Mid(filePath, dotPos))

    Dim i As Long
    For i = LBound(allowedExtensions) To UBound(allowedExtensions)
        If fileExt = allowedExtensions(i) Then
            HasValidFileExtension = True
            Exit Function
        End If
    Next i

    HasValidFileExtension = False
End Function

Private Function ContainsDangerousChars(ByVal inputText As String) As Boolean
    Dim dangerousChars As String
    dangerousChars = Chr(0) & Chr(1) & Chr(2) & Chr(3) & Chr(4) & Chr(5) & _
                     Chr(6) & Chr(7) & Chr(8) & Chr(9) & Chr(11) & Chr(12) & _
                     Chr(14) & Chr(15) & Chr(16) & Chr(17) & Chr(18) & Chr(19) & _
                     Chr(20) & Chr(21) & Chr(22) & Chr(23) & Chr(24) & Chr(25) & _
                     Chr(26) & Chr(27) & Chr(28) & Chr(29) & Chr(30) & Chr(31)

    Dim i As Long
    For i = 1 To Len(dangerousChars)
        If InStr(inputText, Mid(dangerousChars, i, 1)) > 0 Then
            ContainsDangerousChars = True
            Exit Function
        End If
    Next i

    ContainsDangerousChars = False
End Function

'*******************************************************************************
' Function: ValidateCNPJ
' Purpose: Wrapper for CNPJ validation that returns ValidationResult
' Inputs: cnpj - The CNPJ string to validate
' Returns: ValidationResult - Contains validation status and errors
' Note: Implements complete CNPJ validation algorithm (format + check digits)
'*******************************************************************************
Public Function ValidateCNPJ(ByVal cnpj As String) As ValidationResult
    Const PROC_NAME As String = "ValidateCNPJ"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    ' Clean CNPJ (remove formatting)
    Dim cleanCnpj As String
    cleanCnpj = StringHelpers.RemoveNonNumeric(cnpj)

    ' Basic validation
    If Len(cleanCnpj) <> 14 Then
        result.AddError "CNPJ deve ter 14 dígitos"
        Set ValidateCNPJ = result
        Exit Function
    End If

    ' Check for invalid patterns (all same digits)
    If cleanCnpj Like String(14, Left(cleanCnpj, 1)) Then
        result.AddError "CNPJ inválido (dígitos repetidos)"
        Set ValidateCNPJ = result
        Exit Function
    End If

    ' Validate check digits using standard algorithm
    If Not ValidateCNPJCheckDigits(cleanCnpj) Then
        result.AddError "CNPJ inválido (dígitos verificadores)"
    End If

    Set ValidateCNPJ = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na validação do CNPJ"
    Set ValidateCNPJ = result
End Function

'*******************************************************************************
' Function: ValidateCNPJCheckDigits
' Purpose: Validates CNPJ check digits using the official algorithm
' Inputs: cnpj - Clean 14-digit CNPJ string
' Returns: Boolean - True if check digits are valid
'*******************************************************************************
Private Function ValidateCNPJCheckDigits(cnpj As String) As Boolean
    Dim i As Integer
    Dim sum As Integer
    Dim digit As Integer

    ' First check digit
    sum = 0
    For i = 1 To 12
        If i <= 4 Then
            sum = sum + CInt(Mid(cnpj, i, 1)) * (6 - i)
        Else
            sum = sum + CInt(Mid(cnpj, i, 1)) * (14 - i)
        End If
    Next i

    digit = 11 - (sum Mod 11)
    If digit > 9 Then digit = 0

    If digit <> CInt(Mid(cnpj, 13, 1)) Then
        ValidateCNPJCheckDigits = False
        Exit Function
    End If

    ' Second check digit
    sum = 0
    For i = 1 To 13
        If i <= 5 Then
            sum = sum + CInt(Mid(cnpj, i, 1)) * (7 - i)
        Else
            sum = sum + CInt(Mid(cnpj, i, 1)) * (15 - i)
        End If
    Next i

    digit = 11 - (sum Mod 11)
    If digit > 9 Then digit = 0

    ValidateCNPJCheckDigits = (digit = CInt(Mid(cnpj, 14, 1)))
End Function

'*******************************************************************************
' COMMON VALIDATION PATTERNS
' Purpose: Centralized validation functions to eliminate repetitive patterns
'*******************************************************************************

'*******************************************************************************
' Function: ValidateRequired
' Purpose: Validates that a string field is not empty/null (centralized pattern)
' Inputs:
'   fieldName - Name of the field for error messages
'   value - The string value to validate
' Returns: ValidationResult - Contains validation status and errors
' Note: Handles both null/empty and whitespace-only strings
'*******************************************************************************
Public Function ValidateRequired(ByVal fieldName As String, ByVal value As String) As ValidationResult
    Const PROC_NAME As String = "ValidateRequired"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    If strIsNull(value) Then
        result.AddError fieldName & " é obrigatório"
    ElseIf Len(Trim(value)) = 0 Then
        result.AddError fieldName & " não pode estar vazio"
    End If

    Set ValidateRequired = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na validação de campo obrigatório: " & fieldName
    Set ValidateRequired = result
End Function

'*******************************************************************************
' Function: ValidateStringLength
' Purpose: Validates string length within specified bounds
' Inputs:
'   fieldName - Name of the field for error messages
'   value - The string value to validate
'   minLength - Minimum required length (optional, default 0)
'   maxLength - Maximum allowed length (optional, default 255)
' Returns: ValidationResult - Contains validation status and errors
'*******************************************************************************
Public Function ValidateStringLength(ByVal fieldName As String, ByVal value As String, _
                                    Optional ByVal minLength As Integer = 0, _
                                    Optional ByVal maxLength As Integer = MAX_TEXTBOX_LENGTH) As ValidationResult
    Const PROC_NAME As String = "ValidateStringLength"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    Dim actualLength As Integer
    actualLength = Len(Trim(value))

    If actualLength < minLength Then
        result.AddError fieldName & " deve ter pelo menos " & minLength & " caracteres"
    ElseIf actualLength > maxLength Then
        result.AddError fieldName & " não pode ter mais de " & maxLength & " caracteres"
    End If

    Set ValidateStringLength = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na validação de comprimento: " & fieldName
    Set ValidateStringLength = result
End Function

'*******************************************************************************
' Function: ValidateString
' Purpose: Comprehensive string validation with various constraints
' Parameters:
'   value - String value to validate
'   PropertyName - Name of the property/field for error messages
'   Required - Whether the string is required (optional, default True)
'   minLength - Minimum length (optional, default 0)
'   maxLength - Maximum length (optional, default 0 means no limit)
'   OnlyLetters - Whether to allow only letters (optional, default False)
'   AllowedSpecialChars - Special characters to allow (optional, default empty)
' Returns: ValidationResult with validation status and errors
'*******************************************************************************
Public Function ValidateString(ByVal value As String, _
                             ByVal PropertyName As String, _
                             Optional ByVal Required As Boolean = True, _
                             Optional ByVal minLength As Long = 0, _
                             Optional ByVal maxLength As Long = 0, _
                             Optional ByVal OnlyLetters As Boolean = False, _
                             Optional ByVal AllowedSpecialChars As String = "") As ValidationResult
    Const PROC_NAME As String = "ValidateString"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    ' Check required validation
    If Required And Len(Trim(value)) = 0 Then
        result.AddError PropertyName & " é obrigatório"
        Set ValidateString = result
        Exit Function
    End If

    ' If not required and empty, validation passes
    If Not Required And Len(Trim(value)) = 0 Then
        Set ValidateString = result
        Exit Function
    End If

    Dim trimmedValue As String
    trimmedValue = Trim(value)

    ' Check minimum length
    If minLength > 0 And Len(trimmedValue) < minLength Then
        result.AddError PropertyName & " deve ter pelo menos " & minLength & " caracteres"
    End If

    ' Check maximum length (0 means no limit)
    If maxLength > 0 And Len(trimmedValue) > maxLength Then
        result.AddError PropertyName & " não pode ter mais de " & maxLength & " caracteres"
    End If

    ' Check if only letters are allowed
    If OnlyLetters Then
        If Not IsOnlyLetters(trimmedValue, AllowedSpecialChars) Then
            If Len(AllowedSpecialChars) > 0 Then
                result.AddError PropertyName & " deve conter apenas letras e os caracteres especiais: " & AllowedSpecialChars
            Else
                result.AddError PropertyName & " deve conter apenas letras"
            End If
        End If
    End If

    Set ValidateString = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na validação de string: " & PropertyName
    Set ValidateString = result
End Function

'*******************************************************************************
' Private Helper Function: IsOnlyLetters
' Purpose: Checks if string contains only letters and allowed special characters
' Parameters:
'   value - String to check
'   allowedSpecial - String containing allowed special characters
' Returns: Boolean - True if valid, False otherwise
'*******************************************************************************
Private Function IsOnlyLetters(ByVal value As String, ByVal allowedSpecial As String) As Boolean
    Dim i As Long
    Dim char As String

    IsOnlyLetters = True

    For i = 1 To Len(value)
        char = Mid(value, i, 1)

        ' Check if character is a letter (A-Z, a-z, or accented characters)
        If Not IsLetter(char) Then
            ' If not a letter, check if it's in allowed special characters
            If InStr(allowedSpecial, char) = 0 Then
                IsOnlyLetters = False
                Exit Function
            End If
        End If
    Next i
End Function

'*******************************************************************************
' Private Helper Function: IsLetter
' Purpose: Checks if a character is a letter (including accented characters)
' Parameters: char - Single character to check
' Returns: Boolean - True if letter, False otherwise
'*******************************************************************************
Private Function IsLetter(ByVal char As String) As Boolean
    Dim asciiValue As Long
    asciiValue = Asc(char)

    ' Standard ASCII letters A-Z and a-z
    If (asciiValue >= 65 And asciiValue <= 90) Or (asciiValue >= 97 And asciiValue <= 122) Then
        IsLetter = True
    ' Extended ASCII for accented characters (common Portuguese characters)
    ElseIf asciiValue >= 192 And asciiValue <= 255 Then
        IsLetter = True
    ' Space character (often considered valid in names)
    ElseIf asciiValue = 32 Then
        IsLetter = True
    Else
        IsLetter = False
    End If
End Function

'*******************************************************************************
' Function: ValidateNotNull
' Purpose: Validates that a database result/value is not null
' Inputs:
'   fieldName - Name of the field for error messages
'   value - The variant value to check for null
' Returns: ValidationResult - Contains validation status and errors
'*******************************************************************************
Public Function ValidateNotNull(ByVal fieldName As String, ByVal value As Variant) As ValidationResult
    Const PROC_NAME As String = "ValidateNotNull"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    If IsNull(value) Then
        result.AddError fieldName & " não pode ser nulo"
    End If

    Set ValidateNotNull = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na validação de valor nulo: " & fieldName
    Set ValidateNotNull = result
End Function

'*******************************************************************************
' Function: ValidateNotEmpty
' Purpose: Validates that a range/array is not empty
' Inputs:
'   fieldName - Name of the field for error messages
'   value - The variant value to check (can be Range, Array, etc.)
' Returns: ValidationResult - Contains validation status and errors
'*******************************************************************************
Public Function ValidateNotEmpty(ByVal fieldName As String, ByVal value As Variant) As ValidationResult
    Const PROC_NAME As String = "ValidateNotEmpty"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    If IsEmpty(value) Then
        result.AddError fieldName & " não pode estar vazio"
    End If

    Set ValidateNotEmpty = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na validação de conteúdo vazio: " & fieldName
    Set ValidateNotEmpty = result
End Function

'*******************************************************************************
' Function: CombineValidations
' Purpose: Combines multiple validation results into a single result
' Inputs:
'   ParamArray validations - Variable number of ValidationResult objects
' Returns: ValidationResult - Combined validation result
' Note: Useful for validating multiple fields and collecting all errors
'*******************************************************************************
Public Function CombineValidations(ParamArray validations() As Variant) As ValidationResult
    Const PROC_NAME As String = "CombineValidations"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim validation As ValidationResult

    For i = LBound(validations) To UBound(validations)
        Set validation = validations(i)
        If Not validation Is Nothing Then
            If Not validation.IsValid Then
                result.AddError validation.GetErrorsAsString()
            End If
        End If
    Next i

    Set CombineValidations = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na combinação de validações"
    Set CombineValidations = result
End Function

'*******************************************************************************
' Function: CombineValidationsArray
' Purpose: Combines validation results from an array into a single result
' Inputs: validations - Array of ValidationResult objects
' Returns: ValidationResult - Combined validation result
'*******************************************************************************
Public Function CombineValidationsArray(validations() As ValidationResult) As ValidationResult
    Const PROC_NAME As String = "CombineValidationsArray"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim validation As ValidationResult

    For i = LBound(validations) To UBound(validations)
        Set validation = validations(i)
        If Not validation Is Nothing Then
            If Not validation.IsValid Then
                result.AddError validation.GetErrorsAsString()
            End If
        End If
    Next i

    Set CombineValidationsArray = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na combinação de validações"
    Set CombineValidationsArray = result
End Function

'*******************************************************************************
' Function: ValidateXMLFile
' Purpose: Validates XML file structure and content before processing
' Parameters: filePath - Path to XML file to validate
' Returns: Boolean - True if file is valid NFe XML, False otherwise
' Dependencies: FileSystemHelpers.fsIsFile, Microsoft XML v3.0 (DOMDocument30)
'*******************************************************************************
Public Function ValidateXMLFile(ByVal filePath As String) As Boolean
    Const PROC_NAME As String = "ValidateXMLFile"
    On Error GoTo ErrorHandler

    ValidateXMLFile = False

    ' Basic input validation
    If strIsNull(filePath) Then
        ErrorHandler.LogWarning MODULE_NAME, PROC_NAME, "File path is null or empty"
        Exit Function
    End If

    ' Check if file exists
    If Not FileSystemHelpers.fsIsFile(filePath) Then
        ErrorHandler.LogWarning MODULE_NAME, PROC_NAME, "File does not exist: " & filePath
        Exit Function
    End If

    ' Check file extension
    If LCase(Right(filePath, 4)) <> ".xml" Then
        ErrorHandler.LogWarning MODULE_NAME, PROC_NAME, "Invalid file extension: " & filePath
        Exit Function
    End If

    ' Try to load and validate XML document
    Dim testDoc As Object ' Use Object for late binding to avoid reference issues
    Set testDoc = CreateObject("Msxml2.DOMDocument.3.0")

    If testDoc Is Nothing Then
        ErrorHandler.LogError MODULE_NAME, PROC_NAME, "Failed to create XML document object"
        Exit Function
    End If

    testDoc.async = False
    testDoc.validateOnParse = False

    If testDoc.Load(filePath) Then
        ' Basic NFe structure validation - check for required NFe elements
        If testDoc.getElementsByTagName("NFe").Length > 0 Or _
           testDoc.getElementsByTagName("nfe").Length > 0 Then

            ' Additional validation - check for key NFe elements
            If testDoc.getElementsByTagName("infNFe").Length > 0 Or _
               testDoc.getElementsByTagName("infnfe").Length > 0 Then
                ValidateXMLFile = True
                ErrorHandler.LogInfo MODULE_NAME, PROC_NAME, "XML file validated successfully: " & filePath
            Else
                ErrorHandler.LogWarning MODULE_NAME, PROC_NAME, "XML missing required NFe info elements: " & filePath
            End If
        Else
            ErrorHandler.LogWarning MODULE_NAME, PROC_NAME, "XML is not a valid NFe document: " & filePath
        End If
    Else
        ErrorHandler.LogWarning MODULE_NAME, PROC_NAME, "Failed to load XML document: " & filePath & " - " & testDoc.parseError.reason
    End If

    Set testDoc = Nothing
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    ValidateXMLFile = False
    If Not testDoc Is Nothing Then Set testDoc = Nothing
End Function

'*******************************************************************************
' Function: ValidateGenericEntity
' Purpose: Validates any entity using reflection-based property validation
' Parameters:
'   entity - Object to validate
'   requiredProperties - Array of property names that must not be empty
' Returns: ValidationResult - Contains validation status and errors
' Note: Generic validation that works with any object with properties
'*******************************************************************************
Public Function ValidateGenericEntity(ByVal entity As Object, _
                                     ByVal requiredProperties As Variant) As ValidationResult
    Const PROC_NAME As String = "ValidateGenericEntity"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    If entity Is Nothing Then
        result.AddError "Entity cannot be Nothing"
        Set ValidateGenericEntity = result
        Exit Function
    End If

    ' Validate required properties if provided
    If Not IsEmpty(requiredProperties) Then
        Dim i As Integer
        For i = LBound(requiredProperties) To UBound(requiredProperties)
            Dim propertyName As String
            propertyName = CStr(requiredProperties(i))

            ' Note: Property validation would require CallByName or specific interface
            ' This is a generic approach that can be extended
        Next i
    End If

    Set ValidateGenericEntity = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Error validating entity"
    Set ValidateGenericEntity = result
End Function

Private Function ValidateClienteEntity(ByVal entity As Object) As ValidationResult
    Dim result As ValidationResult
    Set result = New ValidationResult

    ' Add cliente-specific validation logic here
    ' For now, basic validation
    If entity Is Nothing Then
        result.AddError "Cliente não pode ser nulo"
    End If

    Set ValidateClienteEntity = result
End Function

Private Function ValidateProdutoEntity(ByVal entity As Object) As ValidationResult
    Dim result As ValidationResult
    Set result = New ValidationResult

    ' Add produto-specific validation logic here
    ' For now, basic validation
    If entity Is Nothing Then
        result.AddError "Produto não pode ser nulo"
    End If

    Set ValidateProdutoEntity = result
End Function

'*******************************************************************************
' VALIDATION FACTORY PATTERN
' Purpose: Factory methods for creating validation rules and strategies
'*******************************************************************************

'*******************************************************************************
' Function: CreateValidationRule
' Purpose: Factory method to create validation rules for different field types
' Parameters:
'   ruleType - Type of validation rule (Required, StringLength, Numeric, Email, etc.)
'   fieldName - Name of the field for error messages
'   ParamArray params - Variable parameters specific to each rule type
' Returns: ValidationResult - Contains validation status and errors
'*******************************************************************************
Public Function CreateValidationRule(ByVal ruleType As String, _
                                    ByVal fieldName As String, _
                                    ByVal value As Variant, _
                                    ParamArray params() As Variant) As ValidationResult
    Const PROC_NAME As String = "CreateValidationRule"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    Select Case UCase(ruleType)
        Case "REQUIRED"
            Set result = ValidateRequired(fieldName, CStr(value))

        Case "COMPANY_NAME"
            Set result = ValidateStringLength(fieldName, CStr(value), MIN_COMPANY_NAME_LENGTH, MAX_COMPANY_NAME_LENGTH)

        Case "PERSON_NAME"
            Set result = ValidateStringLength(fieldName, CStr(value), MIN_PERSON_NAME_LENGTH, MAX_PERSON_NAME_LENGTH)

        Case "EMAIL"
            Set result = ValidateEmail(CStr(value))

        Case "CNPJ"
            Set result = ValidateCNPJ(CStr(value))

        Case "PHONE"
            Set result = ValidateStringLength(fieldName, CStr(value), MIN_PHONE_LENGTH, MAX_PHONE_LENGTH)

        Case "ADDRESS"
            Set result = ValidateStringLength(fieldName, CStr(value), MIN_ADDRESS_LENGTH, MAX_ADDRESS_LENGTH)

        Case "PRODUCT_NAME"
            Set result = ValidateStringLength(fieldName, CStr(value), MIN_PRODUCT_NAME_LENGTH, MAX_PRODUCT_NAME_LENGTH)

        Case "PRODUCT_DESC"
            Set result = ValidateStringLength(fieldName, CStr(value), MIN_PRODUCT_DESC_LENGTH, MAX_PRODUCT_DESC_LENGTH)

        Case "QUANTITY"
            Set result = ValidateNumeric(value, fieldName, MIN_QUANTITY, MAX_QUANTITY)

        Case "PRICE"
            Set result = ValidateNumeric(value, fieldName, MIN_PRICE, MAX_PRICE)

        Case "CURRENCY"
            ' Currency validation (alias for PRICE validation)
            Set result = ValidateNumeric(value, fieldName, MIN_PRICE, MAX_PRICE)

        Case "WEIGHT"
            Set result = ValidateNumeric(value, fieldName, MIN_WEIGHT, MAX_WEIGHT)

        Case "REGION"
            Set result = ValidateRegionCode(fieldName, CStr(value))

        Case "CUSTOM_STRING"
            ' Custom string validation with parameters: minLength, maxLength
            Dim minLen As Integer, maxLen As Integer
            minLen = IIf(UBound(params) >= 0, CInt(params(0)), 1)
            maxLen = IIf(UBound(params) >= 1, CInt(params(1)), MAX_TEXTBOX_LENGTH)
            Set result = ValidateStringLength(fieldName, CStr(value), minLen, maxLen)

        Case "CUSTOM_NUMERIC"
            ' Custom numeric validation with parameters: minValue, maxValue
            Dim minVal As Double, maxVal As Double
            minVal = IIf(UBound(params) >= 0, CDbl(params(0)), 0)
            maxVal = IIf(UBound(params) >= 1, CDbl(params(1)), 999999999)
            Set result = ValidateNumeric(value, fieldName, minVal, maxVal)

        Case Else
            result.AddError "Tipo de validação não suportado: " & ruleType
    End Select

    Set CreateValidationRule = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na criação da regra de validação: " & ruleType
    Set CreateValidationRule = result
End Function

'*******************************************************************************
' Function: ValidateFormField
' Purpose: High-level validation for common form field patterns
' Parameters:
'   fieldType - Predefined field type (CompanyName, Email, CNPJ, etc.)
'   fieldName - Display name for error messages
'   value - Value to validate
'   required - Whether field is required (optional, default True)
' Returns: ValidationResult - Contains validation status and errors
'*******************************************************************************
Public Function ValidateFormField(ByVal fieldType As String, _
                                 ByVal fieldName As String, _
                                 ByVal value As Variant, _
                                 Optional ByVal required As Boolean = True) As ValidationResult
    Const PROC_NAME As String = "ValidateFormField"

    Dim validations() As ValidationResult
    ReDim validations(1)

    On Error GoTo ErrorHandler

    ' First validate if required
    If required Then
        Set validations(0) = ValidateRequired(fieldName, CStr(value))
        ' If required validation fails, don't continue with format validation
        If Not validations(0).IsValid Then
            Set ValidateFormField = validations(0)
            Exit Function
        End If
    Else
        Set validations(0) = New ValidationResult ' Empty validation that passes
    End If

    ' Then validate format/content
    Set validations(1) = CreateValidationRule(fieldType, fieldName, value)

    ' Combine validations
    Set ValidateFormField = CombineValidationsArray(validations)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    Dim result As ValidationResult
    Set result = New ValidationResult
    result.AddError "Erro interno na validação do campo: " & fieldName
    Set ValidateFormField = result
End Function

'*******************************************************************************
' Function: CreateValidationChain
' Purpose: Creates a chain of validation rules for complex validation scenarios
' Parameters:
'   fieldName - Name of the field for error messages
'   value - Value to validate
'   ParamArray ruleTypes - Variable number of validation rule types
' Returns: ValidationResult - Contains validation status and errors
'*******************************************************************************
Public Function CreateValidationChain(ByVal fieldName As String, _
                                     ByVal value As Variant, _
                                     ParamArray ruleTypes() As Variant) As ValidationResult
    Const PROC_NAME As String = "CreateValidationChain"

    Dim validations() As ValidationResult
    ReDim validations(UBound(ruleTypes))

    On Error GoTo ErrorHandler

    Dim i As Integer
    For i = LBound(ruleTypes) To UBound(ruleTypes)
        Set validations(i) = CreateValidationRule(CStr(ruleTypes(i)), fieldName, value)
    Next i

    Set CreateValidationChain = CombineValidationsArray(validations)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    Dim result As ValidationResult
    Set result = New ValidationResult
    result.AddError "Erro interno na cadeia de validação: " & fieldName
    Set CreateValidationChain = result
End Function

'*******************************************************************************
' FORM VALIDATION HELPERS
' Purpose: High-level functions for common form validation scenarios
'*******************************************************************************

'*******************************************************************************
' Function: ValidateFormControls
' Purpose: Validates multiple form controls using a configuration array
' Parameters:
'   controls - Array or Collection of control objects
'   validationConfig - Array of validation configurations
' Returns: ValidationResult - Contains validation status and errors
' Usage: Pass control names and validation rules to validate entire forms
'*******************************************************************************
Public Function ValidateFormControls(ByVal controls As Object, _
                                    ByVal validationConfig As Variant) As ValidationResult
    Const PROC_NAME As String = "ValidateFormControls"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    ' validationConfig should be array of arrays:
    ' Each sub-array: [controlName, fieldType, displayName, required]

    Dim i As Integer
    Dim controlName As String
    Dim fieldType As String
    Dim displayName As String
    Dim required As Boolean
    Dim controlValue As Variant

    For i = LBound(validationConfig) To UBound(validationConfig)
        controlName = validationConfig(i)(0)
        fieldType = validationConfig(i)(1)
        displayName = validationConfig(i)(2)
        required = IIf(UBound(validationConfig(i)) >= 3, validationConfig(i)(3), True)

        ' Get control value
        controlValue = GetControlValue(controls, controlName)

        ' Validate the field
        Dim fieldResult As ValidationResult
        Set fieldResult = ValidateFormField(fieldType, displayName, controlValue, required)

        ' Add errors to main result
        If Not fieldResult.IsValid Then
            result.AddError fieldResult.GetErrorsAsString()
        End If
    Next i

    Set ValidateFormControls = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na validação de controles do formulário"
    Set ValidateFormControls = result
End Function

'*******************************************************************************
' Function: ValidateTextBoxNumeric
' Purpose: Validates numeric textbox controls with automatic correction
' Parameters:
'   textBox - TextBox control to validate
'   fieldName - Display name for errors
'   minValue - Minimum allowed value (optional)
'   maxValue - Maximum allowed value (optional)
'   defaultValue - Value to set if invalid (optional, default 0)
' Returns: ValidationResult - Contains validation status and errors
'*******************************************************************************
Public Function ValidateTextBoxNumeric(ByVal textBox As Object, _
                                      ByVal fieldName As String, _
                                      Optional ByVal minValue As Double = 0, _
                                      Optional ByVal maxValue As Double = 999999999, _
                                      Optional ByVal defaultValue As Double = 0) As ValidationResult
    Const PROC_NAME As String = "ValidateTextBoxNumeric"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    ' Validate the numeric value
    Set result = ValidateNumeric(textBox.value, fieldName, minValue, maxValue)

    ' If validation fails, set default value
    If Not result.IsValid Then
        textBox.value = defaultValue
    End If

    Set ValidateTextBoxNumeric = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na validação de campo numérico: " & fieldName
    textBox.value = defaultValue
    Set ValidateTextBoxNumeric = result
End Function

'*******************************************************************************
' Function: ValidateRequiredFormFields
' Purpose: Validates that all required form fields are filled
' Parameters:
'   formControls - Form controls collection
'   requiredFields - Array of required field names
' Returns: ValidationResult - Contains validation status and errors
'*******************************************************************************
Public Function ValidateRequiredFormFields(ByVal formControls As Object, _
                                          ByVal requiredFields As Variant) As ValidationResult
    Const PROC_NAME As String = "ValidateRequiredFormFields"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim fieldName As String
    Dim controlValue As Variant

    For i = LBound(requiredFields) To UBound(requiredFields)
        fieldName = requiredFields(i)
        controlValue = GetControlValue(formControls, fieldName)

        Dim fieldResult As ValidationResult
        Set fieldResult = ValidateRequired(fieldName, CStr(controlValue))

        If Not fieldResult.IsValid Then
            result.AddError fieldResult.GetErrorsAsString()
        End If
    Next i

    Set ValidateRequiredFormFields = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na validação de campos obrigatórios"
    Set ValidateRequiredFormFields = result
End Function

'*******************************************************************************
' Private Helper Functions for Form Validation
'*******************************************************************************

Private Function GetControlValue(ByVal controls As Object, ByVal controlName As String) As Variant
    On Error GoTo ErrorHandler

    ' Try to get control value from form controls collection
    GetControlValue = controls(controlName).value
    Exit Function

ErrorHandler:
    GetControlValue = vbNullString
End Function

'*******************************************************************************
' Function: ValidateRegionCode
' Purpose: Validates region code according to business rules (moved from RegionHelpers)
' Parameters:
'   fieldName - Name of field for error messages
'   region - Region code to validate
' Returns: ValidationResult - Contains validation status and errors
'*******************************************************************************
Public Function ValidateRegionCode(ByVal fieldName As String, ByVal region As String) As ValidationResult
    Const PROC_NAME As String = "ValidateRegionCode"

    Dim result As ValidationResult
    Set result = New ValidationResult

    On Error GoTo ErrorHandler

    Dim cleanRegion As String
    cleanRegion = Trim(UCase(region))

    ' Basic length validation
    If Len(cleanRegion) < MIN_REGION_LENGTH Or Len(cleanRegion) > MAX_REGION_LENGTH Then
        result.AddError fieldName & " deve ter entre " & MIN_REGION_LENGTH & " e " & MAX_REGION_LENGTH & " caracteres"
        Set ValidateRegionCode = result
        Exit Function
    End If

    ' Check for valid characters (letters only)
    Dim i As Long
    For i = 1 To Len(cleanRegion)
        If Not (Mid(cleanRegion, i, 1) Like "[A-Z]") Then
            result.AddError fieldName & " deve conter apenas letras"
            Set ValidateRegionCode = result
            Exit Function
        End If
    Next i

    ' Additional business rule validation could be added here
    ' For example, check against a list of valid region codes

    Set ValidateRegionCode = result
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError MODULE_NAME, PROC_NAME, False
    result.AddError "Erro interno na validação de região: " & fieldName
    Set ValidateRegionCode = result
End Function