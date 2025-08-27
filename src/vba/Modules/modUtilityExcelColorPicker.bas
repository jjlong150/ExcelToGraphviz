Attribute VB_Name = "modUtilityExcelColorPicker"
Option Explicit

Public Enum ColorPickerType
    cpExcelDialog = 0   ' Uses Application.Dialogs(xlDialogEditColor)
    cpWindowsAPI = 1    ' Uses Windows ChooseColor API
End Enum

#If Win32 Or Win64 Then
    ' Windows API declarations for color chooser
    
    ' === ChooseColor Flags ===
    Public Const CC_RGBINIT               As Long = &H1      ' Use rgbResult to initialize dialog
    Public Const CC_FULLOPEN              As Long = &H2      ' Show custom colors section by default
    Public Const CC_PREVENTFULLOPEN       As Long = &H4      ' Disable expanding to custom colors
    Public Const CC_SHOWHELP              As Long = &H8      ' Show Help button (requires hook)
    Public Const CC_ENABLEHOOK            As Long = &H10     ' Enable hook procedure via lpfnHook
    Public Const CC_ENABLETEMPLATE        As Long = &H20     ' Use custom dialog template via lpTemplateName
    Public Const CC_ENABLETEMPLATEHANDLE  As Long = &H40     ' Use template from memory handle in hInstance
    Public Const CC_SOLIDCOLOR            As Long = &H80     ' Restrict to solid colors only
    Public Const CC_ANYCOLOR              As Long = &H100    ' Show all available colors

    Private Type ChooseColor
        lStructSize As Long
        hwndOwner As LongPtr
        hInstance As LongPtr
        rgbResult As Long
        lpCustColors As LongPtr
        Flags As Long
        lCustData As LongPtr
        lpfnHook As LongPtr
        lpTemplateName As String
    End Type
    
    #If VBA7 Then
        Private Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
        Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
        Private Declare PtrSafe Function VarPtr Lib "vbe7.dll" (Ptr As Any) As LongPtr
    #Else
        Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
        Private Declare Function GetActiveWindow Lib "user32" () As Long
        Private Declare Function VarPtr Lib "vbe6.dll" Alias "VarPtr" (Ptr As Any) As Long
    #End If
#End If

' Function to show color chooser dialog with pre-selected hex color and return RGB as Long
Public Function ShowColorChooser(ByVal colorAsHex As String, Optional ByVal pickerType As ColorPickerType = cpExcelDialog) As Long
    ShowColorChooser = -1   ' Default to fail/cancel until a color is chosen
    
    On Error GoTo ErrorHandler
    
    Dim hexColor As String
    hexColor = colorAsHex   ' Copy input to local variable so we can manipulate the string
    
    Dim r As Long, g As Long, b As Long
    
    ' Validate and clean hex color
    hexColor = UCase(replace(hexColor, "#", ""))
    If Len(hexColor) <> 6 Then
        Debug.Print "ShowColorChooser(): Invalid hex color format. Expected: #RRGGBB, Received: " & hexColor
        ShowColorChooser = -1
        Exit Function
    End If
    
    ' Convert hex to RGB
    On Error Resume Next
    r = CLng("&H" & Left(hexColor, 2))
    g = CLng("&H" & Mid(hexColor, 3, 2))
    b = CLng("&H" & Right(hexColor, 2))
    If Err.number <> 0 Then
        Debug.Print "ShowColorChooser(): Error converting hex to RGB: " & Err.Description
        ShowColorChooser = -1
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Default to no color selected
    ShowColorChooser = -1
        
#If Mac Then
    ' Mac-specific color chooser
    Dim initialRGB As String
    initialRGB = CStr(r) & "," & CStr(g) & "," & CStr(b)
    
    Dim rgbResult As String
    rgbResult = RunAppleScriptTask("pickColor", initialRGB) ' Traps AppleScriptTask errors

    If rgbResult <> "" Then
        ' Parse RGB result
        rgbResult = replace(replace(rgbResult, "(", ""), ")", "")
        Dim rgbArray As Variant
        rgbArray = split(rgbResult, ",")
        If UBound(rgbArray) >= 2 Then
            r = CLng(Trim(rgbArray(0)))
            g = CLng(Trim(rgbArray(1)))
            b = CLng(Trim(rgbArray(2)))
            ShowColorChooser = RGB(r, g, b)
        Else
            Debug.Print "ShowColorChooser(): Invalid RGB result format from color picker."
        End If
    End If
#Else
    Select Case pickerType
        Case cpExcelDialog
            'Open the Application.Dialogs edit color dialog box
            If Application.Dialogs(xlDialogEditColor).show(1, r, g, b) = True Then
                ShowColorChooser = ActiveWorkbook.colors(1)
            End If
    
        Case cpWindowsAPI
            ' Set initial color using RGB order (COLORREF is RGB: R + G*256 + B*65536)
            Dim initialColor As Long
            initialColor = RGB(r, g, b)
            
            ' Set up CHOOSECOLOR structure
            Dim cc As ChooseColor
            Dim customColors(0 To 15) As Long
            With cc
                .lStructSize = LenB(cc)
                .hwndOwner = GetActiveWindow()
                .rgbResult = initialColor
                .lpCustColors = VarPtr(customColors(0))
                .Flags = CC_RGBINIT Or CC_FULLOPEN ' Pre-select color using RGB, and show custom colors section
            End With
            
            'Open the ChooseColor dialog box
            If ChooseColor(cc) <> 0 Then
                ShowColorChooser = cc.rgbResult
            End If
    End Select
#End If
    
    Exit Function

ErrorHandler:
    Debug.Print "ShowColorChooser() Error - " & Err.Description & vbCrLf & _
           "Error Number: " & Err.number
    ShowColorChooser = -1
End Function

' Function to convert RGB Long to hex color string
Public Function RGBToHex(rgbColor As Long) As String
    On Error GoTo ErrorHandler
    
    Dim r As Long, g As Long, b As Long
    
    ' Extract RGB components
    r = rgbColor And 255
    g = (rgbColor \ 256) And 255
    b = (rgbColor \ 65536) And 255
    
    ' Convert to hex and format
    RGBToHex = "#" & Right("0" & Hex(r), 2) & _
                    Right("0" & Hex(g), 2) & _
                    Right("0" & Hex(b), 2)
    
    Exit Function

ErrorHandler:
    Debug.Print "RGBToHex(): Error - " & Err.Description
    RGBToHex = "#000000"
End Function

' Test harness to debug color chooser
Sub TestColorFunctions()
    Dim hexResult As String
    Dim inputHex As String
    
    inputHex = "#00FFFF" ' Test with cyan
    Debug.Print "Input Hex: " & inputHex
    
    Dim selectedColor As Long
    'selectedColor = ShowColorChooser(inputHex, cpExcelDialog)
    selectedColor = ShowColorChooser(inputHex, cpWindowsAPI)
    'selectedColor = ShowColorChooser(inputHex)
    
    If selectedColor <> -1 Then
        hexResult = RGBToHex(selectedColor)
        Debug.Print "Selected RGB Long: " & selectedColor & vbCrLf & _
               "Selected Hex: " & hexResult & vbCrLf & _
               "RGB Components: R=" & (selectedColor And 255) & _
               ", G=" & ((selectedColor \ 256) And 255) & _
               ", B=" & ((selectedColor \ 65536) And 255)
    Else
        Debug.Print "No color selected (or function failed)"
    End If
End Sub




