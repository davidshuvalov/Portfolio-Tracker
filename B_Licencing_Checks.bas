Attribute VB_Name = "B_Licencing_Checks"
'    ts_customer_id = 4209838
#If VBA7 Then
    Private Declare PtrSafe Function MultiWalkIsLicensePro Lib "MultiWalkLicense64.dll" Alias "_MultiWalkIsLicensePro" ( _
        ByVal program_folder As String, _
        ByVal ts_customer_id As Long, _
        ByVal app_name As String) As Integer
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#Else
    Private Declare Function MultiWalkIsLicensePro Lib "MultiWalkLicense32.dll" Alias "_MultiWalkIsLicensePro" ( _
        ByVal program_folder As String, _
        ByVal ts_customer_id As Long, _
        ByVal app_name As String) As Integer
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#End If


Private Declare PtrSafe Function RegGetValue Lib "advapi32.dll" Alias "RegGetValueA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal lpValue As String, _
    ByVal dwFlags As Long, _
    ByVal lpType As Long, _
    ByVal pvData As Any, _
    ByRef pcbData As Long) As Long

Private Const HKEY_CURRENT_USER As Long = &H80000001

Private Const RRF_RT_REG_SZ As Long = &H2

Public MultiWalkProgramFolder As String     ' MultiWalk program folder.  Declared publically so we only need to load from registry once.

' Load the Licensing DLL from the MultiWalk program folder

Public Function LibraryLoaded() As Boolean
    Static isLoaded As Boolean
    Static triedToLoad As Boolean
    Dim dllFileSpec As String

   ' If triedToLoad Then
   '     LibraryLoaded = isLoaded
   '     Exit Function
   ' End If
    
    MultiWalkProgramFolder = GetMultiWalkProgramFolder()
    
    If MultiWalkProgramFolder = "" Then
        MsgBox ("Could not get MultiWalk Program path.  Is MultiWalk installed?")
        isLoaded = False
    Else
         #If VBA7 Then
         dllFileSpec = MultiWalkProgramFolder & "\MultiWalkLicense64.dll"
        #Else
         dllFileSpec = MultiWalkProgramFolder & "\MultiWalkLicense32.dll"
        #End If
        
        isLoaded = LoadLibrary(dllFileSpec)
        
        If Not isLoaded Then
            MsgBox ("Could not load DLL library " & dllFileSpec & ". Is most current version of MulitWalk installed? If you are using a 32-bit version of Excel, please reach out to David Shuvalov for a 32-bit licence key.")
        End If
    End If
    
    triedToLoad = True

    LibraryLoaded = isLoaded
End Function
    

Function GetMultiWalkProgramFolder() As String
    Dim buffer As String * 512
    Dim bufferSize As Long
    Dim result As Long

    bufferSize = Len(buffer)

    result = RegGetValue(HKEY_CURRENT_USER, "SOFTWARE\MultiWalk", "MultiWalkProgramFolder", RRF_RT_REG_SZ, 0, ByVal buffer, bufferSize)

    If result = 0 Then
        GetMultiWalkProgramFolder = left(buffer, InStr(buffer, vbNullChar) - 1)
    Else
        GetMultiWalkProgramFolder = ""
    End If

End Function

 
Function IsLicenseValid() As Boolean
    Dim appName As String
    Dim returnCode As Integer
    Dim customerIDs As Variant
    Dim i As Long
    Dim foundID As Boolean
    
    ' Initialize the output
    IsLicenseValid = False
    
    ' Load the DLL library using the installed MultiWalk Program folder
    If Not LibraryLoaded() Then Exit Function
    
    ' Load customer IDs
    customerIDs = GetCustomerIDs()
    
    ' Check if the customer ID in the range matches any in the list
    foundID = False
    For i = LBound(customerIDs) To UBound(customerIDs)
        If customerIDs(i) = GetNamedRangeValue("TS_Customer_Number") Then
            foundID = True
            Exit For
        End If
    Next i
    
    ' If no matching customer ID was found, exit the function
    If Not foundID Then Exit Function
    
    ' Define your application name
    appName = "ShuvalovPortfolio"
    
    ' Validate the license using the DLL function
    On Error Resume Next
    returnCode = MultiWalkIsLicensePro(MultiWalkProgramFolder, customerIDs(i), appName)
    
    ' Handle the return codes
    Select Case returnCode
        Case 0
            ' License successfully verified
            IsLicenseValid = True
        Case 1
            ' Invalid program folder
            MsgBox "Invalid program folder for Customer ID: " & customerIDs(i), vbCritical
        Case 3
            ' No license key file
            MsgBox "No license key file found for Customer ID: " & customerIDs(i), vbCritical
        Case 4
            ' Multiple license keys found
            MsgBox "Multiple license keys found for Customer ID: " & customerIDs(i), vbCritical
        Case Else
            ' Unhandled or unexpected return code
            MsgBox "Unexpected error for Customer ID: " & customerIDs(i) & " (Code: " & returnCode & ")", vbCritical
    End Select
    
    ' Handle any runtime errors
    If Err.Number <> 0 Then
        MsgBox "Error checking license: " & Err.Description, vbCritical
    End If
    On Error GoTo 0
End Function


 



 

Sub TestLicense()
    Dim customerID As Long
    Dim isValid As Boolean
 

    

    ' Validate the license

    isValid = IsLicenseValid()

    If isValid Then
        MsgBox "License is valid."
    Else
        MsgBox "Invalid or missing license."
    End If
End Sub


