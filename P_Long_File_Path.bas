Attribute VB_Name = "P_Long_File_Path"
Option Explicit

' Windows API Declaration
Private Declare PtrSafe Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, _
     ByVal lpszShortPath As String, _
     ByVal cchBuffer As Long) As Long

Public Function GetShortPath(ByVal longPath As String) As String
    ' Remove duplicate ".csv" extensions if present
    

    ' Check if short path generation is needed
    If Len(longPath) <= 255 Then
        GetShortPath = longPath
        Exit Function
    End If
    
    Dim shortPath As String
    Dim bufferLen As Long
    Dim result As Long
    
    longPath = "\\?\" & longPath
    
    ' Get the required buffer size
    result = GetShortPathName(longPath, vbNullString, 0)
    
    
    If result > 0 Then
        ' Allocate buffer with required size
        shortPath = String(result, Chr(0))
        
        ' Retrieve the short path
        result = GetShortPathName(longPath, shortPath, Len(shortPath) + 1)
        
        Do While Right(shortPath, 4) = ".csv"
            shortPath = left(shortPath, Len(shortPath) - 4)
        Loop
        
        If result > 0 Then
            ' Remove null terminator and return the short path
            GetShortPath = left(shortPath, result)
        Else
            GetShortPath = "Error: Could not retrieve short path"
        End If
    Else
        GetShortPath = "Error: Invalid path or short path unavailable"
    End If
End Function


