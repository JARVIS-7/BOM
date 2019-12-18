Attribute VB_Name = "modCrc32"
Option Explicit
Option Compare Text

Private mlarrCrc32Table(255) As Long

Private Function InitCrc32(Optional ByVal lSeed As Long = &HEDB88320, Optional ByVal lPreCondition As Long = &HFFFFFFFF) As Long
    
    '// Turn on error trapping
    On Error Resume Next
    
    '// Declare counter variable iBytes, counter variable iBits, value variables lCrc32 and lTempCrc32
    Dim iBytes As Integer, iBits As Integer, lCrc32 As Long, lTempCrc32 As Long
    
    
    '// Iterate 256 times
    For iBytes = 0 To 255
        
        '// Initiate lCrc32 to counter variable
        lCrc32 = iBytes
        
        '// Now iterate through each bit in counter byte
        For iBits = 0 To 7
            '// Right shift unsigned long 1 bit
            lTempCrc32 = lCrc32 And &HFFFFFFFE
            lTempCrc32 = lTempCrc32 \ &H2
            lTempCrc32 = lTempCrc32 And &H7FFFFFFF
            
            '// Now check if temporary is less than zero and then mix Crc32 checksum with Seed value
            If (lCrc32 And &H1) <> 0 Then
                lCrc32 = lTempCrc32 Xor lSeed
            Else
                lCrc32 = lTempCrc32
            End If
        Next
        
        '// Put Crc32 checksum value in the holding array
        mlarrCrc32Table(iBytes) = lCrc32
    Next
    
    '// After this is done, set function value to the precondition value
    InitCrc32 = lPreCondition
    
End Function

'// The function above is the initializing function, now we have to write the computation function
Private Function AddCrc32(ByVal sItem As String, ByVal lCrc32 As Long) As Long
    
    '// Turn on error trapping
    On Error Resume Next
    
    '// Declare following variables
    Dim bCharValue As Byte, iCounter As Integer, lIndex As Long
    Dim lAccValue As Long, lTableValue As Long
    
    '// Iterate through the string that is to be checksum-computed
    For iCounter = 1 To Len(sItem)
        
        '// Get ASCII value for the current character
        bCharValue = Asc(Mid$(sItem, iCounter, 1))
        
        '// Right shift an Unsigned Long 8 bits
        lAccValue = lCrc32 And &HFFFFFF00
        lAccValue = lAccValue \ &H100
        lAccValue = lAccValue And &HFFFFFF
        
        '// Now select the right adding value from the holding table
        lIndex = lCrc32 And &HFF
        lIndex = lIndex Xor bCharValue
        lTableValue = mlarrCrc32Table(lIndex)
        
        '// Then mix new Crc32 value with previous accumulated Crc32 value
        lCrc32 = lAccValue Xor lTableValue
    Next
    
    '// Set function value the the new Crc32 checksum
    AddCrc32 = lCrc32
    
End Function

'// At last, we have to write a function so that we can get the Crc32 checksum value at any time
Private Function GetCrc32(ByVal lCrc32 As Long) As Long
   
    '// Turn on error trapping
    On Error Resume Next
    
    '// Set function to the current Crc32 value
    GetCrc32 = lCrc32 Xor &HFFFFFFFF
    
End Function

Public Function Crc32(ByVal sTxt As String) As String
    
    Dim lCrc32Value As Long
    
    On Error Resume Next
    lCrc32Value = InitCrc32()
    lCrc32Value = AddCrc32(sTxt, lCrc32Value)
    Crc32 = Hex$(GetCrc32(lCrc32Value))
    
End Function

