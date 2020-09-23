Attribute VB_Name = "miscFunctionData"
Option Explicit

Function PeekByte(ByVal strPacket As String, ByVal Location As Integer) As String
'Purpose:   Returns a 2-char String Data from a Hex String (eg "01020304")
'Input:     strPacket -- HEX Value in string format, eg "0123456ABCDEF"
'           Location -- Reading position of the retrieved byte, start from 0
'Output:    The retrieved Hex byte (eg. "01" or "02"

    PeekByte = Mid$(strPacket, Location * 2 + 1, 2)
End Function

Function Peek(ByVal strPacket As String, ByVal Location As Integer, ByVal NumberofBytes As Integer, Optional ReverseByteOrder As Boolean = True) As String
'Purpose:   Returns specific amount of character of String Data from a Hex
'           String (eg "01020304")
'Input:     strPacket -- HEX Value in string format, eg "0123456ABCDEF"
'           Location -- Reading position of the retrieved data, start from 0
'           NumberofByte -- Amount of byte to retrieved, one byte = 2 character
'Output:    The retrieved Hex value (eg. "01" or "0203", etc)
    
    Dim Output As String, i As Integer
    
    Output = ""
    For i = Location To Location + (NumberofBytes - 1)
        Output = Output + PeekByte(strPacket, i)
    Next i
    
    If ReverseByteOrder = True Then
      Peek = hDump(Output)
    Else
      Peek = Output
    End If
End Function

Function Poke(ByVal strPacket As String, ByVal InsertString As String, ByVal Location As Integer) As String
'Purpose:   Insert a string data into a Hex Packet
'Input:     strPacket -- HEX Value in string format, eg "0123456ABCDEF"
'           InsertString -- The String value to be inserted
'           Location -- Initial Writing position, location start from 0
'Output:    The altered strPacket
    
    If Len(InsertString) = 0 Then
        Poke = strPacket
        Exit Function
    End If
    
    Mid$(strPacket, Location * 2 + 1) = UCase$(InsertString)
    Poke = strPacket
End Function

Function CutTextL(ByVal InputString As String, ByVal LengthtoCut) As String
'Purpose:   Cut a text string from Left a specific number of character
'Input:     InputString -- String data
'           LengthtoCut -- The number of character to cut from a text
'Output:    String Data

    If LengthtoCut >= Len(InputString) Then
        CutTextL = ""
        Exit Function
    End If
    CutTextL = Right$(InputString, Len(InputString) - LengthtoCut)
End Function

Function CutTextR(ByVal InputString As String, LengthtoCut) As String
'Purpose:   Cut a text string from Right a specific number of character
'Input:     InputString -- String data
'           LengthtoCut -- The number of character to cut from a text
'Output:    String Data
    
    If LengthtoCut >= Len(InputString) Then
        CutTextR = ""
        Exit Function
    End If
    CutTextR = Left$(InputString, Len(InputString) - LengthtoCut)
End Function

Function StrAppend(ParamArray strText())
'Purpose:   Retrieve a series of text and prepare them to a format ready for packet send
'Input:     Array of string
'Output:    String in HexString ready to be placed in a packet

  Dim i As Integer
  StrAppend = ""
  For i = LBound(strText) To UBound(strText)
    StrAppend = StrAppend + _
      Dec_to_Hex(Len(strText(i)) + 1, tInt) + _
      Str_to_Hex(strText(i) + vbNullChar)
  Next i
End Function

Function StrSplitbyChar(ByVal strText As String, ByVal strSplitChar As String, Optional ArrayDimension As Integer = -1)
'Purpose:   Split a single string, to an array by specifying the split character
'Input:     String and a Spliting Character
'Output:    An array of string value
  
  Dim i As Integer, j As Integer, k As Integer, _
    TotalChr As Integer
  
  If ArrayDimension = -1 Then
    i = 1
    Do While InStr(i, strText, strSplitChar) > 0
      i = InStr(i, strText, strSplitChar)
      i = i + 1
      TotalChr = TotalChr + 1
    Loop
  Else
    TotalChr = ArrayDimension
  End If
  
  ReDim Output(TotalChr) As String

  For i = 1 To TotalChr
    j = InStr(1, strText, strSplitChar)
    Output(i - 1) = Left$(strText, j - 1)
    strText = CutTextL(strText, j)
  Next i
  Output(TotalChr) = strText
  
  StrSplitbyChar = Output
End Function


