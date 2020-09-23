Attribute VB_Name = "miscFunction"
Option Explicit
Public Enum VarType
    tByte = 1
    tInt = 2
    tLong = 4
End Enum

Public Enum BitShiftDirection
    RightShift = 1
    LeftShift = -1
End Enum

Public Const HexBinary = _
    "0000" + "0001" + "0010" + "0011" + _
    "0100" + "0101" + "0110" + "0111" + _
    "1000" + "1001" + "1010" + "1011" + _
    "1100" + "1101" + "1110" + "1111"
    
Function hMul(ByVal strVariable1 As String, ByVal strVariable2 As String) As String
'Purpose:   To multiply two hex value, without any size limit so that for example
'           You could multiply 2 256-byte Hex Value and still get an answer without
'           overflow error.
'Input:     2 Hex Number in string format eg. "0123456789ABCDEF"
'Output:    Hex Number in string format

    Dim Multiplier1 As String, _
        Multiplier2 As String, _
        Remainder As String, _
        Resultant As String, _
        SubOutput As String, _
        Output As String, _
        i, j

    For i = Len(strVariable2) To 1 Step -1
        Multiplier1 = "&H" + Mid$(strVariable2, i, 1)
        Remainder = "0"
        SubOutput = ""
        For j = Len(strVariable1) To 1 Step -1
            Multiplier2 = "&H" + Mid$(strVariable1, j, 1)
            Resultant = Hex$((Multiplier1 * Multiplier2) + Val("&H" + Remainder))
            Remainder = CutTextR(Resultant, 1)
            SubOutput = Right$(Resultant, 1) + SubOutput
        Next j
        SubOutput = Remainder + SubOutput + String$(Len(strVariable2) - i, "0")
        Output = hAdd(SubOutput, Output)
    Next i
    
    hMul = Output
End Function

Function hAdd(ByVal strVariable1 As String, ByVal strVariable2 As String) As String
'Purpose:   To add two hex value, without any size limit so that for example
'           You could add 2 256-byte Hex Value and still get an answer without
'           overflow error.
'Input:     2 Hex Number in string format eg. "0123456789ABCDEF"
'Output:    Hex Number in string format
    
    Dim Var1_Length As Long, _
        Var2_Length As Long, _
        Var_Length As Long, _
        Factor1 As String, _
        Factor2 As String, _
        Resultant As String, _
        Remainder As String, _
        SubOutput As String, _
        Output As String, _
        i As Integer

    Var1_Length = Len(strVariable1)
    Var1_Length = Int(Var1_Length / 2) + (Var1_Length Mod 2)
    Var2_Length = Len(strVariable2)
    Var2_Length = Int(Var2_Length / 2) + (Var2_Length Mod 2)
    
    If Var2_Length > Var1_Length Then
        Var_Length = Var2_Length - 1
    Else
        Var_Length = Var1_Length - 1
    End If
    
    strVariable1 = hFill(strVariable1, Var_Length + 1)
    strVariable2 = hFill(strVariable2, Var_Length + 1)
    
    For i = Var_Length To 0 Step -1
        Factor1 = "&H" + PeekByte(strVariable1, i)
        Factor2 = "&H" + PeekByte(strVariable2, i)
        Resultant = HexB$(Val(Factor1 - -Factor2) + Val("&H" + Remainder))
        SubOutput = Right$(Resultant, 2)
        Remainder = CutTextR(Resultant, 2)
        Output = SubOutput + Output
    Next i
    Output = Remainder + Output
    hAdd = Output
End Function

Function BitShift(ByVal strVariable As String, ByVal NoofBits As Integer, ByVal Direction As BitShiftDirection) As String
'Purpose:   Shift the variable left or right in a number of bits.
'Input:     strVariable -- A Hex value in string format eg. "0123456789ABCDEF"
'           NoOfBits -- The amount of binary bits to shift
'           Direction -- LeftShift or RightShift
'Output:    Hex Number in string format

    Dim BinaryValue  As String, _
        Var_Length As Integer, _
        Binary_Length As Integer, _
        Output As String

    If Val("&H" + strVariable) = 0 Then
        BitShift = "0"
        Exit Function
    End If
    
    Var_Length = Len(strVariable) / 2
    
    BinaryValue = Hex_to_Bin(strVariable)
    If Direction = LeftShift Then
        BinaryValue = BinaryValue + String$(NoofBits, "0")
    Else
        BinaryValue = String$(NoofBits, "0") + CutTextR(BinaryValue, NoofBits)
    End If

    Binary_Length = Len(BinaryValue) Mod 4
    If Binary_Length = 0 Then Binary_Length = 4
    BinaryValue = String$(4 - Binary_Length, "0") + BinaryValue
    Output = hFill(Bin_to_Hex(BinaryValue), Var_Length)
    
    BitShift = Output
End Function

Function hAnd(ByVal hexVariable1 As String, ByVal hexVariable2 As String) As String
'Purpose:   Just like the normal AND operator, except this function has no size limit,
'           therefore you would not get any more overflow error
'Input:     2 Hex Number in string format eg. "0123456789ABCDEF"
'Output:    Hex Number in string format

    Dim Var1_Length As Long, _
        Var2_Length As Long, _
        Var_Length As Long, _
        Variable1 As String, _
        Variable2 As String, _
        SubOutput As String, _
        Output As String, _
        i As Integer

    Var1_Length = Len(hexVariable1)
    Var1_Length = Int(Var1_Length / 2) + (Var1_Length Mod 2)
    Var2_Length = Len(hexVariable2)
    Var2_Length = Int(Var2_Length / 2) + (Var2_Length Mod 2)
    
    If Var2_Length > Var1_Length Then
        Var_Length = Var2_Length - 1
    Else
        Var_Length = Var1_Length - 1
    End If
    
    hexVariable1 = hFill(hexVariable1, Var_Length + 1)
    hexVariable2 = hFill(hexVariable2, Var_Length + 1)
    
    For i = 0 To Var_Length
        Variable1 = "&H" + PeekByte(hexVariable1, i)
        Variable2 = "&H" + PeekByte(hexVariable2, i)
        SubOutput = HexB$(Val(Variable1 And Variable2))
        Output = Output + SubOutput
    Next i
    
    hAnd = Output
End Function

Function hXor(ByVal hexVariable1 As String, ByVal hexVariable2 As String) As String
'Purpose:   Just like the normal XOR operator, except this function has no size limit,
'           therefore you would not get any more overflow error
'Input:     2 Hex Number in string format eg. "0123456789ABCDEF"
'Output:    Hex Number in string format

    Dim Var1_Length As Long, _
        Var2_Length As Long, _
        Var_Length As Long, _
        Variable1 As String, _
        Variable2 As String, _
        SubOutput As String, _
        Output As String, _
        i As Integer

    Var1_Length = Len(hexVariable1)
    Var1_Length = Int(Var1_Length / 2) + (Var1_Length Mod 2)
    Var2_Length = Len(hexVariable2)
    Var2_Length = Int(Var2_Length / 2) + (Var2_Length Mod 2)
    
    If Var2_Length > Var1_Length Then
        Var_Length = Var2_Length - 1
    Else
        Var_Length = Var1_Length - 1
    End If
    
    hexVariable1 = hFill(hexVariable1, Var_Length + 1)
    hexVariable2 = hFill(hexVariable2, Var_Length + 1)
    
    For i = 0 To Var_Length
        Variable1 = "&H" + PeekByte(hexVariable1, i)
        Variable2 = "&H" + PeekByte(hexVariable2, i)
        SubOutput = HexB$(Val(Variable1 Xor Variable2))
        Output = Output + SubOutput
    Next i
    
    hXor = Output
End Function


Function hOr(ByVal hexVariable1 As String, ByVal hexVariable2 As String) As String
'Purpose:   Just like the normal OR operator, except this function has no size limit,
'           therefore you would not get any more overflow error
'Input:     2 Hex Number in string format eg. "0123456789ABCDEF"
'Output:    Hex Number in string format
    
    Dim Var1_Length As Long, _
        Var2_Length As Long, _
        Var_Length As Long, _
        Variable1 As String, _
        Variable2 As String, _
        SubOutput As String, _
        Output As String, _
        i As Integer

    Var1_Length = Len(hexVariable1)
    Var1_Length = Int(Var1_Length / 2) + (Var1_Length Mod 2)
    Var2_Length = Len(hexVariable2)
    Var2_Length = Int(Var2_Length / 2) + (Var2_Length Mod 2)
    
    If Var2_Length > Var1_Length Then
        Var_Length = Var2_Length - 1
    Else
        Var_Length = Var1_Length - 1
    End If
    
    hexVariable1 = hFill(hexVariable1, Var_Length + 1)
    hexVariable2 = hFill(hexVariable2, Var_Length + 1)
    
    For i = 0 To Var_Length
        Variable1 = "&H" + PeekByte(hexVariable1, i)
        Variable2 = "&H" + PeekByte(hexVariable2, i)
        SubOutput = HexB$(Val(Variable1 Or Variable2))
        Output = Output + SubOutput
    Next i
    
    hOr = Output
End Function

Function hNot(ByVal hexVariable As String) As String
'Purpose:   Just like the normal NOT operator, except this function has no size limit,
'           therefore you would not get any more overflow error
'Input:     Hex Number in string format eg. "0123456789ABCDEF"
'Output:    Hex Number in string format
    
    Dim BinaryValue As String, _
        Output As String, _
        i As Integer
        
    BinaryValue = Hex_to_Bin(hexVariable)

    Output = ""
    For i = 1 To Len(BinaryValue)
        If Mid$(BinaryValue, i, 1) = "0" Then
            Output = Output + "1"
        Else
            Output = Output + "0"
        End If
    Next i
    
    hNot = Bin_to_Hex(Output)
End Function


Function Hex_to_Bin(ByVal Variable As String) As String
'Purpose:   Convert a Hex Value (String format) to Binary Value (String format)
'Input:     HEX Value in string format, eg "0123456ABCDEF"
'Output:    Binary Value in string format, eg "0000000100100011"
    
    Dim TempValue As Integer, _
        SubOutput As String, _
        Output As String, _
        i As Integer
        
    For i = 1 To Len(Variable)
        TempValue = Val("&H" + Mid$(Variable, i, 1))
        SubOutput = hDump(Peek(HexBinary, TempValue * 2, 2))
        Output = Output + SubOutput
    Next i
    
    Hex_to_Bin = Output
End Function

Function Bin_to_Hex(ByVal Variable As String) As String
'Purpose:   Convert a Binary Value (String format) to Hex Value (String format)
'Input:     Binary Value in string format, eg "0000000100100011"
'Output:    HEX Value in string format, eg "0123456ABCDEF"
    
    Dim Binary_Length As Integer, _
        TempBinary As String, _
        SubOutput As Byte, _
        Output As String, _
        i As Integer

    Binary_Length = Len(Variable) Mod 4
    If Binary_Length = 0 Then Binary_Length = 4
    Variable = String$(4 - Binary_Length, "0") + Variable
    
    For i = 1 To Len(Variable) Step 4
        TempBinary = Mid$(Variable, i, 4)
        
        SubOutput = 0
        If InStr(1, TempBinary, "1") <> 0 Then
            If Mid$(TempBinary, 1, 1) = "1" Then SubOutput = SubOutput + 8
            If Mid$(TempBinary, 2, 1) = "1" Then SubOutput = SubOutput + 4
            If Mid$(TempBinary, 3, 1) = "1" Then SubOutput = SubOutput + 2
            If Mid$(TempBinary, 4, 1) = "1" Then SubOutput = SubOutput + 1
        End If

        Output = Output + Hex$(SubOutput)
    Next i

    Bin_to_Hex = Output
End Function

Function HexB(ByVal Variable) As String
'Purpose:   Like the normat Hex$() Function, except that this return a full byte,
'           eg "0A" instead of just "A". So basically this code just add "0 if
'           required.
'Input:     Decimal Value
'Output:    HEX Value in string format, eg "0123456ABCDEF"
    
    Dim Output As String

    Output = Hex$(Variable)
    If Len(Output) Mod 2 = 1 Then Output = "0" + Output
    HexB = Output
End Function

Function Dec_to_Hex(ByVal Variable, ByVal OutType As VarType) As String
'Purpose:   Convert a Decimal Value to Hex Value (String format)
'Input:     Decimal Value (Whole number, no fractions)
'Output:    HEX Value in string format, eg "0123456ABCDEF"
    
    Dim SubOutput, _
        Output As String
        
    Select Case OutType
        Case tByte: SubOutput = CByte(Variable)
        Case tInt:  SubOutput = CInt(Variable)
        Case tLong: SubOutput = CLng(Variable)
    End Select
    
    Output = Hex$(SubOutput)
    Output = String$((OutType * 2) - Len(Output), "0") + Output
    If OutType = tInt Or OutType = tLong Then Output = hDump(Output)
    
    Dec_to_Hex = Output
End Function

Function Hex_to_Str(ByVal Variable As String) As String
'Purpose:   Convert a Hex Value (String format) to String
'Input:     HEX Value in string format, eg "0123456ABCDEF"
'Output:    String data
    
    Dim Var_Length As Integer, _
        SubOutput As String, _
        Output As String, _
        i As Integer
        
    Var_Length = Len(Variable)
    Var_Length = Int(Var_Length / 2) + (Var_Length Mod 2)
    Variable = hFill(Variable, Var_Length)
    
    For i = 1 To Len(Variable) Step 2
        SubOutput = Chr$("&H" + PeekByte(Variable, (i - 1) / 2))
        Output = Output + SubOutput
    Next i
    Hex_to_Str = Output
End Function

Function Str_to_Hex(ByVal strVariable As String) As String
'Purpose:   Convert a String to Hex Value (String format)
'Input:     String data
'Output:    HEX Value in string format, eg "0123456ABCDEF"
    
    Dim SubOutput As String, _
        Output As String, _
        i As Integer
    
    For i = 1 To Len(strVariable)
        SubOutput = HexB(Asc(Mid$(strVariable, i, 1)))
        Output = Output + SubOutput
    Next i
    Str_to_Hex = Output
End Function

Function IP_to_Hex(ByVal strIP As String) As String
'Purpose:   Convert a IP Addreess (String format) to Hex Value (String format)
'Input:     IP Address (not hostname) in String format
'Output:    HEX Value in string format, eg "0123456ABCDEF"
    
    Dim ReadPos As Integer, _
        NextDotPos As Integer, _
        IPPos As Integer, _
        SubOutput As String, _
        Output As String
    
    ReadPos = 1
    IPPos = 0
    
    Output = "00000000"
    Do
        NextDotPos = InStr(ReadPos, strIP, ".", vbBinaryCompare)
        If NextDotPos = 0 Then
            NextDotPos = Len(strIP)
            SubOutput = Val(Mid$(strIP, ReadPos, NextDotPos - (ReadPos - 1)))
            Output = Poke(Output, HexB$(SubOutput), IPPos)
            Exit Do
        End If
        SubOutput = Val(Mid$(strIP, ReadPos, NextDotPos - (ReadPos - 1)))
        ReadPos = NextDotPos + 1
        Output = Poke(Output, HexB$(SubOutput), IPPos)
        IPPos = IPPos + 1
    Loop Until NextDotPos = 0
    
    IP_to_Hex = Output
End Function

Function Hex_to_IP(ByVal strVariable As String) As String
'Purpose:   Convert a Hex Value (String format) to IP Addreess (String format)
'Input:     HEX Value in string format, eg "0123456ABCDEF"
'Output:    IP Address (not hostname) in String format
    
    Dim SubOutput As String, Output As String, i As Integer
    
    strVariable = hDump(hFill(strVariable, tLong))
    For i = 1 To 4
        SubOutput = PeekByte$(strVariable, i - 1)
        SubOutput = Trim(Str$(Val("&H" + SubOutput)))
        Output = Output + SubOutput + "."
    Next i
    Hex_to_IP = CutTextR(Output, 1)
End Function



Function hDump(ByVal hexVariable As String) As String
'Purpose:   Reverse the order of a Hex Value (String Format), thus creating
'           a hex dump. For example, input = "01ABCDEF", output will equal
'           to "EFCDAB01"
'Input:     HEX Value in string format, eg "0123456ABCDEF"
'Output:    The reversed HEX Value in string format, eg "0123456ABCDEF"
    
    Dim Output As String, i As Integer
    If Len(hexVariable) Mod 2 = 1 Then hexVariable = "0" + hexVariable
    
    Output = ""
    For i = 1 To Len(hexVariable) Step 2
        Output = Mid$(hexVariable, i, 2) + Output
    Next i
    
    hDump = Output
End Function

Function hFill(ByVal hexVariable As String, ByVal OutByte As Integer) As String
'Purpose:   Flush or fill a Hex Value (String Format) to a specific number of byte
'           For example,
'           1) Flush
'               *input = "01ABCDEF" & Outbyte = 2, Output = "CDEF"
'           2) Fill
'               *input = "AB" & Outbyte = 4, Output = "000000AB"
'Input:     hexVariable -- HEX Value in string format, eg "0123456ABCDEF"
'           OutByte -- A Specific number of byte for the output to be
'Output:    HEX Value in string format, eg "0123456ABCDEF"
    
    Dim Output_Length As Integer, Current_Length As Integer
    
    Output_Length = OutByte * 2
    Current_Length = Len(hexVariable)
    
    If Current_Length = Output_Length Then
        hFill = hexVariable
        Exit Function
    End If
    
    If Current_Length < Output_Length Then
        hFill = String$(Output_Length - Current_Length, "0") + hexVariable
    Else
        hFill = Right$(hexVariable, Output_Length)
    End If
End Function

