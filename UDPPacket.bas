Attribute VB_Name = "UDPPacket"
Option Explicit
Public Const ICQTableV5 = _
    "5960376B6562464853614C5960575B3D5E346D36503F6F6753614C5940476339" + _
    "505F5F3F6F47436948333164355A4A425640675341076C49583B4D4668436948" + _
    "333144656246485341076C69483351545D4E6C49384B554A6246483351346D36" + _
    "505F5F5F3F6F4763594067333164355A6A526E3C51346D36505F5F3F4F374B35" + _
    "5A4A6266583B4D66585B5D4E6C49583B4D66583B4D464853614C594067333164" + _
    "556A323E4445526E3C3164556A524E6C694853614C39306F47635960575B3D3E" + _
    "64353A3A5A6A524E6C694853616C49583B4D46686339505F5F3F6F6753412541" + _
    "3C51543D5E545D4E4C39505F5F5F3F6F474369483351545D6E3C3164355A0000"
    
'Public Const ICQTableV4 = _
    "0A5B315D20596F752063616E206D6F646966792074686520736F756E647320494351206D616B65732E204A7573742073" + _
    "656C6563742022536F756E6473222066726F6D207468652022707265666572656E6365732F6D6973632220696E20494351206F722066726F6D20746865202253" + _
    "6F756E64732220696E2074686520636F6E74726F6C2070616E656C2E204372656469743A204572616E0A5B325D2043616E27742072656D656D62657220776861" + _
    "742077617320736169643F2020446F75626C652D636C69636B206F6E2061207573657220746F206765742061206469616C6F67206F6620616C6C206D65737361" + _
    "6765732073656E7420696E636F6D696E"

Function UDP_CreatePacket(ByRef Header As UDP_CLIENT_HEADER) As String
   Dim Packet As String
        
    Header.SeqNum1 = Header.SeqNum1 + 1
    Header.SeqNum2 = Header.SeqNum2 + 1
    
    'Assemble the packet
    Packet = _
        ICQ_UDP_VERSION + _
        "00000000" + _
        Dec_to_Hex(Header.uin, tLong) + _
        Dec_to_Hex(Header.SessionID, tLong) + _
        Dec_to_Hex(Header.Command, tInt) + _
        Dec_to_Hex(Header.SeqNum1, tInt) + _
        Dec_to_Hex(Header.SeqNum2, tInt) + _
        "00000000" + _
        Header.Parameter
    
    UDP_CreatePacket = Hex_to_Str(EncryptPacket(Packet))
End Function

Function UDP_CreatePacketSeq(Header As UDP_CLIENT_HEADER, ByVal Seq As Integer) As String
    Dim Packet As String

    Packet = _
        ICQ_UDP_VERSION + _
        "00000000" + _
        Dec_to_Hex(Header.uin, tLong) + _
        Dec_to_Hex(Header.SessionID, tLong) + _
        Dec_to_Hex(Header.Command, tInt) + _
        Dec_to_Hex(Seq, tInt) + _
        Dec_to_Hex(0, tInt) + _
        "00000000" + _
        Header.Parameter
    
    UDP_CreatePacketSeq = Hex_to_Str(EncryptPacket(Packet))
End Function

Function CryptPacket(ByVal Packet As String, CheckCode As String) As String
    Dim PacketLength As Long, _
        Code1 As String, _
        Code2 As String, _
        ReadWritePos As Integer, _
        TablePos As Integer, _
        SubOutput As String, _
        Output As String
            
    PacketLength = CLng(Len(Packet) / 2)
    
    Code1 = hMul(Hex$(PacketLength), "68656C6C")
    Code1 = hFill(hAdd(Code1, CheckCode), tLong)
    
    ReadWritePos = &HA
    Do While ReadWritePos < PacketLength
        TablePos = ReadWritePos Mod &H100
        Code2 = hFill(hAdd(Code1, PeekByte(ICQTableV5, TablePos)), tLong)
        
        SubOutput = Peek(Packet, ReadWritePos, tLong)
        SubOutput = hXor(SubOutput, Code2)
        SubOutput = hDump(hFill(SubOutput, tLong))
        Packet = hFill(Poke(Packet, SubOutput, ReadWritePos), PacketLength)
        ReadWritePos = ReadWritePos + 4
    Loop
    
    CryptPacket = Packet
End Function

Function DecryptPacket(ByVal strPacket As String) As String
    Dim CheckCode As String
    CheckCode = Descramble(strPacket)
    strPacket = CryptPacket(strPacket, CheckCode)
        
    DecryptPacket = strPacket
End Function

Function EncryptPacket(ByVal strPacket As String) As String
    Dim CheckCode As String
    CheckCode = CalcCheckCode(strPacket)
    strPacket = CryptPacket(strPacket, hDump(CheckCode))
    
    CheckCode = Scramble(hDump(CheckCode))
    strPacket = Poke(strPacket, CheckCode, 20)
    
    EncryptPacket = strPacket
End Function

Function CalcCheckCode(ByVal strPacket As String) As String
    Dim Number1 As String, Number2 As String, _
        Code As String, PacketLength As Integer, _
        B2 As String, B4 As String, B6 As String, B8 As String, _
        X4 As String, X3 As String, X2 As String, X1 As String, _
        R1 As String, R2 As String

    If Len(strPacket) Mod 2 = 1 Then strPacket = "0" + strPacket
    PacketLength = Len(strPacket) / 2

    Randomize Timer
    R1 = Abs(&H18 + Int(Rnd(Timer) * ((PacketLength - &H19) Mod &H100)))
    R2 = Abs(Int(Rnd(Timer) * &HFF))
    
    B2 = hFill(PeekByte(strPacket, 2), tByte)
    B4 = hFill(PeekByte(strPacket, 4), tByte)
    B6 = hFill(PeekByte(strPacket, 6), tByte)
    B8 = hFill(PeekByte(strPacket, 8), tByte)
     
    X4 = hFill(Hex$(R1), tByte)
    X3 = hNot(hFill(PeekByte(strPacket, R1), tByte))
    X2 = hFill(Hex$(R2), tByte)
    X1 = hNot(hFill(PeekByte(ICQTableV5, R2), tByte))
    
    Number1 = B8 + B4 + B2 + B6
    Number2 = X4 + X3 + X2 + X1
     
    Code = hXor(Number1, Number2)
    CalcCheckCode = hDump(Code)
End Function

Function Scramble(ByVal CheckCode As String) As String
    Dim a0$, a1$, a2$, a3$, a4$, TempOut$
    
    a0$ = hFill(hAnd(CheckCode, "1F"), tLong)
    a1$ = hFill(hAnd(CheckCode, "3E003E0"), tLong)
    a2$ = hFill(hAnd(CheckCode, "F8000400"), tLong)
    a3$ = hFill(hAnd(CheckCode, "F800"), tLong)
    a4$ = hFill(hAnd(CheckCode, "41F0000"), tLong)
    
    a0$ = BitShift(a0$, &HC, LeftShift)
    a1$ = BitShift(a1$, &H1, LeftShift)
    a2$ = BitShift(a2$, &HA, RightShift)
    a3$ = BitShift(a3$, &H10, LeftShift)
    a4$ = BitShift(a4$, &HF, RightShift)
    
    TempOut$ = "00"
    TempOut$ = hAdd(TempOut$, a0$)
    TempOut$ = hAdd(TempOut$, a1$)
    TempOut$ = hAdd(TempOut$, a2$)
    TempOut$ = hAdd(TempOut$, a3$)
    TempOut$ = hAdd(TempOut$, a4$)
    
    Scramble = hDump(TempOut$)
End Function

Function Descramble(strPacket As String) As String
    Dim CheckCode$, a0$, a1$, a2$, a3$, a4$, TempOut$
    CheckCode$ = Peek(strPacket, &H14, tLong)
    
    a0$ = hFill(hAnd(CheckCode$, "1F000"), tLong)
    a1$ = hFill(hAnd(CheckCode$, "7C007C0"), tLong)
    a2$ = hFill(hAnd(CheckCode$, "3E0001"), tLong)
    a3$ = hFill(hAnd(CheckCode$, "F8000000"), tLong)
    a4$ = hFill(hAnd(CheckCode$, "83E"), tLong)
    
    a0$ = BitShift(a0$, &HC, RightShift)
    a1$ = BitShift(a1$, &H1, RightShift)
    a2$ = BitShift(a2$, &HA, LeftShift)
    a3$ = BitShift(a3$, &H10, RightShift)
    a4$ = BitShift(a4$, &HF, LeftShift)

    TempOut$ = "00"
    TempOut$ = hAdd(TempOut$, a0$)
    TempOut$ = hAdd(TempOut$, a1$)
    TempOut$ = hAdd(TempOut$, a2$)
    TempOut$ = hAdd(TempOut$, a3$)
    TempOut$ = hAdd(TempOut$, a4$)
    
    Descramble = hFill(TempOut$, tLong)
End Function
