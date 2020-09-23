Attribute VB_Name = "UDPQueue"
Option Explicit

Public Type UDPQueueItem
    Packet As String
    Attempt As Integer      'Amount of resent being done
End Type

Public Type ExpireData
    TotalQueue As Integer
    QueueList As String
End Type

Public qPacket(0 To 65535) As UDPQueueItem
Public qExpire(0 To 59) As ExpireData

Public Sub UDPQueue_Set(ByVal strPacket As String, Seq As Integer, Optional SendAttempt As Integer = 1)
    Dim ExpireTime As Integer
    qPacket(Seq).Packet = strPacket
    qPacket(Seq).Attempt = SendAttempt
    
    ExpireTime = (icqTimer + qIntervalTime) Mod 60
    With qExpire(ExpireTime)
        .TotalQueue = .TotalQueue + 1
        .QueueList = .QueueList + String$((.TotalQueue * 4) - Len(.QueueList), "0")
        .QueueList = _
            Poke( _
                .QueueList, _
                Dec_to_Hex(Seq, tInt), _
                (.TotalQueue - 1) * 2 _
            )
    End With
End Sub

Public Sub UDPQueue_ACK(Seq As Integer)
  Notify_UDPAck Seq
  If qPacket(Seq).Attempt <> 0 Then
    qPacket(Seq).Packet = ""
    qPacket(Seq).Attempt = 0
  End If
End Sub

Public Sub UDPQueue_Reset()
  Dim i As Integer
  For i = 0 To 59
    qExpire(i).QueueList = ""
    qExpire(i).TotalQueue = 0
  Next i
End Sub

Public Sub UDPQueue_CheckTime()
    Dim TimeSlot As Integer, _
        ExpireTime As Integer, _
        Seq As Integer, _
        i As Integer

    TimeSlot = icqTimer Mod 60
    With qExpire(TimeSlot)
    If .TotalQueue > 0 Then
        For i = .TotalQueue To 1 Step -1
            .TotalQueue = .TotalQueue - 1
            Seq = Val("&H" + Peek(.QueueList, (i - 1) * 2, tInt))
            .QueueList = CutTextR(.QueueList, 4)
        
            '---------- Packet Resend Module -----------
            With qPacket(Seq)
                If .Attempt > 0 Then
                    .Attempt = .Attempt + 1
                    Select Case .Attempt - 1
                        Case 6
                            UDPSend_TextCode "B_MESSAGE_ACK"
                            DebugTxt "UDPQueu", "After 6 unsuccesful resend attempt, we send 'B_MESSAGE_ACK'."
                        Case 12
                            UDPSend_Logout
                            DebugTxt "UDPQueu", "After 12 unsuccesful resend attempt, we decide to disconnect"
                            Exit Sub
                        Case Else
                            UDPSock_Send .Packet
                            DebugTxt "UDPQueu", "Resending packet" + Str$(Seq) + " (Attempt" + Str$(.Attempt - 1) + ")"
                    End Select
                    
                    UDPQueue_Set .Packet, Seq, .Attempt
                End If
            End With
            '---------- End Packet Resend Module -------
            
        Next i
    End If
  End With
End Sub
