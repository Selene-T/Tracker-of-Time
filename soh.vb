﻿Module soh
    Private gSaveCtxOff As Integer = 0
    Private sohVersion As Integer = 0
    Private gGameData As Long = 0
    Public gRoomAddr As Long = 0

    Public Function SAV(offset As Integer) As Integer
        ' Somewhere between 5.0.1 and 5.1.4, they shifted the save data 0x0E at the inventory
        Select Case sohVersion
            Case 514
                If offset = &H4 Then offset = -2
                If offset >= &H78 Then offset += &HE
        End Select
        Return gSaveCtxOff + offset
    End Function

    Public Function GDATA(ByVal offset As Integer, Optional bytes As Byte = 4) As UInteger
        Select Case bytes
            Case 1
                Return ReadMemory(Of Byte)(gGameData + offset)
            Case 2
                Return ReadMemory(Of UInt16)(gGameData + offset)
            Case Else
                Return ReadMemory(Of UInteger)(gGameData + offset)
        End Select
    End Function

    Public Function sohVer() As Integer
        Return sohVersion
    End Function

    Private Function getSohVersion(startAddress As Int64) As Integer ' Non-writable area String search
        If ReadMemory(Of String)(startAddress + &HF15490, 20, False) = "BRADLEY ECHO (5.1.4)" Then Return 514
        If ReadMemory(Of String)(startAddress + &HAEFD10, 19, False) = "FLYNN BRAVO (5.0.1)" Then Return 501
        If ReadMemory(Of String)(startAddress + &HAA7570, 19, False) = "ZHORA DELTA (4.0.3)" Then Return 403
        If ReadMemory(Of String)(startAddress + &HAA7FE0, 21, False) = "ZHORA CHARLIE (4.0.2)" Then Return 402
        If ReadMemory(Of String)(startAddress + &HAECD38, 21, False) = "RACHAEL BRAVO (3.0.1)" Then Return 301
        If ReadMemory(Of String)(startAddress + &HAECB78, 20, False) = "RACHAEL ALFA (3.0.0)" Then Return 300
        Return Nothing
    End Function

    Public Sub sohSetup(ByVal startAddress As Int64)
        sohVersion = getSohVersion(startAddress)
        Dim gOffset As Long = 0
        Dim gShift As Long = 0

        Select Case sohVersion
            Case 300
                gSaveCtxOff = &HEC8560
                gOffset = &HE4D878
            Case 301
                gSaveCtxOff = &HECA260
                gOffset = &HE7D878
            Case 402
                gSaveCtxOff = &HEC56A0
                gOffset = &HE492E8
                gShift = -&H40
            Case 403
                ' Rupees: + EC570E
                gSaveCtxOff = &HEC56E0
                gOffset = &HE78398
                gShift = -&H40
            Case 501
                ' Rupees: soh.exe + F07B0E: -2E = F07AE0
                gSaveCtxOff = &HF07AE0
                ' Scene rcx + 1C0: + E9A698
                gOffset = &HE9A698
                gShift = -&H40
                ' Room addr: + D1D858
                ' This moves around too much, while the previous version do not.
                gRoomAddr = &HD1D858
            Case 514
                ' Rupees (2 byte): soh.exe + 1298074: -2E = 1298046
                gSaveCtxOff = &H1298074 - &H2E
                ' Scene (1 byte) rax + 1C0: KF = 85 / Mido's = 40: soh.exe + 1238FF8
                gOffset = &H1238FF8
                gShift = -&H40
                ' Room addr for minimap (1 byte): DT 1F = 0 / DT 2F = 11 > 1 > 2 / DT 3F = 12 > 10: soh.exe + 1061F74
                ' This moves around too much, while the previous version do no.
                gRoomAddr = &H1061F74
            Case Else
                Exit Sub
        End Select

        gGameData = ReadMemory(Of Long)(startAddress + gOffset)
        gGameData = gGameData + gShift

        frmTrackerOfTime.arrLocation(0) = SAV(&H750)
        frmTrackerOfTime.arrLocation(1) = SAV(&H9B8)
        frmTrackerOfTime.arrLocation(2) = SAV(&HA60)
        frmTrackerOfTime.arrLocation(3) = SAV(&H154)
        frmTrackerOfTime.arrLocation(4) = SAV(&H170)
        frmTrackerOfTime.arrLocation(5) = SAV(&H1A8)
        frmTrackerOfTime.arrLocation(6) = SAV(&H1C4)
        frmTrackerOfTime.arrLocation(7) = SAV(&H1E0)
        frmTrackerOfTime.arrLocation(8) = SAV(&H218)
        frmTrackerOfTime.arrLocation(9) = SAV(&H250)
        frmTrackerOfTime.arrLocation(10) = SAV(&H2C0)
        frmTrackerOfTime.arrLocation(11) = SAV(&H2DC)
        frmTrackerOfTime.arrLocation(12) = SAV(&H2F8)
        frmTrackerOfTime.arrLocation(13) = SAV(&H314)
        frmTrackerOfTime.arrLocation(14) = SAV(&H330)
        frmTrackerOfTime.arrLocation(15) = SAV(&H34C)
        frmTrackerOfTime.arrLocation(16) = SAV(&H368)
        frmTrackerOfTime.arrLocation(17) = SAV(&H384)
        frmTrackerOfTime.arrLocation(18) = SAV(&H6E8)
        frmTrackerOfTime.arrLocation(19) = SAV(&H7AC)
        frmTrackerOfTime.arrLocation(20) = SAV(&H8C4)
        frmTrackerOfTime.arrLocation(21) = SAV(&H934)
        frmTrackerOfTime.arrLocation(22) = SAV(&H9F8)
        frmTrackerOfTime.arrLocation(23) = SAV(&HA14)
        frmTrackerOfTime.arrLocation(24) = SAV(&HA68)
        frmTrackerOfTime.arrLocation(25) = SAV(&HAA0)
        frmTrackerOfTime.arrLocation(26) = SAV(&HABC)
        frmTrackerOfTime.arrLocation(27) = SAV(&HAF4)
        frmTrackerOfTime.arrLocation(28) = SAV(&HB64)
        frmTrackerOfTime.arrLocation(29) = SAV(&HB80)
        frmTrackerOfTime.arrLocation(30) = SAV(&HB9C)
        frmTrackerOfTime.arrLocation(31) = SAV(&HD8)
        frmTrackerOfTime.arrLocation(32) = SAV(&HF4)
        frmTrackerOfTime.arrLocation(33) = SAV(&H110)
        frmTrackerOfTime.arrLocation(34) = SAV(&H12C)
        frmTrackerOfTime.arrLocation(35) = SAV(&H148)
        frmTrackerOfTime.arrLocation(36) = SAV(&H164)
        frmTrackerOfTime.arrLocation(37) = SAV(&H180)
        frmTrackerOfTime.arrLocation(38) = SAV(&H19C)
        frmTrackerOfTime.arrLocation(39) = SAV(&H1B8)
        frmTrackerOfTime.arrLocation(40) = SAV(&H1D4)
        frmTrackerOfTime.arrLocation(41) = SAV(&H1F0)
        frmTrackerOfTime.arrLocation(42) = SAV(&H20C)
        frmTrackerOfTime.arrLocation(43) = SAV(&H244)
        frmTrackerOfTime.arrLocation(44) = SAV(&H2D0)
        frmTrackerOfTime.arrLocation(45) = SAV(&H538)
        frmTrackerOfTime.arrLocation(46) = SAV(&H7A0)
        frmTrackerOfTime.arrLocation(47) = SAV(&H7BC)
        frmTrackerOfTime.arrLocation(48) = SAV(&H7D8)
        frmTrackerOfTime.arrLocation(49) = SAV(&H7F4)
        frmTrackerOfTime.arrLocation(50) = SAV(&H8B8)
        frmTrackerOfTime.arrLocation(51) = SAV(&HA24)
        frmTrackerOfTime.arrLocation(52) = SAV(&HA5C)
        frmTrackerOfTime.arrLocation(53) = SAV(&HA78)
        frmTrackerOfTime.arrLocation(54) = SAV(&HAB0)
        frmTrackerOfTime.arrLocation(55) = SAV(&HAE8)
        frmTrackerOfTime.arrLocation(56) = SAV(&HB04)
        frmTrackerOfTime.arrLocation(57) = SAV(&HB20)
        frmTrackerOfTime.arrLocation(58) = SAV(&HB58)
        frmTrackerOfTime.arrLocation(59) = SAV(&HB90)
        frmTrackerOfTime.arrLocation(60) = SAV(&H36)           ' *Biggoron Check
        frmTrackerOfTime.arrLocation(61) = SAV(&HEB0)          ' *Big Fish
        frmTrackerOfTime.arrLocation(62) = SAV(&HEC4)          ' *Events 1: Egg from Malon, Obtained Epona, Won Cow
        frmTrackerOfTime.arrLocation(63) = SAV(&HEC8)          ' *Events 2: Zora Diving Game, Darunia’s Joy
        frmTrackerOfTime.arrLocation(64) = SAV(&HECC)          ' *Events 3: Zelda’s Letter, Song from Impa, Sun Song??, opened Temple of Time, Rainbow Bridge
        frmTrackerOfTime.arrLocation(65) = SAV(&HED4)          ' *Events 5: Scarecrow as Adult
        frmTrackerOfTime.arrLocation(66) = SAV(&HED8)          ' *Events 6: Song at Colossus, Trials
        frmTrackerOfTime.arrLocation(67) = SAV(&HEDC)          ' *Events 7: Saria Gift, Skulltula trades, Barrier Lowered
        frmTrackerOfTime.arrLocation(68) = SAV(&HEE0)          ' *Item Collect #1
        frmTrackerOfTime.arrLocation(69) = SAV(&HEE4)          ' *Item Collection #2
        frmTrackerOfTime.arrLocation(70) = SAV(&HF08)          ' *Item: Rolling Goron as Young + Adult Link
        frmTrackerOfTime.arrLocation(71) = SAV(&HF0E)          ' *Thaw Zora King
        frmTrackerOfTime.arrLocation(72) = SAV(&HF1A)          ' *Items: 1st and 2nd Scrubs, Lost Dog
        frmTrackerOfTime.arrLocation(73) = SAV(&H1288)         ' *Scarecrow Song
        frmTrackerOfTime.arrLocation(74) = SAV(&H9E)           ' *Equipment
        frmTrackerOfTime.arrLocation(75) = SAV(&H36)           ' *Check for Biggoron's Sword
        frmTrackerOfTime.arrLocation(76) = SAV(&HA4)           ' *Upgrades
        frmTrackerOfTime.arrLocation(77) = SAV(&HA8)           ' *Quest Items and Songs
        frmTrackerOfTime.arrLocation(78) = SAV(&HE90)          ' **Gold Skulltulas 1
        frmTrackerOfTime.arrLocation(79) = SAV(&HE94)          ' **Gold Skulltulas 2
        frmTrackerOfTime.arrLocation(80) = SAV(&HE98)          ' **Gold Skulltulas 3
        frmTrackerOfTime.arrLocation(81) = SAV(&HE9C)          ' **Gold Skulltulas 4
        frmTrackerOfTime.arrLocation(82) = SAV(&HEA0)          ' **Gold Skulltulas 5
        frmTrackerOfTime.arrLocation(83) = SAV(&HEA4)          ' **Gold Skulltulas 6
        frmTrackerOfTime.arrLocation(100) = SAV(&H694)
        frmTrackerOfTime.arrLocation(101) = SAV(&H6CC)
        frmTrackerOfTime.arrLocation(102) = SAV(&H11C)
        frmTrackerOfTime.arrLocation(103) = SAV(&HDC)
        frmTrackerOfTime.arrLocation(104) = SAV(&HF8)
        frmTrackerOfTime.arrLocation(105) = SAV(&H114)
        frmTrackerOfTime.arrLocation(106) = SAV(&H130)
        frmTrackerOfTime.arrLocation(107) = SAV(&H14C)
        frmTrackerOfTime.arrLocation(108) = SAV(&H168)
        frmTrackerOfTime.arrLocation(109) = SAV(&H184)
        frmTrackerOfTime.arrLocation(110) = SAV(&H1A0)
        frmTrackerOfTime.arrLocation(111) = SAV(&H1D8)
        frmTrackerOfTime.arrLocation(112) = SAV(&H210)
        frmTrackerOfTime.arrLocation(113) = SAV(&H248)
        frmTrackerOfTime.arrLocation(114) = SAV(&HB2C)
        frmTrackerOfTime.arrLocation(115) = SAV(&HB94)
        frmTrackerOfTime.arrLocation(116) = SAV(&HA0C)
        frmTrackerOfTime.arrLocation(117) = SAV(&H1BC)
        frmTrackerOfTime.arrLocation(118) = SAV(&HBAC)
        frmTrackerOfTime.arrLocation(119) = SAV(&HA40)
        frmTrackerOfTime.arrLocation(120) = SAV(&H810)
        frmTrackerOfTime.arrLocation(121) = SAV(&H784)
        frmTrackerOfTime.arrLocation(122) = SAV(&H74C)
        frmTrackerOfTime.arrLocation(123) = SAV(&H82C)
        frmTrackerOfTime.arrLocation(124) = SAV(&H9B4)
    End Sub
End Module