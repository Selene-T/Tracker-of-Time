'Imports System.Runtime.CompilerServices ' VB doesn't have preprocessor macros, so using Inlining - https://stackoverflow.com/a/37761883

Module soh
    Public Const gSaveCtxOff As Integer = &HEC8560
    Private gGameData As Long = 0

    '<MethodImplAttribute(MethodImplOptions.AggressiveInlining)> 'This will instruct compiler to use aggressive inlining if possible. Should be just before the function definition
    Public Function SAV(offset As Integer) As Integer
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

    Public Sub sohSetup(ByVal startAddress As Int64)
        gGameData = ReadMemory(Of Long)(startAddress + &HE4D878)

        ' Force some settings off until we can get to them
        My.Settings.setShop = 0
        frmTrackerOfTime.updateLTB("ltbShopsanity")
        My.Settings.setScrub = False
        My.Settings.setCow = False
        frmTrackerOfTime.updateSettingsPanel()


        ' Still need 68 - 73

        ' Check that we have not already done this, in case this is triggered twice in one instance
        If frmTrackerOfTime.arrLocation(0) = &H11AD1C Then ' original offset for emulators
            For i = 0 To frmTrackerOfTime.arrLocation.Length - 1
                ' Skip over 60 through 99
                If i = 60 Then i = 100
                frmTrackerOfTime.arrLocation(i) = frmTrackerOfTime.arrLocation(i) + &HDADF94 ' add additional offset from emu into soh RAM
            Next
        End If

        'arrLocation(60) = &H11A640          ' *Biggoron Check, emu
        frmTrackerOfTime.arrLocation(60) = SAV(&H36)            ' - HALP IDK AAAA, seems to be in cur equipped equipment ???  ' *Biggoron Check
        frmTrackerOfTime.arrLocation(61) = SAV(&HEB0)           ' *Big Fish
        frmTrackerOfTime.arrLocation(62) = SAV(&HEC4)          ' *Events 1: Egg from Malon, Obtained Epona, Won Cow
        frmTrackerOfTime.arrLocation(63) = SAV(&HEC8)          ' *Events 2: Zora Diving Game, Darunia’s Joy
        frmTrackerOfTime.arrLocation(64) = SAV(&HECC)          ' *Events 3: Zelda’s Letter, Song from Impa, Sun Song??, opened Temple of Time, Rainbow Bridge
        frmTrackerOfTime.arrLocation(65) = SAV(&HED0)          ' *Events 5: Scarecrow as Adult
        frmTrackerOfTime.arrLocation(66) = SAV(&HED4)          ' *Events 6: Song at Colossus, Trials
        frmTrackerOfTime.arrLocation(67) = SAV(&HED8)          ' *Events 7: Saria Gift, Skulltula trades, Barrier Lowered

        ' NONFUNC/UNTESTED BELOW HERE
        'arrLocation(68) = &H11B4C0          ' *Item Collect #1
        frmTrackerOfTime.arrLocation(68) = SAV(&HEE2)          ' *Item Collect #1
        'arrLocation(69) = &H11B4C4          ' *Item Collection #2
        frmTrackerOfTime.arrLocation(69) = SAV(&HEE6)          ' *Item Collection #2
        'arrLocation(70) = &H11B4E8          ' *Item: Rolling Goron as Young + Adult Link
        frmTrackerOfTime.arrLocation(70) = SAV(&HF0A)          ' *Item: Rolling Goron as Young + Adult Link
        'arrLocation(71) = &H11B4EC          ' *Thaw Zora King
        frmTrackerOfTime.arrLocation(71) = SAV(&HF0E)          ' *Thaw Zora King
        'arrLocation(72) = &H11B4F8          ' *Items: 1st and 2nd Scrubs, Lost Dog
        frmTrackerOfTime.arrLocation(72) = SAV(&HF1A)          ' *Items: 1st and 2nd Scrubs, Lost Dog
        'arrLocation(73) = &H11B894          ' *Scarecrow Song
        frmTrackerOfTime.arrLocation(73) = SAV(&HF28)          ' *Scarecrow Song
        ' END OF NONFUNC/UNTESTED

        'frmTrackerOfTime.arrLocation(62) = SAV(&HEC4)          ' *Events 1: Egg from Malon, Obtained Epona, Won Cow
        'frmTrackerOfTime.arrLocation(63) = SAV(&HEC8)          ' *Events 2: Zora Diving Game, Darunia’s Joy
        'frmTrackerOfTime.arrLocation(64) = SAV(&HECC)          ' *Events 3: Zelda’s Letter, Song from Impa, Sun Song??, opened Temple of Time, Rainbow Bridge
        'frmTrackerOfTime.arrLocation(65) = SAV(&HED0)          ' *Events 5: Scarecrow as Adult
        'frmTrackerOfTime.arrLocation(66) = SAV(&HED4)          ' *Events 6: Song at Colossus, Trials
        'frmTrackerOfTime.arrLocation(67) = SAV(&HED8)          ' *Events 7: Saria Gift, Skulltula trades, Barrier Lowered


        frmTrackerOfTime.arrLocation(74) = SAV(&H9E)        ' *Equipment
        frmTrackerOfTime.arrLocation(75) = SAV(&H36)        ' *Check for Biggoron's Sword
        frmTrackerOfTime.arrLocation(76) = SAV(&HA4)        ' *Upgrades
        frmTrackerOfTime.arrLocation(77) = SAV(&HA8)        ' *Quest Items and Songs

        frmTrackerOfTime.arrLocation(78) = SAV(&HE90)       ' **Gold Skulltulas 1
        frmTrackerOfTime.arrLocation(79) = SAV(&HE94)       ' **Gold Skulltulas 2
        frmTrackerOfTime.arrLocation(80) = SAV(&HE98)       ' **Gold Skulltulas 3
        frmTrackerOfTime.arrLocation(81) = SAV(&HE9C)       ' **Gold Skulltulas 4
        frmTrackerOfTime.arrLocation(82) = SAV(&HEA0)       ' **Gold Skulltulas 5
        frmTrackerOfTime.arrLocation(83) = SAV(&HEA4)       ' **Gold Skulltulas 6
    End Sub
End Module
