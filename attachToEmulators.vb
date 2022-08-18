Module attachToEmulators
    ' Contains all the functions to attack to each emulator
    ' Note: SoH has its own module since it is rather extensive and technically not an emulator

    Public Function attachToProject64(Optional doOffsetScan As Boolean = False) As Process
        With frmTrackerOfTime
            ' First steps, make sure we are in the right bit mode and declare the process
            If frmTrackerOfTime.IS_64BIT Then Return Nothing
            Dim target As Process = Nothing

            ' Try to attach to the first instance of project64
            Try
                target = Process.GetProcessesByName("project64")(0)
            Catch ex As Exception
                Return Nothing
            End Try

            If doOffsetScan Then
                ' So some people's pj64 has a different offset, this will help determine it
                For i = &HDFD00000 To &HE01FFFFF Step 16
                    If Memory.ReadInt32(target, i + &H11A5EC) = 1514490948 Then
                        .rtbAddLine(Hex(i))
                    End If
                Next
                .rtbAddLine("Done")
                Return target
            End If

            ' I have found 3 different addresses when connecting to project 64
            For i = 0 To 3
                Select Case i
                    Case 0
                        .romAddrStart = &HDFE40000
                    Case 1
                        .romAddrStart = &HDFE70000
                    Case 2
                        .romAddrStart = &HDFFB0000
                    Case Else
                        Return Nothing
                End Select

                ' Try to read what should be the first part of the ZELDAZ check
                Dim ootCheck As Integer = 0

                Try
                    ootCheck = Memory.ReadInt32(target, .romAddrStart + &H11A5EC)
                Catch ex As Exception
                    MessageBox.Show("quickRead Problem: " & vbCrLf & ex.Message & vbCrLf & (.romAddrStart + &H11A5EC).ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try

                ' If it matches, set emulator variable and leave the FOR LOOP
                If ootCheck = 1514490948 Then
                    .emulator = "project64"
                    Exit For
                End If
            Next

            ' Return the process
            Return target
        End With
    End Function

    Public Function attachToM64PY() As Process
        ' First steps, make sure we are in the right bit mode and declare the process
        If frmTrackerOfTime.IS_64BIT Then Return Nothing
        Dim target As Process = Nothing


        ' Try to attach to m64py
        Try
            target = Process.GetProcessesByName("m64py")(0)
        Catch ex As Exception
            Return Nothing
        End Try

        ' Grab the address of the DLL
        SetProcessName("m64py")
        Dim addressDLL As Int64 = 0
        For Each mo As ProcessModule In target.Modules
            If LCase(mo.ModuleName) = "mupen64plus-audio-sdl.dll" Then
                addressDLL = mo.BaseAddress.ToInt64
                Exit For
            End If
        Next

        ' If found, then try to check for the first part of ZELDAZ
        If addressDLL = 0 Then Return Nothing
        frmTrackerOfTime.romAddrStart = ReadMemory(Of Integer)(addressDLL + &H172060)
        If ReadMemory(Of Integer)(frmTrackerOfTime.romAddrStart + &H11A5EC) = 1514490948 Then frmTrackerOfTime.emulator = "m64py"

        ' Return the process
        Return target
    End Function

    Public Function attachToBizHawk() As Process
        ' First steps, make sure we are in the right bit mode and declare the process
        If Not frmTrackerOfTime.IS_64BIT Then Return Nothing
        Dim target As Process = Nothing

        ' Try to attach to BizHawk
        Try
            target = Process.GetProcessesByName("emuhawk")(0)
        Catch ex As Exception
            Return Nothing
        End Try

        ' Grab the address of the DLL
        SetProcessName("emuhawk")
        Dim addressDLL As Int64 = 0
        For Each mo As ProcessModule In target.Modules
            If LCase(mo.ModuleName) = "mupen64plus.dll" Then
                addressDLL = mo.BaseAddress.ToInt64
                Exit For
            End If
        Next

        ' If dll is found, then set variables
        If addressDLL = 0 Then Return Nothing
        frmTrackerOfTime.romAddrStart64 = addressDLL + &H658E0
        frmTrackerOfTime.emulator = "emuhawk"

        ' Return the process
        Return target
    End Function

    Public Function attachToRMG() As Process
        ' First steps, make sure we are in the right bit mode and declare the process
        If frmTrackerOfTime.IS_64BIT = False Then Return Nothing
        Dim target As Process = Nothing

        ' Try to attach to RMG
        Try
            target = Process.GetProcessesByName("rmg")(0)
        Catch ex As Exception
            Return Nothing
        End Try

        ' Grab the address of the DLL
        SetProcessName("rmg")
        Dim addressDLL As Int64 = 0
        For Each mo As ProcessModule In target.Modules
            If LCase(mo.ModuleName) = "mupen64plus.dll" Then
                addressDLL = mo.BaseAddress.ToInt64
                Exit For
            End If
        Next

        ' If dll is found, then set variables. RMG is not static, so we have to search for the address
        If addressDLL = 0 Then Return Nothing
        addressDLL = addressDLL + &H29C15D8


        ' Read the first half of the address and make sure it is 8 digits long
        Dim readAddress As String = Hex(ReadMemory(Of Integer)(addressDLL))
        frmTrackerOfTime.fixHex(readAddress)

        ' Since it is a 64bit address, we need the other part of it
        readAddress = Hex(ReadMemory(Of Integer)(addressDLL + 4)) & readAddress

        ' Set it + 0x8000000 as starting address
        frmTrackerOfTime.romAddrStart64 = CLng("&H" & readAddress) + &H80000000&
        frmTrackerOfTime.emulator = "rmg"

        Return target
    End Function

    Public Function attachToM64P(Optional attempt As Byte = 0) As Process
        ' First steps, make sure we are in the right bit mode and declare the process
        If frmTrackerOfTime.IS_64BIT = False Then Return Nothing
        Dim target As Process = Nothing

        ' Try to attach to m64p
        Try
            target = Process.GetProcessesByName("mupen64plus-gui")(0)
        Catch ex As Exception
            Return Nothing
        End Try

        ' Prepare new address variable
        SetProcessName("mupen64plus-gui")
        Dim addressDLL As Int64 = 0
        Dim attemptOffset As Int64 = 0
        Dim attemptAdded As Int64 = 0

        ' Step through all modules to find mupen64plus.dll's base address
        For Each mo As ProcessModule In target.Modules
            If LCase(mo.ModuleName) = "mupen64plus.dll" Then
                addressDLL = mo.BaseAddress.ToInt64
                Exit For
            End If
        Next

        ' m64p attachment process changes depending on the builds
        For i = 0 To 3
            Select Case i
                Case 0
                    ' Builds October 27, 2021 to March 31, 2022
                    attemptOffset = &H29C95D8
                    attemptAdded = &H80000000&
                Case 1
                    ' Builds July 13, 2021 to October 11, 2021
                    attemptOffset = &HCA6B8
                Case 2
                    ' Builds October 27, 2021 to June, 2, 2022

                    ' I hate this sooo much. Rather than the confusing "find the address that has the address", it is:
                    ' "Find the address that, that has an address (with an increase of 0x140), that has the address we need"
                    attemptOffset = &H1177888
                    attemptAdded = 0
                    Dim readAddress As Integer = ReadMemory(Of Integer)(addressDLL + attemptOffset)

                    If Not readAddress = 0 Then
                        readAddress += &H140
                        Dim hexAddress As String = Hex(readAddress)
                        frmTrackerOfTime.fixHex(hexAddress)
                        hexAddress = Hex(ReadMemory(Of Integer)(addressDLL + attemptOffset + 4)) & hexAddress
                        attemptOffset = CLng("&H" & hexAddress) - addressDLL
                    End If
                Case Else
                    Return Nothing
            End Select

            ' Check if mupen64plus.dll was found
            If Not addressDLL = 0 Then
                ' Add location of variable to base address
                addressDLL = addressDLL + attemptOffset
                'rtbOutput.AppendText(Hex(addressDLL) & vbCrLf)
                ' Attach to process and set it as the current emulator
                ' Read the first half of the address
                Dim readAddress As String = Hex(ReadMemory(Of Integer)(addressDLL))
                ' Convert to hex
                If Not readAddress = "0" Then
                    ' Make sure length is 8 digits
                    frmTrackerOfTime.fixHex(readAddress)
                    ' Read the second half of the address
                    readAddress = Hex(ReadMemory(Of Integer)(addressDLL + 4)) & readAddress
                    frmTrackerOfTime.romAddrStart64 = CLng("&H" & readAddress) + attemptAdded
                    If ReadMemory(Of Integer)(frmTrackerOfTime.romAddrStart64 + &H11A5EC) = 1514490948 Then Exit For
                End If
            End If
        Next

        frmTrackerOfTime.emulator = "mupen64plus-gui"
        Return target
    End Function

    Public Function attachToRetroArch() As Process
        ' First steps, make sure we are in the right bit mode and declare the process
        If frmTrackerOfTime.IS_64BIT = False Then Return Nothing
        Dim target As Process = Nothing

        ' Try to retroarch
        Try
            target = Process.GetProcessesByName("retroarch")(0)
        Catch ex As Exception
            Return Nothing
        End Try

        ' Prepare new address variable
        SetProcessName("retroarch")
        Dim addressDLL As Int64 = 0

        ' Step through all modules to find RetroArch's mupen64plus.dll's base address
        For Each mo As ProcessModule In target.Modules
            If LCase(mo.ModuleName) = "mupen64plus_next_libretro.dll" Then
                addressDLL = mo.BaseAddress.ToInt64
                Exit For
            End If
        Next

        ' RetroArch will be coded for two cores, though no idea why someone would use the second. If it found the mupen64plus, go ahead with it
        If addressDLL <> 0 Then
            ' Add location of variable to base address
            addressDLL = addressDLL + &H8E795E0

            ' Read the first half of the address 
            Dim readAddress As String = Hex(ReadMemory(Of Integer)(addressDLL))
            frmTrackerOfTime.fixHex(readAddress)
            ' Read the second half of the address
            readAddress = Hex(ReadMemory(Of Integer)(addressDLL + 4)) & readAddress

            ' Set it + 0x8000000 as starting address
            frmTrackerOfTime.romAddrStart64 = CLng("&H" & readAddress) + &H80000000&
            frmTrackerOfTime.emulator = "retroarch - mupen64plus"
        Else
            ' Check for RetroArch's parallel core, again, for whatever reason but best to cover extra bases
            For Each mo As ProcessModule In target.Modules
                If LCase(mo.ModuleName) = "parallel_n64_libretro.dll" Then
                    addressDLL = mo.BaseAddress.ToInt64
                    Exit For
                End If
            Next
            If addressDLL = 0 Then Return Nothing

            ' Prepare new address variable
            Dim attemptOffset As Int64 = 0

            ' RA's parallel core has version differences to attach
            For i = 0 To 2
                Select Case i
                    Case 0
                        ' 1.9.0 - 1.10.2
                        attemptOffset = &H845000
                    Case 1
                        ' 1.10.3+
                        attemptOffset = &H844000
                    Case Else
                        Return Nothing
                End Select
                frmTrackerOfTime.romAddrStart64 = addressDLL + attemptOffset
                If ReadMemory(Of Integer)(frmTrackerOfTime.romAddrStart64 + &H11A5EC) = 1514490948 Then Exit For
            Next

            frmTrackerOfTime.emulator = "retroarch - parallel"
        End If

        ' Return the process
        Return target
    End Function

    Public Function attachToModLoader64() As Process
        ' This sub, and thus support for ML64, is credited to subenji, who also added the memory function to attach by process ID -- 2022.06.17

        If frmTrackerOfTime.IS_64BIT = False Then Return Nothing
        Dim target As Process = Nothing

        Try
            ' Try to attach to application
            ' target = Process.GetProcessesByName("modloader64-gui")(0)
            ''
            ' ModLoader64 runs 5 or 6 processes with only one of them actually being the emulator window, so we need to check the emulation is running
            ''
            Dim processes As Process() = Process.GetProcessesByName("modloader64-gui")
            If processes.Length = 0 Then
                ' If process was not found, just return
                Return Nothing
            End If
            For Each p As Process In processes
                For Each pModule As ProcessModule In p.Modules
                    If LCase(pModule.ModuleName) = "mupen64plus.dll" Then
                        ''
                        ' As the process is run several times and I need to specify by PID rather than name, I need to attach this once by hand
                        ''
                        If Not OpenProcessHandleById(p.Id) Then
                            ' I changed this to be a msgbox rather than the output RTB to move it to a module. -- Selene Tabacchini 2022.08.12
                            MsgBox("Attachment Problem: Could not open process handle: " & p.Id)
                            Return Nothing
                        End If

                        target = p
                        ' Prepare new address variable
                        Dim addressDLL As Int64 = 0
                        Dim attemptOffset As Int64 = 0
                        Dim attemptAdded As Int64 = 0
                        Dim positiveHit As Boolean = False
                        ''
                        ' These pointers to the Base ROM location were all listed as static, I felt it best just to add them all. The first one worked every time in testing.
                        ' I found adding 0x80000000 was never necessary.
                        ''
                        ' Increased this to 6, so that there will be an overflow handling to return nothing. -- Selene Tabacchini 2022.08.12
                        For attempt = 0 To 6
                            Select Case attempt
                                Case 0
                                    attemptOffset = &H116ECF8
                                Case 1
                                    attemptOffset = &H12EED10
                                Case 2
                                    attemptOffset = &H6A6F0
                                Case 3
                                    attemptOffset = &H6C400
                                Case 4
                                    attemptOffset = &H6C460
                                Case 5
                                    attemptOffset = &H6C538
                                Case Else
                                    Return Nothing
                            End Select

                            ''
                            ' As the emulator is the same mupen64plus module, the rest of the code is the same as the existing attachToM64P function.
                            ''
                            ' Step through all modules to find mupen64plus.dll's base address
                            For Each mo As ProcessModule In target.Modules
                                If LCase(mo.ModuleName) = "mupen64plus.dll" Then
                                    addressDLL = mo.BaseAddress.ToInt64
                                    Exit For
                                End If
                            Next
                            ' Check if mupen64plus.dll was found
                            If Not addressDLL = 0 Then
                                ' Add location of variable to base address
                                addressDLL = addressDLL + attemptOffset
                                ' Set it as the current emulator
                                frmTrackerOfTime.emulator = "modloader64-gui"
                                ' Read the first half of the address
                                Dim readR15 As Integer = ReadMemory(Of Integer)(addressDLL)
                                ' Convert to hex
                                Dim hexR15 As String = Hex(readR15)
                                If Not hexR15 = "0" Then
                                    ' Make sure length is 8 digit for any dropped 0's
                                    frmTrackerOfTime.fixHex(hexR15)
                                    ' Read the second half of the address
                                    readR15 = ReadMemory(Of Integer)(addressDLL + 4)
                                    ' Convert to hex and attach to first half
                                    hexR15 = Hex(readR15) & hexR15
                                    frmTrackerOfTime.romAddrStart64 = CLng("&H" & hexR15) + attemptAdded
                                    If ReadMemory(Of Integer)(frmTrackerOfTime.romAddrStart64 + &H11A5EC) = 1514490948 Then
                                        ' Note as a positive hit and exit the loop
                                        positiveHit = True
                                        Exit For
                                    End If

                                End If
                            End If
                        Next
                        ' If a positive hit, finish up. If not, reset the target to advance to the next process
                        If positiveHit Then
                            Exit For
                        Else
                            target = Nothing
                        End If
                    End If
                Next
                If target IsNot Nothing Then
                    Exit For
                End If
            Next
            ''
            ' I added a line to handle the case that the launcher had started but emulation hadn't yet. Printing here isn't really necessary.
            ''
            If target Is Nothing Then
                ' I changed this to be a msgbox rather than the output RTB to move it to a module. -- Selene Tabacchini 2022.08.12
                MsgBox("Found modloader64 but couldn't find a running emulator.")
                Return Nothing
            End If
        Catch ex As Exception
            If ex.Message = "Index was outside the bounds of the array." Then
                ' This is the expected error if process was not found, just return
                Return Nothing
            Else
                ' Any other error, output error message to textbox
                ' I changed this to be a msgbox rather than the output RTB to move it to a module. -- Selene Tabacchini 2022.08.12
                MsgBox("Attachment Problem: " & ex.Message)
                Return Nothing
            End If
        End Try
        Return target
    End Function
End Module
