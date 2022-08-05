Option Strict On

Imports System.Runtime.InteropServices
Imports System.Text

Module MemPlus
    <DllImport("kernel32.dll")> _
    Private Function OpenProcess(ByVal dwDesiredAccess As UInteger, <MarshalAs(UnmanagedType.Bool)> ByVal bInheritHandle As Boolean, ByVal dwProcessId As Integer) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Function WriteProcessMemory(ByVal hProcess As IntPtr, ByVal lpBaseAddress As IntPtr, ByVal lpBuffer As Byte(), ByVal nSize As IntPtr, <Out()> ByRef lpNumberOfBytesWritten As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Function ReadProcessMemory(ByVal hProcess As IntPtr, ByVal lpBaseAddress As IntPtr, <Out()> ByVal lpBuffer() As Byte, ByVal dwSize As IntPtr, ByRef lpNumberOfBytesRead As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function CloseHandle(ByVal hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Private Const PROCESS_VM_WRITE As UInteger = &H20
    Private Const PROCESS_VM_READ As UInteger = &H10
    Private Const PROCESS_VM_OPERATION As UInteger = &H8
    Private TargetProcess As String = "iw6mp64_ship"
    Private ProcessHandle As IntPtr = IntPtr.Zero
    Private LastKnownPID As Integer = -1
    Private useID As Boolean = False

    Private Function ProcessIDExists(ByVal pID As Integer) As Boolean
        For Each p As Process In Process.GetProcessesByName(TargetProcess)
            If p.Id = pID Then Return True
        Next
        Return False
    End Function

    Public Sub SetProcessName(ByVal processName As String)
        TargetProcess = processName
        If ProcessHandle <> IntPtr.Zero Then CloseHandle(ProcessHandle)
        LastKnownPID = -1
        ProcessHandle = IntPtr.Zero
        useID = False
    End Sub

    Public Function OpenProcessHandleById(id As Integer) As Boolean
        If ProcessHandle <> IntPtr.Zero Then CloseHandle(ProcessHandle)
        ProcessHandle = OpenProcess(PROCESS_VM_READ Or PROCESS_VM_WRITE Or PROCESS_VM_OPERATION, False, id)
        If ProcessHandle = IntPtr.Zero Then Return False
        LastKnownPID = id
        TargetProcess = Process.GetProcessById(id).ProcessName
        useID = True
        Return True
    End Function

    Public Function GetCurrentProcessName() As String
        Return TargetProcess
    End Function

    Public Function UpdateProcessHandle() As Boolean
        Try
            If LastKnownPID = -1 OrElse Not ProcessIDExists(LastKnownPID) Then
                If ProcessHandle <> IntPtr.Zero Then CloseHandle(ProcessHandle)
                Dim p() As Process = Process.GetProcessesByName(TargetProcess)
                If p.Length = 0 Then Return False
                LastKnownPID = p(0).Id
                ProcessHandle = OpenProcess(PROCESS_VM_READ Or PROCESS_VM_WRITE Or PROCESS_VM_OPERATION, False, p(0).Id)
                If ProcessHandle = IntPtr.Zero Then Return False
            End If
            Return True
        Catch ex As Exception
            frmTrackerOfTime.stopScanning()
            Return False
        End Try
    End Function

    Public Function ReadMemory(Of T)(ByVal address As Object) As T
        Return ReadMemory(Of T)(CLng(address))
    End Function

    Public Function ReadMemory(Of T)(ByVal address As Integer) As T
        Return ReadMemory(Of T)(New IntPtr(address), 0, False)
    End Function

    Public Function ReadMemory(Of T)(ByVal address As Long) As T
        Return ReadMemory(Of T)(New IntPtr(address), 0, False)
    End Function

    Public Function ReadMemory(Of T)(ByVal address As IntPtr) As T
        Return ReadMemory(Of T)(address, 0, False)
    End Function

    Public Function ReadMemory(ByVal address As IntPtr, ByVal length As Integer) As Byte()
        Return ReadMemory(Of Byte())(address, length, False)
    End Function

    Public Function ReadMemory(ByVal address As Integer, ByVal length As Integer) As Byte()
        Return ReadMemory(Of Byte())(New IntPtr(address), length, False)
    End Function

    Public Function ReadMemory(ByVal address As Long, ByVal length As Integer) As Byte()
        Return ReadMemory(Of Byte())(New IntPtr(address), length, False)
    End Function

    Public Function ReadMemory(Of T)(ByVal address As IntPtr, ByVal length As Integer, ByVal unicodeString As Boolean) As T
        Dim buffer() As Byte
        If GetType(T) Is GetType(String) Then
            If unicodeString Then buffer = New Byte(length * 2 - 1) {} Else buffer = New Byte(length - 1) {}
        ElseIf GetType(T) Is GetType(Byte()) Then
            buffer = New Byte(length - 1) {}
        Else
            buffer = New Byte(Marshal.SizeOf(GetType(T)) - 1) {}
        End If
        If Not UpdateProcessHandle() Then Return Nothing
        Dim success As Boolean = ReadProcessMemory(ProcessHandle, address, buffer, New IntPtr(buffer.Length), IntPtr.Zero)
        If Not success Then Return Nothing
        If GetType(T) Is GetType(Byte()) Then Return CType(CType(buffer, Object), T)
        If GetType(T) Is GetType(String) Then
            If unicodeString Then Return CType(CType(Encoding.Unicode.GetString(buffer), Object), T)
            Return CType(CType(Encoding.ASCII.GetString(buffer), Object), T)
        End If
        Dim gcHandle As GCHandle = gcHandle.Alloc(buffer, GCHandleType.Pinned)
        Dim returnObject As T = CType(Marshal.PtrToStructure(gcHandle.AddrOfPinnedObject, GetType(T)), T)
        gcHandle.Free()
        Return returnObject
    End Function

    Private Function GetObjectBytes(ByVal value As Object) As Byte()
        If value.GetType() Is GetType(Byte()) Then Return CType(value, Byte())
        Dim buffer(Marshal.SizeOf(value) - 1) As Byte
        Dim ptr As IntPtr = Marshal.AllocHGlobal(buffer.Length)
        Marshal.StructureToPtr(value, ptr, True)
        Marshal.Copy(ptr, buffer, 0, buffer.Length)
        Marshal.FreeHGlobal(ptr)
        Return buffer
    End Function

    Public Function WriteMemory(Of T)(ByVal address As Object, ByVal value As T) As Boolean
        Return WriteMemory(CLng(address), value)
    End Function

    Public Function WriteMemory(Of T)(ByVal address As Object, ByVal value As Object) As Boolean
        Return WriteMemory(CLng(address), CType(value, T))
    End Function

    Public Function WriteMemory(Of T)(ByVal address As Integer, ByVal value As T) As Boolean
        Return WriteMemory(New IntPtr(address), value)
    End Function

    Public Function WriteMemory(Of T)(ByVal address As Integer, ByVal value As Object) As Boolean
        Return WriteMemory(address, CType(value, T))
    End Function

    Public Function WriteMemory(Of T)(ByVal address As Long, ByVal value As T) As Boolean
        Return WriteMemory(New IntPtr(address), value)
    End Function

    Public Function WriteMemory(Of T)(ByVal address As Long, ByVal value As Object) As Boolean
        Return WriteMemory(address, CType(value, T))
    End Function

    Public Function WriteMemory(Of T)(ByVal address As IntPtr, ByVal value As T) As Boolean
        Return WriteMemory(address, value, False)
    End Function

    Public Function WriteMemory(Of T)(ByVal address As IntPtr, ByVal value As Object) As Boolean
        Return WriteMemory(address, CType(value, T), False)
    End Function

    Public Function WriteMemory(Of T)(ByVal address As Object, ByVal value As T, ByVal unicode As Boolean) As Boolean
        Return WriteMemory(CLng(address), value, unicode)
    End Function

    Public Function WriteMemory(Of T)(ByVal address As Integer, ByVal value As T, ByVal unicode As Boolean) As Boolean
        Return WriteMemory(New IntPtr(address), value, unicode)
    End Function

    Public Function WriteMemory(Of T)(ByVal address As Long, ByVal value As T, ByVal unicode As Boolean) As Boolean
        Return WriteMemory(New IntPtr(address), value, unicode)
    End Function

    Public Function WriteMemory(Of T)(ByVal address As IntPtr, ByVal value As T, ByVal unicode As Boolean) As Boolean
        If Not UpdateProcessHandle() Then Return False
        Dim buffer() As Byte
        If TypeOf value Is String Then
            If unicode Then buffer = Encoding.Unicode.GetBytes(value.ToString()) Else buffer = Encoding.ASCII.GetBytes(value.ToString())
        Else
            buffer = GetObjectBytes(value)
        End If
        Dim result As Boolean = WriteProcessMemory(ProcessHandle, address, buffer, New IntPtr(buffer.Length), IntPtr.Zero)
        Return result
    End Function
End Module