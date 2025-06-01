Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Globalization

<Flags()> _
Public Enum ASM_DISPLAY_FLAGS
    VERSION = &H1
    CULTURE = &H2
    PUBLIC_KEY_TOKEN = &H4
    PUBLIC_KEY = &H8
    [CUSTOM] = &H10
    PROCESSORARCHITECTURE = &H20
    LANGUAGEID = &H40
End Enum

Public Enum ASM_NAME
    ASM_NAME_PUBLIC_KEY = 0
    ASM_NAME_PUBLIC_KEY_TOKEN
    ASM_NAME_HASH_VALUE
    ASM_NAME_NAME
    ASM_NAME_MAJOR_VERSION
    ASM_NAME_MINOR_VERSION
    ASM_NAME_BUILD_NUMBER
    ASM_NAME_REVISION_NUMBER
    ASM_NAME_CULTURE
    ASM_NAME_PROCESSOR_ID_ARRAY
    ASM_NAME_OSINFO_ARRAY
    ASM_NAME_HASH_ALGID
    ASM_NAME_ALIAS
    ASM_NAME_CODEBASE_URL
    ASM_NAME_CODEBASE_LASTMOD
    ASM_NAME_NULL_PUBLIC_KEY
    ASM_NAME_NULL_PUBLIC_KEY_TOKEN
    ASM_NAME_CUSTOM
    ASM_NAME_NULL_CUSTOM
    ASM_NAME_MVID
    ASM_NAME_MAX_PARAMS
End Enum

<Flags()> _
Public Enum ASM_CACHE_FLAGS
    ASM_CACHE_ZAP = &H1
    ASM_CACHE_GAC = &H2
    ASM_CACHE_DOWNLOAD = &H4
End Enum

#Region "COM Interface Definitions"

<ComImport(), Guid("CD193BC0-B4BC-11d2-9833-00C04FC31D2E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
Public Interface IAssemblyName
    <PreserveSig()> _
    Function SetProperty(PropertyId As ASM_NAME, pvProperty As IntPtr, cbProperty As Integer) As Integer

    <PreserveSig()> _
    Function GetProperty(PropertyId As ASM_NAME, pvProperty As IntPtr, ByRef pcbProperty As Integer) As Integer

    <PreserveSig()> _
    Overloads Function Finalize() As Integer

    <PreserveSig()> _
    Function GetDisplayName(<Out(), MarshalAs(UnmanagedType.LPWStr)> szDisplayName As StringBuilder, ByRef pccDisplayName As Integer, dwDisplayFlags As ASM_DISPLAY_FLAGS) As Integer

    <PreserveSig()> _
    Function BindToObject(ByRef refIID As Guid, <MarshalAs(UnmanagedType.IUnknown)> pUnkSink As Object, <MarshalAs(UnmanagedType.IUnknown)> pUnkContext As Object, <MarshalAs(UnmanagedType.LPWStr)> szCodeBase As String, llFlags As Long, pvReserved As IntPtr, _
   cbReserved As Integer, ByRef ppv As IntPtr) As Integer

    <PreserveSig()> _
    Function GetName(ByRef lpcwBuffer As Integer, <Out(), MarshalAs(UnmanagedType.LPWStr)> pwzName As StringBuilder) As Integer

    <PreserveSig()> _
    Function GetVersion(ByRef pdwVersionHi As Integer, ByRef pdwVersionLow As Integer) As Integer

End Interface

<ComImport(), Guid("21b8916c-f28e-11d2-a473-00c04f8ef448"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
Public Interface IAssemblyEnum

    <PreserveSig()> _
    Function GetNextAssembly(pvReserved As IntPtr, ByRef ppName As IAssemblyName, dwFlags As Integer) As Integer

    <PreserveSig()> _
    Function Reset() As Integer

    <PreserveSig()> _
    Function Clone(ByRef ppEnum As IAssemblyEnum) As Integer
End Interface

#End Region

Public Class AssemblyCache

    Private Shared oList As New Hashtable

    <DllImport("fusion.dll", SetLastError:=True, PreserveSig:=False)> _
    Private Shared Sub CreateAssemblyEnum(ByRef pEnum As IAssemblyEnum, pUnkReserved As IntPtr, pName As IAssemblyName, dwFlags As ASM_CACHE_FLAGS, pvReserved As IntPtr)
    End Sub

    Public Shared Function LoadClass(ByVal sAssemblyName As String, ByVal sClassName As String) As Object

        Dim sFullName As String = GetFullAssemblyName(sAssemblyName)

        'Get Assembly
        Dim oAssembly As Reflection.Assembly
        Try
            oAssembly = Reflection.[Assembly].Load(sFullName)
        Catch
            Throw New ArgumentException("Can't load assembly " + sAssemblyName)
        End Try

        'Get Class
        Dim oType As Type = oAssembly.GetType(sClassName, False, False)
        If oType Is Nothing Then
            Throw New ArgumentException("Can't load type " + sClassName)
        End If

        'Instantiate
        Dim oTypes(-1) As Type
        Dim oInfo As Reflection.ConstructorInfo = oType.GetConstructor(oTypes)
        Dim oRetObj As Object = oInfo.Invoke(Nothing)
        If oRetObj Is Nothing Then
            Throw New ArgumentException("Can't instantiate type " + sClassName)
        End If

        Return oRetObj
    End Function

    Public Shared Function GetFullAssemblyName(sAssemblyName As String) As String

        If oList.ContainsKey(sAssemblyName) Then
            Return oList(sAssemblyName)
        End If

        Dim ae As IAssemblyEnum = CreateGACEnum()
        Dim an As IAssemblyName = Nothing

        While GetNextAssembly(ae, an) = 0
            If GetName(an) = sAssemblyName Then
                Dim name As System.Reflection.AssemblyName = GetAssemblyName(an)
                Dim sRet As String = name.FullName
                oList.Add(sAssemblyName, sRet)
                Return sRet
            End If
        End While

        Return ""
    End Function

    Public Shared Function GetAssemblyName(nameRef As IAssemblyName) As System.Reflection.AssemblyName
        Dim name As New System.Reflection.AssemblyName()
        name.Name = GetName(nameRef)
        name.Version = GetVersion(nameRef)
        name.CultureInfo = GetCulture(nameRef)
        name.SetPublicKeyToken(GetPublicKeyToken(nameRef))
        Return name
    End Function

    Public Shared Function GetName(name As IAssemblyName) As [String]
        Dim bufferSize As Integer = 255
        Dim buffer As New StringBuilder(CInt(bufferSize))
        name.GetName(bufferSize, buffer)
        Return buffer.ToString()
    End Function

    Public Shared Function GetVersion(name As IAssemblyName) As Version
        Dim major As Integer
        Dim minor As Integer
        name.GetVersion(major, minor)

        If minor < 0 Then
            minor = 0
        End If

        Return New Version(CInt(major) >> 16, CInt(major) And &HFFFF, CInt(minor) >> 16, CInt(minor) And &HFFFF)
    End Function

    Public Shared Function GetPublicKeyToken(name As IAssemblyName) As Byte()
        Dim result As Byte() = New Byte(7) {}
        Dim bufferSize As Integer = 8
        Dim buffer As IntPtr = Marshal.AllocHGlobal(CInt(bufferSize))
        name.GetProperty(ASM_NAME.ASM_NAME_PUBLIC_KEY_TOKEN, buffer, bufferSize)
        For i As Integer = 0 To 7
            result(i) = Marshal.ReadByte(buffer, i)
        Next
        Marshal.FreeHGlobal(buffer)
        Return result
    End Function

    Public Shared Function GetCulture(name As IAssemblyName) As CultureInfo
        Dim bufferSize As Integer = 255
        Dim buffer As IntPtr = Marshal.AllocHGlobal(CInt(bufferSize))
        name.GetProperty(ASM_NAME.ASM_NAME_CULTURE, buffer, bufferSize)
        Dim result As String = Marshal.PtrToStringAuto(buffer)
        Marshal.FreeHGlobal(buffer)
        Return New CultureInfo(result)
    End Function


    Public Shared Function CreateGACEnum() As IAssemblyEnum
        Dim ae As IAssemblyEnum = Nothing
        AssemblyCache.CreateAssemblyEnum(ae, CType(0, IntPtr), Nothing, ASM_CACHE_FLAGS.ASM_CACHE_GAC, CType(0, IntPtr))
        Return ae
    End Function

    Public Shared Function GetNextAssembly(enumerator As IAssemblyEnum, ByRef name As IAssemblyName) As Integer
        Return enumerator.GetNextAssembly(CType(0, IntPtr), name, 0)
    End Function



End Class
