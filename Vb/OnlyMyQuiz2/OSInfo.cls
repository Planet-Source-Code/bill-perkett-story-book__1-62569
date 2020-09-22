VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OSInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  API Declarations - Constants

Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1&
Private Const VER_PLATFORM_WIN32_NT As Long = 2&

' The following declarations are valid only for Windows 2000 and later

' Product Types
Private Const VER_NT_WORKSTATION As Long = &H1
Private Const VER_NT_DOMAIN_CONTROLLER As Long = &H2
Private Const VER_NT_SERVER As Long = &H3

' VerifyVersionInfo Masks
Private Const VER_MINORVERSION As Long = &H1
Private Const VER_MAJORVERSION As Long = &H2
Private Const VER_BUILDNUMBER As Long = &H4
Private Const VER_PLATFORMID As Long = &H8
Private Const VER_SERVICEPACKMINOR As Long = &H10
Private Const VER_SERVICEPACKMAJOR As Long = &H20
Private Const VER_SUITENAME As Long = &H40
Private Const VER_PRODUCT_TYPE As Long = &H80

' API Declarations - Types

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      ' Maintenance string for PSS usage
End Type

' This structure is for Windows 2000 and later

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128        ' Maintenance string for PSS usage
    wSPMajor As Integer                 ' Service Pack Major Version
    wSPMinor As Integer                 ' Service Pack Minor Version
    wSuiteMask As Integer               ' Suite Identifier
    bProductType As Byte                ' Server / Workstation / Domain Controller ?
    bReserved As Byte                   ' Reserved
End Type

' API Declarations - Functions
Private Declare Function GetOSVersion Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetOSVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Private Declare Function VerifyVersionInfo Lib "kernel32" Alias "VerifyVersionInfoA" (lpVersionInformation As OSVERSIONINFOEX, ByVal dwTypeMask As Long, ByVal dwlConditionMask As Currency) As Long
Private Declare Function VerSetConditionMask Lib "kernel32" (ByVal ConditionMask As Currency, ByVal TypeMask As Long, ByVal Condition As Byte) As Currency
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal numBytes As Long)

' Enums
Public Enum OSType
    Win32s
    Win95
    Win95OSR2
    Win98
    Win98R2
    WinME
    WinNT3
    WinNT31
    WinNT35
    WinNT4
    Win2000
    WinDontKnow
End Enum

' The Following Enums are for Windows 2000 or later.
Public Enum ProductType
    OSTypeUnSupported = 0
    OSTypeWorkStation = VER_NT_WORKSTATION
    OSTypeDomainController = VER_NT_DOMAIN_CONTROLLER
    OSTypeServer = VER_NT_SERVER
End Enum

Public Enum CheckOS
    Win9x = 0
    WinNT = 1
End Enum

' Class Variables
Dim m_OS As OSType                  ' OS Enum
Dim m_OSString() As String          ' The Name of the OS

Dim m_Major As Long                 ' Major Version
Dim m_Minor As Long                 ' Minor Version
Dim m_BuildNumber As Long           ' Build Number of the OS
Dim m_SPMajor As Long               ' Service Pack Minor Version. Win2000 or later
Dim m_SPMinor As Long               ' Service Pack Minor Version. Win2000 or later
Dim m_PSSInfo As String             ' PSS Info. Contains Service Pack Info in NT

' The Following vars are for Windows 2000 or later
Dim m_IsVerifyAvailable As Boolean  ' Does the OS support VerifyVersionInfo ?
Dim m_ProductType As ProductType

Private Sub Class_Initialize()

' Calls GetVersionInfo to Initialize the Object

Dim osinfo As OSVERSIONINFO
Dim OSInfoEx As OSVERSIONINFOEX
Dim APIRetVal As Long
Dim I As Long

' Before anything, initialize the OSString Array with elements corresponding to
' the Enum values of OSType

ReDim m_OSString(0 To OSType.WinDontKnow)
m_OSString(Win32s) = "Win32s"
m_OSString(Win95) = "Windows 95"
m_OSString(Win95OSR2) = "Windows 95 OSR 2"
m_OSString(Win98) = "Windows 98"
m_OSString(Win98R2) = "Windows 98 Second Edition"
m_OSString(WinME) = "Windows ME"
m_OSString(WinNT3) = "Windows NT 3.0"
m_OSString(WinNT31) = "Windows NT 3.1"
m_OSString(WinNT35) = "Windows NT 3.5"
m_OSString(WinNT4) = "Windows NT 4.0"
m_OSString(Win2000) = "Windows 2000"
m_OSString(WinDontKnow) = "[Unknown]"

' Initialize the Structure
osinfo.dwOSVersionInfoSize = Len(osinfo)
OSInfoEx.dwOSVersionInfoSize = Len(OSInfoEx)

APIRetVal = GetOSVersion(osinfo)

' Check if the OS is Windows 2000 or later. If so, it provides additional information
' through OSVERSIONINFOEX and Supports VerifyVersionInfo API.
If osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT And osinfo.dwMajorVersion >= 5 Then
    ' Use Windows 2000 OSINFOEX Structure
    GetOSVersionEx OSInfoEx
    ' Build the Results
    m_OS = Win2000
    
    m_Major = OSInfoEx.dwMajorVersion
    m_Minor = OSInfoEx.dwMinorVersion
    m_BuildNumber = OSInfoEx.dwBuildNumber
    m_SPMajor = OSInfoEx.wSPMajor
    m_SPMinor = OSInfoEx.wSPMinor
    m_PSSInfo = OSInfoEx.szCSDVersion
    
    ' Get PSSInfo String
    If Len(m_PSSInfo) > 0 Then
        If InStr(m_PSSInfo, Chr$(0)) > 0 Then
            m_PSSInfo = Left$(m_PSSInfo, InStr(m_PSSInfo, Chr$(0)) - 1)
        End If
    End If
    
    m_IsVerifyAvailable = True
    
    m_ProductType = OSInfoEx.bProductType
Else
    ' May be one of those Win9x or NT 3 to 4.51
    ' The Common Items are Major and Minor Verions and the PSS Strings
    ' As far as the Build Number goes, the NT/2000 systems return the Build Number in
    ' dwBuildNumber of the OSINFO structure. For Win9x, the HighWord of dwBuildNumber
    ' member contains the Major and Minor Versions and the BuildNumber is stored in the Lower
    ' order byte of the member

    ' Fill in the variables
    m_Major = osinfo.dwMajorVersion
    m_Minor = osinfo.dwMinorVersion
    m_BuildNumber = osinfo.dwBuildNumber
    m_PSSInfo = osinfo.szCSDVersion
    ' Features not available other than Windows 2000 or later
    m_IsVerifyAvailable = False
    m_ProductType = OSTypeUnSupported

    ' Get PSSInfo String
    If Len(m_PSSInfo) > 0 Then
        If InStr(m_PSSInfo, Chr$(0)) > 0 Then
            m_PSSInfo = Left$(m_PSSInfo, InStr(m_PSSInfo, Chr$(0)) - 1)
        End If
    End If

    ' Now, Check for Various versions

    Select Case osinfo.dwPlatformId

        Case VER_PLATFORM_WIN32_WINDOWS
            ' Check for Various Win9x Versions
        
            m_BuildNumber = LOWORD(m_BuildNumber)
    
            If m_Major = 4 And m_Minor = 0 Then
                If m_BuildNumber = 950 Then
                    m_OS = Win95
                Else  ' 1111 for OSR 2. For OSR 2.5 = ???
                    m_OS = Win95OSR2
                End If
        
            ElseIf m_Major = 4 And m_Minor = 10 Then
                If m_BuildNumber = 1998 Then
                    m_OS = Win98
                Else  ' Build Number is 2222.
                    m_OS = Win98R2
                End If
        
            ElseIf m_Major >= 4 And m_Minor > 10 Then
                m_OS = WinME  ' Version = 4.90 ?
        
                ' Perhaps, Win ME may support VerifyVersionInfo
                ' m_IsVerifyAvailable=True
            End If

        Case VER_PLATFORM_WIN32_NT
            ' Windows NT
            If m_Major = 3 And m_Minor = 0 Then
                m_OS = WinNT3
            ElseIf m_Major = 3 And m_Minor = 1 Then
                m_OS = WinNT31
            ElseIf m_Major = 4 Then
                m_OS = WinNT4
            End If
    
            m_BuildNumber = m_BuildNumber And &HFFFF&
    
            ' Get the SP Version from the PSS String (Not Tested)

            For I = 1 To 6
                If InStr(m_PSSInfo, Trim$(Str$(I))) > 0 Then
                    m_SPMajor = I
                    Exit For
                End If
            Next
    
            ' Check for 6a. Return 6.1 as the SP Version.
            If InStr(m_PSSInfo, "6a") > 0 Then
                m_SPMajor = 6
                m_SPMinor = 1
            End If
    
        Case Else
            ' The case that will never occur on Win32. It's Win 3.x
            m_OS = Win32s
    
    End Select
End If

End Sub

Public Function OSEnum() As OSType
Attribute OSEnum.VB_Description = "The Operating Systems Enum"

OSEnum = m_OS

End Function

Public Function OSName() As String
Attribute OSName.VB_Description = "Descriptive Name of the OS"
Attribute OSName.VB_UserMemId = 0

OSName = m_OSString(m_OS)

End Function

Public Function OSMajorVersion() As Long
Attribute OSMajorVersion.VB_Description = "OS Major Version"

OSMajorVersion = m_Major

End Function

Public Function OSMinorVersion() As Long
Attribute OSMinorVersion.VB_Description = "OS Minor Version"

OSMinorVersion = m_Minor

End Function

Public Function OSSPMajorVersion() As Long
Attribute OSSPMajorVersion.VB_Description = "Installed Service Pack Major Version (only for Win 2000 and above)"

OSSPMajorVersion = m_SPMajor

End Function

Public Function OSSPMinorVersion() As Long
Attribute OSSPMinorVersion.VB_Description = "Installed Service Pack Minor Version (only for Win 2000 and above)"

OSSPMinorVersion = m_SPMinor

End Function

Public Function OSBuildNumber() As Long
Attribute OSBuildNumber.VB_Description = "OS Build Number"

OSBuildNumber = m_BuildNumber

End Function

Public Function PSSInfo() As String
Attribute PSSInfo.VB_Description = "Product Support Services String"

PSSInfo = m_PSSInfo

End Function

Public Property Get IsNT() As Boolean
Attribute IsNT.VB_Description = "Returns True if the OS is Windows NT/2000"

IsNT = (m_OS >= WinNT3)

End Property

Public Property Get IsWin98() As Boolean
Attribute IsWin98.VB_Description = "Returns True if the OS is Windows 98"

IsWin98 = (m_OS = Win98 Or m_OS = Win98R2)

End Property

Public Property Get IsWin2K() As Boolean
Attribute IsWin2K.VB_Description = "Returns True if the OS is Windows 2000 or later"

IsWin2K = (m_OS = Win2000)

End Property

Public Property Get SupportsVerify() As Boolean
Attribute SupportsVerify.VB_Description = "Returns True if the OS Supports the VerifyVersionInfo API"

SupportsVerify = m_IsVerifyAvailable

End Property

Public Property Get OSProductType() As ProductType

OSProductType = m_ProductType

End Property

Private Function LOWORD(ByVal lData As Long)

' Returns the LoWORD of the Long

If (lData And &HFFFF&) > &H7FFF& Then
    lData = lData - &H10000
Else
    lData = lData And &HFFFF&
End If

LOWORD = lData

End Function

