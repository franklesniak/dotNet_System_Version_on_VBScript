# dotNet_System_Version_on_VBScript
A VBScript implementation of the .NET System.Version class.

Useful because .NET objects are not readily accessible in VBScript, and version-processing/comparison is a common systems administration activity.

## Class Specifications

### Methods

- Clone(ByRef objTargetVersionObject)
- CompareTo(ByVal objOtherVersionObject)
- CompareToString(ByVal strOtherVersion)
- Equals(ByVal objOtherVersionObject)
- GreaterThan(ByVal objOtherVersionObject)
- GreaterThanOrEqual(ByVal objOtherVersionObject)
- InitFromMajorMinor(ByVal lngMajor, ByVal lngMinor)
- InitFromMajorMinorBuild(ByVal lngMajor, ByVal lngMinor, ByVal lngBuild)
- InitFromMajorMinorBuildRevision(ByVal lngMajor, ByVal lngMinor, ByVal lngBuild, ByVal lngRevision)
- InitFromString(ByVal strVersion)
- LessThan(ByVal objOtherVersionObject)
- LessThanOrEqual(ByVal objOtherVersionObject)
- NotEquals(ByVal objOtherVersionObject)
- ToString()

### Properties

- Major (get)
- Minor (get)
- Build (get)
- Revision (get)
- MajorRevision (get)
- MinorRevision (get)

## Example Usage

Example 1:

    Dim objWMI
    Dim colItems
    Dim objItem
    Dim strOSString
    Dim versionOperatingSystem
    Dim intReturnCode

    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("Select Version from Win32_OperatingSystem")
    For Each objItem in colItems
        strOSString = objItem.Version
    Next

    Set versionOperatingSystem = New Version
    intReturnCode = versionOperatingSystem.InitFromString(strOSString)
    If intReturnCode = 0 Then
        ' Success
        If versionOperatingSystem.CompareToString("10.0") >= 0 Then
            WScript.Echo("Windows 10, Windows Server 2016, or newer!")
        Else
            WScript.Echo("Windows 8.1, Windows Server 2012 R2, or older!")
        End If
    Else
        WScript.Echo("An error occurred reading the OS version.")
    End If

Example 2:

    Dim objWMI
    Dim colItems
    Dim objItem
    Dim strOSString
    Dim versionCurrentOperatingSystem
    Dim versionWindows98
    Dim versionWindows98SE
    Dim versionWindowsME
    Dim intReturnCode
    Dim bool9x

    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("Select Version from Win32_OperatingSystem")
    For Each objItem in colItems
        strOSString = objItem.Version
    Next

    Set versionCurrentOperatingSystem = New Version
    intReturnCode = versionCurrentOperatingSystem.InitFromString(strOSString)
    If intReturnCode <> 0 Then
        WScript.Echo("Failed to get the current operating system version!")
    End If

    Set versionWindows98 = New Version
    intReturnCode = versionWindows98.InitFromMajorMinorBuild(4,10,1998)
    Set versionWindows98SE = New Version
    intReturnCode = versionWindows98SE.InitFromMajorMinorBuild(4,10,2222)
    Set versionWindowsME = New Version
    intReturnCode = versionWindowsME.InitFromMajorMinor(4,90)

    bool9x = False
    If versionCurrentOperatingSystem.GreaterThanOrEqual(versionWindows98) And versionCurrentOperatingSystem.LessThanOrEqual(versionWindows98SE) Then
        bool9x = True
    ElseIf (versionCurrentOperatingSystem.Major = versionWindowsME.Major) And (versionCurrentOperatingSystem.Minor = versionWindowsME.Minor) Then
        bool9x = True
    End If

    If bool9x Then
        WScript.Echo("Current OS is Windows 9x. It's 2020 (or later). What are you thinking?")
    Else
        WScript.Echo("Thank the maker! This OS is not Windows 9x.")
    End If

## External References

Since this code is based on it, it is helpful to review the [Microsoft Docs pages for the System.Version class][1].

[1]: https://docs.microsoft.com/en-us/dotnet/api/system.version?view=netcore-3.1
