Attribute VB_Name = "TimeZoneHelpers"
Option Compare Database

    Option Explicit
    
    Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
    End Type
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NOTE: If you are using the Windows WinAPI Viewer Add-In to get
    ' function declarations, not that there is an error in the
    ' TIME_ZONE_INFORMATION structure. It defines StandardName and
    ' DaylightName As 32. This is fine if you have an Option Base
    ' directive to set the lower bound of arrays to 1. However, if
    ' your Option Base directive is set to 0 or you have no
    ' Option Base diretive, the code won't work. Instead,
    ' change the (32) to (0 To 31).
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(0 To 31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(0 To 31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
    End Type
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' These give symbolic names to the time zone
    ' values returned by GetTimeZoneInformation .
    ''''''''''''''''''''''''''''''''''''''''''''''
    
    Private Enum TIME_ZONE

        TIME_ZONE_ID_INVALID = 0        ' Cannot determine DST
        TIME_ZONE_STANDARD = 1          ' Standard Time, not Daylight
        TIME_ZONE_DAYLIGHT = 2          ' Daylight Time, not Standard

    End Enum
    
    Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" _
        (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" _
        (lpSystemTime As SYSTEMTIME)
 
    Function IntArrayToString(V As Variant) As String

        Dim N As Long
        Dim S As String
        For N = LBound(V) To UBound(V)
            S = S & Chr(V(N))
        Next N
        IntArrayToString = S

    End Function

 Public Function CurrentTimeZone() As String

        Dim TZI As TIME_ZONE_INFORMATION
        Dim DST As TIME_ZONE
        
        DST = GetTimeZoneInformation(TZI)
        CurrentTimeZone = IntArrayToString(TZI.StandardName)
 End Function
