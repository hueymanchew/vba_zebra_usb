Attribute VB_Name = "modAPI64bit"
Option Explicit

Public Const DIGCF_PRESENT As Integer = &H2
Public Const DIGCF_DEVICEINTERFACE As Integer = &H10
Public Const DIGCF_ALLCLASSES As Integer = &H4
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_DEVICE = &H40
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_ALWAYS = 4
Public Const OPEN_EXISTING = 3

Public Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Public Type Device_Interface_Data
  cbSize As Long
  InterfaceClassGuid As GUID
  Flags As Long
  ReservedPtr As LongPtr
End Type

Public Type Device_Interface_Detail
  cbSize As Long
  DataPath(256) As Byte
End Type

Public Type Device_Interface_Detail1
  cbSize As Long
  DataPath As Byte
End Type


Public Type SP_DEVINFO_DATA
  cbSize As Long
  InterfaceClassGuid As GUID
  hDevInst As Long
  ReservedPtr As Long
End Type


Public Declare PtrSafe Function SetupDiGetDeviceInterfaceDetail Lib "setupapi.dll" Alias "SetupDiGetDeviceInterfaceDetailA" _
    (ByVal DeviceInfoSet As LongPtr, DeviceInterfaceData As Any, _
     DeviceInterfaceDetailData As Any, ByVal DeviceInterfaceDetailDataSize As Long, RequiredSize As Long, ByVal DeviceInfoData As Long) As Boolean

Public Declare PtrSafe Function SetupDiGetDeviceInterfaceDetail_Ptr Lib "setupapi.dll" Alias "SetupDiGetDeviceInterfaceDetailA" _
    (ByVal DeviceInfoSet As LongPtr, ByVal DeviceInterfaceData As LongPtr, _
     ByVal DeviceInterfaceDetailData As LongPtr, ByVal DeviceInterfaceDetailDataSize As Long, RequiredSize As Long, ByVal DeviceInfoData As Long) As Boolean



Public Declare PtrSafe Function SetupDiEnumDeviceInterfaces Lib "setupapi.dll" _
  (ByVal Handle As LongPtr, ByVal InfoPtr As LongPtr, GuidPtr As LongPtr, ByVal MemberIndex As Long, InterfaceDataPtr As LongPtr) As Boolean


Public Declare PtrSafe Function SetupDiGetClassDevs Lib "setupapi.dll" Alias "SetupDiGetClassDevsA" _
    (GuidPtr As LongPtr, ByVal EnumPtr As Long, ByVal hwndParent As LongPtr, ByVal Flags As Long) As LongPtr

Public Declare PtrSafe Function SetupDiDestroyDeviceInfoList Lib "setupapi.dll" (ByVal DeviceInfoSet As LongPtr) As Boolean

Public Declare PtrSafe Function CreateFileDevice Lib "kernel32" Alias "CreateFileA" _
  (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr

Public Declare PtrSafe Sub CloseHandle Lib "kernel32" (ByVal HandleToClose As LongPtr)

Public Declare PtrSafe Function WriteFile Lib "kernel32" _
  (ByVal Handle As LongPtr, ByVal Buffer As String, ByVal ByteCount As Long, BytesReturnedPtr As Long, ByVal OverlappedPtr As LongPtr) As Long


