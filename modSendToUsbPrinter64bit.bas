Attribute VB_Name = "modSendToUsbPrinter64bit"
Option Explicit

Sub SendToUsbPrinter64bit(strNrArt As String, strEan As String, strName As String)
  
  Dim PrnGUID As GUID
  Dim PnpHandle As LongPtr
  Dim DevIndex As Long, Result As Long, Success As Long, Ret As Long
  Dim BytesReturned As Long, BytesWritten As Long
  Dim DeviceHandle As LongPtr, DeviceName As String
  
  Dim DeviceInterfaceData     As Device_Interface_Data
  Dim FunctionClassDeviceData As Device_Interface_Detail
  Dim DeviceInterfaceDetail   As Device_Interface_Detail1

  Dim PrintOut As String, PrintOutA As String
  
  Dim SendToUsbPrinter As Boolean
  
  PrnGUID.Data1 = &H28D78FAD
  PrnGUID.Data2 = &H5A12
  PrnGUID.Data3 = &H11D1
  PrnGUID.Data4(0) = &HAE
  PrnGUID.Data4(1) = &H5B
  PrnGUID.Data4(2) = &H0
  PrnGUID.Data4(3) = &H0
  PrnGUID.Data4(4) = &HF8
  PrnGUID.Data4(5) = &H3
  PrnGUID.Data4(6) = &HA8
  PrnGUID.Data4(7) = &HC2

  PnpHandle = SetupDiGetClassDevs(ByVal VarPtr(PrnGUID.Data1), 0, 0, DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE)

  If (PnpHandle = -1) Then
    
    MsgBox "Could not attach to PnP node"
  
  Else
    DeviceInterfaceData.cbSize = LenB(DeviceInterfaceData)
    DevIndex = 0

    ' Should be a Do While -> looking for the correct device-name…
    ' If SetupDiEnumDeviceInterfaces(PnpHandle, 0, PrnGUID.Data1, DevIndex, DeviceInterfaceData.cbSize) Then
  
    Result = SetupDiEnumDeviceInterfaces(PnpHandle, 0, ByVal VarPtr(PrnGUID), DevIndex, ByVal VarPtr(DeviceInterfaceData))
    
    If Result Then
      
      'In VBA 32-bit cbSize = 5
      FunctionClassDeviceData.cbSize = 8
      Success = SetupDiGetDeviceInterfaceDetail(PnpHandle, DeviceInterfaceData, FunctionClassDeviceData, UBound(FunctionClassDeviceData.DataPath), BytesReturned, 0)
      
      If Success = 0 Then
        MsgBox "Could not get the name of this device"
      Else
        DeviceName = StrConv(FunctionClassDeviceData.DataPath(), vbUnicode)
        Debug.Print DeviceName
        DeviceHandle = CreateFileDevice(VarPtr(FunctionClassDeviceData.DataPath(0)), GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL + FILE_FLAG_SEQUENTIAL_SCAN, 0)

        If (DeviceHandle = -1) Then
            Debug.Print "Open failed on " & DeviceName; ""
            Debug.Print Err.LastDllError
        Else
          
          PrintOutA = "I8,B,001|Q200,024|q448|rN|S4|D7|ZT|JF|O|R4,0|f100|N|"
          PrintOut = "B33,101,0,E30,2,4,81,B,""" & strEan & """|A24,6,0,2,1,1,N,""" & strName & """|A24,28,0,2,1,1,N,""""|A43,49,0,4,2,2,N,""" & strNrArt & """|P7|N"
          PrintOut = Replace(PrintOutA & PrintOut, "|", Chr(13) & Chr(10))
  
          Ret = WriteFile(DeviceHandle, PrintOut, Len(PrintOut), BytesWritten, 0)
          Debug.Print "Sent; " & BytesWritten & "; bytes."""
          SendToUsbPrinter = True
          CloseHandle DeviceHandle
        End If
      End If
    Else
      MsgBox "Device not connected"
    End If
    
    SetupDiDestroyDeviceInfoList (PnpHandle)
  End If
  
End Sub



