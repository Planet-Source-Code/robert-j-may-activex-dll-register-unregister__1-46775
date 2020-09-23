<div align="center">

## ActiveX Dll Register/UnRegister


</div>

### Description

This code shows how to register and unregister ActiveX dlls programatically, without relying on regsvr32 for the task. It's built into a reuseable class that can be put in your own code or compiled into a dll. Based very loosely on code from Vasudevan S.
 
### More Info
 
The file name

7 flags. See the enumeration and sample code for details.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Robert J May](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/robert-j-may.md)
**Level**          |Intermediate
**User Rating**    |4.5 (54 globes from 12 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/robert-j-may-activex-dll-register-unregister__1-46775/archive/master.zip)

### API Declarations

```
Private Declare Function LoadLibraryRegister Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibraryRegister Lib "KERNEL32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Private Declare Function GetProcAddressRegister Lib "KERNEL32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThreadForRegister Lib "KERNEL32" Alias "CreateThread" (lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "KERNEL32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "KERNEL32" (ByVal dwExitCode As Long)
```


### Source Code

```
Option Explicit
Private Declare Function LoadLibraryRegister Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibraryRegister Lib "KERNEL32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Private Declare Function GetProcAddressRegister Lib "KERNEL32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThreadForRegister Lib "KERNEL32" Alias "CreateThread" (lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "KERNEL32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "KERNEL32" (ByVal dwExitCode As Long)
Private Const STATUS_WAIT_0 = &H0
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)
Private Const NOERRORS As Long = 0
Private Enum stRegisterStatus
  stFileCouldNotBeLoadedIntoMemorySpace = 1
  stNotAValidActiveXComponent = 2
  stActiveXComponentRegistrationFailed = 3
  stActiveXComponentRegistrationSuccessful = 4
  stActiveXComponentUnRegisterSuccessful = 5
  stActiveXComponentUnRegistrationFailed = 6
  stNoFileProvided = 7
End Enum
Public Function Register(ByVal p_sFileName As String) As Variant
  Dim lLib As Long
  Dim lProcAddress As Long
  Dim lThreadID As Long
  Dim lSuccess As Long
  Dim lExitCode As Long
  Dim lThreadHandle As Long
  Dim lRet As Long
  On Error GoTo ErrorHandler
  If lRet = NOERRORS Then
    If p_sFileName = "" Then
      lRet = stNoFileProvided
    End If
  End If
  If lRet = NOERRORS Then
    lLib = LoadLibraryRegister(p_sFileName)
    If lLib = 0 Then
      lRet = stFileCouldNotBeLoadedIntoMemorySpace
    End If
  End If
  If lRet = NOERRORS Then
    lProcAddress = GetProcAddressRegister(lLib, "DllRegisterServer")
    If lProcAddress = 0 Then
      lRet = stNotAValidActiveXComponent
    Else
      lThreadHandle = CreateThreadForRegister(0, 0, lProcAddress, 0, 0, lThreadID)
      If lThreadHandle <> 0 Then
         lSuccess = (WaitForSingleObject(lThreadHandle, 10000) = WAIT_OBJECT_0)
         If lSuccess = 0 Then
          Call GetExitCodeThread(lThreadHandle, lExitCode)
          Call ExitThread(lExitCode)
          lRet = stActiveXComponentRegistrationFailed
         Else
          lRet = stActiveXComponentRegistrationSuccessful
         End If
      End If
    End If
  End If
ExitRoutine:
  Register = lRet
  If lThreadHandle <> 0 Then
    Call CloseHandle(lThreadHandle)
  End If
  If lLib <> 0 Then
    Call FreeLibraryRegister(lLib)
  End If
  Exit Function
ErrorHandler:
  lRet = Err.Number
  GoTo ExitRoutine
End Function
Public Function UnRegister(ByVal p_sFileName As String) As Variant
  Dim lLib As Long
  Dim lProcAddress As Long
  Dim lThreadID As Long
  Dim lSuccess As Long
  Dim lExitCode As Long
  Dim lThreadHandle As Long
  Dim lRet As Long
  On Error GoTo ErrorHandler
  If lRet = NOERRORS Then
    If p_sFileName = "" Then
      lRet = stNoFileProvided
    End If
  End If
  If lRet = NOERRORS Then
    lLib = LoadLibraryRegister(p_sFileName)
    If lLib = 0 Then
      lRet = stFileCouldNotBeLoadedIntoMemorySpace
    End If
  End If
  If lRet = NOERRORS Then
    lProcAddress = GetProcAddressRegister(lLib, "DllUnregisterServer")
    If lProcAddress = 0 Then
      lRet = stNotAValidActiveXComponent
    Else
      lThreadHandle = CreateThreadForRegister(0, 0, lProcAddress, 0, 0, lThreadID)
      If lThreadHandle <> 0 Then
         lSuccess = (WaitForSingleObject(lThreadHandle, 10000) = WAIT_OBJECT_0)
         If lSuccess = 0 Then
          Call GetExitCodeThread(lThreadHandle, lExitCode)
          Call ExitThread(lExitCode)
          lRet = stActiveXComponentUnRegistrationFailed
         Else
          lRet = stActiveXComponentUnRegisterSuccessful
         End If
      End If
    End If
  End If
ExitRoutine:
  UnRegister = lRet
  If lThreadHandle <> 0 Then
    Call CloseHandle(lThreadHandle)
  End If
  If lLib <> 0 Then
    Call FreeLibraryRegister(lLib)
  End If
  Exit Function
ErrorHandler:
  lRet = Err.Number
  GoTo ExitRoutine
End Function
```

