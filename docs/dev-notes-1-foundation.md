# VBA-WebView2 Development Notes, Part 1 — Foundation (Preface to Stage 9.2)

A record of how we established a foundation that receives WebView2's COM callbacks using no external COM library at all,
only machine-code thunks, SAFEARRAY-based memory primitives and a hand-built vtable.
**This volume is frozen. Do not rewrite it.**

---

## Preface — Creating the WebView2Environment

``` VBA
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long
```
### Preparation needed to run this function

### 1. browserExecutableFolder
    0 is fine.
### 2. userDataFolder: specifying the folder path
``` VBA
    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"
```
Pass this userDataPath with `StrPtr(userDataPath)`.
### 3. additionalBrowserArguments
    0 is fine.
### 4. ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler: specifying the handler object
`CreateCoreWebView2EnvironmentWithOptions` launches WebView2.exe.

When startup succeeds, WebView2.exe calls `Invoke`, the 4th function of the object passed as an argument.<br>
To receive that call correctly, we must prepare within VBA a COM object `ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler` that conforms to the specification.<br>
This object is supposed to have the following functions.
```
IUnknown
  0.QueryInterface
  1.AddRef
  2.Release
ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
  3.Invoke
```

## Stage 1 — A scheme in which the corresponding functions are defined in a standard module and an array of their function pointers is laid out in memory as a virtual function table with `CopyMemory`
``` VBA
''''''''''''''''''''''''''''''''''
' --- Module1.bas Stage 1 ---
''''''''''''''''''''''''''''''''''
Option Explicit

' WebView2Loader.dll entry point (Called from c0.CreateWebView2Environment)
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As LongPtr)
Private Const GPTR As Long = &H40 ' Flag for GlobalAlloc uFlags: Allocates fixed and zero-initialized memory
Private Const S_OK As Long = 0

Private vTable(3) As LongPtr
Private pVTable As LongPtr
Private pObject As LongPtr

' Helper function to receive AddressOf as LongPtr
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

' IUnknown::QueryInterface
Private Function Handler_QueryInterface(ByVal this As LongPtr, ByVal riid As LongPtr, ByRef ppvObject As LongPtr) As Long
    ' Normally used to check GUID, but for now it returns itself
    ppvObject = this
    Handler_QueryInterface = S_OK
End Function

' IUnknown::AddRef / Release (Returns 1 as a stub/dummy)
Private Function Handler_AddRef(ByVal this As LongPtr) As Long:
    Handler_AddRef = 1
End Function
Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = 1
End Function

' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
' Receives the initialization result from WebView2
Private Function Handler_Invoke(ByVal this As LongPtr, ByVal errorCode As Long, ByVal pEnvironment As LongPtr) As Long
    ' Does nothing yet
    Debug.Print "ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke is called."
    Handler_Invoke = 0
End Function

' Initializes the WebView2 Environment
Private Sub CreateWebView2Environment()

    ' 1. Store function pointers in the vTable array
    vTable(0) = GetAddr(AddressOf Handler_QueryInterface)
    vTable(1) = GetAddr(AddressOf Handler_AddRef)
    vTable(2) = GetAddr(AddressOf Handler_Release)
    vTable(3) = GetAddr(AddressOf Handler_Invoke)

    ' 2. Allocate heap memory for the vTable
    pVTable = GlobalAlloc(GPTR, (UBound(vTable) + 1) * LenB(vTable(0)))
    CopyMemory ByVal pVTable, vTable(0), (UBound(vTable) + 1) * LenB(vTable(0))

    ' 3. Allocate heap memory for the Object instance itself (points to pVTable)
    pObject = GlobalAlloc(GPTR, LenB(pVTable))
    CopyMemory ByVal pObject, pVTable, LenB(pVTable)

    ' 4. Start creating the WebView2 Environment
    
    ' Prepare a writable user data folder path (Null-terminated Unicode string)
    Dim userDataPath As String
    userDataPath = "C:\Temp\VBA_WebView2" ' Folder will be created automatically if missing

    ' Ensure base directory exists
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    ' Call API (pass the string pointer to the 2nd argument, and custom object to the 4th)
    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, pObject)

    If hr <> S_OK Then
        MsgBox "Failed to initialize WebView2 Environment. HRESULT: " & hr & vbCrLf & _
               "Ensure WebView2Loader.dll is in the same directory as this workbook.", vbCritical, "WebView2 Initialization Error"
    End If

    DoEvents

End Sub
```

## Stage 2 — Staying in a standard module, abolish CopyMemory and build the virtual function table with memory operations via PointerAccesser
``` VBA
''''''''''''''''''''''''''''''''''
' --- Module2.bas Stage 2 ---
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

' https://github.com/WNKLER/RefTypes/discussions/3

''

Option Explicit

#If x64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    Private Const NullPtr As Long = 0&
    Private Const PtrSize = 4
#End If

Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

Private Const S_OK As Long = 0

' WebView2Loader.dll entry point (Called from c0.CreateWebView2Environment)
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

' --- Variable holding the vTable ---
' Thanks to PointerAccessor, these themselves serve as the in-memory body of the COM object
Private Type WebView2Handler
    pVTable As LongPtr     ' Start of the object (pointer to the vTable)
    Functions(3) As LongPtr ' Body of the vTable (array of function pointers)
End Type
Private m_Handler As WebView2Handler

' Helper function to receive AddressOf as LongPtr
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

' IUnknown::QueryInterface
Private Function Handler_QueryInterface(ByVal this As LongPtr, ByVal riid As LongPtr, ByRef ppvObject As LongPtr) As Long
    ' Normally used to check GUID, but for now it returns itself
    ppvObject = this
    Handler_QueryInterface = S_OK
End Function

' IUnknown::AddRef / Release (Returns 1 as a stub/dummy)
Private Function Handler_AddRef(ByVal this As LongPtr) As Long:
    Handler_AddRef = 1
End Function
Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = 1
End Function

' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
' Receives the initialization result from WebView2
Private Function Handler_Invoke(ByVal this As LongPtr, ByVal errorCode As Long, ByVal pEnvironment As LongPtr) As Long
    ' Does nothing yet
    Debug.Print "ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke is called."
    Handler_Invoke = 0
End Function

' Initializes the WebView2 Environment
Private Sub CreateWebView2Environment()

    Static pa As PointerAccessor
    
    ' 1. Set up PointerAccessor (bind pa.arr to pa.sa)
    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = PtrSize
        pa.sa.cLocks = 1
        MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
    End If

    ' 2. Build the vTable inside m_Handler
    ' Place the vTable itself in m_Handler.Functions, and point pVTable at its own Functions array
    m_Handler.pVTable = VarPtr(m_Handler.Functions(0))
    
    ' Use pa.arr to write the contents of the vTable (function pointers)
    ' Point pa.sa.pvData at the write destination and pour the values in through arr(n)
    pa.sa.pvData = VarPtr(m_Handler.Functions(0))
    pa.sa.rgsabound0.cElements = 4
    
    pa.arr(0) = GetAddr(AddressOf Handler_QueryInterface)
    pa.arr(1) = GetAddr(AddressOf Handler_AddRef)
    pa.arr(2) = GetAddr(AddressOf Handler_Release)
    pa.arr(3) = GetAddr(AddressOf Handler_Invoke)
    
    ' 3. Cleanup (safely detach the array)
    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = 0

    ' 4. Call to create the WebView2 environment
    ' Pass the address of m_Handler itself as pObject (its first element is the pointer to the vTable)
    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, VarPtr(m_Handler))

    If hr <> S_OK Then
        MsgBox "Failed to initialize WebView2. HRESULT: " & hr, vbCritical
    End If

    DoEvents

End Sub

Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

'https://github.com/WNKLER/RefTypes/discussions/3#discussion-8595790
Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub
```

## Stage 3 — Same as stage 2, relaying from the standard module's Invoke to the class module's Invoke
``` VBA
''''''''''''''''''''''''''''''''''
' --- Module3.bas Stage 3 ---
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

' https://github.com/WNKLER/RefTypes/discussions/3

''

Option Explicit

#If x64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    Private Const NullPtr As Long = 0&
    Private Const PtrSize = 4
#End If

Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

Private Const S_OK As Long = 0

Private col_Env As New Collection

' WebView2Loader.dll entry point (Called from c0.CreateWebView2Environment)
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

' --- Variable holding the vTable ---
' Thanks to PointerAccessor, these themselves serve as the in-memory body of the COM object
Private Type WebView2Handler
    pVTable As LongPtr     ' Start of the object (pointer to the vTable)
    Functions(3) As LongPtr ' Body of the vTable (array of function pointers)
End Type
Private m_Handler As WebView2Handler

' Helper function to receive AddressOf as LongPtr
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

' IUnknown::QueryInterface
Private Function Handler_QueryInterface(ByVal this As LongPtr, ByVal riid As LongPtr, ByRef ppvObject As LongPtr) As Long
    ' Normally used to check GUID, but for now it returns itself
    ppvObject = this
    Handler_QueryInterface = S_OK
End Function

' IUnknown::AddRef / Release (Returns 1 as a stub/dummy)
Private Function Handler_AddRef(ByVal this As LongPtr) As Long:
    Handler_AddRef = 1
End Function
Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = 1
End Function

' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
' Receives the initialization result from WebView2
Private Function Handler_Invoke(ByVal this As LongPtr, ByVal errorCode As Long, ByVal pEnvironment As LongPtr) As Long
    ' Does nothing yet
    Debug.Print "ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke is called."
    Handler_Invoke = 0
    Dim c1 As Class1
    Set c1 = New Class1
    col_Env.Add c1, CStr(this)
    Call c1.Handler_Invoke(this, errorCode, pEnvironment)
End Function

' Initializes the WebView2 Environment
Private Sub CreateWebView2Environment()

    Static pa As PointerAccessor
    
    ' 1. Set up PointerAccessor (bind pa.arr to pa.sa)
    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = PtrSize
        pa.sa.cLocks = 1
        MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
    End If

    ' 2. Build the vTable inside m_Handler
    ' Place the vTable itself in m_Handler.Functions, and point pVTable at its own Functions array
    m_Handler.pVTable = VarPtr(m_Handler.Functions(0))
    
    ' Use pa.arr to write the contents of the vTable (function pointers)
    ' Point pa.sa.pvData at the write destination and pour the values in through arr(n)
    pa.sa.pvData = VarPtr(m_Handler.Functions(0))
    pa.sa.rgsabound0.cElements = 4
    
    pa.arr(0) = GetAddr(AddressOf Handler_QueryInterface)
    pa.arr(1) = GetAddr(AddressOf Handler_AddRef)
    pa.arr(2) = GetAddr(AddressOf Handler_Release)
    pa.arr(3) = GetAddr(AddressOf Handler_Invoke)
    
    ' 3. Cleanup (safely detach the array)
    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = 0

    ' 4. Call to create the WebView2 environment
    ' Pass the address of m_Handler itself as pObject (its first element is the pointer to the vTable)
    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, VarPtr(m_Handler))

    If hr <> S_OK Then
        MsgBox "Failed to initialize WebView2. HRESULT: " & hr, vbCritical
    End If

    DoEvents

End Sub

Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

'https://github.com/WNKLER/RefTypes/discussions/3#discussion-8595790
Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub

''''''''''''''''''''''''''''''''''
' --- Class1.cls Stage 3 ---
''''''''''''''''''''''''''''''''''
Option Explicit

' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
' Receives the initialization result from WebView2
Public Function Handler_Invoke(ByVal this As LongPtr, ByVal errorCode As Long, ByVal pEnvironment As LongPtr) As Long
    ' Does nothing yet
    Debug.Print "ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke is called."
    Debug.Print "errorCode:", errorCode
    Debug.Print "pEnvironment:", pEnvironment
    Handler_Invoke = 0
End Function
```

## Stage 4 — Abolish the Invoke relay in the standard module and have WebView2 call the class's Invoke directly
 
In stage 3, the standard module's `Handler_Invoke` was called first, and from inside it `Class1.Handler_Invoke` was relayed to.<br>
In stage 4 we abolish the relay completely, so that `Class1.Handler_Invoke` is called directly from WebView2Loader.dll.
 
### Policy
- Building the vTable itself follows the same `PointerAccessor` scheme as stage 3.
- Of the vTable's 4 slots, `[0]QueryInterface` / `[1]AddRef` / `[2]Release` simply receive the `AddressOf` of the standard module.
- Only the `[3]Invoke` slot receives, instead of an `AddressOf`, the address of a **thunk** on an executable region allocated with `VirtualAlloc(PAGE_EXECUTE_READWRITE)`.
- The thunk is 48 bytes of x64 machine code. When COM calls it with the 3 arguments `(this, errorCode, pEnvironment)`, it shifts the arguments one place to the right, wedges `ObjPtr(gHandler)` into the `Me` slot, and `jmp`s to the real address of `Class1.Handler_Invoke`.
- `CopyMemory` is used only for writing into this thunk region (it is a single 48-byte write into the VirtualAlloc'ed executable region, so the impact on Defender countermeasures is acceptable).
### A mechanism to obtain the real address of Class1.Handler_Invoke without ITypeInfo traversal
By placing **only one** `Public` member, `Handler_Invoke`, in `Class1`, the class's vTable layout is pinned to the following.
```
[0] IUnknown::QueryInterface
[1] IUnknown::AddRef
[2] IUnknown::Release
[3] IDispatch::GetTypeInfoCount
[4] IDispatch::GetTypeInfo
[5] IDispatch::GetIDsOfNames
[6] IDispatch::Invoke
[7] Handler_Invoke   ← pinned at offset 56 (= 7 * 8)
```
So the real address can be obtained with `*(LongPtr*)(*(LongPtr*)ObjPtr(gHandler) + 56)`.<br>
Building a function pointer cache with `ITypeInfo::GetTypeAttr` / `GetFuncDesc` / `GetDocumentation` (the equivalent of `BuildFuncPtrCache`), which the standard thunk method required, can be omitted entirely.
 
### An important pitfall: declare Class1.Handler_Invoke as a `Sub`, not a `Function`
When we first implemented it as `Public Function Handler_Invoke(...) As Long`, the body of Invoke was called correctly and reached the `Debug.Print`, but it crashed the moment the function returned.<br>
Cause: a VBA `Function` is internally compiled as **a form with an out argument (`Long*`) for writing back the return value implicitly appended at the end**. COM calls it with 3 arguments by `__stdcall`, yet the VBA class method tries to receive 4 arguments (the last being the pointer for writing back the return value). It then reads a "return value pointer" from an undefined area of the stack at the start of the method, writes `Handler_Invoke = 0` to that invalid address at the end of the method, tramples the return address and saved registers, and crashes at `ret`.<br>
Changing it to `Sub` removes the implicit out argument, so it matches COM's `(this, errorCode, pEnvironment)` signature exactly.<br>
The WebView2 side does not look at the return value (HRESULT) of `ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke`, so making it a `Sub` does no real harm. If you want to return `S_OK` strictly, change the end of the thunk to a `call → xor eax,eax → ret` structure.
 
### Contents of the thunk (x64 machine code)
Win64 calling convention: `RCX=arg0, RDX=arg1, R8=arg2, R9=arg3`
```
Registers when called from COM:
    RCX=this, RDX=errorCode, R8=pEnvironment
   ↓
The thunk shifts the arguments one place to the right while injecting ObjPtr(gHandler) into the Me slot:
    mov r9, r8       ; pEnvironment → arg3
    mov r8, rdx      ; errorCode    → arg2
    mov rdx, rcx     ; this         → arg1
    mov rcx, imm64   ; ObjPtr(gHandler) → arg0 (Me slot)
    mov rax, imm64   ; real address of Class1.Handler_Invoke
    jmp rax
   ↓
Received on the class side:
    Class1.Handler_Invoke(Me=ObjPtr(gHandler), this, errorCode, pEnvironment)
```
No register saving is needed because the destinations being overwritten (`r9`, `r8`, `rdx`) are to the right of the sources (`r8`, `rdx`, `rcx`), so moving them in order from the back is enough; the dependencies run in one direction and do not collide.
 
``` VBA
''''''''''''''''''''''''''''''''''
' --- Module4.bas Stage 4 (CopyMemory version) ---
''''''''''''''''''''''''''''''''''
 
''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/RefTypes/discussions/3
''
 
Option Explicit
 
#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "This module requires 64-bit VBA (x64)"
#End If
 
' --- Structures for SafeArray / PointerAccessor ---
Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum
 
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
 
Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type
 
Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type
 
' --- WebView2Loader.dll ---
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long
 
' --- Win32 API (for the thunk) ---
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long) As LongPtr
 
Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long
 
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, source As Any, ByVal Length As LongPtr)
 
Private Const MEM_COMMIT             As Long = &H1000
Private Const MEM_RESERVE            As Long = &H2000
Private Const MEM_RELEASE            As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const S_OK                   As Long = 0
 
' --- WebView2 Handler structure ---
' Thanks to PointerAccessor, these themselves serve as the in-memory body of the COM object
Private Type WebView2Handler
    pVTable      As LongPtr     ' Start of the object (pointer to the vTable)
    Functions(3) As LongPtr     ' Body of the vTable (4 function pointers)
End Type
Private m_Handler As WebView2Handler
 
' Class instance that receives Invoke. Must be kept alive until the callback completes
Private gHandler As Class1
 
' Executable thunk region (VirtualAlloc)
Private m_pInvokeThunk As LongPtr
 
 
' ============================================================
' Public entry
' ============================================================
Public Sub CreateWebView2Environment()
 
    ' 1. Create the class instance
    Set gHandler = New Class1
 
    ' 2. Get the real address of Class1.Handler_Invoke from vTable slot 7
    Dim pHandlerInvoke As LongPtr
    pHandlerInvoke = GetClassMethodAddrAtFixedSlot(gHandler, 7)
    If pHandlerInvoke = 0 Then
        MsgBox "Failed to obtain the address of Class1.Handler_Invoke", vbCritical
        Exit Sub
    End If
 
    ' 3. Build the thunk for Invoke
    m_pInvokeThunk = BuildInvokeThunk(ObjPtr(gHandler), pHandlerInvoke)
    If m_pInvokeThunk = 0 Then
        MsgBox "Failed to build the Invoke thunk", vbCritical
        Exit Sub
    End If
 
    ' 4. Build the vTable using PointerAccessor
    Static pa As PointerAccessor
    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = PtrSize
        pa.sa.cLocks = 1
        MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
    End If
 
    m_Handler.pVTable = VarPtr(m_Handler.Functions(0))
 
    pa.sa.pvData = VarPtr(m_Handler.Functions(0))
    pa.sa.rgsabound0.cElements = 4
 
    pa.arr(0) = GetAddr(AddressOf Handler_QueryInterface)
    pa.arr(1) = GetAddr(AddressOf Handler_AddRef)
    pa.arr(2) = GetAddr(AddressOf Handler_Release)
    pa.arr(3) = m_pInvokeThunk                           ' ← address of the thunk
 
    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = 0
 
    ' 5. Create the WebView2 environment
    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"
 
    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, VarPtr(m_Handler))
 
    If hr <> S_OK Then
        MsgBox "Failed to initialize WebView2. HRESULT: &H" & Hex(hr), vbCritical
    End If
 
    DoEvents
End Sub
 
Public Sub FinalizeWebView2Environment()
    If m_pInvokeThunk <> 0 Then
        VirtualFree m_pInvokeThunk, 0, MEM_RELEASE
        m_pInvokeThunk = 0
    End If
    Set gHandler = Nothing
End Sub
 
 
' ============================================================
' Build the Invoke thunk (CopyMemory scheme)
' ============================================================
Private Function BuildInvokeThunk( _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr) As LongPtr
 
    Dim addr As LongPtr
    addr = VirtualAlloc(0, 48, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If addr = 0 Then Exit Function
 
    Dim b(47) As Byte, i As Long
    i = 0
 
    ' mov r9, r8       (pEnvironment → arg3)
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx      (errorCode    → arg2)
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx     (this         → arg1)
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    ' mov rcx, imm64   (ObjPtr(gHandler) → arg0 = Me slot)
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    CopyMemory b(i), pSelfObj, 8: i = i + 8
    ' mov rax, imm64   (real address of Class1.Handler_Invoke)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    CopyMemory b(i), pTargetFunc, 8: i = i + 8
    ' jmp rax
    b(i) = &HFF: b(i + 1) = &HE0: i = i + 2
 
    ' Copy the byte sequence into the executable region
    CopyMemory ByVal addr, b(0), CLngPtr(i)
 
    BuildInvokeThunk = addr
End Function
 
 
' ============================================================
' Get the function address of a fixed slot from the class's vTable
'   Class1.Handler_Invoke is the only Public method, so it is settled at slot 7
' ============================================================
Private Function GetClassMethodAddrAtFixedSlot( _
    ByVal cls As Object, ByVal slotIndex As Long) As LongPtr
 
    If cls Is Nothing Then Exit Function
    Dim pObj As LongPtr: pObj = ObjPtr(cls)
    If pObj = 0 Then Exit Function
 
    Dim pVTable As LongPtr
    pVTable = ReadLongPtr(pObj)
    If pVTable = 0 Then Exit Function
 
    GetClassMethodAddrAtFixedSlot = ReadLongPtr(pVTable + slotIndex * PtrSize)
End Function
 
 
' ============================================================
' IUnknown stubs (standard module)
'   Only Invoke was replaced with the thunk, so these still go through AddressOf
' ============================================================
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function
 
Private Function Handler_QueryInterface( _
    ByVal this As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long
    ppvObject = this
    Handler_QueryInterface = S_OK
End Function
 
Private Function Handler_AddRef(ByVal this As LongPtr) As Long
    Handler_AddRef = 1
End Function
 
Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = 1
End Function
 
 
' ============================================================
' Memory-operation primitives via PointerAccessor
' ============================================================
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property
 
Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function
 
Private Sub WritePtrNatively(ByRef ptrs() As LongPtr, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub
 
''''''''''''''''''''''''''''''''''
' --- Class1.cls Stage 4 ---
''''''''''''''''''''''''''''''''''
' The only Public method is Handler_Invoke. This pins Class1's
' vTable layout, guaranteeing that Handler_Invoke lands at slot 7 (= offset 56).
' Never add any other Public method.
'
' ★ Make it a Sub, not a Function.
'   A Function is internally compiled with an out argument for writing back
'   the return value at the end, so it does not match COM's 3-argument __stdcall signature,
'   and corrupts the stack when the method returns, causing a crash. A Sub has no out argument, so it matches.
''''''''''''''''''''''''''''''''''
Option Explicit
 
' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
Public Sub Handler_Invoke( _
    ByVal this As LongPtr, _
    ByVal errorCode As Long, _
    ByVal pEnvironment As LongPtr)
 
    Debug.Print "Class1.Handler_Invoke is called directly from WebView2."
    Debug.Print "  this        :", this
    Debug.Print "  errorCode   :", errorCode
    Debug.Print "  pEnvironment:", pEnvironment
 
End Sub
```
 
### What CopyMemory is still used for after stage 4
Only one place: the 48-byte thunk write inside `BuildInvokeThunk`.<br>
If in the future we replace it with a scheme that writes byte by byte into the thunk region using PointerAccessor (a Byte SafeArray with `pa.sa.cbElements = 1`, pointing pvData at the thunk region), CopyMemory can be abolished completely.


## Stage 5 — Modify the stage 4 thunk so that the class's Invoke can be called as a `Function`
 
In stage 4 we had no choice but to make `Class1.Handler_Invoke` a `Sub`. The cause was that a VBA `Function` is internally compiled with an out argument (`Long*`) for writing back the return value implicitly appended at the end, so with the 3-argument call the thunk prepared, one argument was missing and the stack was corrupted.<br>
In stage 5, the thunk prepares this out argument and uses `call`, which makes declaring it as a `Function` possible. This lets us return an HRESULT explicitly to the COM side (for example, returning `E_NOTIMPL` with `Handler_Invoke = &H80004001`).
 
### Key points of the design
- The thunk is changed from `jmp` to **calling with `call` and coming back**.
- By the x64 calling convention, the 5th argument (the return value pointer) is placed **on the stack at `[rsp+0x20]`** (`[rsp+0x00..0x1F]` is shadow space for the callee).
- The thunk itself **reserves the storage for the return value in its local stack area** and passes that address as the 5th argument.
- After returning from the method, it reads the return value from that storage, loads it into `eax`, and then `ret`s.
### Stack layout (immediately after `sub rsp, 0x38`)
```
[rsp+0x00..0x1F]  shadow space (32 bytes)         ← area the callee may use freely
[rsp+0x20]        arg4 = pRetVal (return value pointer)  ← the thunk writes the address of [rsp+0x28] here
[rsp+0x28]        retval storage (8 bytes)        ← where the Function writes its return value
[rsp+0x30]        alignment padding (8 bytes)
                  ─────────────────────────
                  total 0x38 = 56 bytes
```
 
### Stack alignment check
On x64, RSP immediately before a `call` instruction must be 16-byte aligned (the spec is that, because the call pushes an 8-byte return address, RSP at the callee's entry is `16N + 8`).
- Right after WebView2 `call`s the thunk: RSP is `16N + 8`
- After the thunk does `sub rsp, 0x38` (56): RSP is `16N + 8 - 56 = 16N - 48 = 16(N-3)` → **16-byte aligned ✓**
- Right after the internal `call rax` (callee entry): of the form `16(N-3) - 8 = 16(N-3) + 8` → **the canonical `16N+8` ✓**
### Machine code (54 bytes)
```
48 83 EC 38              sub  rsp, 0x38          ; reserve local area
4D 89 C1                 mov  r9, r8             ; pEnvironment → arg3
49 89 D0                 mov  r8, rdx            ; errorCode    → arg2
48 89 CA                 mov  rdx, rcx           ; this         → arg1
48 B9 <imm64>            mov  rcx, ObjPtr(gHandler) ; inject Me     → arg0
48 8D 44 24 28           lea  rax, [rsp+0x28]    ; address of the return value storage
48 89 44 24 20           mov  [rsp+0x20], rax    ; arg4 = return value pointer
48 B8 <imm64>            mov  rax, pTargetFunc   ; Class.Handler_Invoke
FF D0                    call rax
8B 44 24 28              mov  eax, [rsp+0x28]    ; load the return value
48 83 C4 38              add  rsp, 0x38          ; release local area
C3                       ret
```
 
No register saving is needed because the destinations being overwritten (`r9`/`r8`/`rdx`) are to the right of the sources (`r8`/`rdx`/`rcx`), so moving them in order from the back keeps the dependencies one-directional with no collision (the same reason as stage 4).
 
### Changes from stage 4
- Standard module side: **only the `BuildInvokeThunk` function is completely reworked**. Everything else (vTable construction, PointerAccessor, IUnknown stubs, entry point, etc.) is exactly the same.
- Class module side: `Sub Handler_Invoke` is changed back to `Function Handler_Invoke ... As Long`, and `Handler_Invoke = 0` (S_OK) is added at the end.
``` VBA
''''''''''''''''''''''''''''''''''
' --- Module5.bas Stage 5 (excerpt of BuildInvokeThunk only. The rest is the same as stage 4) ---
''''''''''''''''''''''''''''''''''
 
' ============================================================
' Build the Invoke thunk (CopyMemory scheme, Function support)
'
'   Call conversion:
'     COM side  Invoke(this, errorCode, pEnvironment) → eax = HRESULT
'                ↓
'     class     Handler_Invoke(Me, this, errorCode, pEnvironment, pRetVal*)
'             pRetVal* is the out argument for writing back the Function's return value
' ============================================================
Private Function BuildInvokeThunk( _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr) As LongPtr
 
    Dim addr As LongPtr
    addr = VirtualAlloc(0, 64, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If addr = 0 Then Exit Function
 
    Dim b(63) As Byte, i As Long
    i = 0
 
    ' sub rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4
 
    ' mov r9, r8       (pEnvironment → arg3)
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx      (errorCode    → arg2)
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx     (this         → arg1)
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    ' mov rcx, imm64   (ObjPtr(gHandler) → arg0 = Me slot)
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    CopyMemory b(i), pSelfObj, 8: i = i + 8
 
    ' lea rax, [rsp+0x28]   (address of the return value storage)
    b(i) = &H48: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5
    ' mov [rsp+0x20], rax   (arg4 = retval pointer, placed in the stack argument slot)
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H20: i = i + 5
 
    ' mov rax, imm64   (real address of Class.Handler_Invoke)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    CopyMemory b(i), pTargetFunc, 8: i = i + 8
    ' call rax
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
 
    ' mov eax, [rsp+0x28]   (load the return value the Function wrote)
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H28: i = i + 4
 
    ' add rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4
 
    ' ret
    b(i) = &HC3: i = i + 1
 
    CopyMemory ByVal addr, b(0), CLngPtr(i)
 
    BuildInvokeThunk = addr
End Function
 
''''''''''''''''''''''''''''''''''
' --- Class2.cls Stage 5 ---
''''''''''''''''''''''''''''''''''
' The only Public method is Handler_Invoke. The premise that it is pinned at vTable slot 7 is
' the same as in stage 4. The difference is that it is a Function As Long, not a Sub.
''''''''''''''''''''''''''''''''''
Option Explicit
 
' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
Public Function Handler_Invoke( _
    ByVal this As LongPtr, _
    ByVal errorCode As Long, _
    ByVal pEnvironment As LongPtr) As Long
 
    Debug.Print "Class2.Handler_Invoke is called directly from WebView2."
    Debug.Print "  this        :", this
    Debug.Print "  errorCode   :", errorCode
    Debug.Print "  pEnvironment:", pEnvironment
 
    Handler_Invoke = 0   ' S_OK
End Function
```

## Stage 6 — Change the thunk writes from CopyMemory to the PointerAccesser scheme

Standard module Module6.bas
``` VBA
''''''''''''''''''''''''''''''''''
' --- Module6.bas Stage 6 (CopyMemory completely abolished) ---
'   Remove CopyMemory from the stage 5 thunk construction and
'   unify on PointerAccessor (MemLongPtr).
'
'   Extend the thunk's machine code from 54 → 56 bytes (pad 2 bytes of int3
'   at the end, a safe area never executed because it is after the ret). This
'   rounds it to a multiple of 8, so just 7 MemLongPtr writes complete the transfer to the executable region.
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/RefTypes/discussions/3
''

Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "This module requires 64-bit VBA (x64)"
#End If

' --- Structures for SafeArray / PointerAccessor ---
Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

' --- WebView2Loader.dll ---
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

' --- Win32 API (for the thunk, CopyMemory not used) ---
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long) As LongPtr

Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long

Private Const MEM_COMMIT             As Long = &H1000
Private Const MEM_RESERVE            As Long = &H2000
Private Const MEM_RELEASE            As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const S_OK                   As Long = 0

' --- WebView2 Handler structure ---
Private Type WebView2Handler
    pVTable      As LongPtr
    Functions(3) As LongPtr
End Type
Private m_Handler As WebView2Handler

Private gHandler As Class3
Private m_pInvokeThunk As LongPtr


' ============================================================
' Public entry
' ============================================================
Public Sub CreateWebView2Environment()

    Set gHandler = New Class3

    Dim pHandlerInvoke As LongPtr
    pHandlerInvoke = GetClassMethodAddrAtFixedSlot(gHandler, 7)
    If pHandlerInvoke = 0 Then
        MsgBox "Failed to obtain the address of Class3.Handler_Invoke", vbCritical
        Exit Sub
    End If

    m_pInvokeThunk = BuildInvokeThunk(ObjPtr(gHandler), pHandlerInvoke)
    If m_pInvokeThunk = 0 Then
        MsgBox "Failed to build the Invoke thunk", vbCritical
        Exit Sub
    End If

    Static pa As PointerAccessor
    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = PtrSize
        pa.sa.cLocks = 1
        MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
    End If

    m_Handler.pVTable = VarPtr(m_Handler.Functions(0))

    pa.sa.pvData = VarPtr(m_Handler.Functions(0))
    pa.sa.rgsabound0.cElements = 4

    pa.arr(0) = GetAddr(AddressOf Handler_QueryInterface)
    pa.arr(1) = GetAddr(AddressOf Handler_AddRef)
    pa.arr(2) = GetAddr(AddressOf Handler_Release)
    pa.arr(3) = m_pInvokeThunk

    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = 0

    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, VarPtr(m_Handler))

    If hr <> S_OK Then
        MsgBox "Failed to initialize WebView2. HRESULT: &H" & Hex(hr), vbCritical
    End If

    DoEvents
End Sub

Public Sub FinalizeWebView2Environment()
    If m_pInvokeThunk <> 0 Then
        VirtualFree m_pInvokeThunk, 0, MEM_RELEASE
        m_pInvokeThunk = 0
    End If
    Set gHandler = Nothing
End Sub


' ============================================================
' Build the Invoke thunk (PointerAccessor scheme, CopyMemory completely abolished)
'
'   The machine code is identical to stage 5, but 2 bytes of int3 (0xCC) are
'   padded at the end to make a total of 56 bytes (= 8 x 7). This means
'   just 7 MemLongPtr writes complete the transfer to the executable region.
'   int3 is after the ret so it is not reached in normal execution (a safety measure:
'   if execution ever falls through, the debugger will catch it).
'
'   Machine code (56 bytes):
'     48 83 EC 38              sub  rsp, 0x38
'     4D 89 C1                 mov  r9, r8
'     49 89 D0                 mov  r8, rdx
'     48 89 CA                 mov  rdx, rcx
'     48 B9 <imm64>            mov  rcx, ObjPtr(gHandler)
'     48 8D 44 24 28           lea  rax, [rsp+0x28]
'     48 89 44 24 20           mov  [rsp+0x20], rax
'     48 B8 <imm64>            mov  rax, pTargetFunc
'     FF D0                    call rax
'     8B 44 24 28              mov  eax, [rsp+0x28]
'     48 83 C4 38              add  rsp, 0x38
'     C3                       ret
'     CC CC                    int3 / int3        ← padding (not executed)
' ============================================================
Private Function BuildInvokeThunk( _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr) As LongPtr

    Const THUNK_SIZE As Long = 56  ' = 8 × 7
    Dim addr As LongPtr
    addr = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If addr = 0 Then Exit Function

    Dim b(0 To THUNK_SIZE - 1) As Byte, i As Long
    i = 0

    ' sub rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4

    ' mov r9, r8
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3

    ' mov rcx, imm64   (ObjPtr(gHandler))
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pSelfObj: i = i + 8        ' ← replaced from CopyMemory

    ' lea rax, [rsp+0x28]
    b(i) = &H48: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5
    ' mov [rsp+0x20], rax
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H20: i = i + 5

    ' mov rax, imm64   (pTargetFunc)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pTargetFunc: i = i + 8     ' ← replaced from CopyMemory

    ' call rax
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2

    ' mov eax, [rsp+0x28]
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H28: i = i + 4

    ' add rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4

    ' ret
    b(i) = &HC3: i = i + 1

    ' --- That is 54 bytes so far; pad the remaining 2 bytes ---
    ' int3 (debugger break. Not reached in normal execution)
    b(i) = &HCC: i = i + 1
    b(i) = &HCC: i = i + 1

    ' --- Transfer to the executable region (no CopyMemory; MemLongPtr in 8-byte units, 7 times) ---
    Dim k As Long
    For k = 0 To THUNK_SIZE - 1 Step 8
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k

    BuildInvokeThunk = addr
End Function


' ============================================================
' Take 8 bytes (LongPtr) from the given offset of a byte array
'
'   Read a LongPtr directly from the address of b(offset) using PointerAccessor.
'   A small helper that reinterprets a byte array.
' ============================================================
Private Function ReadLongPtrFromBytes(ByRef b() As Byte, ByVal offset As Long) As LongPtr
    ReadLongPtrFromBytes = ReadLongPtr(VarPtr(b(offset)))
End Function


' ============================================================
' Get the function address of a fixed slot from the class's vTable
' ============================================================
Private Function GetClassMethodAddrAtFixedSlot( _
    ByVal cls As Object, ByVal slotIndex As Long) As LongPtr

    If cls Is Nothing Then Exit Function
    Dim pObj As LongPtr: pObj = ObjPtr(cls)
    If pObj = 0 Then Exit Function

    Dim pVTable As LongPtr
    pVTable = ReadLongPtr(pObj)
    If pVTable = 0 Then Exit Function

    GetClassMethodAddrAtFixedSlot = ReadLongPtr(pVTable + slotIndex * PtrSize)
End Function


' ============================================================
' IUnknown stubs (standard module)
' ============================================================
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

Private Function Handler_QueryInterface( _
    ByVal this As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long
    ppvObject = this
    Handler_QueryInterface = S_OK
End Function

Private Function Handler_AddRef(ByVal this As LongPtr) As Long
    Handler_AddRef = 1
End Function

Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = 1
End Function


' ============================================================
' Memory-operation primitives via PointerAccessor
' ============================================================
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

' https://github.com/WNKLER/RefTypes/discussions/3#discussion-8595790
Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub
```

Class module Class3.cls
```
''''''''''''''''''''''''''''''''''
' --- Class3.cls Stage 6 ---
'   The only Public method is Handler_Invoke (pinned at vTable slot 7).
'   The contents are equivalent to stage 5's Class2; there is no change on the class side.
'   The point of stage 6 is that CopyMemory was completely
'   abolished from the thunk construction on the standard module side.
''''''''''''''''''''''''''''''''''
Option Explicit

' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
Public Function Handler_Invoke( _
    ByVal this As LongPtr, _
    ByVal errorCode As Long, _
    ByVal pEnvironment As LongPtr) As Long

    Debug.Print "Class3.Handler_Invoke is called directly from WebView2."
    Debug.Print "  this        :", this
    Debug.Print "  errorCode   :", errorCode
    Debug.Print "  pEnvironment:", pEnvironment

    Handler_Invoke = 0   ' S_OK
End Function
```

## Stage 7 — Switch the thunk region allocation from VirtualAlloc to the AddressOf acquisition scheme
 
Up to stage 6 we had completely abolished CopyMemory, but the allocation of the thunk region still depended on `VirtualAlloc(PAGE_EXECUTE_READWRITE)`.<br>
In stage 7, following the technique of `MouseScroll.bas` / `LibTimers.bas`, we write the thunk directly into the **code region of an empty Sub** defined in a standard module, thereby abolishing even the `VirtualAlloc` / `VirtualFree` API declarations.
 
### Mechanism
- The VBA runtime allocates a code region of around 100 bytes for each procedure. Even if it contains only the short code equivalent to `End Sub`, the region itself is reserved generously. According to the comments in `MouseScroll.bas`, there is a track record of being able to use up to 91 bytes (x64).
- Our thunk is 56 bytes, so it fits comfortably.
- Define `Private Sub InvokeThunkSlot(): End Sub` in a standard module, obtain its address with `AddressOf`, and overwrite the machine code with PointerAccessor (`MemLongPtr`).
- It is a code region managed by the VBA runtime, so no release processing (the equivalent of `VirtualFree`) is needed.
### Points to note
- **Countermeasure against optimization**: under the `Compile On Demand` / `Background Compile` settings the empty Sub may not actually be materialized, so it must be **called explicitly once to Force-Compile** before writing the thunk (the same technique as `MouseScroll.bas`). Being an empty Sub, calling it does no harm.
- **Syntax restriction of `AddressOf`**: by language specification, VBA's `AddressOf` operator can only be written at the position of an actual argument of a function/Sub. Writing it directly on the right-hand side of an assignment, as in `addr = AddressOf MySub`, is a syntax error. So it must be wrapped as an argument of some function, as in `addr = VBA.Int(AddressOf MySub)`. `VBA.Int` is used because it is the most harmless function that passes an integer through as an integer (the established idiom of MouseScroll and others). The Variant route that turns it into a LongLong to prevent the upper 32 bits from being dropped on x64 is only a secondary effect.
- **Do not call it directly after writing**: if you call `InvokeThunkSlot` directly after overwriting, the thunk runs on the premise of `ICoreWebView2...` arguments, so the behavior is undefined. Do the Force-Compile before writing the thunk.
### Changes from stage 6
- **Delete all** of `Declare PtrSafe Function VirtualAlloc / VirtualFree` and the related constants (`MEM_COMMIT` / `MEM_RESERVE` / `MEM_RELEASE` / `PAGE_EXECUTE_READWRITE`).
- Change how `BuildInvokeThunk` obtains the thunk region from `VirtualAlloc(...)` → `VBA.Int(AddressOf InvokeThunkSlot)`.
- Call `InvokeThunkSlot` at the start of `BuildInvokeThunk` to Force-Compile.
- Newly add `Private Sub InvokeThunkSlot(): End Sub` (the host of the thunk).
- Remove the `VirtualFree` call from `FinalizeWebView2Environment` (not needed because it is a region managed by the VBA runtime).
- The machine-code generation part (assembling the 56 bytes, padding, and writing with MemLongPtr 7 times) is exactly the same as stage 6.

``` VBA
''''''''''''''''''''''''''''''''''
' --- Module7.bas Stage 7 (VirtualAlloc not used) ---
'   Switch to a scheme that writes the thunk region directly
'   into the code region of an empty Sub defined in a standard module.
'   The VBA runtime allocates a code region of around 100 bytes to each procedure
'   (the MouseScroll comment says 91 bytes for x64),
'   so a 56-byte thunk fits comfortably.
'
'   Obtain the address of that region with AddressOf and overwrite
'   the machine code with PointerAccessor. VirtualAlloc / VirtualFree
'   become completely unnecessary.
'
'   Notes:
'   - The empty Sub must be called explicitly and
'     Force-Compiled so that it is not removed by optimization (the same technique as MouseScroll).
'   - On x64, receiving AddressOf as a plain Long drops the upper 32 bits, so
'     obtain it via Variant→LongLong with VBA.Int(AddressOf ...).
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/RefTypes/discussions/3
''

Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "This module requires 64-bit VBA (x64)"
#End If

' --- Structures for SafeArray / PointerAccessor ---
Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

' --- WebView2Loader.dll ---
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

Private Const S_OK As Long = 0

' --- WebView2 Handler structure ---
Private Type WebView2Handler
    pVTable      As LongPtr
    Functions(3) As LongPtr
End Type
Private m_Handler As WebView2Handler

Private gHandler As Class4
Private m_pInvokeThunk As LongPtr


' ============================================================
' Public entry
' ============================================================
Public Sub CreateWebView2Environment()

    Set gHandler = New Class4

    Dim pHandlerInvoke As LongPtr
    pHandlerInvoke = GetClassMethodAddrAtFixedSlot(gHandler, 7)
    If pHandlerInvoke = 0 Then
        MsgBox "Failed to obtain the address of Class4.Handler_Invoke", vbCritical
        Exit Sub
    End If

    m_pInvokeThunk = BuildInvokeThunk(ObjPtr(gHandler), pHandlerInvoke)
    If m_pInvokeThunk = 0 Then
        MsgBox "Failed to build the Invoke thunk", vbCritical
        Exit Sub
    End If

    Static pa As PointerAccessor
    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = PtrSize
        pa.sa.cLocks = 1
        MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
    End If

    m_Handler.pVTable = VarPtr(m_Handler.Functions(0))

    pa.sa.pvData = VarPtr(m_Handler.Functions(0))
    pa.sa.rgsabound0.cElements = 4

    pa.arr(0) = GetAddr(AddressOf Handler_QueryInterface)
    pa.arr(1) = GetAddr(AddressOf Handler_AddRef)
    pa.arr(2) = GetAddr(AddressOf Handler_Release)
    pa.arr(3) = m_pInvokeThunk

    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = 0

    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, VarPtr(m_Handler))

    If hr <> S_OK Then
        MsgBox "Failed to initialize WebView2. HRESULT: &H" & Hex(hr), vbCritical
    End If

    DoEvents
End Sub

' ============================================================
' Cleanup
'   The thunk region is a code region managed by the VBA runtime,
'   so there is no need to release it here. Just tidy up the class instance.
' ============================================================
Public Sub FinalizeWebView2Environment()
    Set gHandler = Nothing
    m_pInvokeThunk = 0
End Sub


' ============================================================
' Build the Invoke thunk (reusing the code region of an empty Sub)
'
'   The thunk's machine code is identical to stage 6 (56 bytes).
'   The difference is the write destination: instead of an executable region allocated with VirtualAlloc,
'   we write directly into the code region of an empty Sub (InvokeThunkSlot) in the standard module.
'
'   Force-Compile (an explicit call to the empty Sub) must be done before
'   overwriting it with the thunk. If this Sub is called after the overwrite, the thunk
'   runs on the premise of ICoreWebView2... arguments and breaks.
' ============================================================
Private Function BuildInvokeThunk( _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr) As LongPtr

    ' Force-Compile: call the empty Sub once (countermeasure against optimization)
    InvokeThunkSlot

    ' Get the address of the empty Sub's code region
    '   On x64, receiving AddressOf as a Long drops the upper 32 bits, so
    '   extract it as a LongLong via a Variant with VBA.Int(AddressOf ...).
    Dim addr As LongPtr
    addr = VBA.Int(AddressOf InvokeThunkSlot)
    If addr = 0 Then Exit Function

    ' --- Machine code generation (56 bytes, identical to stage 6) ---
    Const THUNK_SIZE As Long = 56  ' = 8 × 7
    Dim b(0 To THUNK_SIZE - 1) As Byte, i As Long
    i = 0

    ' sub rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4

    ' mov r9, r8
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3

    ' mov rcx, imm64   (ObjPtr(gHandler))
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pSelfObj: i = i + 8

    ' lea rax, [rsp+0x28]
    b(i) = &H48: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5
    ' mov [rsp+0x20], rax
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H20: i = i + 5

    ' mov rax, imm64   (pTargetFunc)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pTargetFunc: i = i + 8

    ' call rax
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2

    ' mov eax, [rsp+0x28]
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H28: i = i + 4

    ' add rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4

    ' ret
    b(i) = &HC3: i = i + 1

    ' --- Padding (int3) ---
    b(i) = &HCC: i = i + 1
    b(i) = &HCC: i = i + 1

    ' --- Transfer to the code region of the empty Sub (MemLongPtr, 7 times) ---
    Dim k As Long
    For k = 0 To THUNK_SIZE - 1 Step 8
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k

    BuildInvokeThunk = addr
End Function


' ============================================================
' Empty Sub used as the thunk region
'
'   The VBA runtime allocates a code region of around 100 bytes
'   to each procedure, and even if it contains only the short code equivalent to End Sub,
'   the region itself is reserved generously. A 56-byte thunk fits comfortably
'   (MouseScroll's track record shows that up to 91 bytes are usable).
'
'   By writing machine code directly here, we obtain an
'   executable thunk region without relying on VirtualAlloc.
'
'   ★ Important: after this code has been written, never call this Sub directly.
'   The thunk runs on the premise that it receives ICoreWebView2... arguments,
'   so calling it with no arguments is undefined behavior.
' ============================================================
Private Sub InvokeThunkSlot(): End Sub


' ============================================================
' Take 8 bytes (LongPtr) from the given offset of a byte array
' ============================================================
Private Function ReadLongPtrFromBytes(ByRef b() As Byte, ByVal offset As Long) As LongPtr
    ReadLongPtrFromBytes = ReadLongPtr(VarPtr(b(offset)))
End Function


' ============================================================
' Get the function address of a fixed slot from the class's vTable
' ============================================================
Private Function GetClassMethodAddrAtFixedSlot( _
    ByVal cls As Object, ByVal slotIndex As Long) As LongPtr

    If cls Is Nothing Then Exit Function
    Dim pObj As LongPtr: pObj = ObjPtr(cls)
    If pObj = 0 Then Exit Function

    Dim pVTable As LongPtr
    pVTable = ReadLongPtr(pObj)
    If pVTable = 0 Then Exit Function

    GetClassMethodAddrAtFixedSlot = ReadLongPtr(pVTable + slotIndex * PtrSize)
End Function


' ============================================================
' IUnknown stubs (standard module)
' ============================================================
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

Private Function Handler_QueryInterface( _
    ByVal this As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long
    ppvObject = this
    Handler_QueryInterface = S_OK
End Function

Private Function Handler_AddRef(ByVal this As LongPtr) As Long
    Handler_AddRef = 1
End Function

Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = 1
End Function


' ============================================================
' Memory-operation primitives via PointerAccessor
' ============================================================
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

' https://github.com/WNKLER/RefTypes/discussions/3#discussion-8595790
Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub
```

``` VBA
''''''''''''''''''''''''''''''''''
' --- Class4.cls Stage 7 ---
'   The only Public method is Handler_Invoke (pinned at vTable slot 7).
'   The contents are equivalent to stages 5 and 6.
'   The point of stage 7 is that the thunk region allocation on the standard module side
'   was switched from VirtualAlloc to the
'   AddressOf acquisition scheme (reusing the code region of an empty Sub).
''''''''''''''''''''''''''''''''''
Option Explicit

' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
Public Function Handler_Invoke( _
    ByVal this As LongPtr, _
    ByVal errorCode As Long, _
    ByVal pEnvironment As LongPtr) As Long

    Debug.Print "Class4.Handler_Invoke is called directly from WebView2."
    Debug.Print "  this        :", this
    Debug.Print "  errorCode   :", errorCode
    Debug.Print "  pEnvironment:", pEnvironment

    Handler_Invoke = 0   ' S_OK
End Function
```

## Stage 8 — Introduce a two-step jump structure via an EntryPoint stub and make use of the EBMode check
 
Through stage 7 we abolished CopyMemory / VirtualAlloc and completed the thunk scheme that reuses the code region of an empty Sub.<br>
However, the stage 7 scheme called the DummyASM region (= the thunk) **directly** from the vTable, so if WebView2 calls Invoke while Excel is in Break Mode, the VBA interpreter is hit while stopped and there is a danger of a crash.<br>
In stage 8, following the technique of `LibTimers.bas` / `MouseScroll.bas`, we introduce a **two-step jump structure** of an `EntryPoint` stub and a `DummyASM` stub. This makes it possible to incorporate the **EBMode check** that the VBA runtime provides (a safety device that skips the callback while in Break Mode).
 
### Background knowledge: what the address obtained by AddressOf really is
For an empty Sub such as `Private Sub Foo(): End Sub` written by the user in a standard module, the address obtainable with `AddressOf Foo` is **not "the entry point of the user-written Sub body itself" but "the start address of preprocessing code (a stub) that the VBA runtime attached automatically"**.<br>
This stub part has runtime processing such as the EBMode check and argument adjustment built in; on x64 the EBMode check is around +37 bytes from the start, and at position +55 there is embedded **an 8-byte immediate region holding "the address to call next"**.<br>
Normally the user Sub body (for an empty Sub, effectively the address of an "immediate ret") is written in that immediate, but **if you rewrite it, you can branch to arbitrary code**. This is the core mechanism of LibTimers / MouseScroll.
 
### Two-step jump structure
```
WebView2 calls vTable[3]
        ↓
Start of the EntryPoint stub            ← register this address in vTable[3]
        ↓
VBA runtime prologue at EntryPoint+0..+37 (including the EBMode check)
        ↓
   ┌── if EBMode = 2 (Break Mode), skip the rest → immediate ret
   │
   └── in normal operation, call indirectly to whatever the immediate at EntryPoint+55 points to
        ↓
DummyASM stub region (= our 56-byte thunk)   ← write this address into EntryPoint+55
        ↓
The thunk shifts the arguments + injects ObjPtr
        ↓
Class.Handler_Invoke
```
 
The 8-byte immediate region at `EntryPoint+55`, expressed as x64 instructions, is the 8-byte imm64 part of a `mov rax, imm64` opcode (`48 B8`, a 10-byte instruction) that immediately precedes it (from the code analysis of LibTimers). If rewritten, the jump target changes for the `call rax` or `jmp rax` that follows.
 
### Changes from stage 7
- Increase the empty Subs from the single `InvokeThunkSlot` to **two, `EntryPoint` + `DummyASM`**.
- Change the thunk's write destination from `InvokeThunkSlot` to `DummyASM`.
- Change the address registered in vTable[3] from "the start of the thunk" to "`AddressOf EntryPoint`".
- Add processing that writes the address of `DummyASM` with `MemLongPtr` at the position `EntryPoint + 55` (= `LATE_BIND_OFFSET`).
The thunk's own machine code (56 bytes) is exactly the same as in stages 6 to 7. We had worried about stack alignment, but thanks to the invariant of the `call` instruction (RSP at callee entry is of the form `16N+8`), it turned out to line up fine even via EntryPoint.
 
### The value of LATE_BIND_OFFSET
On x64 it is fixed at **55** (the value from LibTimers / MouseScroll's track record). In theory this position could change with the VBA version or the Office bitness, but in the current VBA7 + x64 environment it works fine with 55.

```
''''''''''''''''''''''''''''''''''
' --- Module8.bas Stage 8 (two-step jump via an EntryPoint stub) ---
'
'   In stage 7, the code region of DummyASM (an empty Sub) was called directly from the vTable,
'   so even during Break Mode, if WebView2 called Invoke
'   the thunk would run. Stage 8 follows the technique of LibTimers / MouseScroll
'   and introduces a two-step jump structure, EntryPoint stub → DummyASM stub,
'   making use of the EBMode check (EntryPoint+37) provided by the VBA runtime.
'
'   Structure:
'     vTable[3] = AddressOf EntryPoint  (start of the VBA stub)
'                       ↓
'                 EBMode check          (skip the rest in Break Mode)
'                       ↓
'                 read the "address to call next" at EntryPoint+55 and call indirectly
'                       ↓                 ← write the address of DummyASM here
'                 DummyASM stub region (= our 56-byte thunk)
'                       ↓
'                 inject ObjPtr(gHandler) into Me and go to Class5.Handler_Invoke
'
'   The thunk's machine code itself is exactly the same as stages 6 to 7.
'   However, the way it is called changes from a "direct call" to an "indirect call via EntryPoint",
'   so the stack alignment might be off. If it does not work, an adjustment such as
'   sub rsp, 0x38 → sub rsp, 0x30 is needed.
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/RefTypes/discussions/3
''

Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "This module requires 64-bit VBA (x64)"
#End If

' --- Structures for SafeArray / PointerAccessor ---
Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

' --- WebView2Loader.dll ---
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

Private Const S_OK As Long = 0

' Position within the EntryPoint stub where the "address to call next" is embedded as an immediate
'   Value based on the investigation of LibTimers / MouseScroll (x64)
Private Const LATE_BIND_OFFSET As Long = 55

' --- WebView2 Handler structure ---
Private Type WebView2Handler
    pVTable      As LongPtr
    Functions(3) As LongPtr
End Type
Private m_Handler As WebView2Handler

Private gHandler As Class5
Private m_pEntryPointAddr As LongPtr   ' = AddressOf EntryPoint (the address to register in the vTable)


' ============================================================
' Public entry
' ============================================================
Public Sub CreateWebView2Environment()

    Set gHandler = New Class5

    Dim pHandlerInvoke As LongPtr
    pHandlerInvoke = GetClassMethodAddrAtFixedSlot(gHandler, 7)
    If pHandlerInvoke = 0 Then
        MsgBox "Failed to obtain the address of Class5.Handler_Invoke", vbCritical
        Exit Sub
    End If

    ' Set up the two-step jump structure of EntryPoint and DummyASM
    m_pEntryPointAddr = SetupTwoStageThunk(ObjPtr(gHandler), pHandlerInvoke)
    If m_pEntryPointAddr = 0 Then
        MsgBox "Failed to set up the two-step jump structure", vbCritical
        Exit Sub
    End If

    ' --- Build the vTable ---
    Static pa As PointerAccessor
    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = PtrSize
        pa.sa.cLocks = 1
        MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
    End If

    m_Handler.pVTable = VarPtr(m_Handler.Functions(0))

    pa.sa.pvData = VarPtr(m_Handler.Functions(0))
    pa.sa.rgsabound0.cElements = 4

    pa.arr(0) = GetAddr(AddressOf Handler_QueryInterface)
    pa.arr(1) = GetAddr(AddressOf Handler_AddRef)
    pa.arr(2) = GetAddr(AddressOf Handler_Release)
    pa.arr(3) = m_pEntryPointAddr        ' ← the EntryPoint stub is the entrance

    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = 0

    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, VarPtr(m_Handler))

    If hr <> S_OK Then
        MsgBox "Failed to initialize WebView2. HRESULT: &H" & Hex(hr), vbCritical
    End If

    DoEvents
End Sub

Public Sub FinalizeWebView2Environment()
    Set gHandler = Nothing
    m_pEntryPointAddr = 0
End Sub


' ============================================================
' Set up the two-step jump structure
'
'   (1) Write the 56-byte thunk machine code into the DummyASM stub region
'   (2) Overwrite the "address to call next" at EntryPoint+55 with the address of DummyASM
'
'   Return value: the address to register in the vTable = AddressOf EntryPoint
' ============================================================
Private Function SetupTwoStageThunk( _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr) As LongPtr

    ' Force-Compile (materialize both stubs)
    EntryPoint
    DummyASM

    Dim pEntry  As LongPtr: pEntry = VBA.Int(AddressOf EntryPoint)
    Dim pDummy  As LongPtr: pDummy = VBA.Int(AddressOf DummyASM)
    If pEntry = 0 Or pDummy = 0 Then Exit Function

    ' (1) Write the thunk into the DummyASM region
    WriteThunkMachineCode pDummy, pSelfObj, pTargetFunc

    ' (2) Replace the indirect jump target address at EntryPoint+55 with DummyASM
    MemLongPtr(pEntry + LATE_BIND_OFFSET) = pDummy

    SetupTwoStageThunk = pEntry
End Function


' ============================================================
' Write the thunk's machine code to the specified address (same as stages 6 to 7)
'
'   Machine code (56 bytes):
'     48 83 EC 38              sub  rsp, 0x38
'     4D 89 C1                 mov  r9, r8
'     49 89 D0                 mov  r8, rdx
'     48 89 CA                 mov  rdx, rcx
'     48 B9 <imm64>            mov  rcx, ObjPtr(gHandler)
'     48 8D 44 24 28           lea  rax, [rsp+0x28]
'     48 89 44 24 20           mov  [rsp+0x20], rax
'     48 B8 <imm64>            mov  rax, pTargetFunc
'     FF D0                    call rax
'     8B 44 24 28              mov  eax, [rsp+0x28]
'     48 83 C4 38              add  rsp, 0x38
'     C3                       ret
'     CC CC                    int3 / int3 (padding)
' ============================================================
Private Sub WriteThunkMachineCode( _
    ByVal addr As LongPtr, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr)

    Const THUNK_SIZE As Long = 56
    Dim b(0 To THUNK_SIZE - 1) As Byte, i As Long
    i = 0

    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3

    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pSelfObj: i = i + 8

    b(i) = &H48: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H20: i = i + 5

    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pTargetFunc: i = i + 8

    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H28: i = i + 4
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4
    b(i) = &HC3: i = i + 1

    b(i) = &HCC: i = i + 1
    b(i) = &HCC: i = i + 1

    Dim k As Long
    For k = 0 To THUNK_SIZE - 1 Step 8
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k
End Sub


' ============================================================
' EntryPoint and DummyASM (empty Subs)
'   EntryPoint:
'     The VBA runtime prepends stub code such as the EBMode check.
'     The EBMode check is around +37 bytes, and at +55 the "address
'     to call next" is held as an immediate. In Break Mode (EBMode=2) the
'     check at +37 skips everything after it, so the callback is
'     safely skipped.
'   DummyASM:
'     Write the 56-byte thunk machine code here. By rewriting EntryPoint+55,
'     it is called indirectly from here after passing the EntryPoint
'     check.
'
'   ★ Important: after writing this code, never call this Sub directly.
' ============================================================
Private Sub EntryPoint(): End Sub
Private Sub DummyASM(): End Sub


' ============================================================
' Take 8 bytes (LongPtr) from the given offset of a byte array
' ============================================================
Private Function ReadLongPtrFromBytes(ByRef b() As Byte, ByVal offset As Long) As LongPtr
    ReadLongPtrFromBytes = ReadLongPtr(VarPtr(b(offset)))
End Function


' ============================================================
' Get the function address of a fixed slot from the class's vTable
' ============================================================
Private Function GetClassMethodAddrAtFixedSlot( _
    ByVal cls As Object, ByVal slotIndex As Long) As LongPtr

    If cls Is Nothing Then Exit Function
    Dim pObj As LongPtr: pObj = ObjPtr(cls)
    If pObj = 0 Then Exit Function

    Dim pVTable As LongPtr
    pVTable = ReadLongPtr(pObj)
    If pVTable = 0 Then Exit Function

    GetClassMethodAddrAtFixedSlot = ReadLongPtr(pVTable + slotIndex * PtrSize)
End Function


' ============================================================
' IUnknown stubs (standard module)
' ============================================================
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

Private Function Handler_QueryInterface( _
    ByVal this As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long
    ppvObject = this
    Handler_QueryInterface = S_OK
End Function

Private Function Handler_AddRef(ByVal this As LongPtr) As Long
    Handler_AddRef = 1
End Function

Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = 1
End Function


' ============================================================
' Memory-operation primitives via PointerAccessor
' ============================================================
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

' https://github.com/WNKLER/RefTypes/discussions/3#discussion-8595790
Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub
```

```
''''''''''''''''''''''''''''''''''
' --- Class5.cls Stage 8 ---
'   The only Public method is Handler_Invoke (pinned at vTable slot 7).
'   The contents are equivalent to stages 5 to 7.
'   The point of stage 8 is that the thunk call on the standard module side
'   was changed to a two-step jump via the EntryPoint stub.
''''''''''''''''''''''''''''''''''
Option Explicit

' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
Public Function Handler_Invoke( _
    ByVal this As LongPtr, _
    ByVal errorCode As Long, _
    ByVal pEnvironment As LongPtr) As Long

    Debug.Print "Class5.Handler_Invoke is called directly from WebView2."
    Debug.Print "  this        :", this
    Debug.Print "  errorCode   :", errorCode
    Debug.Print "  pEnvironment:", pEnvironment

    Handler_Invoke = 0   ' S_OK
End Function
```

## Stage 8.5 — Reconnaissance phase: analysis of what the EntryPoint stub really is
 
In stage 8 the two-step jump structure was completed, and from stage 9 onward we needed to design a mechanism for mass-producing handlers.<br>
So, besides the naive pool scheme of "preparing two dedicated empty Subs (EntryPoint + DummyASM) per handler", the idea came up of "**could we mass-produce by reading the EntryPoint stub itself out with `PointerAccessor` and copying it to another region?**"<br>
Whether this scheme can work hinges on **whether the byte sequence of the EntryPoint stub contains any RIP-relative addressing instructions** (a RIP-relative instruction would have its disp32 shifted at the copy destination and crash, so relocation processing would be needed).<br>
So as stage 8.5 we created a dump tool, `Module8_5.bas`, and reconnoitered on a real machine what the stub region obtainable with `AddressOf EntryPoint` contains.
 
```
''''''''''''''''''''''''''''''''''
' --- Module8_5.bas ---
'   A reconnaissance tool that byte-dumps the stub region of an empty Sub obtainable with AddressOf.
'   Used for the design decision from stage 9 onward (whether the stub-copy scheme is feasible).
'
'   Main points to check:
'     - +0..+37   : the body of the EBMode check code added by the VBA runtime
'     - +37..+55  : processing after the check (probably the part just before mov rax, imm64)
'     - +55..+63  : the 8-byte immediate region (the place LibTimers rewrites)
'     - +63..    : the instruction that calls indirectly through the above immediate, and what follows
'
'   ★ Concerns:
'   - Whether there are RIP-relative addressing instructions (opcode starting with 48 8B 05 / FF 15, etc.).
'     If present, copying the stub to another region will not work (the relative offsets shift).
'   - If it consists only of absolute-immediate instructions (48 B8 XX×8 = mov rax, imm64),
'     mass production by copying is feasible.
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/RefTypes/discussions/3
''

Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "This module requires 64-bit VBA (x64)"
#End If

Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type


' ============================================================
' Empty Sub to be dumped
'   Make it exactly the same shape as the EntryPoint used in actual operation.
'   For Force-Compile, always call it once before taking AddressOf.
' ============================================================
Private Sub EntryPoint(): End Sub


' ============================================================
' Entry: dump output to the Immediate window
' ============================================================
Public Sub DumpEntryPointStub()
    ' Force-Compile
    EntryPoint

    Dim addr As LongPtr
    addr = VBA.Int(AddressOf EntryPoint)
    If addr = 0 Then
        Debug.Print "AddressOf EntryPoint failed."
        Exit Sub
    End If

    Debug.Print "==================================================="
    Debug.Print "EntryPoint stub @ 0x" & HexFmt(addr, 16)
    Debug.Print "==================================================="

    ' --- 1. Byte dump (16 bytes wide, 96 bytes worth) ---
    Debug.Print
    Debug.Print "--- Byte dump (16 bytes per line, 96 bytes total) ---"
    DumpBytes addr, 96

    ' --- 2. Dump in LongPtr (8 bytes) units ---
    Debug.Print
    Debug.Print "--- LongPtr (8-byte) view, 12 entries ---"
    Dim i As Long
    For i = 0 To 11
        Dim ofs As Long: ofs = i * 8
        Dim val As LongPtr: val = ReadLongPtr(addr + ofs)
        Debug.Print "  +" & PadL(CStr(ofs), 3) & " : 0x" & HexFmt(val, 16) & _
                   "   (decimal: " & val & ")"
    Next i

    ' --- 3. Focus on the immediate region at +55 (the position LibTimers rewrites) ---
    Debug.Print
    Debug.Print "--- Spotlight: +55 (lateBindOffset, where LibTimers writes) ---"
    Dim val55 As LongPtr: val55 = ReadLongPtr(addr + 55)
    Debug.Print "  Value at +55: 0x" & HexFmt(val55, 16)
    Debug.Print "  → This should be the imm64 part of a 'mov rax, imm64' instruction."
    Debug.Print "    If the value looks like it points to an address inside VBA7.dll or VBE7.dll,"
    Debug.Print "    the 2 bytes before it (addr+53, addr+54) should be 48 B8."

    Dim b53 As Byte: b53 = ReadByte(addr + 53)
    Dim b54 As Byte: b54 = ReadByte(addr + 54)
    Debug.Print "  +53 = 0x" & PadL(Hex(b53), 2, "0") & _
               "  +54 = 0x" & PadL(Hex(b54), 2, "0") & _
               IIf(b53 = &H48 And b54 = &HB8, "  ← mov rax, imm64 confirmed ?", "  ← differs from expected (48 B8)!")

    ' --- 4. Roughly detect the presence of RIP-relative instructions ---
    Debug.Print
    Debug.Print "--- RIP-relative instruction scan ---"
    ScanForRipRelative addr, 96
End Sub


' ============================================================
' Byte dump (16 bytes wide)
' ============================================================
Private Sub DumpBytes(ByVal addr As LongPtr, ByVal count As Long)
    Dim i As Long
    For i = 0 To count - 1 Step 16
        Dim line As String
        Dim ascii As String
        line = "+" & PadL(CStr(i), 3) & " : "

        Dim j As Long
        For j = 0 To 15
            If i + j >= count Then Exit For
            Dim b As Byte: b = ReadByte(addr + i + j)
            line = line & PadL(Hex(b), 2, "0") & " "
            If j = 7 Then line = line & " "  ' blank in the middle

            If b >= 32 And b < 127 Then
                ascii = ascii & Chr$(b)
            Else
                ascii = ascii & "."
            End If
        Next j

        Debug.Print line & " | " & ascii
    Next i
End Sub


' ============================================================
' Rough scan for RIP-relative addressing instructions
'   Look for the opcode patterns of typical RIP-relative instructions on x64.
'
'   This is not a complete disassembler, so false positives / misses are expected.
'   It is only meant to get a rough idea of whether such instructions are likely included.
'
'   Detection patterns (typical ones):
'     48 8B 05 ?? ?? ?? ??  : mov rax, [rip+disp32]
'     48 8B 0D ?? ?? ?? ??  : mov rcx, [rip+disp32]
'     48 8B 15 ?? ?? ?? ??  : mov rdx, [rip+disp32]
'     48 8B 1D / 25 / 2D / 35 / 3D : mov r{bx,sp,bp,si,di}, [rip+disp32]
'     48 8D 05/0D/...       : lea r??, [rip+disp32]
'     FF 15 ?? ?? ?? ??     : call qword ptr [rip+disp32]
'     FF 25 ?? ?? ?? ??     : jmp  qword ptr [rip+disp32]
'     E8 ?? ?? ?? ??        : call rel32  ← this is also relative!
'     E9 ?? ?? ?? ??        : jmp  rel32
' ============================================================
Private Sub ScanForRipRelative(ByVal addr As LongPtr, ByVal count As Long)
    Dim found As Long: found = 0
    Dim i As Long

    For i = 0 To count - 5
        Dim b0 As Byte: b0 = ReadByte(addr + i)
        Dim b1 As Byte: b1 = ReadByte(addr + i + 1)
        Dim b2 As Byte: b2 = ReadByte(addr + i + 2)

        Dim mnemonic As String: mnemonic = ""

        ' E8 / E9 (call rel32 / jmp rel32)
        If b0 = &HE8 Then
            mnemonic = "call rel32"
        ElseIf b0 = &HE9 Then
            mnemonic = "jmp rel32"

        ' 48 8B XX (mov r64, [rip+disp32]) ? the lower 3 bits of XX are rm=101 (= RIP-relative)
        ElseIf b0 = &H48 And b1 = &H8B Then
            If (b2 And &HC7) = &H5 Then
                mnemonic = "mov r64, [rip+disp32]"
            End If

        ' 48 8D XX (lea r64, [rip+disp32])
        ElseIf b0 = &H48 And b1 = &H8D Then
            If (b2 And &HC7) = &H5 Then
                mnemonic = "lea r64, [rip+disp32]"
            End If

        ' FF 15 XX (call qword ptr [rip+disp32])
        ElseIf b0 = &HFF And b1 = &H15 Then
            mnemonic = "call [rip+disp32]"

        ' FF 25 XX (jmp qword ptr [rip+disp32])
        ElseIf b0 = &HFF And b1 = &H25 Then
            mnemonic = "jmp [rip+disp32]"
        End If

        If LenB(mnemonic) > 0 Then
            ' Read the disp32 and compute the absolute target address
            Dim instLen As Long
            Dim dispOffset As Long
            If b0 = &HE8 Or b0 = &HE9 Then
                instLen = 5: dispOffset = 1
            ElseIf b0 = &HFF Then
                instLen = 6: dispOffset = 2
            Else  ' 48 8B / 48 8D
                instLen = 7: dispOffset = 3
            End If

            Dim disp32 As Long: disp32 = ReadLong(addr + i + dispOffset)
            Dim target As LongPtr: target = (addr + i + instLen) + disp32

            Debug.Print "  +" & PadL(CStr(i), 3) & " : " & mnemonic & _
                       "   disp32=" & disp32 & "  → target=0x" & HexFmt(target, 16)
            found = found + 1
        End If
    Next i

    If found = 0 Then
        Debug.Print "  No RIP-relative instructions were detected."
        Debug.Print "  → Mass-producing by copying the stub looks feasible in theory."
    Else
        Debug.Print
        Debug.Print "  ★ " & found & " RIP-relative instruction(s) detected."
        Debug.Print "  → If the stub is copied as is, the places they point to will shift and it will crash."
        Debug.Print "    When copying, disp32 must be recomputed (relocation) to match the new placement."
    End If
End Sub


' ============================================================
' Read 1 byte
'   Read 8 bytes with ReadLongPtr and extract the lowest 1 byte with a mask.
'   This avoids the trouble of preparing a separate ByteAccessor.
'   Note: the 7 bytes around addr are also accessed, but since use inside
'     a code region is assumed, the end-of-region problem can be ignored.
' ============================================================
Private Function ReadByte(ByVal addr As LongPtr) As Byte
    Dim val As LongPtr
    val = ReadLongPtr(addr)
    ReadByte = CByte(val And &HFF)
End Function

' ============================================================
' Read 4 bytes (Long)
'   Implemented by assembling 4 bytes (saving the trouble of
'   preparing a separate Long-typed version of PointerAccessor)
' ============================================================
Private Function ReadLong(ByVal addr As LongPtr) As Long
    Dim b0 As Long: b0 = ReadByte(addr + 0)
    Dim b1 As Long: b1 = ReadByte(addr + 1)
    Dim b2 As Long: b2 = ReadByte(addr + 2)
    Dim b3 As Long: b3 = ReadByte(addr + 3)
    ' Signed Long (little-endian)
    If b3 < &H80 Then
        ReadLong = b0 Or (b1 * &H100&) Or (b2 * &H10000) Or (b3 * &H1000000)
    Else
        ' Negative value: two's complement representation
        ReadLong = b0 Or (b1 * &H100&) Or (b2 * &H10000) Or _
                   ((b3 - &H100) * &H1000000)
    End If
End Function

' ============================================================
' Read 8 bytes (LongPtr)
'   A pattern with a proven track record since stage 6.
' ============================================================
Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function


' ============================================================
' Helper for injecting the SafeArray pointer into the PointerAccessor's internal arr field.
' A pattern with a proven track record since stage 6.
' ============================================================
Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub


' ============================================================
' Formatting helper
' ============================================================
Private Function HexFmt(ByVal val As LongPtr, ByVal width As Long) As String
    Dim h As String: h = Hex(val)
    HexFmt = PadL(h, width, "0")
End Function

Private Function PadL(ByVal s As String, ByVal width As Long, _
                      Optional ByVal padChar As String = " ") As String
    ' Insurance for cases where the Optional default value does not take effect
    If LenB(padChar) = 0 Then padChar = " "

    If Len(s) >= width Then
        PadL = s
    Else
        PadL = String$(width - Len(s), Left$(padChar, 1)) & s
    End If
End Function
```


### Dump result (VBA7 / x64 / Excel)
```
EntryPoint stub @ 0x000001DF171AE9B4
 
+0  : 48 89 4C 24 08          mov  [rsp+8],  rcx       ; save arg0
+5  : 48 89 54 24 10          mov  [rsp+10], rdx       ; save arg1
+10 : 4C 89 44 24 18          mov  [rsp+18], r8        ; save arg2
+15 : 4C 89 4C 24 20          mov  [rsp+20], r9        ; save arg3
+20 : 48 B8 B0 E9 1A 17 DF 01 00 00   mov  rax, 0x000001DF171AE9B0
                                                       ;   ↑ points to EntryPoint start - 4
                                                       ;     (metadata region tied to the Sub?)
+30 : 48 0B C0                or   rax, rax            ; null check
+33 : 74 32                   je   +0x32 (= +67)       ; skip if null
+35 : 48 B8 E4 D1 8D 9E FD 7F 00 00   mov  rax, 0x00007FFD9E8DD1E4
                                                       ;   ↑ EBMode getter function (inside VBA7.dll)
+45 : FF D0                   call rax                  ; call EBMode()
+47 : 48 83 F8 02             cmp  rax, 2              ; Break Mode (= 2) ?
+51 : 74 20                   je   +0x20 (= +85)       ; skip if Break Mode
+53 : 48 B8 AC 40 86 41 E0 01 00 00   mov  rax, 0x000001E0418640AC
                                                       ;   ↑ ★ "address to call next"
                                                       ;     this imm64 is what LibTimers rewrites
                                                       ;     as +55
+63 : 48 8B 4C 24 08          mov  rcx, [rsp+8]        ; restore arg0
+68 : 48 8B 54 24 10          mov  rdx, [rsp+10]       ; restore arg1
+73 : 4C 8B 44 24 18          mov  r8,  [rsp+18]       ; restore arg2
+78 : 4C 8B 4C 24 20          mov  r9,  [rsp+20]       ; restore arg3
+83 : FF E0                   jmp  rax                 ; tail call
+85 : 48 33 C0                xor  rax, rax            ; return value = 0
+88 : C2 00 00                ret  0                   ; return on the Break Mode path
                                                       ; (retn with 0 bytes of arguments)
```
 
### Important findings
 
**1. The stub's actual length is 91 bytes**<br>
Everything from +0 to +90 is meaningful instructions, and +91 onward is a margin buffer reserved by the VBA runtime. The 56-byte thunk written in stages 6 to 7 fits within this buffer.
 
**2. There are zero RIP-relative instructions**<br>
Every instruction of the stub consists of one of the following, so no relocation at copy time is needed at all:
- Relative access to the stack (`[rsp+N]`)
- Absolute address loads via `mov reg, imm64`
- Indirect calls via `call rax` / `jmp rax`
- Short-range `je rel8` relative jumps (`74 32`, `74 20`) — the offsets are closed inside the stub, so they work correctly even when copied
The dump tool's rough RIP-relative scanner had detected a "`jmp rel32` at +23", but that was actually a false positive on the byte `E9` in the imm64 part of the `mov rax, imm64` instruction at +20.
 
**3. The meaning of LibTimers' `+55` is now fully clear**<br>
Of the `48 B8 ...` (`mov rax, imm64`) instruction at +53, **the imm64 field starts at +55**. The 8 bytes written there become "the address to call next if not in Break Mode". The `Const lateBindOffset = 55` in LibTimers / MouseScroll points to this position.
 
**4. The address of the EBMode getter function**<br>
It loads directly, with `mov rax, imm64`, a typical address inside VBA7.dll / VBE7.dll, `0x00007FFD9E8DD1E4`. This is an ASLR-subject address that changes with each process start, but each Sub's stub does not hold its own; they should be calling a common function. This imm64 stays valid as is even when copied.
 
**5. The safe return path in Break Mode**<br>
Via `xor rax, rax` / `ret 0` at +85, the design returns immediately in Break Mode (EBMode = 2) without even restoring the arguments. This keeps the VBA interpreter from crashing even while stopped. When we copy and use the stub, this safety device keeps working as is.
 
**6. The meaning of `EntryPoint - 4`, referenced at the start of the stub**<br>
The immediate `0x000001DF171AE9B0` at +20 points 4 bytes before the start address of EntryPoint (`0x000001DF171AE9B4`). This is presumed to point to a metadata region that VBA ties to the Sub (an identifier? state? unknown).<br>
It is used only for the null check, so if the same value is present after copying, the behavior does not change (it is not null in any case, so it proceeds to the EBMode check). **It does not need to be rewritten at copy time**.
 
### Design options settled for stage 9
 
It is now settled that the mass-production-by-stub-copy scheme **works completely with no relocation needed**. As a result, the approaches available in stage 9 narrow down to the following three:
 
- **Plan X (naive pool)**: line up N pairs of `Private Sub EntryPointNN(): End Sub` and `Private Sub DummyASMNN(): End Sub`. Reliable, but takes many lines of code.
- **Plan Y (carve the DummyASM regions out of one huge Sub + N EntryPoints)**: save space by carving only the DummyASM side out of a huge Sub. EntryPoints are still needed N times.
- **Plan Z (pack N sets of an EntryPoint stub copy + thunk into one huge Sub)**: a new plan based on the stage 8.5 discovery. Copy N 91-byte stub clones into the code region of one huge Sub with PointerAccessor, rewrite the +55 of each clone independently, and have each call a different thunk. There is no need to line up large numbers of EntryPoints or DummyASMs in the standard module. The most elegant.
### Notes on the dump tool `Module8_5.bas`
- It outputs a byte dump (16 bytes wide × 6 rows), a LongPtr view (12 entries), a spotlight on the +55 immediate region, and a rough scan for RIP-relative instructions.
- Because VBA cannot pass a Private UDT as a Variant argument, `ReadByte` reads 8 bytes with `ReadLongPtr` and extracts the lowest 1 byte with a mask (avoiding a dedicated ByteAccessor type).
- A phenomenon was observed where the default value of `PadL`'s Optional `padChar` is passed as an empty string in some situations, so the function includes the insurance `If LenB(padChar) = 0 Then padChar = " "`.
- It can also be reused from stage 9 onward to verify "whether the copy-destination stub has been laid out as expected" (just extend it to take an arbitrary address as an argument).


## Stage 8.6 — Reconnaissance phase: verifying the executability of the module-level array region

In designing the mechanism for managing large numbers of handlers in stage 9, the idea came up for allocating the thunk region: "could we prepare a module-level `Private arr() As Byte` and write machine code directly into that region?"<br>
If this worked, we could avoid `AddressOf`, Force-Compile, the problem of how empty Subs are laid out and so on, and an extremely simple design where the address is obtained plainly with `VarPtr` would be possible.<br>
To check this possibility, as stage 8.6 we created a minimal verification tool, `ArrayExecProbe.bas`, and tried writing machine code into a module-level Byte array and executing it via `DispCallFunc`.

### What was verified
At the start address of the module-level variable `Private m_codeArea(0 To 31) As Byte`, we wrote minimal machine code (11 bytes) that takes no arguments and just returns an immediate value, and called it as a function pointer with `DispCallFunc`:
```
48 B8 78 56 34 12 00 00 00 00    mov rax, 0x12345678
C3                                ret
```

### Result
**Excel crashed.**

### Conclusion
The memory of a module-level array is allocated by the VBA runtime from the process heap as `PAGE_READWRITE` (readable and writable, not executable), so even if machine code is written there, the CPU raises an access violation because of DEP (Data Execution Prevention).<br>
In other words, **an array region cannot be used as a host for machine code**.

### Significance of this discovery
- The reason the authors of `LibTimers.bas` / `MouseScroll.bas` stuck to the Sub procedure region reached via `AddressOf` (= `PAGE_EXECUTE_READ`) is now settled. Since array regions are not executable because of DEP, they presumably reached the same conclusion.
- The means of obtaining an executable region from VBA narrow down, realistically, to the following 3:
  1. The Sub procedure region obtainable with `AddressOf` (adopted from stage 7 of this development onward)
  2. `VirtualAlloc(PAGE_EXECUTE_READWRITE)` (the scheme abolished in stage 6)
  3. Changing the attributes of an existing region to executable with `VirtualProtect` (in the end the same category as VirtualAlloc)
- Therefore, it is reasonable to narrow the reconnaissance from stage 8.7 onward to **how to secure a huge Sub procedure region**.

### Notes on the verification tool `ArrayExecProbe.bas`
- The first version of `WriteByte` tried to implement byte-wise access with bit operations, but `2^56 * &HFF` overflowed the positive range of `LongLong`. The final version was fixed to access directly one byte at a time via a `ByteAccessor` (a SafeArray with cbElements=1).
- A DEP violation cannot be caught with `On Error`, so save the Excel file before running the verification.


## Stage 8.7 — Reconnaissance phase: reconnaissance of a strategy for allocating a huge Sub procedure region
 
Stage 8.6 revealed that the module-level array region is not executable because of DEP, so we continue the reconnaissance narrowed to Sub procedure regions.<br>
To judge the feasibility of plan Z, which packs multiple stub clones into one Sub, we measured the following with `SizeProbe.bas` and `RangeDumper.bas`:
 
(B) **Observing the spacing between consecutively declared empty Subs**<br>
(A) **Measuring the actual size of an empty Sub**<br>
(C) **The size of a Sub with a body and the distance to the next Sub**<br>
(D) **Confirming whether declaration order matches address order**<br>
(E) **Searching for the end of `MediumBody`'s stub (byte dump)**

``` VBA
''''''''''''''''''''''''''''''''''
' --- Module8_7.bas ---
'   Stage 8.7 RangeDumper: a reconnaissance tool that dumps a memory region of arbitrary address and length in detail
'
'   A generalization of StubDumper.bas. It looks at the region of a Sub that has
'   a body, such as MediumBody or LargeBody, from end to end, and grasps the structure of
'   the end of the user-code part and the padding region after it.
'
'   Main usage:
'     DumpRange addr, count          ' byte dump only
'     DumpRangeEx addr, count        ' byte dump + RIP-relative scan + ret detection
'     DumpProcedure procName, count  ' take AddressOf by name and analyze in one go
'
'   Use this to obtain the information needed for the capacity design of plan Z
'   (placing several 91+56=147-byte stub clone + thunk sets in the user-code region of
'   one huge Sub).
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/RefTypes/discussions/3
''

Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "This module requires 64-bit VBA (x64)"
#End If

Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type


' ============================================================
' Group of named Subs to be measured
'   Same structure as SizeProbe.bas, so that they can be dumped immediately by name
' ============================================================
Private Sub Empty00(): End Sub

Private Sub SmallBody()
    Dim a As LongPtr
    a = 1: a = 2: a = 3: a = 4: a = 5
End Sub

Private Sub MediumBody()
    Dim a As LongPtr
    a = 1:  a = 2:  a = 3:  a = 4:  a = 5:  a = 6:  a = 7:  a = 8:  a = 9:  a = 10
    a = 11: a = 12: a = 13: a = 14: a = 15: a = 16: a = 17: a = 18: a = 19: a = 20
    a = 21: a = 22: a = 23: a = 24: a = 25: a = 26: a = 27: a = 28: a = 29: a = 30
    a = 31: a = 32: a = 33: a = 34: a = 35: a = 36: a = 37: a = 38: a = 39: a = 40
    a = 41: a = 42: a = 43: a = 44: a = 45: a = 46: a = 47: a = 48: a = 49: a = 50
    a = 51: a = 52: a = 53: a = 54: a = 55: a = 56: a = 57: a = 58: a = 59: a = 60
    a = 61: a = 62: a = 63: a = 64: a = 65: a = 66: a = 67: a = 68: a = 69: a = 70
    a = 71: a = 72: a = 73: a = 74: a = 75: a = 76: a = 77: a = 78: a = 79: a = 80
    a = 81: a = 82: a = 83: a = 84: a = 85: a = 86: a = 87: a = 88: a = 89: a = 90
    a = 91: a = 92: a = 93: a = 94: a = 95: a = 96: a = 97: a = 98: a = 99: a = 100
End Sub

Private Sub LargeBody()
    Dim a As LongPtr, b As LongPtr, c As LongPtr, d As LongPtr, e As LongPtr
    Dim n As Long
    For n = 1 To 100
        a = n: b = n + 1: c = n + 2: d = n + 3: e = n + 4
        a = n + 5: b = n + 6: c = n + 7: d = n + 8: e = n + 9
    Next n
    a = 1:   a = 2:   a = 3:   a = 4:   a = 5:   a = 6:   a = 7:   a = 8:   a = 9:   a = 10
    a = 11:  a = 12:  a = 13:  a = 14:  a = 15:  a = 16:  a = 17:  a = 18:  a = 19:  a = 20
    a = 21:  a = 22:  a = 23:  a = 24:  a = 25:  a = 26:  a = 27:  a = 28:  a = 29:  a = 30
    a = 31:  a = 32:  a = 33:  a = 34:  a = 35:  a = 36:  a = 37:  a = 38:  a = 39:  a = 40
    a = 41:  a = 42:  a = 43:  a = 44:  a = 45:  a = 46:  a = 47:  a = 48:  a = 49:  a = 50
    a = 51:  a = 52:  a = 53:  a = 54:  a = 55:  a = 56:  a = 57:  a = 58:  a = 59:  a = 60
    a = 61:  a = 62:  a = 63:  a = 64:  a = 65:  a = 66:  a = 67:  a = 68:  a = 69:  a = 70
    a = 71:  a = 72:  a = 73:  a = 74:  a = 75:  a = 76:  a = 77:  a = 78:  a = 79:  a = 80
    a = 81:  a = 82:  a = 83:  a = 84:  a = 85:  a = 86:  a = 87:  a = 88:  a = 89:  a = 90
    a = 91:  a = 92:  a = 93:  a = 94:  a = 95:  a = 96:  a = 97:  a = 98:  a = 99:  a = 100
End Sub


' ============================================================
' Public entry: dump everything by procedure name
'   ProcName : one of "Empty00", "SmallBody", "MediumBody", "LargeBody"
'   count    : number of bytes to dump (800 if omitted)
' ============================================================
Public Sub DumpProcedure(ByVal procName As String, _
                          Optional ByVal count As Long = 800)
    ' Force-Compile: call all the target Subs beforehand
Empty00:     SmallBody: MediumBody: LargeBody

    Dim addr As LongPtr
    Select Case LCase$(procName)
        Case "empty00":    addr = VBA.Int(AddressOf Empty00)
        Case "smallbody":  addr = VBA.Int(AddressOf SmallBody)
        Case "mediumbody": addr = VBA.Int(AddressOf MediumBody)
        Case "largebody":  addr = VBA.Int(AddressOf LargeBody)
        Case Else
            Debug.Print "Unknown procedure: " & procName
            Debug.Print "Available: Empty00, SmallBody, MediumBody, LargeBody"
            Exit Sub
    End Select

    Debug.Print "==================================================="
    Debug.Print " Dumping " & procName & " @ 0x" & HexFmt(addr, 16)
    Debug.Print " Range  : +0 .. +" & (count - 1) & " (" & count & " bytes)"
    Debug.Print "==================================================="
    DumpRangeEx addr, count
End Sub


' ============================================================
' Public entry: dump of arbitrary address and length
' ============================================================
Public Sub DumpRangeEx(ByVal addr As LongPtr, ByVal count As Long)
    ' --- 1. Byte dump (16 bytes wide) ---
    Debug.Print
    Debug.Print "--- Byte dump (16 bytes per line) ---"
    DumpBytes addr, count

    ' --- 2. Positions where ret 0 (C2 00 00) and ret (C3) appear ---
    Debug.Print
    Debug.Print "--- ret 0 (C2 00 00) and ret (C3) positions ---"
    FindRetInstructions addr, count

    ' --- 3. RIP-relative instruction scan ---
    Debug.Print
    Debug.Print "--- RIP-relative instruction scan ---"
    ScanForRipRelative addr, count
End Sub


' ============================================================
' Byte dump (16 bytes wide, with an ASCII column)
' ============================================================
Private Sub DumpBytes(ByVal addr As LongPtr, ByVal count As Long)
    Dim i As Long
    For i = 0 To count - 1 Step 16
        Dim line As String
        Dim ascii As String
        line = "+" & PadL(CStr(i), 4) & " : "

        Dim j As Long
        For j = 0 To 15
            If i + j >= count Then Exit For
            Dim b As Byte: b = ReadByte(addr + i + j)
            line = line & PadL(Hex(b), 2, "0") & " "
            If j = 7 Then line = line & " "

            If b >= 32 And b < 127 Then
                ascii = ascii & Chr$(b)
            Else
                ascii = ascii & "."
            End If
        Next j

        Debug.Print line & " | " & ascii
    Next i
End Sub


' ============================================================
' Detection of ret instructions
'   C2 00 00 : retn 0    (ret with 0 bytes of arguments)
'   C3       : ret       (ret without arguments)
'
'   Shows up to the first 100. Used to identify the boundaries of each VBA Sub.
' ============================================================
Private Sub FindRetInstructions(ByVal addr As LongPtr, ByVal count As Long)
    Dim i As Long
    Dim found As Long: found = 0
    Dim lastAt As Long: lastAt = -1
    For i = 0 To count - 1
        Dim b0 As Byte: b0 = ReadByte(addr + i)
        If b0 = &HC2 And i + 2 < count Then
            Dim b1 As Byte: b1 = ReadByte(addr + i + 1)
            Dim b2 As Byte: b2 = ReadByte(addr + i + 2)
            If b1 = &H0 And b2 = &H0 Then
                Dim deltaStr As String
                If lastAt >= 0 Then
                    deltaStr = "  (Δ +" & (i - lastAt) & ")"
                End If
                Debug.Print "  +" & PadL(CStr(i), 4) & " : C2 00 00  (retn 0)" & deltaStr
                lastAt = i
                found = found + 1
                If found >= 100 Then
                    Debug.Print "  ... (showing only the top 100)"
                    Exit Sub
                End If
            End If
        ElseIf b0 = &HC3 Then
            Debug.Print "  +" & PadL(CStr(i), 4) & " : C3        (ret)"
            found = found + 1
            If found >= 100 Then
                Debug.Print "  ... (showing only the top 100)"
                Exit Sub
            End If
        End If
    Next i
    If found = 0 Then
        Debug.Print "  No ret instructions found."
    End If
End Sub


' ============================================================
' Rough scan for RIP-relative instructions (same logic as StubDumper)
' ============================================================
Private Sub ScanForRipRelative(ByVal addr As LongPtr, ByVal count As Long)
    Dim found As Long: found = 0
    Dim i As Long

    For i = 0 To count - 5
        Dim b0 As Byte: b0 = ReadByte(addr + i)
        Dim b1 As Byte: b1 = ReadByte(addr + i + 1)
        Dim b2 As Byte: b2 = ReadByte(addr + i + 2)

        Dim mnemonic As String: mnemonic = ""

        If b0 = &HE8 Then
            mnemonic = "call rel32"
        ElseIf b0 = &HE9 Then
            mnemonic = "jmp rel32"
        ElseIf b0 = &H48 And b1 = &H8B Then
            If (b2 And &HC7) = &H5 Then mnemonic = "mov r64, [rip+disp32]"
        ElseIf b0 = &H48 And b1 = &H8D Then
            If (b2 And &HC7) = &H5 Then mnemonic = "lea r64, [rip+disp32]"
        ElseIf b0 = &HFF And b1 = &H15 Then
            mnemonic = "call [rip+disp32]"
        ElseIf b0 = &HFF And b1 = &H25 Then
            mnemonic = "jmp [rip+disp32]"
        End If

        If LenB(mnemonic) > 0 Then
            Dim instLen As Long, dispOffset As Long
            If b0 = &HE8 Or b0 = &HE9 Then
                instLen = 5: dispOffset = 1
            ElseIf b0 = &HFF Then
                instLen = 6: dispOffset = 2
            Else
                instLen = 7: dispOffset = 3
            End If

            Dim disp32 As Long: disp32 = ReadLong(addr + i + dispOffset)
            Dim target As LongPtr: target = (addr + i + instLen) + disp32

            Debug.Print "  +" & PadL(CStr(i), 4) & " : " & mnemonic & _
                       "   disp32=" & disp32 & "  → 0x" & HexFmt(target, 16)
            found = found + 1
            If found >= 50 Then
                Debug.Print "  ... (showing only the top 50)"
                Exit Sub
            End If
        End If
    Next i

    If found = 0 Then
        Debug.Print "  No RIP-relative instructions were detected."
    Else
        Debug.Print "  ★ " & found & " RIP-relative instruction(s) detected."
    End If
End Sub


' ============================================================
' Memory-reading primitives
' ============================================================
Private Function ReadByte(ByVal addr As LongPtr) As Byte
    Dim val As LongPtr
    val = ReadLongPtr(addr)
    ReadByte = CByte(val And &HFF)
End Function

Private Function ReadLong(ByVal addr As LongPtr) As Long
    Dim b0 As Long: b0 = ReadByte(addr + 0)
    Dim b1 As Long: b1 = ReadByte(addr + 1)
    Dim b2 As Long: b2 = ReadByte(addr + 2)
    Dim b3 As Long: b3 = ReadByte(addr + 3)
    If b3 < &H80 Then
        ReadLong = b0 Or (b1 * &H100&) Or (b2 * &H10000) Or (b3 * &H1000000)
    Else
        ReadLong = b0 Or (b1 * &H100&) Or (b2 * &H10000) Or _
                   ((b3 - &H100) * &H1000000)
    End If
End Function

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub

Private Function HexFmt(ByVal val As LongPtr, ByVal width As Long) As String
    Dim h As String: h = Hex(val)
    HexFmt = PadL(h, width, "0")
End Function

Private Function PadL(ByVal s As String, ByVal width As Long, _
                      Optional ByVal padChar As String = " ") As String
    If LenB(padChar) = 0 Then padChar = " "
    If Len(s) >= width Then
        PadL = s
    Else
        PadL = String$(width - Len(s), Left$(padChar, 1)) & s
    End If
End Function
```
 
### Measurement results
 
**Subs are not placed contiguously**<br>
The addresses of `Empty00..Empty15` were scattered, and declaration order and address order did not match at all. VBA allocates each Sub at an independent position, and adjacency is not guaranteed.<br>
Therefore the naive scheme of "lining up Subs and expecting their regions to be contiguous" breaks down.
 
**The region each Sub occupies is about 336 bytes**<br>
When we searched from the start of `MediumBody` for `ret 0` (`C2 00 00`) instructions, they were detected at two places, **+88 and +424**. The difference is **+336 bytes**, which is considered to be the unit size VBA allocates for one Sub.
 
**A shocking discovery: the user code (`a = 1` × 100 lines) does not exist in the machine-code region**<br>
When we dumped 600 bytes from +91 of `MediumBody`, most of the region was filled with `00 00 00 00 ...`, and the machine code for the 100 assignment statements was nowhere to be seen.
 
```
+0..+90    : stub #1 (the 91-byte EBMode check stub already confirmed in stage 8.5)
+91..+95   : 00 00 00 00 00     (padding)
+96..+103  : 4C A6 E1 BF 00 01 00 80   (metadata identifying the Sub)
+104..+207 : all 00 (contiguous free region, 104 bytes)
+208..+215 : 55 A6 F8 BF 00 02 00 80   (other metadata)
+216..+319 : all 00 (contiguous free region, 104 bytes)
+320..+335 : 52 A6 F3 BF 00 03 00 91 01 00 00 00   (metadata)
+336..+426 : stub #2 (the stub of another Sub)
...
```
 
### What this discovery means
 
**VBA keeps the body of a Sub not as machine code but as p-code in a separate region, and places only the EBMode check stub and identifier metadata in the machine-code region**. What the immediate at the stub's +55 points to is the entry of the VBA interpreter (the p-code execution engine), from which the Sub's real contents are interpreted and executed.
 
Each Sub is given **a free region of about 200 bytes (mostly filled with `00`)**, which may be usable as an ideal write destination for plan Z. Furthermore, we reconfirmed that the stub part **contains no RIP-relative instructions at all**, so mass production by the copy scheme remains valid.
 
 
---
 
## Stage 8.8 — Reconnaissance phase: verifying the executability of the free space of a Sub procedure
 
We check whether the "free space inside a Sub procedure" discovered in stage 8.7, unlike the array region, is executable.<br>
Hypothesis: the whole Sub procedure should be placed in a region with `PAGE_EXECUTE_READ` (or an equivalent attribute), so the free space inside it should also be executable as machine code.
 
### Verification method
We created `FreeAreaExecProbe.bas`, wrote the same minimal machine code (11 bytes) as in stage 8.6 at `MediumBody + 120` (= inside the free space), and called it via `DispCallFunc`:
```
48 B8 78 56 34 12 00 00 00 00    mov rax, 0x12345678
C3                                ret
```

``` VBA
''''''''''''''''''''''''''''''''''
' --- Module8_8.bas ---
'   Stage 8.8 FreeAreaExecProbe: verifying whether machine code can be written into
'   and executed in the "free space" inside a Sub procedure
'
'   Hypothesis:
'     As the stage 8.7 dump showed, VBA reserves an occupied region of
'     about 336 bytes for a Sub with a body, with the EBMode check stub in the first 91 bytes,
'     metadata at +96..+103, and the rest left as free space.
'     This free space should belong to the same PAGE_EXECUTE_READ region as
'     the whole Sub procedure, so we should be able to write machine code and execute it.
'
'   If this holds:
'     - A 91-byte stub + a 56-byte thunk fit in one Sub
'     - Just lining up N Subs creates N independent handler entrances
'     - Rewrite the +55 of each Sub to make it call the thunk in its own free space
'     - It completes with just Force-Compile and AddressOf, and neither VirtualAlloc
'       nor mass production by stub copying is needed
'
'   Verification method:
'     Write the same minimal machine code as stage 8.6's ArrayExecProbe
'     (mov rax, 0x12345678 / ret) to MediumBody+120 (inside the free space),
'     and call it as a function pointer with DispCallFunc.
'
'   ★ It may crash on failure, so save the Excel file
'     before running.
''''''''''''''''''''''''''''''''''

Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "This module requires 64-bit VBA (x64)"
#End If

Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

Private Type ByteAccessor
    arr() As Byte
    sa As SAFEARRAY_1D
End Type

' --- DispCallFunc (oleaut32) ---
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Integer, _
    ByRef prgpvarg As LongPtr, _
    ByRef pvargResult As Variant) As Long

Private Const CC_STDCALL As Long = 4
Private Const VT_I8      As Integer = 20
Private Const S_OK       As Long = 0


' ============================================================
' Sub under verification
'   Same structure as stage 8.7's MediumBody. Including 100 lines of assignments
'   makes VBA reserve about 336 bytes of code region.
' ============================================================
Private Sub MediumBody()
    Dim a As LongPtr
    a = 1:  a = 2:  a = 3:  a = 4:  a = 5:  a = 6:  a = 7:  a = 8:  a = 9:  a = 10
    a = 11: a = 12: a = 13: a = 14: a = 15: a = 16: a = 17: a = 18: a = 19: a = 20
    a = 21: a = 22: a = 23: a = 24: a = 25: a = 26: a = 27: a = 28: a = 29: a = 30
    a = 31: a = 32: a = 33: a = 34: a = 35: a = 36: a = 37: a = 38: a = 39: a = 40
    a = 41: a = 42: a = 43: a = 44: a = 45: a = 46: a = 47: a = 48: a = 49: a = 50
    a = 51: a = 52: a = 53: a = 54: a = 55: a = 56: a = 57: a = 58: a = 59: a = 60
    a = 61: a = 62: a = 63: a = 64: a = 65: a = 66: a = 67: a = 68: a = 69: a = 70
    a = 71: a = 72: a = 73: a = 74: a = 75: a = 76: a = 77: a = 78: a = 79: a = 80
    a = 81: a = 82: a = 83: a = 84: a = 85: a = 86: a = 87: a = 88: a = 89: a = 90
    a = 91: a = 92: a = 93: a = 94: a = 95: a = 96: a = 97: a = 98: a = 99: a = 100
End Sub


' ============================================================
' Entry: test of free-space executability
' ============================================================
Public Sub TestFreeAreaExecutable()
    ' Force-Compile
    MediumBody

    Debug.Print "==================================================="
    Debug.Print " Stage 8.8: executability test of the free space of a Sub procedure"
    Debug.Print "==================================================="

    Dim base As LongPtr
    base = VBA.Int(AddressOf MediumBody)
    Debug.Print " MediumBody @ 0x" & HexFmt(base, 16)

    ' --- Choose the write position ---
    '   Layout confirmed in stage 8.7:
    '     +0..+90   : stub (do not touch)
    '     +91..+95  : padding
    '     +96..+103 : metadata (do not touch)
    '     +104..+207 : contiguous free space ← place it in here
    '     +208..+215 : metadata
    '
    '   To be safe, leave a margin and write from +120.
    Const WRITE_OFFSET As Long = 120
    Dim addr As LongPtr
    addr = base + WRITE_OFFSET
    Debug.Print " Write destination : MediumBody + " & WRITE_OFFSET & _
               "  (= 0x" & HexFmt(addr, 16) & ")"
    Debug.Print

    ' --- Save the bytes before writing (for confirmation) ---
    Debug.Print " Before writing (12 bytes):"
    DumpInline addr, 12

    ' --- Write the machine code ---
    '   48 B8 78 56 34 12 00 00 00 00    mov rax, 0x12345678
    '   C3                                ret
    Debug.Print " Writing machine code..."
    WriteByte addr + 0, &H48
    WriteByte addr + 1, &HB8
    WriteByte addr + 2, &H78
    WriteByte addr + 3, &H56
    WriteByte addr + 4, &H34
    WriteByte addr + 5, &H12
    WriteByte addr + 6, &H0
    WriteByte addr + 7, &H0
    WriteByte addr + 8, &H0
    WriteByte addr + 9, &H0
    WriteByte addr + 10, &HC3
    Debug.Print

    Debug.Print " After writing (12 bytes):"
    DumpInline addr, 12
    Debug.Print

    ' --- Run the verification ---
    Debug.Print " Attempting the call via DispCallFunc..."
    Debug.Print " (If it crashes here, the free space is not executable."
    Debug.Print "  If it does not crash, it is executable and the return value should be 0x12345678)"
    Debug.Print

    Dim result As Variant
    Dim hr As Long
    Dim dummyVT As Integer
    Dim dummyArg As LongPtr

    hr = DispCallFunc( _
        0, addr, CC_STDCALL, VT_I8, 0, _
        dummyVT, dummyArg, result)

    Debug.Print " ===== Result ====="
    Debug.Print " HRESULT      : 0x" & HexFmt(hr, 8)
    Debug.Print " result       : 0x" & HexFmt(CLngLng(result), 16) & _
               "  (decimal: " & CLngLng(result) & ")"

    If hr = S_OK And CLngLng(result) = &H12345678^ Then
        Debug.Print
        Debug.Print " ★★★ SUCCESS ★★★"
        Debug.Print " The free space of a Sub procedure is executable!"
        Debug.Print " Plan Z (a scheme in which one Sub hosts the stub + thunk"
        Debug.Print " for one handler) is feasible."
    Else
        Debug.Print
        Debug.Print " Note: unexpected result"
        Debug.Print " HRESULT is not S_OK, or result does not match 0x12345678."
    End If
End Sub


' ============================================================
' Inline dump (N bytes on one line)
' ============================================================
Private Sub DumpInline(ByVal addr As LongPtr, ByVal count As Long)
    Dim s As String, i As Long
    s = "  "
    For i = 0 To count - 1
        s = s & " " & PadL(Hex(ReadByte(addr + i)), 2, "0")
    Next i
    Debug.Print s
End Sub


' ============================================================
' Byte-level memory access (via ByteAccessor)
' ============================================================
Private Sub WriteByte(ByVal addr As LongPtr, ByVal value As Byte)
    Dim ba(0 To 0) As ByteAccessor
    With ba(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = 1
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        InjectSafeArrayPtr_Byte ba, VarPtr(.sa)
        .arr(0) = value
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Sub

Private Function ReadByte(ByVal addr As LongPtr) As Byte
    Dim ba(0 To 0) As ByteAccessor
    With ba(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = 1
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        InjectSafeArrayPtr_Byte ba, VarPtr(.sa)
        ReadByte = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Private Sub InjectSafeArrayPtr_Byte(ByRef target() As ByteAccessor, ByVal saPtr As LongPtr)
    Dim slotAddr As LongPtr: slotAddr = VarPtr(target(0))
    MemLongPtr(slotAddr) = saPtr
End Sub


' ============================================================
' LongPtr-level memory access
' ============================================================
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub


' ============================================================
' Formatting helper
' ============================================================
Private Function HexFmt(ByVal val As LongPtr, ByVal width As Long) As String
    Dim h As String: h = Hex(val)
    HexFmt = PadL(h, width, "0")
End Function

Private Function PadL(ByVal s As String, ByVal width As Long, _
                      Optional ByVal padChar As String = " ") As String
    If LenB(padChar) = 0 Then padChar = " "
    If Len(s) >= width Then
        PadL = s
    Else
        PadL = String$(width - Len(s), Left$(padChar, 1)) & s
    End If
End Function
```

 
### Result
**Write completed → call via DispCallFunc succeeded → return value `0x12345678` confirmed ★**<br>
The byte sequence that was `00 00 00 00 ...` before writing was rewritten to `48 B8 78 56 34 12 00 00 00 00 C3 ...`, demonstrating that it can be executed as a function.
 
### What this discovery means
- **The free space inside a Sub procedure allows all of read, write and execute**. Neither writing nor execution is blocked.
- That the writes into the `InvokeThunkSlot` region worked in stage 7 was not a special case but the standard nature of Sub regions.
- It is settled that plan Z (a scheme in which one Sub hosts the stub + thunk for one handler) is technically feasible.
---
 
### Considering a return to the VirtualAlloc scheme
 
In stage 8.8 plan Z, "one Sub hosts the stub + thunk for one handler", turned out to be feasible, but when we compared it against the VirtualAlloc scheme toward the stage 9 implementation, we judged **the VirtualAlloc scheme to be overwhelmingly superior**.
 
### Summary of the comparison
 
| Criterion | Mass placement of procedures | VirtualAlloc + PointerAccessor |
|---|---|---|
| Startup delay | △ Force-Compile initializes all Subs (hundreds of ms expected) | ◎ One API call + only the necessary writes |
| Defender slowdown | ○ Does not occur | ○ Does not occur (CopyMemory not used) |
| Lines of code | × 100 lines × 512 Subs = 51,200 lines+α | ◎ A few dozen lines |
| Flexibility in slot count | × Fixed (changes by recompiling) | ◎ The allocated size can be adjusted dynamically |
| Outlook of the design | △ Must manage the arrangement of the Subs | ◎ Address calculation (`base + N × SLOT_SIZE`) |
 
### Points that became clear in retrospect
 
The original motivation for aiming to abolish CopyMemory in stage 6 was to eliminate the **"slowdown caused by Defender monitoring every `RtlMoveMemory` call"**, and even if VirtualAlloc alone is monitored, the cost is negligible if it is done once at startup.<br>
Going as far as abolishing VirtualAlloc in stage 7 was due to the secondary goal of "wanting to minimize the API patterns Defender monitors", and there was no real harm in operation.<br>
From stage 9 onward we need to manage large numbers of handlers, and since **the VirtualAlloc scheme wins on code simplicity, flexibility and startup performance alike**, we change course here.
 
### Design policy of stage 9
 
1. At startup, allocate an N×147-byte region with `VirtualAlloc(PAGE_EXECUTE_READWRITE)`
2. Read out with `PointerAccessor` the 91 bytes of the stub of the single `EntryPoint` Sub prepared in the standard module, and copy it N times into the allocated region
3. Write into the +55 of each copy the address of "the thunk region in the same set (56 bytes starting at stub end +91)"
4. Bake a different `ObjPtr(handlerN)` and `pHandlerInvoke` into each thunk
5. Register the start address of the corresponding stub clone in vTable[3]
6. At shutdown, free everything at once with `VirtualFree`
Since the stub contains zero RIP-relative instructions (confirmed in stages 8.5/8.7), no relocation processing is needed. It can be mass-produced with a simple bulk copy.


## Stage 8.9 — Minimal implementation of the VirtualAlloc + stub-copy scheme (1 handler)
 
In the reconnaissance phase through stage 8.8, the design policy for stage 9 was settled on **VirtualAlloc + stub copy**.<br>
Before moving on to full-scale management of many handlers, stage 8.9 first confirms that the new scheme works with just 1 handler. We duplicated stage 8's Class5 / Module8 into Class6 / Module8_9 and replaced only the mechanism for allocating the thunk region.
 
### Differences from stage 8
 
| Item | Stage 8 | Stage 8.9 |
|---|---|---|
| Win32 API declarations | None | `VirtualAlloc` / `VirtualFree` |
| Empty Subs | Two: `EntryPoint` + `DummyASM` | Only one: `EntryPoint` |
| Thunk region | The code region of `DummyASM` (prepared by VBA) | 152 bytes allocated with `VirtualAlloc` |
| Stub body | Uses `EntryPoint` (the one VBA placed) directly | A clone copied into the VirtualAlloc region |
| vTable[3] registration | `AddressOf EntryPoint` | Start of the clone stub (= VirtualAlloc return value) |
| Rewrite target of +55 | The address of `DummyASM` | +96 in the same region (start of the thunk) |
 
### Slot layout (152 bytes)
```
+0..+90    : copy of the EntryPoint stub (91 bytes, only +55 rewritten)
+91..+95   : padding (a region that is never executed; overwriting it at copy time is harmless)
+96..+151  : thunk (56 bytes)
```
By rounding SLOT_SIZE up to a multiple of 8, 152 bytes, reads and writes complete in the `MemLongPtr` 8-byte-unit loop.
 
### Division of roles
- **The original EntryPoint (inside the standard module, placed by the VBA runtime)**: only read as the source (template) of the stub's byte sequence. It is called once at startup for Force-Compile, but after that it does not appear on the execution path. It is not registered in the vTable.
- **The clone stub inside the VirtualAlloc region**: the entrance WebView2 actually calls. Thanks to the +55 rewrite, after passing the check it jumps to the thunk in the same region.
- **The thunk**: the same 56 bytes of machine code as from stage 6 onward.
### Execution flow
```
WebView2 calls vTable[3]
        ↓
Clone stub (argument saving, EBMode check)
        ↓
Check passed → indirect jump to wherever the immediate at +55 points
        ↓
VirtualAlloc region +96 = thunk (argument shift + ObjPtr injection)
        ↓
Class6.Handler_Invoke
        ↓
Return to the thunk → ret → return to the stub → ret → to WebView2
```
 
### Outlook toward stage 9
Slots are completely independent with no mutual interference, so scaling up can be done mechanically:
- Allocate N slots at once with `VirtualAlloc(0, SLOT_SIZE * N, ...)`
- Copy the stub into each slot and write a unique thunk
- When creating a handler, take a free slot and register its address in vTable[3]
- At shutdown, free all slots with a single `VirtualFree`
CopyMemory remains unused. Both the copying of the stub and the writing of the thunk are done in `MemLongPtr` loops, so no Defender slowdown occurs. VirtualAlloc / VirtualFree are each called just once at startup and shutdown, so the startup delay is also negligible.


``` VBA
''''''''''''''''''''''''''''''''''
' --- Module8_9.bas Stage 8.9 (VirtualAlloc + stub copy, minimal 1-handler implementation) ---
'
'   Differences from stage 8:
'     - Abolish DummyASM (the empty Sub); the thunk region is allocated with VirtualAlloc
'     - Only one EntryPoint stub is prepared in the standard module, and
'       its byte sequence is read out with PointerAccessor and copied into the VirtualAlloc region
'       for 91 bytes (no relocation needed, confirmed in stage 8.5)
'     - The +55 of the stub clone is rewritten with the address of the thunk in the VirtualAlloc region;
'       the thunk itself is the same 56 bytes of machine code as from stage 6 onward
'     - What is registered in vTable[3] is the stub clone (= start of the VirtualAlloc region)
'
'   Structure:
'     vTable[3] = start of the VirtualAlloc region
'                       ↓
'                 clone of the EntryPoint stub (91 bytes)
'                       ↓
'                 after passing the EBMode check, jump to the address at +55
'                       ↓
'                 thunk in the same VirtualAlloc region (offset +96)
'                       ↓
'                 inject ObjPtr(gHandler) into Me and go to Class6.Handler_Invoke
'
'   Layout (152 bytes allocated):
'     +0..+90    : copy of the EntryPoint stub (91 bytes)
'     +91..+95   : padding (a region that may be overwritten but is never executed)
'     +96..+151  : thunk (56 bytes)
'
'   CopyMemory remains unused. Both copying and rewriting are done
'   via PointerAccessor (MemLongPtr).
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/refTypes/discussions/3
''

Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "This module requires 64-bit VBA (x64)"
#End If

' --- Structures for SafeArray / PointerAccessor ---
Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

' --- WebView2Loader.dll ---
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

' --- Win32 API (VirtualAlloc / VirtualFree only. CopyMemory is not used) ---
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long) As LongPtr

Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long

Private Const MEM_COMMIT             As Long = &H1000
Private Const MEM_RESERVE            As Long = &H2000
Private Const MEM_RELEASE            As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const S_OK                   As Long = 0

' Position within the EntryPoint stub where the "address to call next" is embedded as an immediate
Private Const LATE_BIND_OFFSET As Long = 55

' Actual length of the stub (confirmed in stage 8.5; 91 bytes including the C2 00 00 at +88)
Private Const STUB_LEN As Long = 91

' Length of the thunk (56 bytes, same as from stage 6 onward)
Private Const THUNK_LEN As Long = 56

' Allocation size per slot (rounded to a multiple of 8, 152 bytes)
'   +0..+90    stub clone (91 bytes)
'   +91..+95   padding
'   +96..+151  thunk (56 bytes)
Private Const SLOT_SIZE      As Long = 152
Private Const THUNK_OFFSET   As Long = 96


' --- WebView2 Handler structure ---
Private Type WebView2Handler
    pVTable      As LongPtr
    Functions(3) As LongPtr
End Type
Private m_Handler As WebView2Handler

Private gHandler As Class6
Private m_pSlotBase As LongPtr   ' Start of the region allocated with VirtualAlloc


' ============================================================
' Source of the EntryPoint stub
'   The byte sequence of this stub is copied into the VirtualAlloc region.
'   It must be called once at startup for Force-Compile.
'   ★ After copying, never call this Sub directly (this stub itself is
'     a legitimate entry that calls the interpreter, but the copied stub
'     has its +55 rewritten, so that one is safe).
' ============================================================
Private Sub EntryPoint(): End Sub


' ============================================================
' Public entry
' ============================================================
Public Sub CreateWebView2Environment()

    Set gHandler = New Class6

    Dim pHandlerInvoke As LongPtr
    pHandlerInvoke = GetClassMethodAddrAtFixedSlot(gHandler, 7)
    If pHandlerInvoke = 0 Then
        MsgBox "Failed to obtain the address of Class6.Handler_Invoke", vbCritical
        Exit Sub
    End If

    ' --- VirtualAlloc + stub copy + thunk write ---
    Dim pSlot As LongPtr
    pSlot = AllocateAndBuildSlot(ObjPtr(gHandler), pHandlerInvoke)
    If pSlot = 0 Then
        MsgBox "Failed to build the slot", vbCritical
        Exit Sub
    End If

    ' --- Build the vTable ---
    Static pa As PointerAccessor
    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = PtrSize
        pa.sa.cLocks = 1
        MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
    End If

    m_Handler.pVTable = VarPtr(m_Handler.Functions(0))

    pa.sa.pvData = VarPtr(m_Handler.Functions(0))
    pa.sa.rgsabound0.cElements = 4

    pa.arr(0) = GetAddr(AddressOf Handler_QueryInterface)
    pa.arr(1) = GetAddr(AddressOf Handler_AddRef)
    pa.arr(2) = GetAddr(AddressOf Handler_Release)
    pa.arr(3) = pSlot                    ' ← the start of the stub clone is the entrance

    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = 0

    ' --- Create the WebView2 environment ---
    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, VarPtr(m_Handler))

    If hr <> S_OK Then
        MsgBox "Failed to initialize WebView2. HRESULT: &H" & Hex(hr), vbCritical
    End If

    DoEvents
End Sub

Public Sub FinalizeWebView2Environment()
    If m_pSlotBase <> 0 Then
        VirtualFree m_pSlotBase, 0, MEM_RELEASE
        m_pSlotBase = 0
    End If
    Set gHandler = Nothing
End Sub


' ============================================================
' Allocate the VirtualAlloc region, copy the stub and write the thunk
'
'   Return value: the start address of the slot (= the stub clone)
' ============================================================
Private Function AllocateAndBuildSlot( _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr) As LongPtr

    ' Force-Compile EntryPoint (materialize the source of the stub)
    EntryPoint

    ' Source address of the EntryPoint stub
    Dim pStubSrc As LongPtr
    pStubSrc = VBA.Int(AddressOf EntryPoint)
    If pStubSrc = 0 Then Exit Function

    ' Allocate 152 bytes of executable region with VirtualAlloc
    m_pSlotBase = VirtualAlloc(0, SLOT_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If m_pSlotBase = 0 Then Exit Function

    ' --- Copy the stub for 96 bytes (rounded to a multiple of 8; overwriting +91..+95 is harmless) ---
    Dim k As Long
    For k = 0 To 95 Step 8
        MemLongPtr(m_pSlotBase + k) = ReadLongPtr(pStubSrc + k)
    Next k

    ' --- Write the thunk (56 bytes from offset +96) ---
    WriteThunkMachineCode m_pSlotBase + THUNK_OFFSET, pSelfObj, pTargetFunc

    ' --- Write the thunk's address into +55 of the stub clone ---
    MemLongPtr(m_pSlotBase + LATE_BIND_OFFSET) = m_pSlotBase + THUNK_OFFSET

    AllocateAndBuildSlot = m_pSlotBase
End Function


' ============================================================
' Write the thunk's machine code to the specified address (same as stages 6 to 8)
'
'   For the layout and alignment explanation, see the stage 6 comments.
' ============================================================
Private Sub WriteThunkMachineCode( _
    ByVal addr As LongPtr, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr)

    Const THUNK_SIZE As Long = 56
    Dim b(0 To THUNK_SIZE - 1) As Byte, i As Long
    i = 0

    ' sub rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4
    ' mov r9, r8
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    ' mov rcx, imm64   (ObjPtr(gHandler))
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pSelfObj: i = i + 8
    ' lea rax, [rsp+0x28]
    b(i) = &H48: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5
    ' mov [rsp+0x20], rax
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H20: i = i + 5
    ' mov rax, imm64   (pTargetFunc)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pTargetFunc: i = i + 8
    ' call rax
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
    ' mov eax, [rsp+0x28]
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H28: i = i + 4
    ' add rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4
    ' ret
    b(i) = &HC3: i = i + 1
    ' int3 / int3 (padding)
    b(i) = &HCC: i = i + 1
    b(i) = &HCC: i = i + 1

    ' Write the byte sequence 7 times in 8-byte units
    Dim k As Long
    For k = 0 To THUNK_SIZE - 1 Step 8
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k
End Sub


' ============================================================
' Take 8 bytes (LongPtr) from the given offset of a byte array
' ============================================================
Private Function ReadLongPtrFromBytes(ByRef b() As Byte, ByVal offset As Long) As LongPtr
    ReadLongPtrFromBytes = ReadLongPtr(VarPtr(b(offset)))
End Function


' ============================================================
' Get the function address of a fixed slot from the class's vTable
' ============================================================
Private Function GetClassMethodAddrAtFixedSlot( _
    ByVal cls As Object, ByVal slotIndex As Long) As LongPtr

    If cls Is Nothing Then Exit Function
    Dim pObj As LongPtr: pObj = ObjPtr(cls)
    If pObj = 0 Then Exit Function

    Dim pVTable As LongPtr
    pVTable = ReadLongPtr(pObj)
    If pVTable = 0 Then Exit Function

    GetClassMethodAddrAtFixedSlot = ReadLongPtr(pVTable + slotIndex * PtrSize)
End Function


' ============================================================
' IUnknown stubs (standard module)
' ============================================================
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

Private Function Handler_QueryInterface( _
    ByVal this As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long
    ppvObject = this
    Handler_QueryInterface = S_OK
End Function

Private Function Handler_AddRef(ByVal this As LongPtr) As Long
    Handler_AddRef = 1
End Function

Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = 1
End Function


' ============================================================
' Memory-operation primitives via PointerAccessor
' ============================================================
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub
```

``` VBA
''''''''''''''''''''''''''''''''''
' --- Class6.cls Stage 8.9 ---
'   The only Public method is Handler_Invoke (pinned at vTable slot 7).
'   The contents are equivalent to stages 5 to 8.
'   Stage 8.9 switches the thunk region allocation to VirtualAlloc and, as preparation for
'   mass-producing stubs by copying, first confirms operation with just 1 handler.
''''''''''''''''''''''''''''''''''
Option Explicit

' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
Public Function Handler_Invoke( _
    ByVal this As LongPtr, _
    ByVal errorCode As Long, _
    ByVal pEnvironment As LongPtr) As Long

    Debug.Print "Class6.Handler_Invoke is called directly from WebView2."
    Debug.Print "  this        :", this
    Debug.Print "  errorCode   :", errorCode
    Debug.Print "  pEnvironment:", pEnvironment

    Handler_Invoke = 0   ' S_OK
End Function
```


## Stage 9.0 — Verifying the slot pool + alive flag scheme (1 handler)
 
A first step in extending the "VirtualAlloc + stub copy" scheme established in stage 8.9 into a foundation for managing multiple handlers. We introduced a slot pool (512 slots) allocated in bulk at startup, plus an **alive flag** to give resilience to VBA resets. Operation was again confirmed with just 1 handler.
 
### Differences from stage 8.9
 
| Item | Stage 8.9 | Stage 9.0 |
|---|---|---|
| Region allocation | `VirtualAlloc(0, 152, ...)` (for 1 slot) | `VirtualAlloc(0, 64 + 176×512, ...)` (header + 512 slots) |
| Alive flag | None | 1 byte (0/1) at the start of the region, checked by the thunk every time |
| Thunk length | 56 bytes | 74 bytes (18 bytes of flag check + the existing 56 bytes) |
| SLOT_SIZE | 152 | 176 |
| Slot pool | None (fixed at 1) | Free list (O(1) acquire/return with a `Long` array) |
| API | One `AllocateAndBuildSlot` | `Thunks_Init` / `Thunks_AcquireSlot` / `Thunks_ReleaseSlot` / `Thunks_Shutdown` |
| At Shutdown | `VirtualFree` only | Set the flag to 0 → `VirtualFree` |
 
### Region layout (total 90,176 bytes ≒ 88 KB)
 
```
+0..+7        : alive flag (LongPtr, 8 bytes; only the lowest 1 byte is used)
+8..+63       : header reserved area (for a future sentinel flag-clearing thunk, etc.)
+64           : slot 0 starts (176 bytes)
+240          : slot 1 starts
...
+64+176×511   : slot 511 starts
 
Each slot (176 bytes):
  +0..+90      : clone of the EntryPoint stub (91 bytes)
  +91..+95     : padding
  +96..+169    : thunk (74 bytes)
  +170..+175   : padding (8-byte alignment)
```
 
### Machine code of the extended thunk (74 bytes)
 
At the start of the 56-byte thunk of stages 6 to 8, we insert an 18-byte alive flag check:
 
```asm
+0  : 48 B8 [8 bytes imm64]   mov  rax, pAliveFlag      ; start address of the region
+10 : 80 38 01                cmp  byte ptr [rax], 1
+13 : 74 03                   je   alive (+18)
+15 : 33 C0                   xor  eax, eax              ; S_OK return value
+17 : C3                      ret                        ; return immediately if dead
+18 : (from here on, the same 56-byte thunk body as stages 6 to 8)
```
 
Since this is a tail call that arrives from the stub with `jmp rax`, the thunk's `ret` returns directly to WebView2's caller. `xor eax, eax` zero-extends the whole of RAX, so HRESULT = S_OK is returned correctly. The argument registers (rcx, rdx, r8, r9) are not touched in the flag-check block, so they are passed as they are to the main processing when alive.
 
### Implementation of the slot pool
 
The free list is implemented with a `Long` array. It is a classic singly linked list in which `m_freeHead` holds the index of the first free slot and `m_freeNext(i)` represents "the next free slot if slot i is free". Acquire / Release are both O(1). Since `m_freeNext(i) = -1` expresses "slot i is in use", detecting a double release is also easy.
 
```vba
ReDim m_freeNext(0 To SLOT_COUNT - 1)
For i = 0 To SLOT_COUNT - 2
    m_freeNext(i) = i + 1
Next i
m_freeNext(SLOT_COUNT - 1) = -1
m_freeHead = 0
```
 
### Operation check results
 
- Acquire just 1 handler → reaches Class6.Handler_Invoke (= the same behavior as stage 8.9)
- No crash even when repeating Create / Finalize 5 times in a row
- The bulk copy at startup (MemLongPtr × 6,144 times) is **imperceptible**
- No Defender slowdown either
- However, observation of slot numbers and object leaks is not yet implemented (an item for stage 9.1 onward)
### Design points established in stage 9.0
 
1. **The region is allocated in bulk at startup**: just one `VirtualAlloc(0, 88KB, ...)`. `VirtualFree` at shutdown is also just once.
2. **Header at the start of the region**: 8 bytes of alive flag + 56 bytes reserved. The sentinel's flag-clearing thunk of stage 9.5 is planned to go here.
3. **Check the alive flag at the start of the thunk**: when dead (= flag 0), it returns S_OK immediately before entering the thunk body. This establishes a double defense in which, both at normal exit and at reset, "no crash even if a callback arrives with a handler left behind".
4. **Speed of PointerAccessor**: 6,144 MemLongPtr calls are imperceptible. The cost of the bulk copy is practically negligible.
5. **The class instance's ObjPtr is reused with the same value**: behavior of VBA's internal allocator. Even when repeating Create / Finalize, a new Class6 instance is generated at the same address, but the ObjPtr baked into the thunk's `mov rcx, imm64` is the value at Acquire time, so it is rewritten every time and consistency is maintained.
### Outlook toward stage 9.1
 
- Newly create `WebView2Handler.cls` and make it a GenericHandler that handles multiple kinds via `m_kind` (plan B)
- Have the standard module hold a `m_handlers(0 To 511) As Object` correspondence table so that `Thunks_ReleaseSlot` can be called from `Handler_Release` by reverse lookup
- A real reference count implementation (moving away from the current fixed Release = 1)
- Replace Class6 in one stage, keeping behavioral compatibility up to 9.0


## Stage 9.1 — GenericHandler-ization and a real reference count implementation (1 handler)

On top of the stage 9.0 slot pool, we (a) replaced `Class6` with `Class7`, which "handles multiple kinds with m_kind", and (b) changed AddRef / Release on the standard module side into real reference counter operations. Operation was again confirmed with just 1 handler.

### Differences from stage 9.0

| Item | Stage 9.0 | Stage 9.1 |
|---|---|---|
| Handler class | `Class6` (a single kind) | `Class7` (multiple kinds via `m_kind`; the only Public member is `Handler_Invoke`) |
| Enum | None | `HandlerKind` (6 kinds such as HK_EnvironmentCompleted) |
| Reference count | A mock implementation where `Handler_AddRef`/`Release` always return 1 | A real implementation that actually increments and decrements `Class7.m_refCount` |
| QueryInterface | Accepts all IIDs and returns `S_OK` | Same, but performs the equivalent of AddRef internally |
| Automatic release | None | At refcount = 0, chains ClearOwner → clear m_handlers → ReleaseSlot |
| `m_handlers(0 To 511)` array | None | Newly created. Registered/unregistered by Acquire/Release (not actually used at this stage; groundwork for stage 9.2) |
| Initial refcount | (irrelevant since it was a mock) | Starts at **0** (plan X-1, decided from real-machine logs) |

### Structure of Class7

```vba
Public Enum HandlerKind        ' Placed on the standard module side (so that Class7 can reference it)
    HK_None = 0
    HK_EnvironmentCompleted = 1
    HK_ControllerCompleted = 2
    HK_NavigationStarting = 3
    HK_NavigationCompleted = 4
    HK_WebMessageReceived = 5
    HK_DocumentTitleChanged = 6
End Enum

' Inside Class7.cls
Private m_kind     As HandlerKind
Private m_owner    As Object       ' Upper-level class (used from stage 9.2 onward)
Private m_pSlot    As LongPtr      ' Thunk address that WebView2 calls (= start of the slot)
Private m_refCount As Long

Friend Sub Init(kind, owner, pSlot)         ' Initialization, starting from m_refCount = 0
Friend Property Get/Let RefCount() As Long
Friend Property Get Slot() As LongPtr
Friend Property Get Kind() As HandlerKind
Friend Sub ClearOwner()                      ' Called when refcount=0; cuts the circular reference

Public Function Handler_Invoke( _
    ByVal this As LongPtr, _
    ByVal arg1 As LongPtr, _
    ByVal arg2 As LongPtr) As Long
    ' Branch on m_kind. In stage 9.1 only HK_EnvironmentCompleted was actually confirmed to work
End Function
```

We designed `Handler_Invoke` to receive all 3 arguments as `LongPtr`. For `HK_EnvironmentCompleted`, `arg1` is an HRESULT (32 bit), but since the thunk shifts with `mov r8, rdx` in 64 bit, garbage may be in the upper 32 bits. Inside the class, the `LongPtrLowDword` helper is used to safely extract the lower 32 bits as a signed Long.

### Significant design matter discovered during development

When we started the implementation, the handover prompt said "reverse-lookup the index from `this` (= the slot's start address)", but this **turned out to be inconsistent with the COM calling convention**.

```
[m_vtableObj layout (module-level variable, the only instance)]
m_vtableObj.pVTable      = VarPtr(m_vtableObj.Functions(0))
m_vtableObj.Functions(0) = AddressOf Handler_QueryInterface
m_vtableObj.Functions(1) = AddressOf Handler_AddRef
m_vtableObj.Functions(2) = AddressOf Handler_Release
m_vtableObj.Functions(3) = pSlot  ← thunk address

The object passed to WebView2 = VarPtr(m_vtableObj)

[When WebView2 calls AddRef/Release]
  obj->vtbl->AddRef(obj)
    → Handler_AddRef(this = obj = VarPtr(m_vtableObj))
                                    ↑ not the thunk address!

[When WebView2 calls Invoke]
  obj->vtbl->Invoke(obj, errorCode, pEnv)
    → the thunk starts (rcx = obj = VarPtr(m_vtableObj))
    → the thunk overwrites rcx with mov rcx, ObjPtr(gHandler)
    → Class7.Handler_Invoke(this = ObjPtr(gHandler), arg1 = old rcx, arg2 = old rdx)
```

In other words:

- **`this` of Invoke (class side)** = `ObjPtr(gHandler)` (the value the thunk overwrote)
- **`this` of AddRef/Release (standard module side)** = `VarPtr(m_vtableObj)` (does not pass through the thunk)

Moreover `m_vtableObj` is a module-level **single instance**. To handle multiple handlers, the vtableObj must also be separated per slot.

The scope of stage 9.1 is "confirm operation with 1 handler + a real reference count", so we deferred the region layout change to the next stage and adopted **plan 1 (the compromise for this stage)**:

- `m_vtableObj` stays a single instance
- AddRef/Release ignore `this` and operate on `gHandler` directly
- The `m_handlers(0..511)` array is only declared, with registration/unregistration only in Acquire/Release (groundwork for stage 9.2)

### Reference count implementation of plan 1

```vba
Private Function HandlerAddRefInternal() As Long
    If gHandler Is Nothing Then Exit Function
    Dim n As Long
    n = gHandler.RefCount + 1
    gHandler.RefCount = n
    Debug.Print "  AddRef  → refCount = " & n
    HandlerAddRefInternal = n
End Function

Private Function HandlerReleaseInternal() As Long
    If gHandler Is Nothing Then Exit Function
    Dim n As Long
    n = gHandler.RefCount - 1
    If n < 0 Then n = 0
    gHandler.RefCount = n
    Debug.Print "  Release → refCount = " & n

    If n = 0 Then
        ' Order:
        '   1. Cut the circular reference with ClearOwner
        '   2. Set m_handlers(idx) to Nothing (the Class is still held by gHandler)
        '   3. Return the slot to the free list with Thunks_ReleaseSlot
        '   4. Only with Set gHandler = Nothing is the Class instance finally released
        Dim pSlot As LongPtr: pSlot = gHandler.Slot
        gHandler.ClearOwner
        Dim idx As Long: idx = SlotIndexFromAddr(pSlot)
        If idx >= 0 Then Set m_handlers(idx) = Nothing
        Thunks_ReleaseSlot pSlot
        Set gHandler = Nothing
        Debug.Print "  Release → refCount = 0 → Slot " & idx & " auto-released"
    End If

    HandlerReleaseInternal = n
End Function
```

`Set gHandler = Nothing` is placed last so as to keep the standard module state consistent if `Class_Terminate` is added in the future.

`QueryInterface` was also changed to call `HandlerAddRefInternal` internally, following the COM convention that "the pointer returned on success has been AddRef'ed".

### How the initial refcount came to be 0 (plan X-1 adopted)

When we took logs with the first implementation attempt (plan X-0, initial refcount = 1), the following pattern was observed:

```
  AddRef  → refCount = 2          ← WebView2 calls AddRef once (1 → 2)
Class7.Invoke[HK_EnvironmentCompleted]
  m_refCount  :              2 
  Release → refCount = 1          ← WebView2 calls Release once (2 → 1)
                                   ← no Release comes after that
```

**WebView2 manages symmetrically with one AddRef and one Release, and has no mechanism to offset the VBA side's "initial reference"**. With the initial refcount left at 1, the refcount stops forever at 1 and automatic release never triggers.

So we switched to plan X-1: "the VBA side holds no initial reference; the refcount is only what WebView2 AddRef'ed":

```
  AddRef  → refCount = 1          ← 0 → 1
Class7.Invoke[HK_EnvironmentCompleted]
  m_refCount  :              1 
  Release → refCount = 0          ← 1 → 0, to the automatic release path
  Release → refCount = 0 → Slot 0 auto-released
```

This is also natural in terms of COM semantics: "once you hand a handler to WebView2, WebView2 AddRefs it itself and takes responsibility".

### Operation check results

Real-machine log (Excel 2024 Win64, Defender enabled):

**Create / Finalize alternating (4 times)**:

```
  AddRef  → refCount = 1
Class7.Invoke[HK_EnvironmentCompleted]
  this        :              2365757850744 
  errorCode   :              0 
  pEnvironment:              55972038278160 
  m_refCount  :              1 
  Release → refCount = 0
  Release → refCount = 0 → Slot 0 auto-released
  ... (repeats with the same pattern)
```

**Consecutive Create (no Finalize, 5 times)**:

```
  AddRef  → refCount = 1
Class7.Invoke[HK_EnvironmentCompleted]
  this        :              2365757850744          ← VBA's internal allocator reuses the ObjPtr
  errorCode   :              0 
  pEnvironment:              55972038273040 
  m_refCount  :              1 
  Release → refCount = 0
  Release → refCount = 0 → Slot 0 auto-released   ← always Slot 0, no leak
  ... (repeats with the same pattern; the slot number is always 0)
```

Facts confirmed:

1. ✅ refcount = 1 after AddRef: plan X-1 is working
2. ✅ `Slot 0 auto-released` appears after Release: the automatic release path runs
3. ✅ `Slot 0` is always reused for consecutive Creates — **the slot number does not keep incrementing = decisive evidence that there is no leak**
4. ✅ `errorCode = 0`: the WebView2 environment was created normally
5. ✅ `this` (= ObjPtr(gHandler)) has the same value across consecutive Creates: reuse of ObjPtr by VBA's internal allocator (already observed in stage 9.0)

### Design points established in stage 9.1

1. **GenericHandler structure**: multiple kinds are consolidated into one class via the `m_kind` enum. A new handler kind needs only an addition to the `HandlerKind` enum and an added `Select Case` branch in `Handler_Invoke`.
2. **Real reference count**: `Class7.m_refCount` is incremented and decremented as the real thing, and at refcount = 0 a 4-step release chain (ClearOwner → clear m_handlers → ReleaseSlot → release gHandler) fires.
3. **WebView2's reference management pattern revealed**: symmetric with one AddRef and one Release. The correct design is for the VBA side to **hold no initial reference (start at refcount = 0)**. This is presumed to hold not only for ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler but also for other Completed-type handlers (to be verified).
4. **Preventing circular references**: always set `m_owner` to Nothing with `ClearOwner`. In stage 9.1 owner = Nothing so the effect is not visible, but it becomes essential from stage 9.2 onward when the upper-level class is introduced.
5. **Stability of vTable slot 7**: as long as the number of Public methods is the same (only `Handler_Invoke`), the slot number is maintained even if the argument types change. It worked correctly at the same slot 7 from Class6 to Class7.

### Facts about COM conventions confirmed during development

- When WebView2 calls through the `obj` pointer, **each call to a vTable method is independent**. Only `Invoke` goes through the thunk, which overwrites `obj` with `ObjPtr(gHandler)` and flows to the class method, while `AddRef` / `Release` go directly to functions on the standard module side.
- Therefore, **the `this` value of AddRef/Release is `VarPtr(m_vtableObj)`, not the thunk address**, and with the current single-vtableObj design there is no means of "telling multiple handlers apart".
- Holding an `m_vtableObj` per slot (plan 2 / plan 3) is the orthodox approach, but that is an issue for stage 9.2.

### Constraints of stage 9.1 (planned to be resolved in stage 9.2)

- **Single `m_vtableObj` instance constraint**: only one handler can be active at a time. Multiple WebView2 objects cannot be managed in parallel.
- **The `m_handlers(0..511)` array is not actually used**: it is registered/unregistered by Acquire/Release, but is not referenced from AddRef/Release (which access `gHandler` directly).
- **There is no path to reverse-lookup idx from the `this` of AddRef/Release**: because there is a single vtableObj, all handler calls come with the same `this`.

### Outlook toward stage 9.2

The central issue is redesigning the region layout to support multiple handlers. We compare the following 2 plans:

- **Plan 2: embed the vtableObj in the slot**
  - Extend SLOT_SIZE from 176 → 224 and reserve a vtableObj region (40 bytes) at the end of each slot
  - The address passed to WebView2 = `pSlot + VTABLE_OBJ_OFFSET`
  - Reverse-lookup idx from the `this` of AddRef/Release with `(this - m_pRegionBase - HEADER_SIZE - VTABLE_OBJ_OFFSET) \ SLOT_SIZE`
  - Since the vtableObj lives inside the VirtualAlloc region, it is naturally consistent with the alive flag mechanism (the whole region is managed by a single VirtualAlloc)

- **Plan 3: hold the vtableObj in a VBA array**
  - Add `Private m_vtableObjs(0..511) As (vtable structure type)`
  - The address passed to WebView2 = `VarPtr(m_vtableObjs(idx))`
  - Whether the VarPtrs of VBA array elements can be obtained contiguously needs to be verified
  - It does not touch the VirtualAlloc region, so the impact on the thunk foundation is small

In stage 9.2 we will compare plan 2 / plan 3 and expect to adopt plan 2 (for consistency with the alive flag mechanism and less dependence on VBA specifications).

In addition, we will replace AddRef / Release with a full-fledged reverse lookup via `m_handlers(idx)`. This will abolish plan 1's direct access to `gHandler`.

``` VBA
''''''''''''''''''''''''''''''''''
' --- Module9_1.bas Stage 9.1 WebView2Thunks (GenericHandler-ization and real reference counting) ---
'
'   Differences from stage 9.0 (former Module9_0.bas):
'     - Introduce the HandlerKind enum and replace Class6 with Class7.cls
'     - Perform the equivalent of AddRef inside Handler_QueryInterface
'     - Change Handler_AddRef / Handler_Release into "real counter operations"
'     - At refcount = 0, automatically chain ClearOwner → clear m_handlers → ReleaseSlot
'     - Newly create the SlotIndexFromAddr helper, shared by Release/ReleaseSlot
'     - Change the signature of Thunks_AcquireSlot to receive the Handler object itself
'     - Newly create the m_handlers(0..511) array, registering/unregistering in Acquire/Release
'       (groundwork for the this → idx reverse lookup in stage 9.2)
'
'   Operation checks at this stage:
'     - Acquire just 1 handler and use the same startup flow as 9.0 (Create → Invoke → Finalize)
'     - Observe with Debug.Print the refcount going up and down through the AddRef / Release that WebView2 calls
'     - For a Completed-type handler, refcount=0 after Invoke → the slot is released automatically
'     - Even repeating Create 10 times without Finalize, slots are reclaimed automatically and the pool is not exhausted
'
'   Constraints of stage 9.1 (resolved in stage 9.2):
'     - Because m_vtableObj is a single instance, the slot idx cannot be
'       reverse-looked-up from the this of AddRef/Release
'     - Therefore AddRef/Release operate on gHandler directly under the premise that
'       "only one gHandler is active at a time"
'     - In stage 9.2 the region layout is redesigned and a vtableObj is embedded
'       in each slot, making multiple handlers possible
'
'   Structure (same as 9.0):
'     vTable[3] = start of the Acquired slot
'                       ↓
'                 clone of the EntryPoint stub (91 bytes)
'                       ↓
'                 after passing the EBMode check, indirect jump to the address at +55
'                       ↓
'                 thunk start (+96 = alive flag check)
'                       ↓
'                 if the flag is alive, the main processing (argument shift + ObjPtr injection)
'                       ↓
'                 Class7.Handler_Invoke (branches on m_kind)
'
'   CopyMemory remains unused. All writes go through MemLongPtr.
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/refTypes/discussions/3
''

Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "This module requires 64-bit VBA (x64)"
#End If

' --- Structures for SafeArray / PointerAccessor ---
Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

' --- WebView2Loader.dll ---
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

' --- Win32 API ---
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long) As LongPtr

Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long

Private Const MEM_COMMIT             As Long = &H1000
Private Const MEM_RESERVE            As Long = &H2000
Private Const MEM_RELEASE            As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const S_OK                   As Long = 0

' --- Slot / region layout constants ---

' Position within the EntryPoint stub where the "address to call next" is embedded as an immediate
Private Const LATE_BIND_OFFSET As Long = 55

' Actual length of the stub (confirmed in stage 8.5; 91 bytes including the C2 00 00 at +88)
Private Const STUB_LEN As Long = 91

' Length of the thunk (alive flag check 18 + existing 56 = 74 bytes)
Private Const THUNK_LEN As Long = 74

' Allocation size per slot (rounded to a multiple of 8, 176 bytes)
'   +0..+90      stub clone (91 bytes)
'   +91..+95     padding
'   +96..+169    thunk (74 bytes)
'   +170..+175   padding (8-byte alignment)
Private Const SLOT_SIZE      As Long = 176
Private Const THUNK_OFFSET   As Long = 96

' Number of slots (fixed)
Private Const SLOT_COUNT     As Long = 512

' Header size at the start of the region (alive flag + future reservation)
'   +0          : alive flag (1 byte, but handled in 8-byte units)
'   +1..+63     : reserved (for a future sentinel flag-clearing thunk, etc.)
Private Const HEADER_SIZE    As Long = 64

' Size of the whole region
Private Const REGION_SIZE    As Long = HEADER_SIZE + SLOT_SIZE * SLOT_COUNT


' ============================================================
' HandlerKind enum (referenceable from upper-level modules)
'   Identifies the kind of WebView2 callback. It is stored in Class7.cls's
'   m_kind and used for branching inside Handler_Invoke.
' ============================================================
Public Enum HandlerKind
    HK_None = 0
    HK_EnvironmentCompleted = 1
    HK_ControllerCompleted = 2
    HK_NavigationStarting = 3
    HK_NavigationCompleted = 4
    HK_WebMessageReceived = 5
    HK_DocumentTitleChanged = 6
End Enum


' --- Structure holding a set of vTable + pVTable ---
'   (renamed from the old "WebView2Handler" type, to avoid a clash with the new class name WebView2Handler)
Private Type WebView2VTable
    pVTable      As LongPtr
    Functions(3) As LongPtr
End Type
Private m_vtableObj As WebView2VTable

' --- Single handler created at startup (the old gHandler As Class6 changed to Class7) ---
Private gHandler As Class7

' --- Slot pool state ---
Private m_pRegionBase As LongPtr        ' Start of the region allocated with VirtualAlloc
Private m_freeHead    As Long           ' Head index of the free list (-1 = empty)
Private m_freeNext()  As Long           ' freeNext(i) = the next free slot after slot i
Private m_inUse       As Long           ' Number of slots currently in use (for debugging)

' --- Correspondence table: slot index → Handler object ---
'   In stage 9.1 Acquire/Release only register/unregister, and AddRef/Release do not
'   refer to this array (they use gHandler directly).
'   Only the registration is implemented ahead of time as groundwork for replacing it in stage 9.2
'   with the form "reverse-lookup idx from this → operate on m_handlers(idx)".
Private m_handlers(0 To SLOT_COUNT - 1) As Class7

Private Const THUNK_BUF_SIZE As Long = 80     ' 74 rounded up to a multiple of 8


' ============================================================
' Source of the EntryPoint stub
'   The byte sequence of this stub is copied into the VirtualAlloc region.
'   It must be called once at startup for Force-Compile.
' ============================================================
Private Sub EntryPoint(): End Sub


' ============================================================
' Public entry (equivalent of stage 8.9's CreateWebView2Environment)
'
'   Call Thunks_Init at startup, then take 1 slot with Acquire and
'   request creation of the WebView2 environment just as in stage 8.9.
'
'   Changes in stage 9.1:
'     - The type of gHandler changed from Class6 to Class7
'     - After Acquire, set up kind / owner / pSlot with Handler.Init
'     - Registration into m_handlers(idx) is done inside Thunks_AcquireSlot
' ============================================================
Public Sub CreateWebView2Environment()

    ' --- 1. Initialize the slot pool (if not yet) ---
    If m_pRegionBase = 0 Then
        If Not Thunks_Init() Then
            MsgBox "Thunks_Init failed", vbCritical
            Exit Sub
        End If
    End If

    ' --- 2. Create the handler object ---
    Set gHandler = New Class7

    Dim pHandlerInvoke As LongPtr
    pHandlerInvoke = GetClassMethodAddrAtFixedSlot(gHandler, 7)
    If pHandlerInvoke = 0 Then
        MsgBox "Failed to obtain the address of Class7.Handler_Invoke", vbCritical
        Exit Sub
    End If

    ' --- 3. Take 1 slot, write the thunk + register in m_handlers ---
    Dim pSlot As LongPtr
    pSlot = Thunks_AcquireSlot(gHandler, ObjPtr(gHandler), pHandlerInvoke)
    If pSlot = 0 Then
        MsgBox "Failed to acquire a slot (pool exhausted or initialization failed)", vbCritical
        Exit Sub
    End If

    ' Set the kind, owner and pSlot on the Handler itself
    '   At this stage owner can still be Nothing (the upper-level class comes in stage 9.2 onward)
    gHandler.Init HK_EnvironmentCompleted, Nothing, pSlot

    ' --- 4. Build the vTable ---
    Static pa As PointerAccessor
    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = PtrSize
        pa.sa.cLocks = 1
        MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
    End If

    m_vtableObj.pVTable = VarPtr(m_vtableObj.Functions(0))

    pa.sa.pvData = VarPtr(m_vtableObj.Functions(0))
    pa.sa.rgsabound0.cElements = 4

    pa.arr(0) = GetAddr(AddressOf Handler_QueryInterface)
    pa.arr(1) = GetAddr(AddressOf Handler_AddRef)
    pa.arr(2) = GetAddr(AddressOf Handler_Release)
    pa.arr(3) = pSlot                    ' ← the start of the stub clone is the entrance

    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = 0

    ' --- 5. Create the WebView2 environment ---
    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, VarPtr(m_vtableObj))

    If hr <> S_OK Then
        MsgBox "Failed to initialize WebView2. HRESULT: &H" & Hex(hr), vbCritical
    End If

    DoEvents
End Sub


' ============================================================
' Shutdown processing (formerly FinalizeWebView2Environment)
'
'   By setting the alive flag to 0 before VirtualFree,
'   even if there are unreleased handlers (= even if WebView2 is still
'   holding the vTable address), callbacks are safely cut off with S_OK.
' ============================================================
Public Sub FinalizeWebView2Environment()
    Thunks_Shutdown
    Set gHandler = Nothing
End Sub


' ============================================================
' Thunks_Init
'   Allocate the region with VirtualAlloc, set the alive flag in the header,
'   copy the stub into all slots, and initialize the free list.
'
'   Return value: success = True, failure = False
' ============================================================
Public Function Thunks_Init() As Boolean

    If m_pRegionBase <> 0 Then
        Thunks_Init = True
        Exit Function
    End If

    ' Force-Compile EntryPoint (materialize the source of the stub)
    EntryPoint

    ' Source address of the EntryPoint stub
    Dim pStubSrc As LongPtr
    pStubSrc = VBA.Int(AddressOf EntryPoint)
    If pStubSrc = 0 Then Exit Function

    ' Allocate the region in bulk
    m_pRegionBase = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If m_pRegionBase = 0 Then Exit Function

    ' Zero-initialize the header (VirtualAlloc is already zero-cleared by nature, but just in case)
    Dim k As Long
    For k = 0 To HEADER_SIZE - 1 Step 8
        MemLongPtr(m_pRegionBase + k) = 0^
    Next k

    ' Set the alive flag (1 byte at the start of the region = 1)
    ' Since it is written in 8-byte units, put 0x01 so that only the lowest 1 byte becomes 1
    MemLongPtr(m_pRegionBase) = 1^

    ' Copy the stub into all slots (96 bytes each; the padding at +91..+95 is harmless)
    Dim i As Long, pSlot As LongPtr
    For i = 0 To SLOT_COUNT - 1
        pSlot = SlotAddrAt(i)
        For k = 0 To 95 Step 8
            MemLongPtr(pSlot + k) = ReadLongPtr(pStubSrc + k)
        Next k
    Next i

    ' Initialize the free list
    '   m_freeNext(i) = i + 1 (the last is -1)
    '   m_freeHead    = 0
    ReDim m_freeNext(0 To SLOT_COUNT - 1)
    For i = 0 To SLOT_COUNT - 2
        m_freeNext(i) = i + 1
    Next i
    m_freeNext(SLOT_COUNT - 1) = -1
    m_freeHead = 0
    m_inUse = 0

    ' Clear the Handler correspondence table entirely (it is a fixed array so Nothing at declaration, but just in case)
    For i = 0 To SLOT_COUNT - 1
        Set m_handlers(i) = Nothing
    Next i

    Thunks_Init = True
End Function


' ============================================================
' Thunks_AcquireSlot
'   Take 1 free slot from the free list, write into that slot's
'   thunk region the machine code that "injects pSelfObj into Me and calls pTargetFunc",
'   and write the thunk's address into +55.
'
'   Change in stage 9.1: receives the handler object as the first argument
'   and registers it in m_handlers(idx). This lets Handler_AddRef / Release
'   reverse-lookup idx from the this address and reach the Handler object.
'
'   Return value: start address of the slot (= the value to put in vTable[3]), 0 on failure
' ============================================================
Public Function Thunks_AcquireSlot( _
    ByVal handler As Class7, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr) As LongPtr

    If m_pRegionBase = 0 Then Exit Function
    If m_freeHead < 0 Then Exit Function     ' Pool exhausted
    If handler Is Nothing Then Exit Function

    ' Take one from the free list
    Dim idx As Long
    idx = m_freeHead
    m_freeHead = m_freeNext(idx)
    m_freeNext(idx) = -1                      ' -1 = in use

    Dim pSlot As LongPtr
    pSlot = SlotAddrAt(idx)

    ' Write the thunk (74 bytes from offset +96)
    WriteThunkMachineCode pSlot + THUNK_OFFSET, pSelfObj, pTargetFunc, m_pRegionBase

    ' Write the thunk's address into +55 of the stub clone
    MemLongPtr(pSlot + LATE_BIND_OFFSET) = pSlot + THUNK_OFFSET

    ' Associate the Handler object with idx
    Set m_handlers(idx) = handler

    m_inUse = m_inUse + 1
    Thunks_AcquireSlot = pSlot
End Function


' ============================================================
' Thunks_ReleaseSlot
'   Return the specified slot to the free list.
'   Also clears m_handlers(idx).
'   The contents inside the slot are not zero-cleared; they are overwritten at the next Acquire.
'
'   Note: WebView2 may still be holding vTable[3] (= this slot),
'         so the caller must confirm that it is a timing at which
'         Release may be called.
'         From stage 9.1 onward it is called automatically from Handler_Release
'         at refcount=0, so there is basically no need to call it directly from outside.
' ============================================================
Public Sub Thunks_ReleaseSlot(ByVal pSlot As LongPtr)
    If m_pRegionBase = 0 Then Exit Sub
    If pSlot = 0 Then Exit Sub

    Dim idx As Long
    idx = SlotIndexFromAddr(pSlot)
    If idx < 0 Then Exit Sub
    If m_freeNext(idx) <> -1 Then Exit Sub           ' Already free = double release

    ' Also clear the Handler correspondence table
    Set m_handlers(idx) = Nothing

    ' Return it to the head of the free list
    m_freeNext(idx) = m_freeHead
    m_freeHead = idx

    m_inUse = m_inUse - 1
End Sub


' ============================================================
' SlotIndexFromAddr (internal helper)
'   Computes the slot index backward from the thunk's address
'   (= the slot's start address, the this that WebView2 calls).
'   Returns -1 for an invalid address (outside the region, unaligned).
' ============================================================
Private Function SlotIndexFromAddr(ByVal pSlot As LongPtr) As Long
    SlotIndexFromAddr = -1
    If m_pRegionBase = 0 Then Exit Function

    Dim offset As LongPtr
    offset = pSlot - (m_pRegionBase + HEADER_SIZE)
    If offset < 0 Then Exit Function
    If (offset Mod SLOT_SIZE) <> 0 Then Exit Function

    Dim idx As Long
    idx = CLng(offset \ SLOT_SIZE)
    If idx < 0 Or idx >= SLOT_COUNT Then Exit Function

    SlotIndexFromAddr = idx
End Function


' ============================================================
' Thunks_Shutdown
'   Set the alive flag to 0, then free the region with VirtualFree.
'   After the call, even if a thunk runs it returns S_OK immediately
'   (however, after VirtualFree the region itself is gone, so setting the flag
'   is meaningful only for "the instant until VirtualFree completes").
' ============================================================
Public Sub Thunks_Shutdown()
    If m_pRegionBase = 0 Then Exit Sub

    ' First, set the alive flag to 0
    '   This way, even if a callback arrives from WebView2 addressed to an unreleased handler,
    '   it returns S_OK immediately at the flag check at the start of the thunk.
    MemLongPtr(m_pRegionBase) = 0^

    ' Cut the owner of all Handlers and release the references
    '   Always call ClearOwner to prevent circular references.
    Dim i As Long
    For i = 0 To SLOT_COUNT - 1
        If Not (m_handlers(i) Is Nothing) Then
            m_handlers(i).ClearOwner
            Set m_handlers(i) = Nothing
        End If
    Next i

    ' Free the region
    VirtualFree m_pRegionBase, 0, MEM_RELEASE

    m_pRegionBase = 0
    m_freeHead = -1
    m_inUse = 0
    Erase m_freeNext
End Sub


' ============================================================
' Slot index → slot start address
' ============================================================
Private Function SlotAddrAt(ByVal idx As Long) As LongPtr
    SlotAddrAt = m_pRegionBase + HEADER_SIZE + CLngLng(idx) * SLOT_SIZE
End Function


' ============================================================
' Write the thunk's machine code to the specified address (effective length 74 bytes)
'
'   Layout:
'     +0..+17  : alive flag check (18 bytes)
'     +18..+73 : existing thunk body (56 bytes, same as stages 6 to 8)
'
'   Alive flag check:
'     mov  rax, pAliveFlag      [10 bytes: 48 B8 + 8 byte imm]
'     cmp  byte ptr [rax], 1    [3 bytes:  80 38 01]
'     je   alive (+3)           [2 bytes:  74 03]
'     xor  eax, eax             [2 bytes:  33 C0]   (S_OK)
'     ret                       [1 byte:   C3]
'   alive:
'     (from here on, the existing 56-byte thunk)
'
'   Implementation note:
'     The byte array b is declared with 80 bytes (= a multiple of 8), and the last 6 bytes are
'     treated as zero padding. The write destination, slot +96..+175 (= 80 bytes' worth),
'     is originally padding/unused area, so writing 80 bytes is harmless.
'     This lets us write plainly with the 8-byte-unit loop of MemLongPtr.
' ============================================================

Private Sub WriteThunkMachineCode( _
    ByVal addr As LongPtr, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr, _
    ByVal pRegionBase As LongPtr)

    Dim b(0 To THUNK_BUF_SIZE - 1) As Byte, i As Long
    i = 0

    ' --- Alive flag check (18 bytes) ---
    ' mov rax, imm64   (= pAliveFlag = pRegionBase + 0)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pRegionBase: i = i + 8
    ' cmp byte ptr [rax], 1
    b(i) = &H80: b(i + 1) = &H38: b(i + 2) = &H1: i = i + 3
    ' je +3 (= the alive label, skipping the 3 bytes of xor + ret)
    b(i) = &H74: b(i + 1) = &H3: i = i + 2
    ' xor eax, eax
    b(i) = &H33: b(i + 1) = &HC0: i = i + 2
    ' ret
    b(i) = &HC3: i = i + 1

    ' --- From here on, the same 56-byte thunk body as stages 6 to 8 ---

    ' sub rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4
    ' mov r9, r8
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    ' mov rcx, imm64   (ObjPtr(gHandler))
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pSelfObj: i = i + 8
    ' lea rax, [rsp+0x28]
    b(i) = &H48: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5
    ' mov [rsp+0x20], rax
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H20: i = i + 5
    ' mov rax, imm64   (pTargetFunc)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pTargetFunc: i = i + 8
    ' call rax
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
    ' mov eax, [rsp+0x28]
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H28: i = i + 4
    ' add rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4
    ' ret
    b(i) = &HC3: i = i + 1
    ' int3 / int3 (padding)
    b(i) = &HCC: i = i + 1
    b(i) = &HCC: i = i + 1

    ' --- Write 80 bytes in 8-byte units (the last 6 bytes are 0 padding, harmless) ---
    Dim k As Long
    For k = 0 To THUNK_BUF_SIZE - 1 Step 8
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k
End Sub


' ============================================================
' Take 8 bytes (LongPtr) from the given offset of a byte array
' ============================================================
Private Function ReadLongPtrFromBytes(ByRef b() As Byte, ByVal offset As Long) As LongPtr
    ReadLongPtrFromBytes = ReadLongPtr(VarPtr(b(offset)))
End Function


' ============================================================
' Get the function address of a fixed slot from the class's vTable
' ============================================================
Private Function GetClassMethodAddrAtFixedSlot( _
    ByVal cls As Object, ByVal slotIndex As Long) As LongPtr

    If cls Is Nothing Then Exit Function
    Dim pObj As LongPtr: pObj = ObjPtr(cls)
    If pObj = 0 Then Exit Function

    Dim pVTable As LongPtr
    pVTable = ReadLongPtr(pObj)
    If pVTable = 0 Then Exit Function

    GetClassMethodAddrAtFixedSlot = ReadLongPtr(pVTable + slotIndex * PtrSize)
End Function


' ============================================================
' IUnknown stubs (standard module)
'
'   Changes in stage 9.1:
'     - Perform the equivalent of AddRef inside QueryInterface (increase refcount by 1)
'     - Change AddRef / Release into real reference count operations
'     - When Release brings refcount to 0, automatically chain ClearOwner → clear m_handlers
'       → Thunks_ReleaseSlot
'
'   Constraints of stage 9.1 (planned to be resolved in stage 9.2):
'     - There is no path to reverse-lookup the slot idx from this (= obj = VarPtr(m_vtableObj))
'       (m_vtableObj is a single instance)
'     - Therefore, at the current stage, under the premise that "only one gHandler
'       is active at a time", we ignore this and operate on gHandler directly
'     - The this → idx reverse lookup for multiple-handler support will be introduced in stage 9.2
'       by redesigning the region layout (e.g. the plan of embedding the vtableObj
'       inside the slot)
'
'   The IID check is planned to be introduced from stage 9.5 onward. For now all IIDs are accepted.
' ============================================================
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

Private Function Handler_QueryInterface( _
    ByVal this As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long

    ' For now it accepts all IIDs and returns itself (the IID check is a future task)
    ppvObject = this

    ' The premise is that the pointer QueryInterface returns on success has been AddRef'ed,
    ' so we perform the equivalent of AddRef here
    HandlerAddRefInternal

    Handler_QueryInterface = S_OK
End Function

Private Function Handler_AddRef(ByVal this As LongPtr) As Long
    Handler_AddRef = HandlerAddRefInternal()
End Function

Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = HandlerReleaseInternal()
End Function


' ============================================================
' HandlerAddRefInternal / HandlerReleaseInternal
'   Plan 1: stage 9.1 ignores this and operates on gHandler directly.
'        It will be replaced with the this → idx reverse lookup in stage 9.2.
'
'   Return value: the refcount after the operation (following COM convention)
' ============================================================
Private Function HandlerAddRefInternal() As Long
    If gHandler Is Nothing Then
        HandlerAddRefInternal = 0
        Exit Function
    End If

    Dim n As Long
    n = gHandler.RefCount + 1
    gHandler.RefCount = n
    Debug.Print "  AddRef  → refCount = " & n
    HandlerAddRefInternal = n
End Function

Private Function HandlerReleaseInternal() As Long
    If gHandler Is Nothing Then
        HandlerReleaseInternal = 0
        Exit Function
    End If

    Dim n As Long
    n = gHandler.RefCount - 1
    If n < 0 Then n = 0                ' Defense (should not happen in principle)
    gHandler.RefCount = n
    Debug.Print "  Release → refCount = " & n

    If n = 0 Then
        ' Since refcount has become 0, cut the circular reference and release the slot.
        ' The order matters:
        '   1. Cut the circular reference with ClearOwner
        '   2. Set m_handlers(idx) to Nothing (the Class is still held by gHandler)
        '   3. Return the slot to the free list with Thunks_ReleaseSlot
        '   4. Only with Set gHandler = Nothing is the Class instance finally released
        '
        ' Step 4 goes last so as to keep consistency in case the standard module state is read
        ' inside Class_Terminate (which may be added in the future).

        Dim pSlot As LongPtr
        pSlot = gHandler.Slot

        ' 1. Cut the circular reference (m_owner = Nothing)
        gHandler.ClearOwner

        ' 2. Reverse-lookup the slot index and remove it from m_handlers
        Dim idx As Long
        idx = SlotIndexFromAddr(pSlot)
        If idx >= 0 Then
            Set m_handlers(idx) = Nothing
        End If

        ' 3. Return the slot to the free list
        Thunks_ReleaseSlot pSlot

        ' 4. Cut the gHandler reference (at this point the Class instance is released)
        Set gHandler = Nothing

        Debug.Print "  Release → refCount = 0 → Slot " & idx & " auto-released"
    End If

    HandlerReleaseInternal = n
End Function


' ============================================================
' Memory-operation primitives via PointerAccessor
' ============================================================
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub
```

``` VBA
''''''''''''''''''''''''''''''''''
' --- Class7.cls Stage 9.1 WebView2Handler---
'
'   A general-purpose handler class replacing the old Class6.
'   It handles multiple kinds via m_kind, and the only Public method is Handler_Invoke
'   (pinned at vTable slot 7). As up to stage 9.0, it is called directly from WebView2.
'
'   State held:
'     m_kind     : handler kind (HK_EnvironmentCompleted, etc.)
'     m_owner    : upper-level class (can still be Nothing at this stage)
'     m_pSlot    : start address of the slot where the thunk is placed
'                  (= the this address that WebView2 calls)
'     m_refCount : reference count (starts from 1 in Init)
'
'   Countermeasure against circular references:
'     When refcount = 0, the foundation side (WebView2Thunks.bas) calls
'     ClearOwner to cut m_owner.
'     Because the upper-level class holds the WebView2Handler and the WebView2Handler holds
'     the upper-level class via m_owner, it leaks unless explicitly cut.
''''''''''''''''''''''''''''''''''
Option Explicit

#If Win64 Then
    ' OK
#Else
    #Error "This class requires 64-bit VBA (x64)"
#End If

Private m_kind     As HandlerKind
Private m_owner    As Object
Private m_pSlot    As LongPtr
Private m_refCount As Long


' ============================================================
' Initialization (assumed to be called from the foundation's Thunks_AcquireSlot)
'
'   refcount starts from 0 (plan X-1 adopted):
'     - Following COM semantics, the model is that once a handler is handed to WebView2,
'       WebView2 AddRefs it itself and takes responsibility. The VBA side holds no reference.
'     - It has been confirmed on a real machine that WebView2 calls AddRef inside
'       CreateCoreWebView2EnvironmentWithOptions (refcount: 0 → 1).
'     - After Invoke, WebView2 calls Release and refcount = 0, and
'       the automatic release path of HandlerReleaseInternal runs.
'
'   This initial value was decided because the real-machine logs (the operation check of stage 9.1)
'   revealed that "one AddRef + one Release" are managed symmetrically
'   on the WebView2 side.
'   If plan X-1 is not chosen, refcount stays at 1 and the slot leaks.
' ============================================================
Friend Sub Init( _
    ByVal Kind As HandlerKind, _
    ByVal owner As Object, _
    ByVal pSlot As LongPtr)

    m_kind = Kind
    Set m_owner = owner
    m_pSlot = pSlot
    m_refCount = 0
End Sub


' ============================================================
' Property group (for operation from the foundation side)
' ============================================================
Friend Property Get RefCount() As Long
    RefCount = m_refCount
End Property

Friend Property Let RefCount(ByVal v As Long)
    m_refCount = v
End Property

Friend Property Get Slot() As LongPtr
    Slot = m_pSlot
End Property

Friend Property Get Kind() As HandlerKind
    Kind = m_kind
End Property


' ============================================================
' ClearOwner
'   Called from the foundation side when refcount = 0.
'   Sets m_owner to Nothing to cut the circular reference.
' ============================================================
Friend Sub ClearOwner()
    Set m_owner = Nothing
End Sub


' ============================================================
' Handler_Invoke (pinned at vTable slot 7)
'
'   The only entry that receives callbacks from WebView2.
'   The thunk injects ObjPtr(this) into Me, shifts the arguments one at a time,
'   and reaches this method.
'
'   Argument specification (settled in stage 9.1):
'     this : the ObjPtr of itself that the thunk injects into Me (not used)
'     arg1 : the 1st argument WebView2 put in rcx (received as LongPtr)
'     arg2 : the 2nd argument WebView2 put in rdx (received as LongPtr)
'
'   Interpretation per m_kind:
'     HK_EnvironmentCompleted: arg1 = errorCode (lower 32 bits), arg2 = pEnvironment
'     HK_ControllerCompleted:  arg1 = errorCode (lower 32 bits), arg2 = pController
'     HK_NavigationStarting:   arg1 = sender, arg2 = args
'     HK_NavigationCompleted:  arg1 = sender, arg2 = args
'     HK_WebMessageReceived:   arg1 = sender, arg2 = args
'     HK_DocumentTitleChanged: arg1 = sender, arg2 = args
'
'   In stage 9.1 only HK_EnvironmentCompleted was actually confirmed to work.
'   The other kinds just print "not supported" with Debug.Print.
' ============================================================
Public Function Handler_Invoke( _
    ByVal this As LongPtr, _
    ByVal arg1 As LongPtr, _
    ByVal arg2 As LongPtr) As Long

    Select Case m_kind
        Case HK_EnvironmentCompleted
            Dim errorCode As Long
            errorCode = LongPtrLowDword(arg1)
            Debug.Print "Class7.Invoke[HK_EnvironmentCompleted]"
            Debug.Print "  this        :", this
            Debug.Print "  errorCode   :", errorCode
            Debug.Print "  pEnvironment:", arg2
            Debug.Print "  m_refCount  :", m_refCount

        Case HK_ControllerCompleted
            Debug.Print "Class7.Invoke[HK_ControllerCompleted] (not supported)"

        Case HK_NavigationStarting
            Debug.Print "Class7.Invoke[HK_NavigationStarting] (not supported)"

        Case HK_NavigationCompleted
            Debug.Print "Class7.Invoke[HK_NavigationCompleted] (not supported)"

        Case HK_WebMessageReceived
            Debug.Print "Class7.Invoke[HK_WebMessageReceived] (not supported)"

        Case HK_DocumentTitleChanged
            Debug.Print "Class7.Invoke[HK_DocumentTitleChanged] (not supported)"

        Case Else
            Debug.Print "Class7.Invoke[unknown kind=" & m_kind & "]"
    End Select

    Handler_Invoke = 0   ' S_OK
End Function


' ============================================================
' Extract the lower 32 bits of a LongPtr as a signed Long
'
'   Under the x64 calling convention, when a 32-bit argument (HRESULT, etc.) is passed,
'   the upper 32 bits may be left undefined in rdx and so on.
'   The thunk shifts in 64 bits like "mov rdx, rcx",
'   so receiving it as a LongPtr as is leaves garbage in the upper 32 bits.
'
'   Extract the lower 32 bits and, if 0x80000000 or above, fold it into a negative Long
'   (CLng overflows outside the Long range, so we handle it ourselves).
' ============================================================
Private Function LongPtrLowDword(ByVal v As LongPtr) As Long
    Dim u As LongLong
    u = v And &HFFFFFFFF^
    If u >= &H80000000^ Then
        LongPtrLowDword = CLng(u - &H100000000^)
    Else
        LongPtrLowDword = CLng(u)
    End If
End Function
```


## Stage 9.2 — Multiple handler support + embedded vtableObj + full implementation of the this reverse lookup

The stage that resolves the biggest constraint left by stage 9.1, "because `m_vtableObj` is a single instance, idx cannot be reverse-looked-up from the `this` of AddRef / Release". We embedded a vtableObj region (40 bytes) inside each slot (plan 2), completing a real COM handler mechanism in which AddRef / Release reverse-lookup `m_handlers(idx)` from `this`. Operation was confirmed in two lines: (a) a regression compatible with stage 9.1 (Create / Finalize / consecutive Create with 1 handler), and (b) the newly created pure reverse-lookup unit test `Test_TwoHandlers_Mock`.

### Differences from stage 9.1

| Item | Stage 9.1 | Stage 9.2 |
|---|---|---|
| `SLOT_SIZE` | 176 | **224** (vtableObj 40 bytes + 8 bytes of spare padding added) |
| `VTABLE_OBJ_OFFSET` | None | **176** (= the same value as the old SLOT_SIZE, easy to remember) |
| Whole region | 90,176 bytes (≒88 KB) | **114,752 bytes (≒112 KB)** |
| `m_vtableObj` module variable | Single instance | **Abolished entirely** |
| Address passed to WebView2 | `VarPtr(m_vtableObj)` | **`pSlot + VTABLE_OBJ_OFFSET`** |
| Obtaining function addresses | `AddressOf` at each Acquire | **Once in Thunks_Init**, saved in `m_pHandler_QI` / `_AddRef` / `_Release` |
| vtableObj initialization | Built at Acquire | **All 512 slots at once in Thunks_Init**; Acquire does not touch it |
| AddRef / Release | Direct operation on `gHandler` (plan 1 compromise) | **`this` → `SlotIndexFromVTableObjAddr` → `m_handlers(idx)`** |
| `Handler_QueryInterface` | Call `HandlerAddRefInternal()` with no arguments | Reverse lookup via **`HandlerAddRefInternal(this)`** |
| Operation check tests | Create / Finalize only | **`Test_TwoHandlers_Mock` added** (independence check of 2 handlers) |

### Region layout (total 114,752 bytes ≒ 112 KB)

```
+0..+7        : alive flag (LongPtr, 8 bytes; only the lowest 1 byte is used)
+8..+63       : header reserved area
+64           : slot 0 starts (224 bytes)
+64+224       : slot 1 starts
   ...
+64+224×511   : slot 511 starts

Each slot (224 bytes):
  +0..+90       : EntryPoint stub clone (91 bytes)
  +91..+95      : padding (5 bytes)
  +96..+169     : thunk (74 bytes)
  +170..+175    : padding (6 bytes, for 8-byte alignment)
  +176..+183    : vtableObj.pVTable (= points to +184)         ← VTABLE_OBJ_OFFSET = 176
  +184..+215    : vtableObj.Functions(0..3) (32 bytes)
  +216..+223    : spare padding (8 bytes)
```

### Contents of the vtableObj (+176..+215 of each slot)

```
+176..+183 : pVTable      = +184 (points to its own Functions(0))
+184..+191 : Functions(0) = m_pHandler_QI       (= AddressOf Handler_QueryInterface)
+192..+199 : Functions(1) = m_pHandler_AddRef   (= AddressOf Handler_AddRef)
+200..+207 : Functions(2) = m_pHandler_Release  (= AddressOf Handler_Release)
+208..+215 : Functions(3) = pSlot               (= the start of its own stub clone)
```

`Functions(0..2)` are common function addresses across all 512 slots. Only `Functions(3)` is slot-specific (= the starting point that reaches `Class7.Handler_Invoke` via the thunk inside that slot). They are written for all 512 slots at once inside `Thunks_Init`, and Acquire / Release never touch the vtableObj.

### Why plan 2 was adopted (compared with plan 3 = the VBA array scheme)

| Aspect | Plan 2 (embedded in the slot) | Plan 3 (VBA array) |
|---|---|---|
| Consistency with the alive flag mechanism | ◎ (inside the same VirtualAlloc region) | △ (becomes separately managed) |
| Reverse-lookup arithmetic | One line: `(this - base - HEADER - OFFSET) \ SLOT_SIZE` | Requires a guarantee that the VarPtrs of array elements are contiguous |
| Dependence on VBA specifications | Minimal (just writes to a fixed-address region) | Large (unclear whether contiguity of VarPtr for UDT array elements is specified) |
| Implementation complexity | Simple | Equivalent |
| Memory consumption | +24 KB (88 → 112) | Almost the same |

We chose plan 2 for the cleanliness of "everything managed by a single VirtualAlloc region" and the safety of not depending on undocumented VBA behavior.

### Overall picture of the call paths (final form in stage 9.2)

```
[When WebView2 calls AddRef / Release]
  obj = pSlot + VTABLE_OBJ_OFFSET  ← the this value WebView2 holds
  obj->vtbl = *(LongPtr*)obj = pSlot + VTABLE_OBJ_OFFSET + 8
  obj->vtbl[1] = m_pHandler_AddRef  (Functions(1))
                       ↓
  Handler_AddRef(this = obj)
                       ↓
  HandlerAddRefInternal(this)
                       ↓
  idx = SlotIndexFromVTableObjAddr(this)
      = (this - m_pRegionBase - HEADER_SIZE - VTABLE_OBJ_OFFSET) \ SLOT_SIZE
                       ↓
  m_handlers(idx).RefCount += 1

[When WebView2 calls Invoke]
  obj->vtbl[3] = pSlot  (Functions(3) = start of the stub clone)
                       ↓
  Clone of the EntryPoint stub (91 bytes)
                       ↓
  After passing the EBMode check, indirect jump to the address at +55
                       ↓
  Thunk start (+96 = alive flag check)
                       ↓
  If the flag is alive, the main processing (argument shift + ObjPtr injection)
                       ↓
  Class7.Handler_Invoke (branches on m_kind)
```

`obj` (the COM object as seen from WebView2) is **the start of the vtableObj**. AddRef / Release do not go through the thunk but jump directly to `m_pHandler_AddRef` / `_Release` (standard module functions). Only Invoke goes through the two-step jump stub → thunk via `Functions(3) = pSlot` to reach Class7. Note that the `this` passed to AddRef / Release / QueryInterface is `pSlot + VTABLE_OBJ_OFFSET`, while the `this` the thunk injects into Invoke is `ObjPtr(gHandler)`, so the **values differ**.

### Significance of Test_TwoHandlers_Mock

Launching two real WebView2s in parallel would involve implementing the ControllerCompleted handler, so it is outside the scope of stage 9.2. Instead we introduced a pure unit test that "Acquires 2 handlers, fires AddRef / Release separately from each one's vtableObj address, and confirms that the idx reverse lookup dispatches them correctly".

Items that can be confirmed:
1. The 2 handlers obtain different idx values (e.g. 0 and 1)
2. An AddRef to h1 increases only h1's m_refCount (h2 unchanged)
3. An AddRef to h2 increases only h2's m_refCount (h1 unchanged)
4. Automatic release fires at the moment each independently reaches refcount=0
5. In the end `m_inUse = 0` and both slots return to the free list
6. AddRef / Release to a released slot passes safely as a no-op (return value 0)

Since it does not depend on WebView2, its startup cost is low and it can also be used for regression checks during development.

### Operation check results

Real-machine log (Excel 2024 Win64, Defender enabled):

**Test_TwoHandlers_Mock (key points)**:

```
h1 -> idx 0, pSlot 2364644130880
h2 -> idx 1, pSlot 2364644131104   ← the difference 224 = SLOT_SIZE is correct
m_inUse = 2 (expected 2)

--- AddRef sequence ---
  AddRef [idx=0] → refCount = 1
  AddRef [idx=0] → refCount = 2
  AddRef [idx=1] → refCount = 1
h1.RefCount = 2 (expected 2)
h2.RefCount = 1 (expected 1)
[OK] AddRef dispatching is as expected

--- Release sequence ---
  Release [idx=0] → refCount = 1
  Release [idx=0] → refCount = 0
  Release [idx=0] → refCount = 0 → Slot 0 auto-released
  ...                                  ← h2 stays unchanged at refCount=1 during this time
  Release [idx=1] → refCount = 0
  Release [idx=1] → refCount = 0 → Slot 1 auto-released

--- Post-release safety check ---
AddRef return value to the released pObj1 = 0 (expected 0)
Release return value to the released pObj2 = 0 (expected 0)

m_inUse = 0 (expected 0)
```

**CreateWebView2Environment run consecutively (7 times)**:

```
  AddRef [idx=0] → refCount = 1
Class7.Invoke[HK_EnvironmentCompleted]
  this        :              2364645376240   ← the same value from the 2nd time onward (ObjPtr reuse)
  errorCode   :              0 
  pEnvironment:              55972038486032  ← different every time (newly generated by WebView2)
  m_refCount  :              1 
  Release [idx=0] → refCount = 0
  Release [idx=0] → refCount = 0 → Slot 0 auto-released
   ... (the same pattern repeats 7 times; Slot 0 is reused every time)
```

Facts confirmed:

1. ✅ The 2 handlers obtained different idx values (`idx 0` and `idx 1`)
2. ✅ AddRef dispatching is completely independent (2 AddRefs to h1 give +2 for h1 only, h2 unchanged)
3. ✅ The automatic release path of Release runs independently (`[idx=0] auto-released` and `[idx=1] auto-released` at separate timings)
4. ✅ Operations on a released slot are safely no-ops
5. ✅ The stage 9.1 regression (Slot 0 is always reused on consecutive Creates) still passes
6. ✅ Compared with stage 9.1, the log now shows `[idx=N]`, improving the visibility that the reverse lookup is working

### Design points established in stage 9.2

1. **vtableObj embedded in the slot (plan 2)**: `SLOT_SIZE = 224`, `VTABLE_OBJ_OFFSET = 176`. It is integrated into the same VirtualAlloc region as the alive flag mechanism, unifying resource management.
2. **Bulk initialization in Thunks_Init**: the vtableObj of all 512 slots is written at startup. Acquire does not touch the vtableObj and only writes the thunk code and registers `m_handlers(idx)`. The increase in write count (512 × 5 = 2,560 times) is imperceptible.
3. **The arithmetic of the this reverse lookup**: a single line, `(this - m_pRegionBase - HEADER_SIZE - VTABLE_OBJ_OFFSET) \ SLOT_SIZE`, obtains idx from the `this` on which AddRef / Release were called. It was factored out into the `SlotIndexFromVTableObjAddr` helper.
4. **Reference management via m_handlers(idx) in earnest**: plan 1's direct access to `gHandler` is abolished completely. `Handler_AddRef` / `Handler_Release` / `Handler_QueryInterface` all operate on `m_handlers(idx)` starting from `this`. The `m_handlers` array placed as groundwork in stage 9.1 is actually used here for the first time.
5. **Coexistence of two reverse-lookup helpers**: `SlotIndexFromAddr` (= thunk address → idx, used to derive idx from `gHandler.Slot` in Release's automatic release path) and `SlotIndexFromVTableObjAddr` (= vtableObj address → idx, used by AddRef / Release / QI) coexist. The arithmetic differs subtly (whether or not `VTABLE_OBJ_OFFSET` is subtracted).
6. **Holding a local reference in HandlerReleaseInternal's automatic release path**: before setting `m_handlers(idx) = Nothing` at refcount=0, it saves the handler locally with `Dim h As Class7: Set h = m_handlers(idx)` and then releases in order. If `gHandler` points to the same handler, it is cut in sync.

### Resolving the mystery left by stage 9.1: the true signature of a VBA Function

Regarding the interpretation of Class7.Handler_Invoke's arguments in stage 9.1, after the operation check in stage 9.2 it was completely sorted out, reconfirming the knowledge established in stage 5.

**The thunk's argument shift and the hidden calling convention of a VBA Function**:

The thunk (the 56-byte body within the 74 bytes) does the following:

```asm
sub  rsp, 0x38           ; area for the return value pointer + shadow space + alignment
mov  r9,  r8             ; old r8 → r9
mov  r8,  rdx            ; old rdx → r8
mov  rdx, rcx            ; old rcx → rdx
mov  rcx, imm64          ; ObjPtr(gHandler) → rcx
lea  rax, [rsp+0x28]     ; storage for the return value
mov  [rsp+0x20], rax     ; push the 5th arg (return value pointer)
mov  rax, imm64          ; real address of Class7.Handler_Invoke
call rax
mov  eax, [rsp+0x28]     ; take out the return value into eax
add  rsp, 0x38
ret
```

The **true calling convention** of VBA's `Public Function Handler_Invoke(ByVal this, ByVal arg1, ByVal arg2) As Long`, as found in stage 5, is the following:

```c
HRESULT __stdcall Handler_Invoke(
    Me,           // arg0: rcx          ← the Me that VBA consumes implicitly
    this,         // arg1: rdx
    arg1,         // arg2: r8
    arg2,         // arg3: r9
    Long* pRetVal // arg4: [rsp+0x20]   ← the return value pointer the thunk prepares
);
```

In other words, a VBA Function has a **hidden `Me` argument** before the declared arguments and a **hidden out argument for writing back the return value** at the end, so a method declared with 3 arguments is actually compiled as a 5-argument `__stdcall` function. The reason the thunk takes the form `call + lea + mov + ret` is precisely to prepare this out argument; this is a design established at the point where we solved the problem of "having no choice but to make the Function a Sub" in stage 5.

**Flow of WebView2's `Invoke(obj, errorCode, pEnvironment)`**:

| WebView2's view | Before the thunk (registers) | After the thunk (registers) | VBA Class7's view (declared arguments) | Contents |
|---|---|---|---|---|
| 1st: obj | rcx | rdx | **`this`** | obj (= `pSlot + VTABLE_OBJ_OFFSET`) |
| 2nd: errorCode | rdx | r8 | **`arg1`** | errorCode (32 bit) |
| 3rd: pEnvironment | r8 | r9 | **`arg2`** | pEnvironment |
| (none) | (none) | rcx | **`Me` (implicit)** | `ObjPtr(gHandler)` |
| (none) | (none) | `[rsp+0x20]` | (return value pointer, implicit) | Prepared by the thunk |

**The true nature of the log line "`this : 2364645376240`"**:

The value that Class7.Handler_Invoke prints with `Debug.Print "  this:", this` in the stage 9.1 / 9.2 logs is **in fact `obj` (= `pSlot + VTABLE_OBJ_OFFSET`)**, not `ObjPtr(gHandler)` (= the real `Me`). VBA's implicit `Me` is consumed inside VBA, so even if user code receives it under the name `this`, it cannot be seen.

The reason `this` is always the same value `2364645376240` on consecutive Creates is separate from VBA's internal allocator regenerating the `gHandler` Class instance at the same address; it is merely that **the Acquired slot is always Slot 0 (= the same address), so the values of `pSlot + VTABLE_OBJ_OFFSET` coincide**. In Test_TwoHandlers_Mock h1 and h2 take different pSlots, so if Class7.Handler_Invoke were run on a real machine for both handlers, the values of `this` should differ as well (though this test only does AddRef / Release and does not call Invoke).

### Constraints of stage 9.2 and tasks from stage 9.3 onward

- **Verifying the argument interpretation of multi-argument handlers is for the next stage**: when implementing 2-argument `Invoke(sender, args)` handlers such as NavigationStarting / NavigationCompleted / WebMessageReceived, Class7.Handler_Invoke's declared arguments `this` / `arg1` will receive `sender` / `args` respectively. For HK_EnvironmentCompleted the correspondence was `this = obj` (not used), `arg1 = errorCode`, `arg2 = pEnvironment`, but for 2-argument handlers it becomes `this = sender`, `arg1 = args`, and `arg2` must not be touched (a garbage value from r9). Confirming this on a real machine will be the first task of stage 9.3.
- **Class7's `m_owner` remains Nothing**: it will be used for the first time from stage 9.3 onward, when the upper-level class (such as the Environment class) is introduced.
- **The IID check is not implemented**: `Handler_QueryInterface` accepts all IIDs and returns itself. It should properly return S_OK for only 2, `IUnknown` and `ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler`. Planned for stage 9.5.
- **The single `gHandler` variable remains**: a convenience variable for `CreateWebView2Environment`. When it is replaced in stage 9.3 by handler management via the upper-level class, `gHandler` may also have served its purpose.

### Outlook toward stage 9.3

- Introduce the upper-level class (a class that handles Environment / Controller). `Class7.m_owner` comes alive here. Build a mechanism for handlers to convey event completion notifications to the upper-level class.
- Add cases inside Class7 for multi-argument handlers (NavigationStarting / WebMessageReceived, etc.) and verify the argument interpretation on a real machine.
- Once Environment is complete, implement the flow of generating a Controller from it, aiming at the first milestone "a page is displayed in WebView2".

``` VBA
''''''''''''''''''''''''''''''''''
' --- Module9_2.bas Stage 9.2 WebView2Thunks (multiple handler support + embedded vtableObj) ---
'
'   Differences from stage 9.1 (Module9_1.bas):
'     - Embed a vtableObj region (40 bytes) in each slot (plan 2 adopted)
'     - SLOT_SIZE: 176 → 224 (vtableObj 40 bytes + 8 bytes of padding)
'     - Newly create VTABLE_OBJ_OFFSET = 176 (the same value as the old SLOT_SIZE, easy to remember)
'     - Abolish the module-level variable m_vtableObj
'     - Address passed to WebView2: VarPtr(m_vtableObj) → pSlot + VTABLE_OBJ_OFFSET
'     - Obtain the addresses of Handler_QI / AddRef / Release once at startup and
'       save them in m_pHandler_QI / _AddRef / _Release (not taken at every Acquire)
'     - Bulk-initialize the vtableObj of all 512 slots in Thunks_Init:
'         pVTable      = pFunctions
'         Functions(0) = m_pHandler_QI
'         Functions(1) = m_pHandler_AddRef
'         Functions(2) = m_pHandler_Release
'         Functions(3) = pSlot (= start of the stub clone)
'     - Thunks_AcquireSlot does not touch the vtableObj (already initialized in Init)
'       Also the return value is still pSlot, but on the CreateWebView2Environment side it is changed
'       to pass "pSlot + VTABLE_OBJ_OFFSET" to WebView2
'     - Newly create the SlotIndexFromVTableObjAddr helper (for the AddRef/Release path)
'       Coexists with the existing SlotIndexFromAddr (= for thunk addresses)
'     - Handler_AddRef / Handler_Release are changed to reverse-lookup idx from this and
'       operate on m_handlers(idx) directly (abolishing plan 1's direct access to gHandler)
'     - Add Test_TwoHandlers_Mock for confirming multiple-handler operation
'       (a unit test of the pure reverse-lookup logic, instead of launching two real WebView2s in parallel)
'
'   Operation checks at this stage:
'     - Create / Finalize / consecutive Create with 1 handler keep working (regression)
'     - When 2 handlers are Acquired and AddRef/Release are fired separately from each one's
'       vtableObj address, the corresponding m_handlers entries are operated correctly via the idx reverse lookup
'     - The automatic release paths of the 2 handlers (refcount=0 → ClearOwner → clear m_handlers
'       → ReleaseSlot) run independently
'
'   Stage 9.1 constraints resolved in stage 9.2:
'     - vtableObj was a single instance → independent per slot
'     - AddRef/Release could not reverse-lookup idx from this → can be reverse-looked-up with one line of arithmetic
'     - The m_handlers(0..511) array was not actually used → used in earnest
'
'   Region layout (total 64 + 224×512 = 114,752 bytes ≒ 112 KB):
'     +0..+7        : alive flag (LongPtr, 8 bytes; only the lowest 1 byte is used)
'     +8..+63       : header reserved area
'     +64           : slot 0 starts (224 bytes)
'     +64+224       : slot 1 starts
'       ...
'     +64+224×511   : slot 511 starts
'
'   Each slot (224 bytes):
'     +0..+90       : EntryPoint stub clone (91 bytes)
'     +91..+95      : padding
'     +96..+169     : thunk (74 bytes)
'     +170..+175    : padding (8-byte alignment)
'     +176..+183    : vtableObj.pVTable (= points to +184)         ← VTABLE_OBJ_OFFSET = 176
'     +184..+215    : vtableObj.Functions(0..3) (32 bytes)
'     +216..+223    : spare padding
'
'   Structure (same call path as 9.0 / 9.1):
'     obj held by WebView2 = pSlot + VTABLE_OBJ_OFFSET
'                       ↓
'                 obj->vtbl->Invoke = Functions(3) = pSlot
'                       ↓
'                 clone of the EntryPoint stub (91 bytes)
'                       ↓
'                 after passing the EBMode check, indirect jump to the address at +55
'                       ↓
'                 thunk start (+96 = alive flag check)
'                       ↓
'                 if the flag is alive, the main processing (argument shift + ObjPtr injection)
'                       ↓
'                 Class7.Handler_Invoke (branches on m_kind)
'
'     When WebView2 calls AddRef/Release:
'                 obj->vtbl->AddRef = Functions(1) = Handler_AddRef
'                       ↓
'                 Handler_AddRef(this = pSlot + VTABLE_OBJ_OFFSET)
'                       ↓
'                 reverse-lookup idx with SlotIndexFromVTableObjAddr
'                       ↓
'                 m_handlers(idx).RefCount += 1
'
'   CopyMemory remains unused. All writes go through MemLongPtr.
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/refTypes/discussions/3
''

Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "This module requires 64-bit VBA (x64)"
#End If

' --- Structures for SafeArray / PointerAccessor ---
Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

' --- WebView2Loader.dll ---
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

' --- Win32 API ---
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long) As LongPtr

Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long

Private Const MEM_COMMIT             As Long = &H1000
Private Const MEM_RESERVE            As Long = &H2000
Private Const MEM_RELEASE            As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const S_OK                   As Long = 0

' --- Slot / region layout constants ---

' Position within the EntryPoint stub where the "address to call next" is embedded as an immediate
Private Const LATE_BIND_OFFSET As Long = 55

' Actual length of the stub (confirmed in stage 8.5; 91 bytes including the C2 00 00 at +88)
Private Const STUB_LEN As Long = 91

' Length of the thunk (alive flag check 18 + existing 56 = 74 bytes)
Private Const THUNK_LEN As Long = 74

' Offset of the thunk region (stub 91 + padding 5 = 96)
Private Const THUNK_OFFSET   As Long = 96

' Offset of the vtableObj region (= the same value as the old SLOT_SIZE, easy to remember)
'   From this position, the WebView2VTable structure (40 bytes) is placed:
'     +0..+7   : pVTable (= points to its own +8)
'     +8..+15  : Functions(0) = Handler_QueryInterface
'     +16..+23 : Functions(1) = Handler_AddRef
'     +24..+31 : Functions(2) = Handler_Release
'     +32..+39 : Functions(3) = pSlot (start of the stub clone)
Private Const VTABLE_OBJ_OFFSET As Long = 176

' Allocation size per slot (rounded to a multiple of 8, 224 bytes)
'   +0..+90      stub clone (91 bytes)
'   +91..+95     padding
'   +96..+169    thunk (74 bytes)
'   +170..+175   padding (8-byte alignment)
'   +176..+183   vtableObj.pVTable
'   +184..+215   vtableObj.Functions(0..3)
'   +216..+223   spare padding (8-byte alignment)
Private Const SLOT_SIZE      As Long = 224

' Number of slots (fixed)
Private Const SLOT_COUNT     As Long = 512

' Header size at the start of the region (alive flag + future reservation)
'   +0          : alive flag (1 byte, but handled in 8-byte units)
'   +1..+63     : reserved (for a future sentinel flag-clearing thunk, etc.)
Private Const HEADER_SIZE    As Long = 64

' Size of the whole region (= 64 + 224×512 = 114,752 bytes ≒ 112 KB)
Private Const REGION_SIZE    As Long = HEADER_SIZE + SLOT_SIZE * SLOT_COUNT


' ============================================================
' HandlerKind enum (referenceable from upper-level modules)
'   Identifies the kind of WebView2 callback. It is stored in Class7.cls's
'   m_kind and used for branching inside Handler_Invoke.
' ============================================================
Public Enum HandlerKind
    HK_None = 0
    HK_EnvironmentCompleted = 1
    HK_ControllerCompleted = 2
    HK_NavigationStarting = 3
    HK_NavigationCompleted = 4
    HK_WebMessageReceived = 5
    HK_DocumentTitleChanged = 6
End Enum


' --- Function addresses obtained once at startup and kept ---
'   The values written into each slot's vtableObj.Functions(0..2).
'   The AddressOf operator can only be written at an actual-argument position, so they are obtained via GetAddr.
'   Obtained only once in Thunks_Init and reused at Acquire time.
Private m_pHandler_QI       As LongPtr
Private m_pHandler_AddRef   As LongPtr
Private m_pHandler_Release  As LongPtr

' --- Single handler created at startup (the old gHandler As Class6 changed to Class7) ---
'   In stage 9.2 as well, CreateWebView2Environment still launches 1 handler, so it is kept.
'   Multiple-handler tests are done separately with Test_TwoHandlers_Mock.
Private gHandler As Class7

' --- Slot pool state ---
Private m_pRegionBase As LongPtr        ' Start of the region allocated with VirtualAlloc
Private m_freeHead    As Long           ' Head index of the free list (-1 = empty)
Private m_freeNext()  As Long           ' freeNext(i) = the next free slot after slot i
Private m_inUse       As Long           ' Number of slots currently in use (for debugging)

' --- Correspondence table: slot index → Handler object ---
'   Full-scale use begins in stage 9.2. From Handler_AddRef / Release / QueryInterface,
'   via the path "this (= pSlot + VTABLE_OBJ_OFFSET) → idx (SlotIndexFromVTableObjAddr)
'    → m_handlers(idx)" the corresponding Class7 instance is reached.
'   This array is also cleared in the automatic release path at refcount = 0.
Private m_handlers(0 To SLOT_COUNT - 1) As Class7

Private Const THUNK_BUF_SIZE As Long = 80     ' 74 rounded up to a multiple of 8


' ============================================================
' Source of the EntryPoint stub
'   The byte sequence of this stub is copied into the VirtualAlloc region.
'   It must be called once at startup for Force-Compile.
' ============================================================
Private Sub EntryPoint(): End Sub


' ============================================================
' Public entry (equivalent of stage 8.9's CreateWebView2Environment)
'
'   Call Thunks_Init at startup, then take 1 slot with Acquire and
'   request creation of the WebView2 environment just as in stage 8.9.
'
'   Changes in stage 9.2:
'     - The vTable has already been initialized for all slots in Thunks_Init, so nothing is
'       done here (the old "build the vTable" block becomes entirely unnecessary)
'     - The address passed to WebView2 is not VarPtr(m_vtableObj) but
'       pSlot + VTABLE_OBJ_OFFSET (the start of the vtableObj inside each slot)
' ============================================================
Public Sub CreateWebView2Environment()

    ' --- 1. Initialize the slot pool (if not yet) ---
    If m_pRegionBase = 0 Then
        If Not Thunks_Init() Then
            MsgBox "Thunks_Init failed", vbCritical
            Exit Sub
        End If
    End If

    ' --- 2. Create the handler object ---
    Set gHandler = New Class7

    Dim pHandlerInvoke As LongPtr
    pHandlerInvoke = GetClassMethodAddrAtFixedSlot(gHandler, 7)
    If pHandlerInvoke = 0 Then
        MsgBox "Failed to obtain the address of Class7.Handler_Invoke", vbCritical
        Exit Sub
    End If

    ' --- 3. Take 1 slot, write the thunk + register in m_handlers ---
    Dim pSlot As LongPtr
    pSlot = Thunks_AcquireSlot(gHandler, ObjPtr(gHandler), pHandlerInvoke)
    If pSlot = 0 Then
        MsgBox "Failed to acquire a slot (pool exhausted or initialization failed)", vbCritical
        Exit Sub
    End If

    ' Set the kind, owner and pSlot on the Handler itself
    '   At this stage owner can still be Nothing (the upper-level class comes in stage 9.3 onward)
    gHandler.Init HK_EnvironmentCompleted, Nothing, pSlot

    ' --- 4. Create the WebView2 environment ---
    '   For the 4th argument, pass "the start address of the vtableObj inside the slot".
    '   This address is held by WebView2 as obj and used as the this of the later AddRef/Release/Invoke.
    '   Invoke reaches the thunk via Functions(3) = pSlot.
    Dim pVTableObj As LongPtr
    pVTableObj = pSlot + VTABLE_OBJ_OFFSET

    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, pVTableObj)

    If hr <> S_OK Then
        MsgBox "Failed to initialize WebView2. HRESULT: &H" & Hex(hr), vbCritical
    End If

    DoEvents
End Sub


' ============================================================
' Shutdown processing (formerly FinalizeWebView2Environment)
'
'   By setting the alive flag to 0 before VirtualFree,
'   even if there are unreleased handlers (= even if WebView2 is still
'   holding the vTable address), callbacks are safely cut off with S_OK.
' ============================================================
Public Sub FinalizeWebView2Environment()
    Thunks_Shutdown
    Set gHandler = Nothing
End Sub


' ============================================================
' Test_TwoHandlers_Mock
'
'   Core feature test of stage 9.2: Acquire 2 handlers,
'   fire AddRef/Release separately from each one's vtableObj address, and confirm
'   that the idx reverse lookup works correctly. WebView2 is not called
'   (a pure unit test of the reverse-lookup logic).
'
'   Expected behavior:
'     - h1 and h2 obtain separate slots (e.g. idx=0 and idx=1)
'     - An AddRef on h1's vtableObj increases only h1's m_refCount
'     - An AddRef on h2's vtableObj increases only h2's m_refCount
'     - Automatic release fires at the moment each independently reaches refcount=0
'     - In the end m_inUse = 0 and both slots return to the free list
'
'   This test does not use gHandler (= independent of the handler of
'   CreateWebView2Environment). Only the local variables h1/h2
'   hold the Class instances, so they are released automatically when the Test leaves scope.
'   However, "m_handlers(idx) = Nothing" runs inside the refcount=0
'   automatic release path, so they have most likely already been released during the Test.
' ============================================================
Public Sub Test_TwoHandlers_Mock()

    Debug.Print String(60, "=")
    Debug.Print "Test_TwoHandlers_Mock start"
    Debug.Print String(60, "=")

    ' --- 1. Initialize the pool ---
    If Not Thunks_Init() Then
        Debug.Print "Thunks_Init failed"
        Exit Sub
    End If

    Debug.Print "Thunks_Init OK, m_inUse = " & m_inUse

    ' --- 2. Create 2 handlers + Acquire ---
    Dim h1 As Class7, h2 As Class7
    Set h1 = New Class7
    Set h2 = New Class7

    Dim pInvoke1 As LongPtr, pInvoke2 As LongPtr
    pInvoke1 = GetClassMethodAddrAtFixedSlot(h1, 7)
    pInvoke2 = GetClassMethodAddrAtFixedSlot(h2, 7)

    Dim pSlot1 As LongPtr, pSlot2 As LongPtr
    pSlot1 = Thunks_AcquireSlot(h1, ObjPtr(h1), pInvoke1)
    pSlot2 = Thunks_AcquireSlot(h2, ObjPtr(h2), pInvoke2)

    If pSlot1 = 0 Or pSlot2 = 0 Then
        Debug.Print "Acquire failed"
        Exit Sub
    End If

    h1.Init HK_EnvironmentCompleted, Nothing, pSlot1
    h2.Init HK_ControllerCompleted, Nothing, pSlot2

    Dim idx1 As Long, idx2 As Long
    idx1 = SlotIndexFromAddr(pSlot1)
    idx2 = SlotIndexFromAddr(pSlot2)
    Debug.Print "h1 -> idx " & idx1 & ", pSlot " & pSlot1
    Debug.Print "h2 -> idx " & idx2 & ", pSlot " & pSlot2
    Debug.Print "m_inUse = " & m_inUse & " (expected 2)"

    If idx1 = idx2 Then
        Debug.Print "[NG] Different handlers obtained the same idx"
        Exit Sub
    End If

    ' --- 3. Directly compute each handler's vtableObj address ---
    '   Create the same value as the this that WebView2 would call with.
    Dim pObj1 As LongPtr, pObj2 As LongPtr
    pObj1 = pSlot1 + VTABLE_OBJ_OFFSET
    pObj2 = pSlot2 + VTABLE_OBJ_OFFSET

    ' --- 4. Fire AddRef on each and confirm the counts go up independently ---
    Debug.Print "--- AddRef sequence ---"
    Dim r As Long
    ' h1: 0 -> 1
    r = Handler_AddRef(pObj1)
    Debug.Print "  Handler_AddRef(pObj1) -> " & r
    ' h1: 1 -> 2
    r = Handler_AddRef(pObj1)
    Debug.Print "  Handler_AddRef(pObj1) -> " & r
    ' h2: 0 -> 1
    r = Handler_AddRef(pObj2)
    Debug.Print "  Handler_AddRef(pObj2) -> " & r
    Debug.Print "h1.RefCount = " & h1.RefCount & " (expected 2)"
    Debug.Print "h2.RefCount = " & h2.RefCount & " (expected 1)"

    If h1.RefCount <> 2 Or h2.RefCount <> 1 Then
        Debug.Print "[NG] AddRef dispatching differs from the expected values"
        ' Continue as is and clean up
    Else
        Debug.Print "[OK] AddRef dispatching is as expected"
    End If

    ' --- 5. Fire Release and confirm each independently reaches refcount=0 ---
    Debug.Print "--- Release sequence ---"
    ' h1: 2 -> 1
    r = Handler_Release(pObj1)
    Debug.Print "  Handler_Release(pObj1) -> " & r
    Debug.Print "h1.RefCount = " & h1.RefCount & " (expected 1)"
    Debug.Print "m_inUse = " & m_inUse & " (expected 2, both still in use)"

    ' h1: 1 -> 0, automatic release fires
    r = Handler_Release(pObj1)
    Debug.Print "  Handler_Release(pObj1) -> " & r
    Debug.Print "(after h1 release) m_inUse = " & m_inUse & " (expected 1)"

    ' After h1 is released, m_handlers(idx1) = Nothing, and
    ' only the local variable h1 holds the Class7 instance.
    ' At this point h1.RefCount should be 0.
    Debug.Print "h1.RefCount (after release) = " & h1.RefCount & " (expected 0)"
    Debug.Print "h2.RefCount (independent of h1's release) = " & h2.RefCount & " (expected 1)"

    ' h2: 1 -> 0, automatic release fires
    r = Handler_Release(pObj2)
    Debug.Print "  Handler_Release(pObj2) -> " & r
    Debug.Print "(after h2 release) m_inUse = " & m_inUse & " (expected 0)"
    Debug.Print "h2.RefCount (after release) = " & h2.RefCount & " (expected 0)"

    ' --- 6. Confirm that operations on released slots pass safely as no-ops ---
    Debug.Print "--- Post-release safety check ---"
    r = HandlerAddRefInternal(pObj1)
    Debug.Print "AddRef return value to the released pObj1 = " & r & " (expected 0)"
    r = HandlerReleaseInternal(pObj2)
    Debug.Print "Release return value to the released pObj2 = " & r & " (expected 0)"

    ' --- 7. Cleanup ---
    Set h1 = Nothing
    Set h2 = Nothing

    Debug.Print "--- Final state ---"
    Debug.Print "m_inUse = " & m_inUse & " (expected 0)"
    Debug.Print "m_freeHead = " & m_freeHead

    ' After the Test, fully free the region with Thunks_Shutdown.
    '   - Even if there is a leak in m_handlers on an abnormal path, it is reliably cleaned up
    '   - There is no problem calling CreateWebView2Environment or another Test right after
    Thunks_Shutdown

    Debug.Print "Test_TwoHandlers_Mock done"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Thunks_Init
'   Allocate the region with VirtualAlloc, set the alive flag in the header,
'   copy the stub into all slots, initialize the vtableObj,
'   and initialize the free list.
'
'   Additional processing in stage 9.2:
'     - Obtain the addresses of Handler_QI / AddRef / Release just once
'     - Bulk-initialize the vtableObj of all 512 slots
'         pVTable      = pFunctions
'         Functions(0) = m_pHandler_QI
'         Functions(1) = m_pHandler_AddRef
'         Functions(2) = m_pHandler_Release
'         Functions(3) = pSlot (= start of the stub clone)
'       This removes the need to touch the vtableObj at Acquire time.
'
'   Return value: success = True, failure = False
' ============================================================
Public Function Thunks_Init() As Boolean

    If m_pRegionBase <> 0 Then
        Thunks_Init = True
        Exit Function
    End If

    ' Force-Compile EntryPoint (materialize the source of the stub)
    EntryPoint

    ' Source address of the EntryPoint stub
    Dim pStubSrc As LongPtr
    pStubSrc = VBA.Int(AddressOf EntryPoint)
    If pStubSrc = 0 Then Exit Function

    ' Obtain the addresses of Handler_QI / AddRef / Release just once at startup and save them
    '   AddressOf can only be written at an actual-argument position, so obtain them via GetAddr.
    m_pHandler_QI      = GetAddr(AddressOf Handler_QueryInterface)
    m_pHandler_AddRef  = GetAddr(AddressOf Handler_AddRef)
    m_pHandler_Release = GetAddr(AddressOf Handler_Release)
    If m_pHandler_QI = 0 Or m_pHandler_AddRef = 0 Or m_pHandler_Release = 0 Then
        Exit Function
    End If

    ' Allocate the region in bulk
    m_pRegionBase = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If m_pRegionBase = 0 Then Exit Function

    ' Zero-initialize the header (VirtualAlloc is already zero-cleared by nature, but just in case)
    Dim k As Long
    For k = 0 To HEADER_SIZE - 1 Step 8
        MemLongPtr(m_pRegionBase + k) = 0^
    Next k

    ' Set the alive flag (1 byte at the start of the region = 1)
    ' Since it is written in 8-byte units, put 0x01 so that only the lowest 1 byte becomes 1
    MemLongPtr(m_pRegionBase) = 1^

    ' Initialize all slots (stub copy + vtableObj initialization)
    Dim i As Long, pSlot As LongPtr, pVTableObj As LongPtr, pFunctions As LongPtr
    For i = 0 To SLOT_COUNT - 1
        pSlot = SlotAddrAt(i)

        ' --- Stub copy (96 bytes each; the padding at +91..+95 is harmless) ---
        For k = 0 To 95 Step 8
            MemLongPtr(pSlot + k) = ReadLongPtr(pStubSrc + k)
        Next k

        ' --- vtableObj initialization (40 bytes, written in 8-byte units) ---
        '   Functions(0..2) are the Handler_QI/AddRef/Release common to all slots.
        '   Functions(3) is "the start of its own stub clone". Whichever handler
        '   this slot is assigned to at Acquire/Release, Functions(3) may
        '   always point to the same pSlot (the thunk contents are rewritten
        '   at Acquire time, but the thunk's start address itself is invariant).
        pVTableObj = pSlot + VTABLE_OBJ_OFFSET
        pFunctions = pVTableObj + PtrSize        ' +8, right after pVTable

        MemLongPtr(pVTableObj)                  = pFunctions
        MemLongPtr(pFunctions + 0 * PtrSize)    = m_pHandler_QI
        MemLongPtr(pFunctions + 1 * PtrSize)    = m_pHandler_AddRef
        MemLongPtr(pFunctions + 2 * PtrSize)    = m_pHandler_Release
        MemLongPtr(pFunctions + 3 * PtrSize)    = pSlot
    Next i

    ' Initialize the free list
    '   m_freeNext(i) = i + 1 (the last is -1)
    '   m_freeHead    = 0
    ReDim m_freeNext(0 To SLOT_COUNT - 1)
    For i = 0 To SLOT_COUNT - 2
        m_freeNext(i) = i + 1
    Next i
    m_freeNext(SLOT_COUNT - 1) = -1
    m_freeHead = 0
    m_inUse = 0

    ' Clear the Handler correspondence table entirely (it is a fixed array so Nothing at declaration, but just in case)
    For i = 0 To SLOT_COUNT - 1
        Set m_handlers(i) = Nothing
    Next i

    Thunks_Init = True
End Function


' ============================================================
' Thunks_AcquireSlot
'   Take 1 free slot from the free list, write into that slot's
'   thunk region the machine code that "injects pSelfObj into Me and calls pTargetFunc",
'   and write the thunk's address into +55.
'
'   Change in stage 9.1: receives the handler object as the first argument
'   and registers it in m_handlers(idx). This lets Handler_AddRef / Release
'   reverse-lookup idx from the this address and reach the Handler object.
'
'   Changes in stage 9.2 (effectively zero):
'     - The vtableObj has already been initialized for all slots in Thunks_Init, so it is
'       not touched here. Only the thunk body is written.
'     - The address passed to WebView2 is "pSlot + VTABLE_OBJ_OFFSET", but
'       this is computed by the caller (CreateWebView2Environment, etc.).
'       The return value of this function remains pSlot (= start of the slot).
'
'   Return value: start address of the slot (the reference point used inside Thunks), 0 on failure
' ============================================================
Public Function Thunks_AcquireSlot( _
    ByVal handler As Class7, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr) As LongPtr

    If m_pRegionBase = 0 Then Exit Function
    If m_freeHead < 0 Then Exit Function     ' Pool exhausted
    If handler Is Nothing Then Exit Function

    ' Take one from the free list
    Dim idx As Long
    idx = m_freeHead
    m_freeHead = m_freeNext(idx)
    m_freeNext(idx) = -1                      ' -1 = in use

    Dim pSlot As LongPtr
    pSlot = SlotAddrAt(idx)

    ' Write the thunk (74 bytes from offset +96)
    WriteThunkMachineCode pSlot + THUNK_OFFSET, pSelfObj, pTargetFunc, m_pRegionBase

    ' Write the thunk's address into +55 of the stub clone
    MemLongPtr(pSlot + LATE_BIND_OFFSET) = pSlot + THUNK_OFFSET

    ' Associate the Handler object with idx
    Set m_handlers(idx) = handler

    m_inUse = m_inUse + 1
    Thunks_AcquireSlot = pSlot
End Function


' ============================================================
' Thunks_ReleaseSlot
'   Return the specified slot to the free list.
'   Also clears m_handlers(idx).
'   The contents inside the slot are not zero-cleared; they are overwritten at the next Acquire.
'
'   Note: WebView2 may still be holding vTable[3] (= this slot),
'         so the caller must confirm that it is a timing at which
'         Release may be called.
'         From stage 9.1 onward it is called automatically from Handler_Release
'         at refcount=0, so there is basically no need to call it directly from outside.
' ============================================================
Public Sub Thunks_ReleaseSlot(ByVal pSlot As LongPtr)
    If m_pRegionBase = 0 Then Exit Sub
    If pSlot = 0 Then Exit Sub

    Dim idx As Long
    idx = SlotIndexFromAddr(pSlot)
    If idx < 0 Then Exit Sub
    If m_freeNext(idx) <> -1 Then Exit Sub           ' Already free = double release

    ' Also clear the Handler correspondence table
    Set m_handlers(idx) = Nothing

    ' Return it to the head of the free list
    m_freeNext(idx) = m_freeHead
    m_freeHead = idx

    m_inUse = m_inUse - 1
End Sub


' ============================================================
' SlotIndexFromAddr (internal helper)
'   Computes the slot index backward from the thunk's address (= the slot's
'   start address).
'   Returns -1 for an invalid address (outside the region, unaligned).
'
'   Uses:
'     - HandlerReleaseInternal obtains idx from gHandler.Slot
'     - Double-release detection in Thunks_ReleaseSlot
' ============================================================
Private Function SlotIndexFromAddr(ByVal pSlot As LongPtr) As Long
    SlotIndexFromAddr = -1
    If m_pRegionBase = 0 Then Exit Function

    Dim offset As LongPtr
    offset = pSlot - (m_pRegionBase + HEADER_SIZE)
    If offset < 0 Then Exit Function
    If (offset Mod SLOT_SIZE) <> 0 Then Exit Function

    Dim idx As Long
    idx = CLng(offset \ SLOT_SIZE)
    If idx < 0 Or idx >= SLOT_COUNT Then Exit Function

    SlotIndexFromAddr = idx
End Function


' ============================================================
' SlotIndexFromVTableObjAddr (internal helper)
'   Computes the slot index backward from the vtableObj address (= the value WebView2 passes as this in AddRef/Release,
'                     = pSlot + VTABLE_OBJ_OFFSET).
'
'   Uses:
'     - Handler_AddRef / Handler_Release / Handler_QueryInterface
'       These this values are passed directly from WebView2 rather than via the thunk,
'       so this = pSlot + VTABLE_OBJ_OFFSET
'
'   Returns -1 for an invalid address (outside the region, unaligned).
' ============================================================
Private Function SlotIndexFromVTableObjAddr(ByVal pVTableObj As LongPtr) As Long
    SlotIndexFromVTableObjAddr = -1
    If m_pRegionBase = 0 Then Exit Function

    Dim offset As LongPtr
    offset = pVTableObj - (m_pRegionBase + HEADER_SIZE + VTABLE_OBJ_OFFSET)
    If offset < 0 Then Exit Function
    If (offset Mod SLOT_SIZE) <> 0 Then Exit Function

    Dim idx As Long
    idx = CLng(offset \ SLOT_SIZE)
    If idx < 0 Or idx >= SLOT_COUNT Then Exit Function

    SlotIndexFromVTableObjAddr = idx
End Function


' ============================================================
' Thunks_Shutdown
'   Set the alive flag to 0, then free the region with VirtualFree.
'   After the call, even if a thunk runs it returns S_OK immediately
'   (however, after VirtualFree the region itself is gone, so setting the flag
'   is meaningful only for "the instant until VirtualFree completes").
' ============================================================
Public Sub Thunks_Shutdown()
    If m_pRegionBase = 0 Then Exit Sub

    ' First, set the alive flag to 0
    '   This way, even if a callback arrives from WebView2 addressed to an unreleased handler,
    '   it returns S_OK immediately at the flag check at the start of the thunk.
    MemLongPtr(m_pRegionBase) = 0^

    ' Cut the owner of all Handlers and release the references
    '   Always call ClearOwner to prevent circular references.
    Dim i As Long
    For i = 0 To SLOT_COUNT - 1
        If Not (m_handlers(i) Is Nothing) Then
            m_handlers(i).ClearOwner
            Set m_handlers(i) = Nothing
        End If
    Next i

    ' Free the region
    VirtualFree m_pRegionBase, 0, MEM_RELEASE

    m_pRegionBase = 0
    m_freeHead = -1
    m_inUse = 0
    Erase m_freeNext
End Sub


' ============================================================
' Slot index → slot start address
' ============================================================
Private Function SlotAddrAt(ByVal idx As Long) As LongPtr
    SlotAddrAt = m_pRegionBase + HEADER_SIZE + CLngLng(idx) * SLOT_SIZE
End Function


' ============================================================
' Write the thunk's machine code to the specified address (effective length 74 bytes)
'
'   Layout:
'     +0..+17  : alive flag check (18 bytes)
'     +18..+73 : existing thunk body (56 bytes, same as stages 6 to 8)
'
'   Alive flag check:
'     mov  rax, pAliveFlag      [10 bytes: 48 B8 + 8 byte imm]
'     cmp  byte ptr [rax], 1    [3 bytes:  80 38 01]
'     je   alive (+3)           [2 bytes:  74 03]
'     xor  eax, eax             [2 bytes:  33 C0]   (S_OK)
'     ret                       [1 byte:   C3]
'   alive:
'     (from here on, the existing 56-byte thunk)
'
'   Implementation note:
'     The byte array b is declared with 80 bytes (= a multiple of 8), and the last 6 bytes are
'     treated as zero padding. The write destination, slot +96..+175 (= 80 bytes' worth),
'     is originally padding/unused area, so writing 80 bytes is harmless.
'     This lets us write plainly with the 8-byte-unit loop of MemLongPtr.
' ============================================================

Private Sub WriteThunkMachineCode( _
    ByVal addr As LongPtr, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr, _
    ByVal pRegionBase As LongPtr)

    Dim b(0 To THUNK_BUF_SIZE - 1) As Byte, i As Long
    i = 0

    ' --- Alive flag check (18 bytes) ---
    ' mov rax, imm64   (= pAliveFlag = pRegionBase + 0)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pRegionBase: i = i + 8
    ' cmp byte ptr [rax], 1
    b(i) = &H80: b(i + 1) = &H38: b(i + 2) = &H1: i = i + 3
    ' je +3 (= the alive label, skipping the 3 bytes of xor + ret)
    b(i) = &H74: b(i + 1) = &H3: i = i + 2
    ' xor eax, eax
    b(i) = &H33: b(i + 1) = &HC0: i = i + 2
    ' ret
    b(i) = &HC3: i = i + 1

    ' --- From here on, the same 56-byte thunk body as stages 6 to 8 ---

    ' sub rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4
    ' mov r9, r8
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    ' mov rcx, imm64   (ObjPtr(gHandler))
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pSelfObj: i = i + 8
    ' lea rax, [rsp+0x28]
    b(i) = &H48: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5
    ' mov [rsp+0x20], rax
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H20: i = i + 5
    ' mov rax, imm64   (pTargetFunc)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pTargetFunc: i = i + 8
    ' call rax
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
    ' mov eax, [rsp+0x28]
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H28: i = i + 4
    ' add rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4
    ' ret
    b(i) = &HC3: i = i + 1
    ' int3 / int3 (padding)
    b(i) = &HCC: i = i + 1
    b(i) = &HCC: i = i + 1

    ' --- Write 80 bytes in 8-byte units (the last 6 bytes are 0 padding, harmless) ---
    Dim k As Long
    For k = 0 To THUNK_BUF_SIZE - 1 Step 8
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k
End Sub


' ============================================================
' Take 8 bytes (LongPtr) from the given offset of a byte array
' ============================================================
Private Function ReadLongPtrFromBytes(ByRef b() As Byte, ByVal offset As Long) As LongPtr
    ReadLongPtrFromBytes = ReadLongPtr(VarPtr(b(offset)))
End Function


' ============================================================
' Get the function address of a fixed slot from the class's vTable
' ============================================================
Private Function GetClassMethodAddrAtFixedSlot( _
    ByVal cls As Object, ByVal slotIndex As Long) As LongPtr

    If cls Is Nothing Then Exit Function
    Dim pObj As LongPtr: pObj = ObjPtr(cls)
    If pObj = 0 Then Exit Function

    Dim pVTable As LongPtr
    pVTable = ReadLongPtr(pObj)
    If pVTable = 0 Then Exit Function

    GetClassMethodAddrAtFixedSlot = ReadLongPtr(pVTable + slotIndex * PtrSize)
End Function


' ============================================================
' IUnknown stubs (standard module)
'
'   Changes in stage 9.2:
'     - Reverse-lookup idx from this with SlotIndexFromVTableObjAddr
'     - Operate on m_handlers(idx) directly (abolishing plan 1's direct access to gHandler)
'     - This makes it possible to manage multiple handlers in parallel
'
'   The IID check is planned to be introduced from stage 9.5 onward. For now all IIDs are accepted.
' ============================================================
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

Private Function Handler_QueryInterface( _
    ByVal this As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long

    ' For now it accepts all IIDs and returns itself (the IID check is a future task)
    ppvObject = this

    ' The premise is that the pointer QueryInterface returns on success has been AddRef'ed,
    ' so we perform the equivalent of AddRef here
    HandlerAddRefInternal this

    Handler_QueryInterface = S_OK
End Function

Private Function Handler_AddRef(ByVal this As LongPtr) As Long
    Handler_AddRef = HandlerAddRefInternal(this)
End Function

Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = HandlerReleaseInternal(this)
End Function


' ============================================================
' HandlerAddRefInternal / HandlerReleaseInternal
'
'   Changed in stage 9.2 to the this reverse-lookup scheme:
'     - Obtain idx from this (= pSlot + VTABLE_OBJ_OFFSET)
'     - Operate on the object at m_handlers(idx)
'     - An invalid this (= a released slot or an address outside the region) safely returns 0
'
'   Return value: the refcount after the operation (following COM convention)
' ============================================================
Private Function HandlerAddRefInternal(ByVal this As LongPtr) As Long
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(this)
    If idx < 0 Then
        ' Invalid this (outside the region, unaligned, etc.). Return 0.
        Exit Function
    End If
    If m_handlers(idx) Is Nothing Then
        ' Operation on a released slot. Return 0.
        '   In principle this path should not be reached if the WebView2 side
        '   counts correctly, but we handle it defensively.
        Exit Function
    End If

    Dim n As Long
    n = m_handlers(idx).RefCount + 1
    m_handlers(idx).RefCount = n
    Debug.Print "  AddRef [idx=" & idx & "] → refCount = " & n
    HandlerAddRefInternal = n
End Function

Private Function HandlerReleaseInternal(ByVal this As LongPtr) As Long
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(this)
    If idx < 0 Then Exit Function
    If m_handlers(idx) Is Nothing Then Exit Function

    Dim n As Long
    n = m_handlers(idx).RefCount - 1
    If n < 0 Then n = 0                ' Defense (should not happen in principle)
    m_handlers(idx).RefCount = n
    Debug.Print "  Release [idx=" & idx & "] → refCount = " & n

    If n = 0 Then
        ' Since refcount has become 0, cut the circular reference and release the slot.
        ' The order matters (same as stage 9.1):
        '   1. Cut the circular reference with ClearOwner
        '   2. Hold a local reference to m_handlers(idx), then Set Nothing
        '   3. Return the slot to the free list with Thunks_ReleaseSlot
        '   4. Also set the gHandler reference to Nothing (however, gHandler at this stage is
        '      dedicated to CreateWebView2Environment; among multiple handlers,
        '      it is cut only if it matches m_handlers(idx))
        '
        ' Step 4 goes last so as to keep consistency in case the standard module state is read
        ' inside Class_Terminate (which may be added in the future).

        Dim h As Class7
        Set h = m_handlers(idx)

        Dim pSlot As LongPtr
        pSlot = h.Slot

        ' 1. Cut the circular reference (m_owner = Nothing)
        h.ClearOwner

        ' 2. Set m_handlers(idx) to Nothing
        Set m_handlers(idx) = Nothing

        ' 3. Return the slot to the free list
        Thunks_ReleaseSlot pSlot

        ' 4. If gHandler points to the same handler, cut it too
        '    (only the handler via CreateWebView2Environment is tied to gHandler.
        '     Via Test_TwoHandlers_Mock, gHandler is not used)
        If Not (gHandler Is Nothing) Then
            If gHandler Is h Then
                Set gHandler = Nothing
            End If
        End If

        ' h's reference is cut and released when leaving this procedure
        Set h = Nothing

        Debug.Print "  Release [idx=" & idx & "] → refCount = 0 → Slot " & idx & " auto-released"
    End If

    HandlerReleaseInternal = n
End Function


' ============================================================
' Memory-operation primitives via PointerAccessor
' ============================================================
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub
```

``` VBA
''''''''''''''''''''''''''''''''''
' --- Class7.cls Stage 9.1 WebView2Handler---
'
'   A general-purpose handler class replacing the old Class6.
'   It handles multiple kinds via m_kind, and the only Public method is Handler_Invoke
'   (pinned at vTable slot 7). As up to stage 9.0, it is called directly from WebView2.
'
'   State held:
'     m_kind     : handler kind (HK_EnvironmentCompleted, etc.)
'     m_owner    : upper-level class (can still be Nothing at this stage)
'     m_pSlot    : start address of the slot where the thunk is placed
'                  (= the this address that WebView2 calls)
'     m_refCount : reference count (starts from 1 in Init)
'
'   Countermeasure against circular references:
'     When refcount = 0, the foundation side (WebView2Thunks.bas) calls
'     ClearOwner to cut m_owner.
'     Because the upper-level class holds the WebView2Handler and the WebView2Handler holds
'     the upper-level class via m_owner, it leaks unless explicitly cut.
''''''''''''''''''''''''''''''''''
Option Explicit

#If Win64 Then
    ' OK
#Else
    #Error "This class requires 64-bit VBA (x64)"
#End If

Private m_kind     As HandlerKind
Private m_owner    As Object
Private m_pSlot    As LongPtr
Private m_refCount As Long


' ============================================================
' Initialization (assumed to be called from the foundation's Thunks_AcquireSlot)
'
'   refcount starts from 0 (plan X-1 adopted):
'     - Following COM semantics, the model is that once a handler is handed to WebView2,
'       WebView2 AddRefs it itself and takes responsibility. The VBA side holds no reference.
'     - It has been confirmed on a real machine that WebView2 calls AddRef inside
'       CreateCoreWebView2EnvironmentWithOptions (refcount: 0 → 1).
'     - After Invoke, WebView2 calls Release and refcount = 0, and
'       the automatic release path of HandlerReleaseInternal runs.
'
'   This initial value was decided because the real-machine logs (the operation check of stage 9.1)
'   revealed that "one AddRef + one Release" are managed symmetrically
'   on the WebView2 side.
'   If plan X-1 is not chosen, refcount stays at 1 and the slot leaks.
' ============================================================
Friend Sub Init( _
    ByVal Kind As HandlerKind, _
    ByVal owner As Object, _
    ByVal pSlot As LongPtr)

    m_kind = Kind
    Set m_owner = owner
    m_pSlot = pSlot
    m_refCount = 0
End Sub


' ============================================================
' Property group (for operation from the foundation side)
' ============================================================
Friend Property Get RefCount() As Long
    RefCount = m_refCount
End Property

Friend Property Let RefCount(ByVal v As Long)
    m_refCount = v
End Property

Friend Property Get Slot() As LongPtr
    Slot = m_pSlot
End Property

Friend Property Get Kind() As HandlerKind
    Kind = m_kind
End Property


' ============================================================
' ClearOwner
'   Called from the foundation side when refcount = 0.
'   Sets m_owner to Nothing to cut the circular reference.
' ============================================================
Friend Sub ClearOwner()
    Set m_owner = Nothing
End Sub


' ============================================================
' Handler_Invoke (pinned at vTable slot 7)
'
'   The only entry that receives callbacks from WebView2.
'   The thunk injects ObjPtr(this) into Me, shifts the arguments one at a time,
'   and reaches this method.
'
'   Argument specification (settled in stage 9.1):
'     this : the ObjPtr of itself that the thunk injects into Me (not used)
'     arg1 : the 1st argument WebView2 put in rcx (received as LongPtr)
'     arg2 : the 2nd argument WebView2 put in rdx (received as LongPtr)
'
'   Interpretation per m_kind:
'     HK_EnvironmentCompleted: arg1 = errorCode (lower 32 bits), arg2 = pEnvironment
'     HK_ControllerCompleted:  arg1 = errorCode (lower 32 bits), arg2 = pController
'     HK_NavigationStarting:   arg1 = sender, arg2 = args
'     HK_NavigationCompleted:  arg1 = sender, arg2 = args
'     HK_WebMessageReceived:   arg1 = sender, arg2 = args
'     HK_DocumentTitleChanged: arg1 = sender, arg2 = args
'
'   In stage 9.1 only HK_EnvironmentCompleted was actually confirmed to work.
'   The other kinds just print "not supported" with Debug.Print.
' ============================================================
Public Function Handler_Invoke( _
    ByVal this As LongPtr, _
    ByVal arg1 As LongPtr, _
    ByVal arg2 As LongPtr) As Long

    Select Case m_kind
        Case HK_EnvironmentCompleted
            Dim errorCode As Long
            errorCode = LongPtrLowDword(arg1)
            Debug.Print "Class7.Invoke[HK_EnvironmentCompleted]"
            Debug.Print "  this        :", this
            Debug.Print "  errorCode   :", errorCode
            Debug.Print "  pEnvironment:", arg2
            Debug.Print "  m_refCount  :", m_refCount

        Case HK_ControllerCompleted
            Debug.Print "Class7.Invoke[HK_ControllerCompleted] (not supported)"

        Case HK_NavigationStarting
            Debug.Print "Class7.Invoke[HK_NavigationStarting] (not supported)"

        Case HK_NavigationCompleted
            Debug.Print "Class7.Invoke[HK_NavigationCompleted] (not supported)"

        Case HK_WebMessageReceived
            Debug.Print "Class7.Invoke[HK_WebMessageReceived] (not supported)"

        Case HK_DocumentTitleChanged
            Debug.Print "Class7.Invoke[HK_DocumentTitleChanged] (not supported)"

        Case Else
            Debug.Print "Class7.Invoke[unknown kind=" & m_kind & "]"
    End Select

    Handler_Invoke = 0   ' S_OK
End Function


' ============================================================
' Extract the lower 32 bits of a LongPtr as a signed Long
'
'   Under the x64 calling convention, when a 32-bit argument (HRESULT, etc.) is passed,
'   the upper 32 bits may be left undefined in rdx and so on.
'   The thunk shifts in 64 bits like "mov rdx, rcx",
'   so receiving it as a LongPtr as is leaves garbage in the upper 32 bits.
'
'   Extract the lower 32 bits and, if 0x80000000 or above, fold it into a negative Long
'   (CLng overflows outside the Long range, so we handle it ourselves).
' ============================================================
Private Function LongPtrLowDword(ByVal v As LongPtr) As Long
    Dim u As LongLong
    u = v And &HFFFFFFFF^
    If u >= &H80000000^ Then
        LongPtrLowDword = CLng(u - &H100000000^)
    Else
        LongPtrLowDword = CLng(u)
    End If
End Function


```
