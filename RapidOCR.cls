VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RapidOCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    Copyright
'    Copyright 2024 Danysys. <hello@danysys.com>
'
'    Licensed under the MIT license.
'    See LICENSE file or go to https://opensource.org/licenses/MIT for details.
'    Copyright
'
'    Information
'    Author(s)......: DanysysTeam (Danyfirex & Dany3j)
'    Description....: VBA Wrapper for RapidOCR: A library that empowers VBA users to extract text from images using the robust RapidOCR engine.
'    Version........: 1.0.0
'    Thanks.........:
'                    https://github.com/RapidAI/RapidOCR
'                    Schmidt -  cdeclCallA
'                    John_w  -  UTF8StringToVBAString
'    Information


Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As LongPtr, ByVal paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As Any) As LongPtr
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As LongPtr) As LongPtr
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwflags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwflags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Enum LongPtr
    [_]
End Enum

Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As LongPtr, ByVal paValues As LongPtr, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwflags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwflags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If


Private Type RAPIDOCR_PARAM
    Padding As Long
    MaxSideLen As Long
    BoxScoreThresh As Single
    BoxThresh As Single
    UnClipRatio As Single
    FlagDoAngle As Long
    FlagMostAngle As Long
End Type

Private tRAPIDOCR_PARAM As RAPIDOCR_PARAM

Private Const CP_UTF8 = 65001
Private Const CC_CDECL = 1&

Private VType(0 To 63) As Integer
Private VPtr(0 To 63) As LongPtr
Private LibHdls As New Collection                'Cache


Private IsInitialized As Boolean
Private IsModelsLoaded As Boolean

Private PathDefaultModel As String
Private PathDetModel As String
Private PathClsModel As String
Private PathRecModel As String
Private PathKey As String

Private PathLibRapidOCR As String
Private HRapidOCR As LongPtr

Private Sub Initialize()
    'define Lib Path
    #If Win64 Then
        PathLibRapidOCR = ThisWorkbook.Path & "\RapidOCR\x64\RapidOcrOnnx.dll"
    #Else
        PathLibRapidOCR = ThisWorkbook.Path & "\RapidOCR\x86\RapidOcrOnnx.dll"
    #End If
    
    IsInitialized = True
End Sub

Private Sub LoadDefaultModels()

    'Initialize Models Paths
    PathDefaultModel = ThisWorkbook.Path & "\RapidOCR\onnx-models\"
    PathDetModel = PathDefaultModel & "ch_PP-OCRv3_det_infer.onnx"
    PathClsModel = PathDefaultModel & "ch_ppocr_mobile_v2.0_cls_infer.onnx"
    PathRecModel = PathDefaultModel & "ch_PP-OCRv3_rec_infer.onnx"
    PathKey = PathDefaultModel & "ppocr_keys_v1.txt"
    SetSettings
    LoadModels PathDefaultModel, PathDetModel, PathClsModel, PathRecModel, PathKey
    
End Sub

Public Sub SetSettings(Optional ByVal Padding As Long = 50, Optional ByVal MaxSideLen As Long = 1024, Optional ByVal BoxScoreThresh As Single = 0.6, _
                       Optional ByVal BoxThresh As Single = 0.3, Optional ByVal UnClipRatio As Single = 2#, Optional ByVal FlagDoAngle As Long = 1, Optional ByVal FlagMostAngle As Long = 1)

    tRAPIDOCR_PARAM.Padding = Padding
    tRAPIDOCR_PARAM.MaxSideLen = MaxSideLen
    tRAPIDOCR_PARAM.BoxScoreThresh = BoxScoreThresh
    tRAPIDOCR_PARAM.BoxThresh = BoxThresh
    tRAPIDOCR_PARAM.UnClipRatio = UnClipRatio
    tRAPIDOCR_PARAM.FlagDoAngle = FlagDoAngle
    tRAPIDOCR_PARAM.FlagMostAngle = FlagMostAngle
   
End Sub

Public Sub LoadModels(ByVal PathModels As String, ByVal PathDetMode As String, ByVal PathClsModel As String, ByVal PathRecModel As String, ByVal PathKey As String)
   
    If Not IsInitialized Then
        Initialize
    End If
   
    Dim vbTypeResult As VbVarType
    #If Win64 Then
        vbTypeResult = vbLongLong
    #Else
        vbTypeResult = vbLong
    #End If
    
    HRapidOCR = cdeclCallA(PathLibRapidOCR, "OcrInit", vbTypeResult, PathDetMode, PathClsModel, PathRecModel, PathKey, 1)
    If HRapidOCR = 0 Then Err.Raise 1, , "Error OcrInit"
    IsModelsLoaded = True
End Sub

Public Function ImageToText(ByVal PathImage As String) As String
    Dim Ret As Long
    
    If Not IsInitialized Then
        Initialize
    End If
    
    If Not IsModelsLoaded Then
        LoadDefaultModels
    End If
    
    If Dir(PathImage) = "" Then Err.Raise vbObjectError, , "Image not found: " & PathImage
    Dim Directory As String: Directory = Left(PathImage, InStrRev(PathImage, "\"))
    Dim ImageName As String: ImageName = Right(PathImage, Len(PathImage) - InStrRev(PathImage, "\"))
    

    Ret = cdeclCallA(PathLibRapidOCR, "OcrDetect", vbLong, HRapidOCR, Directory, ImageName, VarPtr(tRAPIDOCR_PARAM))
    Ret = cdeclCallA(PathLibRapidOCR, "OcrGetLen", vbLong, HRapidOCR)
    Dim Result As String
    If Ret Then
        Result = String$(Ret, vbNullChar)
        Ret = cdeclCallA(PathLibRapidOCR, "OcrGetResult", vbLong, HRapidOCR, Result, Ret)
        Result = UTF8StringToVBAString(Result)
    End If
    
    ImageToText = Result
    
End Function

'Utils
Public Function MsgBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "") As VbMsgBoxResult
    MsgBox = MessageBoxW(Application.hWnd, StrPtr(Prompt), StrPtr(Title), Buttons)
End Function

Public Function SaveResultToFile(ByVal FilePath As String, ByRef Result As String) As Boolean
    Dim bytes() As Byte
    bytes = VBAStringToUFT8Bytes(Result)

    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.charset = "utf-8"
    oStream.Type = 1
    oStream.Write bytes
    oStream.SaveToFile FilePath, 2               '2 = overwrite
    oStream.Close
End Function

Private Function VBAStringToUFT8Bytes(ByVal VBAString As String) As Byte()
    Dim bufferSize As Long
    Dim Buffer() As Byte

    If VBAString = "" Then
        Exit Function
    End If
    
    bufferSize = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(VBAString), -1, 0&, 0&, 0&, 0&)
    ReDim Buffer(bufferSize) As Byte
    Call WideCharToMultiByte(CP_UTF8, 0&, StrPtr(VBAString), -1, VarPtr(Buffer(0)), bufferSize, 0&, 0&)

    VBAStringToUFT8Bytes = Buffer
End Function

'https://www.mrexcel.com/board/threads/open-csv-files-and-show-spanish-characters-utf-8-using-vba.1205822/#post-5888866 - John_w
Private Function UTF8StringToVBAString(ByRef UTF8string As String) As String
    
    Dim UTF8bytes() As Byte
    Dim bufferSize As Long
    
    UTF8bytes = StrConv(UTF8string, vbFromUnicode)
    
    'Get required size of output string
    bufferSize = MultiByteToWideChar(CP_UTF8, 0, VarPtr(UTF8bytes(0)), UBound(UTF8bytes) + 1, 0, 0)
    
    'Allocate output string
    UTF8StringToVBAString = String$(bufferSize, 0)
    
    'Convert UTF8 bytes to Unicode output string
    MultiByteToWideChar CP_UTF8, 0, VarPtr(UTF8bytes(0)), UBound(UTF8bytes) + 1, StrPtr(UTF8StringToVBAString), bufferSize

End Function

'https://www.vbforums.com/showthread.php?789217-C-DLL-to-VB6&p=5505940&viewfull=1#post5505940 - Schmidt
Private Function cdeclCallA(sDll As String, sFunc As String, ByVal RetType As VbVarType, ParamArray P() As Variant) As Variant
    Dim i As Long, pFunc As LongPtr, V() As Variant, HRes As Long
 
    V = P                                        'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
    For i = 0 To UBound(V)
        If VarType(P(i)) = vbString Then P(i) = StrConv(P(i), vbFromUnicode): V(i) = StrPtr(P(i))
        VType(i) = VarType(V(i))
        VPtr(i) = VarPtr(V(i))
    Next i
  
    pFunc = GetFuncPtr(sDll, sFunc)
    HRes = DispCallFunc(0, pFunc, CC_CDECL, RetType, i, VarPtr(VType(0)), VarPtr(VPtr(0)), cdeclCallA)
  
    For i = 0 To UBound(P)                       'back-conversion of the ANSI-String-Results
        If VarType(P(i)) = vbString Then P(i) = StrConv(P(i), vbUnicode)
    Next i
    If HRes Then Err.Raise HRes
End Function

Private Function GetFuncPtr(sDll As String, sFunc As String) As LongPtr
    Static hLib As LongPtr, sLib As String
    If sLib <> sDll Then                         'just a bit of caching, to make resolving libHdls faster
        sLib = sDll
        On Error Resume Next
        hLib = 0
        hLib = LibHdls(sLib)
        On Error GoTo 0
    
        If hLib = 0 Then
            hLib = LoadLibrary(StrPtr(sLib))
            If hLib = 0 Then Err.Raise vbObjectError, , "Dll not found (or loadable): " & sLib
            LibHdls.Add hLib, sLib               '<- cache it under the dll-name for the next call
        End If
    End If
    GetFuncPtr = GetProcAddress(hLib, sFunc)
    If GetFuncPtr = 0 Then Err.Raise 453, , "EntryPoint not found: " & sFunc & " in: " & sLib
End Function


