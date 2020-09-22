Attribute VB_Name = "modServices"
Option Explicit

Public Function nSetFileAssociation(EXEPath As String, FileClass As String, AppDesc As String, Extention As String, Icon As String, IconNumber As Integer) As Boolean
    Dim tmp As String
    Dim Handle As Long
    Dim ret As Long
    
    ret = RegCreateKey(HKEY_CLASSES_ROOT, FileClass, Handle)
    If ret <> 0 Then GoTo ErrOut
    ret = RegSetValue(Handle, vbNullString, REG_SZ, AppDesc, 0)
    If ret <> 0 Then GoTo ErrOut
    
    ret = RegCreateKey(HKEY_CLASSES_ROOT, Extention, Handle)
    If ret <> 0 Then GoTo ErrOut
    ret = RegSetValue(Handle, vbNullString, REG_SZ, FileClass, 0)
    If ret <> 0 Then GoTo ErrOut
    
    ret = RegCreateKey(HKEY_CLASSES_ROOT, FileClass, Handle)
    If ret <> 0 Then GoTo ErrOut
    tmp = """" & EXEPath & """" & " ""%1"""
    ret = RegSetValue(Handle, "shell\open\command", REG_SZ, tmp, MAX_PATH)
    If ret <> 0 Then GoTo ErrOut
    
    ret = RegCreateKey(HKEY_CLASSES_ROOT, FileClass, Handle)
    If ret <> 0 Then GoTo ErrOut
    tmp = """" & Icon & """" & "," & IconNumber
    ret = RegSetValue(Handle, "DefaultIcon", REG_SZ, tmp, MAX_PATH)
    If ret <> 0 Then GoTo ErrOut
    
    nSetFileAssociation = True
    Exit Function
    
ErrOut:
    nSetFileAssociation = False
End Function
