VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DiffWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    author Andrew Lippitt, translated from Stuart Gathman's Java Version
'    Stuart D. Gathman, translated from GNU diff 1.15

'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; version 2.

'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal ppDest As Long, ByVal ppSource As Long, ByVal plBytes As Long)

'Roughly 2 megs of null string pointers in an array
Private Const UNREASONABLE_LENGTH As Long = 500000

Private Type DiffRecord
    Inserts As Long
    Deletes As Long
    Block0 As Long
    Block1 As Long
End Type

Public Function GetDiffs(pyFile0() As Byte, pyFile1() As Byte) As Byte()
    Dim plSize As Long
    Dim poDiff As Diff
    Dim poScript As change
    Dim poTmpScript As change
    Dim puDiffRecord As DiffRecord
    Dim plChangeCount As Long
    Dim pyReturn() As Byte
    Dim ppCrawl As Long
    Dim ppDiffRecord As Long
    Dim plDiffRecordLength As Long
    Dim psaFile0() As String
    Dim psaFile1() As String
    Dim plIndex As Long
    
    Set poDiff = New Diff
    
    psaFile0 = Split(pyFile0, vbCrLf)
    psaFile1 = Split(pyFile1, vbCrLf)
    
    poDiff.Init psaFile0, psaFile1
    
    Set poScript = poDiff.Diff(False)
    
    Set poTmpScript = poScript
    Do While Not poTmpScript Is Nothing
        plChangeCount = plChangeCount + 1
        For plIndex = 0 To poTmpScript.Inserted - 1
            plSize = plSize + Len(psaFile1(poTmpScript.plLine1 + plIndex)) * 2 + 4
        Next
        Set poTmpScript = poTmpScript.Link
    Loop
    
    plSize = plSize + 4 + (plChangeCount * Len(puDiffRecord))
    
    ReDim pyReturn(plSize - 1)
    
    ppCrawl = VarPtr(pyReturn(0))
    
    ppDiffRecord = VarPtr(puDiffRecord)
    plDiffRecordLength = Len(puDiffRecord)
    
    CopyMemory ppCrawl, VarPtr(CLng(UBound(psaFile1))), 4
    ppCrawl = ppCrawl + 4
    
    Set poTmpScript = poScript
    Do While Not poTmpScript Is Nothing
        puDiffRecord.Inserts = poTmpScript.Inserted
        puDiffRecord.Deletes = poTmpScript.Deleted
        puDiffRecord.Block0 = poTmpScript.plLine0
        puDiffRecord.Block1 = poTmpScript.plLine1
        
        CopyMemory ppCrawl, ppDiffRecord, plDiffRecordLength
        ppCrawl = ppCrawl + plDiffRecordLength
        
        For plIndex = 0 To puDiffRecord.Inserts - 1
            CopyMemory ppCrawl, VarPtr(CLng(Len(psaFile1(poTmpScript.plLine1 + plIndex)) * 2)), 4
            ppCrawl = ppCrawl + 4
            CopyMemory ppCrawl, StrPtr(psaFile1(poTmpScript.plLine1 + plIndex)), Len(psaFile1(poTmpScript.plLine1 + plIndex)) * 2
            ppCrawl = ppCrawl + Len(psaFile1(poTmpScript.plLine1 + plIndex)) * 2
        Next
        
        Set poTmpScript = poTmpScript.Link
    Loop
    
    GetDiffs = pyReturn
End Function

Public Function ApplyDiffs(pyFile0() As Byte, pyDiffs() As Byte) As Byte()
    Dim psaFile0() As String
    Dim puDiffRecord As DiffRecord
    Dim ppDiffRecord As Long
    Dim plDiffRecordLength As Long
    Dim ppCrawl As Long
    Dim ppDiffBase As Long
    Dim plDiffLastByte As Long
    Dim psTemp As String
    Dim plIndex As Long
    Dim plNewSize As Long
    Dim plNextIndex As Long
    Dim plCopyIndex As Long
    Dim psaReturn() As String
    Dim plStringLength As Long
    Dim plNewFileIndex As Long
    Dim pyReturn() As Byte
    
    psaFile0 = Split(pyFile0, vbCrLf)
    
    ppDiffRecord = VarPtr(puDiffRecord)
    plDiffRecordLength = Len(puDiffRecord)
    ppDiffBase = VarPtr(pyDiffs(0))
    plDiffLastByte = UBound(pyDiffs)
    
    plNewSize = UBound(psaFile0)
    
    plNextIndex = 0
    ppCrawl = 0
    
    If ppCrawl + 4 > plDiffLastByte + 1 Then
        Err.Raise vbObjectError, "DiffWrapper", "Invalid Diff Format"
    End If
    CopyMemory VarPtr(plNewSize), ppDiffBase + ppCrawl, 4
    ppCrawl = ppCrawl + 4
    
    If plNewSize > UNREASONABLE_LENGTH Then
        Err.Raise vbObjectError, "DiffWrapper", "Invalid Diff Format - an unreasonable file line count was specified"
    End If
    If plNewSize >= 0 Then
        ReDim psaReturn(plNewSize)
        
        Do While ppCrawl < plDiffLastByte
            If ppCrawl + plDiffRecordLength > plDiffLastByte + 1 Then
                Err.Raise vbObjectError, "DiffWrapper", "Invalid Diff Format"
            End If
            CopyMemory ppDiffRecord, ppDiffBase + ppCrawl, plDiffRecordLength
            ppCrawl = ppCrawl + plDiffRecordLength
            
            Debug.Assert plNextIndex <= puDiffRecord.Block0
            
            'Fast forward to the right spot
            For plNextIndex = plNextIndex To puDiffRecord.Block0 - 1
                psaReturn(plNewFileIndex) = psaFile0(plNextIndex)
                plNewFileIndex = plNewFileIndex + 1
            Next
            
            'Skip deleted records in source file
            plNextIndex = plNextIndex + puDiffRecord.Deletes
            
            'Start inserting lines
            For plIndex = 0 To puDiffRecord.Inserts - 1
                If ppCrawl + 4 > plDiffLastByte + 1 Then
                    Err.Raise vbObjectError, "DiffWrapper", "Invalid Diff Format"
                End If
                CopyMemory VarPtr(plStringLength), ppDiffBase + ppCrawl, 4
                ppCrawl = ppCrawl + 4
                psTemp = Space(plStringLength \ 2)
                If ppCrawl + plStringLength > plDiffLastByte + 1 Then
                    Err.Raise vbObjectError, "DiffWrapper", "Invalid Diff Format"
                End If
                CopyMemory StrPtr(psTemp), ppDiffBase + ppCrawl, plStringLength
                ppCrawl = ppCrawl + plStringLength
                psaReturn(plNewFileIndex) = psTemp
                plNewFileIndex = plNewFileIndex + 1
            Next
        Loop
        
        'Finish with the rest of the file after the change
        For plNextIndex = plNextIndex To UBound(psaFile0)
            psaReturn(plNewFileIndex) = psaFile0(plNextIndex)
            plNewFileIndex = plNewFileIndex + 1
        Next
        pyReturn = Join(psaReturn, vbCrLf)
    End If
    ApplyDiffs = pyReturn
End Function

'This function will be used to double check the diff algo
'to avoid loss of data until I'm confident that it's bug free
Public Function CheckDiff(pyFile0() As Byte, pyFile1() As Byte, pyDiffs() As Byte) As Boolean
    Dim pyResultantFile() As Byte
    Dim plIndex As Long
    
    pyResultantFile = ApplyDiffs(pyFile0, pyDiffs)
    
    If Not IsArrayInitialized(pyResultantFile) Then
        If IsArrayInitialized(pyFile1) Then
            Exit Function
        Else
            CheckDiff = True
        End If
    Else
        If IsArrayInitialized(pyFile1) Then
            If UBound(pyResultantFile) = UBound(pyFile1) Then
                For plIndex = 0 To UBound(pyResultantFile)
                    If pyResultantFile(plIndex) <> pyFile1(plIndex) Then
                        Exit Function
                    End If
                Next
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    CheckDiff = True
End Function

Private Function IsArrayInitialized(pyFile() As Byte) As Boolean
    Dim plUBound As Long
    On Error Resume Next
    plUBound = UBound(pyFile)
    If Err.Number = 0 Then
        IsArrayInitialized = plUBound <> -1
    End If
End Function
