VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'Vector mapping virtual line numbers (not counting discarded lines)
'to real ones (counting those lines).  Both are origin-0.
Private mlaRealIndexes() As Long

'Number of elements (lines) in this file.
'final int mlBufferedLines;
Private mlBufferedLines As Long

'Vector, indexed by line number, containing an equivalence code for
'each line.  It is this vector that is actually compared with that
'of another file to generate differences.
Private mlaEquivs() As Long

'Vector, like the previous one except that
'the elements for discarded lines have been squeezed out.
Private mlaUndiscarded() As Long

'Total number of nondiscarded lines.
Private mlNonDiscardedLines As Long

'Array, indexed by real origin-1 line number,
'containing true for a line that is an insertion or a deletion.
'The results of comparison are stored here.
Private mbaChangedFlag() As Boolean
Private mbNoDiscards As Boolean

Public Sub Clear()
    'Allocate a flag for each line of each file, saying whether that line
    'is an insertion or deletion.  Allocate an extra element, always zero,
    'at each end of each vector.

    ReDim mbaChangedFlag(mlBufferedLines + 1)
End Sub

'Return equiv_count[I] as the number of lines in this file
'that fall in equivalence class I.

Public Function EquivCount() As Long()
    Dim plaEquivCount() As Long
    Dim plIndex As Long
    
    ReDim plaEquivCount(glEquivMax - 1)
    
    For plIndex = 0 To mlBufferedLines - 1
        plaEquivCount(mlaEquivs(plIndex)) = plaEquivCount(mlaEquivs(plIndex)) + 1
    Next
    
    EquivCount = plaEquivCount
End Function

'Discard lines that have no matches in another file.
'
'A line which is discarded will not be considered by the actual
'comparison algorithm; it will be as if that line were not in the file.
'The file 's `realindexes' table maps virtual line numbers
'(which don't count the discarded lines) into real line numbers;
'this is how the actual comparison algorithm produces results
'that are comprehensible when the discarded lines are counted.
'
'When we discard a line, we also mark it as a deletion or insertion
'so that it will be printed in the output.
Public Sub DiscardConfusingLines(poOtherFile As FileData)
    Dim pyaDiscarded() As Byte
    
    Clear

    'Set up table of which lines are going to be discarded.
    pyaDiscarded = Discardable(poOtherFile.EquivCount())

    'Don't really discard the provisional lines except when they occur
    'in a run of discardables, with nonprovisionals at the beginning
    'and end.
    
    FilterDiscards pyaDiscarded
    
    Discard pyaDiscarded
End Sub

'Mark to be discarded each line that matches no line of another file.
'If a line matches many lines, mark it as provisionally discardable.
Private Function Discardable(plaCounts() As Long) As Byte()
    Dim plEnd As Long
    Dim pyaDiscards() As Byte
    Dim plMany As Long
    Dim plThreshold As Long
    Dim plIndex As Long
    Dim plMatches As Long
    
    plEnd = mlBufferedLines
    ReDim pyaDiscards(plEnd)
    plMany = 5
    plThreshold = plEnd / 64

    'Multiply MANY by approximate square root of number of lines.
    'That is the threshold for provisionally discardable lines.
    
    plThreshold = plThreshold * 2
    Do While plThreshold > 0
        plMany = plMany * 2
        plThreshold = plThreshold \ 2
    Loop
    
    For plIndex = 0 To plEnd - 1
        If mlaEquivs(plIndex) <> 0 Then
            plMatches = plaCounts(mlaEquivs(plIndex))
            If plMatches = 0 Then
                pyaDiscards(plIndex) = 1
            ElseIf (plMatches > plMany) Then
                pyaDiscards(plIndex) = 2
            End If
        End If
    Next
    Discardable = pyaDiscards
End Function

'       Don't really discard the plProvisional lines except when they occur
'       in a run of discardables, with nonprovisionals at the beginning
'       and end.
Private Sub FilterDiscards(pyaDiscards() As Byte)
    Dim plEnd As Long
    Dim plIndex As Long
    Dim plInnerIndex As Long
    Dim plLength As Long
    Dim plProvisional As Long
    Dim plConsec As Long
    Dim plMinimum As Long
    Dim plThreshold As Long
    
    plEnd = mlBufferedLines
    
    For plIndex = 0 To plEnd - 1
        If pyaDiscards(plIndex) = 2 Then
            pyaDiscards(plIndex) = 0
        ElseIf pyaDiscards(plIndex) <> 0 Then
            
            'We have found a nonprovisional discard.
            plProvisional = 0
            
            'Find end of this run of discardable lines.
            'Count how many are provisionally discardable.
            For plInnerIndex = plIndex To plEnd - 1
                If pyaDiscards(plInnerIndex) = 0 Then
                    Exit For
                End If
                If pyaDiscards(plInnerIndex) = 2 Then
                    plProvisional = plProvisional + 1
                End If
            Next

            'Cancel plProvisional pyaDiscards at end, and shrink the run.
            Do While plInnerIndex > plIndex And pyaDiscards(plInnerIndex - 1) = 2
                plInnerIndex = plInnerIndex - 1
                pyaDiscards(plInnerIndex) = 0
                plProvisional = plProvisional - 1
            Loop
            
            'Now we have the plLength of a run of discardable lines
            'whose first and last are not plProvisional.
            plLength = plInnerIndex - plIndex
            
            'If 1/4 of the lines in the run are plProvisional,
            'cancel discarding of all plProvisional lines in the run.
            If plProvisional * 4 > plLength Then
                Do While plInnerIndex > plIndex
                    plInnerIndex = plInnerIndex - 1
                    If pyaDiscards(plInnerIndex) = 2 Then
                        pyaDiscards(plInnerIndex) = 0
                    End If
                Loop
            Else
                
                
                plMinimum = 1
                plThreshold = plLength / 4
                
                'plMinimum is approximate square root of plLength/4.
                'A subrun of two or more provisionals can stand
                'when plLength is at least 16.
                'A subrun of 4 or more can stand when plLength >= 64.  */
                plThreshold = plThreshold / 2
                Do While (plThreshold > 0)
                    plMinimum = plMinimum * 2
                    plThreshold = plThreshold / 2
                Loop
                plMinimum = plMinimum + 1
                
                'Cancel any subrun of plMinimum or more provisionals
                'within the larger run.
                plConsec = 0
                For plInnerIndex = 0 To plLength - 1
                    If pyaDiscards(plIndex + plInnerIndex) <> 2 Then
                        plConsec = 0
                    Else
                        'Back up to start of subrun, to cancel it all.  */
                        plConsec = plConsec + 1
                        If plMinimum = plConsec Then
                            plInnerIndex = plInnerIndex - plConsec
                        ElseIf plMinimum < plConsec Then
                            pyaDiscards(plIndex + plInnerIndex) = 0
                        End If
                    End If
                Next
            
                'Scan from beginning of run
                'until we find 3 or more nonprovisionals in a row
                'or until the first nonprovisional at least 8 lines in.
                'Until that point, cancel any provisionals.
                
                plConsec = 0
                For plInnerIndex = 0 To plLength - 1
                    If plInnerIndex >= 8 And pyaDiscards(plIndex + plInnerIndex) = 1 Then
                        Exit For
                    End If
                    If pyaDiscards(plIndex + plInnerIndex) = 2 Then
                        plConsec = 0
                        pyaDiscards(plIndex + plInnerIndex) = 0
                    ElseIf pyaDiscards(plIndex + plInnerIndex) = 0 Then
                        plConsec = 0
                    Else
                        plConsec = plConsec + 1
                    End If
                
                    If plConsec = 3 Then
                        Exit For
                    End If
                Next
                plIndex = plIndex + plLength - 1
            
                'Same thing, from end.
                plConsec = 0
                For plInnerIndex = 0 To plLength - 1
                    If plInnerIndex >= 0 And pyaDiscards(plIndex - plInnerIndex) = 1 Then
                        Exit For
                    End If
                    If pyaDiscards(plIndex - plInnerIndex) = 2 Then
                        plConsec = 0
                        pyaDiscards(plIndex - plInnerIndex) = 0
                    ElseIf pyaDiscards(plIndex - plInnerIndex) = 0 Then
                        plConsec = 0
                    Else
                        plConsec = plConsec + 1
                    End If
                    If plConsec = 3 Then
                        Exit For
                    End If
                Next
            End If
        End If
    Next
End Sub

'Actually discard the lines.
Private Sub Discard(pyaDiscards() As Byte)
    Dim plEnd As Long
    Dim plIndex As Long
    Dim plMapIndex As Long
    
    plEnd = mlBufferedLines
    plMapIndex = 0
    
    For plIndex = 0 To plEnd - 1
        If mbNoDiscards Or pyaDiscards(plIndex) = 0 Then
            mlaUndiscarded(plMapIndex) = mlaEquivs(plIndex)
            mlaRealIndexes(plMapIndex) = plIndex
            plMapIndex = plMapIndex + 1
        Else
            mbaChangedFlag(plIndex + 1) = True
        End If
        
        mlNonDiscardedLines = plMapIndex
    Next
End Sub

Public Sub Init(psaLines() As String, poDictionary As Dictionary, pbNoDiscards As Boolean)
    Dim plIndex As Long
    
    mbNoDiscards = pbNoDiscards
    mlBufferedLines = UBound(psaLines) + 1
    
    ReDim mlaEquivs(-1 To mlBufferedLines - 1)
    ReDim mlaUndiscarded(-1 To mlBufferedLines - 1)
    ReDim mlaRealIndexes(-1 To mlBufferedLines - 1)
    
    For plIndex = 0 To UBound(psaLines)
        If poDictionary.Exists(psaLines(plIndex)) Then
            mlaEquivs(plIndex) = poDictionary(psaLines(plIndex))
        Else
            poDictionary(psaLines(plIndex)) = glEquivMax
            mlaEquivs(plIndex) = glEquivMax
            glEquivMax = glEquivMax + 1
        End If
    Next
End Sub

'Adjust inserts/deletes of blank lines to join changes
'as much as possible.
'
'We do something when a run of changed lines include a blank
'line at one end and have an excluded blank line at the other.
'We are free to choose which blank line is included.
''compareseq' always chooses the one at the beginning,
'but usually it is cleaner to consider the following blank line
'to be the "Change".  The only exception is if the plPreceding blank line
'would join this Change to other changes.
Public Sub ShiftBoundaries(poOtherFile As FileData)
    Dim pbaOtherChanged() As Boolean
    Dim plIndex As Long
    Dim plOtherIndex As Long
    Dim plLastLine As Long
    Dim plPreceding As Long
    Dim plOtherPreceding As Long
    Dim plStart As Long
    Dim plEnd As Long
    Dim plOtherStart As Long
    
    pbaOtherChanged = poOtherFile.GetChangedFlagArray()
    plIndex = 0
    plOtherIndex = 0
    plLastLine = mlBufferedLines
    plPreceding = -1
    plOtherPreceding = -1
    
    Do While True
        
        'Scan forwards to find beginning of another run of changes.
        'Also keep track of the corresponding point in the other file.
        Do While plIndex < plLastLine And Not mbaChangedFlag(plIndex + 1)
            Do While pbaOtherChanged(plOtherIndex + 1)
                plOtherIndex = plOtherIndex + 1
                plOtherPreceding = plOtherIndex
            Loop
            plOtherIndex = plOtherIndex + 1
            plIndex = plIndex + 1
        Loop
        
        If plIndex = plLastLine Then
            Exit Do
        End If
        
        plStart = plIndex
        plOtherStart = plOtherIndex
        
        Do While True
            'Now find the end of this run of changes.
            Do While plIndex < plLastLine And mbaChangedFlag(plIndex + 1)
                plIndex = plIndex + 1
            Loop
            plEnd = plIndex
            
            'If the first changed line matches the following unchanged one,
            'and this run does not follow right after a previous run,
            'and there are no lines deleted from the other file here,
            'then classify the first changed line as unchanged
            'and the following line as changed in its place.
            '
            'You might ask, how could this run follow right after another?
            'Only because the previous run was shifted here.
            If (plEnd <> plLastLine _
                And Not pbaOtherChanged(plOtherIndex + 1) _
                And plEnd <> plLastLine _
                And Not ((plPreceding >= 0 And plStart = plPreceding) _
                    Or (plOtherPreceding >= 0 _
                    And plOtherStart = plOtherPreceding))) Then
                    
                If Not mlaEquivs(plStart) = mlaEquivs(plEnd) Then
                    Exit Do
                End If
                
                mbaChangedFlag(plEnd + 1) = True
                plEnd = plEnd + 1
                mbaChangedFlag(plStart + 1) = False
                plStart = plStart + 1
                
                'Since one line-that-matches is now before this run
                'instead of after, we must advance in the other file
                'to keep in sync.
                plIndex = plIndex + 1
                plOtherIndex = plOtherIndex + 1
            Else
                Exit Do
            End If
        Loop
        plPreceding = plIndex
        plOtherPreceding = plOtherIndex
    Loop
End Sub

Public Property Get GetUndiscardedArray() As Long()
    GetUndiscardedArray = mlaUndiscarded
End Property

Public Property Get RealIndexes(plIndex As Long) As Long
    RealIndexes = mlaRealIndexes(plIndex)
End Property

Public Function GetChangedFlagArray() As Boolean()
    GetChangedFlagArray = mbaChangedFlag()
End Function

Public Sub SetChangedFlag(plIndex As Long, pbValue As Boolean)
    mbaChangedFlag(plIndex) = pbValue
End Sub

Public Property Get NonDiscaredLines() As Long
    NonDiscaredLines = mlNonDiscardedLines
End Property

Public Property Get BufferedLines() As Long
    BufferedLines = mlBufferedLines
End Property

Public Property Let NoDiscards(pbNewValue As Boolean)
    mbNoDiscards = pbNewValue
End Property
