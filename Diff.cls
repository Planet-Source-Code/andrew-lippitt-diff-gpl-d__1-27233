VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Diff"
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

Private moFileVec(1) As FileData

'When set to true, the comparison uses a heuristic to speed it up.
'With this heuristic, for files with a constant small density
'of changes, the algorithm is linear in the file size.
Private mbHeuristic As Boolean

'When set to true, the algorithm returns a guarranteed minimal
'set of changes.  This makes things slower, sometimes much slower.
Private mbNoDiscards As Boolean

Private mlaXVector() As Long        'Vectors being compared
Private mlaYVector() As Long

Private mlaForwardDiag() As Long    'Vector, indexed by diagonal, containing
                                    'the plX coordinate of the point furthest
                                    'along the given diagonal in the forward
                                    'search of the edit matrix.
                                    
Private mlaBackwardDiag() As Long   'Vector, indexed by diagonal, containing
                                    'the plX coordinate of the point furthest
                                    'along the given diagonal in the backward
                                    'search of the edit matrix.
                                    
Private mlaForwardDiagOffset As Long, mlaBackwardDiagOffset As Long

Private mlCost As Long

Private Sub Class_Initialize()
    glEquivMax = 1
    Heuristic = False
End Sub

'Prepare to find differences between two arrays.  Each element of
'the arrays is translated to an "equivalence number" based on
'the result of <code>equals</code>.  The original Object arrays
'are no longer needed for computing the differences.  They will
'be needed again later to print the results of the comparison as
'an edit poScript, if desired.
Public Sub Init(psaFile0() As String, psaFile1() As String)
    Dim poDictionary As Dictionary
    
    Set poDictionary = New Dictionary
    
    Set moFileVec(0) = New FileData
    Set moFileVec(1) = New FileData
    moFileVec(0).Init psaFile0, poDictionary, mbNoDiscards
    moFileVec(1).Init psaFile1, poDictionary, mbNoDiscards
    
End Sub

'Find the midpoint of the shortest edit poScript for a specified
'portion of the two files.
'
'We scan from the beginnings of the files, and simultaneously from the ends,
'doing a breadth-first search through the space of edit-sequence.
'When the two searches meet, we have found the midpoint of the shortest
'edit sequence.
'
'The value returned is the number of the diagonal on which the midpoint lies.
'The diagonal number equals the number of inserted lines minus the number
'of deleted lines (counting only lines before the midpoint).
'The edit cost is stored into COST; this is the total number of
'lines inserted or deleted (counting only lines before the midpoint).
'
'This function assumes that the first lines of the specified portions
'of the two files do not match, and likewise that the last lines do not
'match.  The caller must trim matching lines from the beginning and end
'of the portions it is going to specify.
'
'Note that if we return the "wrong" diagonal value, or if
'the value of mlaBackwardDiag at that diagonal is "wrong",
'the worst this can do is cause suboptimal diff output.
'It cannot cause incorrect diff output.
Private Function GetDiagonal(plXVectorOffset As Long, plXVectorLimit As Long, plYVectorOffset As Long, plYVectorLimit As Long) As Long
    Dim plDiagMin As Long       'Minimum valid diagonal
    Dim plDiagMax As Long       'Maximum valid diagonal
    Dim plForwardMid As Long    'Center diagonal of top-down search
    Dim plBackwardMid As Long   'Center diagonal of bottom-up search
    Dim plForwardMin As Long    'Limits of top-down search.
    Dim plForwardMax As Long
    Dim plBackwardMin As Long   'Limits of bottom-up search.
    Dim plBackwardMax As Long
    Dim pbOdd As Boolean        'True if southeast corner is on an pbOdd diagonal with respect to the northwest.
    Dim plCost As Long
    Dim plDiag As Long             'Active diagonal
    Dim pbBigSnake As Boolean   'True if southeast corner is on an pbOdd
                                'diagonal with respect to the northwest.
    Dim plX As Long
    Dim plY As Long
    Dim plOldX As Long
    Dim plLow As Long
    Dim plHigh As Long
    
    Dim plBest As Long
    Dim plBestPos As Long
                
    Dim plDiagDelta As Long
    Dim plIndex As Long
    
    plDiagMin = plXVectorOffset - plYVectorLimit
    plDiagMax = plXVectorLimit - plYVectorOffset
    plForwardMid = plXVectorOffset - plYVectorOffset
    plBackwardMid = plXVectorLimit - plYVectorLimit
    
    plForwardMin = plForwardMid
    plForwardMax = plForwardMid
    
    plBackwardMin = plBackwardMid
    plBackwardMax = plBackwardMid
    
    pbOdd = (plForwardMid - plBackwardMid And 1) <> 0
    
    mlaForwardDiag(mlaForwardDiagOffset + plForwardMid) = plXVectorOffset
    mlaBackwardDiag(mlaBackwardDiagOffset + plBackwardMid) = plXVectorLimit
    
    plCost = 1
    Do While True
        'Extend the top-down search by an edit step in each diagonal
        
        If plForwardMin > plDiagMin Then
            plForwardMin = plForwardMin - 1
            mlaForwardDiag(mlaForwardDiagOffset + plForwardMin - 1) = -1
        Else
            plForwardMin = plForwardMin + 1
        End If
        If plForwardMax < plDiagMax Then
            plForwardMax = plForwardMax + 1
            mlaForwardDiag(mlaForwardDiagOffset + plForwardMax + 1) = -1
        Else
            plForwardMax = plForwardMax - 1
        End If
        
        plDiag = plForwardMax
        Do While plDiag >= plForwardMin
            plLow = mlaForwardDiag(mlaForwardDiagOffset + plDiag - 1)
            plHigh = mlaForwardDiag(mlaForwardDiagOffset + plDiag + 1)
            
            If plLow >= plHigh Then
                plX = plLow + 1
            Else
                plX = plHigh
            End If
            
            plOldX = plX
            
            plY = plX - plDiag
            Do While plX < plXVectorLimit And plY < plYVectorLimit
                If Not mlaXVector(plX) = mlaYVector(plY) Then
                    Exit Do
                End If
                plX = plX + 1
                plY = plY + 1
            Loop
            
            If plX - plOldX > 20 Then
                pbBigSnake = True
            End If
            mlaForwardDiag(mlaForwardDiagOffset + plDiag) = plX
            If pbOdd And plBackwardMin <= plDiag And plDiag <= plBackwardMax And mlaBackwardDiag(mlaBackwardDiagOffset + plDiag) <= mlaForwardDiag(mlaForwardDiagOffset + plDiag) Then
                mlCost = 2 * plCost - 1
                GetDiagonal = plDiag
                Exit Function
            End If
            
            plDiag = plDiag - 2
        Loop
        
        'Similar extend the bottom-up search
        
        If plBackwardMin > plDiagMin Then
            plBackwardMin = plBackwardMin - 1
            mlaBackwardDiag(mlaBackwardDiagOffset + plBackwardMin - 1) = &H7FFFFFFF
        Else
            plBackwardMin = plBackwardMin + 1
        End If
        If plBackwardMax < plDiagMax Then
            plBackwardMax = plBackwardMax + 1
            mlaBackwardDiag(mlaBackwardDiagOffset + plBackwardMax + 1) = &H7FFFFFFF
        Else
            plBackwardMax = plBackwardMax - 1
        End If
        
        plDiag = plBackwardMax
        Do While plDiag >= plBackwardMin
            
            plLow = mlaBackwardDiag(mlaBackwardDiagOffset + plDiag - 1)
            plHigh = mlaBackwardDiag(mlaBackwardDiagOffset + plDiag + 1)
            
            If plLow < plHigh Then
                plX = plLow
            Else
                plX = plHigh - 1
            End If
            plOldX = plX
            plY = plX - plDiag
            Do While plX > plXVectorOffset And plY > plYVectorOffset
                If Not (mlaXVector(plX - 1) = mlaYVector(plY - 1)) Then
                    Exit Do
                End If
                plX = plX - 1
                plY = plY - 1
            Loop
            If plOldX - plX > 20 Then
                pbBigSnake = True
            End If
            mlaBackwardDiag(mlaBackwardDiagOffset + plDiag) = plX
            If (Not pbOdd) And plForwardMin <= plDiag And plDiag <= plForwardMax And mlaBackwardDiag(mlaBackwardDiagOffset + plDiag) <= mlaForwardDiag(mlaForwardDiagOffset + plDiag) Then
                mlCost = 2 * plCost
                GetDiagonal = plDiag
                Exit Function
            End If
            
            plDiag = plDiag - 2
        Loop
        
        'Heuristic: check occasionally for a diagonal that has made
        'lots of progress compared with the edit distance.
        'If we have any such, find the one that has made the most
        'progress and return it as if it had succeeded.
        '
        'With this heuristic, for files with a constant small density
        'of changes, the algorithm is linear in the file size.
        If plCost > 200 And pbBigSnake And Heuristic Then
            plBest = 0
            plBestPos = -1
            
            plDiag = plForwardMax
            Do While plDiag >= plForwardMin
                plDiagDelta = plDiag - plForwardMid
                
                If (mlaForwardDiag(mlaForwardDiagOffset + plDiag) - plXVectorOffset) * 2 - plDiagDelta > 12 * (plCost + (IIf(plDiagDelta > 0, plDiagDelta, -plDiagDelta))) Then
                    If mlaForwardDiag(mlaForwardDiagOffset + plDiag) * 2 - plDiagDelta > plBest And _
                        mlaForwardDiag(mlaForwardDiagOffset + plDiag) - plXVectorOffset > 20 And _
                        mlaForwardDiag(mlaForwardDiagOffset + plDiag) - plDiag - plYVectorOffset > 20 Then
                    
                        
                        plX = mlaForwardDiag(mlaForwardDiagOffset + plDiag)
                        
                        'We have a good enough plBest diagonal, now insist that it end
                        'with a significant snake
                        
                        For plIndex = 1 To 21
                            If mlaXVector(plX - plIndex) <> mlaYVector(plX - plDiag - plIndex) Then
                                Exit For
                            End If
                        Next
                        
                        If plIndex = 21 Then
                            plBest = mlaForwardDiag(mlaForwardDiagOffset + plDiag) * 2 - plDiagDelta
                            plBestPos = plDiag
                        End If
                    End If
                End If
                
                If plBest > 0 Then
                    mlCost = 2 & plCost - 1
                    GetDiagonal = plBestPos
                    Exit Function
                End If
                            
                plBest = 0
                plDiag = plBackwardMax
                Do While plDiag >= plBackwardMin
                    plDiagDelta = plDiag - plBackwardMid
                    If (plXVectorLimit - mlaBackwardDiag(mlaBackwardDiagOffset + plDiag)) * 2 + plDiagDelta > 12 * (plCost + (IIf(plDiagDelta > 0, plDiagDelta, -plDiagDelta))) Then
                        If (plXVectorLimit - mlaBackwardDiag(mlaBackwardDiagOffset + plDiag)) * 2 + plDiagDelta > plBest And _
                            plXVectorLimit - mlaBackwardDiag(mlaBackwardDiagOffset + plDiag) > 20 And _
                            plYVectorLimit - (mlaBackwardDiag(mlaBackwardDiagOffset + plDiag) - plDiag) > 20 Then
                        
                            'We have a good enough plBest diagonal, now insist that it end with a significant snake
                            plX = mlaBackwardDiag(mlaBackwardDiagOffset + plDiag)
                            For plIndex = 0 To 19
                                'TODO: what to do if subscript out of range on this next
                                'statement.  Find out if the original java would have dealt
                                'with it, or if I have a bug in the port.
                                If mlaXVector(plX + plIndex) <> mlaYVector(plX - plDiag + plIndex) Then
                                    Exit For
                                End If
                            Next
                            
                            If plIndex = 20 Then
                                plBest = (plXVectorLimit - mlaBackwardDiag(mlaBackwardDiagOffset + plDiag)) * 2 + plDiagDelta
                                plBestPos = plDiag
                            End If
                        End If
                    End If
                            
                    plDiag = plDiag - 2
                Loop

                If plBest > 0 Then
                    mlCost = 2 * plCost - 1
                    GetDiagonal = plBestPos
                    Exit Function
                End If
                
                plDiag = plDiag - 2
            Loop
        End If
        plCost = plCost + 1
    Loop
End Function



'Compare in detail contiguous subsequences of the two files
'which are known, as a whole, to match each other.
'
'The results are recorded in the vectors filevec[N].changed_flag, by
'storing a 1 in the element for each line that is an insertion or deletion.
'
'The subsequence of file 0 is (plXVectorOffset, plXVectorLimit) and likewise for file 1.
'
'Note that plXVectorLimit, plYVectorLimit are exclusive bounds.
'All line numbers are origin-0 and discarded lines are not counted.
Private Sub CompareSequence(plXVectorOffset As Long, plXVectorLimit As Long, plYVectorOffset As Long, plYVectorLimit As Long)
    Dim plDiag As Long
    Dim plCost As Long
    Dim plForwardCoord As Long
    Dim plBackwardCoord As Long

    'Slide down the bottom initial diagonal.
    Do While plXVectorOffset < plXVectorLimit And plYVectorOffset < plYVectorLimit
        If Not mlaXVector(plXVectorOffset) = mlaYVector(plYVectorOffset) Then
            Exit Do
        End If
        
        plXVectorOffset = plXVectorOffset + 1
        plYVectorOffset = plYVectorOffset + 1
    Loop

    'Slide up the top initial diagonal.
    Do While plXVectorLimit > plXVectorOffset And plYVectorLimit > plYVectorOffset
        If Not mlaXVector(plXVectorLimit - 1) = mlaYVector(plYVectorLimit - 1) Then
            Exit Do
        End If
        plXVectorLimit = plXVectorLimit - 1
        plYVectorLimit = plYVectorLimit - 1
    Loop
    
    'Handle simple cases.
    If plXVectorOffset = plXVectorLimit Then
        Do While plYVectorOffset < plYVectorLimit
            moFileVec(1).SetChangedFlag moFileVec(1).RealIndexes(plYVectorOffset) + 1, True
            plYVectorOffset = plYVectorOffset + 1
        Loop
    ElseIf plYVectorOffset = plYVectorLimit Then
        Do While plXVectorOffset < plXVectorLimit
            moFileVec(0).SetChangedFlag moFileVec(0).RealIndexes(plXVectorOffset) + 1, True
            plXVectorOffset = plXVectorOffset + 1
        Loop
    Else
        'Find a point of correspondence in the middle of the files.
        
        plDiag = GetDiagonal(plXVectorOffset, plXVectorLimit, plYVectorOffset, plYVectorLimit)
        plCost = mlCost
        plForwardCoord = mlaForwardDiag(mlaForwardDiagOffset + plDiag)
        plBackwardCoord = mlaBackwardDiag(mlaBackwardDiagOffset + plDiag)
        
        If plCost = 1 Then
            'This should be impossible, because it implies that of of the
            'two subsequences is empty, and that case was handled above
            'without calling diag
            'Lets verify that this is true
            
            Err.Raise vbObjectError, "Diff", "Invalid Subsequence"
        Else
            'Use that point to split this problem into two subproblems
            
            CompareSequence plXVectorOffset, plBackwardCoord, plYVectorOffset, plBackwardCoord - plDiag
            CompareSequence plBackwardCoord, plXVectorLimit, plBackwardCoord - plDiag, plYVectorLimit
        End If
    End If
End Sub

'Discard lines from one file that have no matches in the other file.
Private Sub DiscardConfusingLines()
    moFileVec(0).DiscardConfusingLines moFileVec(1)
    moFileVec(1).DiscardConfusingLines moFileVec(0)
End Sub

'Adjust inserts/deletes of blank lines to join changes as much as possible.
Private Sub ShiftBoundries()
    moFileVec(0).ShiftBoundaries moFileVec(1)
    moFileVec(1).ShiftBoundaries moFileVec(0)
End Sub

'Scan the tables of which lines are inserted and deleted, producing an edit poScript in reverse order.
Private Function BuildReverseScript() As change
    Dim plIndex0 As Long, plIndex1 As Long
    Dim poScript As change
    Dim pbaChanged0() As Boolean
    Dim pbaChanged1() As Boolean
    Dim plLen0 As Long
    Dim plLen1 As Long
    Dim plLine0 As Long
    Dim plLine1 As Long
    Dim poNewScript As change
    
    pbaChanged0 = moFileVec(0).GetChangedFlagArray()
    pbaChanged1 = moFileVec(1).GetChangedFlagArray()
    plLen0 = moFileVec(0).BufferedLines
    plLen1 = moFileVec(1).BufferedLines

    'Note that changedN[plLen0] does exist, and contains 0.
    plIndex0 = 0
    plIndex1 = 0
    
    Do While plIndex0 < plLen0 Or plIndex1 < plLen1
        If pbaChanged0(plIndex0 + 1) Or pbaChanged1(plIndex1 + 1) Then
            plLine0 = plIndex0
            plLine1 = plIndex1
            
            'Find # lines changed here in each file.
            Do While pbaChanged0(plIndex0 + 1)
                plIndex0 = plIndex0 + 1
            Loop
            Do While pbaChanged1(plIndex1 + 1)
                plIndex1 = plIndex1 + 1
            Loop

            'Record this Change.
            Set poNewScript = New change
            poNewScript.Init plLine0, plLine1, plIndex0 - plLine0, plIndex1 - plLine1, poScript
            Set poScript = poNewScript
        End If

        'We have reached lines in the two files that match each other.
        plIndex0 = plIndex0 + 1
        plIndex1 = plIndex1 + 1
    Loop
    Set BuildReverseScript = poScript
End Function

'Scan the tables of which lines are inserted and deleted,
'producing an edit poScript in forward order.
Private Function BuildScript() As change
    Dim poScript As change
    Dim poNewScript As change
    Dim pbaChanged0() As Boolean
    Dim pbaChanged1() As Boolean
    Dim plLen0 As Long
    Dim plLen1 As Long
    Dim plIndex0 As Long
    Dim plIndex1 As Long
    Dim plLine0 As Long
    Dim plLine1 As Long
    
    pbaChanged0 = moFileVec(0).GetChangedFlagArray()
    pbaChanged1 = moFileVec(1).GetChangedFlagArray()
    plLen0 = moFileVec(0).BufferedLines
    plLen1 = moFileVec(1).BufferedLines
    plIndex0 = plLen0
    plIndex1 = plLen1

    'Note that changedN[-1] does exist, and contains 0.
    
    Do While plIndex0 >= 0 Or plIndex1 >= 0
        If pbaChanged0(plIndex0) Or pbaChanged1(plIndex1) Then
            plLine0 = plIndex0
            plLine1 = plIndex1
            
            Do While pbaChanged0(plIndex0)
                plIndex0 = plIndex0 - 1
            Loop
            
            Do While pbaChanged1(plIndex1)
                plIndex1 = plIndex1 - 1
            Loop
            ' Record this Change.
            
            Set poNewScript = New change
            poNewScript.Init plIndex0, plIndex1, plLine0 - plIndex0, plLine1 - plIndex1, poScript
            Set poScript = poNewScript
        End If

        'We have reached lines in the two files that match each other.
        plIndex0 = plIndex0 - 1
        plIndex1 = plIndex1 - 1
    Loop
    Set BuildScript = poScript
End Function

'Report the differences of two files.  DEPTH is the current directory depth.
Public Function Diff(reverse As Boolean) As change
    Dim diags As Long
    
    'Some lines are obviously insertions or deletions
    'because they don't match anything.  Detect them now,
    'and avoid even thinking about them in the main comparison algorithm.
    
    DiscardConfusingLines

    'Now do the main comparison algorithm, considering just the undiscarded lines.

    mlaXVector = moFileVec(0).GetUndiscardedArray
    mlaYVector = moFileVec(1).GetUndiscardedArray
    
    diags = moFileVec(0).NonDiscaredLines + moFileVec(1).NonDiscaredLines + 3
    
    ReDim mlaForwardDiag(diags - 1)
    ReDim mlaBackwardDiag(diags - 1)
    
    mlaForwardDiagOffset = moFileVec(1).NonDiscaredLines + 1
    mlaBackwardDiagOffset = moFileVec(1).NonDiscaredLines + 1

    CompareSequence 0, moFileVec(0).NonDiscaredLines, _
        0, moFileVec(1).NonDiscaredLines
    
    'Modify the results slightly to make them prettier
    'in cases where that can validly be done.
    ShiftBoundries
    
    'Get the results of comparison in the form of a chain
    'of Change's -- an edit Script.
    If reverse Then
        Set Diff = BuildReverseScript
    Else
        Set Diff = BuildScript
    End If
End Function

Public Property Get Heuristic() As Boolean
    Heuristic = mbHeuristic
End Property

Public Property Let Heuristic(pbNewValue As Boolean)
    mbHeuristic = pbNewValue
End Property

Public Property Get NoDiscards() As Boolean
    NoDiscards = mbNoDiscards
End Property

Public Property Let NoDiscards(pbNewValue As Boolean)
    mbNoDiscards = pbNewValue
End Property
