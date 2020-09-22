VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "change"
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

'The result of comparison is an "edit script": a chain of Change objects.
'Each Change represents one place where some lines are deleted
'and some are inserted.
'
'plLine0 and plLine1 are the first affected lines in the two files (origin 0).
'DELETED is the number of lines deleted here from file 0.
'INSERTED is the number of lines inserted here in file 1.
'
'If DELETED is 0 then plLine0 is the number of the line before
'which the insertion was done; vice versa for INSERTED and plLine1.
    
'Previous or next edit command.
Public Link As change

'# lines of file 1 changed here.
Public Inserted As Long

'# lines of file 0 changed here.
Public Deleted As Long
    
'Line number of 1st deleted line.
Public plLine0 As Long
    
'Line number of 1st inserted line.
Public plLine1 As Long
    
'plLine0 and plLine1 are the first affected lines in the two files (origin 0).
'DELETED is the number of lines deleted here from file 0.
'INSERTED is the number of lines inserted here in file 1.
'
'If DELETED is 0 then plLine0 is the number of the line before
'which the insertion was done; vice versa for INSERTED and plLine1.
Public Sub Init(plLine0 As Long, plLine1 As Long, Deleted As Long, Inserted As Long, poNext As change)
    Me.plLine0 = plLine0
    Me.plLine1 = plLine1
    Me.Inserted = Inserted
    Me.Deleted = Deleted
    Set Me.Link = poNext
End Sub

