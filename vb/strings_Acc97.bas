Option Compare Binary
Option Explicit

' Module: strings_Acc97 -- functions for manipulating strings that are
' provided in VBA for Access 2003 but not in Access 97

'    Freetext Matching Algorithm: natural language analysis system for clinical text
'    Copyright: Anoop Dinesh Shah, 2012, 2013
'    Email: anoop@doctors.org.uk

'    This file is part of the Freetext Matching Algorithm.

'    The Freetext Matching Algorithm is free software: you can
'    redistribute it and/or modify it under the terms of the
'    GNU General Public License as published by the Free Software
'    Foundation, either version 3 of the License, or (at your option)
'    any later version.

'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>

Function replace(bigstring As String, lookstring As String, replacestring As String) As String
' Returns bigstring with every instance of lookstring replaced with replacestring
Dim a As Integer, b As Integer, c As Integer
replace = bigstring
a = InStr(1, bigstring, lookstring)
If lookstring = "" Or a = 0 Then Exit Function
b = Len(lookstring)
c = Len(replacestring)
Do
    replace = Left(replace, a - 1) & replacestring & Mid(replace, a + b)
    a = InStr(a + c, replace, lookstring)
Loop Until a = 0
End Function

Function dissect3(in_string As String, Optional delimiter As String, _
    Optional number As Long) As String
' Equivalent to the VBA.Split() function in Access 2003, so this program can run in Access 97.
Dim part_string As String
Dim start As Long
Dim finish As Long
Dim looper As Long
Dim current_pos As Long
If delimiter = "" Then delimiter = " " ' default is space. Default number is 1.

' if the first section is requested, the solution is simpler
If number = 0 Or number = 1 Then
    finish = InStr(1, in_string, delimiter, vbTextCompare)
    If finish = 0 Then
        ' delimiter is not present in in_string
        dissect3 = in_string
        Exit Function
    Else
        dissect3 = Left(in_string, finish - 1)
        Exit Function
    End If
End If

looper = number
current_pos = 0
' find position of delimiter just before required data
Do While looper > 1
    current_pos = InStr(current_pos + 1, in_string, delimiter)
    looper = looper - 1
    If current_pos = 0 Then
        ' current_pos is after the end of the string
        Exit Do
    End If
Loop

If current_pos = 0 Then dissect3 = "": Exit Function

' current_pos is now the position of the delimiter just before the start of the data

start = current_pos + 1
' start is the position of the start of the data

finish = InStr(current_pos + 1, in_string, delimiter)
' finish is the position of the delimiter just after the data

If start = finish Then
    dissect3 = "" ' null data
    Exit Function
ElseIf finish = 0 Then
    part_string = Mid(in_string, start)
Else
    part_string = Mid(in_string, start, finish - start)
End If

dissect3 = part_string
End Function

Function monthname(number As Integer, short As Boolean) As String
' Name of the month (either full name or short name).
If short = False Then
    Select Case number
    Case 1: monthname = "january"
    Case 2: monthname = "february"
    Case 3: monthname = "march"
    Case 4: monthname = "april"
    Case 5: monthname = "may"
    Case 6: monthname = "june"
    Case 7: monthname = "july"
    Case 8: monthname = "august"
    Case 9: monthname = "september"
    Case 10: monthname = "october"
    Case 11: monthname = "november"
    Case 12: monthname = "december"
    End Select
Else
    Select Case number
    Case 1: monthname = "jan"
    Case 2: monthname = "feb"
    Case 3: monthname = "mar"
    Case 4: monthname = "apr"
    Case 5: monthname = "may"
    Case 6: monthname = "jun"
    Case 7: monthname = "jul"
    Case 8: monthname = "aug"
    Case 9: monthname = "sep"
    Case 10: monthname = "oct"
    Case 11: monthname = "nov"
    Case 12: monthname = "dec"
    End Select
End If
End Function
