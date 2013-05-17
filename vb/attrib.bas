Option Compare Binary
Option Explicit

' Module: attrib -- code related to the attributes table. The table is loaded
' from a text file by the 'import' function

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

Const maxattrib = 400

Dim w(5, maxattrib) As String ' pattern of up to 5 words
Dim p(5, maxattrib) As String ' options for punctuation associated with each word
Dim a(5, maxattrib) As String ' attribute associated with each word
Dim death_only(maxattrib) As Boolean ' whether this attribute pattern is only applicable in 'death' mode
Dim numwd(maxattrib) As Long ' number of words (1 to 5) in this pattern
Dim order(maxattrib) As Double ' order of this row in the lookup table; not used in the actual algorithm but loaded for debug purposes.
Dim num As Integer


Function import(filename As String) As String
' Imports attributes lookup table and returns a string stating what was
' imported. The table must be already be sorted in order; this is checked
' but not corrected.
Dim fileno As Integer, firstline As String, rawinput As String
Dim b As Long ' for rows of data imported
Dim previousorder As Double ' to ensure that order is strictly increasing

fileno = freefile
Open filename For Input As fileno
b = 1 ' row number
previousorder = 0

Line Input #fileno, firstline
If firstline = "order" & vbTab & "w1" & vbTab & "p1" & vbTab & "w2" & vbTab & "p2" & _
    vbTab & "w3" & vbTab & "p3" & vbTab & "w4" & vbTab & "p4" & _
    vbTab & "w5" & vbTab & "p5" & _
    vbTab & "a1" & vbTab & "a2" & vbTab & "a3" & vbTab & "a4" & vbTab & "a5" & _
    vbTab & "death_only" & vbTab & "comment" Then
    Do
        Line Input #fileno, rawinput
        If rawinput <> "" Then
            ' Extract the individual variables for the arrays
            ' Number of words in this pattern
            numwd(b) = 0
            order(b) = CDbl(dissect(rawinput, 1, vbTab))
            If order(b) < previousorder Then
                import = "ERROR: Rows " & order(b) & " and " & _
                    previousorder & " are out of order. The order must be strictly ascending."
                Close #fileno
                Exit Function
            Else
                previousorder = order(b)
            End If
            w(1, b) = dissect2_options(dissect(rawinput, 2, vbTab))
            w(2, b) = dissect2_options(dissect(rawinput, 4, vbTab))
            w(3, b) = dissect2_options(dissect(rawinput, 6, vbTab))
            w(4, b) = dissect2_options(dissect(rawinput, 8, vbTab))
            w(5, b) = dissect2_options(dissect(rawinput, 10, vbTab))
            If w(1, b) <> "" Then numwd(b) = 1
            If w(2, b) <> "" Then numwd(b) = 2
            If w(3, b) <> "" Then numwd(b) = 3
            If w(4, b) <> "" Then numwd(b) = 4
            If w(5, b) <> "" Then numwd(b) = 5
            
            ' Punctuation
            p(1, b) = dissect(rawinput, 3, vbTab)
            p(2, b) = dissect(rawinput, 5, vbTab)
            p(3, b) = dissect(rawinput, 7, vbTab)
            p(4, b) = dissect(rawinput, 9, vbTab)
            p(5, b) = dissect(rawinput, 11, vbTab)
            
            ' Attributes
            a(1, b) = dissect(rawinput, 12, vbTab)
            a(2, b) = dissect(rawinput, 13, vbTab)
            a(3, b) = dissect(rawinput, 14, vbTab)
            a(4, b) = dissect(rawinput, 15, vbTab)
            a(5, b) = dissect(rawinput, 16, vbTab)
            
            ' Whether death only
            If dissect(rawinput, 17, vbTab) = "TRUE" Then
                death_only(b) = True
            Else
                death_only(b) = False
            End If
            b = b + 1
        End If
    Loop Until EOF(fileno)
    num = b - 1
    import = "Loaded the attribute table (" & num & " rows) from " & filename
Else
    import = "ERROR: Unable to load attribute table from " & filename
End If
End Function

Function dissect2_options(instring As String) As String
' Counts the number of options in a string and puts it at the front of
' the string, for future use by the dissect2 function. e.g. 'word|another|option' is
' converted to '3|word|another|option'.
If instring = "" Then Exit Function
Dim pos As Integer, n As Integer
pos = 0: n = 0
Do
    pos = InStr(pos + 1, instring, "|"): n = n + 1
Loop Until pos = 0
dissect2_options = n & "|" & instring
End Function

Sub pd_search2(Optional debug_ As Boolean, Optional death As Boolean)
' Tries each attribute pattern in turn to see whether it applies to the free text
' being analysed. Results are added to attribute fields in the arrays in module pd.
Dim b As Long, start As Long, c As Long, temp As String
For b = 1 To num
    If death = True Or death_only(b) = False Then ' whether or not to use this attrib pattern
        For start = 1 To pd.max
            If pd.matchpattern(start, w(1, b), p(1, b), w(2, b), p(2, b), w(3, b), _
                p(3, b), w(4, b), p(4, b), w(5, b), p(5, b)) Then
            
                For c = 1 To numwd(b)
                    If a(c, b) <> "_" Then pd.set_attr a(c, b), start + c - 1
                    If a(c, b) = "." Then pd.set_attr "", start + c - 1
                    If InStr(1, a(c, b), " ") Then ' more than one word i.e. meaning included
                        pd.set_mean (dissect2(a(c, b), , 1) & " " & pd.text(start + c - 1)), start + c - 1
                    End If
                Next

                If debug_ Then
                    debug_string = debug_string & "Attrib phrase (search position " & order(b) & "): " & _
                        w(1, b) & p(1, b) & " " & w(2, b) & p(2, b) & " " & w(3, b) & p(3, b) & " " & _
                        w(4, b) & p(4, b) & " " & w(5, b) & p(5, b) & Chr$(13) & Chr$(10)
                    debug_string = debug_string & "Matches to: " & pd.part_punc_nospace(start, start + numwd(b) - 1) _
                        & Chr$(13) & Chr$(10)
                End If
            End If
        Next
    End If
Next
End Sub

