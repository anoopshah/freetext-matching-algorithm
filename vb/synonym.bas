Option Compare Binary
Option Explicit

' Module: synonym -- code for handling synonyms

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

Const maxsynonym = 20000
Dim s_used As Long

' Tables of synonyms
' for get_search_summary - searching on s1
Dim s1_sorted(maxsynonym) As String ' sorted text word/phrase (duplicates are allowed)
Dim s1_result(maxsynonym) As String ' priority and numwords, used for get_search_summary
Dim s1_s2(maxsynonym) As String ' Read word/phrase

' for trylink2 - searching on s2 (Read Term fragment)
Dim s2_sorted(maxsynonym) As String ' sorted Read word/phrase (duplicates are allowed)
Dim s2_s2num(maxsynonym) As Long ' number of words in Read word/phrase
Dim s2_s1num(maxsynonym) As Long ' number of words in text word/phrase
Dim s2_priority(maxsynonym) As Long ' priority of synonym pair
Dim s2_s1(maxsynonym) As String ' text word/phrase

' Priority codes:
' 5 = exact match (both ways) e.g. chronic obstructive pulmonary disease = copd;
' 4 = almost exact match (both ways) e.g. cancer = malignant neoplasm;
' 3 = moderate match (s1 is ) e.g. b pne = bronchopneumonia;
'     e.g. carcinoma (is a type of) = malignant neoplasm;
' 2 = non-standard abbreviation or distorted form; possible one-way match (s2 wider than s1)
'     e.g. rsi = repetitive strain injury;
' 1 = loosely associated (s2 wider than s1) e.g. foot (is a part of) = lower limb;
' -100 = opposite.

Function numrows() As Long
' Returns the number of synonyms (s_used) for use by external functions.
numrows = s_used
End Function

Function import(filename As String) As String
' Imports the synonym table from the text lookup file. Returns a string stating whether the
' table was imported successfully.
Dim fileno As Integer, firstline As String, rawinput As String
Dim b As Long ' for rows of data imported


fileno = freefile
Open filename For Input As fileno
b = 1 ' row number

Line Input #fileno, firstline
If firstline = "text" & vbTab & "read" & vbTab & "priority" & vbTab & "comment" Then
    Do
        Line Input #fileno, rawinput
        If rawinput <> "" Then
            ' Extract the individual variables for the arrays
            ' Number of words in this pattern
            s2_s1(b) = dissect(rawinput, 1, vbTab)
            s2_sorted(b) = dissect(rawinput, 2, vbTab)
            s2_priority(b) = CLng(dissect(rawinput, 3, vbTab))
            b = b + 1
        End If
    Loop Until EOF(fileno)
    s_used = b - 1
    import = "Loaded synonyms table (" & s_used & " rows) from " & filename
Else
    import = "ERROR: " & filename & " is not in the correct format."
    Exit Function
End If
Close #fileno

' Sort the s2 array by s2_sorted
Dim i As Long, iMin As Long, iMax As Long
Dim s2_sorted_swap As String
Dim s2_s1_swap As String
Dim s2_priority_swap As Long

iMin = 1
iMax = s_used
For i = (iMax + iMin) \ 2 To iMin Step -1
    heap_s2 i, iMin, iMax
Next i
For i = iMax To iMin + 1 Step -1
    s2_sorted_swap = s2_sorted(i)
    s2_sorted(i) = s2_sorted(iMin)
    s2_sorted(iMin) = s2_sorted_swap
    
    s2_s1_swap = s2_s1(i)
    s2_s1(i) = s2_s1(iMin)
    s2_s1(iMin) = s2_s1_swap
    
    s2_priority_swap = s2_priority(i)
    s2_priority(i) = s2_priority(iMin)
    s2_priority(iMin) = s2_priority_swap
    heap_s2 iMin, iMin, i - 1
Next i

' Fill in the remaining columns in s2, and create the s1 array
For b = 1 To s_used
    s2_s2num(b) = numwords(s2_sorted(b))
    s2_s1num(b) = numwords(s2_s1(b))
    s1_sorted(b) = s2_s1(b)
    s1_result(b) = s2_priority(b) & " " & s2_s1num(b)
    s1_s2(b) = s2_sorted(b)
Next

' Sort the s1 array
Dim s1_sorted_swap As String
Dim s1_s2_swap As String
Dim s1_result_swap As String

iMin = 1
iMax = s_used
For i = (iMax + iMin) \ 2 To iMin Step -1
    heap_s1 i, iMin, iMax
Next i
For i = iMax To iMin + 1 Step -1
    s1_sorted_swap = s1_sorted(i)
    s1_sorted(i) = s1_sorted(iMin)
    s1_sorted(iMin) = s1_sorted_swap
    
    s1_s2_swap = s1_s2(i)
    s1_s2(i) = s1_s2(iMin)
    s1_s2(iMin) = s1_s2_swap
    
    s1_result_swap = s1_result(i)
    s1_result(i) = s1_result(iMin)
    s1_result(iMin) = s1_result_swap
    heap_s1 iMin, iMin, i - 1
Next i

End Function

Sub heap_s2(ByVal i As Long, iMin As Long, iMax As Long)
' Heap helper function for sorting the synonym table by Read word.
Dim leaf As Long
Dim s2_sorted_swap As String
Dim s2_s1_swap As String
Dim s2_priority_swap As Long
Dim done As Boolean
done = False

Do
    leaf = i + i - (iMin - 1)
    If leaf > iMax Then
        done = True
    Else
        If leaf < iMax Then
            If s2_sorted(leaf + 1) > s2_sorted(leaf) Then
                leaf = leaf + 1
            End If
        End If
        If s2_sorted(i) > s2_sorted(leaf) Then
            done = True
        Else
            s2_sorted_swap = s2_sorted(i)
            s2_sorted(i) = s2_sorted(leaf)
            s2_sorted(leaf) = s2_sorted_swap
            
            s2_s1_swap = s2_s1(i)
            s2_s1(i) = s2_s1(leaf)
            s2_s1(leaf) = s2_s1_swap
            
            s2_priority_swap = s2_priority(i)
            s2_priority(i) = s2_priority(leaf)
            s2_priority(leaf) = s2_priority_swap
            i = leaf
        End If
    End If
Loop Until done
End Sub

Sub heap_s1(ByVal i As Long, iMin As Long, iMax As Long)
' Heap helper function for sorting the synonym table by text word.
Dim leaf As Long
Dim s1_sorted_swap As String
Dim s1_s2_swap As String
Dim s1_result_swap As String
Dim done As Boolean
done = False

Do
    leaf = i + i - (iMin - 1)
    If leaf > iMax Then
        done = True
    Else
        If leaf < iMax Then
            If s1_sorted(leaf + 1) > s1_sorted(leaf) Then
                leaf = leaf + 1
            End If
        End If
        If s1_sorted(i) > s1_sorted(leaf) Then
            done = True
        Else
            s1_sorted_swap = s1_sorted(i)
            s1_sorted(i) = s1_sorted(leaf)
            s1_sorted(leaf) = s1_sorted_swap
            
            s1_s2_swap = s1_s2(i)
            s1_s2(i) = s1_s2(leaf)
            s1_s2(leaf) = s1_s2_swap
            
            s1_result_swap = s1_result(i)
            s1_result(i) = s1_result(leaf)
            s1_result(leaf) = s1_result_swap
            i = leaf
        End If
    End If
Loop Until done
End Sub

Function get_search_summary(instring As String) As String
' Returns the s1_result for an entry in the s1_sorted column (text word/phrase).
' Uses a binary search algorithm.
Dim top As Long, bot As Long, trial As Long
top = 1: bot = s_used
Do
    trial = Int((top + bot) / 2)
    If s1_sorted(trial) > instring Then
        bot = trial - 1
    ElseIf s1_sorted(trial) < instring Then
        top = trial + 1
    ElseIf s1_sorted(trial) = instring Then
        top = trial: bot = trial
    End If
Loop Until bot - top < 2
If s1_sorted(top) = instring Then get_search_summary = "CLIN " & s1_result(top)
If s1_sorted(bot) = instring Then get_search_summary = "CLIN " & s1_result(bot)
End Function

Function trylink_2(ByVal read_term_segment As String, pd_start As Long, pd_fin As Long, _
    cur_true As Boolean) As String
' Tries to match a Read term segment to pd (the text being analysed
' between pd_start and pd_fin). The algorithm starts from the beginning
' of the Read term segment, trying to match the whole of pd between
' pd_start and pd_fin, then tries to get the largest possible match.
' If not possible, it tries smaller segments of the Read term but
' always starting from the beginning. The output is a string with the
' following values (space separated):
' priority position_within_pd_start position_within_pd_fin read_fin.
' If the Read term segment is identical to the text (pd), the output
' has priority 6.
Dim a As Long, b As Long, c As Long, pdstring As String, match_start As Long, match_fin As Long
pdstring = pd.part_nopunc(pd_start, pd_fin)

For a = numwords(read_term_segment) To 1 Step -1
    read_term_segment = strfunc.words(read_term_segment, 1, a)
    ' first try for an exact match
    If InStr(1, pdstring, read_term_segment) Then
        c = 0
        Do ' find exact position of match
            c = c + 1
        Loop Until strfunc.words(pdstring, c, a) = read_term_segment Or c > 1 + pd_fin - pd_start
        If c <= pd_fin - pd_start + 1 Then ' match found
            If (pd.Attr(pd_start) = "negative" And cur_true = False) Or _
                (pd.Attr(pd_start) <> "negative" And cur_true = True) Then
                match_start = c ' start position in pd (relative to pd_start)
                match_fin = c + a - 1 ' finish position in pd, relative to pd_start
                trylink_2 = "6 " & match_start & " " & match_fin & " " & a
                Exit Function
            End If
        End If
    End If
    
    b = s2_pos(read_term_segment) ' search for s2 match
    
    If b > 0 Then
        Do ' loop through all the matches, selecting the best one
            If InStr(1, pdstring, s2_s1(b)) Then ' it might match up; need to double check
                c = 0
                Do ' find exact position of match
                    c = c + 1
                Loop Until strfunc.words(pdstring, c, s2_s1num(b)) = s2_s1(b) Or _
                    c > 1 + pd_fin - pd_start
                If c <= pd_fin - pd_start + 1 Then ' match found
                    If (pd.Attr(pd_start) = "negative" And cur_true = False) Or _
                        (pd.Attr(pd_start) <> "negative" And cur_true = True) Then
                        match_start = c ' start position in pd (relative to pd_start)
                        match_fin = c + s2_s1num(b) - 1 ' finish position in pd,
                            ' relative to pd_start
                        trylink_2 = s2_priority(b) & " " & match_start & " " & _
                            match_fin & " " & s2_s2num(b)
                        Exit Function
                    End If
                End If
            End If
            b = b + 1 ' advancing to next link in synonym table
        Loop Until s2_sorted(b) <> read_term_segment
    End If
Next
' Note that this function only looks for single level links -
' this will be improved in future versions
End Function

Function s2_pos(s2_text As String) As Long
' Returns the topmost position of s2 (partial Read term) text in the s2 sorted list
Dim top As Long, bot As Long, trial As Long
top = 1: bot = s_used
Do
    trial = Int((top + bot) / 2)
    If s2_sorted(trial) > s2_text Then
        bot = trial - 1
    ElseIf s2_sorted(trial) < s2_text Then
        top = trial + 1
    ElseIf s2_sorted(trial) = s2_text Then
        bot = trial
    End If
Loop Until bot - top < 2

If top > bot Then Exit Function
For trial = top To bot
    If s2_sorted(trial) = s2_text And s2_sorted(trial - 1) <> s2_text Then
        s2_pos = trial:  Exit Function  ' top of list of s2 matches
    End If
Next
End Function

Function s1_pos(s1_text As String) As Long
' Returns the topmost position of s1 text in the s1 sorted list.
Dim top As Long, bot As Long, trial As Long
top = 1: bot = s_used
Do
    trial = Int((top + bot) / 2)
    If s1_sorted(trial) > s1_text Then
        bot = trial - 1
    ElseIf s1_sorted(trial) < s1_text Then
        top = trial + 1
    ElseIf s1_sorted(trial) = s1_text Then
        bot = trial
    End If
Loop Until bot - top < 2

If top > bot Then Exit Function
For trial = top To bot
    If s1_sorted(trial) = s1_text And s1_sorted(trial - 1) <> s1_text Then
        s1_pos = trial:  Exit Function  ' top of list of s1 matches
    End If
Next
End Function

Function s2(s1_pos As Long) As String
' Returns the part Read term (s2) at a particular position in the s1 table.
If s1_pos > s_used Then
    s2 = ""
Else
    s2 = s1_s2(s1_pos)
End If
End Function

Function s1(s1_pos As Long) As String
' Returns the part text (s1) at a particular position in the s1 table.
If s1_pos > s_used Then
    s1 = ""
Else
    s1 = s1_sorted(s1_pos)
End If
End Function

Function s1_priority(s1_pos As Long) As Long
' Returns the priority at a particular position in the s1 table.
s1_priority = Val(dissect2(s1_result(s1_pos), , 1))
End Function

