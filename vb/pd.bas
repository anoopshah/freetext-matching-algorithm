Option Compare Binary
Option Explicit

' Module: pd -- arrays for holding individual words of the text being analysed
' (limit of 1000 words), and functions for pattern matching

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

Const maxpartdata = 1000
Dim partdata_used As Long ' number of words in the input text
Dim partdata(maxpartdata) As String ' array containing individual words in the input free text
Dim punc(maxpartdata) As String ' punctuation
Dim attrib(maxpartdata) As String ' attribute e.g. negative, family etc.
Dim meaning(maxpartdata) As String ' interpreted meaning e.g. Read code or date

Sub check_compressed(Optional maybe_pregnant As Boolean, Optional labtest As String)
' Checks that attributes and values are consistent. This function must be
' run after sub compress. It also converts gestational ages into a 'LABS'
' output data type, checks that there is only one gestational age and checks
' that systolic blood pressure is greater than diastolic. It also checks that
' dateprev, datenext etc. refer to clinical events.

Dim a As Long, gest As Long, newtermref As Long
gest = 0: a = 0

Do
    a = a + 1
    Select Case words(Attr(a), 1)
    Case "LABS"
        If attrib(a) = "LABS sysbp" Then ' checking that blood pressure measurements are sensible
            If Attr(a + 1) <> "LABS diabp" Or _
                (Val(words(meaning(a), 2)) <= Val(words(meaning(a + 1), 2))) _
                Or attrib(a - 1) = "normalrange" Then
                pd.remove a: a = a - 1
            End If
        ElseIf attrib(a) = "LABS diabp" Then
            If attrib(a - 1) <> "sysbp" Then pd.remove a: a = a - 1
        End If
        set_attr words(attrib(a), 2), a ' removing the word LABS from the attribute
    Case "negative", "negpmh", "negfamily"
        ' eliminate double negatives
        If Left(meaning(a), 4) = "READ" Then
            If Not terms.true_term(Val(dissect2(meaning(a), , 2))) Then
                set_attr "", a
            End If
        ElseIf meaning(a) = "ATTR" And labtest <> "" Then ' convert 'negative' to lab result
            set_mean "LABS negative", a
            set_attr "", a
        End If
        ' use machinequery if diagnosis is followed immediately by negative diagnosis
        If words(meaning(a), 2, 1) = words(meaning(a - 1), 2, 1) Then
            ' same value, not necessarily same readscore (which is also incorporated in meaning)
            If Not in_set(Attr(a - 1), "negative", "query", "negpmh", "negfamily") Then
                pd.remove a - 1: a = a - 1
                pd.set_attr "machinequery", a
            End If
        ElseIf words(meaning(a), 2, 1) = words(meaning(a + 1), 2, 1) Then
            If Not in_set(Attr(a + 1), "negative", "query", "negpmh", "negfamily") Then
                pd.remove a + 1
                pd.set_attr "machinequery", a
            End If
        End If
    Case "duraprev", "dateprev", "dueto"
        If Left(meaning(a - 1), 4) <> "READ" Then Call set_attr("", a)
    Case "duranext", "datenext", "causing"
        If Left(meaning(a + 1), 4) <> "READ" Then Call set_attr("", a)
    Case "edd"
        If attrib(a - 1) = "edd" Then Call set_attr("", a)
        maybe_pregnant = True
    Case "dob", "lmp"
        If attrib(a - 1) = attrib(a) Then Call set_attr("", a)
    Case "admitdate", "deathdate", "certdate", "dischdate"
        ' to prevent diagnoses being recorded as pmh if they follow a date which is
        ' not necessarily the date that the event occurred
        If Left(meaning(a + 1), 4) = "READ" Then
            If Right(attrib(a + 2), 4) = "prev" Then
            Else
                If attrib(a + 1) = "pmh" Then Call set_attr("", a + 1)
            End If
        End If
        If attrib(a - 1) = attrib(a) Then Call set_attr("", a)
    End Select
    
    If labtest = "normal" Then
        ' non-numeric mode; if text contains Read terms, remove LABS result
        If dissect2(mean(a), " ", 1) = "READ" Then Call remove_from_compressed(, "LABS")
    End If

    If words(meaning(a), 1) = "DURA_gest" Then
        Call set_mean("LABS " & words(meaning(a), 2), a)
        Call set_attr("gest", a)
        If gest = 0 Then
            gest = Val(words(meaning(a), 2))
        ElseIf gest = Val(words(meaning(a), 2)) Then
            ' duplicate gest; delete
            pd.remove a: a = a - 1
        Else
            gest = 999
        End If
    ElseIf words(meaning(a), 1) = "ATTR" And (attrib(a) = "ignore" Or attrib(a) = "") Then
            ' attribute data type - but no attribute!
            pd.remove a: a = a - 1
    End If
    
    ' check for sensible gestational age
    If attrib(a) = "gest" Then
        If Val(words(meaning(a), 2)) < 5 Or Val(words(meaning(a), 2)) > 45 Then
            pd.remove a
            a = a - 1
        End If
    End If
    
    If a > 0 Then
        If words(meaning(a), 1, 2) = words(meaning(a - 1), 1, 2) And _
            (attrib(a) = attrib(a - 1) Or attrib(a) = "") Then
            pd.remove a
            a = a - 1
        End If
    End If

    If Left(meaning(a), 4) = "READ" Then
        newtermref = terms.linkto(Val(dissect2(meaning(a), , 2)))
        meaning(a) = "READ " & newtermref
    End If
Loop Until a >= partdata_used

' Gestational age: convert from duration
a = 0
Do
    a = a + 1
    If Left(meaning(a), 9) = "DURA_wks_" Then
        'find out if this number of weeks is actually gestational age
        If gest = 0 Then ' if gestational age already found in this text, do not bother
            If attrib(a) = "" And maybe_pregnant Then
                If Val(words(meaning(a), 2)) < 43 Then
                    Call set_mean("LABS " & words(meaning(a), 2), a)
                    Call set_attr("gest", a)
                End If
            End If
            gest = Val(words(meaning(a), 2))
        End If
    End If
Loop Until a >= partdata_used

If gest = 999 Then Call remove_from_compressed("gest")
End Sub

Sub remove_from_compressed(Optional ByVal attr_to_remove As String, _
    Optional ByVal type_to_remove As String)
' Removes all entries with a certain attribute from the pd arrays
' if there is a risk it might be wrong.
Dim a As Long
If attr_to_remove = "" Then attr_to_remove = "zzzz"
If type_to_remove = "" Then type_to_remove = "zzzz"
a = 0
Do
    a = a + 1
    If attrib(a) = attr_to_remove Then Call remove(a): a = a - 1
    If dissect2(mean(a), " ", 1) = type_to_remove Then Call remove(a): a = a - 1
Loop Until a = partdata_used
End Sub

Sub compress()
' Converts the pd arrays from a list of words from the original text (i.e. one
' entry per text) to a list of interpreted results (i.e. one entry per output value).
' The original text and punctuation are removed. This is used as an intermediate stage
' in the construction of the final output.
Dim read_row As Long, write_row As Long, family As Boolean, attr_only As Boolean
read_row = 1 = 1: write_row = 0
Dim prev_meaning As String ' use to avoid duplicate entries
prev_meaning = "ZZZZ"
family = False ' detection of family attribute. Family attribute alone
attr_only = False ' if meaning consists solely of attribute e.g. query or negative
If partdata_used = 1 Then ' if text consists solely of 'query' or 'negative'
    If Attr(1) = "negative" Or Attr(1) = "query" Then meaning(1) = "ATTR"
    attr_only = True
End If
Do
    If in_set(Left(pd.mean(read_row), 4), "DATE", "DURA", "LABS", "READ") Then
    'if this row contains useful data then
        If pd.mean(read_row) <> prev_meaning And Not in_set(pd.Attr(read_row), "anon", "possibility") Then
            ' if this is a new meaning add a new row and record meaning and possibly attribute
            write_row = write_row + 1
            Call set_mean(mean(read_row), write_row)
            If correct_attr(read_row) Then
                Call set_attr(Attr(read_row), write_row)
            Else
                Call set_attr("", write_row)
            End If
            prev_meaning = pd.mean(read_row)
        Else
            ' if current attribute is blank and read_row has correct attr, use this one.
            If Attr(write_row) = "" And correct_attr(read_row) And _
                Attr(read_row) <> "negative" Then
                ' negative attributes only apply if they are present at th
                ' beginning of the Read term match, otherwise they imply that the
                ' text matches to a negative part of a Read term
                ' (e.g. 'Retained placenta no haemorrhage')
                Call set_attr(Attr(read_row), write_row)
            End If
        End If
    End If
    If pd.Attr(read_row) = "family" Then family = True
    read_row = read_row + 1
Loop Until read_row > partdata_used

' if text contains only 'father' etc.
If write_row = 0 And (family = True Or attr_only = True) Then
    meaning(1) = "ATTR"
    write_row = 1
End If

Call remove(write_row + 1, partdata_used)

' Remove original text and punctuation (for clarity when pd is displayed)
If write_row > 0 Then
    Do
        partdata(write_row) = ""
        punc(write_row) = ""
        write_row = write_row - 1
    Loop Until write_row = 0
End If
End Sub

Function correct_attr(pos As Long) As Boolean
' Returns True if the attribute is appropriate for the extracted data type
correct_attr = False
Select Case dissect2(mean(pos), , 1) ' first word of meaning = Data Type
Case "READ"
    If in_set(attrib(pos), "", "family", "negative", "negfamily", "query", "dueto", _
        "causing", "pmh", "negpmh", "normalrange") Or Left(pd.Attr(pos), 10) = "deathcause" Then
        correct_attr = True
    End If
Case "DATE_full" ', "DATE_part" - date_part no longer exists
    If in_set(attrib(pos), "", "admitdate", "dischdate", "certdate", "deathdate", _
        "dateprev", "datenext", "lmp", "edd", "dob", "followup", "sicknote") Then
        correct_attr = True
    End If
Case "DATE_year"
    If in_set(attrib(pos), "", "dateprev", "datenext", "dob") Then correct_attr = True
Case "DATE_time"
    If in_set(attrib(pos), "", "certdate", "deathdate", "dateprev", "datenext") Then
        correct_attr = True
    End If
Case "DURA_yrs_"
    If in_set(attrib(pos), "", "duraprev", "duranext", "followup", "age", "ageprev") Then
        correct_attr = True
    End If
Case "DURA_mths"
    If in_set(attrib(pos), "", "duraprev", "duranext", "followup", "age", "ageprev") Then
        correct_attr = True
    End If
Case "DURA_wks_"
    If in_set(attrib(pos), "", "lmp", "duraprev", "duranext", "followup", "age", _
        "sicknote", "ageprev") Then
        correct_attr = True
    End If
Case "DURA_days"
    If in_set(attrib(pos), "", "lmp", "duraprev", "duranext", "followup", "sicknote") Then
        correct_attr = True
    End If
Case "LABS"
    If Left(attrib(pos), 4) = "LABS" Then correct_attr = True
    ' NB at this stage all LABS attributes should be in the form 'LABS xxx'
Case "ATTR"
    If in_set(attrib(pos), "negative", "query", "family") Then correct_attr = True
End Select

If correct_attr = False Then Exit Function
If Right(attrib(pos), 4) = "prev" Then
    ' check that there is a preceding Read term
    If Left(meaning(pos - 1), 4) = "READ" Then Exit Function
    If pos > 1 Then
        Select Case Left(meaning(pos - 2), 4)
        Case "READ": Exit Function
        Case "", "IGNO" ' continue
        Case Else: correct_attr = False: Exit Function
        End Select
    End If
    If pos > 2 Then
        Select Case Left(meaning(pos - 3), 4)
        Case "READ": Exit Function
        Case "", "IGNO"   ' continue
        Case Else: correct_attr = False: Exit Function
        End Select
    End If
    correct_attr = False

ElseIf Right(attrib(pos), 4) = "next" Then
    ' check that there is a following Read term
    If Left(meaning(pos + 1), 4) = "READ" Then Exit Function
    Select Case Left(meaning(pos + 2), 4)
    Case "READ": Exit Function
    Case "", "IGNO"  ' continue
    Case Else: correct_attr = False: Exit Function
    
    End Select
    Select Case Left(meaning(pos + 3), 4)
    Case "READ": Exit Function
    Case "", "IGNO"   ' continue
    Case Else: correct_attr = False: Exit Function
    End Select
    
    correct_attr = False
End If

End Function

Sub show_all_2()
' Adds the whole of the pd arrays to the debug string, for use when analysing
' a single text in debug mode.
Dim a As Long
debug_string = debug_string & "Word : Punctuation : Meaning : Attribute" & Chr$(13) & Chr$(10)
For a = 1 To partdata_used
    debug_string = debug_string & partdata(a) & " : " & punc(a) & " : " & meaning(a) & _
        " : " & attrib(a) & Chr$(13) & Chr$(10)
Next
debug_string = debug_string & Chr$(13) & Chr$(10)
End Sub

Function true_(pos As Long) As Boolean
' Returns True if the attribute at this position is not 'negative'.
If pd.Attr(pos) = "negative" Then true_ = False Else true_ = True
End Function

Function Attr(pos As Long) As String
' Returns the attribute at this position.
If pos < 0 Then Exit Function
Attr = attrib(pos)
End Function

Function mean(pos As Long) As String
' Returns the interpreted meaning at this position.
If pos < 0 Then Exit Function
mean = meaning(pos)
End Function

Sub set_attr(new_attribute As String, pos As Long)
' Sets the attribute at this position to a specific value.
attrib(pos) = new_attribute
End Sub

Sub set_mean(new_meaning As String, pos As Long)
' Sets the interpreted meaning at this position to a specific value.
meaning(pos) = new_meaning
End Sub

Sub add_attr(new_attribute As String, pos_start As Long, Optional pos_fin As Long, _
    Optional ignore_if_already As Boolean)
' Sets the attribute for a range of positions to a specific value.
Dim a As Long
If ignore_if_already Then
    ' if there is already an attribute of any type at this position, exit sub
    If pos_fin > pos_start Then
        For a = pos_start To pos_fin
            If attrib(a) <> "" Then Exit Sub
        Next
    Else
        If attrib(pos_start) <> "" Then Exit Sub
    End If
End If
attrib(pos_start) = new_attribute
If pos_fin > pos_start Then
    For a = pos_start + 1 To pos_fin
        attrib(a) = new_attribute ' code for 'same as previous'
    Next
End If
End Sub

Sub add_mean(new_meaning As String, pos_start As Long, Optional pos_fin As Long, _
    Optional ignore_if_already As Boolean)
' Sets the interpreted meaning for a range of positions to a specific value.
Dim a As Long
If ignore_if_already Then
    ' if there is already an meaning of any type at this position, exit sub
    If pos_fin > pos_start Then
        For a = pos_start To pos_fin
            If meaning(a) <> "" Or a > partdata_used Then Exit Sub
        Next
    Else
        If meaning(pos_start) <> "" Or a > partdata_used Then Exit Sub
    End If
End If
meaning(pos_start) = new_meaning
If pos_fin > pos_start Then
    For a = pos_start + 1 To pos_fin
        If a <= partdata_used Then meaning(a) = new_meaning ' code for 'same as previous'
    Next
End If
End Sub

Function part_nopunc(Optional start As Long, Optional ByVal fin As Long) As String
' Returns a string containing a defined set of words from the text with no punctuation.
If start > pd.max Then Exit Function
If fin < start Then Exit Function
If start = 0 Then start = 1
If fin > pd.max Or fin = 0 Then fin = pd.max
Dim a As Long
For a = start To fin
    If a = start Then
        part_nopunc = partdata(a)
    Else
        part_nopunc = part_nopunc & " " & partdata(a)
    End If
Next
End Function

Function part_punc_nospace(start As Long, fin As Long) As String
' Returns a string containing a defined set of words and punctuation but without
' spaces either side of punctuation.
If start > pd.max Then Exit Function
If fin < start Then Exit Function
If fin > pd.max Then fin = pd.max
Dim a As Long
For a = start To fin
    If a = fin Then ' ignore punctuation for the last one
        part_punc_nospace = part_punc_nospace & partdata(a)
    Else
        If punc(a) = "" Then
            part_punc_nospace = part_punc_nospace & partdata(a) & " "
        Else
            part_punc_nospace = part_punc_nospace & partdata(a) & punc(a)
        End If
    End If
Next
part_punc_nospace = RTrim(part_punc_nospace)
End Function

Function matchpattern(partdata_pos As Long, w1 As String, p1 As String, _
    w2 As String, p2 As String, w3 As String, p3 As String, w4 As String, p4 As String, _
    w5 As String, p5 As String) As Boolean
' Returns True if the set of up to 5 words or meanings (w1-w5) with punctuation (p1-p5)
' match a set of entries in partdata
matchpattern = True
If Not matchposition(partdata_pos, w1, p1) Then matchpattern = False: Exit Function
If w2 = "" Then Exit Function
If Not matchposition(partdata_pos + 1, w2, p2) Then matchpattern = False: Exit Function
If w3 = "" Then Exit Function
If Not matchposition(partdata_pos + 2, w3, p3) Then matchpattern = False: Exit Function
If w4 = "" Then Exit Function
If Not matchposition(partdata_pos + 3, w4, p4) Then matchpattern = False: Exit Function
If w5 = "" Then Exit Function
If Not matchposition(partdata_pos + 4, w5, p5) Then matchpattern = False
End Function

Function matchposition(partdata_pos As Long, ByVal word As String, ByVal punct As String) As Boolean
' Returns True if there is a match between the search word and
' text. The argument 'word' can represent either text or meaning (if enclosed in []).
Dim a As Long, b As Long
Dim temp As Boolean, tempstr As String
matchposition = False
a = CInt(dissect2(word, "|", 1))
For b = 2 To a + 1
    temp = matchoption(partdata_pos, dissect2(word, "|", b), punct)
    If temp = True Then Exit For
Next
If temp = True Then matchposition = True
End Function

Function matchoption(partdata_pos As Long, ByVal word As String, ByVal punct As String) As Boolean
' Match the free text and single position match meaning / words
Dim upp As Single, Low As Single, actual As Single
matchoption = False
If Left(word, 1) = "[" Then
    If Left(word, 6) = "[ATTR " Then
        ' match attribute
        If attrib(partdata_pos) = Mid(Left(word, Len(word) - 1), 7) Then matchoption = True
    ElseIf Left(word, 6) = "[NUMB_" Then
        ' indicates a range of numbers (e.g. viable blood pressures or weeks gestation
        If Left(meaning(partdata_pos), 4) = "NUMB" Then
            Low = Val(dissect(word, 2, "_")): upp = Val(dissect(word, 3, "_"))
            actual = Val(words(meaning(partdata_pos), 2))
            If actual >= Low And actual <= upp Then matchoption = True
        End If
    ElseIf Left(meaning(partdata_pos), InStr(1, word, "]") - 2) = Mid(word, 2, InStr(1, word, "]") - 2) Then
        matchoption = True ' e.g. [DATE_full] or [DATE]
    End If
Else
    If word = "*" Or word = text(partdata_pos) Then
        matchoption = True
    End If
End If
' now match punctuation (the pattern string must CONTAIN the punctuation string;
' i.e. pattern string can give several options. blank = no punctuation, * = any)
If punct = "*" Then Exit Function
If InStr(1, punct, "_") And punc(partdata_pos) = "" Then Exit Function
    ' if _ is included in punctuation pattern, it means blank is acceptable
If punct = "" Or punct = "_" Then ' punctuation  MUST be blank
    If punc(partdata_pos) <> "" Then matchoption = False
    Exit Function
Else ' must NOT be blank
    If punc(partdata_pos) = "" Then matchoption = False
End If
If InStr(1, punct, punc(partdata_pos)) = 0 Then matchoption = False
End Function

Sub init_read(instring As String)
' Initialises the 'partdata' and 'punc' arrarys in the pd module with
' words and punctuation from the free text, by parsing the raw free text string.
' Also converts symbols '+' and  '&' to the word 'and', and '###' (used for CPRD
' anonymised words) to the word 'anonymised', to avoid it being recognised as
' part of a Read term.

Call clear ' clears pd
instring = LCase(instring) & " "

' st_type Code: 0=space, 1=word, 2=number, 3=punctuation
Dim a1 As Long ' text position - start of word/phrase
Dim a2  As Long ' text position - end of word/phrase
Dim b As Long ' partdata position
Dim cur As Long, nxt As Long, pattern As String, a As Integer
a1 = 1: a2 = 1
b = 1
Do
    cur = st_type(Mid(instring, a1, 1))
    'advance a2 to end of block
    Do While st_type(Mid(instring, a2, 1)) = cur And a2 < Len(instring)
        a2 = a2 + 1
    Loop
    nxt = st_type(Mid(instring, a2, 1))
    Select Case cur
    Case 0 ' current section is a space. do nothing
    Case 1 ' current section is a word but ignore 's and (s) and (es)
        If (Mid(instring, a1, (a2 - a1)) = "s" Or Mid(instring, a1, (a2 - a1)) = "es") And _
            ((punc(b - 1) = "(" And Mid(instring, a2, 1) = ")") Or punc(b - 1) = "'") Then
            ' ignore
        Else
            partdata(b) = Mid(instring, a1, (a2 - a1))
            b = b + 1
        End If
    Case 2 ' current section is a number
    
        ' If the next section is a dot (e.g. 12.6)
        If Mid(instring, a2, 1) = "." And is_numeric(Mid(instring, a2 + 1, 1)) Then
            a = 0: pattern = ""
            Do ' generate pattern string
                If IsNumeric(Mid(instring, a1 + a, 1)) Then
                    pattern = pattern & "#"
                Else
                    pattern = pattern & Mid(instring, a1 + a, 1)
                End If
                a = a + 1
            Loop Until Not IsNumeric(Mid(instring, a1 + a, 1)) And Mid(instring, a1 + a, 1) <> "."
            Select Case pattern
            Case "#.#.##", "#.#.####"
                instring = Mid(instring, 1, a1) & "/" & Mid(instring, a1 + 2, 1) & _
                    "/" & Mid(instring, a1 + 4)
            Case "#.##.##", "#.##.####"
                instring = Mid(instring, 1, a1) & "/" & Mid(instring, a1 + 2, 2) & _
                    "/" & Mid(instring, a1 + 5)
            Case "##.#.##", "##.#.####"
                instring = Mid(instring, 1, a1 + 1) & "/" & Mid(instring, a1 + 3, 1) & _
                    "/" & Mid(instring, a1 + 5)
            Case "##.##.##", "##.##.####"
                instring = Mid(instring, 1, a1 + 1) & "/" & Mid(instring, a1 + 3, 2) & _
                    "/" & Mid(instring, a1 + 6)
            Case Else
                ' combining numbers with decimal points
                Do
                    a2 = a2 + 1
                Loop Until Not is_numeric(Mid(instring, a2, 1))
            End Select
            partdata(b) = Mid(instring, a1, (a2 - a1))
        ElseIf Mid(instring, a2, 2) = ". " And a2 - a1 < 3 And _
            is_numeric(Mid(instring, a2 + 2, 1)) And Mid(instring, a2 + 3, 1) = "%" Then
            ' If the percentage has a space, e.g. 12. 3% or 5. 4%
            partdata(b) = Mid(instring, a1, a2 - a1) & "." & Mid(instring, a2 + 2, 1)
            punc(b) = "%"
            a2 = a2 + 4
        Else
            partdata(b) = Mid(instring, a1, (a2 - a1))
        End If
        b = b + 1
    Case 3 ' current section is punctuation
        Select Case Mid(instring, a1, (a2 - a1))
        Case "+", "&"
            partdata(b) = "and"
            b = b + 1
        Case "###" ' anonymised term
            partdata(b) = "anonymised"
            b = b + 1
        Case "?", "??"
            partdata(b) = "query"
            b = b + 1
        Case "/"
            ' / is generally treated as a word rather than punctuation
            If nxt = 2 Or Len(partdata(b - 1)) = 1 Then
                ' if the next section is a number, include / as punc rather than a word
                punc(b - 1) = punc(b - 1) & Mid(instring, a1, (a2 - a1))
            Else
                partdata(b) = "/"
                b = b + 1
            End If
        Case Else
            punc(b - 1) = punc(b - 1) & Mid(instring, a1, (a2 - a1))
        End Select
    End Select
    a1 = a2
Loop Until a1 >= Len(instring) Or b > maxpartdata - 2 ' for safety - in case line too long!
partdata_used = b - 1

End Sub

Function st_type(instring As String) As Long
' Returns the type of a text string; 0 if it is a single space, 1 if it is
' part of a word, 2 if it is a number, and 3 if it does not fit into any of the
' other categories (i.e. if it is punctuation).
If instring = " " Then st_type = 0: Exit Function
If is_text(instring) Then st_type = 1: Exit Function
If is_numeric(instring) Then st_type = 2: Exit Function
st_type = 3 ' punctuation
End Function

Sub clear()
' Clears the 'partdata', 'punc', 'attrib' and 'meaning' arrays in the pd module.
Dim counter As Long
For counter = 0 To partdata_used
    partdata(counter) = ""
    punc(counter) = ""
    attrib(counter) = ""
    meaning(counter) = ""
Next
partdata_used = 0
End Sub

Sub remove(pos1 As Long, Optional pos2 As Long)
' Removes data from the arrays in the pd module between the specified positions
If pos1 > partdata_used Then Exit Sub
If pos2 > partdata_used Then pos2 = partdata_used
Dim diff As Long
If pos2 > 0 Then diff = pos2 - pos1 + 1 Else diff = 1
Dim counter As Long
For counter = pos1 To partdata_used
    If counter + diff > maxpartdata Then
        partdata(counter) = ""
        punc(counter) = ""
        attrib(counter) = ""
        meaning(counter) = ""
    Else
        punc(counter) = punc(counter + diff)
        partdata(counter) = partdata(counter + diff)
        attrib(counter) = attrib(counter + diff)
        meaning(counter) = meaning(counter + diff)
    End If
Next
punc(partdata_used + 1) = ""
partdata(partdata_used + 1) = ""
attrib(partdata_used + 1) = ""
meaning(partdata_used + 1) = ""
partdata_used = partdata_used - diff
End Sub

Function text(position As Long) As String
' Returns the word at a particular position (from the 'partdata' array).
If position > maxpartdata Then Exit Function
If position < 1 Then Exit Function
text = partdata(position)
End Function

Sub set_text(new_text As String, position As Long)
' Replaces the word at a particular position (in the 'partdata' array).
If position > maxpartdata Then Exit Sub
If position < 1 Then Exit Sub
partdata(position) = new_text
End Sub

Function punct(position As Long) As String
' Returns the punctuation at a particular position (from the 'punc' array).
If position > maxpartdata Then Exit Function
If position < 1 Then Exit Function
punct = punc(position)
End Function

Function max() As Long
' Returns the total number of words in the input text
max = partdata_used
End Function

