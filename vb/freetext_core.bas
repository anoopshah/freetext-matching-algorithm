Option Compare Binary
Option Explicit

' Module: freetext_core -- core algorithm

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

Const wordmatchthreshold = 0.73 ' used by readscore

Public debug_string As String ' stores analysis report for an individual text, when running in debug mode
Dim death As Boolean ' whether Read term implies death
Dim gest As Boolean ' whether Read term refers to weeks gestation
Dim spell As Boolean ' whether to use spelling correction

Sub main_termref(instring As String, Termref As Long, Optional spell_ As Boolean, _
    Optional debug_ As Boolean, Optional ByVal append_term As Boolean)
' Calls main_analyse with appropiate analysis option based on the Read term associated
' with the text, and depending on the append_term argument it may also append
' the text to the end of the Read term to appear as it would on the GP's computer.
Dim datatype As String, termstring As String
datatype = terms.read_type(Termref)
termstring = terms.std_term(Termref)
Dim tempterm As Long, tempterm2 As Long

' Find out what FMA extracts from original Read term, in order to remove it from
' the final output
If append_term And Not in_set(datatype, "L", "N", "T", "S") Then
    Call main_analyse(instring:=termstring, spell_:=False, debug_:=False, _
        append_term:=False)
    If Left(pd.mean(1), 4) = "READ" Then tempterm = Val(dissect2(pd.mean(1), " ", 2))
End If

Select Case datatype
Case "D" ' death
    Call main_analyse(instring:=instring, death_:=True, spell_:=spell_, debug_:=debug_, _
        termstring:=termstring, append_term:=append_term)
Case "P" ' pregnant
    Call main_analyse(instring:=instring, pregnant_:=True, spell_:=spell_, debug_:=debug_, _
        termstring:=termstring, append_term:=append_term)
Case "L" ' lab value
    Call main_analyse(instring:=instring, labtest:="value", spell_:=spell_, debug_:=debug_, _
        termstring:=termstring, append_term:=False)
Case "N" ' normal / abnormal
    Call main_analyse(instring:=instring, labtest:="normal", spell_:=spell_, debug_:=debug_, _
        termstring:=termstring, append_term:=False)
Case "T" ' date only - Read term specifies that free text contains date
    Call main_analyse(instring:=instring, date_only:=True, spell_:=spell_, debug_:=debug_, _
        termstring:=termstring, append_term:=False)
Case "S" ' sicknote - Dates and durations are classified as 'sicknote'
    Call main_analyse(instring:=instring, sicknote:=True, spell_:=spell_, debug_:=debug_, _
        termstring:=termstring, append_term:=False)
Case Else
    Call main_analyse(instring:=instring, spell_:=spell_, debug_:=debug_, _
        termstring:=termstring, append_term:=append_term)
End Select

If append_term And Not in_set(datatype, "L", "N", "T", "S") Then
    tempterm2 = Val(dissect2(pd.mean(1), " ", 2))
    If (tempterm2 = tempterm Or tempterm2 = Termref) And pd.Attr(1) = "" Then
        pd.remove 1
    End If
    ' remove first interpreted value if it is the same as the existing value / term
    ' and there is no attribute.
End If

End Sub

Sub main_analyse(ByVal instring As String, Optional death_ As Boolean, _
    Optional pregnant_ As Boolean, Optional debug_ As Boolean, _
    Optional labtest As String, Optional spell_ As Boolean, Optional date_only As Boolean, _
    Optional termstring As String, Optional append_term As Boolean, _
    Optional sicknote As Boolean)
' This is the main part of the Freetext Matching Algorithm which calls functions
' to perform each of the major steps in the analysis of an input text (instring).
death = death_: gest = pregnant_: If death = True Then gest = False
spell = spell_

' Initialise readscore (because it stores previous results in order to save time)
Dim dummy As Single
dummy = readscore(0, 0, 0, False, True)

' If running in test mode, start to create an analysis report, storing it
' in debug_string
If debug_ Then
    debug_string = "Analysis options: " & _
        "Death = " & death_ & ", Pregnancy = " & pregnant_ & Chr$(13) & Chr$(10) & _
        "Lab test = " & labtest & ", Date only = " & date_only & _
        ", Sicknote = " & sicknote & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & _
        "INITIAL_SEARCH, ATTRIB.PD_SEARCH2" & _
        Chr$(13) & Chr$(10)
Else
    debug_string = ""
End If

' Remove semi-structured ignorable phrases (in the 'ignorephrase' table)
If append_term Then
    instring = termstring & " " & remove_ignore_phrases(instring) & " "
Else
    instring = " " & remove_ignore_phrases(instring) & " "
End If

' Initialise the arrays in the pd module with the free text
pd.init_read (instring)

' Identify text fragments as dates, numbers, non-clinical words or words
' which might be part of a Read term. Try to correct spelling mistakes.
Call initial_search(debug_)

' Search for attributes
attrib.pd_search2 debug_, death

If debug_ Then
    pd.show_all_2
    debug_string = debug_string & "ATTRIB_SEARCH, ANALYSE_PD" & Chr$(13) & Chr$(10)
End If

' Extend attributes to nearby words according to hard-coded rules.
Call attrib_search(debug_)

' Match sequences of words to Read terms
Call analyse_pd(debug_, labtest)

If debug_ Then
    pd.show_all_2
    debug_string = debug_string & _
        "PD.COMPRESS, PD.CHECK_COMPRESSED" & Chr$(13) & Chr$(10)
End If

' Discard the original text and keep the sequence of extracted structured information
pd.compress

' Check whether certain extracted medcodes are invalidated or validated
' by the text
checkterms.check_all instring, debug_, sicknote, death_, date_only

' Check that attributes are valid for the data type
pd.check_compressed gest, labtest

If debug_ Then pd.show_all_2

' Now the Freetext Matching Algorithm output is in the pd arrays
' (meaning and attribute). This can be further processed by the do_analysis]
' function in the fma_gold module to produce output in a format similar to
' fma_gold
End Sub

Function import_all_lookups(lookupfolder As String) As String
' Imports all lookup tables from text files by calling the appropriate
' import functions in the modules attrib, checkterms, synonym, terms and wordlist.
' The text files must have standard names as in the master repository
' (https://github.com/anoopshah/freetext-matching-algorithm-lookups).
' Returns a string stating what was imported.
Dim out As String, newline As String
newline = Chr$(13) & Chr$(10)

' Load all the lookup tables
out = attrib.import(lookupfolder & "attributes.txt")
out = out & newline & checkterms.import(lookupfolder & "checkterms.txt")
out = out & newline & synonym.import(lookupfolder & "synonyms.txt")
out = out & newline & wordlist.import_ignore(lookupfolder & "ignore.txt", _
    lookupfolder & "ignorephrase.txt")

' Terms must be loaded in the order native, virtual, alternate
out = out & newline & terms.import(lookupfolder & "nativeterms.txt", "native")
out = out & newline & terms.import(lookupfolder & "virtualterms.txt", "virtual")
out = out & newline & terms.import(lookupfolder & "alternateterms.txt", "alternate")
    
' Wordlist must be loaded after synonyms and terms
out = out & newline & wordlist.import_wordlist(lookupfolder & "2of4brif.txt")

import_all_lookups = out
End Function

Sub initial_search(Optional debug_ As Boolean)
' Identifies synonyms, words which might be part of a Read term, numbers
' and dates in the free text, recording the results in the 'meaning' array
' in the pd module.
If pd.max = 0 Then Exit Sub ' if there is no text in partdata, exit sub
Dim a As Long ' start position of phrase
Dim b As Long ' number of words in phrase
Dim found As Boolean ' whether a match is found
Dim temp_long As Long
Dim temp_date As String
Dim temp_string As String
Dim temp_pd As String
Dim maxlen As Long
a = 1

Do
    If (pd.max - a + 1) >= 5 Then maxlen = 5 Else maxlen = (pd.max - a + 1)
    found = False
    If maxlen > 1 Then
        For b = 5 To 2 Step -1 ' analyses up to 5 words at a time
            temp_pd = pd.part_nopunc(a, a + b - 1)
    
            ' date search (only get times in DEATH mode)
            temp_date = strfunc.get_date(pd.part_punc_nospace(a, a + b - 1), death)
            If temp_date <> "" Then
                pd.add_mean temp_date, a, a + b - 1
                found = True
                Exit For
            End If
            
            ' synonym search
            temp_string = synonym.get_search_summary(temp_pd)
            If temp_string <> "" Then
                pd.add_mean temp_string, a, a + b - 1
                found = True
                Exit For
            End If
            
            If found Then Exit For
        Next ' loop b
    End If ' if maxlen>1
    
    ' if still not found, try searching on the single word
    If Not found Then
        temp_pd = pd.text(a)
        b = 1
        ' look up ignorable list
        If ignorable(temp_pd) Then
            pd.add_attr "ignore", a
            pd.add_mean "IGNO", a
            found = True
        End If
    End If
        
    If Not found Then
        ' Word is a number? (lab results mode TRUE, i.e. words such as 'normal' are numbers)
        If strfunc.is_numeric(temp_pd, True) Then
            temp_date = strfunc.get_date(temp_pd, death)
            If temp_date <> "" Then
                pd.add_mean temp_date, a, a, False
            Else
                pd.add_mean "NUMB " & temp_pd, a, a, False
            End If
            found = True
        End If
    End If
    
    If Not found Then
        ' synonym match?
        temp_string = synonym.get_search_summary(temp_pd)
        If temp_string <> "" Then
            pd.add_mean temp_string, a, a, False
            found = True
        End If
    End If
       
    If Not found Then
        ' Word is part of Read term?
        temp_string = wordsearch(temp_pd, spell)
        If temp_string <> "" Then
            pd.add_mean temp_string, a, a, False
            If Left(temp_string, 4) = "CLIN" Then
                pd.set_text new_text:=Mid(temp_string, 6), position:=a
            End If
        End If
    End If

    a = a + b
Loop Until a > pd.max
End Sub

Sub attrib_search(Optional debug_ As Boolean)
' Extends context attributes found on pattern matching (attrib.pd_search2)
' to nearby words based on hard-coded patterns.
 Dim a As Long, b As Long

If pd.max = 1 Then Exit Sub
Dim cur_attr As String
Dim length As Long
length = 1
cur_attr = pd.Attr(1)
For a = 2 To pd.max
    If pd.Attr(a) = "" Or pd.Attr(a) = "ignore" Then

        ' stick with previous attribute until you come to punctuation
        If (length > 2 And Not in_set(Left(pd.mean(a), 4), "DATE", "DURA", "CLIN", "LABS") _
            Or (Len(pd.text(a - 1)) > 1 And in_set(pd.punct(a - 1), ".", ").", ".)", ".("))) Then
            ' if more than two ignorable words in between or full stop then stop attribute.
            Select Case cur_attr
            Case "family", "normalrange", "possibility"
                ' stop family or normalrange attribute only after a full stop or new attribute.
                If in_set(pd.punct(a - 1), ".", ").", ".)", ".(") Then
                    cur_attr = ""
                Else
                    pd.set_attr cur_attr, a
                End If
            Case Else
                cur_attr = "" ' stop other attributes
            End Select
        Else
            Select Case cur_attr
            Case "ignore"
            Case "normalrange", "possibility"
                pd.set_attr cur_attr, a ' keep going until a new attribute appears
            Case "anon"
                cur_attr = ""
            Case "negative"
                If (in_set(pd.punct(a - 1), ".", ").", ":", ";", "-") Or _
                    in_set(pd.text(a - 1), "but", "has", "some", "slight", "seems", "looks", _
                        "feels", "usually", "sometimes", "than", "mild", "severe", _
                        "except", "unless", "until") Or _
                    in_set("normal", pd.text(a - 1), pd.text(a), pd.text(a + 1), _
                        pd.text(a + 2))) Then
                    cur_attr = "" ' stop negative attribute after full stop or 'but' or 'except'
                    ' or if the following phrase contains 'normal'
                Else
                    pd.set_attr "negative", a ' continue current attribute
                End If
            Case Else
                pd.set_attr cur_attr, a ' continue current attribute
            End Select
        End If
        If Not in_set(Left(pd.mean(a), 4), "DATE", "DURA", "CLIN", "LABS") Then
            length = length + 1
        End If
    Else
        ' record current attribute and leave it as it is
        ' (however do not carry forward 'dueto' or 'causing')
        If in_set(pd.Attr(a), "dueto", "causing", "duranext") Then
            cur_attr = ""
        Else
            cur_attr = pd.Attr(a)
        End If
        length = 1
    End If
Next

cur_attr = "" ' copying family attribute backwards, up to 5 words (counted by 'length')

' search ahead for formation 'no xxx, yyy, yyy, no xxx' or even 'no xxx yyy yyy no xxx'
' only with NEGATIVE attribute, not with negpmh or negfamily
Dim carryfamily As Boolean, midsequence As Boolean
carryfamily = False: midsequence = False: length = 0
For a = pd.max To 1 Step -1
    If pd.Attr(a) = "family" Then carryfamily = True: length = 0
    If a < pd.max And (in_set(pd.punct(a), ":", ".", "-", ";") Or length > 5) Then
        carryfamily = False
    End If
    ' don't carry family attribute through to previous sentence or clause.
    If carryfamily Then
        If pd.Attr(a) = "" And in_set(Left(pd.mean(a), 4), "", "CLIN") Then
            pd.set_attr "family", a
            length = length + 1
        End If
    End If
    
    ' if midsequence i.e. if in the middle of a list of findings, just before a negative finding
    If pd.text(a) = "or" Then midsequence = False
    If midsequence Then
        If pd.Attr(a) = "negative" Then pd.set_attr "", a
        If in_set(pd.text(a), "no", "not", "neg", "negative") Then
            ' ensure that following clinical term is negative.
            ' continue negative attribute for 2 terms (3 if immediate next term is not CLIN)
            ' or until punctuation.
            If Left(pd.mean(a + 1), 4) = "CLIN" Then b = a + 1 Else b = a + 2
            If pd.punct(b - 1) = "" Then pd.set_attr "negative", b
            If pd.punct(b) = "" Then pd.set_attr "negative", b + 1
        End If
        If (pd.text(a) = "no" Or pd.text(a) = "not") And pd.Attr(a) = "negative" Then
            midsequence = True
        End If
    End If
    If pd.punct(a) = "." Then midsequence = False

Next
End Sub


Sub analyse_pd(Optional debug_ As Boolean, Optional labtest As String)
' Attempts to map sequences of words to Read terms.
Dim st As Long, fin As Long ' start and finish positions
Dim curattr As String ' current attribute
' it will only analyse stretches in which the attribute is the same (or converts to
' negative, cause, etc.) and there is no date.
Dim readmatch As String, counter As Long
' advance to first 'useful' word
st = 1
fin = 1
Do
    Do While (Not in_set(Left(pd.mean(st), 4), "IGNO", "CLIN", "NUMB")) Or _
        in_set(pd.Attr(st), "ignore", "anon", "normalrange", "possibility")
        ' Original code:
        ' keep advancing until the word is in the required group and not to be ignored
        ' it cannot be a number (i.e. program will not recognise a number as start of
        ' clinical term)
        st = st + 1
        If st > pd.max Then Exit Do
    Loop
    If st > pd.max Then Exit Do
    
    curattr = pd.Attr(st)
    fin = st
    Do While fin < pd.max And fin < st + 8 And _
        (in_set(pd.Attr(fin + 1), "", curattr, "negative", "query", "dueto", _
            "causing", "ignore", "pmh", "duraprev", "dateprev", "ageprev")) And _
            (in_set(Left(pd.mean(fin + 1), 4), "IGNO", "CLIN", "NUMB")) And _
            pd.Attr(fin + 1) <> "anon"
        ' keep advancing as long as the word is still in the required group
        If pd.punct(fin) = "," Then
            If strfunc.in_set(Left(pd.mean(fin + 1), 4), "CLIN", "NUMB") Then Exit Do
        End If ' modified 12 sep 2004
        If pd.punct(fin) = "." Then
            If Len(pd.text(fin)) > 2 Then Exit Do
            ' stop at end of sentences, but not if e.g. m.i. or ca.
        End If
        fin = fin + 1
    Loop
    
    Do
        readmatch = list.bestmatch(st, fin, debug_)
        fin = fin - 1
    Loop Until readmatch <> "" Or fin < st
    fin = fin + 1
    
    If readmatch <> "" Then
        For counter = st To fin
            pd.set_mean "READ " & readmatch, counter
        Next
    End If
    
    st = fin + 1
Loop Until st > pd.max

Dim labresult As Boolean
labresult = False
' make the first number into a lab result if current data type is labs
If labtest <> "" Then
    st = 1
    Do
        If in_set(Left(pd.mean(st), 4), "LABS", "DURA", "DATE", "READ") Then Exit Do
        If Left(pd.mean(st), 4) = "NUMB" And pd.Attr(st) = "" Then
            labresult = True
        ElseIf Mid(pd.mean(st), 6) = "negative" And pd.text(st) = "negative" Then
            labresult = True
        ElseIf Mid(pd.mean(st), 6) = "neg" And pd.text(st) = "neg" Then
            labresult = True
        End If
        If labresult Then
            If labtest = "value" Or _
                Not IsNumeric(Mid(pd.mean(st), 6)) Then
                pd.set_mean "LABS " & strfunc.words(pd.mean(st), 2), st
                Exit Do
            End If
        End If
        st = st + 1
    Loop Until st > pd.max
End If

If death = False Then Exit Sub
' extra bit, to deal with a ... b ... (without number 1, 2 etc.) only in Death mode.
st = 9999
If pd.text(1) = "a" And Left(pd.mean(1), 4) <> "READ" Then
    For st = 2 To 4
        If Left(pd.mean(st), 4) = "READ" Then Exit For
    Next
    If st = 5 Then st = 9999
Else ' added 20sep04 to deal with 1a ... b ...
    For st = 1 To pd.max
        If pd.Attr(st) = "deathcause1a" Then Exit For
    Next
End If
If st < pd.max Then
    pd.set_attr "deathcause1a", st
    Do
        If pd.text(st) = "b" And Left(pd.mean(st), 4) <> "READ" And _
            Left(pd.mean(st + 1), 4) = "READ" Then Exit Do
        st = st + 1
        If pd.Attr(st) = "deathcause1b" Then st = 9999: Exit Do
    Loop Until st >= pd.max
End If
If st < pd.max Then
    pd.set_attr "deathcause1b", st + 1
    Do
        If pd.text(st) = "c" And Left(pd.mean(st), 4) <> "READ" And _
            Left(pd.mean(st + 1), 4) = "READ" Then Exit Do
        If pd.Attr(st) = "deathcause1c" Then st = 9999: Exit Do
        st = st + 1
    Loop Until st >= pd.max
End If
If st < pd.max Then
    pd.set_attr "deathcause1c", st + 1
End If

Dim cause1a_passed As Boolean
cause1a_passed = False
st = 1
Do
    If pd.Attr(st) = "deathcause2" Then pd.set_attr "", st ' cannot have cause2 before cause1 !!
    If pd.Attr(st) = "deathcause1a" Then cause1a_passed = True Else st = st + 1
Loop Until cause1a_passed Or st > pd.max
If Not cause1a_passed Then Exit Sub ' if there is no cause1a then no need to do the next bit
Do
    st = st - 1
    If Left(pd.Attr(st), 10) = "deathcause" And pd.Attr(st) <> "deathcause" Then
        pd.set_attr "", st ' if attribute is deathcause category, ignore, because it might be wrong.
    End If
Loop Until st = 0
End Sub


Function remove_ignorable(ByVal instring As String, _
    Optional remove_right_left As Boolean) As String
' Removes ignorable words from a phrase. The argument instring must have
' one space between words and no punctuation.
Dim temp As String, temp2 As String, a As Long
instring = Trim(instring)
temp = "": temp2 = ""
If instring = "" Then Exit Function

' check whether letters r and l are to be considered as left and right
Dim remove_r_l As Boolean
remove_r_l = remove_right_left
If numwords(instring) >= Len(instring) / 2 Then ' acronym
    remove_r_l = False
ElseIf InStr(1, instring, "dopa") Then ' e.g. L dopa
    remove_r_l = False
ElseIf InStr(1, instring, "lumbar") Then ' e.g. L 5 vertebra
    remove_r_l = False
ElseIf InStr(1, instring, "interval") Or InStr(1, instring, "type") Then
    remove_r_l = False
ElseIf InStr(1, instring, "ventric") Or InStr(1, instring, "atrium") Or _
    InStr(1, instring, "atrial") Or InStr(1, instring, "coron") Then
    'always differentiate between ventricles and coronary arteries
    remove_r_l = False
    remove_right_left = False
End If

For a = 1 To strfunc.numwords(instring)
    temp2 = dissect2(instring, , a)
    If remove_r_l Then
        If Not ignorable(temp2) And Not in_set(temp2, "right", "left", "lt", "rt", "l", "r") Then
            temp = temp & temp2 & " "
        End If
    ElseIf remove_right_left Then ' remove rt, lt, right, left only
        ' only remove r and l if not part of an acronym.
        If Not ignorable(temp2) And Not in_set(temp2, "right", "left", "lt", "rt") Then
            temp = temp & temp2 & " "
        End If
    Else ' remove ignorable only
        If Not ignorable(temp2) Then
            temp = temp & temp2 & " "
        End If
    End If
Next
remove_ignorable = Trim(temp)
End Function

Function readscore(pd_start As Long, pd_fin As Long, Termref As Long, _
    Optional debug_ As Boolean, Optional clear_memory As Boolean) As Single
' Returns a score (0 to 100) based on the accuracy and completeness of match between
' a sequence of words in the free text and a candidate Read term.
Const maxrd = 40 ' maximum for read and partdata term
Const mem_max = 8 ' memory of previous searches
If debug_ Then debug_string = ""

Static mem_start(mem_max) As Long
Static mem_fin(mem_max) As Long
Static mem_termref(mem_max) As Long
Static mem_score(mem_max) As Single
Static pos As Integer
pos = pos + 1
If pos > mem_max Then pos = 1
Dim a As Long, b As Long, c As Long, d As Long ' counting variables
' clear memory if necessary
If clear_memory Then
    For a = 1 To mem_max
    mem_termref(a) = 0
    mem_start(a) = 0
    mem_fin(a) = 0
    mem_score(a) = 0
    Next
End If

' see if it has already been done
For a = 1 To mem_max
    If mem_termref(a) = Termref Then
        If mem_start(a) = pd_start Then
            If mem_fin(a) = pd_fin Then
               readscore = mem_score(a): Exit Function
            End If
        End If
    End If
Next
mem_start(pos) = pd_start
mem_fin(pos) = pd_fin
mem_termref(pos) = Termref

Dim Read(maxrd) As String
Dim match(maxrd) As Long ' whether each Read word is matched up, and exactness of match
Dim pd_pos(maxrd) As Long ' position of each Read word in pd (relative to pd_start=1)
Dim num As Long ' number of words in Read term
Dim pd_match(maxrd) As Long ' whether each pd term is matched, and priority
Dim read_std_term As String
Dim match_rightleft As Boolean ' whether the Read term specifies left or right

read_std_term = terms.std_term(Termref)
If InStr(1, read_std_term, " right ") Or InStr(1, read_std_term, " left ") Then
    match_rightleft = True
Else
    match_rightleft = False
End If '  whether or not need to match right and left
read_std_term = Trim(read_std_term)
num = strfunc.numwords(read_std_term)
Dim pd_string As String
pd_string = pd.part_nopunc(pd_start, pd_fin)
Dim temp As Long, tempstr As String, tempmatch As Single, anymatch As Boolean
anymatch = False

Dim read_attr As String ' read attribute string
read_attr = terms.attrib_str(Termref)
' if first pd attribute is negative and read term is all TRUE, invert the Read attribute
If Left(pd.Attr(pd_start), 3) = "neg" And Not InStr(1, read_attr, "F") Then
    read_attr = replace(read_attr, "T", "F")
End If

' first check for exact match
If pd_string = read_std_term Then readscore = 100: Exit Function

' transfer read term to array
For a = 1 To num
    Read(a) = dissect2(read_std_term, , a)
Next ' next a (i.e. Read)

Dim cur_true As Boolean
' try using synonym dictionary
' Read term is allowed to be wider than pd but not more specific
a = 1 'word position in Read term
Do
    ' move to next unattached, non-ignorable word of Read term
    Do While (Mid(read_attr, a, 1) = "I" Or match(a) <> 0) And a < num
        a = a + 1
    Loop
        
    ' find out whether this phrase is supposed to be true or not
    If Mid(read_attr, a, 1) = "F" Then cur_true = False Else cur_true = True
    
    b = a
    ' b is word position in Read term at end of current phrase
    ' b=a if the phrase is a single word
    If (Not Mid(read_attr, a, 1) = "I") And match(a) = 0 And a <= num Then
    ' move to end of unattached phrase
        If cur_true = False Then
            Do While Mid(read_attr, a, 1) = "F" And match(a) = 0 And b <= num And b < a + 6
                b = b + 1
            Loop ' find the end of the 'false' phrase
        Else
            Do While Mid(read_attr, a, 1) = "T" And match(a) = 0 And b <= num And b < a + 6
                b = b + 1
            Loop ' find the end of the 'true' phrase
        End If
        If b > a Then b = b - 1
        
        ' b is now the end of the largest Read phrase to try to match
        tempstr = synonym.trylink_2(words(read_std_term, a, , b), pd_start, pd_fin, cur_true)
        If tempstr <> "" Then b = a + Val(dissect2(tempstr, , 4)) - 1 Else b = a
        ' b is now end of current phrase being matched
        
        ' if match found, analyse it and add results to the arrays
        If tempstr <> "" Then
            If debug_ Then debug_string = debug_string & "Link: " & temp & " : " & _
                strfunc.words(read_std_term, a, , b) & " --> " & pd.part_nopunc(pd_start, pd_fin) & _
                Chr$(13) & Chr$(10) & "Tempstr: " & Chr$(13) & Chr$(10)
            For c = a To b
                match(c) = Val(dissect2(tempstr, , 1)) ' allocating the match - for Read
                pd_pos(c) = Val(dissect2(tempstr, , 2)) ' recording the pd position of Read term word
            Next
            For c = Val(dissect2(tempstr, , 2)) To Val(dissect2(tempstr, , 3))
                pd_match(c) = match(a) ' c is the position in this section of partdata
            Next
            anymatch = True
            
        Else
            
            For c = pd_start To pd_fin
                If Read(a) = pd.text(c) Then ' added 18 jan 04 (moved from prev position)
                    If pd.true_(c) = cur_true Then
                        If debug_ Then
                            debug_string = debug_string & "Word: " & Read(a) & " --> " & pd.text(c) & Chr$(13) & Chr$(10)
                        End If
                        match(a) = 6: pd_match(c - pd_start + 1) = 2
                        Exit For
                    Else
                        If debug_ Then debug_string = debug_string & "True/false mismatch" _
                            & Chr$(13) & Chr$(10)
                        GoTo final ' zero readscore
                    End If
                End If
            Next
        End If
    End If
    a = b + 1
Loop Until a > num

' Now to add up the readscore
Dim tot As Single, ok As Single, scor As Single ' number of linked words and total
Dim optiondone As Boolean
Dim partscore As Single

' Read term
tot = 0: ok = 0: scor = 0: optiondone = False
For a = 1 To num
    If Mid(read_attr, a, 1) = "T" Then
        tot = tot + 1
        If match(a) > 0 Then ok = ok + 1
    ElseIf Mid(read_attr, a, 1) = "F" Then
        tot = tot + 1
        If match(a) > 0 Then ok = ok + 1 Else ok = ok + 0.5
    ElseIf Mid(read_attr, a, 5) = "OIOIO" Then
        If match(a) > 0 Then optiondone = True
    ElseIf Mid(read_attr, a, 3) = "OIO" Then
        If match(a) > 0 Then optiondone = True
    ElseIf Mid(read_attr, a, 1) = "O" Then
        If match(a) > 0 Then optiondone = True
        If optiondone = True Then ok = ok + 1
        tot = tot + 1
        optiondone = False
    ElseIf Mid(read_attr, a, 1) = "I" Then
        If match(a) = 0 Then ok = ok - 0.1 ' lose a few points for unmatched ignorable words
    End If
    If match(a) < 0 Then GoTo final ' bad match
    scor = scor + match(a) / 6
Next
If tot = 0 Then GoTo final
partscore = 63 * (ok / tot) ' completeness
partscore = partscore + 7 * (scor / tot)  ' accuracy
partscore = partscore - 27

If partscore < 33 Then GoTo final
' try to match any unmatched terms in candidate text
' but only bother if Read term already matched.
For a = pd_start To pd_fin ' position in pd
    If pd_match(a - pd_start + 1) = 0 Then
        For b = 1 To num
            If in_set(pd.text(a), "right", "left", "rt", "lt", "r", "l") And match_rightleft = False Then
                pd_match(a - pd_start + 1) = 6 ' array pd_match, which applies to current phrase only
                Exit For
            End If
            If in_set(pd.text(a), "not", "without", "absent", "lack", "no") And a < pd_fin Then
                pd_match(a - pd_start + 1) = 6 ' array pd_match, which applies to current phrase only
                Exit For ' these negative words would have influenced the attributes
            End If
            If Left(Attr(a), 3) = "neg" Then
                tempstr = synonym.trylink_2(Read(b), a, a, False)
            Else
                tempstr = synonym.trylink_2(Read(b), a, a, True)
            End If
            If tempstr <> "" Then
                If debug_ Then debug_string = debug_string & "Link: " & temp & " : " & _
                    Read(b) & " --> " & pd.text(a) & _
                    Chr$(13) & Chr$(10) & "Tempstr: " & Chr$(13) & Chr$(10)
                pd_match(a - pd_start + 1) = Val(dissect2(tempstr, , 1))
            End If
        Next
    End If
Next

If debug_ Then
    debug_string = debug_string & "CANDIDATE TEXT" & Chr$(13) & Chr$(10)
    For a = pd_start To pd_fin
        debug_string = debug_string & pd.text(a) & "  " & pd.Attr(a) & "  " & pd_match(a - pd_start + 1) & _
            Chr$(13) & Chr$(10)
    Next
    debug_string = debug_string & Chr$(13) & Chr$(10) & "READ TERM" & Chr$(13) & Chr$(10)
    For a = 1 To num
        debug_string = debug_string & Read(a) & " " & Mid(read_attr, a, 1) & " " & match(a) & _
            " --> " & pd_pos(a) & " " & pd.text(pd_pos(a)) & Chr$(13) & Chr$(10)
    Next
End If


' candidate text (get bonus scor marks for converting ignorable words)
tot = 0: ok = 0: scor = 0
For a = pd_start To pd_fin
    If Not ignorable(pd.text(a)) Then
        tot = tot + 1
        If pd_match(a - pd_start + 1) > 0 Then ok = ok + 1
    End If
    If pd_match(a - pd_start + 1) < 0 Then
        If debug_ Then debug_string = debug_string & "Pd match " & a - pd_start + 1 & " : " & _
            pd_match(a - pd_start + 1) & Chr$(13) & Chr$(10)
        GoTo final ' for bad matches
    End If
    scor = scor + pd_match(a - pd_start + 1) / 6
Next
If tot = 0 Then
    readscore = 0
    mem_score(pos) = 0
    GoTo final
End If
partscore = partscore + 52 * (ok / tot) ' completeness
partscore = partscore + 5 * (scor / tot) ' accuracy

If partscore < 0 Then
    readscore = 0
ElseIf partscore > 100 Then
    readscore = 100
Else
    readscore = partscore
End If

final:
mem_score(pos) = readscore
End Function

Function fuzzylink(ref_word As String, test_word As String) As Long
' Whether the two words are almost the same (maximum one character difference).
' Assume the first character is the same and they differ in length by at most 1.
' Gives a score (letter position of difference, zero if too different).
Dim a As Long, b As Long, error As Boolean, wordlength As Long
b = 1: error = False: wordlength = Len(test_word)

fuzzylink = wordlength

For a = 2 To wordlength
    b = b + 1
    If Mid(ref_word, b, 1) <> Mid(test_word, a, 1) Then
        If error = True Then
            If fuzzylink = a - 1 Then
                If Mid(ref_word, a - 1, 1) <> Mid(test_word, a, 1) Then fuzzylink = 0: Exit Function
                If Mid(ref_word, a, 1) <> Mid(test_word, a - 1, 1) Then fuzzylink = 0: Exit Function
                ' if two adjacent letters are swapped the fuzzylink is maintained.
            Else
                If wordlength > 9 And wordlength = a Then
                    fuzzylink = fuzzylink - 1
                Else
                    fuzzylink = 0: Exit Function
                End If
            End If
            ' up to 2 extra errors allowed at the end of the word as long as it is >9 chars long
        Else
            error = True: fuzzylink = a
            If Len(ref_word) = 1 + wordlength Then b = b + 1
        End If
    End If
Next
End Function

