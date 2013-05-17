Option Compare Binary ' CASE SENSITIVE.
Option Explicit

' Module: wordlist -- clinical and non-clinical words for spelling correction

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

Const maxwords = 100000 ' Maximum number of entries in the 'w' arrays (list of all words)
Const maxignore = 100 ' Number of entries in the 'ignore' table
Const maxletters = 30 ' Number of letters per word

' 1 ALL - a list of all individual words in Read terms used for matching
' (i.e. Read std_term for table a in the terms table) and non-clinical words
' from a general wordlist. Sorted by wordlength then alphabetically.
Dim w_words(maxwords) As String ' array of clinical and non-clinical words (no duplicates)
Dim w_clinical(maxwords) As Boolean ' whether the word is possibly part of a clinical term
Dim w_top(maxletters) As Long ' start position for words of different lengths
Dim w_max As Long ' total number of words

' 2 IGNORABLE list
Dim ignorelist(maxignore) As String ' words which can be ignored e.g. if, and, of, the
Dim ignorelistnum As Long ' number of words in ignorable list

' 3 IGNORABLE PHRASES (only used for initial processing of text phrase)
Dim ignorephrase(maxignore) As String ' words which can be ignored e.g. if, and, of, the
Dim ignorephrasenum As Long ' number of phrases in ignorable phrases list


Sub add_to_wordlist(words_to_add As String)
' Adds a word or words to wordlist, and automatically sorts and compresses
' the wordlist when necessary
If words_to_add = "" Then Exit Sub

Dim nwords As Long, i As Long
nwords = strfunc.numwords(words_to_add)

' If approaching within 10 words of the maximum permitted number of words ...
If w_max > maxwords - nwords - 10 Then
    sort_and_compress_wordlist
End If

For i = 1 To nwords
    If w_max + i > maxwords Then
        Exit Sub
    End If
    w_words(w_max + i) = strfunc.words(words_to_add, i)
Next

' Update the total number of words
w_max = w_max + nwords

End Sub

Function import_wordlist(wordlistfile As String) As String
' Creates a list of words in clinical terms and other English words.
' Gets text words from the synonyms table. Returns a string stating
' what was imported.

' First set the existing number of words in wordlist to zero
w_max = 0

Dim rawtext As String, i As Long, j As Long
For i = 1 To synonym.numrows
    add_to_wordlist synonym.s1(i)
Next

' Get words from Read std_terms to be coded to bag of words
For i = 1 To terms.numrows_a_c
    add_to_wordlist terms.get_bagofwords(i)
Next

' Other English words in 2of4brif
i = 0 ' to count number of words read
Dim fileno As Integer, rawstring As String
fileno = freefile
Open wordlistfile For Input As #fileno
Do While Not EOF(fileno)
    Line Input #fileno, rawstring
    rawstring = Trim(rawstring)
    If rawstring <> "" Then
        i = i + 1
        add_to_wordlist rawstring & "_" ' to mark it as non-clinical
    End If
Loop
Close #fileno

import_wordlist = "Loaded " & CStr(i) & " words from " & wordlistfile

' Add the number of words at the beginning of the word and then re-sort
For i = 1 To w_max
    If Right(w_words(i), 1) = "_" Then
        w_words(i) = Format(Len(w_words(i)) - 1, "00") & w_words(i)
    Else
        w_words(i) = Format(Len(w_words(i)), "00") & w_words(i)
    End If
Next

sort_and_compress_wordlist

' Now place the clinical information in the clinical variable. Non clinical
' words are denoted by the _ suffix. Also remove the number of characters
' prefix from each word and fill in the w_top array
' Remove non-clinical words if there is also a clinical word.
Dim curwordlength As Long, thiswordlength As Long ' current word length
Dim wordlengths As Long ' a counter for wordlengths
w_top(1) = 1
curwordlength = 1
j = 0 ' position to write to
For i = 1 To w_max
    thiswordlength = CLng(Left(w_words(i), 2))
    ' If it ends in _, it is a non-clinical word
    If Mid(w_words(i), 3) <> w_words(j) & "_" Then
        j = j + 1
        ' Check that it is not a non-clinical version of the previous word
        If Right(w_words(i), 1) = "_" Then
            w_clinical(j) = False
            ' remove the prefix (number of letters) and suffix (_)
            w_words(j) = Mid(Left(w_words(i), Len(w_words(i)) - 1), 3)
        Else
            w_clinical(j) = True
            ' remove the prefix (number of letters)
            w_words(j) = Mid(w_words(i), 3)
        End If
    End If
    
    If thiswordlength > curwordlength Then
        ' start of a section of longer words
        For wordlengths = curwordlength + 1 To thiswordlength
            If wordlengths < maxletters Then
                ' Position of start of words of this length
                w_top(wordlengths) = j
            End If
        Next
        curwordlength = thiswordlength
    End If
Next
' Reset the maximum number of terms
w_max = j

import_wordlist = import_wordlist & _
    "; total number of distinct words in this wordlist, Read terms or synonyms = " & w_max

' Fill in the remaining wordlengths
If curwordlength < maxletters Then
    For j = curwordlength + 1 To maxletters
        w_top(j) = w_max + 1
    Next
End If
End Function

Sub sort_and_compress_wordlist()
' Sorts the wordlist and removes duplicates. All words are preceded by the
' number of letters so they are sorted by number of letters then the text (alphabetically)
quicksort w_words, 1, w_max
Dim readpos As Long, writepos As Long
readpos = 1
writepos = 1
Do While readpos < w_max
    Do While Mid(w_words(readpos), 3) = Mid(w_words(writepos), 3) And readpos < w_max
        readpos = readpos + 1
    Loop
    writepos = writepos + 1
    If readpos <> writepos Then
        w_words(writepos) = w_words(readpos)
    End If
Loop
w_max = writepos - 1
End Sub

Sub quicksort(ByRef tosort As Variant, ByVal start As Long, ByVal finish As Long)
' Sorts a vector of strings
Dim swap As String
Dim low_pos As Long, high_pos As Long
Dim pivot As String

' Initial positions of high and low markers
low_pos = start
high_pos = finish
pivot = tosort((start + finish) \ 2)

Do
    While tosort(low_pos) < pivot
        low_pos = low_pos + 1
    Wend
   
    While tosort(high_pos) > pivot
        high_pos = high_pos - 1
    Wend
    
    If low_pos <= high_pos Then
        swap = tosort(low_pos)
        tosort(low_pos) = tosort(high_pos)
        tosort(high_pos) = swap
        low_pos = low_pos + 1
        high_pos = high_pos - 1
    End If
Loop Until low_pos > high_pos

If high_pos > start Then quicksort tosort, start, high_pos
If low_pos < finish Then quicksort tosort, low_pos, finish
End Sub

Function import_ignore(ignorefile As String, ignorephrase_file As String) As String
' Imports ignore.txt and ignorephrase.txt. ignore.txt should be sorted alphabetically
' but is re-sorted to ensure that the string comparison order is identical to that which
' will be used for binary searching. ignorephrase.txt contains semi-structured phrases
' which might be found in the raw text and should be ignored. Neither file has a header row.
Dim fileno As Integer, i As Integer, rawinput As String

fileno = freefile
Open ignorefile For Input As fileno
ignorelistnum = 0
Do
    Line Input #fileno, rawinput
    If rawinput <> "" Then
        ignorelistnum = ignorelistnum + 1
        ignorelist(ignorelistnum) = Trim(rawinput)
    End If
Loop Until EOF(fileno)
Close #fileno

' Sort the ignorable words list
quicksort ignorelist, 1, ignorelistnum

Open ignorephrase_file For Input As fileno
ignorephrasenum = 0
Do
    Line Input #fileno, rawinput
    If rawinput <> "" Then
        ignorephrasenum = ignorephrasenum + 1
        ignorephrase(ignorephrasenum) = Trim(rawinput)
    End If
Loop Until EOF(fileno)
Close #fileno

import_ignore = "Loaded " & CStr(ignorelistnum) & " words/phrases to ignore from " & _
    ignorefile & " and " & CStr(ignorephrasenum) & " raw phrases to ignore from " & _
    ignorephrase_file
End Function

Function in_wordlist(instring As String) As String
' Returns CLIN for clinical words, WORD for non-clinical words and
' blank for words not found in the wordlist.

Dim top As Long, bot As Long, trial As Long, wordlength As Long
in_wordlist = ""
wordlength = Len(instring)
top = w_top(wordlength): bot = w_top(wordlength + 1) - 1
If top = 0 Or bot < 1 Or bot < top Then Exit Function
Do
    trial = Int((top + bot) / 2)
    If w_words(trial) > instring Then
        bot = trial - 1
    ElseIf w_words(trial) < instring Then
        top = trial + 1
    ElseIf w_words(trial) = instring Then
        top = trial: bot = trial
    End If
Loop Until bot - top < 2
If w_words(top) = instring Then
    If w_clinical(top) = True Then in_wordlist = "CLIN" Else in_wordlist = "WORD"
ElseIf w_words(bot) = instring Then
    If w_clinical(bot) = True Then in_wordlist = "CLIN" Else in_wordlist = "WORD"
End If
End Function

Function approx_wordlist(instring As String) As Long
' Approximate position of a word in the wordlist (sorted by wordlength, then word)

Dim top As Long, bot As Long, wordlength As Long
wordlength = Len(instring)
top = w_top(wordlength): bot = w_top(wordlength + 1) - 1
If top = 0 Or bot < 1 Or bot < top Then Exit Function
Do
    approx_wordlist = Int((top + bot) / 2)
    If w_words(approx_wordlist) > instring Then
        bot = approx_wordlist - 1
    ElseIf w_words(approx_wordlist) < instring Then
        top = approx_wordlist + 1
    End If
Loop Until bot - top < 2
End Function

Function wordsearch(ByVal word As String, Optional do_spellcheck As Boolean) As String
' Tries to convert a word into a standard form (or without spelling mistakes) which is in
' wordlist. Returns CLIN (for a clinical word) or WORD (for any other word) followed by
' the correctly spelled word, or blank if the spelling cannot be corrected

Dim temp As String
temp = in_wordlist(word)
If temp <> "" Then wordsearch = temp & " " & word: Exit Function

If Len(word) < 6 Or do_spellcheck = False Then Exit Function

' look for spelling mistakes
Dim maxpos As Long, maxscore As Long, tempscore As Long
Dim a As Long, b As Long, top As Long, bot As Long

' missed out single letter - search for same length, then one letter longer
For b = 2 To 1 Step -1
    top = approx_wordlist(Left(word, 2) & Left("aaaaaaaaaaaaaaaaaaaa", Len(word) - b))
    bot = approx_wordlist(Left(word, 2) & Left("zzzzzzzzzzzzzzzzzzzz", Len(word) - b))
    maxscore = 0: maxpos = 0
    For a = top To bot
        tempscore = fuzzylink(w_words(a), word)
        If tempscore > maxscore Then
            maxscore = fuzzylink(w_words(a), word)
            maxpos = a
        ElseIf tempscore = maxscore Then
            maxpos = 0 ' more than one possible link --> dont link unless a better link found
        End If
    Next
    If maxpos > 0 Then
        If w_clinical(maxpos) = True Then
            wordsearch = "CLIN " & w_words(maxpos)
        Else
            wordsearch = "WORD " & w_words(maxpos)
        End If
        Exit Function
    End If
Next

If Right(word, 1) = "s" Then wordsearch = wordsearch(Left(word, Len(word) - 1))
End Function

Function ignorable(instring As String) As Boolean
' Returns True if a word is in the ignorable list for Read matching.
' Uses a binary search algorithm. The ignorable list must be sorted.

ignorable = False
If ignorelistnum = 0 Then Exit Function
Dim top As Integer, bot As Integer, a As Integer
top = 1: bot = ignorelistnum
a = Int(ignorelistnum / 2)
Do
    If ignorelist(a) = instring Then
        ignorable = True
        Exit Function
    ElseIf top + 1 = bot Then
        If ignorelist(top) = instring Then ignorable = True
        If ignorelist(bot) = instring Then ignorable = True
        Exit Function
    ElseIf instring < ignorelist(a) Then
        bot = a
        a = Int((top + bot) / 2)
    Else
        top = a
        a = Int((top + bot) / 2)
    End If
Loop
End Function

Function remove_ignore_phrases(instring As String) As String
' Returns instring with phrases found in 'ignorephrase' list removed. The function
' tries each phrase to remove in turn in the order they appear in the table.
Dim a As Integer, searchtext As String
remove_ignore_phrases = instring
For a = 1 To ignorephrasenum
    searchtext = ignorephrase(a)
    If Left(searchtext, 6) = "START_" Then
        If Mid(searchtext, 7) = Left(remove_ignore_phrases, Len(searchtext) - 6) Then
            remove_ignore_phrases = Mid(remove_ignore_phrases, Len(searchtext) - 5)
        End If
    Else
        remove_ignore_phrases = replace(remove_ignore_phrases, searchtext, "")
    End If
Next
a = InStr(1, instring, "Hospital=")
If a > 1 Then
    remove_ignore_phrases = Left(instring, a - 1) & " " & initial_process(Mid(instring, a))
Else
    remove_ignore_phrases = initial_process(instring)
End If
End Function

Function initial_process(instring As String) As String
' Pre-processor to remove semi-structured Vision text in specific formats.

Dim spec As String, priv As Boolean, cancelled As Boolean, urgent As Boolean
Dim temp As Long, temp2 As Long
Select Case Left(instring, 9)
Case ":SOURCE N", ":SOURCE G" ' cervical smear
    temp = InStr(9, instring, "NON-CANCER RESULT")
    If temp > 0 Then
        temp2 = InStr(temp, instring, ":")
        If temp2 = 0 Then temp2 = 500
        initial_process = "cervical smear " & _
            replace(Mid(instring, 17 + temp, temp2 - temp - 17), "None", "negative")
        ' pasting result of cervical smear
        temp = InStr(9, instring, "INFECTION")
        If temp > 0 Then
            temp2 = InStr(temp, instring, ":")
            If temp2 = 0 Then temp2 = 500
            initial_process = initial_process & ". " & _
                replace(Mid(instring, 9 + temp, temp2 - temp - 9), "None", "no") & " infection"
            ' pasting infection result
        End If
    End If
Case ":LOCATION"
    initial_process = "": Exit Function
Case ":HOSPITAL"
    ' colons separate structured data; capital field names;
    If InStr(9, instring, "CATEGORY Private") Then priv = True
    If InStr(9, instring, "URGENCY Immediate") Then urgent = True
    temp = InStr(9, instring, "DEPARTMENT")
    If temp > 0 Then
        temp2 = InStr(temp, instring, ":") ' finding next colon after department name
        If temp2 = 0 Then temp2 = 500
        If urgent Then initial_process = "Urgent referral. "
        If priv Then initial_process = initial_process & "Private "
        initial_process = initial_process & "referral to " & _
            Trim(Mid(instring, 10 + temp, temp2 - temp - 10))
        If initial_process = "referral to " Then initial_process = instring: Exit Function
        ' pasting clinical specialty
        temp = InStr(9, instring, "COMPLAINT")
        If temp > 0 Then
            temp2 = InStr(temp, instring, ":")
            If temp2 = 0 Then temp2 = 500
            initial_process = initial_process & ". " & Mid(instring, 9 + temp, temp2 - temp - 9)
            ' pasting indication for referral
        End If
    End If
Case "Hospital="
    If InStr(9, instring, "Private=Yes") Then priv = True
    If InStr(9, instring, "Outcome=Cancelled") Then cancelled = True
    If InStr(9, instring, "Urgency=Immediate") Then urgent = True
    ' colon and =; lower case field names
    temp = InStr(9, instring, "Specialisation=")
    If temp > 0 Then
        temp2 = InStr(temp, instring, ":") ' finding next colon after department name
        If temp2 = 0 Then temp2 = 500
        If urgent Then initial_process = "Urgent referral. "
        If priv Then initial_process = initial_process & "Private "
        initial_process = initial_process & "referral to " & _
            Trim(Mid(instring, 15 + temp, temp2 - temp - 15))
        If initial_process = "referral to " Then initial_process = instring: Exit Function
        ' pasting clinical specialty
        temp = InStr(9, instring, "Reason=")
        If temp > 0 Then
            temp2 = InStr(temp, instring, ":")
            If temp2 = 0 Then temp2 = 500
            initial_process = initial_process & ". " & Mid(instring, 7 + temp, temp2 - temp - 7)
            ' pasting indication for referral
        End If
        If cancelled = True Then initial_process = initial_process & ". Appt cancelled by doctor"
    End If
Case "Location="
    If InStr(9, instring, "Private Referral") Then priv = True
Case "Hospital:"
    If InStr(9, instring, "Consultant:") Then
        initial_process = Mid(instring, InStr(9, instring, "Reason:") + 7)
    End If
End Select

If initial_process = "" Then initial_process = instring

End Function

