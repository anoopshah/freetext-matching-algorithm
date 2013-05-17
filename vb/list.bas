Option Compare Binary
Option Explicit

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

' Module: list
' Code for handling lists of phrase variants for matching

Const maxtermlist = 50 ' maximum number of terms to consider
Const threshold_high = 91 ' (for readscore - don't analyse further)
Const threshold = 87 ' (for readscore - minimum)

Type termlist
    Termref(maxtermlist) As Long ' medcode that the term maps to
    words(maxtermlist) As String ' phrase variant which maps to this medcode
    score(maxtermlist) As Single ' readscore
    num As Long ' number of terms in termlist
End Type

Function bestmatch(pd_start As Long, pd_fin As Long, Optional debug_ As Boolean) As String
' Returns a string containing the medcode and readscore for the best possible Read term match
' for a portion of the text. The match may be to an 'alternate' Read term, which is
' converted to the linked preferred term in the final output.
Dim candidate As termlist  ' list of candidate Read termrefs
Dim word As String 'search strings to use for words
Dim max_rs_val As Single ' maximum readscore
Dim max_rs_pos As Long ' position in list of term with maximum readscore
Dim curscore As Single
Dim searchstring As String
Dim a As Long, b As Long  ' counter
searchstring = ""
Dim found As Boolean

' 0. Try exact match
If terms.exact_read_termref(pd.part_nopunc(pd_start, pd_fin)) > 0 Then
    bestmatch = Format(terms.exact_read_termref(pd.part_nopunc(pd_start, pd_fin))) & " 100"
    Exit Function
End If

' 1. generate list of words to search
For a = pd_start To pd_fin
    searchstring = searchstring & " " & pd.text(a)
Next

' 2. generate initial matchlist
candidate = getlist(searchstring, pd_start, pd_fin, 0)

For b = 1 To 3
    ' 3 analyse matchlist - select best match
    max_rs_val = 0: max_rs_pos = 0
    If candidate.num > 0 Then
        For a = candidate.num To 1 Step -1
            If candidate.score(a) > max_rs_val Then
                max_rs_val = candidate.score(a)
                max_rs_pos = a
            End If
            If candidate.score(a) > threshold_high Then Exit For
        Next
    End If
    
    ' 4 if the best match is higher than high threshold, or higher than low threshold and
    ' set has already been expanded, select this match
    If max_rs_val > threshold_high Or (max_rs_val > threshold And b > 1) Then
        bestmatch = Format(candidate.Termref(max_rs_pos)) & " " & Format(max_rs_val)
        If debug_ Then Call display(candidate)
        Exit Function
    End If
    
    ' 5 if none of the matches has a high enough score, expand and try again.
    If b = 1 Then candidate = expand(candidate, pd_start, pd_fin, 0)
    If b = 2 Then candidate = expand(candidate, pd_start, pd_fin, 0) ' leeway=0
Next

If debug_ Then Call display(candidate)
End Function

Function expand(in_list As termlist, Optional pd_start As Long, Optional pd_fin As Long, _
    Optional leeway As Long) As termlist
' Returns a termlist which expands the input termlist by generating variants of the
' text fragment using the synonym table. The function searches for up to 5 words at a
' time, starting with longer possible matches.
Dim in_list_num As Long
in_list_num = in_list.num
Dim numwd As Long ' number of words in phrase
Dim cur_words As String ' don't repeat the search for duplicate words
cur_words = ""
expand = in_list
Dim top As Long, pos As Long
Dim candidate As String, maxwords As Long, newwords As String
Dim a As Long, b As Long, c As Long ' looping variables
Dim temp_termlist As termlist

For a = 0 To in_list_num ' looping through terms
    If cur_words <> in_list.words(a) Then ' this is a new distinct phrase
        cur_words = in_list.words(a)
        ' dissect2 into groups of words
        
        numwd = strfunc.numwords(cur_words)
        If numwd < 5 Then maxwords = numwd Else maxwords = 5
        
        For b = maxwords To 1 Step -1 ' number of words in phrase to search for (start with 5)
            For c = 1 To numwd - b + 1 ' position of first word of phrase
                candidate = strfunc.words(cur_words, c, b)
                ' search for dictionary matches
                top = synonym.s1_pos(candidate)
                If top > 0 Then
                ' for each match found: -
                    pos = top
                    Do
                        If synonym.s1_priority(pos) > 0 Then
                            ' i.e. not an 'opposite' match, with priority -100
                            ' generate new phrase
                            If c = 1 Then
                                newwords = synonym.s2(pos) & " " & _
                                    strfunc.words(cur_words, b + c, 100)
                            Else
                                newwords = strfunc.words(cur_words, 1, c - 1) & " " & _
                                    synonym.s2(pos) & " " & strfunc.words(cur_words, b + c, numwd)
                            End If
                            
                            ' generate new set of termrefs  and  append to existing list
                            temp_termlist = getlist(newwords, pd_start, pd_fin, leeway)
                            If temp_termlist.num > 0 Then
                                expand = add_termlists(expand, temp_termlist)
                            Else
                                ' record that this search has been attempted
                                expand.num = expand.num + 1
                                expand.Termref(expand.num) = 0
                                expand.score(expand.num) = 0
                                expand.words(expand.num) = newwords
                            End If
                            If expand.score(expand.num) > threshold_high Then Exit Function
                                ' to make it faster
                            If expand.num > maxtermlist - 10 Then Exit Function
                                ' to prevent excessively long lists
                        End If

                        pos = pos + 1
                    Loop Until synonym.s1(pos) <> candidate
                End If  ' if top>0
            Next ' next c; next sub-phrase
        Next ' next b; next number of words in sub-phrase
    End If
Next ' next a; next term

End Function

Function add_termlists(t1 As termlist, t2 As termlist) As termlist
' Appends one termlist to another, and returns the combined termlist.
add_termlists = t1
If t2.num = 0 Then Exit Function
Dim a As Long
For a = 1 To t2.num
    If a + t1.num > maxtermlist Then Exit For
    add_termlists.words(a + t1.num) = t2.words(a)
    add_termlists.Termref(a + t1.num) = t2.Termref(a)
    add_termlists.score(a + t1.num) = t2.score(a)
Next
add_termlists.num = t1.num + a - 1
End Function

Function getlist(ByVal words As String, Optional pd_start As Long, _
    Optional pd_fin As Long, Optional leeway As Long) As termlist
' Creates a list of potential Read term matches to a text phrase, returning
' the result as a termlist object. Calculates the readscore for each match.
' If no matches are found, the function removes the words 'left' and 'right'
' from the text and tries again (by recursion). The leeway argument is
' currently not used, but it may be possible in the future to alter this
' to allow the function to attempt to match to terms with a different
' number of non-ignorable words.

words = remove_ignorable(words, False)
If words = "" Then Exit Function

Dim curscore As Single ' current readscore

getlist.words(0) = words
getlist.num = 0

Dim numwords As Long: numwords = strfunc.numwords(words)

' Get a list of termrefs matching the bag of words.
' Calculate the readscore for each, and return if one is greater than threshold.
Dim term_top As Long, term_bot As Long, thisbag As String
Dim i As Long, a As Long
thisbag = bag_of_words(words)
term_top = terms.pos_bagofwords(True, thisbag)
term_bot = terms.pos_bagofwords(False, thisbag)
a = 1 ' counter for getlist
    
If term_top <> 0 And term_bot <> 0 Then
    For i = term_top To term_bot
        curscore = readscore(pd_start, pd_fin, terms.termref_bagofwords(i))
        If curscore > threshold Then
            getlist.words(a) = words
            getlist.Termref(a) = terms.termref_bagofwords(i)
            getlist.score(a) = curscore
            a = a + 1
        End If
        If curscore > threshold_high Then getlist.num = a - 1: GoTo enough
            ' no need to bother searching for more terms
        If a >= maxtermlist Then Exit For
    Next
    getlist.num = a - 1
End If

' if no terms entered with high enough threshold, record the phrase tested.
enough:
If getlist.num > 0 Then Exit Function
Dim tempwords As String
tempwords = remove_ignorable(words, True)
If tempwords = words Then getlist.words(1) = words: Exit Function ' no left/right
words = tempwords
getlist = getlist(words, pd_start, pd_fin, leeway)     ' recursive
' the next time in the loop tempwords will equal words so the function
' will be called at most twice.
End Function

Sub display(in_list As termlist)
' Adds the contents of termlist to the debug string. This is used when
' running the program in test mode, to produce an analysis report for a single
' text.
Dim a As Long
debug_string = debug_string & Chr$(13) & Chr$(10) & "List of candidate terms" & _
    Chr$(13) & Chr$(10)
For a = 0 To in_list.num
    If in_list.Termref(a) = 0 Then
        debug_string = debug_string & in_list.score(a) & " " & in_list.Termref(a) & " " & _
            std_term(in_list.Termref(a)) & " [" & in_list.words(a) & "]" & Chr$(13) & Chr$(10)
    Else
        debug_string = debug_string & in_list.score(a) & " " & in_list.Termref(a) & " " & _
            std_term(in_list.Termref(a)) & Chr$(13) & Chr$(10) & "          [" & _
            in_list.words(a) & "]" & Chr$(13) & Chr$(10)
    End If
Next
debug_string = debug_string & "Total " & in_list.num & " terms" & Chr$(13) & Chr$(10)
End Sub

