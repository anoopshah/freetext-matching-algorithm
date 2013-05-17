Option Compare Binary
Option Explicit

' Module: terms --  Read terms as used by the program

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

Const max_usedterms = 100000
Const max_allterms = 150000

Dim a_std_term(max_usedterms) As String ' array of std_term (sorted) to get termref
Dim a_termref(max_usedterms) As Long ' termref (medcode) linked to a_std_term
Dim a_terms_used As Long ' number of entries in a_std_term and a_termref

Dim b_termref(max_allterms) As Long ' all terms (native, virtual or alternate), sorted by termref
' All termrefs that may be associated with a Read code in the text are included,
' as are any virtual or alternate terms.
' Terms marked as 'include' are transferred to table a. This table is ordered by termref
' (must be ordered when loaded, as it is not sorted by the program)
Dim b_std_term(max_allterms) As String ' standardised Read term
Dim b_attrib_str(max_allterms) As String ' attribute string
Dim b_type(max_allterms) As String ' data type of Read term (pregnancy, labtest, death etc.)
Dim b_linkto(max_allterms) As Long ' the actual medcode in the output
Dim b_terms_used As Long ' number of records in the 'b' arrays

Dim c_bagofwords(max_usedterms) As String ' sorted array of 'bag of non-ignorable words'
Dim c_termref(max_usedterms) As Long ' termref (medcode) for each bag of words

' Headers for lookup files
Const headerNative = "medcode" & vbTab & "readcode" & vbTab & "term" & vbTab & _
    "stdterm" & vbTab & "attrstring" & vbTab & "include" & vbTab & "type" & vbTab & "comment" _
    ' headings for nativeterms lookup file
Const headerVirtual = "medcode" & vbTab & "term" & vbTab & _
    "stdterm" & vbTab & "attrstring" & vbTab & "comment" _
    ' headings for virtualterms lookup file
Const headerAlternate = "medcode" & vbTab & "stdterm" & vbTab & "attrstring" & vbTab & _
    "linkto" & vbTab & "comment" _
    ' headings for alternateterms lookup file

    
Function numrows_a_c() As Long
' Returns a_terms_used for use by external functions. Tables a and c contain
' only terms used to match to
numrows_a_c = a_terms_used
End Function

Function numrows_b() As Long
' Returns b_terms_used for use by external functions. Table b contains all
' Read terms, including those that may be associated with text but are not used
' for matching.
numrows_b = b_terms_used
End Function

Function get_bagofwords(pos As Long) As String
' Returns the value of c_bagofwords for a particular position, for use
' by external functions.
If pos > a_terms_used Then
    get_bagofwords = ""
Else
    get_bagofwords = c_bagofwords(pos)
End If
End Function

Function import(filename As String, termsection As String) As String
' Imports the text files with native Read terms, Virtual Read terms for coding and
' alternate terms (variants of native or virtual terms which have identical meaning).
' Not all the native terms may be coded to; only those with include=TRUE.
' These term files are stored on the GitHub repository.
' Argument: termsection = native, virtual or alternate. They must be loaded
' in that order, because the medcodes must be in order.

Dim fileno As Integer
Dim rawstring As String
fileno = freefile
Dim thistermref  As Long ' to check that termref is sorted in ascending order.
Dim Include As Boolean ' whether the current term should be included in tables a and c
Dim includestr As String ' temporary string version of include

If Not in_set(termsection, "native", "virtual", "alternate") Then
    import = "ERROR: invalid termsection " & termsection
    Exit Function
End If

' When native terms are loaded, start from position zero.
If termsection = "native" Then
    thistermref = 0
    b_terms_used = 0
    a_terms_used = 0
Else
    thistermref = b_termref(b_terms_used)
End If

Open filename For Input As #fileno
' Inputs everything from file, default filenumber = 1
Dim b As Long: b = 1 + b_terms_used
Line Input #fileno, rawstring
If (termsection = "native" And rawstring = headerNative) Or _
    (termsection = "virtual" And rawstring = headerVirtual) Or _
    (termsection = "alternate" And rawstring = headerAlternate) Then
    Do
        Line Input #fileno, rawstring
        If rawstring <> "" Then
            b_termref(b) = CLng(dissect(rawstring, 1, vbTab))
            If b_termref(b) <= thistermref Then
                import = "ERROR: " & b_termref(b) & _
                    "is out of order; medcodes should be strictly increasing. "
                Close
                Exit Function
            End If
            thistermref = b_termref(b)

            ' Load the other columns
            Select Case termsection
            Case "native"
                b_std_term(b) = " " & Trim(dissect(rawstring, 4, vbTab)) & " "
                b_attrib_str(b) = dissect(rawstring, 5, vbTab)
                includestr = dissect(rawstring, 6, vbTab)
                If in_set(Left(includestr, 1), "T", "t", "Y", "y", "1") Then
                    Include = True
                Else
                    Include = False
                End If
                b_type(b) = dissect(rawstring, 7, vbTab)
                b_linkto(b) = b_termref(b)
            Case "virtual"
                b_std_term(b) = " " & Trim(dissect(rawstring, 3, vbTab)) & " "
                b_attrib_str(b) = dissect(rawstring, 4, vbTab)
                b_linkto(b) = b_termref(b)
            Case "alternate"
                b_std_term(b) = " " & Trim(dissect(rawstring, 2, vbTab)) & " "
                b_attrib_str(b) = dissect(rawstring, 3, vbTab)
                b_linkto(b) = CLng(dissect(rawstring, 4, vbTab))
            End Select
            
            ' If this term is included for matching to, add it to table a as well
            ' Note that table a is currently unsorted, but will be sorted by the
            ' init_and_sort subroutine afterwards.
            If Include = True Or termsection = "virtual" Or termsection = "alternate" Then
                a_terms_used = a_terms_used + 1
                a_std_term(a_terms_used) = b_std_term(b)
                a_termref(a_terms_used) = b_termref(b)
            End If
            b = b + 1
        End If
    Loop Until EOF(fileno)
    b_terms_used = b - 1
    import = "Loaded " & termsection & " terms from " & filename & _
        "; " & b_terms_used & " terms in total of which " & a_terms_used & _
        " can be matched to."
Else
    import = "ERROR: " & filename & " is not in the correct format."
End If
Close fileno

If termsection = "alternate" And Left(import, 6) <> "ERROR:" Then
    ' last section
    import = import & " Sorting lookup tables."
    Call init_and_sort
End If
End Function

Sub init_and_sort()
' Initialises c using table a, and sorts tables a and c. This must be
' run after tables a and b have been filled by the import function.
Dim i As Long
For i = 1 To a_terms_used
    c_bagofwords(i) = strfunc.bag_of_words(remove_ignorable(Trim(a_std_term(i)), _
        remove_right_left:=False))
    c_termref(i) = a_termref(i)
Next

' Now sort both tables (heap sort)
Dim iMin As Long
Dim iMax As Long
Dim bagswap As String
Dim termrefswap As Long
Dim stdswap As String

' Bag of words
iMin = 1
iMax = a_terms_used
For i = (iMax + iMin) \ 2 To iMin Step -1
    heap_bagofwords i, iMin, iMax
Next i
For i = iMax To iMin + 1 Step -1
    bagswap = c_bagofwords(i)
    c_bagofwords(i) = c_bagofwords(iMin)
    c_bagofwords(iMin) = bagswap
    termrefswap = c_termref(i)
    c_termref(i) = c_termref(iMin)
    c_termref(iMin) = termrefswap
    heap_bagofwords iMin, iMin, i - 1
Next i

' stdterm table
iMin = 1
iMax = a_terms_used
For i = (iMax + iMin) \ 2 To iMin Step -1
    heap_std_term i, iMin, iMax
Next i
For i = iMax To iMin + 1 Step -1
    stdswap = a_std_term(i)
    a_std_term(i) = a_std_term(iMin)
    a_std_term(iMin) = stdswap
    termrefswap = a_termref(i)
    a_termref(i) = a_termref(iMin)
    a_termref(iMin) = termrefswap
    heap_std_term iMin, iMin, i - 1
Next i

End Sub

Sub heap_bagofwords(ByVal i As Long, iMin As Long, iMax As Long)
' Heap helper function for sorting the bag of words vectors (table c).
Dim leaf As Long
Dim bagswap As String
Dim termrefswap As Long
Dim done As Boolean
done = False

Do
    leaf = i + i - (iMin - 1)
    If leaf > iMax Then
        done = True
    Else
        If leaf < iMax Then
            If c_bagofwords(leaf + 1) > c_bagofwords(leaf) Then
                leaf = leaf + 1
            End If
        End If
        If c_bagofwords(i) > c_bagofwords(leaf) Then
            done = True
        Else
            bagswap = c_bagofwords(i)
            c_bagofwords(i) = c_bagofwords(leaf)
            c_bagofwords(leaf) = bagswap
            termrefswap = c_termref(i)
            c_termref(i) = c_termref(leaf)
            c_termref(leaf) = termrefswap
            i = leaf
        End If
    End If
Loop Until done
End Sub

Sub heap_std_term(ByVal i As Long, iMin As Long, iMax As Long)
' Heap helper function for sorting by std_term (table a).
Dim leaf As Long
Dim stdswap As String
Dim termrefswap As Long
Dim done As Boolean
done = False

Do
    leaf = i + i - (iMin - 1)
    If leaf > iMax Then
        done = True
    Else
        If leaf < iMax Then
            If a_std_term(leaf + 1) > a_std_term(leaf) Then
                leaf = leaf + 1
            End If
        End If
        If a_std_term(i) > a_std_term(leaf) Then
            done = True
        Else
            stdswap = a_std_term(i)
            a_std_term(i) = a_std_term(leaf)
            a_std_term(leaf) = stdswap
            termrefswap = a_termref(i)
            a_termref(i) = a_termref(leaf)
            a_termref(leaf) = termrefswap
            i = leaf
        End If
    End If
Loop Until done
End Sub

Function pos_bagofwords(search_top As Boolean, instring As String) As Long
' Returns the position of first or last termref for which the bag of words
' (c_bagofwords) matches instring. If there is no match, zero is returned.
' Specify search_top = True to return the topmost match, or search_top = False
' for the last match.
Dim top As Long, bot As Long, trial As Long
top = 1: bot = a_terms_used
Do
    trial = Int((top + bot) / 2)
    If c_bagofwords(trial) > instring Then
        bot = trial - 1
    ElseIf c_bagofwords(trial) < instring Then
        top = trial + 1
    ElseIf c_bagofwords(trial) = instring Then
        If search_top = True Then bot = trial Else top = trial
    End If
Loop Until bot - top < 2

If top > bot Then Exit Function
For trial = top To bot
    If c_bagofwords(trial) = instring Then
        If search_top Then
            If c_bagofwords(trial - 1) <> instring Then
                pos_bagofwords = trial ' top of list of termrefs
            End If
        Else
            If c_bagofwords(trial + 1) <> instring Then
                pos_bagofwords = trial ' bottom of list of termrefs
            End If
        End If
    End If
Next
End Function

Function termref_bagofwords(position As Long) As Long
' Returns the termref from the bag of words table (c_termref) in a particular position.
If position <= a_terms_used Then
    termref_bagofwords = c_termref(position)
Else
    termref_bagofwords = 0
End If
End Function

Function true_term(Termref As Long) As Boolean
' Whether a term contains any true parts.
Dim attrib_string As String
attrib_string = attrib_str(Termref)
Dim a As Integer
For a = 1 To Len(attrib_string)
    If Mid(attrib_string, a, 1) = "F" Then true_term = False: Exit Function
    If Mid(attrib_string, a, 1) = "T" Then true_term = True: Exit Function
    If Mid(attrib_string, a, 1) = "O" Then true_term = True: Exit Function
Next
End Function

Function exact_read_termref(ByVal search_term As String) As Long
' Attempts to find an exact match to Read std_term, and returns the
' medcode (termref) of the match. Binary search of a_std_term.
Dim top As Long, bot As Long, trial As Long
If search_term = "" Then Exit Function
' Ensure there is one space before and after the word
search_term = " " & Trim(search_term) & " "
top = 1: bot = a_terms_used
Do
    trial = Int((top + bot) / 2)
    If a_std_term(trial) > search_term Then
        bot = trial - 1
    ElseIf a_std_term(trial) < search_term Then
        top = trial + 1
    ElseIf a_std_term(trial) = search_term Then
        top = trial: bot = trial
    End If
Loop Until bot - top < 2
If a_std_term(top) = search_term Then exact_read_termref = a_termref(top)
If a_std_term(bot) = search_term Then exact_read_termref = a_termref(bot)
End Function

Function read_type(Termref As Long) As String
' Returns the type code of the Read Term (whether pregnancy, death, labtest etc.)
' by binary search on table b.
Dim top As Long, bot As Long, trial As Long
top = 1: bot = b_terms_used
Do
    trial = Int((top + bot) / 2)
    If b_termref(trial) > Termref Then
        bot = trial - 1
    ElseIf b_termref(trial) < Termref Then
        top = trial + 1
    ElseIf b_termref(trial) = Termref Then
        top = trial: bot = trial
    End If
Loop Until bot - top < 2
If b_termref(top) = Termref Then read_type = b_type(top)
If b_termref(bot) = Termref Then read_type = b_type(bot)
End Function

Function std_term(Termref As Long) As String
' Returns the standardised term for a termref, by a binary search
' on table b.
Dim top As Long, bot As Long, trial As Long
top = 1: bot = b_terms_used
Do
    trial = Int((top + bot) / 2)
    If b_termref(trial) > Termref Then
        bot = trial - 1
    ElseIf b_termref(trial) < Termref Then
        top = trial + 1
    ElseIf b_termref(trial) = Termref Then
        top = trial: bot = trial
    End If
Loop Until bot - top < 2
If b_termref(top) = Termref Then std_term = b_std_term(top)
If b_termref(bot) = Termref Then std_term = b_std_term(bot)
End Function

Function attrib_str(Termref As Long) As String
' Returns the attribute string for a termref, by a binary search
' on table b.
Dim top As Long, bot As Long, trial As Long
top = 1: bot = b_terms_used
Do
    trial = Int((top + bot) / 2)
    If b_termref(trial) > Termref Then
        bot = trial - 1
    ElseIf b_termref(trial) < Termref Then
        top = trial + 1
    ElseIf b_termref(trial) = Termref Then
        top = trial: bot = trial
    End If
Loop Until bot - top < 2
If b_termref(top) = Termref Then attrib_str = b_attrib_str(top)
If b_termref(bot) = Termref Then attrib_str = b_attrib_str(bot)
End Function

Function linkto(Termref As Long) As String
' Returns the linked termref (e.g. for alternate Read terms) by binary
' search on table b.
Dim top As Long, bot As Long, trial As Long
top = 1: bot = b_terms_used
Do
    trial = Int((top + bot) / 2)
    If b_termref(trial) > Termref Then
        bot = trial - 1
    ElseIf b_termref(trial) < Termref Then
        top = trial + 1
    ElseIf b_termref(trial) = Termref Then
        top = trial: bot = trial
    End If
Loop Until bot - top < 2
If b_termref(top) = Termref Then linkto = b_linkto(top)
If b_termref(bot) = Termref Then linkto = b_linkto(bot)
End Function

