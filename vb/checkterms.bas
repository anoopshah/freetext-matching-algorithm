Option Compare Binary
Option Explicit

' Module: checkterms -- checks for occurence (or not) of words in the
' text to validate or invalidate some termrefs (medcodes)

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

Const maxcheckterms = 100

Dim Termref(maxcheckterms) As Long ' medcode of output term
Dim Qualify(maxcheckterms) As String ' word fragments which must be present in the text for the medcode to be returned
Dim Dequalify(maxcheckterms) As String ' word fragments which must not be present in the text for the medcode to be returned' verbatim words or phrases
Dim used As Long ' number of entries

Function import(filename As String) As String
' Imports the checkterms table from text file. Returns a text statement stating
' whether it was successful. This text can be displayed on screen or added to a log file.

Dim fileno As Integer
Dim rawstring As String
fileno = freefile
Dim thistermref  As Long ' to check that termref is sorted in ascending order.
thistermref = 0

Open filename For Input As #fileno
' Inputs everything from file, default filenumber = 1
Dim b As Long: b = 1
Line Input #fileno, rawstring
If rawstring = "medcode" & vbTab & "qualify" & vbTab & "dequalify" Then
    Do
        Line Input #fileno, rawstring
        If rawstring <> "" Then
            Termref(b) = CLng(dissect(rawstring, 1, vbTab))
            If Termref(b) <= thistermref Then
                import = "ERROR: " & Termref(b) & _
                    "is out of order; medcodes should be strictly increasing. "
                Exit Do
            End If
            thistermref = Termref(b)
            Qualify(b) = dissect(rawstring, 2, vbTab)
            Dequalify(b) = dissect(rawstring, 3, vbTab)
            b = b + 1
        End If
    Loop Until EOF(fileno)
    used = b - 1
    import = "Loaded checkterms (" & used & " rows) from " & filename
Else
    import = "ERROR: " & filename & " is not in the correct format."
End If
Close fileno

End Function

Sub check_all(checkstring As String, Optional debug_ As Boolean, Optional sicknote As Boolean, _
    Optional death As Boolean, Optional date_only As Boolean)
' Carries out the actual checking
Dim a As Long, b As Long, tmptermref As Long, c As Long
If pd.max = 0 Then Exit Sub
Dim qual As Boolean, dequal As Boolean, needqual As Boolean
Dim countdates As Integer: countdates = 0
Dim curdate As String: curdate = ""

a = 1
Do
    If sicknote Then
        If Left(pd.mean(a), 1) = "D" Then pd.set_attr "sicknote", a ' date or duration
    End If
    
    If death Then
        If Left(pd.mean(a), 9) = "DATE_full" And pd.Attr(a) = "" And (curdate <> pd.mean(a)) Then
            countdates = countdates + 1
            curdate = pd.mean(a) ' two identical dates are acceptable
            pd.set_attr "date", a
        End If
        If Left(pd.mean(a), 4) = "LABS" Then
            If debug_ Then debug_string = debug_string & _
                pd.mean(a) & "no LABS data allowed in Death mode" & Chr$(13) & Chr$(10)
            pd.remove a: a = a - 1  ' remove all LABS data
        End If
    ElseIf date_only Then
        If Left(pd.mean(a), 9) = "DATE_full" And pd.Attr(a) = "" And (curdate <> pd.mean(a)) Then
            countdates = countdates + 1
            curdate = pd.mean(a) ' two identical dates are acceptable
            pd.set_attr "date", a
        Else
            pd.remove a: a = a - 1 ' remove all non-date data
        End If
    End If
       
    If pd.Attr(a) = "normalrange" Then
        If debug_ Then debug_string = debug_string & _
            pd.mean(a) & " removed because refers to target or normal range" & Chr$(13) & Chr$(10)
        pd.remove a: a = a - 1
    End If
    
    If Left(pd.mean(a), 4) = "READ" Then
        tmptermref = Val(dissect2(pd.mean(a), " ", 2))
        b = in_list(tmptermref)
        If b > 0 Then
            If Qualify(b) = "" Then needqual = False Else needqual = True
            qual = if_qualify(b, checkstring)
            dequal = if_dequalify(b, checkstring)
            If qual = True And dequal = True Then
                pd.set_attr "machinequery", a
            ElseIf qual = False And dequal = True Then
                If debug_ Then debug_string = debug_string & _
                    "Term " & tmptermref & " removed because of ambiguity" & Chr$(13) & Chr$(10)
                pd.remove a: a = a - 1 ' remove this entry
            ElseIf needqual = True And qual = False Then
                If debug_ Then debug_string = debug_string & _
                    "Term " & tmptermref & " removed because of ambiguity" & Chr$(13) & Chr$(10)
                pd.remove a: a = a - 1 ' remove this entry
            End If
        End If
    End If
     
    a = a + 1
Loop Until a > pd.max

If countdates > 1 Then ' ambiguous date - convert all dates to 'machinequery'
    ' Note that machinequery dates are not returned by fma_gold as they do not have an
    ' entity type; they are simply ignored.
    If debug_ Then debug_string = debug_string & _
        "More than one date detected -- to be removed" & Chr$(13) & Chr$(10)
    For a = 1 To pd.max
        If Left(pd.mean(a), 9) = "DATE_full" Then pd.set_attr "machinequery", a
    Next
End If
End Sub

Function in_list(in_termref As Long) As Long
' Returns the row number of the termref (medcode) in the checkterms table
If in_termref = 0 Then Exit Function
Dim top As Long, bot As Long, trial As Long
top = 1: bot = used: in_list = 0
Do
    trial = Int((top + bot) / 2)
    If Termref(trial) > in_termref Then
        bot = trial - 1
    ElseIf Termref(trial) < in_termref Then
        top = trial + 1
    ElseIf Termref(trial) = in_termref Then
        top = trial: bot = trial
    End If
Loop Until bot - top < 2
If Termref(top) = in_termref Then in_list = top
If Termref(bot) = in_termref Then in_list = bot
End Function

Function if_qualify(pos As Long, checkstring As String) As Boolean
' Returns TRUE if one of the qualifying phrases is present in the text,
' FALSE otherwise.
Dim a As Long, src As String
a = 1: if_qualify = False
Do
    src = dissect2(Qualify(pos), ",", a)
    If src = "" Then Exit Function
    If InStr(1, checkstring, src) Then if_qualify = True: Exit Function
    a = a + 1
Loop
End Function

Function if_dequalify(pos As Long, checkstring As String) As Boolean
' whether one of the dequalifying terms is present in the text
Dim a As Long, src As String
a = 1: if_dequalify = False
Do
    src = dissect2(Dequalify(pos), ",", a)
    If src = "" Then Exit Function
    If InStr(1, checkstring, src) Then if_dequalify = True: Exit Function
    a = a + 1
Loop
End Function

