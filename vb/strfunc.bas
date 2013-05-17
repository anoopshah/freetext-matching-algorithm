Option Compare Binary
Option Explicit

' Module: strfunc -- functions for manipulating strings

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

Function get_date(s As String, Optional get_time As Boolean) As String
' Attempts to identify dates and durations in almost any format,
' returning a string stating the date or duration in a standardised format.
' The first 9 characters are the date/duration type, followed by a single
' space, followed by a number or date. The possible types are:
' DURA_gest (gestational age),
' DURA_days (duration in weeks),
' DURA_wks_ (duration in weeks),
' DURA_mths (duration in weeks),
' DURA_yrs_ (duration in years),
' DATE_time (time, e.g. 12:15),
' DATE_full (full date, e.g. 9-May-2013),
' DATE_year (year only)

Const mindate = #1/1/1910# ' earliest recognised date
Dim s1 As String, maybeduration As Boolean
Dim n1 As Long, n2string As String, n3 As Long
If s = "" Then Exit Function
If s = "1 a" Then Exit Function
If s = "12 am" Or s = "00.00" Or s = "00:00" Or _
    s = "12.00 am" Or s = "0000" Or s = "12:00 am" Then
    get_date = "DATE_time 00:00": Exit Function '
End If
If in_set(s, "24 h", "24 hr", "24 hour") Then get_date = "DURA_days 1"
If in_set(s, "48 h", "48 hr", "48 hour") Then get_date = "DURA_days 2"

If Left(s, 3) = "on " Then
    s = Mid(s, 4) ' CANNOT BE A DURATION
    maybeduration = False
ElseIf Left(s, 4) = "date" Then
    s = Mid(s, 6)
    maybeduration = False
Else
    maybeduration = True
End If

If IsNumeric(s) Then
    If (Len(s) <> 4 Or Val(s) > 2400) Then Exit Function ' added 13 jul 05
End If

Dim pattern As String, a As Integer
For a = 1 To 12
    s = replace(s, monthname(a, False), monthname(a, True))
Next
s = replace(s, "1 st", "1"): s = replace(s, "one ", "1 ")
s = replace(s, "2 nd", "2"): s = replace(s, "two ", "2 ")
s = replace(s, "3 rd", "3"): s = replace(s, "three ", "3 ")
s = replace(s, "4 th", "4")
s = replace(s, "5 th", "5")
s = replace(s, "6 th", "6")
s = replace(s, "7 th", "7")
s = replace(s, "8 th", "8")
s = replace(s, "9 th", "9")
s = replace(s, "0 th", "0")

For a = 1 To Len(s)
    If IsNumeric(Mid(s, a, 1)) Then
        pattern = pattern & "#"
    Else
        pattern = pattern & Mid(s, a, 1)
    End If
Next

For a = 1 To 12
    pattern = replace(pattern, monthname(a, True), "_m_")
Next

s1 = dissect2(s, , 1)

If maybeduration Then
    ' DURATION: MONTHS, WEEKS, DAYS
    If Left(pattern, 8) = "# to ## " Or Left(pattern, 8) = "# or ## " Then
        s1 = get_date_average(s1, Mid(s, 6, 2)): s = s1 & Mid(s, 8)
        pattern = Left("##", Len(s1)) & Mid(pattern, 8)
    ElseIf Left(pattern, 7) = "# to # " Or Left(pattern, 7) = "# or # " Then
        s1 = get_date_average(s1, Mid(s, 6, 1)): s = s1 & Mid(s, 7)
        pattern = Left("##", Len(s1)) & Mid(pattern, 7)
    ElseIf Left(pattern, 5) = "#-## " Then
        s1 = get_date_average(Left(s, 1), Mid(s, 3, 2)): s = s1 & Mid(s, 5)
        pattern = Left("##", Len(s1)) & Mid(pattern, 5)
    ElseIf Left(pattern, 4) = "#-# " Then
        s1 = get_date_average(Left(s, 1), Mid(s, 3, 1)): s = s1 & Mid(s, 4)
        pattern = Left("##", Len(s1)) & Mid(pattern, 4)
    End If
    
    Select Case pattern
    Case "#/#"
        If Right(s, 1) = "7" Then
            If Val(Left(s, 1)) >= 1 And Val(Left(s, 1)) <= 7 Then
                get_date = "DURA_days " & Left(s, 1): Exit Function
            End If
        End If
    Case "#/##"
        If Right(s, 2) = "52" Then ' weeks
            get_date = "DURA_wks_ " & Left(s, 1): Exit Function
        ElseIf Right(s, 2) = "40" Then ' gestation
            get_date = "DURA_wks_ " & Left(s, 1): Exit Function
        ElseIf Right(s, 2) = "12" Then
            get_date = "DURA_mths " & Left(s, 1): Exit Function
        End If
    Case "##/##"
        If Right(s, 2) = "52" And Val(Left(s, 2)) <= 52 Then ' weeks
            get_date = "DURA_wks_ " & Left(s, 2): Exit Function
        ElseIf Right(s, 2) = "40" And Val(Left(s, 2)) < 42 Then  ' gestation
            get_date = "DURA_gest " & Left(s, 2): Exit Function
        ElseIf Right(s, 2) = "12" And in_set(Left(s, 2), "10", "11") Then
            get_date = "DURA_mths " & Left(s, 2): Exit Function
        End If
    Case "##/## and #"
        If Mid(s, 4, 2) = "40" And Val(Left(s, 2)) < 42 And Val(Right(s, 1)) < 8 Then
            get_date = "DURA_gest " & Left(s, 2): Exit Function ' weeks gestation
        End If
    Case "##+/##"
        If Right(s, 2) = "40" And Val(Left(s, 2)) < 42 Then
            get_date = "DURA_gest " & Left(s, 2): Exit Function
        End If
    Case "## and #/##"
        If Right(s, 2) = "40" And Val(Left(s, 2)) < 42 And Val(Mid(s, 8, 1)) Then
            get_date = "DURA_gest " & Left(s, 2): Exit Function
        End If
    Case "# days", "# day", "# dy", "## days", "## day", "## dy", _
        "# d ago", "# d before", "# d history", "# d hx"
        If Val(s1) < 40 And Val(s1) > 0 Then
            get_date = "DURA_days " & s1: Exit Function
        End If
    Case "# weeks", "# wks", "# wk", "# week", "## weeks", "## wks", "## wk", "# week", _
        "# and weeks", "# and wks", "# and wk", "# and week", "## and weeks", "## and wks", _
        "## and wk", "# and week", "# w ago", "# w before", "# w history", "# w hx"
        ' translated from 30+ weeks etc.
        If Val(s1) < 45 And Val(s1) > 0 Then
            get_date = "DURA_wks_ " & s1: Exit Function
        End If
    Case "# months", "# mths", "# mth", "# month", "## months", "## mths", "## mth", "# month"
        If Val(s1) < 40 And Val(s1) > 0 Then
            get_date = "DURA_mths " & s1: Exit Function
        End If
    Case "# years", "# year", "# yrs", "# yr", "## years", "## year", "## yrs", "# yr"
        If Val(s1) < 80 And Val(s1) > 0 Then
            get_date = "DURA_yrs_ " & s1: Exit Function
        End If
    End Select
End If

Dim temp As Date
On Error Resume Next
Select Case pattern
Case "in ####"  ' if the date is a year
    If Val(Right(s, 4)) > 1900 And Val(Right(s, 4)) <= Year(now) Then
        get_date = "DATE_year " & Right(s, 4): Exit Function
    End If
Case "at ####"
    If get_time Then
        temp = CDate(Right(s, 4))
        If Hour(temp) > 0 Then get_date = "DATE_time " & Format(temp, "hh:mm")
        Exit Function
    End If
Case "####"
    If Val(s) > 1945 And Val(s) <= Year(now) + 1 Then ' maximum year is now+1
        get_date = "DATE_year " & Right(s, 4): Exit Function ' year
    ElseIf Val(s) < 2400 And get_time Then
        temp = CDate(Left(s, 2) & ":" & Right(s, 2)) ' time
        If Hour(temp) > 0 Then get_date = "DATE_time " & Format(temp, "hh:mm")
        Exit Function
    End If
Case "#/#/##", "#/##/##", "##/#/##", "##/##/##", "#/#/####", "#/##/####", "##/#/####", "##/##/####"
    n1 = dissect2(s, "/", 1): n2string = monthname(dissect2(s, "/", 2), True): n3 = dissect2(s, "/", 3)
Case "#.#.##", "#.##.##", "##.#.##", "##.##.##", "#.#.####", "#.##.####", "##.#.####", "##.##.####"
    n1 = dissect2(s, ".", 1): n2string = monthname(dissect2(s, ".", 2), True): n3 = dissect2(s, ".", 3)
Case "#;#;##", "#;##;##", "##;#;##", "##;##;##", "#;#;####", "#;##;####", "##;#;####", "##;##;####"
    n1 = dissect2(s, ";", 1): n2string = monthname(dissect2(s, ";", 2), True): n3 = dissect2(s, ";", 3)
Case "#-#-##", "#-##-##", "##-#-##", "##-##-##", "#-#-####", "#-##-####", "##-#-####", "##-##-####"
    n1 = dissect2(s, "-", 1): n2string = monthname(dissect2(s, "-", 2), True): n3 = dissect2(s, "-", 3)
Case "# _m_ ##", "## _m_ ##", "# _m_ ####", "## _m_ ####"
    n1 = dissect2(s, " ", 1): n2string = dissect2(s, " ", 2): n3 = dissect2(s, " ", 3)
Case "#/_m_/##", "##/_m_/##", "#/_m_/####", "##/_m_/####"
    n1 = dissect2(s, "/", 1): n2string = dissect2(s, "/", 2): n3 = dissect2(s, "/", 3)
Case "#-_m_-##", "##-_m_-##", "#-_m_-####", "##-_m_-####"
    n1 = dissect2(s, "-", 1): n2string = dissect2(s, "-", 2): n3 = dissect2(s, "-", 3)
Case "_m_ #,##", "_m_ ##,##", "_m_ #,####", "_m_ ##,####"
    n1 = dissect2(dissect2(s, " ", 2), ",", 1): n2string = dissect2(s, " ", 1): n3 = dissect2(s, ",", 2)
Case "# _m_,##", "## _m_,##", "# _m_,####", "## _m_,####"
    n2string = dissect2(dissect2(s, " ", 2), ",", 1): n1 = dissect2(s, " ", 1): n3 = dissect2(s, ",", 2)
End Select
temp = CDate(Format(n1) & " " & n2string & " " & Format(n3))

If (temp > mindate And temp < now + 400) Then
    get_date = "DATE_full " & Format(temp, "d-mmm-yyyy")
    Exit Function
End If

' TIMES
'/ am, pm
If Not get_time Then Exit Function
Select Case pattern
Case "#.##", "#:##"
    If Left(s, 1) > "0" And Val(Mid(s, 3, 1)) <= 5 Then temp = CDate(s)
Case "##.##", "##:##"
    If Val(Mid(s, 4, 1)) <= 5 Then temp = CDate(s)
Case "# am", "# pm"
    temp = CDate(Left(s, 1) & ":00 " & Right(s, 2))
Case "## am", "## pm"
    temp = CDate(Left(s, 2) & ":00 " & Right(s, 2))
Case "#### hrs", "####"
    If Val(s) < 2400 Then temp = CDate(Left(s, 2) & ":" & Mid(s, 3, 2))
Case "#.## hrs"
    temp = CDate(Left(s, 1) & ":" & Mid(s, 3, 2))
Case "##.## hrs", "##:## hrs"
    temp = CDate(Left(s, 2) & ":" & Mid(s, 4, 2))
End Select
If temp > 0 Then get_date = "DATE_time " & Format(temp, "hh:mm")
End Function

Function get_date_average(s1 As String, s2 As String) As String
' Provides a replacement for the first number (s1) from phrases such
' as 2-3 weeks, 5-6 days etc. The average duration is used, rounded up
' (no fractions in the result).
Dim s1_val As Single, s2_val As Single
s1_val = Val(s1): s2_val = Val(s2)
If (s1_val >= s2_val Or s2_val > 2 * s1_val) Then get_date_average = "ZZZ": Exit Function
get_date_average = CInt((s1_val + s2_val) / 2 + 0.5)
End Function

Function words(ByVal phrase As String, start As Long, Optional numwd As Long, _
    Optional finish As Long) As String
' Extracts individual words from a string, assuming one space between words
' and no spaces at the beginning of the string.
On Error GoTo help:
If numwd = 0 Then
    If finish = 0 Then
        words = dissect2(Trim(phrase), , start)
        Exit Function
    Else
        numwd = finish + 1 - start
    End If
End If
If start > numwords(phrase) Then Exit Function

Dim a As Long, b As Long, c As Long
phrase = Trim(phrase) & " "
b = 0: a = 0
Do While b < start - 1
    a = InStr(a + 1, phrase, " ")
    b = b + 1
Loop
a = a + 1 ' start position
c = a

' now find finish position (c is the position of space just after the last desired word
For b = 1 To numwd
    c = InStr(c + 1, phrase, " ")
    If c = 0 Then Exit For
Next

If c < 2 Then words = Mid(phrase, a) Else words = Mid(phrase, a, c - a)
help:
End Function

Function in_set(Target As String, a As String, b As String, Optional c As String, _
    Optional d As String, Optional e As String, Optional f As String, _
    Optional g As String, Optional h As String, Optional i As String, _
    Optional j As String, Optional k As String, Optional l As String, _
    Optional m As String, Optional n As String, Optional o As String) As Boolean
' Whether target is one of a, b, c, d, e etc. The function does not consider
' any entries after the first empty string.
in_set = False
If a = "" And Target = "" Then in_set = True: Exit Function
If Target = a Then in_set = True: Exit Function
If Target = b Then in_set = True: Exit Function
If Not c = "" Then
    If Target = c Then in_set = True: Exit Function
Else: Exit Function
End If
If Not d = "" Then
    If Target = d Then in_set = True: Exit Function
Else: Exit Function
End If
If Not e = "" Then
    If Target = e Then in_set = True: Exit Function
Else: Exit Function
End If
If Not f = "" Then
    If Target = f Then in_set = True: Exit Function
Else: Exit Function
End If
If Not g = "" Then
    If Target = g Then in_set = True: Exit Function
Else: Exit Function
End If
If Not h = "" Then
    If Target = h Then in_set = True: Exit Function
Else: Exit Function
End If
If Not i = "" Then
    If Target = i Then in_set = True: Exit Function
Else: Exit Function
End If
If Not j = "" Then
    If Target = j Then in_set = True: Exit Function
Else: Exit Function
End If
If Not k = "" Then
    If Target = k Then in_set = True: Exit Function
Else: Exit Function
End If
If Not l = "" Then
    If Target = l Then in_set = True: Exit Function
Else: Exit Function
End If
If Not m = "" Then
    If Target = l Then in_set = True: Exit Function
Else: Exit Function
End If
If Not n = "" Then
    If Target = l Then in_set = True: Exit Function
Else: Exit Function
End If
If Not o = "" Then
    If Target = l Then in_set = True: Exit Function
Else: Exit Function
End If
End Function


Function is_text(instring As String) As Boolean
' Whether a string consists entirely of lower case text.
Dim counter As Long
Dim result As Boolean

result = True

For counter = 1 To Len(instring)
    If Asc(Mid(instring, counter, 1)) > 122 Then result = False
    If Asc(Mid(instring, counter, 1)) < 97 Then result = False
Next

is_text = result
End Function

Function numwords(ByVal instring As String) As Long
' Returns the number of words in a string, assuming exactly one
' space between adjacent words.
Dim pos As Long ' assumes only one space between words
instring = Trim(instring)
If instring = "" Then numwords = 0: Exit Function
pos = 0: numwords = 0
Do
    pos = InStr(pos + 1, instring, " "): numwords = numwords + 1
Loop Until pos = 0
End Function

Function num_diff_char(str1 As String, str2 As String) As Long
' Counts the number of characters which are different between str1 and str2.
' Ignores any differences beyond the length of the shorter string.
' If there are more than 3 differences, num_diff_char returns '4' and the exact
' number of differences is not counted.
num_diff_char = 0
Dim length As Long
length = Len(str1)
If Len(str2) < length Then length = Len(str2)
If length = 0 Then num_diff_char = -1: Exit Function  ' one of the strings is empty
Do
    If Mid(str1, length, 1) <> Mid(str2, length, 1) Then num_diff_char = num_diff_char + 1
    length = length - 1
    If num_diff_char > 3 Then Exit Function
Loop Until length = 0
End Function

Function dissect(in_string As String, number As Long, _
    Optional delimiter As String) As String
' Extracts part of a string between two delimiters. Uses the
' VBA.split function via 'dissect2'. The functions dissect and dissect2 are
' identical apart from the order of the arguments.
dissect = dissect2(in_string, delimiter, number)
End Function

Function dissect2(in_string As String, Optional delimiter As String, _
    Optional number As Long) As String
' Extracts part of a string between two delimiters. Uses the
' VBA.split function via 'dissect2', with a fallback to the dissect3 function
' in the strings_Acc97 module if this function is not found. The functions
' dissect and dissect2 are identical apart from the order of the arguments.
If number = 0 Then number = 1
If delimiter = "" Then delimiter = " "
On Error GoTo endsub
dissect2 = VBA.split(in_string, delimiter)(number - 1)
Exit Function
endsub:
dissect2 = strings_Acc97.dissect3(in_string, delimiter, number)
End Function

Function is_numeric(instring As String, Optional lab_results_mode As Boolean, _
    Optional dont_ignore_large_numbers As Boolean) As Boolean
' Determines whether a string contains only a single number or part of a
' single number. If lab_results_mode is TRUE, words like 'normal', 'abnormal' etc.
' are considered to be numbers.
Dim points As Long ' number of decimal points
Dim counter As Long
Dim result As Boolean

If lab_results_mode Then
    If in_set(instring, "normal", "abnormal", "nil", "negative", "positive", "nad", "neg") Then
        is_numeric = True: Exit Function
    End If
End If

If instring = "." Then is_numeric = False: Exit Function
If instring = "" Then
    is_numeric = False
    Exit Function
End If

If Len(instring) = 1 Then
    If Asc(instring) > 47 And Asc(instring) < 58 Then is_numeric = True
    Exit Function
End If

result = True
points = 0

For counter = 1 To Len(instring)
    If Asc(Mid(instring, counter, 1)) > 57 Then result = False
    If Asc(Mid(instring, counter, 1)) < 48 Then result = False
    If Mid(instring, counter, 1) = "." Then
        points = points + 1
        result = True
    End If
    If result = False Then Exit For
Next
    
If points > 1 Then result = False
If Val(instring) > 20000 And dont_ignore_large_numbers = False Then result = False
    ' larger numbers are almost certainly ID numbers or dates without delimiters.
is_numeric = result

End Function

Function bag_of_words(instring As String) As String
' Creates a bag-of-words representation of a string: all words in
' alphabetical order, no duplicates, one space between words.
' This function can only handle up to 10 words; any additional words
' are ignored.

Const maxwords = 10
Dim thewords(maxwords) As String
Dim i As Long, j As Long
Dim numwords_ As Long
numwords_ = numwords(instring)
If numwords_ > maxwords Then
    numwords_ = maxwords
End If

' Get separate words
For i = 1 To numwords_
    If (i <= maxwords) Then
        thewords(i) = words(instring, i)
    End If
Next

' Sort the words
wordlist.quicksort thewords, 1, numwords_

' Remove duplicates
i = 2
Do While i <= numwords_
    If thewords(i) = thewords(i - 1) Then
        ' This is a duplicate, so move all subsequent words
        ' one row forward
        If i < numwords_ Then
            For j = i To numwords_ - 1
                thewords(j) = thewords(j + 1)
            Next
        End If
        numwords_ = numwords_ - 1
    Else
        i = i + 1
    End If
Loop

' Reconstitute the words with a single space
Dim output As String
output = thewords(1)
If numwords_ > 1 Then
    For i = 2 To numwords_
        output = output & " " & thewords(i)
    Next
End If
bag_of_words = output
End Function

