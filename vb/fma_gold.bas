Option Compare Binary
Option Explicit

' Module: fma_gold -- functions for analysis of free text and output in
' a format similar to the Clinical Practice Research Datalink 'GOLD' format.

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

' Maximum number of texts to be analysed
Const maxtexts = 200001
Const delim = "," ' delimiter
Dim newline As String ' newline character, will be defined in main_fma_gold

' Maximum words per text, maximum output rows per text (should be the same as max pd)
Const maxrows = 1000
Dim outdata(maxrows) As String ' output data staging area
Dim outrows As Long ' number of rows in output for a single text

' Arrays for storage of practice, textid and medcode (for analysing text in
' context of associated Read code)
Dim pracid(maxtexts) As Long ' ordered practice identifier
Dim textid(maxtexts) As Long ' ordered text ID (unique within practice)
Dim medcode(maxtexts) As Long ' medcode (may be multiple medcodes for each pracid / textid combination)
Dim ntexts As Long ' actual number of texts

Sub do_analysis(logfile As String, lookups As String, _
    Optional infile As String, Optional medcodefile As String, _
    Optional outfile As String, Optional freetext As String, _
    Optional medcode As String, Optional origmedcode As Long)
' FMA gold analysis of free text. medcodefile is the file with medcodes
' to be appended to the free text to provide the analysis modes.
' This file is optional. If not provided, medcode is assumed to be
' zero for all files. If freetext is supplied as an argument
' to the function, it is analysed (together with origmedcode)
' and the text and debug output are written in the log file.

' Check that folder names end in '\'
If Not Right(lookups, 1) = "\" Then lookups = lookups & "\"
If medcode <> "" Then origmedcode = CLng(medcode)

newline = Chr$(13) & Chr$(10) ' Windows-style new line

Dim logfileno As Integer
Dim numtexts As Long
Dim numwd As Long

logfileno = freefile
numtexts = 0
numwd = 0

Open logfile For Output As #logfileno
Print #logfileno, "Freetext Matching Algorithm version 15."
Print #logfileno, "Analysis started at " & Time & " on " & date & newline

' Initialise lookups
Dim lookupfile As Integer
lookupfile = freefile

Print #logfileno, "Loading lookup tables from " & lookups
Print #logfileno, freetext_core.import_all_lookups(lookups) & newline

If freetext = "" Then
    ' Load from file
    
    ' Load medcodes
    Print #logfileno, loadMedcodes(medcodefile)
    
    Dim infileno As Integer
    Dim outfileno As Integer
    infileno = freefile
    
    ' Open input file
    Open infile For Input As #infileno
    Print #logfileno, "Opening input file " & infile
    
    ' Open output file
    outfileno = freefile
    Open outfile For Output As #outfileno
    Print #logfileno, "Opening output file " & outfile
    Print #outfileno, "pracid,textid,origmedcode,medcode,enttype,data1,data2,data3,data4"
    
    Dim origmedcodes As String
    Dim nummedcodes As Long
    Dim thistextid As Long
    Dim thispracid As Long
    Dim rawtext As String
    Dim i As Long ' counter for medcodes
    Dim j As Long ' counter for outdata
    
    ' Loop through the input file, interpreting each text
    Do While Not EOF(infileno)
        ' Read the next line of text
        Line Input #infileno, rawtext
        ' textid and pracid must be valid
        thistextid = gettextid(rawtext)
        thispracid = getpracid(rawtext)
        freetext = dissect(rawtext, 3, Chr$(9))
        origmedcodes = getmedcodes(targetpracid:=thispracid, _
            targettextid:=thistextid)
        nummedcodes = CLng(dissect(origmedcodes, 1, "|"))
        If nummedcodes = 0 Then
            origmedcodes = "1|0"
            nummedcodes = 1
        End If
        
        For i = 1 To nummedcodes
            ' update statistics (each text with different medcode counts as separate)
            numwd = numwd + numwords(freetext)
            numtexts = numtexts + 1
            
            origmedcode = CLng(dissect(origmedcodes, 1 + i, "|"))
            If origmedcode = 0 Then
                freetext_core.main_analyse instring:=freetext, _
                    spell_:=True, debug_:=False
            Else
                freetext_core.main_termref instring:=freetext, _
                    Termref:=origmedcode, spell_:=True, debug_:=False, append_term:=True
            End If
            ' Convert to FMA GOLD format, and store the output in outdata
            pd_to_fma_gold origmedcode
            
            If outrows > 0 Then
                For j = 1 To outrows
                    Print #outfileno, thispracid & delim & thistextid & delim & _
                        origmedcode & delim & outdata(j)
                Next
            End If
        Next
    Loop
    
    Print #logfileno, newline & "Finished analysing " & numtexts & _
        " texts (approx " & numwd & " words) at " & Time
    
    ' Close all files
    Close #infileno
    Close #outfileno

Else
    ' Analyse a single text, and put the debug output in the logfile
    Print #logfileno, newline & "Analysing a single text:"
    If origmedcode = 0 Then
        freetext_core.main_analyse instring:=freetext, _
            spell_:=True, debug_:=True
        Print #logfileno, freetext & newline
    Else
        freetext_core.main_termref instring:=freetext, _
             Termref:=origmedcode, spell_:=True, debug_:=True, append_term:=True
        Print #logfileno, "Original Read term:" & terms.std_term(origmedcode) & newline & _
            "Free text: " & freetext & newline
    End If
    Print #logfileno, debug_string
    pd_to_fma_gold
    Print #logfileno, newline & "Output:"
    ' Show the Read term (std_term) in the output as well
    Print #logfileno, "pracid,textid,medcode,enttype,data1,data2,data3,data4,std_term"
    For j = 1 To outrows
        Print #logfileno, thispracid & delim & thistextid & delim & _
            delim & outdata(j) & delim & terms.std_term(CLng(dissect(outdata(j), 1, delim)))
    Next
End If
Close #logfileno
End Sub

Function gettextid(str As String) As Long
' Finds textid in a string, at the second position, tab separated
' In a separate function for error trapping purposes
gettextid = 0
On Error GoTo errortrapping:
gettextid = CLng(dissect(str, 2, Chr$(9)))
errortrapping:
End Function

Function getpracid(str As String) As Long
' Finds pracid in a string, at the first position, tab separated
' In a separate function for error trapping purposes
getpracid = 0
On Error GoTo errortrapping:
getpracid = CLng(dissect(str, 1, Chr$(9)))
errortrapping:
End Function

Function pdYYYYMMDD(str As String) As Double
' Converts a date to YYYYMMDD format

' In case of a type conversion error:
On Error GoTo errortrapping:

Dim typeofdate As String
Dim thevalue As String
Dim thedate As Date
typeofdate = dissect(str, 1)
thevalue = dissect(str, 2)

If (typeofdate = "DATE_year") Then
    pdYYYYMMDD = CDbl(thevalue)
ElseIf (typeofdate = "DATE_full") Then
    thedate = CDate(thevalue)
    pdYYYYMMDD = 10000 * Year(thedate) + 100 * Month(thedate) + Day(thedate)
End If
errortrapping:
End Function

Function pdValue(str As String) As Double
' Returns the value (e.g. medcode, LABS value, duration number)

' In case of a type conversion error:
On Error GoTo errortrapping:

Dim thevalue As String
thevalue = dissect(str, 2)
If is_numeric(thevalue, dont_ignore_large_numbers:=True) Then
    pdValue = CDbl(dissect(str, 2))
Else
    ' Convert to TQU lookup
    Select Case thevalue
    Case "nil"
        pdValue = 15
    Case "nad", "normal"
        pdValue = 9
    Case "abnormal"
        pdValue = 12
    Case "neg", "negative"
        pdValue = 22
    Case "positive"
        pdValue = 21
    End Select
End If
errortrapping:
End Function

Function pdAge(str As String) As Double
' Returns the age in years
Dim thevalue As Double

' In case of a type conversion error:
On Error GoTo errortrapping:

thevalue = CDbl(dissect(str, 2))
Select Case dissect(str, 1)
Case "DURA_yrs_"
    pdAge = thevalue
Case "DURA_wks_"
    pdAge = thevalue / 52
Case "DURA_mths"
    pdAge = thevalue / 12
Case "DURA_days"
    pdAge = thevalue / 365.25
End Select
errortrapping:
End Function

Function pdDurUnits(str As String) As Double
' Returns the SUM lookup value for the duration units
Select Case dissect(str, 1)
Case "DURA_yrs_"
    pdDurUnits = 148
Case "DURA_wks_"
    pdDurUnits = 147
Case "DURA_mths"
    pdDurUnits = 101
Case "DURA_days"
    pdDurUnits = 41
End Select
End Function

Function pdDurValue(str As String) As Double
' Returns the SUM lookup value for the duration units
' Error trapping in case of type conversion error
On Error GoTo errortrapping:
If Left(str, 4) = "DURA" Then
    pdDurValue = CDbl(dissect(str, 2))
End If
errortrapping:
End Function

Sub addOutputRow(medcode_ As Double, enttype_ As Double, _
    Optional data1 As Double, Optional data2 As Double, _
    Optional data3 As Double, Optional data4 As Double)
' Adds data to the output rows. All arguments are required to be double.
' Zero values are ignored and considered as missing.
If (outrows < maxrows) Then
    outrows = outrows + 1
    outdata(outrows) = medcode_ & delim & enttype_ & delim & _
        blankIfZero(data1) & delim & blankIfZero(data2) & delim & _
        blankIfZero(data3) & delim & blankIfZero(data4)
End If
End Sub

Function blankIfZero(number As Double) As String
' Converts a number to a string, returning an empty string if the number
' is zero.
If (number = 0) Then
    blankIfZero = ""
Else
    blankIfZero = CStr(number)
End If
End Function

Sub pd_to_fma_gold(Optional origmedcode As Long)
' Extracts information from pd and converts it to FMA gold format.
' Stores the extracted information in the outdata array.

Dim i As Long ' looper for pd
Dim currentoutput As Long ' current position in output array
i = 1
outrows = 0 ' reset the output rows
If pd.max > 0 Then
    Do While i <= pd.max
        ' Try two rows of pd, if possible
        currentoutput = outrows
        If i < pd.max Then
            If pd.Attr(i) = "sysbp" And pd.Attr(i + 1) = "diabp" Then
                addOutputRow 1, 1001, pdValue(pd.mean(i + 1)), pdValue(pd.mean(i))
            ElseIf pd.Attr(i) = "admitdate" And pd.Attr(i + 1) = "dischdate" Then
                If (pdValue(pd.mean(i)) <= pdValue(pd.mean(i + 1))) Then
                    ' check that admission and discharge dates are in the right order
                    addOutputRow 43828, 2009, pdYYYYMMDD(pd.mean(i)), pdYYYYMMDD(pd.mean(i + 1))
                End If
            ElseIf (in_set(pd.Attr(i), "dueto", "causing") Or pd.Attr(i) = "") And _
                dissect(pd.mean(i), 1) = "READ" Then
                ' Diagnosis followed by date / duration / age (entity 2005)
                Select Case pd.Attr(i + 1)
                Case "duraprev"
                    addOutputRow pdValue(pd.mean(i)), 2005, _
                        data2:=pdDurValue(pd.mean(i + 1)), _
                        data3:=pdDurUnits(pd.mean(i + 1))
                Case "dateprev"
                    addOutputRow pdValue(pd.mean(i)), 2005, _
                        data1:=pdYYYYMMDD(pd.mean(i + 1))
                Case "ageprev"
                    ' Convert age to years if it is in weeks or months
                    addOutputRow pdValue(pd.mean(i)), 2005, _
                        data4:=pdAge(pd.mean(i + 1))
                End Select
            ElseIf (in_set(pd.Attr(i + 1), "dueto", "causing") Or pd.Attr(i) = "") And _
                dissect(pd.mean(i + 1), 1) = "READ" Then
                ' Diagnosis preceded by date / duration / age
                Select Case pd.Attr(i)
                Case "duranext"
                    addOutputRow pdValue(pd.mean(i + 1)), 2005, _
                        data2:=pdDurValue(pd.mean(i)), _
                        data3:=pdDurUnits(pd.mean(i))
                Case "datenext"
                    addOutputRow pdValue(pd.mean(i + 1)), 2005, _
                        data1:=pdYYYYMMDD(pd.mean(i))
                End Select
            ElseIf pd.Attr(i) = "pmh" Then
                ' Diagnosis followed by date / duration / age (entity 1002)
                Select Case pd.Attr(i + 1)
                Case "duraprev"
                    addOutputRow pdValue(pd.mean(i)), 1002, _
                        data2:=pdDurValue(pd.mean(i + 1)), _
                        data3:=pdDurUnits(pd.mean(i + 1))
                Case "dateprev"
                    addOutputRow pdValue(pd.mean(i)), 1002, _
                        data1:=pdYYYYMMDD(pd.mean(i + 1))
                Case "ageprev"
                    addOutputRow pdValue(pd.mean(i)), 1002, _
                        data4:=pdAge(pd.mean(i + 1))
                End Select
            ElseIf pd.Attr(i + 1) = "pmh" Then
                ' Diagnosis preceded by date / duration / age
                Select Case pd.Attr(i)
                Case "duranext"
                    addOutputRow pdValue(pd.mean(i + 1)), 1002, _
                        data2:=pdDurValue(pd.mean(i)), _
                        data3:=pdDurUnits(pd.mean(i))
                Case "datenext"
                    addOutputRow pdValue(pd.mean(i + 1)), 1002, _
                        data1:=pdYYYYMMDD(pd.mean(i))
                End Select
            End If
        End If
        If outrows > currentoutput Then
            ' Two rows have been used; advance counter and don't try any single row pd
            ' just yet in case it is possible to do another two rows
            i = i + 1
        Else
            ' Try a single row, all entities in entity spreadsheet
            Select Case pd.Attr(i)
            Case "", "dueto", "causing"
                ' Rows with a blank attribute only if they are Read or LABS
                If (dissect(pd.mean(i), 1) = "READ") Then
                    addOutputRow pdValue(pd.mean(i)), 2005
                ElseIf (dissect(pd.mean(i), 1) = "LABS") Then
                    If is_numeric(dissect(pd.mean(i), 2), _
                        dont_ignore_large_numbers:=True) Then
                        ' Quantitative result
                        addOutputRow CDbl(origmedcode), 2011, 3, pdValue(pd.mean(i))
                    Else
                        ' Qualitative: nil, normal, abnormal, negative or positive
                        addOutputRow CDbl(origmedcode), 2011, pdValue(pd.mean(i))
                    End If
                End If
            Case "pmh": addOutputRow pdValue(pd.mean(i)), 1002
            Case "negative": addOutputRow 72907, 1085, pdValue(pd.mean(i))
            Case "family": addOutputRow 17485, 1087, pdValue(pd.mean(i))
            Case "pulse": addOutputRow 6154, 1131, pdValue(pd.mean(i))
            Case "deathdate", "certdate"
                addOutputRow 43009, 1148, pdYYYYMMDD(pd.mean(i))
            Case "deathcause1", "deathcause1a"
                addOutputRow pdValue(pd.mean(i)), 1149, 1
            Case "deathcause1b"
                addOutputRow pdValue(pd.mean(i)), 1149, 2
            Case "deathcause1c"
                addOutputRow pdValue(pd.mean(i)), 1149, 3
            Case "deathcause2"
                addOutputRow pdValue(pd.mean(i)), 1149, 4
            Case "deathcause"
                ' cause of death without category stated
                addOutputRow pdValue(pd.mean(i)), 2001
            Case "negpmh": addOutputRow 11435, 2002, pdValue(pd.mean(i))
            Case "negfamily": addOutputRow 13240, 2003, pdValue(pd.mean(i))
            Case "query": addOutputRow 5494, 2004, pdValue(pd.mean(i))
            Case "gest": addOutputRow 55352, 2006, pdValue(pd.mean(i))
            Case "edd": addOutputRow 8879, 2007, pdYYYYMMDD(pd.mean(i))
            Case "lmp": addOutputRow 6769, 2008, pdYYYYMMDD(pd.mean(i))
            Case "admitdate": addOutputRow 43828, 2009, data1:=pdYYYYMMDD(pd.mean(i))
            Case "dischdate": addOutputRow 43828, 2009, data2:=pdYYYYMMDD(pd.mean(i))
            Case "sicknote"
                ' Can be a date or duration; try to extract both
                ' (if it is a date, duration will automatically be zero/blank)
                addOutputRow 5761, 2010, _
                    data1:=pdYYYYMMDD(pd.mean(i)), _
                    data2:=pdDurValue(pd.mean(i)), _
                    data3:=pdDurUnits(pd.mean(i))
            Case "followup"
                ' Can be a date or duration
                addOutputRow 1793, 2013, _
                    data1:=pdYYYYMMDD(pd.mean(i)), _
                    data2:=pdDurValue(pd.mean(i)), _
                    data3:=pdDurUnits(pd.mean(i))
            
            ' Standard quantitative lab results
            Case "albumin": addOutputRow 31969, 1152, 3, pdValue(pd.mean(i))
            Case "cobalamin": addOutputRow 7926, 1157, 3, pdValue(pd.mean(i))
            Case "calcium": addOutputRow 77, 1159, 3, pdValue(pd.mean(i))
            Case "cholesterol": addOutputRow 12, 1163, 3, pdValue(pd.mean(i))
            Case "creatinine": addOutputRow 5, 1165, 3, pdValue(pd.mean(i))
            Case "ferritin": addOutputRow 8491, 1169, 3, pdValue(pd.mean(i))
            Case "folate": addOutputRow 13748, 1170, 3, pdValue(pd.mean(i))
            Case "haemoglobin": addOutputRow 4, 1173, 3, pdValue(pd.mean(i))
            Case "hdl": addOutputRow 44, 1175, 3, pdValue(pd.mean(i))
            Case "ldl": addOutputRow 65, 1177, 3, pdValue(pd.mean(i))
            Case "mcv": addOutputRow 10, 1182, 3, pdValue(pd.mean(i))
            Case "platelets": addOutputRow 7, 1189, 3, pdValue(pd.mean(i))
            Case "trithyroid": addOutputRow 13791, 1197, 3, pdValue(pd.mean(i))
            Case "tetrathyroid": addOutputRow 941, 1198, 3, pdValue(pd.mean(i))
            Case "triglycerides": addOutputRow 37, 1202, 3, pdValue(pd.mean(i))
            Case "tsh": addOutputRow 13598, 1203, 3, pdValue(pd.mean(i))
            Case "urea": addOutputRow 18587, 1204, 3, pdValue(pd.mean(i))
            Case "wbc": addOutputRow 13818, 1207, 3, pdValue(pd.mean(i))
            Case "esr": addOutputRow 46, 1273, 3, pdValue(pd.mean(i))
            Case "glyhb": addOutputRow 14051, 1275, 3, pdValue(pd.mean(i))
            Case "pefr": addOutputRow 11772, 1311, 3, pdValue(pd.mean(i))
            Case "inr": addOutputRow 71, 1323, 3, pdValue(pd.mean(i))
            Case "rdw": addOutputRow 64, 2000, 3, pdValue(pd.mean(i))
            End Select
        End If
        ' Advance the counter by one row
        i = i + 1
    Loop
End If
End Sub

Function getmedcodes(targetpracid As Long, targettextid As Long) As String
' Returns the medcode mapping the given pracid and textid,
' or 0 if it is not found. The sorting is by pracid then textid.
' This function uses a binary search, comparing both pracid
' and textid with the target.
If ntexts = 0 Or (targetpracid = 0 And targettextid = 0) Then
    getmedcodes = "0"
    Exit Function
End If

Dim top As Long, bot As Long, trial As Long
top = 0
bot = ntexts
Do
    trial = Int((top + bot) / 2)
    If targetpracid < pracid(trial) Or _
        (targetpracid = pracid(trial) And targettextid < textid(trial)) Then
        bot = trial - 1
    ElseIf targetpracid > pracid(trial) Or _
        (targetpracid = pracid(trial) And targettextid > textid(trial)) Then
        top = trial + 1
    ElseIf (targetpracid = pracid(trial) And targettextid = textid(trial)) Then
        top = trial: bot = trial
    End If
Loop Until bot - top < 2

Dim pos_top As Long
Dim pos_bot As Long

If (targetpracid = pracid(top) And targettextid = textid(top) And top > 0) Then
    pos_top = top
    pos_bot = top
ElseIf (targetpracid = pracid(bot) And targettextid = textid(bot)) Then
    pos_top = bot
    pos_bot = bot
Else
    getmedcodes = "0"
    Exit Function
End If

' Find the top and bottom position of this run of medcodes
Do While (targetpracid = pracid(pos_top - 1) And targettextid = textid(pos_top - 1))
    If pos_top > 1 Then
        pos_top = pos_top - 1
    End If
Loop
Do While (targetpracid = pracid(pos_bot + 1) And targettextid = textid(pos_bot + 1))
    If pos_bot < maxtexts Then
        pos_bot = pos_bot + 1
    End If
Loop

' First character(s) are the number of medcodes
getmedcodes = pos_bot - pos_top + 1
Dim i As Long
For i = pos_top To pos_bot
    getmedcodes = getmedcodes & "|" & CStr(medcode(i))
Next
End Function

Function loadMedcodes(filename As String) As String
' Loads text id and medcodes from a comma separated text file with optional header:
' pracid, textid, medcode. If no header, it is assumed that the columns are in this order,
' otherwise the column names are used and additional columns are allowed.
' Returns a message stating whether the load was sucessful.

Dim fileno As Integer
Dim pracid_col As Long
Dim textid_col As Long
Dim medcode_col As Long
Dim rawstring As String
Dim message As String

On Error GoTo errortrapping:

message = "Loading pracid,textid,medcode from " & filename & newline
fileno = freefile

' Open the file and read the first line
Open filename For Input As #fileno
Line Input #fileno, rawstring

If (isHeader(rawstring) = False) Then
    ' No header, so assume the columns are in order
    Close #fileno
    Open filename For Input As #fileno
    pracid_col = 1
    textid_col = 2
    medcode_col = 3
Else
    ' Get column order from row numbers
    pracid_col = findColumn("pracid", rawstring)
    If pracid_col = 0 Then ' it may be quoted
        pracid_col = findColumn("""pracid""", rawstring)
    End If
    message = message & "pracid is in column " & pracid_col & newline
    textid_col = findColumn("textid", rawstring)
    If textid_col = 0 Then ' it may be quoted
        textid_col = findColumn("""textid""", rawstring)
    End If
    message = message & "textid is in column " & textid_col & newline
    medcode_col = findColumn("medcode", rawstring)
    If medcode_col = 0 Then ' it may be quoted
        medcode_col = findColumn("""medcode""", rawstring)
    End If
    message = message & "medcode is in column " & medcode_col & newline
End If

If (pracid_col > 0 And textid_col > 0 And medcode_col > 0) Then
    ' Load from text file
    ntexts = 0
    Do While Not EOF(fileno) And ntexts < maxtexts
        Line Input #fileno, rawstring
        If (rawstring <> "") Then
            ntexts = ntexts + 1
            pracid(ntexts) = dissect(rawstring, pracid_col, delim)
            textid(ntexts) = dissect(rawstring, textid_col, delim)
            medcode(ntexts) = dissect(rawstring, medcode_col, delim)
        End If
    Loop
    Close #fileno
    ' Trying to load too many texts
    If (ntexts >= maxtexts - 1) Then
        message = message & "Too many rows; maxtexts = " & maxtexts - 1 & newline
    End If
    loadMedcodes = message & ntexts & " rows loaded." & newline
    Exit Function
Else
    message = message & "Unable to locate columns" & newline
End If

errortrapping:
loadMedcodes = message & "ERROR"
End Function


Function isHeader(str As String) As Boolean
' Whether str is a possible header in a comma separated file
' If any of the columns are non-numeric, isHeader is True

isHeader = False
Dim thisSection As String
Dim i As Long
i = 0
Do
    i = i + 1
    thisSection = dissect(str, i, delim)
    If (thisSection <> "") Then
        If (is_numeric(Trim(thisSection), _
            dont_ignore_large_numbers:=True) = False) Then
            isHeader = True
        End If
    End If
Loop Until thisSection = "" Or isHeader = True
End Function

Function findColumn(colName As String, allNames As String) As Long
' Finds out the column number (first position) of colName in allNames, with comma delimiter
' Returns 0 if column name not found

findColumn = 0
Dim thisSection As String
Dim i As Long
i = 0
Do
    i = i + 1
    thisSection = dissect(allNames, i, delim)
    If (Trim(thisSection) = colName) Then
        findColumn = i
    End If
Loop Until thisSection = "" Or findColumn > 0
End Function

