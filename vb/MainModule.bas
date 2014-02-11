' Module: MainModule -- to invoke the program from a command line

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

Sub Main()
' Loads arguments from configuration file and runs the analysis.
' The command-line argument (Command) is the location of the configuration file.
do_analysis logfile:=getParameterFromFile("logfile", Command), _
    lookups:=getParameterFromFile("lookups", Command), _
    infile:=getParameterFromFile("infile", Command), _
    medcodefile:=getParameterFromFile("medcodefile", Command), _
    outfile:=getParameterFromFile("outfile", Command), _
    freetext:=getParameterFromFile("freetext", Command), _
    medcode:=getParameterFromFile("medcode", Command), _
    ignoreerrors:=getParameterFromFile("ignoreerrors", Command)
End Sub

Function getParameterFromFile(parameterName As String, filename As String) As String
' Gets parameter from file, where each line of the file has the format:
' parameter   value
' (separated by at least one space)
Dim fileno As Integer, rawtext As String
fileno = freefile
Open filename For Input As fileno
Do
    Line Input #fileno, rawtext
    If Left(rawtext, Len(parameterName) + 1) = parameterName & " " Then
        getParameterFromFile = Trim(Mid(rawtext, Len(parameterName) + 1))
        Close fileno
        Exit Function
    End If
Loop Until EOF(fileno)
Close fileno
End Function
