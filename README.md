freetext-matching-algorithm
===========================

Source code for the Freetext Matching Algorithm, a natural language processing system for clinical text. This program is used together with lookup tables which are in another public Git repository: https://github.com/anoopshah/freetext-matching-algorithm-lookups

This program is licensed under the GNU General Public Licence Version 3 (http://www.gnu.org/licenses/gpl-3.0-standalone.html).

If you use this program, please cite the following:

Shah AD, Martinez C, Hemingway H. The freetext matching algorithm: a computer program to extract diagnoses and causes of death from unstructured text in electronic health records. BMC Med Inform Decis Mak 2012;12:88 doi: 10.1186/1472-6947-12-88 http://www.biomedcentral.com/1472-6947/12/88/

Please send feedback, bug reports and suggested modifications to the lookup tables to anoop (@) doctors.org.uk.

## Acknowledgements

This software was developed as part of the CALIBER programme, funded by the Wellcome Trust (086091/Z/08/Z) and the National Institute for Health Research (NIHR) under its Programme Grants for Applied Research programme (RP-PG-0407-10314). The author is supported by a Wellcome Trust Clinical Research Training Fellowship (0938/30/Z/10/Z).

## Compiling and using the algorithm

The folder *vb* contains the source code (Visual Basic 6.0). It can be compiled using the Microsoft Visual Basic compiler or imported into a Visual Basic for Applications runtime environment such as Microsoft Access.

There are two pre-compiled executables in the binaries folder, both of which were compiled using Microsoft Visual Basic 6.0:

* fma15command.exe -- a command line version, which takes as its single argument the path to a configuration file. 
* fma15gui.exe -- a Visual Basic form, with a dialog box for entering the names of input and output files

The lookups must be downloaded from the repository and saved in a folder which is accessible to the program. Do not change the names of these files. If modifying the lookup tables, ensure that they remain in the same format (see https://github.com/anoopshah/freetext-matching-algorithm-lookups/blob/master/README.md for details). The binaries have been tested on Microsoft Windows and wine-1.5.26.

### Command-line version (fma15command.exe)

This program can be run on Windows from the command line thus:

    fma15command.exe argument

On Linux:

    wine fma15command.exe argument

where 'argument' is the path to a configuration file. An example of a configuration file is given in the *testing* folder. This executable is designed to work with the *CALIBERfma* R package, to facilitate the development and review of algorithms.

The configuration file must be a plain text file with the parameter name at the start of the line, followed by one or more spaces and then the parameter value (no quotes). The parameters can be listed in any order and some are optional. They are as follows:

* infile -- full filepath to input file with pracid, textid and free text
* medcodefile -- full filepath of file mapping pracid and textid to medcode
* outfile -- full filepath of output file (it will be over-written silently if a file of this name already exists)
* logfile -- full filepath of log file (it will be over-written silently if a file of this name already exists)
* lookups -- full path to folder containing lookup tables
* freetext -- a single free text to analyse in test mode. If supplied, the infile, medcodefile and outfile parameters are ignored and instead this single text is analysed 
* medcode -- a single medcode associated with freetext (text to analyse in test mode)

e.g.

    freetext     hypertensive 160/90
    medcode      1
    logfile      Z:\home\log1.log
    lookups      Z:\home\lookups\

### Graphical version (fma15gui.exe)

Type the input parameters in the dialog box and press 'START' to start the analysis. If a single free text is supplied, it will be analysed instead of the text file, and the result will be given in the log file. There are two slight differences from the command line version:

* If the lookups folder argument is left blank, it defaults to the folder containing the program itself.
* There is no box to enter a single medcode, so it is not possible to analyse a single free text with medcode in test mode using the graphical version.

## Testing

To test the program, supply a single free text string instead of input and output files. The program will return a detailed analysis report in the log file. However when analysing a text file, no text is written to the output file or log file.

## Input file format

All files must have Windows-style line endings.

* infile -- tab separated file with no quotes, and 3 columns without column headers:
    * Column 1: unique practice identifier (pracid)
    * Column 2: unique identifier for free text string within practice (textid)
    * Column 3: free text

* medcodefile -- comma separated values, with columns pracid, textid and medcode, sorted by pracid and textid. Column names are optional.
    * pracid -- unique practice identifier
    * textid -- unique identifier for free text string within practice
    * medcode -- medcode (can have multiple medcodes for each pracid / textid combination)

## Output file format

* outfile -- comma separated values file, with the following columns:
    * pracid
    * textid
    * origmedcode -- corresponds to medcode in medcodefile
    * medcode -- new medcode extracted from free text. This can be interpreted in a similar way to medcodes in the original Clinical Practice Research Datalink GOLD data format; a medcode in this column is for a past or present event for this patient.
    * enttype -- virtual entity type for information extracted from text
    * data1 ... data4 -- additional information (e.g. laboratory values, family history)

* logfile -- log file reporting which files were loaded and the number of texts analysed. In test mode, the log file also shows analysis information and results.



