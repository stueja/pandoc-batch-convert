@echo off
REM cfc.bat
REM Jan Stühler, 2018-07-03

REM PROCESS:
REM 1.) create mediawiki formatted file (e. g. a Workbook)
REM 2.) create a file reference.docx with configured styles (e. g. headlines, enumerations, page numbers, ...)
REM 3.) have this script convert the mediawiki file to docx
REM 4.) open the (new) word file

REM cfc.bat converts a mediawiki formatted text file to docx format
REM REQUIREMENTS:
REM - pandoc
REM - a reference file
REM   (that is a Word document with configured styles)

REM USAGE:
REM invoke via command line (cbc.bat c:\path\to\mediawiki.txt) or
REM drag a file onto it in Windows Explorer

set "pandocexe=.\pandoc.exe"
set "reffile=.\reference.docx"
set "tocdepth=5"
set "outfile=.\pandoc.out.docx"

IF %* == "" GOTO Error1

REM %* = all arguments, to catch file names with whitespace
REM using it without double quotes (%* instead of "%*"), because Windows adds them itself if whitespace is contained
echo "%pandocexe%" -s -f mediawiki -t docx --reference-doc="%reffile%" --toc --toc-depth=%tocdepth% %* -o "%outfile%"
"%pandocexe%" -s -f mediawiki -t docx --reference-doc="%reffile%" --toc --toc-depth=%tocdepth% %* -o "%outfile%"
exit

:Error1
echo no input file given
exit
