dim digits			' Anzahl der Stellen, auf die die Seriennummer erweitert wird
dim modus			'1=Dateiname in Kleinbuchstaben, 2=Dateiname in Großbuchstaben
dim counter			'laufende Seriennummer der automatisch generierten Lims-Dateinamen
dim max_counter	'maximale Seriennummer der automatisch generierten Lims-Dateinamen
dim Filename		'Dateiname der kopierten TEMP_C.DAT für das LIMS
dim Zeile
dim compound(100),conc(100),unit(100)
dim sid_date,sampid,comp,conc1
dim date1,date2,datetime,HHMMSS,year,day,month,paramfile
dim ImportFileName, OutputFileName, Auto_Manual, Task, meth, methode1
dim scriptpath

' Verbindung zu Dateisystembefehlen herstellen
set fs = CreateObject("Scripting.FileSystemObject")
set WSHShell = CreateObject("WScript.Shell")				'nötig für Run-Befehl

counter=0
modus = 2

Set args = Wscript.Arguments
paramfile=args(0)
Filename = args(1)
argsnumber = args.length
'MsgBox ("argsnumber:" & argsnumber)

'MsgBox ("args(0):" & args(0))
set par = fs.OpenTextFile(paramfile)
do until par.atEndOfStream
	inpar = par.ReadLine
	param = Split(inpar, "=", -1, 0)
	if param(0)="max_counter" then max_counter=param(1)
	if param(0)="digits" then digits=param(1)
	if param(0)="append" then append=param(1)
	if param(0)="DecimalSeparator" then DecSep=param(1)
	if param(0)="Units" then Units=param(1)
	if param(0)="Separator" then sep=param(1)
	if param(0)="ImportFileName" then ImportFileName=param(1)
	if param(0)="ScriptPath" then scriptpath=param(1)
loop
par.close

sep1=Chr(9)
if sep="tab" then sep1=Chr(9)
if sep="semicolon" then sep1=";"
if sep="comma" then sep1=","
sep=sep1

logbuch = scriptpath & "output_to_csv.log"	' Name des Logbuchs
counterfile = scriptpath & "output_to_csv_counter.TXT"

if append=0 then
	if not fs.FileExists(counterfile) then
		problem = counterfile + " does not exist! Reset file counter !"
		'Logbuch im Append-Modus (8) öffnen und anlegen, falls es noch
		'gar nicht existiert (vbTrue):
	  	set logfile = fs.OpenTextFile(logbuch, 8, vbTrue)
		logfile.WriteLine Now & " : " & problem
		logfile.close

	'	MsgBox problem, vbExclamation

	else
	  	set LIMS_Transfile = fs.OpenTextFile(counterfile, 1, vbTrue)
		counter = LIMS_Transfile.ReadLine
		LIMS_Transfile.close

	end if
end if


if not fs.FileExists(ImportFileName) then
	problem = ImportFileName + " does not exist!"
	'Logbuch im Append-Modus (8) öffnen und anlegen, falls es noch
	'gar nicht existiert (vbTrue):
  	set logfile = fs.OpenTextFile(logbuch, 8, vbTrue)
	logfile.WriteLine Now & " : " & problem
	logfile.close
	WScript.Quit

'	MsgBox problem, vbExclamation
'	WScript.Quit
end if

if append = 0 and counter = max_counter then
	counter = 0
	problem = "File counter = " + Cstr(max_counter) + " ! Reset counter !"
	'Logbuch im Append-Modus (8) öffnen und anlegen, falls es noch
	'gar nicht existiert (vbTrue):
  	set logfile = fs.OpenTextFile(logbuch, 8, vbTrue)
	logfile.WriteLine Now & " : " & problem
	logfile.close

end if


set eingabe = fs.OpenTextFile(ImportFileName)
methode1 = eingabe.ReadLine 'MM name
'MsgBox ("MM File:" & methode1)
methode2=Split(methode1, "\", -1, 0)
if IsArray(methode2) then
	methode_index=UBound(methode2)
	methode1=methode2(methode_index)
	'MsgBox ("methode_index:" & Cstr(methode_index) & "  methode2(methode_index):" & methode2(methode_index))
end if
methode3=Split(methode1, ".", -1, 0)
methode1=methode3(0)
gelesen = eingabe.ReadLine 'Skip Line 2
sampid = eingabe.ReadLine 'Sample Id
'MsgBox ("Sample Id:" & sampid)
'samp_type = trim(left(sampid, 9))
'MsgBox ("samp_type:" & samp_type)
'if samp_type<>"Standard-" then dateiname = dateiname & LimsExt
'if samp_type="Standard-" then dateiname = dateiname & LimsExt_Standards
gelesen = eingabe.ReadLine 'Skip Line 4
gelesen = eingabe.ReadLine 'Skip Line 5
gelesen = eingabe.ReadLine 'Skip Line 6

endlist = vbFalse

ind = 7
nc=0
do until eingabe.atEndOfStream
	gelesen = eingabe.ReadLine
	'MsgBox ("Zeile:" & ind & " " & gelesen)
	if not endlist then
		if len(gelesen)>1 then
			ind = ind + 1
			nc = nc + 1
			Compound(nc) = trim(left(gelesen, 32))
			conc(nc) = trim(mid(gelesen,65,16))
			if len(DecSep)>0 then
				conc(nc) = replace(conc(nc),".",DecSep)
			end if
			'MsgBox ("Conc=" & Cstr(conc(nc)))
			unit(nc) = trim(mid(gelesen,82,8))
			'MsgBox ("Unit=" & unit(nc))
			'MsgBox ("Compound=" & Compound(nc) & "  " & conc(nc) & " " & unit(nc))

		else
			endlist = vbTrue
			znrendlist = eingabe.line - 1
			'MsgBox ("# Leerzeile:" & znrendlist)

		end if
	else
		if eingabe.line = znrendlist + 3 then
			date1 = gelesen ' Date and Time
			'MsgBox ("Date and Time:" & datum1)
		end if
		if eingabe.line = znrendlist + 6 then
			operator = gelesen ' Operator Name
			'MsgBox ("Operator:" & operator)
			operator = mid(operator,11) ' Operator Name
			'MsgBox ("Operator:" & operator)
		end if
	end if
loop
eingabe.close

if append = 0 and fs.FileExists(Filename) then
	problem = Filename + " exists already ! Will be overwritten !"
	'Logbuch im Append-Modus (8) öffnen und anlegen, falls es noch
	'gar nicht existiert (vbTrue):
  	set logfile = fs.OpenTextFile(logbuch, 8, vbTrue)
	logfile.WriteLine Now & " : " & problem
	logfile.close
end if


datetime = Split(date1, " ", -1, 0)

headzeile= "SampleName" & sep & "Date" & sep & "Time"
if Units="header" then
	unitszeile= "----------" & sep & "----" & sep & "----"
end if
zeile=sampid & sep & datetime(0) & sep & datetime(1)

if Units="header" then
	for i = 1 to nc
		headzeile = headzeile & sep & Compound(i)
		unitszeile = unitszeile & sep & unit(i)
		zeile = zeile & sep & conc(i)
	next
end if

if Units="col" then
	for i = 1 to nc
		headzeile=headzeile & sep & Compound(i)& sep & " "
		zeile=zeile & sep & conc(i)& sep & unit(i)
	next
end if

'MsgBox ("headzeile=" & headzeile)
'MsgBox ("zeile=" & zeile)

if append = 0 and fs.FileExists(Filename) then
	problem = Filename + " exists already ! Will be overwritten !"
	'Logbuch im Append-Modus (8) öffnen und anlegen, falls es noch
	'gar nicht existiert (vbTrue):
  	set logfile = fs.OpenTextFile(logbuch, 8, vbTrue)
	logfile.WriteLine Now & " : " & problem
	logfile.close
end if

if append = 0 then
	writemode = 2
	set Output = fs.OpenTextFile(Filename, writemode, vbTrue)
	Output.WriteLine headzeile
	if Units="header" then
		Output.WriteLine unitszeile
	end if
	Output.WriteLine zeile
	Output.close
end if

Fileexists = fs.FileExists(Filename)
if append = 1 and not Fileexists then
	writemode = 2
	set Output = fs.OpenTextFile(Filename, writemode, vbTrue)
	Output.WriteLine headzeile
	if Units="header" then
		Output.WriteLine unitszeile
	end if
	Output.WriteLine zeile
	Output.close
end if

if append = 1 and Fileexists then
	writemode = 8	'open file in append mode
	set Output = fs.OpenTextFile(Filename, writemode, vbTrue)
	Output.WriteLine zeile
	Output.close
end if


if append = 0 then
	set LIMS_Transfile = fs.OpenTextFile(counterfile, 2, vbTrue)
	LIMS_Transfile.WriteLine counter
	LIMS_Transfile.close
end if


function format(strvar,digits,filler)
	' formatiert eine Zeichenkette durch vorangestelltes Zeichen (filler), z.B. "0", und erweitert die Zahl auf diese
	' Weise auf die in der Variablen digit festgelegten Stellen

	format = string(digits-len(strvar), filler) & strvar
end function
