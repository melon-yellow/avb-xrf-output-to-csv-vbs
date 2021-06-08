
dim paramfile, filename
dim argsnumber, scriptPath, dataPath
dim fileContents, lines, counted, timesSkip
dim folder, files, order, Torder

Const FOR_READING = 1 
Const FOR_WRITING = 2
set fs = CreateObject("Scripting.FileSystemObject")
set WSHShell = CreateObject("WScript.Shell")

Set args = Wscript.Arguments
paramfile = args(0)
filename = args(1)
argsnumber = args.length

set par = fs.OpenTextFile(paramfile, FOR_READING)
do until par.atEndOfStream
	inpar = par.ReadLine
	param = Split(inpar, "=", -1, 0)
	if param(0)="ScriptPath" then scriptPath=param(1)
loop
par.close

' Set Const Parameters
Const Frame = 32
dataPath = scriptPath & "data\"

' Set Reference Lines
refLines = getContents(filename)

validate(filename)

' Get Length of Array
Function length(arr)
	length = (UBound(arr) + 1)
End Function

' Get File Contents
Function getContents(filename)
	dim file, fileContents
	Set file = fs.OpenTextFile(filename, FOR_READING)
	fileContents = file.ReadAll
	file.Close
	getContents = Split(fileContents, vbNewLine)
End Function

' Set Top of Csv
Function settop(file)
	file.WriteLine refLines(0)
	file.Write refLines(1)
End Function

' Validate File
Function validate(filename)

	dim lines, relines, counted
	dim file, timesSkip

	' Read Temp File
	lines = getContents(filename)
	counted = length(lines)

	' Check if File Exceeds Max Length
	If counted > (Frame + 2) Then
		' Get Times to Skip
		timesSkip = (counted - Frame)
		' Construct Files
		relines = Slice(lines, 2, (timesSkip - 1))
		Construct(relines)
		' Write to Temp
		Set file = fs.OpenTextFile(filename, FOR_WRITING)
		settop(file)
		file.Write vbNewLine
		For i = timesSkip To (counted - 1)
			file.Write lines(i)
			If i < (counted - 1) Then
				file.Write vbNewLine
			End If
		Next
		file.Close
	End If

End Function

' Validate Historian Files
Function revalidate(filename)

	dim file, lines
	dim relines, counted

	' Read Temp File
	lines = getContents(filename)
	counted = length(lines)

	' Check if File Exceeds Max Length
	If counted > (Frame + 2) Then
		' Construct Files
		relines = Slice(lines, (Frame + 2), (counted - 1))
		Construct(relines)
		' Write to File
		Set file = fs.OpenTextFile(filename, FOR_WRITING)
		For i = 0 To (Frame + 1)
			file.Write lines(i)
			If i < (Frame + 1) Then
				file.Write vbNewLine
			End If
		Next
		file.Close
	End If

End Function

' Create Order File
Function newOrder(order)
	dim file, orderName
	orderName = dataPath & order & ".csv"
	If Not fs.FileExists(orderName) Then
		Set file = fs.CreateTextFile(orderName, true)
		settop(file)
		file.Close
	End If
	newOrder = orderName
End Function

' Construct Files by Order
Function Construct(relines)

	dim folder, files
	dim file, lines, counted
	dim order, Torder, orderName
	dim append

	' Get Order File
	Set folder = fs.GetFolder(dataPath)
	Set files = folder.Files
	order = 1
	For Each file In files
		if InStr(file.Name, ".csv") Then
			Torder = Split(file.Name, ".csv")(0)
			If IsNumeric(Torder) Then
				If (Cint(Torder) > Cint(order)) Then
					order = Torder
				End If
			End If
		End If
	Next

	' Create Order File
	orderName = newOrder(order)

	' Read Order File
	lines = getContents(orderName)
	counted = length(lines)

	' Write to Order
	If counted >= (Frame + 2) Then
		order = (order + 1)
		orderName = newOrder(order)
		lines = getContents(orderName)
		counted = length(lines)
	End If

	' Write to Order
	Set file = fs.OpenTextFile(orderName, FOR_WRITING)
	For i = 0 To (counted - 1)
		file.WriteLine lines(i)
	Next
	append = Join(relines, vbNewLine)
	file.Write append
	file.Close

	' Call Recursive
	revalidate(orderName)

	' Return Order
	Construct = order
	
End Function

Function Slice (aInput, Byval aStart, Byval aEnd)
    If IsArray(aInput) Then
        Dim i
        Dim intStep
        Dim arrReturn
        If aStart < 0 Then
            aStart = aStart + Ubound(aInput) + 1
        End If
        If aEnd < 0 Then
            aEnd = aEnd + Ubound(aInput) + 1
        End If
        Redim arrReturn(Abs(aStart - aEnd))
        If aStart > aEnd Then
            intStep = -1
        Else
            intStep = 1
        End If
        For i = aStart To aEnd Step intStep
            If Isobject(aInput(i)) Then
                Set arrReturn(Abs(i-aStart)) = aInput(i)
            Else
                arrReturn(Abs(i-aStart)) = aInput(i)
            End If
        Next
        Slice = arrReturn
    Else
        Slice = Null
    End If
End Function
