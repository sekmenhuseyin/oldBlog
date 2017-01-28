<!--METADATA
	TYPE="TypeLib"
	NAME="Microsoft ActiveX Data Objects 2.5 Library"
	UUID="{00000205-0000-0010-8000-00AA006D2EA4}"
	VERSION="2.5"
-->

<%
' Field class represents interface to data passed within one field
'
' Two available methods of getting a field:
'	Set objField=objUpload.Fields("File1")
'	Set objField=objUpload("File1")
'
'
'	objField.Name
'		Name of the field as defined on the form
'
'	objFiled.Filepath
'		Path that file was sent from
'
'		ie: C:\Documents and Settings\lmoten\Desktop\Photo.gif
'
'	objField.FileDir
'		Directory that file was sent from
'
'		ie: C:\Documents and Settings\lmoten\Desktop
'
'	objField.FileExt
'		Uppercase Extension of the file
'
'		ie: GIF
'
'	objField.FileName
'		Name of the file
'
'		use: Response.AddHeader "Content-Disposition", "filename=""" & objField.FileName & """"
'
'		ie: Photo.gif
'
'	objField.ContentType
'		Type of binary data
'
'		use: Response.ContentType=objField.ContentType
'
'		ie: image/gif
'
'	objField.Value
'		Unicode value passed from form.  This value is empty if the field is binary data.
'
'		use: Response.Write "The value of this field is: " & objField.Value
'
'	objField.BinaryData
'		Contents of files binary data. (Integer SubType Array)
'
'		use: Response.BinaryWrite objField.BinaryData
'
'	objField.BLOB
'		Same thing as BinaryData but with a shorter name.  Added to help prevent
'		confusion with database access.
'
'		use: Call lobjRs.Fields("Image").AppendChunk(objField.BLOB)
'
'	objField.Length
'		byte size of Value or BinaryData - depending on type of field
'
'		use: Response.Write "The size of this file is: " & objField.Length
'
'	objField.BinaryAsText()
'		Converts binary data into unicode format.  Useful when you expect the user
'		to upload a text file and you have the need to interact with it.
'
'		use: Response.Write objField.BinaryAsText()
'
'	objField.SaveAs()
'		Saves binary data to a specified path.  This will overwrite any existing files.
'
'		use: objField.SaveAs(Server.MapPath("/Uploads/") & "\" & objField.FileName)
'
' ------------------------------------------------------------------------------
Class clsField
	
	Public Name				' Name of the field defined in form

	Private mstrPath		' Full path to file on visitors computer
							' C:\Documents and Settings\lmoten\Desktop\Photo.gif
	
	Public FileDir			' Directory that file existed in on visitors computer
							' C:\Documents and Settings\lmoten\Desktop

	Public FileExt			' Extension of the file
							' GIF

	Public FileName			' Name of the file
							' Photo.gif
	
	Public ContentType		' Content / Mime type of file
							' image/gif
							
	Public Value			' Unicode value of field (used for normail form fields - not files)
	
	Public BinaryData		' Binary data passed with field (for files)

	Public Length			' byte size of value or binary data

	Private mstrText		' Text buffer 
								' If text format of binary data is requested more then
								' once, this value will be read to prevent extra processing
	
' ------------------------------------------------------------------------------
	Public Property Get BLOB()
		BLOB=BinaryData
	End Property
' ------------------------------------------------------------------------------
	Public Function BinaryAsText()
		
		' Binary As Text returns the unicode equivilant of the binary data.
		' this is useful if you expect a visitor to upload a text file that
		' you will need to work with.
		
		' NOTICE:
		' NULL values will prematurely terminate your Unicode string.
		' NULLs are usually found within binary files more often then plain-text files.
		' a simple way around this may consist of replacing null values with another character
		' such as a space " "

		Dim lbinBytes
		Dim lobjRs
		
		' Don't convert binary data that does not exist
		If Length=0 Then Exit Function
		If LenB(BinaryData)=0 Then Exit Function
		
		' If we previously converted binary to text, return the buffered content
		If Not Len(mstrText)=0 Then
			BinaryAsText=mstrText
			Exit Function
		End If
		
		' Convert Integer Subtype Array to Byte Subtype Array
		lbinBytes=ASCII2Bytes(BinaryData)
   		
   		' Convert Byte Subtype Array to Unicode String
   		mstrText=Bytes2Unicode(lbinBytes)
   		
   		' Return Unicode Text
    	BinaryAsText=mstrText

	End Function
' ------------------------------------------------------------------------------
	Public Sub SaveAs(ByRef pstrFileName)
on error resume next
		Dim lobjStream
		Dim lobjRs
		Dim lbinBytes
		
		' Don't save files that do not posess binary data
		If Length=0 Then Exit Sub
		If LenB(BinaryData)=0 Then Exit Sub
		
		' Create magical objects from never never land
		Set lobjStream=Server.CreateObject("ADODB.Stream")
		
		' Let stream know we are working with binary data
		lobjStream.Type=adTypeBinary
		
		' Open stream
		Call lobjStream.Open()

		' Convert Integer Subtype Array to Byte Subtype Array
		lbinBytes=ASCII2Bytes(BinaryData)
		
		' Write binary data to stream
		Call lobjStream.Write(lbinBytes)
		
		' Save the binary data to file system
		'	Overwrites file if previously exists!
		Call lobjStream.SaveToFile(pstrFileName, adSaveCreateOverWrite)
		
		' Close the stream object
		Call lobjStream.Close()
		
		' Release objects
		Set lobjStream=Nothing
	if err.number=0 then pass=1
	End Sub
' ------------------------------------------------------------------------------
	Public Property Let FilePath(ByRef pstrPath)
		
		mstrPath=pstrPath
		
		' Parse File Ext
		If Not InStrRev(pstrPath, ".")=0 Then
			FileExt=Mid(pstrPath, InStrRev(pstrPath, ".") + 1)
			FileExt=UCase(FileExt)
		End If
		
		' Parse File Name
		If Not InStrRev(pstrPath, "\")=0 Then
			FileName=Mid(pstrPath, InStrRev(pstrPath, "\") + 1)
		End If
		
		' Parse File Dir
		If Not InStrRev(pstrPath, "\")=0 Then
			FileDir=Mid(pstrPath, 1, InStrRev(pstrPath, "\") - 1)
		End If
		
	End Property
' ------------------------------------------------------------------------------
	Public Property Get FilePath()
		FilePath=mstrPath
	End Property
' ------------------------------------------------------------------------------
	Private Function ASCII2Bytes(ByRef pbinBinaryData)
	
		Dim lobjRs
		Dim llngLength
		Dim lbinBuffer
		
		' get number of bytes
		llngLength=LenB(pbinBinaryData)
		
		Set lobjRs=Server.CreateObject("ADODB.Recordset")
		
		' create field in an empty recordset to hold binary data
		Call lobjRs.Fields.Append("BinaryData", adLongVarBinary, llngLength)
		
		' Open recordset
		Call lobjRs.Open()
		
		' Add a new record to recordset
		Call lobjRs.AddNew()
		
		' Populate field with binary data
		Call lobjRs.Fields("BinaryData").AppendChunk(pbinBinaryData & ChrB(0))
		
		' Update / Convert Binary Data
			' Although the data we have is binary - it has still been
			' formatted as 4 bytes to represent each byte.  When we
			' update the recordset, the Integer Subtype Array that we
			' passed into the Recordset will be converted into a
			' Byte Subtype Array
		Call lobjRs.Update()
		
		' Request binary data and save to stream
		lbinBuffer=lobjRs.Fields("BinaryData").GetChunk(llngLength)
		
		' Close recordset
		Call lobjRs.Close()
		
		' Release recordset from memory
		Set lobjRs=Nothing
		
		' Return Bytes
		ASCII2Bytes=lbinBuffer
		
	End Function
' ------------------------------------------------------------------------------
	Private Function Bytes2Unicode(ByRef pbinBytes)

		Dim lobjRs
		Dim llngLength
		Dim lstrBuffer
		
		llngLength=LenB(pbinBytes)
				
		Set lobjRs=Server.CreateObject("ADODB.Recordset")

		' Create field in an empty recordset to hold binary data
    	Call lobjRs.Fields.Append("BinaryData", adLongVarChar, llngLength)
    	
    	' Open Recordset
    	Call lobjRs.Open()
    	
    	' Add a new record to recordset
    	Call lobjRs.AddNew()
    	
    	' Populate field with binary data
    	Call lobjRs.Fields("BinaryData").AppendChunk(pbinBytes)
    	
    	' Update / Convert.
    		' Ensure bytes are proper subtype
    	Call lobjRs.Update()
    	
    	' Request unicode value of binary data
    	lstrBuffer=lobjRs.Fields("BinaryData").Value
    	
    	' Close recordset
    	Call lobjRs.Close()

    	' Release recordset from memory
    	Set lobjRs=Nothing
	
		' Return Unicode
		Bytes2Unicode=lstrBuffer
		
	End Function

' ------------------------------------------------------------------------------
End Class
' ------------------------------------------------------------------------------



' Upload class retrieves multi-part form data posted to web page
' and parses it into objects that are easy to interface with.
' Requires MDAC (ADODB) COM components found on most servers today
' Additional compenents are not necessary.
'
' Demo:
'	Set objUpload=new clsUpload
'		Initializes object and parses all posted multi-part from data.
'		Once this as been done, Access to the Request object is restricted
'
'	objUpload.Count
'		Number of fields retrieved
'
'		use: Response.Write "There are " & objUpload.Count & " fields."
'
'	objUpload.Fields
'		Access to field objects.  This is the default propert so it does
'		not necessarily have to be specified.  You can also determine if
'		you wish to specify the field index, or the field name.
'
'		Use:
'			Set objField=objUpload.Fields("File1")
'			Set objField=objUpload("File1")
'			Set objField=objUpload.Fields(0)
'			Set objField=objUpload(0)
'			Response.Write objUpload("File1").Name
'			Response.Write objUpload(0).Name
'
' ------------------------------------------------------------------------------
'
' List of all fields passed:
'
'	For i=0 To objUpload.Count - 1
'		Response.Write objUpload(i).Name & "<br />"
'	Next
'
' ------------------------------------------------------------------------------
'
' HTML needed to post multipart/form-data
'
'<FORM method="post" encType="multipart/form-data" action="Upload.asp">
'	<INPUT type="File" name="File1">
'	<INPUT type="Submit" value="Upload">
'</FORM>

Class clsUpload
' ------------------------------------------------------------------------------

	Private mbinData			' bytes visitor sent to server
	Private mlngChunkIndex		' byte where next chunk starts
	Private mlngBytesReceived	' length of data
	Private mstrDelimiter		' Delimiter between multipart/form-data (43 chars)

	Private CR					' ANSI Carriage Return
	Private LF					' ANSI Line Feed
	Private CRLF				' ANSI Carriage Return & Line Feed
	
	Private mobjFieldAry()		' Array to hold field objects
	Private mlngCount			' Number of fields parsed
	
' ------------------------------------------------------------------------------
	Private Sub RequestData

		Dim llngLength		' Number of bytes received
		
		' Determine number bytes visitor sent
		mlngBytesReceived=Request.TotalBytes
		
		' Store bytes recieved from visitor
		mbinData=Request.BinaryRead(mlngBytesReceived)
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub ParseDelimiter()

		' Delimiter seperates multiple pieces of form data
			' "around" 43 characters in length
			' next character afterwards is carriage return (except last line has two --)
			' first part of delmiter is dashes followed by hex number
			' hex number is possibly the browsers session id?

		' Examples:

		' -----------------------------7d230d1f940246
		' -----------------------------7d22ee291ae0114

		mstrDelimiter=MidB(mbinData, 1, InStrB(1, mbinData, CRLF) - 1)
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub ParseData()

		' This procedure loops through each section (chunk) found within the
		' delimiters and sends them to the parse chunk routine
		
		Dim llngStart	' start position of chunk data
		Dim llngLength	' Length of chunk
		Dim llngEnd		' Last position of chunk data
		Dim lbinChunk	' Binary contents of chunk
		
		' Initialize at first character
		llngStart=1
		
		' Find start position
		llngStart=InStrB(llngStart, mbinData, mstrDelimiter & CRLF)
		
		' While the start posotion was found
		While Not llngStart=0
			
			' Find the end position (after the start position)
			llngEnd=InStrB(llngStart + 1, mbinData, mstrDelimiter) - 2
			
			' Determine Length of chunk
			llngLength=llngEnd - llngStart
			
			' Pull out the chunk
			lbinChunk=MidB(mbinData, llngStart, llngLength)
			
			' Parse the chunk
			Call ParseChunk(lbinChunk)
			
			' Look for next chunk after the start position
			llngStart=InStrB(llngStart + 1, mbinData, mstrDelimiter & CRLF)
			
		Wend
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub ParseChunk(ByRef pbinChunk)
	
		' This procedure gets a chunk passed to it and parses its contents.
		' There is a general format that the chunk follows.

		' First, the deliminator appears

		' Next, headers are listed on each line that define properties of the chunk.

		'	Content-Disposition: form-data: name="File1"; filename="C:\Photo.gif"
		'	Content-Type: image/gif
	
		' After this, a blank line appears and is followed by the binary data.
		
		Dim lstrName			' Name of field
		Dim lstrFileName		' File name of binary data
		Dim lstrContentType		' Content type of binary data
		Dim lbinData			' Binary data
		Dim lstrDisposition		' Content Disposition
		Dim lstrValue			' Value of field
		
		' Parse out the content dispostion
		lstrDisposition=ParseDisposition(pbinChunk)

			' And Parse the Name
			lstrName=ParseName(lstrDisposition)

			' And the file name
			lstrFileName=ParseFileName(lstrDisposition)

		' Parse out the Content Type
		lstrContentType=ParseContentType(pbinChunk)
		
		' If the content type is not defined, then assume the
		' field is a normal form field
		If lstrContentType="" Then

			' Parse Binary Data as Unicode
			lstrValue=CStrU(ParseBinaryData(pbinChunk))
		
		' Else assume the field is binary data
		Else
			
			' Parse Binary Data
			lbinData=ParseBinaryData(pbinChunk)

		End If
		
		' Add a new field
		Call AddField(lstrName, lstrFileName, lstrContentType, lstrValue, lbinData)
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub AddField(ByRef pstrName, ByRef pstrFileName, ByRef pstrContentType, ByRef pstrValue, ByRef pbinData)

		Dim lobjField		' Field object class
		
		' Add a new index to the field array
		' Make certain not to destroy current fields
		ReDim Preserve mobjFieldAry(mlngCount)

		' Create new field object
		Set lobjField=New clsField
		
		' Set field properties
		lobjField.Name=pstrName
		lobjField.FilePath=pstrFileName				
		lobjField.ContentType=pstrContentType

		' If field is not a binary file
		If LenB(pbinData)=0 Then
			
			lobjField.BinaryData=ChrB(0)
			lobjField.Value=pstrValue
			lobjField.Length=Len(pstrValue)

		' Else field is a binary file
		Else

			lobjField.BinaryData=pbinData
			lobjField.Length=LenB(pbinData)
			lobjField.Value=""

		End If

		' Set field array index to new field
		Set mobjFieldAry(mlngCount)=lobjField
		
		' Incriment field count
		mlngCount=mlngCount + 1
		
	End Sub
' ------------------------------------------------------------------------------
	Private Function ParseBinaryData(ByRef pbinChunk)
	
		' Parses binary content of the chunk
		
		Dim llngStart	' Start Position

		' Find first occurence of a blank line
		llngStart=InStrB(1, pbinChunk, CRLF & CRLF)
		
		' If it doesn't exist, then return nothing
		If llngStart=0 Then Exit Function
		
		' Incriment start to pass carriage returns and line feeds
		llngStart=llngStart + 4
		
		' Return the last part of the chunk after the start position
		ParseBinaryData=MidB(pbinChunk, llngStart)
		
	End Function
' ------------------------------------------------------------------------------
	Private Function ParseContentType(ByRef pbinChunk)
		
		' Parses the content type of a binary file.
		'	example: image/gif is the content type of a GIF image.
		
		Dim llngStart	' Start Position
		Dim llngEnd		' End Position
		Dim llngLength	' Length
		
		' Fid the first occurance of a line starting with Content-Type:
		llngStart=InStrB(1, pbinChunk, CRLF & CStrB("Content-Type:"), vbTextCompare)
		
		' If not found, return nothing
		If llngStart=0 Then Exit Function
		
		' Find the end of the line
		llngEnd=InStrB(llngStart + 15, pbinChunk, CR)
		
		' If not found, return nothing
		If llngEnd=0 Then Exit Function
		
		' Adjust start position to start after the text "Content-Type:"
		llngStart=llngStart + 15
		
		' If the start position is the same or past the end, return nothing
		If llngStart >= llngEnd Then Exit Function
		
		' Determine length
		llngLength=llngEnd - llngStart
		
		' Pull out content type
		' Convert to unicode
		' Trim out whitespace
		' Return results
		ParseContentType=Trim(CStrU(MidB(pbinChunk, llngStart, llngLength)))

	End Function
' ------------------------------------------------------------------------------
	Private Function ParseDisposition(ByRef pbinChunk)
	
		' Parses the content-disposition from a chunk of data
		'
		' Example:
		'
		'	Content-Disposition: form-data: name="File1"; filename="C:\Photo.gif"
		'
		'	Would Return:
		'		form-data: name="File1"; filename="C:\Photo.gif"
		
		Dim llngStart	' Start Position
		Dim llngEnd		' End Position
		Dim llngLength	' Length
		
		' Find first occurance of a line starting with Content-Disposition:
		llngStart=InStrB(1, pbinChunk, CRLF & CStrB("Content-Disposition:"), vbTextCompare)
		
		' If not found, return nothing
		If llngStart=0 Then Exit Function
		
		' Find the end of the line
		llngEnd=InStrB(llngStart + 22, pbinChunk, CRLF)
		
		' If not found, return nothing
		If llngEnd=0 Then Exit Function
		
		' Adjust start position to start after the text "Content-Disposition:"
		llngStart=llngStart + 22
		
		' If the start position is the same or past the end, return nothing
		If llngStart >= llngEnd Then Exit Function
		
		' Determine Length
		llngLength=llngEnd - llngStart
		
		' Pull out content disposition
		' Convert to Unicode
		' Return Results
		ParseDisposition=CStrU(MidB(pbinChunk, llngStart, llngLength))

	End Function
' ------------------------------------------------------------------------------
	Private Function ParseName(ByRef pstrDisposition)

		' Parses the name of the field from the content disposition
		'
		' Example
		'
		'	form-data: name="File1"; filename="C:\Photo.gif"
		'
		'	Would Return:
		'		File1
		
		Dim llngStart	' Start Position
		Dim llngEnd		' End Position
		Dim llngLength	' Length
		
		' Find first occurance of text name="
		llngStart=InStr(1, pstrDisposition, "name=""", vbTextCompare)
		
		' If not found, return nothing
		If llngStart=0 Then Exit Function
		
		' Find the closing quote
		llngEnd=InStr(llngStart + 6, pstrDisposition, """")
		
		' If not found, return nothing
		If llngEnd=0 Then Exit Function
		
		' Adjust start position to start after the text name="
		llngStart=llngStart + 6
		
		' If the start position is the same or past the end, return nothing
		If llngStart >= llngEnd Then Exit Function
		
		' Determine Length
		llngLength=llngEnd - llngStart
		
		' Pull out field name
		' Return results
		ParseName=Mid(pstrDisposition, llngStart, llngLength)
		
	End Function
' ------------------------------------------------------------------------------
	Private Function ParseFileName(ByRef pstrDisposition)
		' Parses the name of the field from the content disposition
		'
		' Example
		'
		'	form-data: name="File1"; filename="C:\Photo.gif"
		'
		'	Would Return:
		'		C:\Photo.gif
		
		Dim llngStart	' Start Position
		Dim llngEnd		' End Position
		Dim llngLength	' Length
		
		' Find first occurance of text filename="
		llngStart=InStr(1, pstrDisposition, "filename=""", vbTextCompare)
		
		' If not found, return nothing
		If llngStart=0 Then Exit Function
		
		' Find the closing quote
		llngEnd=InStr(llngStart + 10, pstrDisposition, """")
		
		' If not found, return nothing
		If llngEnd=0 Then Exit Function
		
		' Adjust start position to start after the text filename="
		llngStart=llngStart + 10
		
		' If the start position is the same of past the end, return nothing
		If llngStart >= llngEnd Then Exit Function
		
		' Determine length
		llngLength=llngEnd - llngStart
		
		' Pull out file name
		' Return results
		ParseFileName=Mid(pstrDisposition, llngStart, llngLength)
		
	End Function
' ------------------------------------------------------------------------------
	Public Property Get Count()
		
		' Return number of fields found
		Count=mlngCount
		
	End Property
' ------------------------------------------------------------------------------
	
	Public Default Property Get Fields(ByVal pstrName)

		Dim llngIndex	' Index of current field
		
		' If a number was passed
		If IsNumeric(pstrName) Then
			
			llngIndex=CLng(pstrName)
			
			' If programmer requested an invalid number
			If llngIndex > mlngCount - 1 Or llngIndex < 0 Then
				' Raise an error
				Call Err.Raise(vbObjectError + 1, "clsUpload.asp", "Object does not exist within the ordinal reference.")
				Exit Property
			End If
				
			' Return the field class for the index specified
			Set Fields=mobjFieldAry(pstrName)
		
		' Else a field name was passed
		Else
		
			' convert name to lowercase
			pstrName=LCase(pstrname)
			
			' Loop through each field
			For llngIndex=0 To mlngCount - 1
				
				' If name matches current fields name in lowercase
				If LCase(mobjFieldAry(llngIndex).Name)=pstrName Then
					
					' Return Field Class
					Set Fields=mobjFieldAry(llngIndex)
					Exit Property
					
				End If
			
			Next
		
		End If

		' If matches were not found, return an empty field
		Set Fields=New clsField
		
'		' ERROR ON NonExistant:
'		' If matches were not found, raise an error of a non-existent field
'		Call Err.Raise(vbObjectError + 1, "clsUpload.asp", "Object does not exist within the ordinal reference.")
'		Exit Property

	End Property
' ------------------------------------------------------------------------------
	Private Sub Class_Terminate()
		
		' This event is called when you destroy the class.
		'
		' Example:
		'	Set objUpload=Nothing
		'
		' Example:
		'	Response.End
		'
		' Example:
		'	Page finnishes executing ...
		
		Dim llngIndex	' Current Field Index
		
		' Loop through fields
		For llngIndex=0 To mlngCount - 1
			
			' Release field object
			Set mobjFieldAry(llngIndex)=Nothing
			
		Next
		
		' Redimension array and remove all data within
		ReDim mobjFieldAry(-1)
		
	End Sub
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		
		' This event is called when you instantiate the class.
		'
		' Example:
		'	Set objUpload=New clsUpload
		
		' Redimension array with nothing
		ReDim mobjFieldAry(-1)
		
		' Compile ANSI equivilants of carriage returns and line feeds
		
		CR=ChrB(Asc(vbCr))	' vbCr		Carriage Return
		LF=ChrB(Asc(vbLf))	' vbLf		Line Feed
		CRLF=CR & LF			' vbCrLf	Carriage Return & Line Feed

		' Set field count to zero
		mlngCount=0
		
		' Request data
		Call RequestData
		
		' Parse out the delimiter
		Call ParseDelimiter()
		
		' Parse the data
		Call ParseData
		
	End Sub
' ------------------------------------------------------------------------------
	Private Function CStrU(ByRef pstrANSI)
		
		' Converts an ANSI string to Unicode
		' Best used for small strings
		
		Dim llngLength	' Length of ANSI string
		Dim llngIndex	' Current position
		
		' determine length
		llngLength=LenB(pstrANSI)
		
		' Loop through each character
		For llngIndex=1 To llngLength
		
			' Pull out ANSI character
			' Get Ascii value of ANSI character
			' Get Unicode Character from Ascii
			' Append character to results
			CStrU=CStrU & Chr(AscB(MidB(pstrANSI, llngIndex, 1)))
		
		Next

	End Function
' ------------------------------------------------------------------------------
	Private Function CStrB(ByRef pstrUnicode)

		' Converts a Unicode string to ANSI
		' Best used for small strings
		
		Dim llngLength	' Length of ANSI string
		Dim llngIndex	' Current position
		
		' determine length
		llngLength=Len(pstrUnicode)
		
		' Loop through each character
		For llngIndex=1 To llngLength
		
			' Pull out Unicode character
			' Get Ascii value of Unicode character
			' Get ANSI Character from Ascii
			' Append character to results
			CStrB=CStrB & ChrB(Asc(Mid(pstrUnicode, llngIndex, 1)))
		
		Next
		
	End Function
' ------------------------------------------------------------------------------
End Class
' ------------------------------------------------------------------------------
%>
