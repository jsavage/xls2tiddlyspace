Attribute VB_Name = "Module1"
Function ttts(userpass As String, command As String, data As String, contenttype As String, url As String, category As String) As String
'Transfer To TiddlySpace TTTS V0.1
'Usage: put the following in a cell at the end of each row of data that describes the tiddler(s) you want created/ updated
'=ttts("user:password", "PUT",<cellContainingBodyofTiddler>, "application/json",<urlroot>&<tiddlerName>,<tag>)
'<cellContainingBodyofTiddler> should be a reference to a cell containing the body of the tiddler
'<urlroot> should be a reference to a cell containing the baseurl for all tiddlers in the space
'eg http://<space>.tiddlyspace.com:80/bags/<space>_public/tiddlers/
'<tiddlerName> should be a ref to the cell containing the name of the tiddler to be created
'<tag> should be a ref to the tag. Only a single tag is supported at present

'The purpose of this user-defined function is to facilitate migration/ transfer of data from excel to tiddlyspaace
'The reverse function is a simpler use case and can easily be implemented if required.
'Limitations:
'No warranty or guarantees with this.  Only use this if you understand what it is doing.
'There is no collision detection & minimal error handling
'Success recorded as OK + the time.
'When called the function over-writes the destination tiddler with the data from excel.
'So Excel acts as the master.  Any changes made to the excel spreadsheet are immediately replicated to tiddlyspace
'No provision for configuring a proxy server as yet
'Base address is embedded on First page of spreadsheet cell $E1
'Only one Tag is currently supported by the json wrapper.

'James Savage July 2012

Dim resptext As String
Dim WinHttpReq As WinHttp.WinHttpRequest
Set WinHttpReq = New WinHttpRequest
' Switch the mouse pointer to an hourglass while busy.
MousePointer = vbHourglass
'Retain as Proxy may need to be set
'WinHttpReq.WinHttpGetProxyForUrl
'WinHttpReq.SetProxy HTTPREQUEST_PROXYSETTING_PROXY,mvarProxyAddress & ":" & mvarProxyPort
WinHttpReq.Open command, Trim(url), False
WinHttpReq.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
WinHttpReq.SetRequestHeader "Authorization", "Basic " + Base64Encode(userpass)

WinHttpReq.SetRequestHeader "Content-Type", "application/json"

WinHttpReq.Send jsonouterwrap(jsonaddfield("text", Trim(data)) + "," + jsonaddarray("tags", Trim(category)))

' Get all response headers and status
resptext1 = WinHttpReq.GetAllResponseHeaders()
resptext2 = WinHttpReq.ResponseText()
resptext3 = Trim(Str(WinHttpReq.Status))
resptext4 = WinHttpReq.StatusText
' Switch the mouse pointer back to default.
MousePointer = vbDefault
If resptext3 = "204" Then
    ttts = "OK " + Str(Time()) ' no need to report detail as this is the result expected
Else
    ttts = resptext4 + ":" + resptext3 + ":" + resptext2 + ":" + resptext1
End If
'Application.Volatile
End Function

Function jsonouterwrap(text As String) As String
' puts outer wrap on json contents eg ' "[....]"
jsonouterwrap = "{" + text + "}"
End Function
Function jsonaddfield(field As String, text2wrap As String) As String
'When used as follows: jsonaddfield("text", Trim(data)))
'this Function returns a string formatted as follows "text":"..." where ... is replaced by the data
'Data is trimmed for leading and trailing spaces on assumption these are not significant.
jsonaddfield = Chr(34) + field + Chr(34) + ":" + Chr(34) + text2wrap + Chr(34)
End Function
Function jsonaddarray(field As String, text2wrap As String) As String
'This function is needed to support addition of tags to tiddlyspace. The tags need to be formatted as an array.
'When called as follows: jsonaddarray("tags", Trim("critical") results in "tags":["critical"]  where "critical" is the tag.
'This function only supports one tag at present
jsonaddarray = Chr(34) + field + Chr(34) + ":" + "[" + Chr(34) + text2wrap + Chr(34) + "]"
End Function

'Some unit tests for the json related functions
Sub testjsonaddfield()
MsgBox (jsonaddfield("text", "james"))
End Sub
Sub testjsonaddarray()
MsgBox (jsonaddarray("text", "tag1"))
End Sub
Sub testjsonwrap()
MsgBox (jsonouterwrap(jsonaddfield("text", "xxyyzz") + "," + jsonaddarray("tags", "noun")))
End Sub
'simple unit test for Base64 encoding
' compare output to that produced using a REST Client or other similar tool
Sub testb64()
 MsgBox (Base64Encode("user:mypass"))
End Sub

Function Base64Encode(sText)
'Acknowledgement: from http://stackoverflow.com/a/506992
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.createElement("base64")
    oNode.DataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function


'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
'Acknowledgement: from http://stackoverflow.com/a/506992
Function Stream_StringToBinary(text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string
Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And get binary data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function



