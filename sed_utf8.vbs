'����� ������ https://stackoverflow.com/questions/127318/is-there-any-sed-like-utility-for-cmd-exe
'����������� ��� � ������ �� ��������� ���.
'
'��� �������� ������� � ��������� ����� ������ ������ 
'��� � ����� ������ https://stackoverflow.com/questions/10091711/how-to-pass-a-command-with-spaces-and-quotes-as-a-single-parameter-to-cscript
'��� ����� ������� ������ ��������� ������� � �������� ������� ��� ���� � �������
'
'������� ��� ������ UTF-8
'https://stackoverflow.com/questions/13851473/read-utf-8-text-file-in-vbscript
'
'������ ������
'cscript //nologo sed.vbs "s/(installingUser value=~.*~)/installingUser value=~%username%~/g" < %installdir%\configuration.xml > %installdir%\configuration2.xml
'��� �������� � �������� XML <installingUser value="reviakin_admin" /> �� <installingUser value="�������_������������" />

Option Explicit

Dim strPattern, strTokens, objRegexp, objStream, strData

strPattern = WScript.Arguments(0)
strTokens = Split(replace(strPattern,"~",chr(34)),"/")

Set objRegexp = new RegExp

objRegexp.Global = True
objRegexp.Multiline = False
objRegexp.Pattern = tokens(1)

Set objStream = CreateObject("ADODB.Stream")

objStream.CharSet = "utf-8"
objStream.Open
objStream.LoadFromFile("C:\Users\admin\Desktop\ArtistCG\folder.txt")

strData = objStream.ReadText()

objStream.Close
Set objStream = Nothing
Do While Not WScript.StdIn.AtEndOfStream
  inp = WScript.StdIn.ReadLine()
  WScript.Echo rxp.Replace(inp, replace(patparts(2),"~",chr(34)))
Loop