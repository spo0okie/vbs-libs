'����� ������ https://stackoverflow.com/questions/127318/is-there-any-sed-like-utility-for-cmd-exe
'����������� ��� � ������ �� ��������� ���.
'��� �������� ������� � ��������� ����� ������ ������ 
'��� � ����� ������ https://stackoverflow.com/questions/10091711/how-to-pass-a-command-with-spaces-and-quotes-as-a-single-parameter-to-cscript
'��� ����� ������� ������ ��������� ������� � �������� ������� ��� ���� � �������
'������ ������
'cscript //nologo sed.vbs "s/(installingUser value=~.*~)/installingUser value=~%username%~/g" < %installdir%\configuration.xml > %installdir%\configuration2.xml
'��� �������� � �������� XML <installingUser value="reviakin_admin" /> �� <installingUser value="�������_������������" />
Dim pat, patparts, rxp, inp
pat = WScript.Arguments(0)
patparts = Split(pat,"/")
Set rxp = new RegExp
rxp.Global = True
rxp.Multiline = False
rxp.Pattern = replace(patparts(1),"~",chr(34))
Do While Not WScript.StdIn.AtEndOfStream
  inp = WScript.StdIn.ReadLine()
  WScript.Echo rxp.Replace(inp, replace(patparts(2),"~",chr(34)))
Loop