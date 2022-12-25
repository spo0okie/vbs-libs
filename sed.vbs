'украл отсюда https://stackoverflow.com/questions/127318/is-there-any-sed-like-utility-for-cmd-exe
'естественно это и близко не настоящий сед.
'для передачи кавычек в аргументы нужно пихать тильды 
'как я понял отсюда https://stackoverflow.com/questions/10091711/how-to-pass-a-command-with-spaces-and-quotes-as-a-single-parameter-to-cscript
'это самый простой способ затолкать кавычки в аргумент который уже взят в кавычки
'пример вызова
'cscript //nologo sed.vbs "s/(installingUser value=~.*~)/installingUser value=~%username%~/g" < %installdir%\configuration.xml > %installdir%\configuration2.xml
'это заменяет в исходном XML <installingUser value="reviakin_admin" /> на <installingUser value="текущий_пользователь" />
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