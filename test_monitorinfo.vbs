'РўРµСЃС‚ РґР»СЏ РїСЂРѕРІРµСЂРєРё СЂР°Р±РѕС‚С‹ С„СѓРЅРєС†РёРё GetMonitorInfo РёР· lib_displayinfo.vbs
Option Explicit

Dim fso, testFile, result

'РЎРѕР·РґР°РµРј РѕР±СЉРµРєС‚ РґР»СЏ СЂР°Р±РѕС‚С‹ СЃ С„Р°Р№Р»РѕРІРѕР№ СЃРёСЃС‚РµРјРѕР№
Set fso = CreateObject("Scripting.FileSystemObject")

'РџРѕРґРєР»СЋС‡Р°РµРј Р±РёР±Р»РёРѕС‚РµРєСѓ
ExecuteGlobal GetFileContent("lib_core.vbs")
ExecuteGlobal GetFileContent("lib_displayinfo.vbs")

'РЎРѕР·РґР°РµРј С‚РµСЃС‚РѕРІС‹Р№ С„Р°Р№Р»
Set testFile = fso.CreateTextFile("test_result.txt", True)

'Р’С‹Р·С‹РІР°РµРј С„СѓРЅРєС†РёСЋ GetMonitorInfo РёР· Р±РёР±Р»РёРѕС‚РµРєРё
result = GetMonitorInfo()

'РџСЂРѕРІРµСЂСЏРµРј СЂРµР·СѓР»СЊС‚Р°С‚
If Len(result) > 0 Then
    testFile.WriteLine "РўРµСЃС‚ РїСЂРѕР№РґРµРЅ: С„СѓРЅРєС†РёСЏ РІРµСЂРЅСѓР»Р° РЅРµРїСѓСЃС‚СѓСЋ СЃС‚СЂРѕРєСѓ"
    testFile.WriteLine "Р РµР·СѓР»СЊС‚Р°С‚: " & result
Else
    testFile.WriteLine "РўРµСЃС‚ РЅРµ РїСЂРѕР№РґРµРЅ: С„СѓРЅРєС†РёСЏ РІРµСЂРЅСѓР»Р° РїСѓСЃС‚СѓСЋ СЃС‚СЂРѕРєСѓ"
End If

testFile.Close

'Р¤СѓРЅРєС†РёСЏ РґР»СЏ С‡С‚РµРЅРёСЏ СЃРѕРґРµСЂР¶РёРјРѕРіРѕ С„Р°Р№Р»Р°
Function GetFileContent(filePath)
    Dim file, content
    Set file = fso.OpenTextFile(filePath, 1)
    content = file.ReadAll
    file.Close
    Set file = Nothing
    GetFileContent = content
End Function