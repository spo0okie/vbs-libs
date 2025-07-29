Option Explicit
'проверка необходимости перезагрузки с пояснением причин

' Глобальная переменная для хранения причины перезагрузки
Public g_RebootReason

' Основная функция (совместима с lib_reboot.vbs)
Function isRebootPending()
    g_RebootReason = "No reboot required" ' Сбрасываем причину по умолчанию
    isRebootPending = False
    
    'Что проверяет: Указывает, что обновление Windows или установка ПО требует перезагрузки.
    'Когда требует перезагрузки: Если значение существует и не равно 0.
    'Типичные причины:
    '  Установка обновлений Windows.
    '  Установка/удаление программ, которые заменяют исполняемые файлы.
    Dim testValue
    testValue = regRead("HKLM\SOFTWARE\Microsoft\Updates\UpdateExeVolatile")
    If (Not testValue = False) And (Not testValue = 0) Then
        g_RebootReason = "UpdateExeVolatile registry key indicates reboot is required"
        isRebootPending = True
        Exit Function
    End If

    
    'Что проверяет: Очередь операций переименования/удаления файлов, которые выполнятся при перезагрузке.
    'Когда требует перезагрузки: Если ключ существует и содержит непустые строки.
    'Формат данных:
    'Мультистроковый REG_MULTI_SZ, где каждая операция состоит из двух строк:
    '    "исходный_путь" -> "новый_путь" (переименование)
    '    "путь_к_файлу" -> "" (удаление)
    'Типичные причины:
    '    Обновление системных файлов (например, C:\Windows\system32\file.dll -> C:\Windows\system32\file.dll.new).
    '    Удаление временных файлов инсталлятора.
    Dim fileRenameOps
    fileRenameOps = GetPendingFileRenameOperations()
    If fileRenameOps <> "" Then
        g_RebootReason = "Pending file rename operations:" & vbCrLf & fileRenameOps
        isRebootPending = True
        Exit Function
    End If

    
    'Что проверяет: Флаг, что Windows Update установил обновления и ждет перезагрузки.
    'Когда требует перезагрузки: Если ключ существует (даже пустой).
    'Типичные причины:
    '    Установлены критические обновления безопасности.
    '    Обновления ядра Windows.
    If regExists("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") Then 
        g_RebootReason = "Windows Update requires reboot (HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired)"
        isRebootPending = True
    End if

    
    'Windows Update использует его для отправки отчетов после автоматической перезагрузки.
    'Особенность: Сам по себе не инициирует перезагрузку, но связан с автоматическим процессом.
    'If regExists(checkExistance("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting\")) Then 
    '    g_RebootReason = "Windows Update post-reboot reporting pending (HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting)"
    '    isRebootPending = True
    'End if

    
    'Что проверяет: Сигнал от установщика Windows (например, при обновлении ОС).
    'Когда требует перезагрузки: Если ключ существует.
    'Типичные причины:
    '    Установка Windows с DVD/USB.
    '    Крупное обновление (например, с Windows 10 до 11).
    If regExists("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\DVDRebootSignal") Then 
        g_RebootReason = "DVD reboot signal detected (HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\DVDRebootSignal)"
        isRebootPending = True
    End if


    'Что проверяет: Флаг перезагрузки для Component-Based Servicing (CBS) — системы обновления компонентов Windows.
    'Когда требует перезагрузки: Если ключ существует.
    'Типичные причины:
    '    Установка/удаление компонентов Windows (например, .NET Framework, языковых пакетов).
    '    Повреждение системных файлов (SFC / DISM требует перезагрузки).
    If regExists("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") Then 
        g_RebootReason = "Component Based Servicing reboot pending (HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending)"
        isRebootPending = True
    End if


    'SKIP!!!!
    'Указывает, что перезагрузка уже начата системой для завершения установки компонентов.
    'Важно: Этот ключ появляется во время перезагрузки, а не до нее.
    'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress



    'Что проверяет: Ожидание перезагрузки для завершения присоединения к домену.
    'Когда требует перезагрузки: Если ключ существует.
    'Типичные причины:
    '    Компьютер недавно присоединился к домену Active Directory.
    If regExists("HKLM\SYSTEM\CurrentControlSet\Services\Netlogon\JoinDomain") Then 
        g_RebootReason = "Domain join operation pending (HKLM\SYSTEM\CurrentControlSet\Services\Netlogon\JoinDomain)"
        isRebootPending = True
    End if



    'Что проверяет: Количество попыток перезагрузки для завершения установки ролей сервера.
    'Когда требует перезагрузки: Если значение > 0.
    'Типичные причины:
    '    Установка/удаление ролей Windows Server (DNS, Active Directory, Hyper-V).
    If regExists("HKLM\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttempts") Then 
        g_RebootReason = "Domain join operation pending (HKLM\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttempts)"
        isRebootPending = True
    End if


    'SKIP!!!!
    'Что проверяет:
    'Этот ключ связан с регистрацией Service Principal Names (SPN) в Active Directory. Он указывает, что системе нужно избегать автоматической регистрации SPN до перезагрузки.
    'Когда требует перезагрузки:
    '    Если ключ существует (обычно со значением 1).
    '    Когда компьютер присоединен к домену, но SPN не были зарегистрированы из-за проблем с сетевыми настройками или правами.
    'Типичные причины:
    '    Присоединение к домену, где учетная запись компьютера не имеет прав на запись SPN.
    '    Конфликты SPN (например, дублирование имен служб).
    '    Проблемы с синхронизацией между локальной системой и контроллером домена.
    'Как работает:
    '    Система обнаруживает, что SPN не могут быть зарегистрированы сразу.
    '    Устанавливает флаг AvoidSpnSet=1, чтобы отложить регистрацию.
    '    После перезагрузки служба Netlogon повторно пытается зарегистрировать SPN.
    'HKLM\SYSTEM\CurrentControlSet\Services\Netlogon\AvoidSpnSet



    Dim arrSubkeys, subkey

    'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending
    'Что проверяет:
    'Этот ключ указывает на наличие пакетов обновлений компонентов Windows (CBS — Component-Based Servicing), которые были установлены, но требуют перезагрузки для завершения настройки.
    'Когда требует перезагрузки:
    '    Если ключ существует и содержит подразделы (имена пакетов).
    '    Когда CBS (например, DISM или TrustedInstaller) отложил применение изменений до перезагрузки.
    'Типичные причины:
    '    Установка или удаление системных компонентов через DISM (DISM /Online /Add-Package).
    '    Обновление встроенных функций Windows (например, .NET Framework, языковые пакеты).
    '    Повреждение хранилища компонентов (C:\Windows\WinSxS), требующее восстановления.
    'Как работает:
    '    Windows сохраняет список пакетов, которые не могут быть применены "на лету" из-за блокировки системных файлов.
    '    При перезагрузке TrustedInstaller завершает установку перед загрузкой ОС.
    arrSubkeys = RegEnumKeys("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending")
    If Not IsNull(arrSubKeys) Then
        For Each subkey In arrSubKeys
            g_RebootReason = "Component Based Servicing packages pending (HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending\"&subkey&")"
            isRebootPending = True
            Exit Function
        Next
    End If

    
    ' Что проверяет: Незавершенные установки обновлений Windows
    ' Когда требует перезагрузки: При наличии подразделов в указанном пути
    ' Типичные причины:
    '   - Скачанные, но не установленные обновления
    '   - Ошибки во время установки обновлений    Dim arrSubKeys, subkey
    arrSubkeys = RegEnumKeys("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services\Pending")
    If Not IsNull(arrSubKeys) Then
        For Each subkey In arrSubKeys
            g_RebootReason = "Pending Windows Update installation detected (HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services\Pending\"&subkey&")"
            isRebootPending = True
            Exit Function
        Next
    End If
End Function

' Вспомогательная функция для получения операций с файлами
Function GetPendingFileRenameOperations()
    Dim regValue, operations, i, result
    result = ""
    
    ' Проверяем оба ключа
    Dim regKeys, key
    regKeys = Array( _
        "PendingFileRenameOperations", _
        "PendingFileRenameOperations2" _
    )
    
    For Each key In regKeys
        If regExists("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager" & "\" & key) Then
            operations = RegGetMultiStringValue("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager",key)
            on error resume next
            If operations <> false  Then
                For i = 0 To UBound(operations) - 1 Step 2
                    If operations(i) <> "" Then
			msg operations(i)
                        If operations(i+1) = "" Then
                            result = result & "Delete: " & operations(i) & vbCrLf
                        Else
                            result = result & "Rename: " & operations(i) & " -> " & operations(i+1) & vbCrLf
                        End If
                    End If
                Next
            End If
            on error goto 0
        End If
    Next
    
    GetPendingFileRenameOperations = result
End Function