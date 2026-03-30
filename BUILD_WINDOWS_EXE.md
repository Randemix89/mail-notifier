# Сборка `.exe` на Windows

Этот проект можно запускать на Windows как обычную программу (двойным кликом), если собрать `.exe` через PyInstaller.

## Быстрый способ (рекомендуется)

1) Установи **Python 3.10+** с сайта python.org  
   В установщике поставь галочку **Add Python to PATH**.

2) В папке проекта запусти файл:

- `build_windows_exe.bat`

3) Готовый файл появится здесь:

- `dist\MailNotifier\MailNotifier.exe`

## Если Python на Windows не установлен (portable, без установки)

Можно собрать `.exe` вообще без установки Python — через PowerShell-скрипт, который сам скачает portable WinPython.

1) Открой PowerShell в папке проекта.

2) Запусти:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_windows_exe.ps1
```

3) Готовый файл:

- `dist\MailNotifier\MailNotifier.exe`

Лог сборки:

- `build_windows_exe_ps.log`

## Примечания

- Для Windows рекомендуется собирать из `app_tk.py` (tkinter), он менее капризный, чем Qt.
- `config.json` будет создаваться/обновляться рядом с запуском приложения.

