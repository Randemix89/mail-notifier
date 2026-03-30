# Mail Notifier (macOS)

Небольшое приложение для отправки **уведомлений клиентам** по списку email из CSV (1 колонка `email`).

## Возможности
- SMTP (Gmail / Mail.ru / Custom)
- Тема, текст письма
- Вложения
- Старт/стоп отправки
- Очередь + безопасная скорость (rate limit)
- Статистика и лог ошибок

## Запуск

1) Установить Python 3.10+

2) Установить зависимости:

```bash
cd mail_notifier_macos
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

3) Запустить:

```bash
python app.py
```

## Если PySide6/Qt не запускается (ошибка cocoa)
На некоторых версиях macOS Qt может не подхватить `cocoa` плагин. В этом случае запускай версию на `tkinter` (встроен в Python):

```bash
python app_tk.py
```

## GitHub (сборка Windows `.exe` в Releases)

Важно: `config.json` содержит пароли — **не добавляй его в репозиторий**. Используй `config.example.json` как шаблон.

### Публикация релиза с готовым `.exe`

Workflow `/.github/workflows/windows-release.yml` собирает Windows `.exe`, когда ты пушишь тег вида `vX.Y.Z`.

Пример:

```bash
git tag v0.1.0
git push origin v0.1.0
```

Скачать: GitHub → Releases → latest.

## Windows `.exe` через GitHub Releases (без сборки у пользователей)

Если ты выложил проект в GitHub, можно собирать готовый Windows `.exe` автоматически:

1) В репозитории есть workflow: `.github/workflows/windows-release.yml`

2) Чтобы выпустить новую версию:

- обнови файл `VERSION` (например `0.1.1`)
- создай git-тег `v0.1.1` и запушь его в GitHub

После этого в GitHub появится Release с файлами:

- `MailNotifier.exe`
- `MailNotifier-vX.Y.Z-windows.zip`

## Gmail / Mail.ru заметки
- Для Gmail обычно нужен **App Password** (если включена 2FA) или OAuth. В этом MVP используется пароль/app-password.
- SMTP Gmail: `smtp.gmail.com:587` (STARTTLS)
- SMTP Mail.ru: `smtp.mail.ru:587` (STARTTLS)

## CSV формат
- Поддерживается CSV с заголовком `email` **или** без заголовка (тогда берётся первая колонка).

