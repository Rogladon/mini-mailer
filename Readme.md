# Bulk Mailer Electron App

Кросс‑платформенное (Windows‑фокус) мини‑приложение для массовой отправки разовых писем через SMTP Яндекса из Excel‑таблицы. Написано на **React + TypeScript** (renderer) и **Electron/Node** (main).

## Возможности

* Чтение получателей из `.xlsx` (колонки `email`, `name` обрабатываются автоматически)
* Подстановка имени в тему/тело письма
* Пауза 20‑40 с между отправкой (джиттер) для снижения риска «Спам»
* Отчёт об отправке (OK/FAIL) в UI + экспорт `report.xlsx`
* Конфигурация трёх SMTP‑ящиков @yandex.ru
* Portable `.exe` (без установки Node), размер ≈ 80 МБ

## Требования

* **Node.js >= 18** (только для сборки)
* **Windows 10/11** для конечного EXE. На macOS/Linux будет `.dmg` / `AppImage`.

## Быстрый старт (режим разработки)

```bash
# 1. Клонируем
git clone https://example.com/bulk-mailer-electron.git
cd bulk-mailer-electron

# 2. Ставим зависимости
npm install   # или pnpm install / yarn

# 3. Запуск dev‑режима
npm run dev   # откроется окно Electron + React hot‑reload
```

## Сборка portable‑инсталлятора

```bash
npm run build    # создаст dist/BulkMailer‑Setup‑0.1.0.exe
```

Раздайте EXE операторам: двойной клик → установка → ярлык «Bulk Mailer». Авто‑обновления отключены.

## Создание пароля приложения Яндекса

1. passport.yandex.ru → «Пароли приложений» → «Почта»
2. Скопируйте пароль и сохраните — он потребуется в поле «Пароль».

## Структура проекта

```
├─ package.json
├─ electron.vite.config.ts           # конфиг electron‑vite
├─ tsconfig.json
├─ src/
│  ├─ main/                          # Node‑часть
│  │   ├─ main.ts
│  │   ├─ preload.ts
│  │   └─ mailer.ts
│  └─ renderer/                      # React‑часть
│      ├─ main.tsx
│      ├─ App.tsx
│      └─ index.css
└─ assets/
```

## Замена HTML‑шаблона письма

Файл `assets/template.html` содержит базовый шаблон. Используйте плейсхолдер `{{name}}` для подстановки.