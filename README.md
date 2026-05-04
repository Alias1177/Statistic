# Энергомонитор

Дашборд для анализа энергопотребления. Загружаешь `.docx` с расчётами — получаешь интерактивные графики, сравнение по объектам, динамику по дням и структуру нагрузок. Парсинг Word-документа полностью на клиенте, без бэкенда.

## Стек
- React 18 (UMD, без сборки React)
- esbuild — компиляция JSX и минификация
- JSZip + DOMParser — чтение `.docx` в браузере
- SVG-графики собственного производства, IntersectionObserver-анимации

## Структура
```
.
├── index.html               # шаблон, __BUILD__ → хеш билда
├── src/
│   ├── app.jsx              # вся логика приложения
│   └── styles.css           # стили
├── public/
│   ├── favicon.svg
│   └── templates/           # примеры .docx для загрузки
│       ├── template-piti.docx
│       └── template-mall.docx
├── build.mjs                # esbuild + копирование public/
├── package.json
├── vercel.json              # cache headers, output = dist/
└── dist/                    # результат сборки (gitignore)
```

## Локально
```bash
npm install
npm run build      # → dist/
npm run preview    # http://localhost:5173
```

## Деплой на Vercel

### Через CLI
```bash
npm i -g vercel
vercel              # первый раз — настройка проекта
vercel --prod       # прод-деплой
```
Vercel сам прочитает `vercel.json`, выполнит `npm run build` и опубликует `dist/`.

### Через GitHub
1. Запушь репозиторий на GitHub
2. На vercel.com → New Project → Import репозиторий
3. Framework Preset: **Other** (всё уже настроено в `vercel.json`)
4. Deploy

## Что оптимизировано
- **Без babel-standalone в рантайме** — JSX компилируется на этапе билда (–250 KB загрузки)
- **Хеши билда в URL** (`?v=abc123`) — иммутабельный кэш + автоматический cache-bust при деплое
- **Cache-Control: immutable** на JS/CSS/SVG, `must-revalidate` на HTML
- **Pinned версии** React/JSZip с CDN + `preconnect` для быстрого TLS
- **Минификация** JS и CSS через esbuild
- **Кириллические имена убраны** из URL (template-piti.docx, template-mall.docx)
- **Безопасность**: X-Content-Type-Options, Referrer-Policy, X-Frame-Options, Permissions-Policy

## Формат входного .docx
Документ должен содержать три раздела с таблицами:
1. **Сводные данные по месяцам** — Объект | Период A | Период B | Экономия %
2. **Детализация по дням** — два подраздела с таблицами День | Потребление
3. **Разбивка по типам нагрузки** — два подраздела с таблицами Тип | Потребление | Доля

Парсер находит таблицы по форме (числовые колонки, формат даты), а не по названиям, так что заголовки могут отличаться. Название объекта вытаскивается из заголовка H1 («Расчёт … — ТЦ "Меридиан"») или из имени файла.
