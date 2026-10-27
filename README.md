# OpenChamber Office Preview

Read-only Excel, Word and PowerPoint previews inside OpenChamber Files.

![Excel preview](https://raw.githubusercontent.com/d0kur0/openchamber-office-preview/main/docs/images/excel.png)

Локальный просмотрщик Excel, Word и PowerPoint внутри Files. Исходный файл не
редактируется. Нет фонового сервиса, запросов к внешним сайтам, прав на папки
или доступа к разговорам. Данные выбранного файла передаёт сам OpenChamber.

## Возможности

- XLSX/XLSM: листы, значения, числовые форматы, оформление, объединения, формула
  выбранной ячейки, основные диаграммы по сохранённым данным.
- DOCX: страницы, заголовки, колонтитулы, таблицы и встроенные изображения.
- PPTX: слайды, миниатюры, переключение стрелками, масштаб.
- PDF, CSV и TSV остаются штатным просмотрщикам OpenChamber.

Формулы Excel не пересчитываются, используются сохранённые значения. Нет запуска
макросов, редактирования, поддержки старых DOC/XLS/PPT и защищённых паролем файлов.
Сложные SmartArt, анимации, нестандартные шрифты и точное печатное совпадение
с Microsoft Office не гарантируются. Ограничения: 20 МБ файла, 100 МБ распакованных
данных, до 2000 видимых строк/80 столбцов листа; дополнительные строки раскрываются кнопкой.

## Установка

Требуется OpenChamber >=2.5.2.

1. Откройте Settings → Extensions → Add extension.
2. Вставьте `https://github.com/d0kur0/openchamber-office-preview` и нажмите Add.
3. Откройте XLSX, DOCX или PPTX в Files.

Node.js и npm для установки в OpenChamber не нужны: сборка включена в репозиторий.
Можно установить ZIP из [Releases](https://github.com/d0kur0/openchamber-office-preview/releases)
или выбрать локальную папку клона. Релизный ZIP содержит папку `office-preview`
с манифестом и готовой сборкой.

Пакет также распространяется в npm как `openchamber-office-preview`.
После `npm install openchamber-office-preview` выберите абсолютный путь
`node_modules/openchamber-office-preview` в Extensions. Это расширение
OpenChamber, его не нужно добавлять в `plugin` конфигурации OpenCode.

Файлы XLSX/DOCX/PPTX после открытия в Files используют новый просмотрщик.
Панель Office также принимает выбранный пользователем файл и drag-and-drop.
Можно отключить расширение в Settings → Extensions, штатные просмотры вернутся.

## Разработка и проверки

```powershell
npm install --legacy-peer-deps
npm run build
npm test
```

В репе сохранён собранный `panel/main.js`, OpenChamber не собирает расширения.
Зависимости и их лицензии перечислены в THIRD_PARTY_NOTICES.md, точные версии
зафиксированы package-lock.json. Ключи и пользовательские документы не нужны.

## Релизы

Workflow `release.yml` проверяет push в `main` и pull request. После успешных
проверок в main запускается semantic-release. Версии и теги вручную не меняются:

- `fix: ...` выпускает patch.
- `feat: ...` выпускает minor.
- `feat!: ...` или `BREAKING CHANGE:` выпускает major.
- `ci(release): ...` выпускает patch для изменений механизма выпуска.
- Обычные `docs:`, `test:`, `chore:` и `ci:` сами по себе релиз не создают.

Semantic-release обновляет package.json и lockfile, ведёт CHANGELOG.md,
коммитит версию и готовую сборку с `[skip ci]`, создаёт тег и публикует npm
и GitHub Release с ZIP, TGZ и SHA256SUMS.txt. Одновременные выпуски сериализованы.
В pull request выполняются только проверки. Ручной запуск workflow на main
повторяет анализ коммитов; без новых изменений для выпуска версия не создаётся.

Для npm используется trusted publisher GitHub Actions: owner `d0kur0`,
repository `openchamber-office-preview`, workflow `release.yml`.
Его необходимо настроить в npm package settings после первой ручной публикации.
Workflow требует разрешения на прямой `npm publish`; постоянный NPM_TOKEN
в GitHub Secrets не требуется. [Документация npm](https://docs.npmjs.com/trusted-publishers/).

До первого автоматического выпуска нужны ручная публикация владельцем и
настройка trusted publisher. Workflow останавливается до создания новых тегов,
если пакет ещё не существует. Первая публикация: `npm login`, затем
`npm publish --access public`; npm требует 2FA или подходящий granular token.

## Примеры

На снимках только синтетические данные.

![Word preview](https://raw.githubusercontent.com/d0kur0/openchamber-office-preview/main/docs/images/word.png)
![PowerPoint preview](https://raw.githubusercontent.com/d0kur0/openchamber-office-preview/main/docs/images/powerpoint.png)

Независимое расширение сообщества. Не является продуктом Microsoft или
официальным расширением команды OpenChamber. Лицензия MIT; лицензии библиотек
сохранены в `licenses/` и THIRD_PARTY_NOTICES.md.
