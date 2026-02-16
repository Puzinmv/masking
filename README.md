# Система маскирования персональных и корпоративных данных

Утилита для автоматического маскирования чувствительных данных в документах с использованием Microsoft Presidio и кастомных распознавателей для РФ.

## Что умеет текущий релиз

- Обрабатывает форматы: `txt`, `docx`, `doc`, `pdf`, а также изображения (`png`, `jpg`, `jpeg`, `bmp`, `tif`, `tiff`, `gif`, `webp`, `heic`, `heif`, `jp2`, `j2k`, `jpx`, `jpf`, `jpm`, `jxr`, `ppm`, `pgm`, `pbm`, `pnm`, `ico`).
- Рекурсивно обходит входную директорию и сохраняет структуру папок в выходной.
- Маскирует сущности токенами вида `<TYPE_0001>`.
- Формирует `deanonymization_map.json` для обратного сопоставления токенов и оригиналов.
- Сохраняет детерминизм в рамках одного запуска: одинаковое значение получает одинаковый токен.
- Для `docx` маскирует текст в абзацах и таблицах с сохранением структуры документа.
- Для `pdf`:
  - извлекает текст из текстовых PDF;
  - при `--ocr` дополнительно распознает текст на страницах с изображениями и пишет отдельный `*.ocr.txt`.
- Для изображений маскирование работает только при флаге `--ocr`; результат пишется в `*.ocr.txt`.
- Поддерживает стоп-слова из `stopwords.yaml` (точные строки и regex), включая фильтрацию по леммам.

## Поддерживаемые типы сущностей

По умолчанию в `config.yaml`:

- `PERSON`
- `PHONE_NUMBER`
- `EMAIL_ADDRESS`
- `ORGANIZATION`
- `LOCATION`
- `INN`
- `SNILS`
- `OGRN`
- `BANK_ACCOUNT`
- `ORG_INN`
- `ORG_KPP`
- `ORG_OGRN`

Кроме стандартных recognizers Presidio используются кастомные regex-распознаватели для российских форматов (ИНН/КПП/ОГРН/СНИЛС/телефон/email/адрес и др.).

## Установка

Требования:

- Python 3.10+
- Windows / Linux / macOS

Установка зависимостей:

```bash
python -m venv venv
venv\Scripts\activate  # Windows
# source venv/bin/activate  # Linux/macOS

pip install -r requirements.txt
```

Установка модели SpaCy (для русского языка):

```bash
python -m spacy download ru_core_news_md
```

Опционально для OCR (если нужен флаг `--ocr`) установите:

```bash
pip install paddleocr Pillow numpy pymupdf
```

## Быстрый запуск

```bash
python mask_data.py \
  --input ./input_texts \
  --output ./masked_texts \
  --map ./deanonymization_map.json \
  --config ./config.yaml \
  --log-level INFO
```

На Windows можно использовать перенос строк через `^`.

## Параметры CLI

- `--input` путь к входной папке (по умолчанию `./input_texts`)
- `--output` путь к выходной папке (по умолчанию `./masked_texts`)
- `--map` путь к JSON-карте деанонимизации (по умолчанию `./deanonymization_map.json`)
- `--config` путь к YAML-конфигу (по умолчанию `config.yaml`)
- `--language` язык анализа (если указан, переопределяет `language` из `config.yaml`)
- `--log-level` один из `DEBUG`, `INFO`, `WARNING`, `ERROR`
- `--ocr` включить OCR для изображений и сканированных PDF

## Конфигурация (`config.yaml`)

Используются поля:

- `language`
- `entity_types`
- `score_threshold`
- `logging.level` (только как значение в файле; реальный уровень задается через `--log-level`)

Пример:

```yaml
language: ru

entity_types:
  - PERSON
  - PHONE_NUMBER
  - EMAIL_ADDRESS
  - ORGANIZATION
  - LOCATION
  - INN
  - SNILS
  - OGRN
  - BANK_ACCOUNT
  - ORG_INN
  - ORG_KPP
  - ORG_OGRN

score_threshold: 0.35

logging:
  level: INFO
```

## Стоп-слова (`stopwords.yaml`)

Стоп-слова позволяют исключать ложные срабатывания глобально, для любых типов сущностей.

Поддерживается:

- точное совпадение строки;
- regex через префикс `re:`.

Пример:

```yaml
- "СОГЛАСОВАНО"
- "re:^([А-ЯЁ]{2})\\.\\d+$"
- "КАТЕГОРИЯ"
```

Если `stopwords.yaml` отсутствует, скрипт продолжает работу с пустым списком исключений.

## Что получается на выходе

- Маскированные файлы в `--output` с исходной структурой директорий.
- `deanonymization_map.json` с токенами и оригинальными значениями.
- Для OCR-результатов: отдельные файлы `*.ocr.txt`.

## Ограничения

- Качество распознавания зависит от модели SpaCy и качества исходного текста.
- Для `doc` нужен `textract`, иначе файл пропускается с ошибкой в логе.
- Выходной `pdf` формируется как простой текстовый PDF без сохранения исходной верстки.
- Нумерация токенов не фиксируется между разными запусками.
