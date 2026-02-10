# Система маскирования персональных и корпоративных данных в текстовых документах

Утилита для автоматического маскирования персональных данных физических лиц и данных о юридических лицах в текстовых документах с использованием Microsoft Presidio.

## 1. Возможности

- **Поддерживаемые форматы входных файлов**: `txt`, `docx`, `doc`  
- **Сохранение структуры папок** при обработке.  
- **Маскирование сущностей** с помощью Microsoft Presidio:
  - персональные данные: ФИО, телефоны, email, адреса, ИНН, СНИЛС и др.;
  - данные юридических лиц: названия организаций, ИНН/КПП/ОГРН, банковские реквизиты и др.
- **Токены маскирования**: вида `<TYPE_XXXX>`, например:
  - `Иванов Иван Иванович` → `<PERSON_0001>`
  - `ООО "Ромашка"` → `<ORGANIZATION_0001>`
  - `+7 912 345-67-89` → `<PHONE_NUMBER_0001>`
- **Словарь демаскирования** в формате JSON:
  ```json
  {
    "<PERSON_0001>": {
      "type": "PERSON",
      "original": "Иванов Иван Иванович"
    }
  }
  ```
- **Детерминированность в рамках одного запуска**: одно и то же значение → один и тот же токен независимо от файла.
- **Логирование**: количество файлов, количество сущностей по типам, ошибки.

## 2. Установка

### 2.1. Требования

- **Python** 3.10+
- ОС: **Linux / macOS / Windows**

### 2.2. Установка зависимостей

```bash
cd d:/PycharmProjects/masking
python -m venv venv
venv\Scripts\activate  # Windows
# source venv/bin/activate  # Linux/macOS

pip install -r requirements.txt
```

### 2.3. Модель SpaCy для русского языка

Presidio использует NLP‑движок. Для русского языка рекомендуются модели `ru_core_news_md` или `ru_core_news_lg`:

```bash
python -m spacy download ru_core_news_md
```

При необходимости можно изменить имя модели в функции `build_nlp_engine` файла `mask_data.py`.

## 3. Запуск

Пример запуска:

```bash
python mask_data.py ^
  --input ./input_texts ^
  --output ./masked_texts ^
  --map ./deanonymization_map.json ^
  --language ru ^
  --log-level INFO ^
  --config ./config.yaml
```

Для Linux/macOS:

```bash
python mask_data.py \
  --input ./input_texts \
  --output ./masked_texts \
  --map ./deanonymization_map.json \
  --language ru \
  --log-level INFO \
  --config ./config.yaml
```

### Обязательные параметры

- `--input` — входная папка с файлами (`txt`, `docx`, `doc`).
- `--output` — выходная папка, куда будут записаны маскированные файлы.
- `--map` — путь к JSON-файлу, куда будет записан словарь демаскирования.

### Опциональные параметры

- `--language` — язык текста (по умолчанию `ru` или значение из `config.yaml`).
- `--log-level` — уровень логирования: `INFO`, `DEBUG`, `WARNING`, `ERROR`.
- `--config` — путь к конфигурационному YAML-файлу (по умолчанию `config.yaml`).

## 4. Структура выходных данных

- **Папка маскированных файлов**: структура директорий и имена файлов совпадают с входными.
- **Словарь демаскирования** (`deanonymization_map.json`):

```json
{
  "<PERSON_0001>": {
    "type": "PERSON",
    "original": "Иванов Иван Иванович"
  },
  "<ORGANIZATION_0001>": {
    "type": "ORGANIZATION",
    "original": "ООО \"Ромашка\""
  },
  "<PHONE_NUMBER_0001>": {
    "type": "PHONE_NUMBER",
    "original": "+7 912 345-67-89"
  }
}
```

## 5. Конфигурация (`config.yaml`)

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
```

- **language** — язык текста.
- **entity_types** — список типов сущностей, которые нужно маскировать.
- **score_threshold** — минимальный порог уверенности для учёта сущности.

## 6. Добавление новых recognizers

Кастомные recognizers добавляются в функции `build_analyzer` (`mask_data.py`):

1. **Определить шаблон (Pattern)**:

   ```python
   custom_pattern = Pattern(
       name="custom_ru",
       regex=r"ВАШ_РЕГУЛЯРНЫЙ_ШАБЛОН",
       score=0.6,
   )
   ```

2. **Создать `PatternRecognizer`**:

   ```python
   custom_recognizer = PatternRecognizer(
       supported_entity="CUSTOM_ENTITY",
       patterns=[custom_pattern],
   )
   ```

3. **Зарегистрировать его в `RecognizerRegistry`**:

   ```python
   registry.add_recognizer(custom_recognizer)
   ```

4. **Добавить новый тип сущности в `config.yaml`**:

   ```yaml
   entity_types:
     - CUSTOM_ENTITY
   ```

После этого новый тип будет маскироваться токенами вида `<CUSTOM_ENTITY_0001>`, `<CUSTOM_ENTITY_0002>` и т.д.

## 7. Пример входных и выходных данных

### Вход (`input_texts/example.txt`)

```text
Иванов Иван Иванович проживает по адресу: г. Москва, ул. Ленина, д. 10.
Телефон: +7 912 345-67-89, email: ivanov@example.com.
ООО "Ромашка" ИНН 7701234567, р/с 40702810900000000001.
```

### Выход (`masked_texts/example.txt`)

```text
<PERSON_0001> проживает по адресу: <LOCATION_0001>.
Телефон: <PHONE_NUMBER_0001>, email: <EMAIL_ADDRESS_0001>.
<ORGANIZATION_0001> <ORG_INN_0001>, р/с <BANK_ACCOUNT_0001>.
```

### Словарь (`deanonymization_map.json`)

```json
{
  "<PERSON_0001>": {
    "type": "PERSON",
    "original": "Иванов Иван Иванович"
  },
  "<LOCATION_0001>": {
    "type": "LOCATION",
    "original": "г. Москва, ул. Ленина, д. 10"
  },
  "<PHONE_NUMBER_0001>": {
    "type": "PHONE_NUMBER",
    "original": "+7 912 345-67-89"
  },
  "<EMAIL_ADDRESS_0001>": {
    "type": "EMAIL_ADDRESS",
    "original": "ivanov@example.com"
  },
  "<ORGANIZATION_0001>": {
    "type": "ORGANIZATION",
    "original": "ООО \"Ромашка\""
  },
  "<ORG_INN_0001>": {
    "type": "ORG_INN",
    "original": "7701234567"
  },
  "<BANK_ACCOUNT_0001>": {
    "type": "BANK_ACCOUNT",
    "original": "40702810900000000001"
  }
}
```

## 8. Обработка ошибок и отчётность

- Некорректные или не читаемые файлы **не прерывают** общий процесс.
- Об ошибках чтения/анализа/записи файлов пишется в лог (`ERROR`).
- В конце работы в лог выводится:
  - число успешно обработанных файлов;
  - число ошибок;
  - количество найденных сущностей по каждому типу.

## 9. Ограничения решения

- Качество распознавания зависит от:
  - используемой модели SpaCy для русского языка;
  - встроенных и кастомных recognizers Presidio;
  - качества и структуры исходного текста.
- Формат `.doc` поддерживается через пакет `textract`; при его отсутствии такие файлы будут пропускаться с ошибкой в логе.
- Форматирование сложных `.docx` документов может быть упрощено (маскирование выполняется на уровне текста абзацев).
- Детерминированность гарантируется **в рамках одного запуска**; между разными запусками токены могут иметь другие номера.

