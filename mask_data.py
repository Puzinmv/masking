import argparse
import json
import logging
import os
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Iterable, Optional

import yaml
from presidio_analyzer import AnalyzerEngine, RecognizerRegistry, PatternRecognizer, Pattern, RecognizerResult
from presidio_analyzer.nlp_engine import NlpEngineProvider
from presidio_analyzer.nlp_engine import NlpEngine

from docx import Document


LOGGER = logging.getLogger("mask_data")


SUPPORTED_EXTENSIONS = {".txt", ".docx", ".doc"}

# Глобальный словарь "стоп-слов" по типам сущностей, загружается из YAML.
# Ключ: тип сущности (PERSON, ORGANIZATION, ...), значение: множество строк (в верхнем регистре).
ENTITY_STOPWORDS: Dict[str, set[str]] = {}

# Кеш лемм и NLP-движок для морфологической фильтрации стоп-слов.
STOPWORD_LEMMA_CACHE: Dict[str, set[str]] = {}
STOPWORDS_NLP_ENGINE: Optional[NlpEngine] = None
STOPWORDS_LANGUAGE: str = "ru"

# Heuristics for expanding organization abbreviations like "ФГБОУ".
ORG_ABBR_RE = re.compile(r"^[A-ZА-ЯЁ]{2,10}(?:\s+[A-ZА-ЯЁ]{2,10}){0,3}$")
ORG_UPPER_WORD_RE = re.compile(r"^[A-ZА-ЯЁ]{2,10}$")
ORG_TITLE_WORD_RE = re.compile(r"^[A-ZА-ЯЁ][a-zа-яё]+(?:-[A-ZА-ЯЁ][a-zа-яё]+)*$")


@dataclass
class MaskingConfig:
    language: str
    entity_types: List[str]
    score_threshold: float
    custom_dictionary: Dict[str, List[str]]


def load_config(path: Optional[Path]) -> MaskingConfig:
    """
    Загрузка конфигурации из YAML. При отсутствии файла используются значения по умолчанию.
    """
    default = {
        "language": "ru",
        "entity_types": [
            "PERSON",
            "PHONE_NUMBER",
            "EMAIL_ADDRESS",
            "ORGANIZATION",
            "LOCATION",
            "INN",
            "SNILS",
            "OGRN",
            "BANK_ACCOUNT",
            "ORG_INN",
            "ORG_KPP",
            "ORG_OGRN",
        ],
        "score_threshold": 0.35,
        "custom_dictionary": {},
    }

    if path is None or not path.is_file():
        cfg = default
    else:
        with path.open("r", encoding="utf-8") as f:
            file_cfg = yaml.safe_load(f) or {}
        cfg = {**default, **file_cfg}

    raw_custom_dict = cfg.get("custom_dictionary") or {}
    custom_dictionary: Dict[str, List[str]] = {}
    if isinstance(raw_custom_dict, dict):
        for ent_type, values in raw_custom_dict.items():
            if not isinstance(values, list):
                continue
            ent_type_u = str(ent_type).upper()
            normalized = [str(v).strip() for v in values if str(v).strip()]
            if normalized:
                custom_dictionary[ent_type_u] = normalized

    return MaskingConfig(
        language=cfg["language"],
        entity_types=list(cfg["entity_types"]),
        score_threshold=float(cfg["score_threshold"]),
        custom_dictionary=custom_dictionary,
    )


def load_stopwords(path: Optional[Path]) -> Dict[str, set[str]]:
    """
    Загрузка стоп-слов из YAML-файла.
    Формат:
      ORGANIZATION:
        - "СОГЛАСОВАНО"
        - "УТВЕРЖДАЮ"
      PERSON:
        - "КАТЕГОРИЯ"
        - "КАТЕГОРИЯ III"
    Все ключи и значения приводятся к верхнему регистру.
    """
    if path is None:
        # По умолчанию пробуем файл stopwords.yaml в рабочей директории.
        path = Path("stopwords.yaml").resolve()

    if not path.is_file():
        LOGGER.info("Файл стоп-слов не найден: %s (будет использован пустой список)", path)
        return {}

    try:
        with path.open("r", encoding="utf-8") as f:
            raw = yaml.safe_load(f) or {}
    except Exception as exc:  # noqa: BLE001
        LOGGER.error("Ошибка чтения файла стоп-слов %s: %s", path, exc)
        return {}

    if not isinstance(raw, dict):
        LOGGER.error("Некорректный формат stopwords YAML (ожидается mapping): %s", path)
        return {}

    result: Dict[str, set[str]] = {}
    for ent_type, values in raw.items():
        if not isinstance(values, list):
            continue
        et = str(ent_type).upper()
        normalized = {str(v).upper() for v in values}
        if normalized:
            result[et] = normalized

    LOGGER.info("Загружено стоп-слов: %d типов из %s", len(result), path)
    return result


def _get_lemmas_for_text(text: str) -> set[str]:
    """
    Получить множество лемм для строки, используя тот же NLP-движок, что и Presidio.
    Леммы используются для того, чтобы одно стоп-слово (например, "ПОЛОЖЕНИЕ")
    отбрасывало все его склонения ("ПОЛОЖЕНИЯ", "ПОЛОЖЕНИИ" и т.п.).
    """
    original = text.strip()
    if not original:
        return set()

    cache_key = original.upper()

    cached = STOPWORD_LEMMA_CACHE.get(cache_key)
    if cached is not None:
        return cached

    if STOPWORDS_NLP_ENGINE is None:
        # NLP-движок ещё не инициализирован или недоступен —
        # возвращаем пустое множество, чтобы не ломать основной поток.
        lemmas: set[str] = set()
        STOPWORD_LEMMA_CACHE[cache_key] = lemmas
        return lemmas

    try:
        # Для корректной морфологии ру‑модель лучше работает с нижним регистром.
        doc = STOPWORDS_NLP_ENGINE.process_text(original.lower(), STOPWORDS_LANGUAGE)
    except Exception:  # noqa: BLE001
        lemmas = set()
        STOPWORD_LEMMA_CACHE[cache_key] = lemmas
        return lemmas

    # process_text() в Presidio возвращает NlpArtifacts, а не сам spaCy-doc.
    # Используем его поля lemmas/tokens.
    lemmas: set[str] = set()
    doc_lemmas = getattr(doc, "lemmas", None)
    doc_tokens = getattr(doc, "tokens", None)

    if doc_lemmas:
        lemmas.update(str(l).upper() for l in doc_lemmas if str(l).strip())
    elif doc_tokens:
        # Fallback: если леммы недоступны, используем сами токены.
        lemmas.update(str(t).upper() for t in doc_tokens if str(t).strip())

    STOPWORD_LEMMA_CACHE[cache_key] = lemmas
    return lemmas


def _is_in_stopwords(ent_type: str, original_value: str) -> bool:
    """
    Проверяет, является ли распознанная сущность стоп-словом.

    Логика:
    1) сначала точное сравнение строки с элементами из ENTITY_STOPWORDS;
    2) затем сравнение по леммам, чтобы одно слово в стоп-списке
       отбрасывало все его формы.
    """
    ent_type_u = ent_type.upper()
    stopwords = ENTITY_STOPWORDS.get(ent_type_u)
    if not stopwords:
        return False

    value_u = original_value.upper().strip()
    if not value_u:
        return False

    # Точное совпадение по строке.
    if value_u in stopwords:
        return True

    # Сравнение по леммам.
    value_lemmas = _get_lemmas_for_text(original_value)
    if not value_lemmas:
        return False

    for sw in stopwords:
        sw_lemmas = _get_lemmas_for_text(sw)
        if value_lemmas & sw_lemmas:
            return True

    return False


def build_nlp_engine(language: str) -> NlpEngine:
    """
    Создаёт NLP-движок для Presidio на базе SpaCy.
    Для русского языка предполагается установленная модель ru_core_news_md/ru_core_news_lg.
    """
    # Можно переопределить через переменные окружения или конфиг при необходимости.
    nlp_configuration = {
        "nlp_engine_name": "spacy",
        "models": [
            {"lang_code": language, "model_name": f"{language}_core_news_lg"},
        ],
    }
    provider = NlpEngineProvider(nlp_configuration=nlp_configuration)
    return provider.create_engine()


def _debug_log(msg: str, data: dict) -> None:
    # #region agent log
    try:
        log_path = Path(__file__).resolve().parent / ".cursor" / "debug.log"
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with open(log_path, "a", encoding="utf-8") as f:
            import json, time
            f.write(json.dumps({"message": msg, "data": data, "timestamp": int(time.time() * 1000), "runId": "post-fix"}, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion

def build_analyzer(config: MaskingConfig) -> AnalyzerEngine:
    """
    Создаёт AnalyzerEngine Presidio, регистрируя кастомные recognizers для РФ.
    """
    # #region agent log
    _debug_log("build_analyzer entry", {"config_language": config.language, "hypothesisId": "C"})
    # #endregion
    nlp_engine = build_nlp_engine(config.language)
    registry = RecognizerRegistry(supported_languages=[config.language])
    # #region agent log
    _debug_log("after RecognizerRegistry()", {"registry_supported_languages": getattr(registry, "supported_languages", "N/A"), "hypothesisId": "A"})
    # #endregion
    registry.load_predefined_recognizers()
    # #region agent log
    _debug_log("before AnalyzerEngine", {"registry_supported_languages": getattr(registry, "supported_languages", "N/A"), "analyzer_langs": [config.language], "hypothesisId": "B"})
    # #endregion

    # Кастомные recognizers для РФ.

    # ИНН (10 или 12 цифр).
    inn_pattern = Pattern(name="inn_ru", regex=r"\b\d{10}(\d{2})?\b", score=0.6)
    inn_recognizer = PatternRecognizer(
        supported_entity="INN",
        patterns=[inn_pattern],
        supported_language=config.language,
    )
    registry.add_recognizer(inn_recognizer)

    # ОГРН (13 цифр).
    ogrn_pattern = Pattern(name="ogrn_ru", regex=r"\b\d{13}\b", score=0.6)
    ogrn_recognizer = PatternRecognizer(
        supported_entity="OGRN",
        patterns=[ogrn_pattern],
        supported_language=config.language,
    )
    registry.add_recognizer(ogrn_recognizer)

    # СНИЛС (формат 000-000-000 00).
    snils_pattern = Pattern(
        name="snils_ru",
        regex=r"\b\d{3}-\d{3}-\d{3}\s\d{2}\b",
        score=0.7,
    )
    snils_recognizer = PatternRecognizer(
        supported_entity="SNILS",
        patterns=[snils_pattern],
        supported_language=config.language,
    )
    registry.add_recognizer(snils_recognizer)

    # Банковский счёт (простейший шаблон для 20-значного счёта РФ).
    bank_acc_pattern = Pattern(
        name="bank_account_ru",
        regex=r"\b\d{20}\b",
        score=0.5,
    )
    bank_acc_recognizer = PatternRecognizer(
        supported_entity="BANK_ACCOUNT",
        patterns=[bank_acc_pattern],
        supported_language=config.language,
    )
    registry.add_recognizer(bank_acc_recognizer)

    # ИНН/КПП/ОГРН организаций как отдельные типы (примеры, можно доработать).
    org_inn_pattern = Pattern(
        name="org_inn_ru",
        regex=r"\b\d{10}\b",
        score=0.55,
    )
    org_inn_recognizer = PatternRecognizer(
        supported_entity="ORG_INN",
        patterns=[org_inn_pattern],
        supported_language=config.language,
    )
    registry.add_recognizer(org_inn_recognizer)

    org_kpp_pattern = Pattern(
        name="org_kpp_ru",
        regex=r"\b\d{9}\b",
        score=0.55,
    )
    org_kpp_recognizer = PatternRecognizer(
        supported_entity="ORG_KPP",
        patterns=[org_kpp_pattern],
        supported_language=config.language,
    )
    registry.add_recognizer(org_kpp_recognizer)

    org_ogrn_pattern = Pattern(
        name="org_ogrn_ru",
        regex=r"\b\d{13}\b",
        score=0.55,
    )
    org_ogrn_recognizer = PatternRecognizer(
        supported_entity="ORG_OGRN",
        patterns=[org_ogrn_pattern],
        supported_language=config.language,
    )
    registry.add_recognizer(org_ogrn_recognizer)

    # Дополнительные recognizers для e‑mail, телефонов и ИНН/КПП с явными метками.
    email_pattern = Pattern(
        name="email_ru_simple",
        regex=r"\b[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9.-]+\b",
        score=0.8,
    )
    email_recognizer = PatternRecognizer(
        supported_entity="EMAIL_ADDRESS",
        patterns=[email_pattern],
        supported_language=config.language,
    )
    registry.add_recognizer(email_recognizer)

    # Полный почтовый адрес РФ как единая LOCATION.
    # Пример: "640004, Курганская область, г. Курган, ул. Панфилова 22"
    # Логика: 6‑значный индекс, затем запятая и остальная часть строки до конца
    # строки или до точки с запятой.
    full_address_pattern = Pattern(
        name="ru_full_postal_address",
        regex=r"\b\d{6},[^\n;]+",
        score=0.9,
    )
    full_address_recognizer = PatternRecognizer(
        supported_entity="LOCATION",
        patterns=[full_address_pattern],
        supported_language=config.language,
    )
    registry.add_recognizer(full_address_recognizer)

    # ФИО в формате инициалов и фамилии: "Д.И. Анучин", "О.П. Чувардин",
    # а также варианты со пробелами между инициалами: "Д. И. Анучин".
    initials_person_pattern = Pattern(
        name="person_ru_initials_surname",
        regex=r"\b[А-ЯЁ]\.\s*[А-ЯЁ]\.\s*[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?",
        score=0.8,
    )
    initials_person_recognizer = PatternRecognizer(
        supported_entity="PERSON",
        patterns=[initials_person_pattern],
        supported_language=config.language,
    )
    registry.add_recognizer(initials_person_recognizer)

    phone_patterns = [
        Pattern(
            name="phone_ru_international",
            regex=r"(?:\+7|8)\s*(?:\(\d{3}\)|\d{3})[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}",
            score=0.8,
        ),
        Pattern(
            name="phone_ru_generic",
            regex=r"\b\+?\d[\d\-\s()]{9,}\d\b",
            score=0.5,
        ),
    ]
    phone_recognizer = PatternRecognizer(
        supported_entity="PHONE_NUMBER",
        patterns=phone_patterns,
        supported_language=config.language,
    )
    registry.add_recognizer(phone_recognizer)

    org_inn_labeled_pattern = Pattern(
        name="org_inn_labeled_ru",
        regex=r"\bИНН\s+(?:\d{10}|\d{12})\b",
        score=0.8,
    )
    org_inn_labeled_recognizer = PatternRecognizer(
        supported_entity="ORG_INN",
        patterns=[org_inn_labeled_pattern],
        supported_language=config.language,
    )
    registry.add_recognizer(org_inn_labeled_recognizer)

    org_kpp_labeled_pattern = Pattern(
        name="org_kpp_labeled_ru",
        regex=r"\bКПП\s+\d{9}\b",
        score=0.8,
    )
    org_kpp_labeled_recognizer = PatternRecognizer(
        supported_entity="ORG_KPP",
        patterns=[org_kpp_labeled_pattern],
        supported_language=config.language,
    )
    registry.add_recognizer(org_kpp_labeled_recognizer)

    analyzer = AnalyzerEngine(
        nlp_engine=nlp_engine,
        registry=registry,
        supported_languages=[config.language],
    )
    # Инициализируем NLP-движок для морфологической фильтрации стоп-слов.
    global STOPWORDS_NLP_ENGINE, STOPWORDS_LANGUAGE, STOPWORD_LEMMA_CACHE
    STOPWORDS_NLP_ENGINE = nlp_engine
    STOPWORDS_LANGUAGE = config.language
    STOPWORD_LEMMA_CACHE.clear()
    return analyzer


def get_files(input_dir: Path) -> Iterable[Path]:
    """
    Рекурсивный обход входной директории с фильтрацией по поддерживаемым расширениям.
    """
    for root, _, files in os.walk(input_dir):
        root_path = Path(root)
        for name in files:
            ext = Path(name).suffix.lower()
            if ext in SUPPORTED_EXTENSIONS:
                yield root_path / name


def read_text_from_file(path: Path) -> str:
    """
    Чтение текста из файла в зависимости от расширения.
    Для .doc используется textract (при наличии), иначе логируем ошибку.
    """
    suffix = path.suffix.lower()

    if suffix == ".txt":
        with path.open("r", encoding="utf-8", errors="ignore") as f:
            return f.read()

    if suffix == ".docx":
        doc = Document(str(path))
        paragraphs_text = "\n".join(p.text for p in doc.paragraphs)
        tables_text = ""
        if doc.tables:
            # Собираем текст из таблиц только для статистики (пока не используем при маскировании).
            tables_text = " ".join(
                cell.text
                for table in doc.tables
                for row in table.rows
                for cell in row.cells
            )
        # #region agent log
        _debug_log(
            "docx read stats",
            {
                "file_name": path.name,
                "paragraphs": len(doc.paragraphs),
                "tables": len(doc.tables),
                "paragraphs_text_len": len(paragraphs_text),
                "tables_text_len": len(tables_text),
                "hypothesisId": "H1",
            },
        )
        # #endregion
        return paragraphs_text

    if suffix == ".doc":
        try:
            import textract  # type: ignore
        except ImportError:
            LOGGER.error("Для обработки .doc требуется пакет textract. Файл %s будет пропущен.", path)
            raise

        text = textract.process(str(path))
        return text.decode("utf-8", errors="ignore")

    raise ValueError(f"Неподдерживаемое расширение файла: {path}")


def write_text_to_file(path: Path, text: str, original_suffix: str) -> None:
    """
    Запись маскированного текста в файл.
    Для .docx создаётся простой документ Word, для .txt — текстовый файл.
    Для .doc по умолчанию пишем .txt, сохраняя расширение .doc в имени.
    """
    suffix = original_suffix.lower()

    if suffix == ".txt" or suffix == ".doc":
        # Для .doc создаём текстовый файл с тем же именем.
        with path.open("w", encoding="utf-8") as f:
            f.write(text)
        return

    if suffix == ".docx":
        # #region agent log
        _debug_log(
            "docx write stats",
            {
                "file_name": path.name,
                "lines": len(text.splitlines()),
                "text_len": len(text),
                "hypothesisId": "H2",
            },
        )
        # #endregion
        doc = Document()
        for line in text.splitlines():
            doc.add_paragraph(line)
        doc.save(str(path))
        return

    # На всякий случай, как текст.
    with path.open("w", encoding="utf-8") as f:
        f.write(text)


class TokenGenerator:
    """
    Генератор токенов вида <TYPE_XXXX>, детерминированный в рамках запуска.
    """

    def __init__(self) -> None:
        self._counters: Dict[str, int] = defaultdict(int)
        self._mapping: Dict[Tuple[str, str], str] = {}

    @property
    def mapping(self) -> Dict[str, Dict[str, str]]:
        """
        Возвращает словарь для сохранения в JSON:
        {
          "<PERSON_0001>": {"type": "PERSON", "original": "Иванов Иван Иванович"},
          ...
        }
        """
        result: Dict[str, Dict[str, str]] = {}
        for (ent_type, original), token in self._mapping.items():
            result[token] = {"type": ent_type, "original": original}
        return result

    def get_token(self, ent_type: str, original: str) -> str:
        key = (ent_type, original)
        if key in self._mapping:
            return self._mapping[key]

        self._counters[ent_type] += 1
        num = self._counters[ent_type]
        token = f"<{ent_type.upper()}_{num:04d}>"
        self._mapping[key] = token
        return token


def _looks_like_org_word(word: str) -> bool:
    if ORG_UPPER_WORD_RE.match(word):
        return True
    if ORG_TITLE_WORD_RE.match(word):
        return True
    return False


def _expand_organization_span(text: str, start: int, end: int) -> Tuple[int, int]:
    """
    Expand short organization abbreviations to the right to capture full names.
    Example: "ФГБОУ ВО "..." "..."".
    """
    original = text[start:end].strip()
    if not ORG_ABBR_RE.match(original):
        return start, end

    i = end
    new_end = end

    while i < len(text):
        # skip whitespace
        while i < len(text) and text[i].isspace():
            i += 1
        if i >= len(text):
            break

        ch = text[i]
        if ch in ("\"", "«"):
            close = "»" if ch == "«" else "\""
            j = text.find(close, i + 1)
            if j == -1:
                break
            new_end = j + 1
            i = j + 1
            continue

        if ch in "([;:,.])" or ch in "–—-":
            break

        j = i
        while j < len(text) and (text[j].isalnum() or text[j] in "-/"):
            j += 1
        word = text[i:j]
        if _looks_like_org_word(word):
            new_end = j
            i = j
            continue
        break

    return start, new_end


def _expand_organization_results(
    text: str, results: List[RecognizerResult]
) -> List[RecognizerResult]:
    expanded: List[RecognizerResult] = []
    for res in results:
        if res.entity_type.upper() == "ORGANIZATION":
            new_start, new_end = _expand_organization_span(text, res.start, res.end)
            if new_start != res.start or new_end != res.end:
                res = RecognizerResult(
                    entity_type=res.entity_type,
                    start=new_start,
                    end=new_end,
                    score=res.score,
                )
        expanded.append(res)
    return expanded


def _merge_overlapping_organizations(
    results: List[RecognizerResult],
) -> List[RecognizerResult]:
    orgs = [r for r in results if r.entity_type.upper() == "ORGANIZATION"]
    if not orgs:
        return results

    orgs_sorted = sorted(
        orgs, key=lambda r: (r.start, -(r.end - r.start), -r.score)
    )
    merged: List[RecognizerResult] = []
    for res in orgs_sorted:
        if not merged:
            merged.append(res)
            continue
        last = merged[-1]
        if res.start < last.end:
            len_res = res.end - res.start
            len_last = last.end - last.start
            if len_res > len_last or (len_res == len_last and res.score > last.score):
                merged[-1] = res
            continue
        merged.append(res)

    others = [r for r in results if r.entity_type.upper() != "ORGANIZATION"]
    return merged + others


def _remove_results_inside_organizations(
    results: List[RecognizerResult],
) -> List[RecognizerResult]:
    org_spans = [(r.start, r.end) for r in results if r.entity_type.upper() == "ORGANIZATION"]
    if not org_spans:
        return results

    filtered: List[RecognizerResult] = []
    for res in results:
        if res.entity_type.upper() == "ORGANIZATION":
            filtered.append(res)
            continue
        if any(start <= res.start and res.end <= end for start, end in org_spans):
            continue
        filtered.append(res)
    return filtered


def _dedupe_results(results: List[RecognizerResult]) -> List[RecognizerResult]:
    seen = set()
    deduped: List[RecognizerResult] = []
    for res in results:
        key = (res.entity_type, res.start, res.end)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(res)
    return deduped


def _term_to_regex(term: str) -> str:
    parts = [re.escape(p) for p in re.split(r"\s+", term.strip()) if p]
    return r"\s+".join(parts)


def _find_dictionary_entities(
    text: str,
    custom_dictionary: Dict[str, List[str]],
    enabled_entity_types: List[str],
) -> List[RecognizerResult]:
    if not custom_dictionary:
        return []

    enabled = {e.upper() for e in enabled_entity_types}
    results: List[RecognizerResult] = []
    for ent_type, values in custom_dictionary.items():
        ent_type_u = ent_type.upper()
        if ent_type_u not in enabled:
            continue
        for term in values:
            pattern = _term_to_regex(term)
            if not pattern:
                continue
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                results.append(
                    RecognizerResult(
                        entity_type=ent_type_u,
                        start=match.start(),
                        end=match.end(),
                        score=0.9,
                    )
                )
    return results


def mask_text(
    text: str,
    analyzer: AnalyzerEngine,
    config: MaskingConfig,
    token_gen: TokenGenerator,
) -> Tuple[str, List[RecognizerResult]]:
    """
    Запускает анализ Presidio и заменяет найденные сущности на токены.
    Возвращает маскированный текст и список исходных результатов.
    """
    results: List[RecognizerResult] = analyzer.analyze(
        text=text,
        language=config.language,
        entities=config.entity_types,
        score_threshold=config.score_threshold,
    )

    # Фильтрация результатов:
    # 1) не трогаем уже сгенерированные плейсхолдеры вида <TYPE_0001>, чтобы избежать
    #    "маскирования плейсхолдеров плейсхолдерами" и дублей в JSON;
    # 2) отбрасываем заведомые ложные срабатывания по словарю ENTITY_STOPWORDS.
    placeholder_re = re.compile(r"^<?[A-Z_]+_\d{4}>?$")

    filtered_results: List[RecognizerResult] = []
    for res in results:
        original_value_raw = text[res.start : res.end]
        original_value = original_value_raw.strip()

        # Пропускаем, если это уже плейсхолдер.
        if placeholder_re.match(original_value):
            continue

        # Пропускаем заведомо неверные сущности по словарю стоп-слов,
        # в том числе с учётом всех их склонений (через лемматизацию).
        if _is_in_stopwords(res.entity_type, original_value):
            continue

        filtered_results.append(res)

    # Add dictionary-based entities (explicit allow-list).
    filtered_results.extend(
        _find_dictionary_entities(
            text=text,
            custom_dictionary=config.custom_dictionary,
            enabled_entity_types=config.entity_types,
        )
    )

    # Expand organization abbreviations and clean overlaps.
    filtered_results = _expand_organization_results(text, filtered_results)
    filtered_results = _merge_overlapping_organizations(filtered_results)
    filtered_results = _remove_results_inside_organizations(filtered_results)
    filtered_results = _dedupe_results(filtered_results)

    # Сортируем по убыванию start, чтобы не смещать индексы при подстановке.
    results_sorted = sorted(filtered_results, key=lambda r: r.start, reverse=True)

    masked_text = text
    for res in results_sorted:
        original_value = masked_text[res.start : res.end]
        token = token_gen.get_token(res.entity_type, original_value)

        masked_text = masked_text[: res.start] + token + masked_text[res.end :]

    return masked_text, filtered_results


def process_docx_file(
    file_path: Path,
    output_file_path: Path,
    analyzer: AnalyzerEngine,
    config: MaskingConfig,
    token_gen: TokenGenerator,
    stats: Dict[str, int],
) -> None:
    """
    Обработка .docx-файла с сохранением структуры документа (таблицы, абзацы).
    Текст маскируется по абзацам и ячейкам таблиц.
    """
    try:
        doc = Document(str(file_path))
    except Exception as e:  # noqa: BLE001
        LOGGER.error("Ошибка открытия .docx файла %s: %s", file_path, e)
        stats["errors"] += 1
        return

    paragraphs = list(doc.paragraphs)
    table_paragraphs = [
        p
        for table in doc.tables
        for row in table.rows
        for cell in row.cells
        for p in cell.paragraphs
    ]

    total_par_len = sum(len(p.text) for p in paragraphs)
    total_table_len = sum(len(p.text) for p in table_paragraphs)

    # #region agent log
    _debug_log(
        "docx process before",
        {
            "file_name": file_path.name,
            "paragraphs": len(paragraphs),
            "table_paragraphs": len(table_paragraphs),
            "paragraphs_text_len": total_par_len,
            "tables_text_len": total_table_len,
            "hypothesisId": "H3",
        },
    )
    # #endregion

    all_results: List[RecognizerResult] = []

    def _mask_paragraphs(par_list: List["docx.text.paragraph.Paragraph"]) -> None:  # type: ignore[name-defined]
        nonlocal all_results
        for para in par_list:
            text = para.text
            if not text:
                continue
            masked_text, results = mask_text(
                text=text,
                analyzer=analyzer,
                config=config,
                token_gen=token_gen,
            )
            para.text = masked_text
            all_results.extend(results)

    _mask_paragraphs(paragraphs)
    _mask_paragraphs(table_paragraphs)

    total_par_len_after = sum(len(p.text) for p in paragraphs)
    total_table_len_after = sum(len(p.text) for p in table_paragraphs)

    # #region agent log
    _debug_log(
        "docx process after",
        {
            "file_name": file_path.name,
            "paragraphs_text_len_after": total_par_len_after,
            "tables_text_len_after": total_table_len_after,
            "entities_found": len(all_results),
            "hypothesisId": "H4",
        },
    )
    # #endregion

    # Обновление статистики по типам сущностей.
    for res in all_results:
        stats[res.entity_type] += 1

    try:
        LOGGER.debug("Сохранение .docx файла %s", output_file_path)
        output_file_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_file_path))
        stats["files_processed"] += 1
    except Exception as e:  # noqa: BLE001
        LOGGER.error("Ошибка записи .docx файла %s: %s", output_file_path, e)
        stats["errors"] += 1


def process_file(
    file_path: Path,
    input_dir: Path,
    output_dir: Path,
    analyzer: AnalyzerEngine,
    config: MaskingConfig,
    token_gen: TokenGenerator,
    stats: Dict[str, int],
) -> None:
    """
    Обработка одного файла: чтение, маскирование, запись результата.
    Для .docx используется специальный путь с сохранением структуры документа.
    """
    try:
        rel_path = file_path.relative_to(input_dir)
    except ValueError:
        # На случай, если файл не лежит строго внутри input_dir.
        rel_path = file_path.name

    output_file_path = output_dir / rel_path
    output_file_path.parent.mkdir(parents=True, exist_ok=True)

    suffix = file_path.suffix.lower()

    if suffix == ".docx":
        process_docx_file(
            file_path=file_path,
            output_file_path=output_file_path,
            analyzer=analyzer,
            config=config,
            token_gen=token_gen,
            stats=stats,
        )
        return

    try:
        LOGGER.debug("Чтение файла %s", file_path)
        text = read_text_from_file(file_path)
    except Exception as e:  # noqa: BLE001
        LOGGER.error("Ошибка чтения файла %s: %s", file_path, e)
        stats["errors"] += 1
        return

    try:
        masked_text, results = mask_text(
            text=text,
            analyzer=analyzer,
            config=config,
            token_gen=token_gen,
        )
    except Exception as e:  # noqa: BLE001
        LOGGER.error("Ошибка анализа файла %s: %s", file_path, e)
        stats["errors"] += 1
        return

    # Обновление статистики по типам сущностей.
    for res in results:
        stats[res.entity_type] += 1

    try:
        LOGGER.debug("Запись маскированного файла %s", output_file_path)
        write_text_to_file(output_file_path, masked_text, file_path.suffix)
        stats["files_processed"] += 1
    except Exception as e:  # noqa: BLE001
        LOGGER.error("Ошибка записи файла %s: %s", output_file_path, e)
        stats["errors"] += 1


def setup_logging(log_level: str) -> None:
    logging.basicConfig(
        level=getattr(logging, log_level.upper(), logging.INFO),
        format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Утилита маскирования персональных и корпоративных данных в текстовых документах (Microsoft Presidio)."
    )
    parser.add_argument(
        "--input",
        default="./input_texts",
        help="Путь к входной папке с текстовыми файлами (doc, docx, txt). По умолчанию ./input_texts.",
    )
    parser.add_argument(
        "--output",
        default="./masked_texts",
        help="Путь к выходной папке для маскированных файлов. По умолчанию ./masked_texts.",
    )
    parser.add_argument(
        "--map",
        default="./deanonymization_map.json",
        help="Путь к JSON-файлу со словарём демаскирования. По умолчанию ./deanonymization_map.json.",
    )
    parser.add_argument(
        "--language",
        default=None,
        help="Язык текста (по умолчанию ru или значение из config.yaml).",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Уровень логирования.",
    )
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Путь к YAML-конфигу с параметрами Presidio и списка сущностей.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    setup_logging(args.log_level)

    input_dir = Path(args.input).resolve()
    output_dir = Path(args.output).resolve()
    map_path = Path(args.map).resolve()
    config_path = Path(args.config).resolve() if args.config else None

    if not input_dir.is_dir():
        raise SystemExit(f"Входная директория не найдена: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)

    config = load_config(config_path)
    if args.language:
        config.language = args.language

    # Загрузка стоп-слов из YAML (по умолчанию stopwords.yaml рядом с config.yaml).
    global ENTITY_STOPWORDS
    stopwords_path: Optional[Path] = None
    if config_path is not None:
        stopwords_path = config_path.with_name("stopwords.yaml")
    ENTITY_STOPWORDS = load_stopwords(stopwords_path)

    LOGGER.info("Загрузка Presidio Analyzer (язык=%s)...", config.language)
    analyzer = build_analyzer(config)

    token_gen = TokenGenerator()

    stats: Dict[str, int] = defaultdict(int)

    files = list(get_files(input_dir))
    LOGGER.info("Найдено файлов для обработки: %d", len(files))

    for file_path in files:
        LOGGER.debug("Обработка файла %s", file_path)
        try:
            process_file(
                file_path=file_path,
                input_dir=input_dir,
                output_dir=output_dir,
                analyzer=analyzer,
                config=config,
                token_gen=token_gen,
                stats=stats,
            )
        except Exception as e:  # noqa: BLE001
            LOGGER.error("Необработанное исключение при обработке файла %s: %s", file_path, e)
            stats["errors"] += 1

    # Сохранение словаря демаскирования.
    LOGGER.info("Сохранение словаря демаскирования в %s", map_path)
    with map_path.open("w", encoding="utf-8") as f:
        json.dump(token_gen.mapping, f, ensure_ascii=False, indent=2)

    # Формирование отчёта.
    LOGGER.info("Обработка завершена.")
    LOGGER.info("Файлов успешно обработано: %d", stats.get("files_processed", 0))
    LOGGER.info("Ошибок: %d", stats.get("errors", 0))

    LOGGER.info("Найденные сущности по типам:")
    for key, value in sorted(stats.items()):
        if key in {"files_processed", "errors"}:
            continue
        LOGGER.info("  %s: %d", key, value)


if __name__ == "__main__":
    main()

