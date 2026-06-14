# UX Analyz — AI-анализ UX веб-сайтов

**Streamlit-приложение для автоматизированного UX-анализа сайтов с помощью LLM (Ollama или GigaChat), скриншотов через Selenium и генерации отчета в PowerPoint.**

Введите URL, получите скриншоты страниц, запустите AI-оценку UX по структурированному промпту и экспортируйте результаты в презентацию `.pptx`.

---

## Скриншоты

> Добавьте скриншоты в `docs/screenshots/` и раскомментируйте строки ниже.

<!--
![Главный экран](docs/screenshots/01-main.png)
![Результат анализа](docs/screenshots/02-analysis.png)
![Экспорт PPTX](docs/screenshots/03-pptx.png)
-->

---

## Функциональность

| Функция | Описание |
|---------|----------|
| **Обход страниц** | Анализ до N связанных страниц от стартового URL |
| **Скриншоты** | Автоматический захват экрана через Selenium + Chrome |
| **LLM-анализ** | UX-оценка через Ollama (локально) или GigaChat (облако) |
| **Структурированный промпт** | Настраиваемый шаблон анализа (`ux_prompt.txt`) |
| **Экспорт PPTX** | Генерация презентации со скриншотами и выводами |
| **Режимы анализа** | Базовый, расширенный, полный (настраивается) |

---

## Стек технологий

| Слой | Технологии |
|------|------------|
| **Интерфейс** | Streamlit |
| **Автоматизация браузера** | Selenium + Chrome |
| **LLM** | Ollama (llama3) / GigaChat API |
| **Генерация отчетов** | python-pptx |
| **Конфигурация** | YAML (`config.yaml`) |
| **Парсинг HTML** | requests, BeautifulSoup |

---

## Архитектура

```
Ввод URL (Streamlit)
        │
        ▼
┌───────────────────┐
│  Selenium         │  Скриншоты страниц
│  (Chrome)         │
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐     ┌──────────────────┐
│  Извлечение HTML  │────▶│  LLM (Ollama /   │
│  + ux_prompt.txt  │     │  GigaChat)       │
└─────────┬─────────┘     └──────────────────┘
          │
          ▼
┌───────────────────┐
│  Отчет PPTX       │  Слайды со скриншотами и выводами
└───────────────────┘
```

---

## Быстрый старт

### Требования

- Python 3.10+
- [Ollama](https://ollama.com/) запущен локально (`ollama serve`)
- Модель: `ollama pull llama3`
- Установлен Google Chrome / Chromium

### Установка

```powershell
git clone https://github.com/Sergey051291/ux_analyz.git
cd ux_analyz

python -m venv .venv
.\.venv\Scripts\activate

pip install -r requirements.txt
```

### Настройка

Отредактируйте `config.yaml`:

```yaml
llm:
  provider: "ollama"
  ollama:
    base_url: "http://localhost:11434"
    model: "llama3"

prompt_template_path: "ux_prompt.txt"
```

Для GigaChat укажите `provider: "gigachat"` и настройте credentials в `ux_analyz.py`.

### Запуск

Терминал 1:
```powershell
ollama serve
```

Терминал 2:
```powershell
streamlit run ux_analyz.py
```

Откройте `http://localhost:8501`, введите URL и запустите анализ.

---

## Структура проекта

```
ux_analyz/
├── ux_analyz.py       # Основное Streamlit-приложение
├── ux_prompt.txt      # Шаблон промпта для UX-анализа
├── config.yaml        # Конфигурация
├── requirements.txt
└── docs/screenshots/
```

---

## Параметры конфигурации

| Параметр | По умолчанию | Описание |
|----------|--------------|----------|
| `max_pages` | 5 | Максимум страниц за один запуск |
| `max_html_length` | 15000 | Ограничение длины HTML для LLM |
| `analysis_modes` | Базовый / Расширенный / Полный | Уровни глубины анализа |

---

## Автор

Личный проект — инструмент автоматизированного UX-аудита с локальной LLM.

**Стек:** Python · Streamlit · Selenium · Ollama · python-pptx
