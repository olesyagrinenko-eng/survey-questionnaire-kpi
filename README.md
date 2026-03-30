## survey-questionnaire-kpi

Мини‑приложение: **конструктор анкеты** по файлу `data/KPI_framework_ads_FULL.xlsx` с выгрузкой в **DOCX**. На главной странице показатели выбираются **галочками** (с формулировками), отдельно настраивается вывод вариантов ответов в DOCX.

### Локально

```bash
cd survey-questionnaire-kpi
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export FLASK_SECRET_KEY="your-secret"
python app.py
```

Откройте `http://localhost:5000`.

### Данные KPI

По умолчанию читается `data/KPI_framework_ads_FULL.xlsx` в корне репозитория.  
Чтобы указать другой файл (например, на сервере), задайте переменную окружения **`KPI_FRAMEWORK_PATH`** — полный путь к `.xlsx`.

### Render

- В репозитории есть `.python-version` (`3.11.11`) и `render.yaml` с `PYTHON_VERSION=3.11.11`.
- После деплоя при необходимости задайте `KPI_FRAMEWORK_PATH`, если не используете встроенный файл в `data/`.
