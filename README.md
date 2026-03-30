## survey-questionnaire-kpi

Веб‑интерфейс для анализа опросных данных и конструктор анкеты по KPI Framework (выгрузка в DOCX).

Упрощённая веб‑версия логики Telegram‑бота `survey_analysis_bot_FINAL.py`.

### Что умеет

- **Загрузка файлов**: `CSV`, `XLS`, `XLSX`, `SAV (SPSS)`.
- **Выбор переменных**:
  - для **строк** (вопросы/значения);
  - для **столбцов** (группы/сегменты, в т.ч. Total).
- **Метрики**:
  - количество (`N`);
  - процент по столбцу.
- **Веса**: опциональное использование столбца `weight` (как в боте).
- **Z‑тест против Total**:
  - классический пропорционный z‑тест;
  - для каждой группы сравнивает долю с `Total`, считает `z` и `p`;
  - выводит список ячеек, где `p < 0.05` (выше или ниже Total).
- **Экспорт в Excel**: выгрузка рассчитанной таблицы в `survey_analysis.xlsx`.
- **Конструктор анкеты**: страница `/questionnaire-builder` — сборка DOCX по `KPI_framework_ads_FULL.xlsx` (путь к файлу задаётся в `app.py`).

### Локальный запуск

```bash
git clone https://github.com/<ваш-логин>/survey-questionnaire-kpi.git
cd survey-questionnaire-kpi
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
export FLASK_SECRET_KEY="your-secret"  # опционально
python app.py
```

Откройте в браузере: `http://localhost:5000`.

### Деплой на Render

1. Создайте на GitHub репозиторий **`survey-questionnaire-kpi`** и запушьте этот код (корень репозитория = корень приложения).
2. В Render создайте **Web Service** из этого репо **или** используйте Blueprint по `render.yaml`.
3. Если сервис создаёте вручную:
   - Environment: `Python`;
   - Root directory: `.` (корень репозитория);
   - Build command: `pip install -r requirements.txt`;
   - Start command: `gunicorn app:app` (или как в `render.yaml`).
4. (Опционально) задайте переменную `FLASK_SECRET_KEY`.

