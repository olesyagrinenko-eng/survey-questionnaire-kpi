from __future__ import annotations

import os
import re
import tempfile
from typing import Any, List

import pandas as pd
from docx import Document
from flask import Flask, after_this_request, flash, render_template, request, send_file

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_DEFAULT_KPI = os.path.join(_BASE_DIR, "data", "KPI_framework_ads_FULL.xlsx")

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "dev-secret-key-change-me")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024


def _kpi_excel_path() -> str:
    env = (os.getenv("KPI_FRAMEWORK_PATH") or "").strip()
    if env and os.path.isfile(env):
        return env
    return _DEFAULT_KPI


def _load_kpi_framework_df() -> pd.DataFrame:
    path = _kpi_excel_path()
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Файл KPI не найден: {path}")
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    needed = [
        "Блок",
        "KPI",
        "Метка",
        "Формулировка вопроса",
        "Тип / шкала",
        "Варианты ответа (сокращенно)",
        "Обязательность",
        "Когда использовать",
    ]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"В KPI-файле нет колонок: {', '.join(missing)}")
    return df[needed].copy()


def _split_answer_options(raw_value: Any) -> List[str]:
    text = str(raw_value or "").strip()
    if not text or text == "-":
        return []
    parts = [p.strip() for p in re.split(r"\s*[/;]\s*", text) if p.strip()]
    return parts if parts else [text]


def _build_questionnaire_docx(
    df_filtered: pd.DataFrame,
    selected_blocks: List[str],
    selected_obligatory: List[str],
    when_filter: str,
    include_answers: bool,
    include_rotation: bool,
    add_intro_instructions: bool,
) -> str:
    doc = Document()
    doc.add_heading("Анкета количественного опроса (по KPI Framework)", level=1)
    doc.add_paragraph("Документ сформирован автоматически (survey-questionnaire-kpi).")

    criteria = [
        f"Блоки: {', '.join(selected_blocks) if selected_blocks else 'все'}",
        f"Обязательность: {', '.join(selected_obligatory) if selected_obligatory else 'все'}",
        (
            f"Когда использовать: содержит '{when_filter}'"
            if when_filter
            else "Когда использовать: без фильтра"
        ),
        f"Включать варианты ответов: {'да' if include_answers else 'нет'}",
        f"Добавлять ротацию: {'да' if include_rotation else 'нет'}",
    ]
    doc.add_paragraph("Критерии сборки:\n- " + "\n- ".join(criteria))

    if add_intro_instructions:
        doc.add_heading("Общие инструкции интервьюеру", level=2)
        doc.add_paragraph(
            "Показывайте креатив перед блоком оценки. Соблюдайте скрининг, "
            "не подсказывайте ответы, фиксируйте ответ дословно там, где вопрос открытый."
        )

    for idx, row in df_filtered.reset_index(drop=True).iterrows():
        label = str(row["Метка"]).strip()
        kpi_name = str(row["KPI"]).strip()
        block = str(row["Блок"]).strip()
        question = str(row["Формулировка вопроса"]).strip()
        q_type = str(row["Тип / шкала"]).strip()
        obligatory = str(row["Обязательность"]).strip()
        when_use = str(row["Когда использовать"]).strip()

        doc.add_heading(f"Q{idx + 1}. [{label}] {kpi_name}", level=2)
        doc.add_paragraph(f"Блок: {block}")
        doc.add_paragraph(f"Формулировка: {question}")
        doc.add_paragraph(f"Тип вопроса: {q_type}")
        doc.add_paragraph(f"Обязательность: {obligatory}")
        doc.add_paragraph(f"Условие показа: {when_use}")

        if include_answers:
            options = _split_answer_options(row["Варианты ответа (сокращенно)"])
            if options:
                doc.add_paragraph("Варианты ответа:")
                for opt in options:
                    doc.add_paragraph(opt, style="List Bullet")
            else:
                doc.add_paragraph("Варианты ответа: ввести вручную / открытый ответ.")

        if include_rotation and any(token in q_type.lower() for token in ["multi", "single"]):
            doc.add_paragraph("Ротация: ротировать порядок вариантов ответа.")

    fd, output_path = tempfile.mkstemp(suffix=".docx", prefix="questionnaire_")
    os.close(fd)
    doc.save(output_path)
    return output_path


@app.route("/", methods=["GET", "POST"])
def index():
    try:
        framework_df = _load_kpi_framework_df()
    except Exception as exc:
        flash(f"Не удалось прочитать KPI framework: {exc}", "error")
        return render_template(
            "index.html",
            blocks=[],
            obligatory_values=[],
            selected_blocks=[],
            selected_obligatory=[],
            when_filter="",
            include_answers=True,
            include_rotation=True,
            add_intro_instructions=True,
            total_questions=0,
            kpi_source_label=_kpi_excel_path(),
        )

    blocks = sorted([str(v) for v in framework_df["Блок"].dropna().unique()])
    obligatory_values = sorted([str(v) for v in framework_df["Обязательность"].dropna().unique()])

    selected_blocks = request.form.getlist("blocks")
    selected_obligatory = request.form.getlist("obligatory")
    when_filter = (request.form.get("when_filter") or "").strip()
    include_answers = bool(request.form.get("include_answers"))
    include_rotation = bool(request.form.get("include_rotation"))
    add_intro_instructions = bool(request.form.get("add_intro_instructions"))

    if request.method == "POST":
        filtered = framework_df.copy()
        if selected_blocks:
            filtered = filtered[filtered["Блок"].astype(str).isin(selected_blocks)]
        if selected_obligatory:
            filtered = filtered[filtered["Обязательность"].astype(str).isin(selected_obligatory)]
        if when_filter:
            filtered = filtered[
                filtered["Когда использовать"]
                .astype(str)
                .str.contains(when_filter, case=False, na=False)
            ]

        if filtered.empty:
            flash("По выбранным критериям не найдено вопросов.", "error")
            return render_template(
                "index.html",
                blocks=blocks,
                obligatory_values=obligatory_values,
                selected_blocks=selected_blocks,
                selected_obligatory=selected_obligatory,
                when_filter=when_filter,
                include_answers=include_answers,
                include_rotation=include_rotation,
                add_intro_instructions=add_intro_instructions,
                total_questions=int(framework_df.shape[0]),
                kpi_source_label=_kpi_excel_path(),
            )

        try:
            output_path = _build_questionnaire_docx(
                filtered,
                selected_blocks=selected_blocks,
                selected_obligatory=selected_obligatory,
                when_filter=when_filter,
                include_answers=include_answers,
                include_rotation=include_rotation,
                add_intro_instructions=add_intro_instructions,
            )
        except Exception as exc:
            flash(f"Ошибка генерации DOCX: {exc}", "error")
            return render_template(
                "index.html",
                blocks=blocks,
                obligatory_values=obligatory_values,
                selected_blocks=selected_blocks,
                selected_obligatory=selected_obligatory,
                when_filter=when_filter,
                include_answers=include_answers,
                include_rotation=include_rotation,
                add_intro_instructions=add_intro_instructions,
                total_questions=int(framework_df.shape[0]),
                kpi_source_label=_kpi_excel_path(),
            )

        @after_this_request
        def _cleanup_doc(response):  # noqa: ANN001
            try:
                os.remove(output_path)
            except OSError:
                pass
            return response

        return send_file(
            output_path,
            as_attachment=True,
            download_name="questionnaire_from_kpi_framework.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    return render_template(
        "index.html",
        blocks=blocks,
        obligatory_values=obligatory_values,
        selected_blocks=[],
        selected_obligatory=[],
        when_filter="",
        include_answers=True,
        include_rotation=True,
        add_intro_instructions=True,
        total_questions=int(framework_df.shape[0]),
        kpi_source_label=_kpi_excel_path(),
    )


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=True)
