from __future__ import annotations

import json
import re
from datetime import datetime
from typing import Any, Dict, List

import pandas as pd
import streamlit as st

from case_schema import CASE_SLIDES, archive_summary, make_default_deck
from docx_builder import build_review_text_docx, build_word_summary
from github_storage import (
    GitHubArchiveError,
    generate_archive_id,
    github_backup_is_configured,
    github_config_status_message,
    list_drafts_from_github,
    load_draft_from_github,
    save_draft_to_github,
)
from pptx_builder import build_powerpoint
from printable_form_builder import build_printable_planning_form

APP_TITLE = "Pediatric Case Conference Builder"
PROJECT_VERSION = "0.1.0"


# -----------------------------
# Utility functions
# -----------------------------


def count_words(text: Any) -> int:
    return len(re.findall(r"\b\w+\b", str(text or "")))


def nonempty_lines(text: Any) -> List[str]:
    return [line.strip() for line in str(text or "").splitlines() if line.strip()]


def clear_widget_state() -> None:
    for key in list(st.session_state.keys()):
        if key.startswith(("widget__", "table__", "cell__")):
            del st.session_state[key]


def normalize_deck(deck_or_payload: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    """Merge uploaded/old draft data into the current schema."""
    deck = deck_or_payload.get("deck", deck_or_payload) if isinstance(deck_or_payload, dict) else {}
    default = make_default_deck()
    for slide in CASE_SLIDES:
        sid = slide["id"]
        if isinstance(deck, dict) and sid in deck and isinstance(deck[sid], dict):
            for field in slide["fields"]:
                fkey = field["key"]
                if fkey in deck[sid]:
                    default[sid][fkey] = deck[sid][fkey]
    return default


def initialize_state() -> None:
    if "deck" not in st.session_state:
        st.session_state.deck = make_default_deck()
    if "selected_slide_id" not in st.session_state:
        st.session_state.selected_slide_id = CASE_SLIDES[0]["id"]
    if "include_facilitator_notes" not in st.session_state:
        st.session_state.include_facilitator_notes = True
    if "archive_id" not in st.session_state:
        st.session_state.archive_id = ""
    if "archive_path" not in st.session_state:
        st.session_state.archive_path = ""
    if "archive_index_rows" not in st.session_state:
        st.session_state.archive_index_rows = []


def nav_label(slide: Dict[str, Any]) -> str:
    labels = {
        "title_goal": "Title",
        "opening_case": "Opening Case",
        "history_questions": "Ask: History",
        "history_reveal": "History Reveal",
        "exam_questions": "Ask: Exam",
        "exam_reveal": "Exam Reveal",
        "problem_representation": "Problem Representation",
        "differential": "Differential",
        "diagnostic_studies_question": "Ask: Diagnostics",
        "diagnostic_interpretation": "Interpret Diagnostics",
        "management_decision": "Management",
        "disease_teaching": "Disease Teaching",
        "major_learning_points": "Major Points",
        "return_to_patient": "Return to Patient",
        "final_bottom_line": "Bottom Line",
    }
    return labels.get(slide["id"], slide["label"])


def slide_display_title(slide: Dict[str, Any]) -> str:
    return str(slide.get("export_title") or slide.get("label") or "Untitled slide").strip()


def field_is_visible(slide_data: Dict[str, Any], field: Dict[str, Any]) -> bool:
    condition = field.get("show_if")
    if not condition:
        return True
    for controlling_key, expected_value in condition.items():
        if slide_data.get(controlling_key) != expected_value:
            return False
    return True


def sync_selected_slide() -> None:
    label_to_id = {nav_label(slide): slide["id"] for slide in CASE_SLIDES}
    selected_label = st.session_state.get("selected_slide_label")
    if selected_label in label_to_id:
        st.session_state.selected_slide_id = label_to_id[selected_label]


def validate_field(value: Any, field: Dict[str, Any]) -> List[str]:
    problems: List[str] = []
    label = field.get("label", field.get("key", "Field"))

    if field.get("required"):
        if field.get("type") == "table":
            rows = value if isinstance(value, list) else []
            if not rows or not any(any(str(cell or "").strip() for cell in row.values()) for row in rows):
                problems.append(f"{label} is required.")
        elif str(value or "").strip() == "":
            problems.append(f"{label} is required.")

    if field.get("type") == "table":
        rows = value if isinstance(value, list) else []
        max_rows = field.get("max_rows")
        if max_rows is not None and len(rows) > int(max_rows):
            problems.append(f"{label} has {len(rows)}/{max_rows} rows.")
        for row_idx, row in enumerate(rows, start=1):
            for col, cell_value in row.items():
                if count_words(cell_value) > 18:
                    problems.append(f"{label} row {row_idx}, {col} is long ({count_words(cell_value)} words).")
        return problems

    text = str(value or "")
    if "max_words" in field:
        words = count_words(text)
        if words > int(field["max_words"]):
            problems.append(f"{label} has {words}/{field['max_words']} words.")
    if "max_lines" in field:
        lines = len(nonempty_lines(text))
        if lines > int(field["max_lines"]):
            problems.append(f"{label} has {lines}/{field['max_lines']} lines.")
    if "max_words_per_line" in field:
        for line_number, line in enumerate(nonempty_lines(text), start=1):
            words = count_words(line)
            if words > int(field["max_words_per_line"]):
                problems.append(f"{label} line {line_number} has {words}/{field['max_words_per_line']} words.")
    return problems


def validate_deck(deck: Dict[str, Dict[str, Any]]) -> List[str]:
    problems: List[str] = []
    for slide in CASE_SLIDES:
        sid = slide["id"]
        slide_data = deck.get(sid, {})
        for field in slide["fields"]:
            if not field_is_visible(slide_data, field):
                continue
            value = slide_data.get(field["key"], "")
            for problem in validate_field(value, field):
                problems.append(f"{slide['label']} — {problem}")
    return problems


def progress_summary(deck: Dict[str, Dict[str, Any]]) -> tuple[int, int]:
    visible_fields = 0
    filled_fields = 0
    for slide in CASE_SLIDES:
        slide_data = deck.get(slide["id"], {})
        for field in slide["fields"]:
            if not field_is_visible(slide_data, field):
                continue
            visible_fields += 1
            value = slide_data.get(field["key"], "")
            if field.get("type") == "table":
                if value:
                    filled_fields += 1
            elif str(value or "").strip():
                filled_fields += 1
    return filled_fields, visible_fields


def safe_filename(value: str, fallback: str = "case-conference") -> str:
    slug = re.sub(r"[^a-zA-Z0-9]+", "-", str(value or "").strip().lower()).strip("-")
    return slug or fallback


# -----------------------------
# Field rendering
# -----------------------------


def render_text_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> Any:
    key = field["key"]
    widget_key = f"widget__{slide_id}__{key}"
    if widget_key not in st.session_state:
        st.session_state[widget_key] = slide_data.get(key, field.get("default", ""))

    label = field["label"]
    help_text = field.get("guide")

    if field["type"] == "textarea":
        value = st.text_area(label, key=widget_key, help=help_text, height=field.get("height", 150))
    else:
        value = st.text_input(label, key=widget_key, help=help_text)

    slide_data[key] = value
    return value


def render_select_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> Any:
    key = field["key"]
    options = field.get("options", [])
    current = slide_data.get(key, field.get("default", options[0] if options else ""))
    if current not in options and options:
        current = options[0]
    widget_key = f"widget__{slide_id}__{key}"
    index = options.index(current) if current in options else 0

    value = st.selectbox(field["label"], options=options, index=index, key=widget_key, help=field.get("guide"))
    slide_data[key] = value
    return value


def render_table_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> Any:
    key = field["key"]
    table_key = f"table__{slide_id}__{key}"
    columns = field.get("columns", [])
    current_rows = slide_data.get(key, field.get("default", []))
    df = pd.DataFrame(current_rows)
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    df = df[columns]

    max_rows = int(field.get("max_rows", 6))
    st.caption(field.get("guide", ""))
    edited = st.data_editor(
        df,
        key=table_key,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_order=columns,
    )
    edited = edited.fillna("")
    rows = edited.to_dict(orient="records")[:max_rows]
    rows = [{col: str(row.get(col, "")).strip() for col in columns} for row in rows]
    rows = [row for row in rows if any(value for value in row.values())]
    slide_data[key] = rows
    return rows


def render_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> Any:
    if not field_is_visible(slide_data, field):
        return None
    st.markdown(f"#### {field['label']}") if field.get("type") == "table" else None
    if field["type"] in {"text", "textarea"}:
        value = render_text_field(slide_id, slide_data, field)
    elif field["type"] == "select":
        value = render_select_field(slide_id, slide_data, field)
    elif field["type"] == "table":
        value = render_table_field(slide_id, slide_data, field)
    else:
        value = render_text_field(slide_id, slide_data, field)

    problems = validate_field(value, field)
    if problems:
        for problem in problems:
            st.warning(problem)
    else:
        if field.get("type") != "table" and ("max_words" in field or "max_lines" in field):
            details = []
            if "max_words" in field:
                details.append(f"{count_words(value)}/{field['max_words']} words")
            if "max_lines" in field:
                details.append(f"{len(nonempty_lines(value))}/{field['max_lines']} lines")
            if details:
                st.caption(" | ".join(details))
    return value


# -----------------------------
# UI sections
# -----------------------------


def render_sidebar() -> None:
    st.sidebar.title("Case Builder")
    labels = [nav_label(slide) for slide in CASE_SLIDES]
    id_to_label = {slide["id"]: nav_label(slide) for slide in CASE_SLIDES}
    current_label = id_to_label.get(st.session_state.selected_slide_id, labels[0])
    st.sidebar.radio(
        "Section",
        options=labels,
        index=labels.index(current_label),
        key="selected_slide_label",
        on_change=sync_selected_slide,
    )

    filled, total = progress_summary(st.session_state.deck)
    st.sidebar.progress(filled / total if total else 0)
    st.sidebar.caption(f"{filled}/{total} fields started")

    st.sidebar.checkbox("Include facilitator notes appendix", key="include_facilitator_notes")

    if st.sidebar.button("Clear all fields"):
        st.session_state.deck = make_default_deck()
        st.session_state.archive_id = ""
        st.session_state.archive_path = ""
        clear_widget_state()
        st.rerun()


def render_current_slide() -> None:
    selected_id = st.session_state.selected_slide_id
    slide = next((s for s in CASE_SLIDES if s["id"] == selected_id), CASE_SLIDES[0])
    slide_data = st.session_state.deck.setdefault(selected_id, {})

    st.header(slide_display_title(slide))
    st.caption(slide.get("label", ""))

    for field in slide["fields"]:
        render_field(selected_id, slide_data, field)
        st.divider()


def render_validation_panel() -> None:
    problems = validate_deck(st.session_state.deck)
    with st.expander("Validation check", expanded=bool(problems)):
        if not problems:
            st.success("No required-field or word-limit issues found.")
        else:
            st.warning(f"{len(problems)} item(s) need attention before export.")
            for problem in problems[:50]:
                st.write(f"- {problem}")


def render_export_panel() -> None:
    archive = archive_summary(st.session_state.deck)
    case_slug = safe_filename(archive.get("case", "case-conference"))
    timestamp = datetime.now().strftime("%Y%m%d")
    base_name = f"case_conference_{timestamp}_{case_slug}"

    with st.expander("Export PowerPoint, Word documents, worksheet, or JSON", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            pptx_bytes = build_powerpoint(st.session_state.deck, include_facilitator_notes=st.session_state.include_facilitator_notes)
            st.download_button(
                "Download PowerPoint",
                data=pptx_bytes,
                file_name=f"{base_name}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

            summary_docx = build_word_summary(st.session_state.deck)
            st.download_button(
                "Download one-page summary",
                data=summary_docx,
                file_name=f"{base_name}_summary.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

            worksheet = build_printable_planning_form()
            st.download_button(
                "Download printable planning worksheet",
                data=worksheet,
                file_name="case_conference_printable_planning_form.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        with col2:
            review_docx = build_review_text_docx(st.session_state.deck)
            st.download_button(
                "Download mentor review document",
                data=review_docx,
                file_name=f"{base_name}_mentor_review.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

            payload = {
                "archive_id": st.session_state.archive_id or generate_archive_id(),
                "archive_path": st.session_state.archive_path,
                "presenter": archive.get("presenter", ""),
                "case": archive.get("case", ""),
                "learning_point": archive.get("learning_point", ""),
                "app_version": PROJECT_VERSION,
                "saved_at_local": datetime.now().isoformat(timespec="seconds"),
                "deck": st.session_state.deck,
            }
            st.download_button(
                "Download editable JSON draft",
                data=json.dumps(payload, indent=2, ensure_ascii=False).encode("utf-8"),
                file_name=f"{base_name}.json",
                mime="application/json",
                use_container_width=True,
            )

            uploaded = st.file_uploader("Reload from JSON draft", type=["json"])
            if uploaded is not None:
                try:
                    payload = json.loads(uploaded.getvalue().decode("utf-8"))
                    st.session_state.deck = normalize_deck(payload)
                    st.session_state.archive_id = str(payload.get("archive_id", ""))
                    st.session_state.archive_path = str(payload.get("archive_path", ""))
                    clear_widget_state()
                    st.success("Draft loaded. The fields have been restored.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Could not load JSON draft: {exc}")


def render_archive_panel() -> None:
    with st.expander("Case Archive", expanded=False):
        st.write("Archive index columns: **Presenter**, **Case**, and **Learning Point**.")
        if github_backup_is_configured():
            st.success(github_config_status_message())
        else:
            st.info(github_config_status_message())
            st.caption("GitHub archive is optional. Local JSON download/reload still works without it.")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Save current case to GitHub archive", use_container_width=True):
                try:
                    result = save_draft_to_github(
                        st.session_state.deck,
                        app_version=PROJECT_VERSION,
                        archive_id=st.session_state.archive_id or None,
                        existing_path=st.session_state.archive_path or None,
                    )
                    st.session_state.archive_path = result.path
                    if not st.session_state.archive_id:
                        # Reload payload to get the generated archive id into state.
                        loaded = load_draft_from_github(result.path)
                        st.session_state.archive_id = str(loaded.get("archive_id", ""))
                    st.success(f"Saved to archive: {result.path}")
                except GitHubArchiveError as exc:
                    st.error(str(exc))
                except Exception as exc:
                    st.error(f"Unexpected archive error: {exc}")

        with col2:
            search_text = st.text_input("Search archive", value="", placeholder="presenter, case, or learning point")
            if st.button("Refresh archive index", use_container_width=True):
                try:
                    st.session_state.archive_index_rows = list_drafts_from_github(search_text)
                except GitHubArchiveError as exc:
                    st.error(str(exc))
                except Exception as exc:
                    st.error(f"Unexpected archive error: {exc}")

        rows = st.session_state.archive_index_rows
        if rows:
            df = pd.DataFrame(rows)
            display_cols = ["saved_date", "presenter", "case", "learning_point", "path"]
            st.dataframe(df[display_cols], use_container_width=True, hide_index=True)
            csv_bytes = df[["presenter", "case", "learning_point", "saved_date", "path"]].to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "Download archive index CSV",
                data=csv_bytes,
                file_name="case_conference_archive_index.csv",
                mime="text/csv",
                use_container_width=True,
            )
            options = {f"{row['saved_date']} | {row['presenter']} | {row['case']}": row["path"] for row in rows}
            selected = st.selectbox("Load a saved case", options=list(options.keys()))
            if st.button("Load selected archived case", use_container_width=True):
                try:
                    payload = load_draft_from_github(options[selected])
                    st.session_state.deck = normalize_deck(payload)
                    st.session_state.archive_id = str(payload.get("archive_id", ""))
                    st.session_state.archive_path = str(payload.get("archive_path", options[selected]))
                    clear_widget_state()
                    st.success("Archived case loaded.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Could not load archived case: {exc}")
        else:
            st.caption("Refresh the archive index to list saved cases.")


# -----------------------------
# Main app
# -----------------------------


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="🩺", layout="wide")
    initialize_state()
    render_sidebar()

    st.title(APP_TITLE)
    st.caption("Build a structured, progressive, interactive pediatric case conference without wrestling with slide formatting.")

    archive = archive_summary(st.session_state.deck)
    metric_cols = st.columns(3)
    metric_cols[0].metric("Presenter", archive.get("presenter") or "—")
    metric_cols[1].metric("Case", archive.get("case") or "—")
    metric_cols[2].metric("Learning point", archive.get("learning_point")[:45] + ("…" if len(archive.get("learning_point", "")) > 45 else "") if archive.get("learning_point") else "—")

    left, right = st.columns([0.62, 0.38], gap="large")
    with left:
        render_current_slide()
    with right:
        render_validation_panel()
        render_export_panel()
        render_archive_panel()

    st.caption(f"Version {PROJECT_VERSION}")


if __name__ == "__main__":
    main()
