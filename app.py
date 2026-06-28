from __future__ import annotations

import json
import re
from copy import deepcopy
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
PROJECT_VERSION = "0.1.1"


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
    if "archive_panel" not in st.session_state:
        st.session_state.archive_panel = ""
    if "advanced_panel" not in st.session_state:
        st.session_state.advanced_panel = ""
    if "archive_index_rows" not in st.session_state:
        st.session_state.archive_index_rows = []
    if "archive_index_loaded" not in st.session_state:
        st.session_state.archive_index_loaded = False


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
                rows = value if isinstance(value, list) else []
                if rows and any(any(str(cell or "").strip() for cell in row.values()) for row in rows):
                    filled_fields += 1
            elif str(value or "").strip():
                filled_fields += 1
    return filled_fields, visible_fields


def safe_filename(value: str, fallback: str = "case-conference") -> str:
    slug = re.sub(r"[^a-zA-Z0-9]+", "-", str(value or "").strip().lower()).strip("-")
    return slug or fallback


def truncate_text(value: Any, limit: int = 80) -> str:
    text = str(value or "").strip()
    return text[:limit] + ("…" if len(text) > limit else "")


def current_case_summary() -> Dict[str, str]:
    return archive_summary(st.session_state.deck)


def apply_loaded_payload_to_session(loaded: Dict[str, Any], source_path: str = "") -> None:
    st.session_state.deck = normalize_deck(loaded)
    st.session_state.archive_id = str(loaded.get("archive_id", "") if isinstance(loaded, dict) else "")
    st.session_state.archive_path = str(
        (loaded.get("archive_path", "") if isinstance(loaded, dict) else "") or source_path
    ).strip().lstrip("/")
    clear_widget_state()


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


def slug_for_widget(value: Any) -> str:
    slug = re.sub(r"[^a-zA-Z0-9]+", "_", str(value or "").strip()).strip("_")
    return slug or "column"


def clear_table_cell_state(slide_id: str, table_key: str) -> None:
    prefix = f"cell__{slide_id}__{table_key}__"
    for session_key in list(st.session_state.keys()):
        if session_key.startswith(prefix):
            del st.session_state[session_key]


def normalize_table_rows(rows: Any, columns: List[str]) -> List[Dict[str, str]]:
    if not isinstance(rows, list):
        rows = []

    normalized: List[Dict[str, str]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        old_columns = list(row.keys())
        normalized_row: Dict[str, str] = {}
        for col_index, col in enumerate(columns):
            if col in row:
                value = row.get(col, "")
            elif col_index < len(old_columns):
                value = row.get(old_columns[col_index], "")
            else:
                value = ""
            normalized_row[col] = str(value or "")
        normalized.append(normalized_row)
    return normalized


def render_table_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> Any:
    """Render tables like the Journal Club Builder: stable text cells, not st.data_editor."""
    key = field["key"]
    columns = list(field.get("columns", []))
    if not columns:
        st.info("No table columns are configured for this field.")
        slide_data[key] = []
        return []

    rows = normalize_table_rows(deepcopy(slide_data.get(key, field.get("default", []))), columns)
    if not rows:
        rows = [{col: "" for col in columns}]

    max_rows = int(field.get("max_rows", 6) or 6)
    st.caption(field.get("guide", ""))

    control_cols = st.columns([1, 1, 3])
    with control_cols[0]:
        if st.button("Add row", key=f"table_add__{slide_id}__{key}", disabled=len(rows) >= max_rows, use_container_width=True):
            rows.append({col: "" for col in columns})
            slide_data[key] = rows
            clear_table_cell_state(slide_id, key)
            st.rerun()
    with control_cols[1]:
        if st.button("Remove row", key=f"table_remove__{slide_id}__{key}", disabled=len(rows) <= 1, use_container_width=True):
            rows = rows[:-1]
            slide_data[key] = rows
            clear_table_cell_state(slide_id, key)
            st.rerun()
    with control_cols[2]:
        st.caption(f"{len(rows)} / {max_rows} rows")

    header_cols = st.columns([1 for _ in columns])
    for header_col, column_name in zip(header_cols, columns):
        header_col.markdown(f"**{column_name}**")

    updated_rows: List[Dict[str, str]] = []
    column_signature = "__".join(slug_for_widget(col) for col in columns)

    for row_index, row in enumerate(rows):
        row_cols = st.columns([1 for _ in columns])
        updated_row: Dict[str, str] = {}
        for col_index, column_name in enumerate(columns):
            cell_key = (
                f"cell__{slide_id}__{key}__{row_index}__"
                f"{column_signature}__{slug_for_widget(column_name)}"
            )
            if cell_key not in st.session_state:
                st.session_state[cell_key] = row.get(column_name, "")
            updated_row[column_name] = row_cols[col_index].text_input(
                f"{column_name} row {row_index + 1}",
                key=cell_key,
                label_visibility="collapsed",
            )
        updated_rows.append(updated_row)

    slide_data[key] = updated_rows
    return updated_rows


def render_field(slide_id: str, slide_data: Dict[str, Any], field: Dict[str, Any]) -> None:
    if not field_is_visible(slide_data, field):
        return

    ftype = field["type"]
    if ftype in {"text", "textarea"}:
        value = render_text_field(slide_id, slide_data, field)
    elif ftype == "select":
        value = render_select_field(slide_id, slide_data, field)
    elif ftype == "table":
        st.markdown(f"#### {field['label']}")
        value = render_table_field(slide_id, slide_data, field)
    else:
        st.error(f"Unsupported field type: {ftype}")
        return

    if ftype in {"text", "textarea"}:
        metric_parts = []
        if "max_words" in field:
            metric_parts.append(f"{count_words(value)} / {field['max_words']} words")
        if "max_lines" in field:
            metric_parts.append(f"{len(nonempty_lines(value))} / {field['max_lines']} lines")
        if metric_parts:
            st.caption(" · ".join(metric_parts))

    for problem in validate_field(value, field):
        st.warning(problem)


def render_slide_preview(slide: Dict[str, Any], slide_data: Dict[str, Any]) -> None:
    preview_title = slide_display_title(slide)
    st.caption(f"PowerPoint slide title: {preview_title}")

    for field in slide["fields"]:
        if not field_is_visible(slide_data, field):
            continue
        value = slide_data.get(field["key"], "")
        if field["type"] == "table":
            st.markdown(f"**{field['label']}**")
            st.dataframe(pd.DataFrame(value), hide_index=True, use_container_width=True)
        elif field["type"] == "select":
            st.caption(f"{field['label']}: {value}")
        else:
            st.markdown(f"**{field['label']}**")
            st.write(value if str(value).strip() else "—")


# -----------------------------
# Sidebar and identity display
# -----------------------------


def render_case_identity_card(location: str = "main") -> None:
    archive = current_case_summary()
    presenter = archive.get("presenter") or "—"
    case = archive.get("case") or "—"
    learning_point = archive.get("learning_point") or "—"

    #if location == "sidebar":
        #with st.container(border=True):
        #    st.markdown("**Current case**")
        #    st.caption(f"**Presenter:** {presenter}")
        #    st.caption(f"**Case:** {truncate_text(case, 60)}")
        #    st.caption(f"**Learning point:** {truncate_text(learning_point, 90)}")
        #return

    #with st.container(border=True):
    #    c1, c2, c3 = st.columns([0.8, 1.2, 2.0])
    #    c1.markdown("**Presenter**")
    #    c1.caption(presenter)
    #    c2.markdown("**Case**")
    #    c2.caption(case)
    #    c3.markdown("**Learning point**")
    #    c3.caption(learning_point)


def render_sidebar() -> None:
    with st.sidebar:
        st.header("Slides")
        st.caption("Pick a slide here. Edit the fields on the main page.")

        label_to_id = {nav_label(slide): slide["id"] for slide in CASE_SLIDES}
        id_to_label = {slide["id"]: nav_label(slide) for slide in CASE_SLIDES}

        if "selected_slide_label" not in st.session_state:
            st.session_state.selected_slide_label = id_to_label[st.session_state.selected_slide_id]
        elif st.session_state.selected_slide_label not in label_to_id:
            st.session_state.selected_slide_label = id_to_label[CASE_SLIDES[0]["id"]]
            st.session_state.selected_slide_id = CASE_SLIDES[0]["id"]

        st.radio(
            "Choose slide",
            list(label_to_id.keys()),
            key="selected_slide_label",
            on_change=sync_selected_slide,
        )

        st.divider()
        #render_case_identity_card(location="sidebar")

        #st.divider()
        if st.button("Advanced Drafts/Reset", key="open_advanced_panel_button", use_container_width=True):
            st.session_state.advanced_panel = "" if st.session_state.get("advanced_panel") == "advanced" else "advanced"

        if st.session_state.get("advanced_panel") == "advanced":
            render_advanced_panel()


def render_advanced_panel() -> None:
    with st.container(border=True):
        st.markdown("**Advanced drafts/reset**")

        st.checkbox(
            "Include facilitator notes appendix",
            key="include_facilitator_notes",
            help="Adds facilitator prompts and notes as an appendix in the PowerPoint export.",
        )

        st.caption(
            "Mentor review export creates an editable DOCX with the slide text and facilitator notes."
        )
        review_docx_bytes = build_review_text_docx(st.session_state.deck)
        st.download_button(
            "Download PowerPoint Text Review DOCX",
            data=review_docx_bytes,
            file_name=f"case_conference_text_review_{datetime.now().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        st.divider()

        archive = current_case_summary()
        case_slug = safe_filename(archive.get("case", "case-conference"))
        timestamp = datetime.now().strftime("%Y%m%d")
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
            "Download Editable JSON Draft",
            data=json.dumps(payload, indent=2, ensure_ascii=False).encode("utf-8"),
            file_name=f"case_conference_{timestamp}_{case_slug}.json",
            mime="application/json",
            use_container_width=True,
        )

        uploaded = st.file_uploader("Reload From JSON Draft", type=["json"], key="advanced_uploaded_json")
        if uploaded is not None:
            if st.button("Load Uploaded JSON Draft", key="load_uploaded_json_button", use_container_width=True):
                try:
                    loaded = json.loads(uploaded.getvalue().decode("utf-8"))
                    apply_loaded_payload_to_session(loaded)
                    st.success("Draft loaded.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Could not load JSON draft: {exc}")

        st.divider()

        if st.button("Reset to Case Example", key="reset_case_example_button", use_container_width=True):
            st.session_state.deck = make_default_deck()
            st.session_state.archive_id = ""
            st.session_state.archive_path = ""
            clear_widget_state()
            st.success("Reset complete.")
            st.rerun()

        if st.button("Clear All Fields", key="clear_all_fields_button", use_container_width=True):
            st.session_state.deck = make_default_deck()
            st.session_state.archive_id = ""
            st.session_state.archive_path = ""
            for slide in CASE_SLIDES:
                for field in slide["fields"]:
                    if field.get("type") == "table":
                        st.session_state.deck[slide["id"]][field["key"]] = []
                    else:
                        st.session_state.deck[slide["id"]][field["key"]] = ""
            clear_widget_state()
            st.success("All fields cleared.")
            st.rerun()

        if st.button("Close Advanced Panel", key="close_advanced_panel_button", use_container_width=True):
            st.session_state.advanced_panel = ""
            st.rerun()


# -----------------------------
# Right panel: validation, archive, downloads
# -----------------------------


def render_validation_panel(deck: Dict[str, Dict[str, Any]]) -> None:
    problems = validate_deck(deck)
    filled, total = progress_summary(deck)
    st.progress(filled / total if total else 0)
    st.caption(f"{filled}/{total} visible fields completed")

    if problems:
        with st.expander(f"Validation Warnings ({len(problems)})", expanded=False):
            for problem in problems:
                st.write(f"- {problem}")
    else:
        st.success("All visible fields are within limits.")


def render_save_archive_panel(deck: Dict[str, Dict[str, Any]]) -> None:
    with st.container(border=True):
        st.markdown("#### Save Draft To Archive")
        st.caption("Uses Presenter, Case, and Learning Point from the Title slide.")

        archive = archive_summary(deck)
        if not archive.get("presenter"):
            st.warning("Presenter is blank. Complete the Title slide before saving.")
        if not archive.get("case"):
            st.warning("Case title is blank. Complete the Title slide before saving.")
        if not archive.get("learning_point"):
            st.warning("Learning point is blank. Complete the Title slide before saving.")

        render_case_identity_card(location="main")

        if github_backup_is_configured():
            st.success(github_config_status_message())
        else:
            st.info(github_config_status_message())
            st.caption("Add Streamlit secrets first. Local JSON download/reload still works without GitHub.")

        button_cols = st.columns([1, 1])
        with button_cols[0]:
            save_clicked = st.button("Save Draft To Archive", key="save_draft_to_archive_button", use_container_width=True)
        with button_cols[1]:
            if st.button("Close Save Panel", key="close_save_archive_button", use_container_width=True):
                st.session_state.archive_panel = ""
                st.rerun()

        if save_clicked:
            if not archive.get("presenter") or not archive.get("case") or not archive.get("learning_point"):
                st.error("Please complete Presenter, Case Title, and Learning Point before saving.")
                return
            try:
                result = save_draft_to_github(
                    deck,
                    app_version=PROJECT_VERSION,
                    archive_id=st.session_state.archive_id or None,
                    existing_path=st.session_state.archive_path or None,
                )
                st.session_state.archive_path = result.path
                if not st.session_state.archive_id:
                    loaded = load_draft_from_github(result.path)
                    st.session_state.archive_id = str(loaded.get("archive_id", ""))
                st.session_state.archive_index_loaded = False
                st.success("Draft saved to Archive.")
                st.caption(f"Archive path: {result.path}")
            except GitHubArchiveError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Unexpected archive error: {exc}")


def render_reload_archive_panel() -> None:
    with st.container(border=True):
        st.markdown("#### Reload Saved Draft From Archive")

        if github_backup_is_configured():
            st.success(github_config_status_message())
        else:
            st.info(github_config_status_message())
            st.caption("Add Streamlit secrets first. Local JSON reload is available in Advanced Drafts/Reset.")

        search_text = st.text_input(
            "Search archive",
            value="",
            placeholder="presenter, case, or learning point",
            key="archive_search_text",
        )

        search_cols = st.columns([1, 1])
        with search_cols[0]:
            find_clicked = st.button("Find Saved Cases", key="find_saved_cases_button", use_container_width=True)
        with search_cols[1]:
            if st.button("Close Reload Panel", key="close_reload_archive_button", use_container_width=True):
                st.session_state.archive_panel = ""
                st.rerun()

        if find_clicked:
            try:
                rows = list_drafts_from_github(search_text)
                st.session_state.archive_reload_rows = rows
                if rows:
                    st.success(f"Found {len(rows)} archived case(s).")
                else:
                    st.info("No archived cases found.")
            except GitHubArchiveError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Unexpected archive error: {exc}")

        rows = st.session_state.get("archive_reload_rows", []) or []
        if not rows:
            return

        options = {
            f"{row.get('presenter', '')} | {row.get('case', '')} | {truncate_text(row.get('learning_point', ''), 45)}": row.get("path", "")
            for row in rows
            if row.get("path")
        }
        if not options:
            st.info("No loadable archive paths were found.")
            return

        selected = st.selectbox("Choose a saved case", options=list(options.keys()), key="selected_archived_case_label")
        if st.button("Load Selected Archived Case", key="load_selected_archived_case_button", use_container_width=True):
            try:
                path = options[selected]
                payload = load_draft_from_github(path)
                apply_loaded_payload_to_session(payload, source_path=path)
                st.success("Archived case loaded.")
                st.rerun()
            except Exception as exc:
                st.error(f"Could not load archived case: {exc}")


def render_archive_index_panel() -> None:
    with st.container(border=True):
        st.markdown("#### Archive Index")
        st.caption("Lists saved cases with Presenter, Case, and Learning Point.")

        if github_backup_is_configured():
            st.success(github_config_status_message())
        else:
            st.info(github_config_status_message())
            st.caption("Add Streamlit secrets first. Local JSON download/reload still works without GitHub.")

        action_cols = st.columns([1, 1])
        with action_cols[0]:
            refresh_clicked = st.button("Refresh Archive Index", key="refresh_archive_index_button", use_container_width=True)
        with action_cols[1]:
            if st.button("Close Archive Index Panel", key="close_archive_index_button", use_container_width=True):
                st.session_state.archive_panel = ""
                st.rerun()

        if refresh_clicked or not st.session_state.get("archive_index_loaded", False):
            try:
                rows = list_drafts_from_github("")
                st.session_state.archive_index_rows = rows
                st.session_state.archive_index_loaded = True
                if rows:
                    st.success(f"Found {len(rows)} archived case(s).")
                else:
                    st.info("No archived cases found.")
            except GitHubArchiveError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Unexpected archive index error: {exc}")

        rows = st.session_state.get("archive_index_rows", []) or []
        if rows:
            df = pd.DataFrame(rows)
            display_cols = ["presenter", "case", "learning_point"]
            df_display = df[[col for col in display_cols if col in df.columns]].copy()
            df_display.columns = ["Presenter", "Case", "Learning Point"][: len(df_display.columns)]
            st.dataframe(df_display, hide_index=True, use_container_width=True)

            csv_df = df[[col for col in display_cols if col in df.columns]].copy()
            csv_bytes = csv_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "Download Archive Index CSV",
                data=csv_bytes,
                file_name="case_conference_archive_index.csv",
                mime="text/csv",
                use_container_width=True,
            )


def render_archive_controls(deck: Dict[str, Dict[str, Any]]) -> None:
    if st.button("Save Draft To Archive", key="open_save_archive_panel_button", use_container_width=True):
        st.session_state.archive_panel = "save"

    if st.button("Reload Saved Draft From Archive", key="open_reload_archive_panel_button", use_container_width=True):
        st.session_state.archive_panel = "reload"

    if st.button("View Archive Index", key="open_archive_index_panel_button", use_container_width=True):
        st.session_state.archive_panel = "index"

    panel = st.session_state.get("archive_panel", "")
    if panel == "save":
        render_save_archive_panel(deck)
    elif panel == "reload":
        render_reload_archive_panel()
    elif panel == "index":
        render_archive_index_panel()


def render_downloads(deck: Dict[str, Dict[str, Any]]) -> None:
    problems = validate_deck(deck)
    archive = archive_summary(deck)
    case_slug = safe_filename(archive.get("case", "case-conference"))
    timestamp = datetime.now().strftime("%Y%m%d")
    base_name = f"case_conference_{timestamp}_{case_slug}"

    render_archive_controls(deck)
    st.divider()

    pptx_bytes = build_powerpoint(deck, include_facilitator_notes=st.session_state.include_facilitator_notes)
    st.download_button(
        "Download PowerPoint",
        data=pptx_bytes,
        file_name=f"{base_name}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        disabled=bool(problems),
        use_container_width=True,
    )

    summary_docx = build_word_summary(deck)
    st.download_button(
        "Download 1-Page Summary",
        data=summary_docx,
        file_name=f"{base_name}_summary.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        disabled=bool(problems),
        use_container_width=True,
    )

    worksheet = build_printable_planning_form()
    st.download_button(
        "Download Printable Planning Form",
        data=worksheet,
        file_name="case_conference_printable_planning_form.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )


# -----------------------------
# Main app
# -----------------------------


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="🩺", layout="wide")
    initialize_state()
    render_sidebar()

    st.title(APP_TITLE)
    st.caption("Choose a slide on the left, complete the fields in the main workspace, then export a standardized, progressive case conference.")

    #render_case_identity_card(location="main")

    selected_slide = next(slide for slide in CASE_SLIDES if slide["id"] == st.session_state.selected_slide_id)
    selected_slide_data = st.session_state.deck[selected_slide["id"]]

    editor_col, export_col = st.columns([2.2, 0.9], gap="large")

    with editor_col:
        #st.markdown(f"## {slide_display_title(selected_slide)}")
        #st.caption("Fill out the fields below. The sidebar is only for moving between slides.")

        with st.container(border=True):
            st.markdown("### Edit this slide")
            for field in selected_slide["fields"]:
                render_field(selected_slide["id"], selected_slide_data, field)

        with st.expander("Preview this slide", expanded=False):
            render_slide_preview(selected_slide, selected_slide_data)

    with export_col:
        st.markdown("### Export")
        render_validation_panel(st.session_state.deck)
        st.divider()
        render_downloads(st.session_state.deck)

    st.caption(f"Version {PROJECT_VERSION}")


if __name__ == "__main__":
    main()
