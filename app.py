"""Energy Data Cleaner — Streamlit web app.

Files are processed entirely in memory.
Nothing is written to disk on the server.
"""

import contextlib
import io
from datetime import datetime

import pandas as pd
import streamlit as st

from clean_data import fill_to_year, process_format
from format_detector import detect_format

COLUMNS = ["NMI", "Date_time", "the_date", "the_interval", "net_kWh", "CE_kWh", "SOE_kWh", "notes"]

st.set_page_config(page_title="Energy Data Cleaner", page_icon="⚡", layout="centered")

st.title("⚡ Energy Data Cleaner")
st.caption(
    "Upload one or more meter files. "
    "The app auto-detects the format and exports a single combined clean CSV."
)

# ── Sidebar options ──────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Options")
    fill_gaps = st.checkbox(
        "Fill gaps with weekday-average forecast",
        value=False,
        help=(
            "When enabled, missing 30-min slots within the 12-month window "
            "are filled using a weekday-hour average profile. "
            "Filled rows are marked notes='forecast'."
        ),
    )
    st.divider()
    st.markdown(
        "**Output schema**\n"
        "- `NMI` — 10-char meter ID\n"
        "- `Date_time` — interval end (d/MM/yyyy H:mm)\n"
        "- `the_date` — trading date\n"
        "- `the_interval` — 1–48\n"
        "- `net_kWh` — CE − |SOE|\n"
        "- `CE_kWh` — consumption\n"
        "- `SOE_kWh` — export (≤ 0)\n"
        "- `notes` — quality flag\n"
    )

# ── File uploader ────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Drop files here or click to browse",
    accept_multiple_files=True,
    type=["csv", "xlsx", "xls", "xlsb", "tsv"],
    help="Supports NEM12, Gentrack, Citipower, Origin, and many other formats.",
)

if not uploaded:
    st.info("Upload files above to get started.")
    st.stop()

st.write(f"**{len(uploaded)} file(s) selected.**")

if st.button("▶ Process files", type="primary", use_container_width=True):

    all_rows: list = []
    results:  list = []
    log_lines: list = []

    bar    = st.progress(0.0, text="Starting…")
    status = st.empty()

    for idx, f in enumerate(uploaded):
        bar.progress(idx / len(uploaded), text=f"Processing {f.name}…")
        status.text(f.name)

        file_bytes = f.getvalue()

        try:
            cfg = detect_format(file_bytes, f.name)

            if cfg is None:
                results.append({
                    "File":   f.name,
                    "Status": "⚠️ Format not detected",
                    "Rows":   "-",
                    "NMIs":   "-",
                })
                log_lines.append(f"--- {f.name} ---\nFormat not detected.\n")
                continue

            log_buf = io.StringIO()
            with contextlib.redirect_stdout(log_buf):
                rows, success = process_format(cfg, file_bytes=file_bytes)
                if fill_gaps:
                    rows = fill_to_year(rows)

            log_lines.append(f"--- {f.name} ---\n{log_buf.getvalue()}")

            if success and rows:
                all_rows.extend(rows)
                nmis = sorted(set(r["NMI"] for r in rows))
                results.append({
                    "File":   f.name,
                    "Status": "✅ OK",
                    "Rows":   len(rows),
                    "NMIs":   ", ".join(nmis),
                })
            else:
                results.append({
                    "File":   f.name,
                    "Status": "❌ No rows produced",
                    "Rows":   0,
                    "NMIs":   "-",
                })

        except Exception as exc:
            results.append({
                "File":   f.name,
                "Status": f"❌ {str(exc)[:100]}",
                "Rows":   "-",
                "NMIs":   "-",
            })
            log_lines.append(f"--- {f.name} ---\nERROR: {exc}\n")

    bar.progress(1.0, text="Done")
    status.empty()

    # ── Results summary ──────────────────────────────────────────────────────
    st.subheader("Results")
    st.dataframe(pd.DataFrame(results), use_container_width=True, hide_index=True)

    if all_rows:
        distinct_nmis = sorted(set(r["NMI"] for r in all_rows))
        st.success(
            f"{len(all_rows):,} rows · {len(distinct_nmis)} NMI(s): "
            + ", ".join(distinct_nmis)
        )

        ts     = datetime.now().strftime("%Y%m%d_%H%M%S")
        df_out = pd.DataFrame(all_rows, columns=COLUMNS)

        st.download_button(
            label="⬇ Download Clean CSV",
            data=df_out.to_csv(index=False),
            file_name=f"clean_output_{ts}.csv",
            mime="text/csv",
            type="primary",
            use_container_width=True,
        )
    else:
        st.warning("No rows were produced. Check the processing log below.")

    # ── Log expander ─────────────────────────────────────────────────────────
    with st.expander("Processing log", expanded=not all_rows):
        st.code("\n\n".join(log_lines) if log_lines else "(no output)")
