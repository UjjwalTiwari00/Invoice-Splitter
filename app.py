import streamlit as st
import pandas as pd

from invoice_splitter import (
    parse_particulars,
    process_dataframe,
    apply_hsn_codes,
    dataframe_to_excel_bytes,
    load_hsn_codes,
    save_hsn_codes,
    _norm,
    EMPTY,
)

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Invoice Splitter",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
[data-testid="metric-container"] {
    background: #f7fafc;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 1rem 1.25rem;
}
</style>
""", unsafe_allow_html=True)


# ── Session state ─────────────────────────────────────────────────────────────

if 'hsn_codes' not in st.session_state:
    st.session_state.hsn_codes = load_hsn_codes()

for key in ('out_df', 'stats', 'last_file', 'raw_df'):
    if key not in st.session_state:
        st.session_state[key] = None


def reset_results():
    st.session_state.out_df  = None
    st.session_state.stats   = None
    st.session_state.raw_df  = None
    st.session_state.last_file = None


# ── Helpers ───────────────────────────────────────────────────────────────────

def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Drop all-empty and Unnamed padding columns that Tally writes."""
    df = df.dropna(axis=1, how='all')
    named = [c for c in df.columns if not str(c).startswith('Unnamed:')]
    return df[named] if named else df


def display_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepare DataFrame for display.
    Converts _is_split into a human-readable 'Status' column and drops the flag.
    Avoids pandas Styler entirely (Styler + PyArrow serialisation segfaults on
    Streamlit Cloud with pandas >= 3.x).
    """
    df = df.reset_index(drop=True).copy()
    if '_is_split' in df.columns:
        df.insert(0, 'Status', df['_is_split'].map({True: '🟢 New', False: ''}))
        df = df.drop(columns=['_is_split'])
    return df


# ── Header ────────────────────────────────────────────────────────────────────

st.title("🧾 Invoice Splitter")
st.caption(
    "Splits multi-product Tally Sales Export rows into individual product rows "
    "and auto-fills GST values and HSN codes."
)

tab_process, tab_hsn = st.tabs(["Split Invoices", "Manage HSN Codes"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Split Invoices
# ══════════════════════════════════════════════════════════════════════════════

with tab_process:

    with st.sidebar:
        st.header("How it works")
        st.markdown("""
**Particulars format:**
```
Product x Qty @ Rate (Tax GST%: GSTAmt)
```
**Per-product output:**

| Column | Formula |
|---|---|
| Rate | Value after `@` |
| Value | Qty × Rate |
| Gross Total | Value + GST |
| CGST / SGST | GST ÷ 2 |
| HSN Code | Matched by name |

**Row colours:**
- **Green** = newly split row
- White = unchanged row
""")

    uploaded = st.file_uploader(
        "Upload your Tally Sales Export Excel file",
        type=["xlsx", "xls"],
    )

    if uploaded is None:
        reset_results()
        st.info("Upload an Excel file above to get started.")

    else:
        # Reset when a new file is chosen
        if uploaded.name != st.session_state.last_file:
            reset_results()
            st.session_state.last_file = uploaded.name

        raw_df = clean_columns(pd.read_excel(uploaded, dtype=str))
        st.session_state.raw_df = raw_df

        if 'Particulars' not in raw_df.columns:
            st.error("Column **Particulars** not found in this file. Please check your Excel.")

        else:
            # Input preview
            multi_count = (
                raw_df['Particulars'].dropna()
                .apply(lambda v: len(parse_particulars(str(v))) > 1)
                .sum()
            )
            with st.expander(
                f"Preview input — {len(raw_df)} rows · {len(raw_df.columns)} columns"
                f" · {multi_count} row(s) will be split",
                expanded=False,
            ):
                st.dataframe(raw_df, use_container_width=True, height=240)

            st.divider()

            # Process button
            if st.button("Process File", type="primary", use_container_width=True):
                out_df, stats = process_dataframe(raw_df)
                out_df = apply_hsn_codes(out_df, st.session_state.hsn_codes)
                st.session_state.out_df = out_df
                st.session_state.stats  = stats

            # Results
            if st.session_state.out_df is not None:
                out_df = st.session_state.out_df
                stats  = st.session_state.stats

                st.divider()

                # Metrics
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Input Rows",     stats['input_rows'])
                c2.metric("Output Rows",    stats['output_rows'])
                c3.metric("Rows Split",     stats['split_rows'])
                c4.metric("Passed Through", stats['passthrough_rows'])

                # Totals validation
                mismatches = []
                for col in ['Rate', 'Value', 'Gross Total']:
                    if col in out_df.columns and col in raw_df.columns:
                        try:
                            orig = pd.to_numeric(raw_df[col].replace(EMPTY, None), errors='coerce').sum()
                            new  = pd.to_numeric(out_df[col].replace(EMPTY, None), errors='coerce').sum()
                            if abs(orig - new) > 0.10:
                                mismatches.append(f"{col}: original={orig:.2f}, output={new:.2f}")
                        except Exception:
                            pass
                if mismatches:
                    for m in mismatches:
                        st.warning(m)
                else:
                    st.success("Totals match — Rate, Value and Gross Total verified.")

                # HSN match summary
                if 'HSN Code' in out_df.columns:
                    matched   = (out_df['HSN Code'] != '').sum()
                    unmatched = (out_df['HSN Code'] == '').sum()
                    if unmatched:
                        st.warning(
                            f"HSN: **{matched}** matched, **{unmatched}** unmatched. "
                            "Add missing entries in the **Manage HSN Codes** tab."
                        )

                # Parse errors
                if stats['errors']:
                    with st.expander(f"{len(stats['errors'])} row(s) could not be parsed", expanded=False):
                        for e in stats['errors']:
                            st.warning(e)

                # Download button — above the table
                st.divider()
                try:
                    excel_bytes   = dataframe_to_excel_bytes(out_df)
                    download_name = uploaded.name.rsplit('.', 1)[0] + "_split.xlsx"
                    st.download_button(
                        label="⬇  Download Output Excel",
                        data=excel_bytes,
                        file_name=download_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"Excel generation failed: {e}")

                # Output table
                st.divider()
                st.markdown(
                    "**Output data** — "
                    "**🟢 New** = row created by splitting"
                )

                gst_cols = [c for c in [
                    'HSN Code', 'CGST 2.50% Output', 'SGST 2.50% Output',
                    'CGST 9% Output', 'SGST 9% Output', 'IGST 18%',
                ] if c in out_df.columns]

                tab_full, tab_key = st.tabs(["Full table", "Key columns"])

                with tab_full:
                    st.dataframe(display_df(out_df), use_container_width=True, height=450)

                with tab_key:
                    key_cols = (
                        ['Date', 'Buyer', 'Particulars', 'HSN Code',
                         'Quantity', 'Rate', 'Value', 'Gross Total', 'Local Sales']
                        + [c for c in gst_cols if c != 'HSN Code']
                    )
                    key_cols = [c for c in key_cols if c in out_df.columns]
                    key_with_flag = key_cols + (['_is_split'] if '_is_split' in out_df.columns else [])
                    st.dataframe(display_df(out_df[key_with_flag]), use_container_width=True, height=450)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Manage HSN Codes
# ══════════════════════════════════════════════════════════════════════════════

with tab_hsn:
    st.subheader("HSN Code Manager")
    st.caption(
        "Add, edit, or delete product → HSN code mappings. "
        "Use the **+** row at the bottom to add entries, the trash icon to delete. "
        "Click **Save Changes** when done."
    )

    # Build editable DataFrame from session state
    hsn_df = pd.DataFrame(
        [(k, v) for k, v in sorted(st.session_state.hsn_codes.items())],
        columns=["Product Name", "HSN Code"],
    )

    edited_df = st.data_editor(
        hsn_df,
        num_rows="dynamic",
        use_container_width=True,
        height=500,
        column_config={
            "Product Name": st.column_config.TextColumn(
                "Product Name",
                help="Case-insensitive. Partial name matching is supported.",
                width="large",
            ),
            "HSN Code": st.column_config.TextColumn(
                "HSN Code",
                help="e.g. 33049990",
                width="medium",
            ),
        },
        hide_index=True,
    )

    st.divider()

    col_save, col_reset, col_dl, col_ul = st.columns(4)

    # Save
    with col_save:
        if st.button("Save Changes", type="primary", use_container_width=True):
            new_map, skipped = {}, 0
            for _, row in edited_df.iterrows():
                name = str(row["Product Name"]).strip()
                code = str(row["HSN Code"]).strip()
                if name and name.upper() not in ('NAN', '') and code and code.upper() not in ('NAN', ''):
                    new_map[_norm(name)] = code
                else:
                    skipped += 1
            st.session_state.hsn_codes = new_map
            try:
                save_hsn_codes(new_map)
                msg = f"Saved {len(new_map)} entries to hsn_codes.json."
            except Exception as e:
                msg = f"Session updated ({len(new_map)} entries). File write failed: {e}"
            if skipped:
                msg += f" {skipped} blank row(s) skipped."
            st.success(msg)
            reset_results()

    # Reset
    with col_reset:
        if st.button("Reset to File", use_container_width=True):
            st.session_state.hsn_codes = load_hsn_codes()
            st.success("Reloaded from hsn_codes.json.")
            reset_results()
            st.rerun()

    # Download CSV
    with col_dl:
        csv_bytes = edited_df.to_csv(index=False).encode()
        st.download_button(
            "Download CSV",
            data=csv_bytes,
            file_name="hsn_codes.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # Import CSV
    with col_ul:
        csv_upload = st.file_uploader(
            "Import CSV",
            type=["csv"],
            help="CSV must have columns: Product Name, HSN Code",
            label_visibility="collapsed",
        )
        if csv_upload:
            try:
                imp_df = pd.read_csv(csv_upload, dtype=str).fillna('')
                imp_df.columns = [c.strip() for c in imp_df.columns]
                if "Product Name" not in imp_df.columns or "HSN Code" not in imp_df.columns:
                    st.error("CSV must have columns: Product Name, HSN Code")
                else:
                    imp_map = {
                        _norm(r["Product Name"]): str(r["HSN Code"]).strip()
                        for _, r in imp_df.iterrows()
                        if r["Product Name"].strip() and r["HSN Code"].strip()
                    }
                    st.session_state.hsn_codes.update(imp_map)
                    st.success(f"Imported {len(imp_map)} entries. Click Save to persist.")
                    st.rerun()
            except Exception as e:
                st.error(f"Import failed: {e}")

    st.divider()
    st.caption(f"**{len(st.session_state.hsn_codes)} entries** currently loaded in memory.")
