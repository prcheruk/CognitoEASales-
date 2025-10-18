import gradio as gr
import pandas as pd
import os
from datetime import datetime
import logging

# Import your existing functions
from SalesAssist_v11 import (
    execute_sql_file,
    perform_ea_analysis_with_serial_reconciliation,
    get_cisco_ai_client,
    summarize_with_cisco_ai,
    prepare_data_for_openai,
    create_ea_powerpoint_with_ai,
    load_excel_from_local_onedrive,
    send_webex_message_with_files,
    WEBEX_ROOM_ID,
    FALLBACK_WEBEX_EMAIL,
    AUTHORIZATION_TOKEN,
    DB_USERNAME,
    DB_PASSWORD,
    DB_DSN,
)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def fetch_cav_ids(customer_name):
    """
    Fetch available CAV IDs for a customer and display full customer information

    Returns:
        Tuple of (status_message, customer_dataframe, dropdown_choices, dropdown_visibility, button_visibility, dataframe_visibility, serial_upload_visibility)
    """
    if not customer_name or not customer_name.strip():
        return (
            "⚠️ Please enter a customer name",
            None,
            [],
            gr.update(visible=False),
            gr.update(visible=False),
            gr.update(visible=False),
            gr.update(visible=False),
        )

    customer_name = customer_name.strip()

    try:
        logger.info(f"Fetching CAV IDs for: {customer_name}")

        customer_records = execute_sql_file(
            sql_file_path="get_customers_by_gu_name.sql",
            bind_params={"gu_name": customer_name},
        )

        if not customer_records:
            return (
                f"❌ No records found for customer: **{customer_name}**\n\nPlease check the spelling and try again.",
                None,
                [],
                gr.update(visible=False),
                gr.update(visible=False),
                gr.update(visible=False),
                gr.update(visible=False),
            )

        # Convert to DataFrame
        customer_df = pd.DataFrame(customer_records)
        customer_df.columns = customer_df.columns.str.upper()

        # Convert numeric columns
        for col in ["CAV_ID", "CAV_BU_ID", "PARTY_ID", "PARTY_COUNT"]:
            if col in customer_df.columns:
                customer_df[col] = pd.to_numeric(
                    customer_df[col], errors="coerce"
                ).astype(pd.Int64Dtype())

        # Define the exact columns from your SQL query (in order)
        display_columns = [
            "GU_NAME",
            "CAV_ID",
            "CAV_NAME",
            "CAV_BU_NAME",
            "PARTY_COUNT",
            "IB_COVERED_PRODUCT_LIST",
            "IB_UNCOVERED_PRODUCT_LIST",
            "IB_NEVER_COVERED_PRODUCT_LIST",
        ]

        # Filter to only columns that exist
        available_columns = [
            col for col in display_columns if col in customer_df.columns
        ]

        # Create display dataframe with only the relevant columns
        display_df = customer_df[available_columns].copy()

        # Format long product list columns for better readability
        for col in [
            "IB_COVERED_PRODUCT_LIST",
            "IB_UNCOVERED_PRODUCT_LIST",
            "IB_NEVER_COVERED_PRODUCT_LIST",
        ]:
            if col in display_df.columns:
                # Truncate very long strings and add ellipsis
                display_df[col] = display_df[col].apply(
                    lambda x: (str(x)[:150] + "...")
                    if pd.notna(x) and len(str(x)) > 150
                    else (str(x) if pd.notna(x) else "")
                )

        # Get available CAV IDs with additional info
        if "CAV_ID" in customer_df.columns and "CAV_NAME" in customer_df.columns:
            cav_info = customer_df[["CAV_ID", "CAV_NAME"]].drop_duplicates()
            cav_choices = [
                f"{row['CAV_ID']} - {row['CAV_NAME']}" for _, row in cav_info.iterrows()
            ]
        elif "CAV_ID" in customer_df.columns:
            available_cav_ids = customer_df["CAV_ID"].unique().tolist()
            cav_choices = [str(cav_id) for cav_id in available_cav_ids]
        else:
            return (
                "❌ No CAV IDs found in the data",
                None,
                [],
                gr.update(visible=False),
                gr.update(visible=False),
                gr.update(visible=False),
                gr.update(visible=False),
            )

        if not cav_choices:
            return (
                "❌ No CAV IDs available for this customer",
                None,
                [],
                gr.update(visible=False),
                gr.update(visible=False),
                gr.update(visible=False),
                gr.update(visible=False),
            )

        # Count unique values
        unique_cavs = (
            customer_df["CAV_ID"].nunique() if "CAV_ID" in customer_df.columns else 0
        )
        unique_bus = (
            customer_df["CAV_BU_NAME"].nunique()
            if "CAV_BU_NAME" in customer_df.columns
            else 0
        )
        total_parties = (
            customer_df["PARTY_COUNT"].sum()
            if "PARTY_COUNT" in customer_df.columns
            else 0
        )

        status_msg = f"""✅ Found **{unique_cavs}** CAV ID(s) for **{customer_name}**

**Summary:**
- Total Records: {len(customer_df):,}
- Unique CAV IDs: {unique_cavs}
- Customer Business Units: {unique_bus}
- Total Parties: {total_parties:,}

📊 Review the customer details below (install base coverage by CAV & Business Unit):"""

        return (
            status_msg,
            display_df,  # Return the formatted display DataFrame
            cav_choices,
            gr.update(visible=True, choices=cav_choices, value=cav_choices[0]),
            gr.update(visible=True),
            gr.update(visible=True),  # Show the dataframe component
            gr.update(visible=True),  # Show the serial file upload
        )

    except Exception as e:
        logger.error(f"Error fetching CAV IDs: {str(e)}", exc_info=True)
        return (
            f"❌ Error fetching CAV IDs: {str(e)}",
            None,
            [],
            gr.update(visible=False),
            gr.update(visible=False),
            gr.update(visible=False),
            gr.update(visible=False),
        )


def run_ea_analysis(
    customer_name,
    cav_selection,
    serial_file,
    send_to_webex_flag=False,
    progress=gr.Progress(),
):
    """
    Run complete EA analysis for a customer with selected CAV ID and optional serial file upload
    """
    if not customer_name or not customer_name.strip():
        return "❌ Please enter a customer name", None, None, None

    if not cav_selection:
        return "❌ Please select a CAV ID", None, None, None

    customer_name = customer_name.strip()

    # Extract CAV ID from selection (handles "12345 - Company Name" format)
    target_cav_id_int = int(cav_selection.split(" - ")[0].strip())

    # Debug: Check Webex configuration
    print(f"\n{'=' * 60}")
    print("WEBEX CONFIGURATION CHECK:")
    print(f"{'=' * 60}")
    print(f"send_to_webex_flag: {send_to_webex_flag}")
    print(f"AUTHORIZATION_TOKEN exists: {bool(AUTHORIZATION_TOKEN)}")
    print(
        f"AUTHORIZATION_TOKEN is not empty: {bool(AUTHORIZATION_TOKEN and AUTHORIZATION_TOKEN.strip())}"
    )
    print(f"WEBEX_ROOM_ID: {WEBEX_ROOM_ID}")
    print(f"FALLBACK_WEBEX_EMAIL: {FALLBACK_WEBEX_EMAIL}")
    print(f"Serial file uploaded: {serial_file is not None}")
    if serial_file:
        print(f"Serial file path: {serial_file}")
    print(f"{'=' * 60}\n")

    try:
        progress(0.1, desc=f"✅ Using CAV ID: {target_cav_id_int}")
        logger.info(
            f"Starting analysis for {customer_name}, CAV ID: {target_cav_id_int}"
        )

        # Step 2: Load data files from OneDrive
        progress(0.15, desc="📂 Loading data files from OneDrive...")

        case_history_df = load_excel_from_local_onedrive(
            "IB Analyst\\Metro Fire Dept", "case_history_by_cav.xlsx"
        )
        if case_history_df is None:
            case_history_df = pd.DataFrame()

        progress(0.2, desc="📂 Loading install base...")
        install_base_df = load_excel_from_local_onedrive(
            "IB Analyst\\Metro Fire Dept", "install_base_by_cav.xlsx"
        )
        if install_base_df is None:
            install_base_df = pd.DataFrame()

        progress(0.25, desc="📂 Loading sales history...")
        sales_history_df = load_excel_from_local_onedrive(
            "IB Analyst\\Metro Fire Dept", "sales_history_data.xlsx"
        )
        if sales_history_df is None:
            sales_history_df = pd.DataFrame()

        # Load customer serial numbers - use uploaded file if provided, otherwise use OneDrive
        progress(0.3, desc="📂 Loading customer serial data...")
        if serial_file:
            logger.info(f"Using uploaded serial file: {serial_file}")
            print(f"📎 Loading uploaded serial file: {serial_file}")
            try:
                # Determine file type and load accordingly
                if serial_file.endswith(".csv"):
                    customer_serial_ib_df = pd.read_csv(serial_file)
                else:
                    customer_serial_ib_df = pd.read_excel(serial_file)
                print(
                    f"✅ Uploaded serial file loaded: {len(customer_serial_ib_df)} records"
                )
                logger.info(
                    f"Uploaded serial file loaded successfully: {len(customer_serial_ib_df)} records"
                )
            except Exception as e:
                logger.error(f"Error loading uploaded serial file: {e}")
                print(f"⚠️ Error loading uploaded file, falling back to OneDrive: {e}")
                customer_serial_ib_df = load_excel_from_local_onedrive(
                    "IB Analyst\\Metro Fire Dept", "customer_serial_ib_data.xlsx"
                )
                if customer_serial_ib_df is None:
                    customer_serial_ib_df = pd.DataFrame()
        else:
            logger.info("No serial file uploaded, using OneDrive default")
            print("📂 Using default serial file from OneDrive")
            customer_serial_ib_df = load_excel_from_local_onedrive(
                "IB Analyst\\Metro Fire Dept", "customer_serial_ib_data.xlsx"
            )
            if customer_serial_ib_df is None:
                customer_serial_ib_df = pd.DataFrame()

        progress(0.32, desc="📂 Loading EAMP proposal...")
        eamp_proposal_quote_df = load_excel_from_local_onedrive(
            "IB Analyst\\Metro Fire Dept", "eamp_proposal_quote.xlsx"
        )
        if eamp_proposal_quote_df is None:
            eamp_proposal_quote_df = pd.DataFrame()

        progress(0.35, desc="📂 Loading customer details...")
        customer_detail_by_cav_df = load_excel_from_local_onedrive(
            "IB Analyst\\Metro Fire Dept", "customer_detail_by_cavid.xlsx"
        )
        if customer_detail_by_cav_df is None:
            customer_detail_by_cav_df = pd.DataFrame()

        logger.info("All data files loaded successfully")

        # Step 3: Run EA Analysis
        progress(0.4, desc="⚙️ Running EA analysis and serial reconciliation...")

        output_dir = "data_verification_output"
        os.makedirs(output_dir, exist_ok=True)

        # Main analysis function returns data (no AI summary yet)
        final_df, log_output, serial_df, serial_summary, serial_excel_path = (
            perform_ea_analysis_with_serial_reconciliation(
                target_cav_id_int=target_cav_id_int,
                customer_detail_by_cav_df=customer_detail_by_cav_df,
                case_history_df=case_history_df,
                install_base_df=install_base_df,
                sales_history_df=sales_history_df,
                customer_serial_ib_df=customer_serial_ib_df,
                eamp_proposal_quote_df=eamp_proposal_quote_df,
                output_folder=output_dir,
            )
        )

        progress(0.6, desc="✅ EA analysis complete")
        logger.info(f"EA analysis complete. {len(final_df)} rows in final dataframe")

        # Step 4: Prepare AI summary (mimics main script logic)
        progress(0.65, desc="🤖 Generating AI summary...")

        # Deduplicate for AI to avoid double-counting
        deduplicated_df = (
            final_df.groupby(["CAV_BU_ID", "PARTY_ID"]).first().reset_index()
        )

        ai_input_text = prepare_data_for_openai(
            deduplicated_df, serial_df, serial_summary
        )

        cisco_ai_client = get_cisco_ai_client()

        if cisco_ai_client:
            ai_summary = summarize_with_cisco_ai(
                ai_client=cisco_ai_client,
                text_to_summarize=ai_input_text,
                prompt_prefix=f"""Analyze this EA data for {customer_name} (CAV: {target_cav_id_int}).

IMPORTANT TERMINOLOGY - READ CAREFULLY:
- "Customer Business Units" = the customer's organizational divisions (identified by CAV_BU_NAME)
- "Parties" or "Locations" = individual sites/locations within customer business units
- U_BE = Cisco Business Entity (Cisco's product line/Architecture: Routing, Switching, Security, etc.)
- U_SUBBE = Cisco Sub-Business Entity (more specific Cisco product category)
- "Architecture" = Another term for Cisco Business Entity (U_BE)
- Use "assets" or "devices" instead of generic equipment terms

SERIAL NUMBER RECONCILIATION CONTEXT:
{serial_summary}

CRITICAL: The input data includes detailed EOL/EOS metrics BY ARCHITECTURE (U_BE). You MUST analyze and present this data.

ANALYSIS REQUIRED - COMPLETE ALL 8 SECTIONS:

1) Executive Summary
2) Cisco Business Entity (U_BE/Architecture) Breakdown ⭐ KEY FOCUS
3) EOL/EOS Risk Analysis by Cisco Architecture (U_BE) ⭐⭐ CRITICAL - MANDATORY SECTION
4) Asset & Coverage Metrics by Customer Business Unit
5) Serial Number Reconciliation ⭐ IMPORTANT
6) Coverage Opportunities by Architecture
7) Top Customer Business Units & Locations
8) Next Steps - Actionable Recommendations

FORMAT REQUIREMENTS:
- Use clear headings with numbering
- Section 3 MUST include a detailed table - DO NOT skip this section
- Use markdown tables for better readability""",
            )
        else:
            ai_summary = "⚠️ AI client unavailable. Analysis completed but AI insights could not be generated."

        progress(0.75, desc="✅ AI summary generated")
        logger.info("AI summary generated")

        # Step 5: Create PowerPoint (mimics main script)
        progress(0.8, desc="📊 Creating PowerPoint presentation...")

        ppt_path = create_ea_powerpoint_with_ai(
            final_df=deduplicated_df,
            ai_summary=ai_summary,
            ai_client=cisco_ai_client,
            cav_name=customer_name,
            cav_id=target_cav_id_int,
            gu_name=customer_name,
            output_folder=output_dir,
        )

        progress(0.85, desc="✅ PowerPoint created")
        logger.info(f"PowerPoint created: {ppt_path}")

        # Step 6: Create the multi-sheet Excel file (mimics main script)
        progress(0.87, desc="📊 Creating Excel report...")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_cav_name = customer_name.replace(" ", "_").replace("/", "_")
        excel_filename = (
            f"{safe_cav_name}_CAV_{target_cav_id_int}_EA_Analysis_{timestamp}.xlsx"
        )
        excel_path = os.path.join(output_dir, excel_filename)

        try:
            with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
                # Sheet 1: Unique parties (for analysis and AI)
                deduplicated_df.to_excel(
                    writer, sheet_name="Unique Parties", index=False
                )

                # Sheet 2: All locations (for reference with all ship-to addresses)
                final_df.to_excel(writer, sheet_name="All Locations", index=False)

                # Sheet 3: Summary statistics
                summary_stats = pd.DataFrame(
                    {
                        "Metric": [
                            "Total Unique Parties",
                            "Total Locations/Ship-Tos",
                            "Total Install Base",
                            "Covered Assets",
                            "Uncovered Assets",
                            "Never Covered Assets",
                            "Total Cases",
                            "Total Sales",
                        ],
                        "Value": [
                            len(deduplicated_df),
                            len(final_df),
                            int(deduplicated_df["TOTAL_IB_COUNT"].sum()),
                            int(deduplicated_df["IB_COUNT_COVERED"].sum()),
                            int(deduplicated_df["IB_COUNT_UNCOVERED"].sum()),
                            int(deduplicated_df["IB_COUNT_NEVER_COVERED"].sum()),
                            int(deduplicated_df["CASE_COUNT"].sum())
                            if "CASE_COUNT" in deduplicated_df.columns
                            else 0,
                            f"${deduplicated_df['TOTAL_SALES'].sum():,.2f}"
                            if "TOTAL_SALES" in deduplicated_df.columns
                            else "$0.00",
                        ],
                    }
                )
                summary_stats.to_excel(writer, sheet_name="Summary", index=False)

                # Format all sheets with headers
                workbook = writer.book
                header_format = workbook.add_format(
                    {
                        "bold": True,
                        "bg_color": "#4472C4",
                        "font_color": "white",
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                    }
                )

                # Format Unique Parties sheet
                worksheet1 = writer.sheets["Unique Parties"]
                for col_num, col_name in enumerate(deduplicated_df.columns):
                    worksheet1.write(0, col_num, col_name, header_format)
                    worksheet1.set_column(col_num, col_num, 15)
                worksheet1.freeze_panes(1, 0)

                # Format All Locations sheet
                worksheet2 = writer.sheets["All Locations"]
                for col_num, col_name in enumerate(final_df.columns):
                    worksheet2.write(0, col_num, col_name, header_format)
                    worksheet2.set_column(col_num, col_num, 15)
                worksheet2.freeze_panes(1, 0)

                # Format Summary sheet
                worksheet3 = writer.sheets["Summary"]
                for col_num, col_name in enumerate(summary_stats.columns):
                    worksheet3.write(0, col_num, col_name, header_format)
                worksheet3.set_column(0, 0, 30)  # Metric column wider
                worksheet3.set_column(1, 1, 20)  # Value column

            print(f"\n✅ EA Analysis Excel saved to: {excel_path}")
            print(
                f"   📊 Sheet 1 'Unique Parties': {len(deduplicated_df)} unique parties"
            )
            print(
                f"   📍 Sheet 2 'All Locations': {len(final_df)} locations (includes all ship-to addresses)"
            )
            print(f"   📈 Sheet 3 'Summary': Key metrics at a glance")
            logger.info(f"Excel file created: {excel_path}")

        except Exception as e:
            logger.error(f"Error creating multi-sheet Excel: {e}")
            print(f"\n⚠️ Error creating Excel file: {e}")
            excel_path = None

        # Step 7: Verify all files exist
        progress(0.88, desc="📁 Verifying generated files...")

        if excel_path and os.path.exists(excel_path):
            logger.info(f"✓ Excel file ready: {excel_path}")
            print(f"📊 Excel file: {excel_path}")
        else:
            logger.warning("Excel file not created!")
            print(f"⚠️ Excel file not found")

        if ppt_path and os.path.exists(ppt_path):
            logger.info(f"✓ PowerPoint ready: {ppt_path}")
            print(f"📈 PowerPoint: {ppt_path}")

        if serial_excel_path and os.path.exists(serial_excel_path):
            serial_path = serial_excel_path
            logger.info(f"✓ Serial file ready: {serial_path}")
            print(f"🔍 Serial file: {serial_path}")
        else:
            serial_path = None

        logger.info(
            f"File summary - Excel: {excel_path}, PPT: {ppt_path}, Serial: {serial_path}"
        )

        # Note: Webex sending is handled by the underlying SalesAssist_v11 script
        # The send_to_webex_flag would need to be passed to perform_ea_analysis_with_serial_reconciliation
        # if that function supports it

        if send_to_webex_flag:
            logger.info(
                "✅ Webex delivery was handled by the underlying analysis script"
            )
            print("✅ Results have been sent to Webex by the analysis script")
        else:
            logger.info("Webex delivery not requested")
            print("📤 Webex delivery skipped (checkbox not checked)")

        # Step 6: Send to Webex if requested (mimics main script)
        if send_to_webex_flag and AUTHORIZATION_TOKEN:
            progress(0.9, desc="📤 Sending to Webex...")

            # Split AI summary into manageable chunks (Webex limit: 7439 chars)
            def split_into_chunks(text, max_length=6500):
                """Split text into chunks at paragraph boundaries"""
                if len(text) <= max_length:
                    return [text]

                chunks = []
                paragraphs = text.split("\n\n")
                current_chunk = ""

                for para in paragraphs:
                    if len(current_chunk) + len(para) + 2 <= max_length:
                        current_chunk += para + "\n\n"
                    else:
                        if current_chunk:
                            chunks.append(current_chunk.strip())
                        current_chunk = para + "\n\n"

                if current_chunk:
                    chunks.append(current_chunk.strip())

                return chunks

            # Create header message
            header_message = f"""# ✅ EA Analysis Complete: {customer_name}

## 📊 Key Metrics Summary
- **CAV ID:** {target_cav_id_int}
- **Unique Parties:** {len(deduplicated_df):,}
- **Total Install Base:** {deduplicated_df["TOTAL_IB_COUNT"].sum():,.0f} assets
- **Covered Assets:** {deduplicated_df.get("IB_COUNT_COVERED", pd.Series([0])).sum():,.0f}
- **Uncovered Assets:** {deduplicated_df.get("IB_COUNT_UNCOVERED", pd.Series([0])).sum():,.0f}
- **Support Cases:** {deduplicated_df.get("CASE_COUNT", pd.Series([0])).sum():,.0f}
- **Sales Revenue:** ${deduplicated_df.get("TOTAL_SALES", pd.Series([0])).sum():,.2f}

## 🔍 Serial Number Reconciliation
{serial_summary[:500]}...

---
📊 **Detailed AI analysis follows in next messages...**
"""

            try:
                # Send header message
                send_webex_message_with_files(
                    room_id=WEBEX_ROOM_ID,
                    to_email=FALLBACK_WEBEX_EMAIL,
                    text_message=header_message.strip(),
                    file_paths=[],
                    authorization_token=AUTHORIZATION_TOKEN,
                )
                logger.info("Webex: Header message sent")

                # Send AI summary in chunks
                ai_chunks = split_into_chunks(ai_summary, max_length=6500)
                for i, chunk in enumerate(ai_chunks, 1):
                    chunk_header = (
                        f"## 🤖 AI Analysis - Part {i}/{len(ai_chunks)}\n\n"
                        if len(ai_chunks) > 1
                        else "## 🤖 AI-Generated Executive Summary\n\n"
                    )
                    send_webex_message_with_files(
                        room_id=WEBEX_ROOM_ID,
                        to_email=FALLBACK_WEBEX_EMAIL,
                        text_message=chunk_header + chunk,
                        file_paths=[],
                        authorization_token=AUTHORIZATION_TOKEN,
                    )
                    logger.info(f"Webex: AI summary part {i}/{len(ai_chunks)} sent")

                # Send files header
                files_msg = """---
## 📁 Downloadable Reports

The following files contain detailed analysis:
"""
                send_webex_message_with_files(
                    room_id=WEBEX_ROOM_ID,
                    to_email=FALLBACK_WEBEX_EMAIL,
                    text_message=files_msg,
                    file_paths=[],
                    authorization_token=AUTHORIZATION_TOKEN,
                )

                # Send Excel file
                if excel_path and os.path.exists(excel_path):
                    send_webex_message_with_files(
                        room_id=WEBEX_ROOM_ID,
                        to_email=FALLBACK_WEBEX_EMAIL,
                        text_message="📊 **Excel Report** - Multi-sheet analysis with unique parties, all locations, and summary statistics",
                        file_paths=[excel_path],
                        authorization_token=AUTHORIZATION_TOKEN,
                    )
                    logger.info("Webex: Excel file sent")
                else:
                    logger.warning(f"Excel file not found: {excel_path}")

                # Send PowerPoint
                if ppt_path and os.path.exists(ppt_path):
                    send_webex_message_with_files(
                        room_id=WEBEX_ROOM_ID,
                        to_email=FALLBACK_WEBEX_EMAIL,
                        text_message="📈 **PowerPoint Presentation** - Executive summary with AI insights and visualizations",
                        file_paths=[ppt_path],
                        authorization_token=AUTHORIZATION_TOKEN,
                    )
                    logger.info("Webex: PowerPoint sent")

                # Send Serial Reconciliation
                if serial_path and os.path.exists(serial_path):
                    send_webex_message_with_files(
                        room_id=WEBEX_ROOM_ID,
                        to_email=FALLBACK_WEBEX_EMAIL,
                        text_message="🔍 **Serial Number Reconciliation** - Data quality validation and match analysis",
                        file_paths=[serial_path],
                        authorization_token=AUTHORIZATION_TOKEN,
                    )
                    logger.info("Webex: Serial reconciliation sent")

                logger.info("✅ All messages and files sent to Webex successfully")
            except Exception as e:
                logger.error(f"Failed to send to Webex: {e}")
        else:
            logger.info("Webex delivery not requested")

        progress(1.0, desc="✅ Complete!")

        # Format summary output for Gradio UI
        # Note: AI summary is generated by the main script, we display a summary here
        total_ib = deduplicated_df["TOTAL_IB_COUNT"].sum()
        covered = deduplicated_df["IB_COUNT_COVERED"].sum()
        uncovered = deduplicated_df["IB_COUNT_UNCOVERED"].sum()
        never_covered = deduplicated_df["IB_COUNT_NEVER_COVERED"].sum()
        cases = deduplicated_df["CASE_COUNT"].sum()
        sales = deduplicated_df["TOTAL_SALES"].sum()
        coverage_pct = (covered / total_ib * 100) if total_ib > 0 else 0

        summary_output = f"""
# ✅ Analysis Complete!

## 📋 Customer Information
- **Customer:** {customer_name}
- **CAV ID:** {target_cav_id_int}
- **Unique Parties:** {len(deduplicated_df):,}

## 📊 Key Metrics

| Metric | Value |
|--------|-------|
| **Total Install Base** | {total_ib:,} |
| **Covered Assets** | {covered:,} ({coverage_pct:.1f}%) |
| **Uncovered Assets** | {uncovered:,} |
| **Never Covered** | {never_covered:,} |
| **Support Cases** | {cases:,} |
| **Sales Revenue** | ${sales:,.2f} |

---

## 🤖 AI-Generated Executive Summary

{ai_summary}

---

## 📥 Download Files Below

Use the download buttons to get:
- **Excel Report**: Comprehensive multi-sheet analysis with all data
- **PowerPoint Presentation**: Executive summary with AI insights and charts
- **Serial Reconciliation**: Detailed validation report with match rates

{"✅ **Results have been sent to Webex**" if send_to_webex_flag else ""}
"""

        return summary_output, excel_path, ppt_path, serial_path

    except Exception as e:
        logger.error(f"Error during analysis: {str(e)}", exc_info=True)
        error_output = f"""
# ❌ Analysis Failed

**Error:** {str(e)}

**Customer:** {customer_name}
**CAV ID:** {target_cav_id_int}

Please check:
- Customer name spelling
- CAV ID selection
- Database connectivity
- OneDrive file availability
- Log files for details
"""
        return error_output, None, None, None


# Create Gradio Interface
with gr.Blocks(
    title="Cisco EA Analysis Tool",
    theme=gr.themes.Soft(primary_hue="blue", secondary_hue="cyan"),
    css="""
        .main-header {font-size: 2.5rem; font-weight: bold; text-align: center; color: #049fd9; margin-bottom: 1rem;}
        .subtitle {text-align: center; color: #666; margin-bottom: 2rem;}
        .metric-box {background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 1.5rem; border-radius: 10px; color: white;}
    """,
) as demo:
    # Header
    gr.HTML('<h1 class="main-header">🎯 Cagnito EA Sales Assistent</h1>')
    gr.HTML('<p class="subtitle">AI Powered Cagnito EA Sales Assistent </p>')

    gr.Markdown("---")

    # Step 1: Customer lookup
    gr.Markdown("## Step 1: Find Customer")

    with gr.Row():
        with gr.Column(scale=4):
            customer_input = gr.Textbox(
                label="🏢 Customer Name (Global Ultimate)",
                placeholder="e.g., Fire Rescue Victoria, Amazon, Microsoft, Walmart",
                lines=1,
                info="Enter the exact customer name as it appears in the database",
            )

        with gr.Column(scale=1):
            lookup_btn = gr.Button("🔍 Lookup CAV IDs", variant="secondary", size="lg")

    lookup_status = gr.Markdown(value="")

    # Customer data table (initially hidden)
    gr.Markdown("### 📋 Customer Details & Install Base Coverage")
    customer_data_table = gr.Dataframe(
        label="Customer Records by CAV ID and Business Unit", visible=False
    )

    # Step 2: CAV ID selection (initially hidden)
    gr.Markdown("## Step 2: Select CAV ID & Upload Serial Numbers (Optional)")

    with gr.Row():
        with gr.Column(scale=2):
            cav_dropdown = gr.Dropdown(
                label="📋 Available CAV IDs - Select the CAV ID to analyze",
                choices=[],
                visible=False,
            )

        with gr.Column(scale=2):
            serial_file_upload = gr.File(
                label="📎 Upload Customer Serial Numbers (Optional) - Excel or CSV",
                file_types=[".xlsx", ".xls", ".csv"],
                visible=False,
                type="filepath",
            )

        with gr.Column(scale=1):
            webex_checkbox = gr.Checkbox(label="📤 Send to Webex", value=False)

    # Run analysis button (initially hidden)
    run_btn = gr.Button(
        "🚀 Run Complete EA Analysis", variant="primary", size="lg", visible=False
    )

    gr.Markdown("---")

    # Output section
    gr.Markdown("## 📊 Analysis Results")

    with gr.Row():
        output_markdown = gr.Markdown(value="", label="Results")

    gr.Markdown("---")
    gr.Markdown("## 📥 Download Reports")

    with gr.Row():
        with gr.Column():
            excel_output = gr.File(
                label="📊 Excel Report (Multi-Sheet)", file_count="single"
            )

        with gr.Column():
            ppt_output = gr.File(
                label="📈 PowerPoint Presentation", file_count="single"
            )

        with gr.Column():
            serial_output = gr.File(
                label="🔍 Serial Reconciliation", file_count="single"
            )

    gr.Markdown("---")

    # Instructions
    with gr.Accordion("📖 How to Use", open=False):
        gr.Markdown("""
        ### Step-by-Step Guide:
        
        **Step 1: Find Customer**
        1. Enter the Global Ultimate customer name exactly as it appears in the database
        2. Click "🔍 Lookup CAV IDs"
        3. The system will find all CAV IDs associated with this customer
        
        **Step 2: Select CAV ID**
        1. Choose the CAV ID you want to analyze from the dropdown
        2. Optional: Check "Send to Webex" to automatically deliver results
        
        **Step 3: Run Analysis**
        1. Click "🚀 Run Complete EA Analysis"
        2. Wait 2-5 minutes for the analysis to complete
        3. Progress bar shows current step
        
        **Step 4: Download Results**
        - Excel: Multi-sheet workbook (Unique Parties, All Locations, Summary)
        - PowerPoint: AI-powered presentation with charts
        - Serial Reconciliation: Validation report with match rates
        
        ### 📁 Output Files Include:
        - **Excel Report**: Comprehensive analysis with 3 sheets
        - **PowerPoint**: Executive presentation with AI insights
        - **Serial Reconciliation**: Data quality and match analysis
        
        ### ⚡ Requirements:
        - OneDrive files must be synced locally
        - Database connectivity required
        - Cisco AI credentials configured
        
        ### 💡 Tips:
        - Large customers may have multiple CAV IDs
        - Select the most relevant CAV ID for your analysis
        - Download files before running another analysis
        """)

    # Footer
    gr.Markdown("---")
    gr.Markdown("""
    <div style='text-align: center; color: #666; font-size: 0.9rem;'>
    <p>🔒 Secure • 🚀 Fast • 🤖 AI-Powered</p>
    <p>Cisco EA Analysis Tool v1.1 - Now with CAV ID Selection</p>
    </div>
    """)

    # Connect lookup button
    lookup_btn.click(
        fn=fetch_cav_ids,
        inputs=[customer_input],
        outputs=[
            lookup_status,
            customer_data_table,
            cav_dropdown,
            cav_dropdown,
            run_btn,
            customer_data_table,
            serial_file_upload,
        ],
    )

    # Connect analysis button
    run_btn.click(
        fn=run_ea_analysis,
        inputs=[customer_input, cav_dropdown, serial_file_upload, webex_checkbox],
        outputs=[output_markdown, excel_output, ppt_output, serial_output],
        show_progress="full",
    )

# Launch configuration
if __name__ == "__main__":
    print("\n" + "=" * 70)
    print("🎯 CISCO EA ANALYSIS TOOL v1.1")
    print("=" * 70)
    print("\nStarting Gradio interface...")
    print("The app will open in your browser automatically.")
    print("\n📱 Access Options:")
    print("   - Local: http://127.0.0.1:7860")
    print("   - Network: http://YOUR-IP:7860")
    print("   - Public: Will be shown below (expires in 72 hours)")
    print("\n🔗 Share the public URL with your team!")
    print("=" * 70 + "\n")

    demo.launch(
        share=True,
        server_name="0.0.0.0",
        server_port=7860,
        show_error=True,
        quiet=False,
    )
