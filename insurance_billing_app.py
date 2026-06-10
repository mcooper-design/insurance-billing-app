import streamlit as st
from openai import OpenAI
import pypdf
import io
import json
import email
from email import policy
from email.utils import parsedate_to_datetime
from datetime import date
import re
import pandas as pd
import extract_msg
import platform

st.set_page_config(page_title="Insurance Billing Auto-Entry", layout="wide")
st.title("📧 Email → Compliant Insurance Billing Entry (A+L LEDES/UTBMS)")

def extract_matter_names(subject):
    if not subject:
        return ""
    subject = re.sub(r'^(Re:|FW:|Fwd:|RE:|FW:)\s*', '', subject, flags=re.I).strip()
    pattern = r'\b[A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)*\b'
    candidates = re.findall(pattern, subject)
    exclude = {'Re', 'FW', 'Fwd', 'Vs', 'V', 'The', 'And', 'Claim', 'Matter', 'Case'}
    names = [c for c in candidates if c not in exclude]
    return " ".join(names) if names else ""

def process_email_to_billing_entry(client, full_text, service_date, matter_reference, source_name="email"):
    """Single place for the AI call (avoids duplication)"""
    system_prompt = """You are an expert insurance-defense legal biller using LEDES 1998B + UTBMS standards.
You ALWAYS respond with valid JSON only. Never include explanations outside the JSON."""

    user_prompt = f"""Create ONE billing entry for this email.
**Priority:**
1. Choose the best **task_code** (L-code) based on the main legal work.
2. Then choose an appropriate **activity_code** (A-code).
**Important Rules for Codes:**
- Return **ONLY the code** (example: "L310" or "A101").
- Do **NOT** include any description or text in parentheses next to the code.
Use these values:
- date_of_service: "{service_date}"
- matter_reference: "{matter_reference}"
CRITICAL INSTRUCTIONS:
- Never mention image files.
- For PDFs: Always include the accurate page count.
- Max 2-3 concise sentences. No block billing.
REQUIRED JSON FORMAT:
{{
  "matter_reference": "string",
  "date_of_service": "{service_date}",
  "hours": number (in 0.1 increments),
  "activity_code": "Axxx",
  "task_code": "Lxxx",
  "narrative": "Clear narrative..."
}}
Email + Attachments:
{full_text[:25000]}
Return ONLY the JSON object. No markdown or extra text."""

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0.2,
        response_format={"type": "json_object"}
    )
    return json.loads(response.choices[0].message.content.strip())

# === API Key ===
api_key = st.text_input("OpenAI API Key (sk-...)", type="password", value="")
if not api_key:
    st.warning("Enter your OpenAI API key to enable smart parsing")
    st.stop()

client = OpenAI(api_key=api_key)

if "billing_entries" not in st.session_state:
    st.session_state.billing_entries = []

# === HEADER + INSTRUCTIONS ===
st.subheader("🎯 Import Emails for Billing Entries")

with st.expander("📋 How to add .msg / .eml files (Critical for Outlook on laptop)", expanded=True):
    st.markdown("""
    **Direct drag from Outlook usually fails** on Windows laptops (especially the new Outlook app).  
    This is a Microsoft limitation — not a problem with your app.

    **Recommended workflow (works every time):**
    1. In Outlook, open the email (or select it in the list).
    2. **Drag the email to your Desktop** first → this automatically creates a `.msg` file.
    3. Then drag that `.msg` file (or multiple files) from your Desktop into the box below.

    **Alternative methods:**
    - Right-click the email → **Save As** → choose **Outlook Message Format (*.msg)** → save to Desktop → drag in.
    - Works with both `.msg` and `.eml` files.
    - The app reads the full email body + any PDF, Word, or Excel attachments (skips photos/images).

    You can drop **multiple emails at once**.
    """)

# === FILE UPLOADER ===
uploaded_files = st.file_uploader(
    "Drop .msg or .eml files here (or click to browse)",
    accept_multiple_files=True,
    type=["msg", "eml"]
)

st.markdown("---")

# === PASTE FALLBACK ===
st.subheader("📝 Or paste email content (fallback)")
pasted_content = st.text_area(
    "Paste the full email here (subject, from, date, body + attachment text)",
    height=160,
    placeholder="SUBJECT: ...\nFROM: ...\nDATE: ...\n\nBody text here...\n\n--- PDF ATTACHMENT: motion.pdf ---\nPage content...",
    key="pasted_email"
)

# === OUTLOOK DIRECT IMPORT (only on Windows) ===
can_import_from_outlook = False
if platform.system() == "Windows":
    try:
        import win32com.client
        can_import_from_outlook = True
    except ImportError:
        can_import_from_outlook = False

if can_import_from_outlook:
    st.markdown("---")
    st.subheader("📥 Or import directly from Outlook (no drag, no Save As)")

    if st.button("📥 Import Currently Selected Email from Outlook", type="secondary"):
        try:
            import tempfile
            import os

            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            explorer = outlook.Application.ActiveExplorer()
            inspector = outlook.Application.ActiveInspector()

            mail_item = None
            if inspector is not None and inspector.CurrentItem.Class == 43:
                mail_item = inspector.CurrentItem
            elif explorer is not None and explorer.Selection.Count > 0:
                selected = explorer.Selection.Item(1)
                if selected.Class == 43:
                    mail_item = selected

            if mail_item is None:
                st.warning("No email selected or open in Outlook. Please select or open an email first.")
            else:
                with st.spinner("Importing email from Outlook..."):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
                        temp_path = tmp.name
                    mail_item.SaveAs(temp_path, 3)

                    outlook_msg = extract_msg.Message(temp_path)
                    subject_line = outlook_msg.subject or ""
                    from_addr = outlook_msg.sender or ""
                    email_date_str = outlook_msg.date
                    body = outlook_msg.body or ""
                    attachment_text = ""

                    for att in outlook_msg.attachments:
                        att_name = getattr(att, 'filename', None) or getattr(att, 'longFilename', None) or 'Unknown'
                        if not att_name:
                            continue
                        ext = str(att_name).lower().split('.')[-1] if '.' in str(att_name) else ''
                        if ext in ['jpg', 'jpeg', 'png']:
                            continue
                        if ext == 'pdf' and hasattr(att, 'data') and att.data:
                            attachment_text += f"\n--- PDF ATTACHMENT: {att_name} ---\n"
                            try:
                                pdf_reader = pypdf.PdfReader(io.BytesIO(att.data))
                                page_count = len(pdf_reader.pages)
                                attachment_text += f"PDF - {page_count} pages (accurate count)\n"
                                for i, page in enumerate(pdf_reader.pages, 1):
                                    attachment_text += f"\n--- Page {i} ---\n{page.extract_text() or ''}\n"
                            except:
                                attachment_text += "[PDF text could not be extracted]\n"
                        else:
                            attachment_text += f"\n--- ATTACHMENT: {att_name} ({ext.upper()} file) ---\n"
                            try:
                                if hasattr(att, 'data') and att.data:
                                    attachment_text += str(att.data)[:2500] + "\n"
                            except:
                                attachment_text += "[Attachment content not extracted]\n"

                    try:
                        os.unlink(temp_path)
                    except:
                        pass

                    service_date = date.today()
                    if email_date_str:
                        try:
                            dt = parsedate_to_datetime(email_date_str) if isinstance(email_date_str, str) else email_date_str
                            service_date = dt.date()
                        except:
                            pass

                    matter_reference = extract_matter_names(subject_line) or "Outlook Import"
                    full_text = f"SUBJECT: {subject_line}\nFROM: {from_addr}\n\n{body}\n\nATTACHMENTS:\n{attachment_text}"

                    result = process_email_to_billing_entry(client, full_text, service_date, matter_reference, "Outlook")
                    if "narrative" in result and "hours" in result:
                        st.session_state.billing_entries.append(result)
                        st.success("✅ Email imported directly from Outlook and billing entry created!")
                    else:
                        st.warning("AI response was incomplete.")

        except Exception as e:
            st.error(f"Could not import from Outlook: {str(e)}")

# === GENERATE BUTTON (handles files + paste) ===
if st.button("🔥 Generate Billing Entry", type="primary"):
    new_count = 0

    # Process uploaded files
    if uploaded_files:
        with st.spinner("AI is analyzing uploaded email(s) + attachments..."):
            for file in uploaded_files:
                filename = file.name.lower()
                if not (filename.endswith('.eml') or filename.endswith('.msg')):
                    continue
                try:
                    if filename.endswith('.eml'):
                        msg = email.message_from_bytes(file.getvalue(), policy=policy.default)
                        subject_line = msg.get("subject") or ""
                        from_addr = msg.get("from") or ""
                        email_date_str = msg.get("Date") or msg.get("Received")
                        body = ""
                        attachment_text = ""
                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_type() == "text/plain":
                                    body = part.get_payload(decode=True).decode(errors="ignore")
                                elif part.get_filename():
                                    att_name = part.get_filename()
                                    ext = att_name.lower().split('.')[-1] if '.' in att_name else ''
                                    if ext in ['jpg', 'jpeg', 'png']:
                                        continue
                                    if ext == 'pdf':
                                        attachment_text += f"\n--- PDF ATTACHMENT: {att_name} ---\n"
                                        try:
                                            pdf_data = io.BytesIO(part.get_payload(decode=True))
                                            pdf_reader = pypdf.PdfReader(pdf_data)
                                            page_count = len(pdf_reader.pages)
                                            attachment_text += f"PDF - {page_count} pages (accurate count)\n"
                                            for i, page in enumerate(pdf_reader.pages, 1):
                                                attachment_text += f"\n--- Page {i} ---\n{page.extract_text() or ''}\n"
                                        except:
                                            attachment_text += "[PDF text could not be extracted]\n"
                                    else:
                                        attachment_text += f"\n--- ATTACHMENT: {att_name} ({ext.upper()} file) ---\n"
                                        try:
                                            content = part.get_payload(decode=True)
                                            if content:
                                                attachment_text += str(content)[:2500] + "\n"
                                        except:
                                            attachment_text += "[Attachment content not extracted]\n"
                        else:
                            payload = msg.get_payload(decode=True)
                            body = payload.decode(errors="ignore") if payload else ""
                    else:  # .msg
                        outlook_msg = extract_msg.Message(io.BytesIO(file.getvalue()))
                        subject_line = outlook_msg.subject or ""
                        from_addr = outlook_msg.sender or ""
                        email_date_str = outlook_msg.date
                        body = outlook_msg.body or ""
                        attachment_text = ""
                        for att in outlook_msg.attachments:
                            att_name = getattr(att, 'filename', None) or getattr(att, 'longFilename', None) or 'Unknown'
                            if not att_name:
                                continue
                            ext = str(att_name).lower().split('.')[-1] if '.' in str(att_name) else ''
                            if ext in ['jpg', 'jpeg', 'png']:
                                continue
                            if ext == 'pdf' and hasattr(att, 'data') and att.data:
                                attachment_text += f"\n--- PDF ATTACHMENT: {att_name} ---\n"
                                try:
                                    pdf_reader = pypdf.PdfReader(io.BytesIO(att.data))
                                    page_count = len(pdf_reader.pages)
                                    attachment_text += f"PDF - {page_count} pages (accurate count)\n"
                                    for i, page in enumerate(pdf_reader.pages, 1):
                                        attachment_text += f"\n--- Page {i} ---\n{page.extract_text() or ''}\n"
                                except:
                                    attachment_text += "[PDF text could not be extracted]\n"
                            else:
                                attachment_text += f"\n--- ATTACHMENT: {att_name} ({ext.upper()} file) ---\n"
                                try:
                                    if hasattr(att, 'data') and att.data:
                                        attachment_text += str(att.data)[:2500] + "\n"
                                except:
                                    attachment_text += "[Attachment content not extracted]\n"

                    service_date = date.today()
                    if email_date_str:
                        try:
                            dt = parsedate_to_datetime(email_date_str) if isinstance(email_date_str, str) else email_date_str
                            service_date = dt.date()
                        except:
                            pass

                    matter_reference = extract_matter_names(subject_line) or "Unknown Matter"
                    full_text = f"SUBJECT: {subject_line}\nFROM: {from_addr}\n\n{body}\n\nATTACHMENTS:\n{attachment_text}"

                    result = process_email_to_billing_entry(client, full_text, service_date, matter_reference, file.name)
                    if "narrative" in result and "hours" in result:
                        st.session_state.billing_entries.append(result)
                        new_count += 1
                    else:
                        st.warning(f"AI response missing required fields for {file.name}")

                except Exception as e:
                    st.error(f"Error processing {file.name}: {str(e)}")

    # Process pasted content
    if pasted_content and pasted_content.strip():
        with st.spinner("AI is analyzing pasted email..."):
            try:
                subject_line = ""
                from_addr = ""
                for line in pasted_content.split('\n')[:15]:
                    lower = line.lower().strip()
                    if lower.startswith("subject:"):
                        subject_line = line.split(":", 1)[1].strip()
                    elif lower.startswith("from:"):
                        from_addr = line.split(":", 1)[1].strip()

                service_date = date.today()
                matter_reference = extract_matter_names(subject_line) or "Pasted Email"
                full_text = pasted_content

                result = process_email_to_billing_entry(client, full_text, service_date, matter_reference, "pasted")
                if "narrative" in result and "hours" in result:
                    st.session_state.billing_entries.append(result)
                    new_count += 1
                else:
                    st.warning("AI response missing required fields for pasted email")
            except Exception as e:
                st.error(f"Error processing pasted email: {str(e)}")

    if new_count > 0:
        st.success(f"✅ Successfully created {new_count} billing entry(ies)!")

# === DISPLAY SECTION ===
if st.session_state.billing_entries:
    st.subheader(f"📋 Generated Billing Entries ({len(st.session_state.billing_entries)} total)")

    df = pd.DataFrame(st.session_state.billing_entries)
    column_order = ['matter_reference', 'date_of_service', 'hours', 'task_code', 'activity_code', 'narrative']
    df = df[column_order]

    edited_df = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        key="billing_editor"
    )

    if st.button("💾 Save Edits to Current Batch"):
        st.session_state.billing_entries = edited_df.to_dict("records")
        st.success("Edits saved!")
        st.rerun()

    col1, col2, col3 = st.columns(3)
    with col1:
        csv = edited_df.to_csv(index=False)
        st.download_button("📥 Download as CSV", csv, file_name=f"insurance_billing_{date.today()}.csv", mime="text/csv")
    with col2:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            edited_df.to_excel(writer, index=False, sheet_name="Billing Entries")
        st.download_button("📥 Download as Excel", buffer.getvalue(), file_name=f"insurance_billing_{date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col3:
        if st.button("🆕 Clear All Entries & Start Fresh", type="secondary"):
            st.session_state.billing_entries = []
            st.rerun()

st.caption("One billing entry per email. L-code shown before A-code. Clean codes only (no descriptions).")
