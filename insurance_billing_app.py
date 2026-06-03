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

st.set_page_config(page_title="Insurance Billing Auto-Entry", layout="wide")
st.title("📧 Email → Compliant Insurance Billing Entry (A+L LEDES/UTBMS)")

# Helper: Extract ONLY names from subject line for matter name
def extract_matter_names(subject):
    if not subject:
        return ""
    subject = re.sub(r'^(Re:|FW:|Fwd:|RE:|FW:)\s*', '', subject, flags=re.I).strip()
    pattern = r'\b[A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)*\b'
    candidates = re.findall(pattern, subject)
    exclude = {'Re', 'FW', 'Fwd', 'Vs', 'V', 'The', 'And', 'Claim', 'Matter', 'Case'}
    names = [c for c in candidates if c not in exclude]
    return " ".join(names) if names else ""

api_key = st.text_input("OpenAI API Key (sk-...)", type="password", value="")
if not api_key:
    st.warning("Enter your OpenAI API key to enable smart parsing")
    st.stop()

client = OpenAI(api_key=api_key)

timekeeper = st.text_input("Timekeeper Name", value="S. Michael Cooper")

if "billing_entries" not in st.session_state:
    st.session_state.billing_entries = []

st.subheader("🎯 Drag & drop MULTIPLE .msg or .eml emails here")
st.caption("**Outlook tip**: Just drag the .msg files directly from Outlook — the app processes PDF, Word, and Excel attachments (ignores only jpeg/png images). Accurate PDF page counts are included. No need to save as .eml!")

uploaded_files = st.file_uploader(
    "Upload .msg or .eml email files (multiple supported)",
    accept_multiple_files=True,
    type=["msg", "eml"]
)

if st.button("🔥 Generate One Compliant Entry Per Email", type="primary"):
    if not uploaded_files:
        st.warning("Please upload at least one email file")
    else:
        with st.spinner("AI is analyzing every email + attachments..."):
            new_count = 0
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
                                    
                                    # Ignore only jpeg and png image files
                                    if ext in ['jpg', 'jpeg', 'png']:
                                        continue
                                    
                                    if ext == 'pdf':
                                        # PDF: accurate page count + full text extraction
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
                                        # Word, Excel, or other allowed attachments - reference by name/type
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
                    else:  # .msg file
                        outlook_msg = extract_msg.Message(io.BytesIO(file.getvalue()))
                        subject_line = outlook_msg.subject or ""
                        from_addr = outlook_msg.sender or ""
                        email_date_str = outlook_msg.date
                        body = outlook_msg.body or ""
                        attachment_text = ""
                        for att in outlook_msg.attachments:
                            att_name = (
                                getattr(att, 'filename', None) or
                                getattr(att, 'longFilename', None) or
                                getattr(att, 'shortFilename', None) or
                                getattr(att, 'name', None) or
                                getattr(att, 'displayName', None) or
                                'Unknown_Attachment'
                            )
                            if not att_name:
                                continue
                            
                            ext = str(att_name).lower().split('.')[-1] if '.' in str(att_name) else ''
                            
                            # Ignore only jpeg and png image files
                            if ext in ['jpg', 'jpeg', 'png']:
                                continue
                            
                            if ext == 'pdf' and hasattr(att, 'data') and att.data:
                                # PDF: accurate page count + full text extraction
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
                                # Word, Excel, or other allowed attachments - reference by name/type
                                attachment_text += f"\n--- ATTACHMENT: {att_name} ({ext.upper()} file) ---\n"
                                try:
                                    if hasattr(att, 'data') and att.data:
                                        attachment_text += str(att.data)[:2500] + "\n"
                                except:
                                    attachment_text += "[Attachment content not extracted]\n"
                    # Get real email date as service date (or today fallback)
                    service_date = date.today()
                    if email_date_str:
                        try:
                            if isinstance(email_date_str, str):
                                dt = parsedate_to_datetime(email_date_str)
                            else:
                                dt = email_date_str
                            service_date = dt.date()
                        except:
                            pass
                    matter_reference = extract_matter_names(subject_line) or "Unknown Matter"
                    full_text = f"SUBJECT: {subject_line}\nFROM: {from_addr}\n\n{body}\n\nATTACHMENTS:\n{attachment_text}"
                    # Strong prompt for insurance-defense compliant narratives
                    system_prompt = """You are an expert insurance-defense legal biller using LEDES 1998B + UTBMS standards.
You ALWAYS respond with valid JSON only. Never include explanations outside the JSON."""
                    user_prompt = f"""Create ONE perfect daily time entry for insurance defense billing.
Use these exact values:
- timekeeper: "{timekeeper}"
- date_of_service: "{service_date}"
- matter_reference: "{matter_reference}"

CRITICAL INSTRUCTIONS FOR NARRATIVE (must follow exactly):
- Start with a strong action verb (e.g., "Reviewed", "Analyzed", "Drafted", "Corresponded with").
- Reference PDF, Word, and Excel attachments when relevant to the work performed. Never mention or describe jpeg, png, or other image files.
- For PDFs: ALWAYS use the accurate page count provided (e.g., "Reviewed the 14-page medical records PDF from Dr. Smith"). The page count is exact, not estimated.
- For Word/Excel attachments: Reference them by filename and type when they are relevant (e.g., "the attached settlement agreement.docx" or "the Excel damages spreadsheet").
- Be highly specific to the email subject, body, and allowed attachment contents.
- Max 2-3 concise sentences. NO block billing. Use insurance-carrier friendly language.
- Focus on the legal work performed and its purpose for the defense.

UTBMS CODE SUGGESTIONS (required):
- Carefully analyze the full email content (subject, body, and any PDF/Word/Excel text).
- Select the SINGLE most appropriate activity_code (A-series) and task_code (L-series) that best match the actual legal work described in THIS email.
- Base your code choices strictly on the specific tasks performed (research, drafting, correspondence, document review, etc.). Use the most specific fitting code rather than generic defaults.

REQUIRED JSON FORMAT (exact keys):
{{
  "matter_reference": "string (from subject or Unknown Matter)",
  "date_of_service": "{service_date}",
  "timekeeper": "{timekeeper}",
  "hours": number (in 0.1 increments, e.g. 0.5 or 1.2),
  "activity_code": "Best matching UTBMS activity code based on content",
  "task_code": "Best matching UTBMS task code based on content",
  "narrative": "Detailed, compliant narrative here..."
}}

Email + FULL Attachment Content:
{full_text[:25000]}

Return ONLY the JSON object. No markdown, no extra text."""
                    response = client.chat.completions.create(
                        model="gpt-4o",
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": user_prompt}
                        ],
                        temperature=0.2,
                        response_format={"type": "json_object"}
                    )
                    result = json.loads(response.choices[0].message.content.strip())
                    # Basic validation
                    if "narrative" in result and "hours" in result:
                        st.session_state.billing_entries.append(result)
                        new_count += 1
                    else:
                        st.warning(f"AI response missing required fields for {file.name}")
                except Exception as e:
                    st.error(f"Error processing {file.name}: {str(e)}")
        if new_count > 0:
            st.success(f"✅ Successfully created {new_count} new compliant billing entries!")
        else:
            st.info("No valid .msg or .eml files were processed.")

# Display section
if st.session_state.billing_entries:
    st.subheader(f"📋 Generated Billing Entries ({len(st.session_state.billing_entries)} total)")
    
    # Use data_editor for easy inline editing (newer Streamlit feature)
    edited_df = st.data_editor(
        pd.DataFrame(st.session_state.billing_entries),
        use_container_width=True,
        num_rows="dynamic",
        key="billing_editor"
    )
    
    # Update session state if edited
    if st.button("💾 Save Edits to Current Batch"):
        st.session_state.billing_entries = edited_df.to_dict("records")
        st.success("Edits saved!")
        st.rerun()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        csv = edited_df.to_csv(index=False)
        st.download_button(
            "📥 Download as CSV",
            csv,
            file_name=f"insurance_billing_{date.today()}.csv",
            mime="text/csv"
        )
    with col2:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            edited_df.to_excel(writer, index=False, sheet_name="Billing Entries")
        st.download_button(
            "📥 Download as Excel",
            buffer.getvalue(),
            file_name=f"insurance_billing_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col3:
        if st.button("🆕 Clear All Entries & Start Fresh", type="secondary"):
            st.session_state.billing_entries = []
            st.rerun()

st.caption("Creates daily time entries compliant with insurance company guidelines (LEDES 1998B / UTBMS). Processes PDF/Word/Excel attachments (ignores jpeg/png images). Accurate PDF page counts + intelligent UTBMS code suggestions. Optimized for carriers like Travelers, AmTrust, etc.")