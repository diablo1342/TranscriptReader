#!/usr/bin/env python3

import os
import sys
import json
import requests
import streamlit as st
from msal import PublicClientApplication
from openai import OpenAI


GRAPH_SCOPE = ["Mail.Send", "User.Read", "OnlineMeetings.Read"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_BETA = "https://graph.microsoft.com/beta"



def _save_json(payload: dict, path: str, debug: bool):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
        if debug:
            st.write(f"[DEBUG] Saved JSON -> {path}")
            st.json(payload)
    except Exception as e:
        st.warning(f"Failed to save {path}: {e}")

def acquire_user_token_device_code(client_id: str, tenant_id: str) -> str:
    app = PublicClientApplication(client_id=client_id, authority=f"https://login.microsoftonline.com/{tenant_id}")
    flow = app.initiate_device_flow(scopes=GRAPH_SCOPE)
    if "user_code" not in flow:
        raise RuntimeError("Failed to initiate device code flow.")
    st.info(f"Go to {flow['verification_uri']} and enter code: {flow['user_code']}")
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Could not acquire token: {result}")
    return result["access_token"]



def resolve_meeting_id_from_link(access_token: str, meeting_link: str, debug: bool) -> str:
    url = f"{GRAPH_BASE}/me/onlineMeetings?$filter=JoinWebUrl eq '{meeting_link}'"
    headers = {"Authorization": f"Bearer {access_token}"}
    r = requests.get(url, headers=headers, timeout=30)

    try:
        data = r.json()
    except Exception:
        data = {"_non_json_response": r.text}

    _save_json(data, "graph_meeting_lookup.json", debug)

    if r.status_code != 200:
        raise RuntimeError(f"Failed to resolve meeting ID: {r.status_code} {r.text}")

    if "value" not in data or not data["value"]:
        raise RuntimeError("No meeting found for the given link (lookup returned empty).")

    return data["value"][0]["id"]


def fetch_transcript_by_id(access_token: str, meeting_id: str, debug: bool) -> str:
    headers = {"Authorization": f"Bearer {access_token}"}
    tried = []

    def _download_content_stream(url: str, label: str) -> str:
        accept_candidates = ["text/plain", "text/vtt", "application/vnd.ms-graph.vtt", "application/octet-stream"]
        last_err = None
        for accept in accept_candidates:
            try:
                r = requests.get(
                    url,
                    headers={"Authorization": headers["Authorization"], "Accept": accept},
                    timeout=60,
                )
                if debug:
                    st.write(f"[DEBUG] GET {label} content (Accept: {accept}) -> {r.status_code}")
                if r.status_code == 200 and r.content is not None:
                    try:
                        text = r.content.decode("utf-8", errors="replace").strip()
                    except Exception:
                        text = r.text.strip()
                    return text
                last_err = f"{r.status_code} {r.text}"
            except Exception as e:
                last_err = str(e)
        raise RuntimeError(f"Failed to download transcript content stream ({label}). Last error: {last_err}")

    for base, dump_path in (
        (GRAPH_BASE, "graph_transcripts_v1.json"),
        (GRAPH_BETA, "graph_transcripts_beta.json"),
    ):
        url = f"{base}/me/onlineMeetings/{meeting_id}/transcripts"
        r = requests.get(url, headers=headers, timeout=30)
        tried.append((base, r.status_code))
        try:
            data = r.json()
        except Exception:
            data = {"_non_json_response": r.text}
        _save_json(data, dump_path, debug)

        if r.status_code == 200 and isinstance(data, dict) and data.get("value"):
            t0 = data["value"][0]
            inline = t0.get("content")
            if inline:
                return inline.strip()
            stream_url = t0.get("transcriptContentUrl")
            if stream_url:
                return _download_content_stream(stream_url, base)

    details = "; ".join([f"{b}={code}" for b, code in tried])
    raise RuntimeError(f"No usable transcript returned. Statuses: {details}.")


def summarize_text_with_openai(text: str) -> str:
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not set.")
    client = OpenAI(api_key=api_key)
    prompt = f"""
You are a helpful assistant. Analyze the following meeting transcript and create a structured summary.

Instructions:
- Identify the 3 most important topics discussed in the meeting.
- For each topic, provide 3–5 concise bullet points.
- Format the response as clean HTML with:
  - <h2> for topic headers
  - <ul><li> for bullet points
- Make it professional, suitable for sending as a meeting recap email.

Transcript:
{text}
"""
    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You summarize Teams meeting transcripts."},
            {"role": "user", "content": prompt},
        ],
    )
    summary = resp.choices[0].message.content


    if summary.startswith("```html"):
        summary = summary.strip("`")  # strip backticks
        summary = summary.replace("html", "", 1).strip()  # remove 'html' after backticks
        if summary.endswith("```"):
            summary = summary[:-3].strip()

    return summary





def send_email(access_token: str, recipients: list[str], subject: str, body_html: str):
    url = f"{GRAPH_BASE}/me/sendMail"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    message = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": addr.strip()}} for addr in recipients],
        }
    }
    r = requests.post(url, headers=headers, json=message, timeout=30)
    if r.status_code not in (200, 202):
        raise RuntimeError(f"Graph sendMail failed: {r.status_code} {r.text}")
    st.success(f"✅ Email sent to {', '.join(recipients)}")



st.title("📊 Teams Meeting Summarizer")

meeting_link = st.text_input("Paste Teams Meeting Link:")
recipients = st.text_input("Recipients (comma-separated emails):")
subject = st.text_input("Email Subject:", "Teams Call Summary")
debug = st.checkbox("Enable debug logs", value=False)

if st.button("Run Summarizer"):
    try:
        client_id = os.getenv("AZURE_CLIENT_ID")
        tenant_id = os.getenv("AZURE_TENANT_ID")
        if not client_id or not tenant_id:
            st.error("AZURE_CLIENT_ID and AZURE_TENANT_ID must be set as environment variables.")
            st.stop()

        st.write("🔑 Getting Graph token...")
        token = acquire_user_token_device_code(client_id, tenant_id)

        st.write("📌 Resolving meeting...")
        meeting_id = resolve_meeting_id_from_link(token, meeting_link, debug)
        st.success(f"Meeting ID: {meeting_id}")

        st.write("📝 Fetching transcript...")
        transcript = fetch_transcript_by_id(token, meeting_id, debug)

        st.write("🤖 Summarizing transcript with OpenAI...")
        summary_html = summarize_text_with_openai(transcript)

        st.markdown("### 📄 AI Summary")
        st.markdown(summary_html, unsafe_allow_html=True)

        st.write("📧 Sending email...")
        send_email(token, recipients.split(","), subject, summary_html)

    except Exception as e:
        st.error(f"Error: {e}")
