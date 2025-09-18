#!/usr/bin/env python3

import argparse
import os
import sys
import textwrap
import json
import requests
from msal import PublicClientApplication

# --- OpenAI (simple SDK usage) ---
try:
    from openai import OpenAI
except ImportError:
    print("Please pip install openai.", file=sys.stderr)
    raise

# ✅ Correct Graph scopes (no transcript-only scope)
GRAPH_SCOPE = ["Mail.Send", "User.Read", "OnlineMeetings.Read"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_BETA = "https://graph.microsoft.com/beta"


def _save_json(payload: dict, path: str, debug: bool):
    """Save JSON to disk and optionally print a trimmed preview."""
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
        if debug:
            print(f"[DEBUG] Saved JSON -> {path}")
            print(json.dumps(payload, indent=2)[:2000])  # avoid flooding console
    except Exception as e:
        print(f"[WARN] Failed to save {path}: {e}")


def resolve_meeting_id_from_link(access_token: str, meeting_link: str, debug: bool) -> str:
    """
    Use the joinWebUrl (full meeting link) to look up the Graph meeting id.
    Requires OnlineMeetings.Read (delegated) and that the signed-in user can see the meeting.
    """
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

    return data["value"][0]["id"]  # Graph meetingId (not the join URL token)


def fetch_transcript_by_id(access_token: str, meeting_id: str, debug: bool) -> str:
    """
    Fetch transcript text for a given Teams meeting ID.
    - Try v1.0, then beta.
    - If the list endpoint returns a transcript item without inline `content`,
      download the content from `transcriptContentUrl` with explicit Accept headers.
    - Saves raw responses to:
        graph_transcripts_v1.json
        graph_transcripts_beta.json
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    tried = []

    def _download_content_stream(url: str, label: str) -> str:
        # Graph rejects Accept: */* for transcript content. Try explicit formats.
        accept_candidates = [
            "text/plain",
            "text/vtt",
            "application/vnd.ms-graph.vtt",
            "application/octet-stream",
        ]
        last_err = None
        for accept in accept_candidates:
            try:
                r = requests.get(
                    url,
                    headers={"Authorization": headers["Authorization"], "Accept": accept},
                    timeout=60,
                )
                if debug:
                    print(f"[DEBUG] GET {label} content (Accept: {accept}) -> {r.status_code} {r.headers.get('Content-Type')}")
                if r.status_code == 200 and r.content is not None:
                    # best-effort decode
                    try:
                        text = r.content.decode("utf-8", errors="replace").strip()
                    except Exception:
                        text = r.text.strip()

                    # Save a copy to disk for inspection
                    out_name = f"graph_transcript_content_{label.split('/')[-1].replace('.', '_')}.txt"
                    try:
                        with open(out_name, "w", encoding="utf-8") as f:
                            f.write(text)
                        if debug:
                            print(f"[DEBUG] Saved transcript content -> {out_name} ({len(text)} chars)")
                    except Exception as e:
                        print(f"[WARN] Could not save transcript content: {e}")

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
            # 1) Inline content
            inline = t0.get("content")
            if inline:
                return inline.strip()
            # 2) Content stream
            stream_url = t0.get("transcriptContentUrl")
            if stream_url:
                return _download_content_stream(stream_url, base)
            if debug:
                print(f"[DEBUG] {base}: transcript item present but no `content` or `transcriptContentUrl`.")

    details = "; ".join([f"{b}={code}" for b, code in tried])
    raise RuntimeError(
        f"No usable transcript returned. Statuses: {details}. "
        f"Check graph_transcripts_v1.json / graph_transcripts_beta.json."
    )


def summarize_text_with_openai(text: str) -> str:
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "OPENAI_API_KEY is not set. For PowerShell:\n"
            '  $env:OPENAI_API_KEY="sk-..."\n'
            "then re-run this script."
        )

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
    return resp.choices[0].message.content


def acquire_user_token_device_code(client_id: str, tenant_id: str) -> str:
    app = PublicClientApplication(client_id=client_id, authority=f"https://login.microsoftonline.com/{tenant_id}")
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes=GRAPH_SCOPE, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=GRAPH_SCOPE)
    if "user_code" not in flow:
        raise RuntimeError("Failed to initiate device code flow.")
    print(f"\nTo sign in, go to {flow['verification_uri']} and enter code: {flow['user_code']}\n")
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Could not acquire token: {result}")
    return result["access_token"]


def send_email(access_token: str, recipients: list[str], subject: str, body_html: str):
    url = f"{GRAPH_BASE}/me/sendMail"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
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
    print(f"✅ Email sent to {', '.join(recipients)}")


def main():
    parser = argparse.ArgumentParser(
        description="Fetch a Teams transcript by meeting link, summarize with AI, and email via Microsoft Graph.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""
            Example:
              python teams.py --meeting-link "PASTE_TEAMS_LINK_HERE" --to "alice@contoso.com,bob@contoso.com" --subject "Call Summary"
        """),
    )
    parser.add_argument("--meeting-link", required=True, help="Full Microsoft Teams meeting join link")
    parser.add_argument("--to", required=True, help="Comma-separated email addresses")
    parser.add_argument("--subject", default="Teams Call Summary", help="Email subject")
    parser.add_argument("--debug", action="store_true", help="Print and save raw Graph responses")
    args = parser.parse_args()

    client_id = os.getenv("AZURE_CLIENT_ID")
    tenant_id = os.getenv("AZURE_TENANT_ID")
    if not client_id or not tenant_id:
        raise RuntimeError("AZURE_CLIENT_ID and AZURE_TENANT_ID must be set for device-code flow.")

    print("Acquire Microsoft Graph token (device code flow)...")
    token = acquire_user_token_device_code(client_id, tenant_id)

    print("Resolving meeting ID from link...")
    meeting_id = resolve_meeting_id_from_link(token, args.meeting_link, args.debug)
    print(f"Graph returned meeting ID: {meeting_id}")

    print(f"Fetching transcript for meeting {meeting_id}...")
    transcript = fetch_transcript_by_id(token, meeting_id, args.debug)

    print("Summarizing transcript with OpenAI...")
    summary_html = summarize_text_with_openai(transcript)

    print("Sending summary email...")
    recipients = [e.strip() for e in args.to.split(",")]
    send_email(token, recipients, args.subject, summary_html)

    print("Done. Email submitted to Graph.")


if __name__ == "__main__":
    main()
