#!/usr/bin/env python3


import argparse
import os
import sys
import textwrap
import json
from datetime import datetime

import requests
from msal import PublicClientApplication

# --- OpenAI (simple SDK usage) ---
try:
    from openai import OpenAI
except ImportError:
    print("Please `pip install openai`.", file=sys.stderr)
    raise

GRAPH_SCOPE = ["Mail.Send", "User.Read"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def read_transcript(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read().strip()


def summarize_with_openai(transcript: str, model: str, base_url: str | None = None) -> str:

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY not set.")

    client_kwargs = {"api_key": api_key}
    if base_url:
        client_kwargs["base_url"] = base_url
    client = OpenAI(**client_kwargs)

    system_msg = (
        "You are an expert meeting summarizer. Produce a crisp, factual, "
        "non-redundant summary suitable to email to stakeholders who did not attend. "
        "Prefer bullets. Infer assignees/deadlines only if clear."
    )
    user_msg = f"""
Below is a raw transcript from a Microsoft Teams call. Please produce:

1) Executive Summary (3–6 bullets)
2) Key Decisions (with rationale if stated)
3) Open Questions & Risks
4) Action Items (Assignee • Task • Due date if mentioned)
5) Notable Quotes (optional, brief)

Transcript:
---
{transcript[:120000]}  # safeguard size, truncate if extremely large
---
"""
    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_msg},
            {"role": "user", "content": user_msg},
        ],
        temperature=0.2,
    )
    return resp.choices[0].message.content.strip()


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


def build_html_email(summary_md: str, transcript_snippet: str) -> str:

    def esc(s: str) -> str:
        return (
            s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
        )

    lines = summary_md.splitlines()
    html_lines = ['<div style="font-family:Segoe UI,Arial,sans-serif; line-height:1.5">']
    for ln in lines:
        if ln.strip().startswith("###") or ln.strip().startswith("##"):
            html_lines.append(f"<h3>{esc(ln.strip('# ').strip())}</h3>")
        elif ln.strip().startswith("- ") or ln.strip().startswith("* "):

            if not (html_lines and html_lines[-1].startswith("<ul")):
                html_lines.append("<ul>")
            html_lines.append(f"<li>{esc(ln[2:].strip())}</li>")

        else:

            if html_lines and html_lines[-1].startswith("<li>"):
                html_lines.append("</ul>")
            if ln.strip():
                html_lines.append(f"<p>{esc(ln.strip())}</p>")
            else:
                html_lines.append("<br/>")

    if html_lines and html_lines[-1].startswith("<li>"):
        html_lines.append("</ul>")

    
    html_lines.append("<hr/>")
    html_lines.append("<p><strong>Transcript snippet (first ~1,000 chars):</strong></p>")
    html_lines.append(f"<pre style='white-space:pre-wrap'>{esc(transcript_snippet[:1000])}</pre>")
    html_lines.append(f"<p style='color:#888'>Sent {esc(datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC'))}</p>")
    html_lines.append("</div>")
    return "\n".join(html_lines)


def send_mail_graph(access_token: str, to_addresses: list[str], subject: str, html_body: str) -> None:
    url = f"{GRAPH_BASE}/me/sendMail"
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": addr.strip()}} for addr in to_addresses if addr.strip()],
        },
        "saveToSentItems": True,
    }
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    r = requests.post(url, headers=headers, data=json.dumps(payload), timeout=30)
    if r.status_code not in (202, 200):
        raise RuntimeError(f"Graph sendMail failed: {r.status_code} {r.text}")


def main():
    parser = argparse.ArgumentParser(
        description="Summarize a Teams transcript with AI and email via Microsoft Graph.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""
            Examples:
              python summarize_and_email.py --transcript call.txt --to "alice@contoso.com,bob@contoso.com" --subject "Call Summary"
        """),
    )
    parser.add_argument("--transcript", required=True, help="Path to the Teams transcript .txt")
    parser.add_argument("--to", required=True, help="Comma-separated email addresses")
    parser.add_argument("--subject", default="Teams Call Summary", help="Email subject")
    parser.add_argument("--model", default=os.getenv("OPENAI_MODEL", "gpt-4o-mini"), help="OpenAI model name")
    parser.add_argument("--openai-base-url", default=os.getenv("OPENAI_BASE_URL"), help="Override base URL (e.g., Azure OpenAI endpoint)")
    parser.add_argument("--from-user", action="store_true", help="Send as the signed-in user (device-code flow)")

    args = parser.parse_args()

    transcript = read_transcript(args.transcript)
    print("Summarizing transcript with AI...")
    summary = summarize_with_openai(transcript, args.model, args.openai_base_url)

    html = build_html_email(summary, transcript)

    if not args.from_user:
        print("This script currently supports sending as the signed-in user. Use --from-user.")
        sys.exit(2)

    client_id = os.getenv("AZURE_CLIENT_ID")
    tenant_id = os.getenv("AZURE_TENANT_ID")
    if not client_id or not tenant_id:
        raise RuntimeError("AZURE_CLIENT_ID and AZURE_TENANT_ID must be set for device-code flow.")

    print("Acquire Microsoft Graph token (device code flow)...")
    token = acquire_user_token_device_code(client_id, tenant_id)

    print("Sending email via Microsoft Graph...")
    to_addresses = [e.strip() for e in args.to.split(",")]
    send_mail_graph(token, to_addresses, args.subject, html)

    print("Done. Email submitted to Graph.")


if __name__ == "__main__":
    main()
