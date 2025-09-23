## Teams Meeting Transcript Summarizer

A tool that takes a Microsoft Teams meeting link, fetches the transcript via Microsoft Graph API, generates a structured AI-powered summary using OpenAI, and automatically emails the recap to participants.


## Features

-Authentication with Microsoft Graph (device code flow)

-Resolve meeting IDs directly from Teams meeting links

-Summarize transcripts into clean HTML using OpenAI GPT models

-Send AI-generated summaries via email to chosen recipients

-Streamlit-based web interface for easy use by non-technical users

## Tech Stack

-Python

-Streamlit

-MSAL (Microsoft Authentication Library) – device code authentication

-Microsoft Graph API – meeting & transcript data, email sending

-OpenAI API – transcript summarization with GPT

-Requests / JSON – API requests and data parsing

## Setup
  1. Clone Repository
  
  2. Install Dependencies (pip install -r requirements.txt)
  
  3. Azure App Registration (One-Time Setup)
  
    -To use the Microsoft Graph API, you need to register an app in Azure Active Directory:
    
    -Go to the Azure Portal, go to app registrations and click new registration.
    
    -Give your app a name.
    
    -Copy your Application (client) ID and Directory (tenant) ID — you’ll need these later.
    
    -Under API Permissions, add:
    
      -OnlineMeetings.Read
      
      -Mail.Send
      
      -User.Read
  
    -Click Grant admin consent so the app can access these permissions (If admin has already not given permission)

  4. Environment Variables
  
    -Set the following environment variables:
      1. OPENAI_API_KEY
      2. AZURE_CLIENT_ID
      3. AZURE_TENANT_ID
  
  5. Run Streamlit App in terminal
    streamlit run app.py

## 🚧 Limitations

-This app integrates with the Microsoft Graph API, which requires authentication through an Azure AD tenant.

-Currently, only users within the Ignitiv Azure tenant (or another properly configured tenant) can log in and use the Teams integration features.

-If you are outside the organization, you won’t be able to authenticate with Microsoft Graph.

-Here is a link to the app for people within Ignitiv: https://transcriptreader-m2ajqeey8jsqdmshttnna6.streamlit.app/

## Demo Video
https://www.loom.com/share/4c21d2988d1c48959ec4635ebde526e0?sid=da241a99-3908-436d-a05e-e47cfac71c5e


## Future Improvements

  - Add support for Zoom/Google Meet transcripts

  - Enhance summary customization (tone, length, bullet vs paragraph)
