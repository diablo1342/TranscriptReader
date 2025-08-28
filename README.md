<h1> TranscriptReader </h1>

## 📝 Overview 📝

The Transcript Reader is a cloud-integrated application that retrieves and processes academic transcript data from Microsoft services. It uses Azure AD Authentication and the Microsoft Graph API to securely access user data, extract relevant information, and present it in a structured format.

## How I built it
Authentication: Azure Active Directory (Azure AD) for secure user login and API access.

API Integration: Microsoft Graph API to access user transcripts and related data.

Backend: Python / FastAPI (or Flask) to handle API calls, token exchange, and data parsing.

## Key Features
Secure OAuth2 authentication via Azure AD.

Fetches transcript data via Microsoft Graph API endpoints.

Parses academic records into readable and structured formats.

Can be extended to support GPA calculations, course tracking, and visual dashboards.

## Built with
Python

FastAPI / Flask

Azure Active Directory

Microsoft Graph API

