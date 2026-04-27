# Lecturer AI Agent

An autonomous AI-powered lecturer assistant that automates lecture preparation, lecturer approval workflows, and student content delivery.

## Project Overview

Lecturer AI Agent is an end-to-end academic automation system designed to reduce manual workload in teaching workflows.

The agent intelligently checks upcoming lectures from Google Calendar, analyzes syllabus content, generates lecture notes and PowerPoint slides using AI, sends them to the lecturer for review via email, handles revision requests automatically, and publishes approved materials to Canvas LMS while notifying students.

---

## What the Agent Does

-  Fetches and analyzes lecture schedules from Google Calendar
-  Prioritizes urgent upcoming lectures
-  Parses syllabus PDFs to identify current academic week topics
-  Generates structured lecture notes using Gemini AI
-  Creates 15-slide lecture presentations automatically
-  Sends slides to lecturers for approval via Gmail
-  Detects approval or requested revisions
-  Regenerates slides based on lecturer feedback
-  Uploads approved slides to Canvas LMS
-  Posts announcements and notifies students automatically

---

##  How It Interacts with the User

### Lecturer Workflow

1. Agent checks lecture schedule automatically
2. Generates lecture content and slides
3. Emails slides to lecturer for review
4. Lecturer replies:
   - **Approved** → Slides uploaded to Canvas automatically
   - **Changes requested** → Agent revises and resends
5. Students receive lecture material notification

---


## 🛠 Tech Stack

- Python
- Streamlit
- LangGraph
- LangChain
- Google Gemini API
- Google Calendar API
- Gmail API
- Canvas LMS API
- APScheduler

---

##  Workflow Architecture

```text
Calendar Check
      ↓
Lecture Priority Analysis
      ↓
Syllabus Parsing
      ↓
Lecture Notes Generation
      ↓
Slide Generation
      ↓
Lecturer Review Email
      ↓
Approval / Revision Loop
      ↓
Canvas LMS Upload
      ↓
Student Notification
