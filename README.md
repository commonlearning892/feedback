# Pre-Primary Feedback Analytics

This document describes the metrics and aggregates used to generate the Pre-Primary Feedback Dashboard and PowerPoint report for Sri Chaitanya Techno Schools – Tamil Nadu.

- Data source: `Parent Feedback Form – Academic & Administrative Review - Pre Primary.csv`
- Generated outputs:
  - `feedback_stats.json` (all computed statistics)
  - `Pre_Primary_Feedback_Analysis.pptx` (styled, clickable PPT)
- Logo used: `srichaitanya.jpg`

## How to run
- Generate analytics (JSON + PPT):
  - Run: `python3 analyze_feedback.py`
- View the dashboard locally (optional):
  - Run a simple HTTP server in this folder and open `http://localhost:8000/dashboard.html`

## Summary KPIs
- Total responses: 5,063
- Classes
  - IK - 2: 2,171
  - IK - 1: 1,982
  - Pre - KG: 822
  - Unknown: 88
- Orientations
  - Techno: 3,887
  - Star Mavericks: 1,089
  - Unknown: 87
- Languages
  - Tamil: 4,180
  - Hindi: 663
  - Unknown: 220
- Overall average rating (0–5): 4.56

## Category Scores (0–5)
- Academics: 4.54
- Administration: 4.56
- Environment: 4.66
- Infrastructure: 4.38
- App: 4.57
- Transport: 4.13

## Recommendation
- Distribution
  - Yes: 4,190
  - No: 107
  - Maybe: 619
  - Not Applicable: 0
- Yes %: 85.23%

## PTM Effectiveness (0–5)
- Average: 4.65

## Teaching Indicators (0–5)
- Concept Clarity: 4.64
- Teacher Approachability: 4.69
- Engagement: 4.64
- Communication Skills: 4.51

## Environment Focus (0–5)
- Interest in attending school: 4.65
- Campus safety: 4.70
- Moral values: 4.57

## Communication & Administration (0–5)
- Leadership Access: 4.62
- Front Office Support: 4.61
- App Usability: 4.34
- Timely Updates: 4.71

## Concern Handling & Resolution
- Role-wise concern handling: Not available (no role-wise columns detected in source)
- Concern Resolution distribution
  - Yes: 3,049
  - No: 1,815
  - Maybe: 0
  - Not Applicable: 0

## Subject-wise Performance (0–5)
- Literacy (English)
  - Average: 4.59
  - Distribution: Excellent 3,113 | Good 1,432 | Average 248 | Poor 16
- Numeracy (Math)
  - Average: 4.60
  - Distribution: Excellent 3,029 | Good 1,368 | Average 214 | Poor 10
- General Awareness
  - Average: 4.55
  - Distribution: Excellent 2,834 | Good 1,462 | Average 263 | Poor 19
- Second Language
  - Average: 4.26
  - Distribution: Excellent 1,770 | Good 1,434 | Average 526 | Poor 86

## Additional Breakdowns (used in dashboard/PPT)
- Branch performance: per-branch averages across Academics, Environment, Infrastructure, Parent-Teacher, Admin, and Overall.
- Orientation performance: per-orientation averages and counts.
- Class performance: per-class averages and counts.
- Branch recommendation %: per-branch “Recommend Yes %”.

These breakdowns are charted in the dashboard and PPT; refer to those views for ranked lists and visuals (Top Branches, Branch Recommendation %, and scatter charts).

## PowerPoint Report Features
- Clickable navigation:
  - A Dashboard Menu slide with section buttons.
  - Persistent tab bar on all slides; active section highlighted.
  - "Menu" button on each slide to return to the main menu.
- Native styling to mirror the dashboard:
  - KPI stat cards.
  - Card-style containers for charts and content.
  - Consistent colors, borders, and fonts.

## File Locations
- CSV input: `/Users/venkubabugollapudi/Desktop/Feedback/Feed Back/Parent Feedback Form – Academic & Administrative Review - Pre Primary.csv`
- JSON output: `/Users/venkubabugollapudi/Desktop/Feedback/Feed Back/feedback_stats.json`
- PPT output: `/Users/venkubabugollapudi/Desktop/Feedback/Feed Back/Pre_Primary_Feedback_Analysis.pptx`
- Dashboard: `dashboard.html` (open via local server)

## Notes
- Ratings are normalized on a 0–5 scale.
- Null/"Not Applicable" responses are excluded from averages.
- Category labels are derived from the CSV; some labels are abbreviated for layout.
