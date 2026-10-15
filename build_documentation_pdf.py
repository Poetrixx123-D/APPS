"""
build_documentation_pdf.py
Generates a comprehensive, publication-quality technical PDF documentation
for the APPS (AI Predictive Personnel Stress - Welfare and Monitoring System) system.
"""

import os
import sys
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.pdfgen import canvas
from xml.sax.saxutils import escape as xml_escape

class NumberedCanvas(canvas.Canvas):
    """
    Two-pass canvas to dynamically compute total pages and render running
    headers, footers, rules, and page numbers.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()
        print(f"Total pages rendered: {num_pages}")

    def draw_page_decorations(self, page_count):
        self.saveState()
        page_w, page_h = A4

        # Skip headers and footers on Cover Page (Page 1)
        if self._pageNumber > 1:
            # Header
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.HexColor("#475569"))
            self.drawString(36, page_h - 26, "APPS - AI Predictive Personnel Stress | Technical Architecture")
            self.drawRightString(page_w - 36, page_h - 26, "CONFIDENTIAL & PROPRIETARY")

            # Header rule
            self.setStrokeColor(colors.HexColor("#CBD5E1"))
            self.setLineWidth(0.5)
            self.line(36, page_h - 30, page_w - 36, page_h - 30)

            # Footer rule
            self.line(36, 36, page_w - 36, 36)

            # Footer text
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.HexColor("#64748B"))
            self.drawString(36, 24, "System Specification & Implementation Guide | Version 2.0")
            page_str = f"Page {self._pageNumber} of {page_count}"
            self.drawRightString(page_w - 36, 24, page_str)

        self.restoreState()


def create_docs_pdf(output_path="APPS_System_Documentation.pdf"):
    # Margins: 36pt (0.5 inch) left/right, 42pt top/bottom
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=42,
        bottomMargin=42
    )

    page_width, page_height = A4
    printable_width = page_width - 72  # 523.27 pt

    styles = getSampleStyleSheet()

    # Custom typography styles
    c_primary = colors.HexColor("#0F172A")    # Deep Navy
    c_accent = colors.HexColor("#2563EB")     # Royal Blue
    c_teal = colors.HexColor("#0D9488")       # Deep Teal
    c_body = colors.HexColor("#1E293B")       # Dark Slate
    c_muted = colors.HexColor("#64748B")      # Slate Gray
    c_border = colors.HexColor("#E2E8F0")     # Light Border
    c_bg_light = colors.HexColor("#F8FAFC")   # Light Background
    c_success = colors.HexColor("#059669")    # Green
    c_amber = colors.HexColor("#D97706")      # Amber

    title_style = ParagraphStyle(
        'CoverTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=24,
        leading=28,
        textColor=c_primary,
        alignment=0
    )

    subtitle_style = ParagraphStyle(
        'CoverSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=12,
        leading=16,
        textColor=c_accent,
        alignment=0
    )

    h1_style = ParagraphStyle(
        'SectionHeading',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=14,
        leading=18,
        textColor=c_primary,
        spaceBefore=14,
        spaceAfter=6,
        keepWithNext=True
    )

    h2_style = ParagraphStyle(
        'SubSectionHeading',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10.5,
        leading=14,
        textColor=c_accent,
        spaceBefore=8,
        spaceAfter=4,
        keepWithNext=True
    )

    body_style = ParagraphStyle(
        'BodyDark',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=12,
        textColor=c_body,
        spaceBefore=2,
        spaceAfter=4
    )

    body_bold = ParagraphStyle(
        'BodyBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=12,
        textColor=c_body
    )

    code_style = ParagraphStyle(
        'CodeSnippet',
        parent=styles['Normal'],
        fontName='Courier',
        fontSize=7.5,
        leading=10,
        textColor=colors.HexColor("#0F172A")
    )

    table_cell = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10.5,
        textColor=c_body
    )

    table_header = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=11,
        textColor=colors.white
    )

    badge_pill = ParagraphStyle(
        'BadgePill',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7,
        leading=9,
        textColor=c_accent
    )

    callout_text = ParagraphStyle(
        'CalloutText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.2,
        leading=11.5,
        textColor=colors.HexColor("#1E3A8A")
    )

    def p(text, style=body_style):
        return Paragraph(text, style)

    def callout(title_txt, body_txt, bg_color="#EFF6FF", border_color="#3B82F6"):
        content = [
            Paragraph(f"<b>{xml_escape(title_txt)}</b>", ParagraphStyle('CT', parent=callout_text, fontName='Helvetica-Bold', textColor=colors.HexColor(border_color))),
            Spacer(1, 3),
            Paragraph(xml_escape(body_txt), callout_text)
        ]
        t = Table([[content]], colWidths=[printable_width])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor(bg_color)),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor(border_color)),
            ('PADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ]))
        return t

    story = []

    # =========================================================================
    # COVER / HEADER BANNER
    # =========================================================================
    header_table_data = [
        [
            Paragraph("<b>APPS</b> &bull; AI PREDICTIVE PERSONNEL STRESS", ParagraphStyle('BannerTop', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, textColor=c_accent)),
            Paragraph("WELFARE & MONITORING SYSTEM", ParagraphStyle('BannerRight', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=c_muted, alignment=2))
        ]
    ]
    t_top = Table(header_table_data, colWidths=[320, printable_width - 320])
    t_top.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(t_top)
    story.append(Spacer(1, 8))

    story.append(HRFlowable(width="100%", thickness=2, color=c_accent, spaceBefore=0, spaceAfter=14))

    story.append(Paragraph("Complete Technical Documentation", title_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph("Frontend, Backend Architecture, Database Schema, Machine Learning & Recent Enhancements", subtitle_style))
    story.append(Spacer(1, 14))

    meta_data = [
        [
            Paragraph("<b>Project:</b> AI Predictive Personnel Stress (APPS)", table_cell),
            Paragraph("<b>Environment:</b> Vercel Serverless / Cloud Blob", table_cell),
        ],
        [
            Paragraph("<b>Backend:</b> Python Flask 3.x / Scikit-Learn", table_cell),
            Paragraph("<b>Database:</b> SQLite3 + Two-Way Cloud Blob Sync", table_cell),
        ],
        [
            Paragraph("<b>Frontend:</b> Responsive HTML5 / CSS3 / Jinja2", table_cell),
            Paragraph("<b>Status:</b> Production Live Verified", table_cell),
        ],
        [
            Paragraph("<b>Live URL:</b> https://apps-nu-henna.vercel.app", table_cell),
            Paragraph("<b>Generated:</b> October 2026 | Version 2.2 (Continuous Sync)", table_cell),
        ]
    ]
    t_meta = Table(meta_data, colWidths=[printable_width * 0.5, printable_width * 0.5])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F1F5F9")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#CBD5E1")),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 14))

    # =========================================================================
    # SECTION 1: EXECUTIVE SUMMARY
    # =========================================================================
    story.append(Paragraph("1. Executive Summary & System Overview", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_primary, spaceBefore=2, spaceAfter=8))

    exec_summary_text = (
        "<b>APPS (AI Predictive Personnel Stress - Welfare and Monitoring System)</b> is a full-stack, mission-critical digital psychological "
        "welfare platform designed to monitor, assess, and support the mental wellbeing of military personnel. "
        "The system incorporates an automated <b>Machine Learning Random Forest classifier</b> to analyze daily physiological "
        "and operational metrics (sleep duration, active duty hours, resting heart rate, and mood score), flagging soldiers experiencing "
        "heightened stress levels and delivering personalized coping recommendations.<br/><br/>"
        "The platform establishes a strict, role-governed operational hierarchy among three distinct user categories: "
        "<b>(1) Personnel (Soldiers)</b> who submit daily check-ins and request assistance, "
        "<b>(2) Welfare Officers</b> who monitor unit-wide personnel rosters, author privileged counseling assessments, track physiological risk trends, and resolve welfare requests collaboratively without caseload locking barriers, and "
        "<b>(3) Senior Administrators</b> who monitor force-wide wellness trends, unit-level risk distributions, generate access credentials, "
        "and export audit-ready reports."
    )
    story.append(Paragraph(exec_summary_text, body_style))
    story.append(Spacer(1, 8))

    # Key pillars table
    pillars_data = [
        [
            Paragraph("<b>Core Module</b>", table_header),
            Paragraph("<b>Primary Responsibility</b>", table_header),
            Paragraph("<b>Technologies & Protocols</b>", table_header)
        ],
        [
            Paragraph("<b>Frontend Client</b>", table_cell),
            Paragraph("Adaptive UI for smartphones, tablets, and laptops. Role-based dashboards, survey forms, live password strength indicators.", table_cell),
            Paragraph("HTML5, Custom Glassmorphic CSS3, Vanilla JS, Jinja2 Templates", table_cell)
        ],
        [
            Paragraph("<b>Backend Engine</b>", table_cell),
            Paragraph("REST routing, role decorators (RBAC), CSRF & rate limit protections, automated officer load-balancing, ML inference.", table_cell),
            Paragraph("Python, Flask, Werkzeug Security, Scikit-learn, Joblib", table_cell)
        ],
        [
            Paragraph("<b>Predictive AI</b>", table_cell),
            Paragraph("Predicts stress levels (Low, Moderate, High) from physiological telemetry; generates automated non-clinical wellness advice.", table_cell),
            Paragraph("Scikit-learn Random Forest Classifier, Pandas, NumPy", table_cell)
        ],
        [
            Paragraph("<b>Database & Cloud Sync</b>", table_cell),
            Paragraph("Relational persistence with real-time bidirectional synchronization between local SQLite and Vercel Blob Object Storage.", table_cell),
            Paragraph("SQLite3, Vercel Blob Private REST API, SyncedConnection", table_cell)
        ],
        [
            Paragraph("<b>Audit & Reporting</b>", table_cell),
            Paragraph("Generates unit stress analytics, Excel spreadsheets with injection defense, sanitized PDFs, and CSV telemetry exports.", table_cell),
            Paragraph("ReportLab, OpenPyXL, CSV, ActivityLog audit trail", table_cell)
        ],
    ]
    t_pillars = Table(pillars_data, colWidths=[100, 240, printable_width - 340])
    t_pillars.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.5, c_border),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, c_bg_light]),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_pillars)
    story.append(Spacer(1, 14))

    # =========================================================================
    # SECTION 2: FRONTEND ARCHITECTURE & UX
    # =========================================================================
    story.append(Paragraph("2. Frontend Architecture & User Interface Design", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_primary, spaceBefore=2, spaceAfter=8))

    fe_overview = (
        "The APPS frontend is architected as a lightweight, lightning-fast, zero-dependency server-side rendered application. "
        "It avoids bloated JavaScript frameworks in favor of clean semantic HTML5, Jinja2 templates, and a custom CSS design system "
        "built with modern CSS variables, responsive flexbox/grid layouts, and glassmorphic aesthetic cards."
    )
    story.append(Paragraph(fe_overview, body_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("A. Mobile-First & Desktop Ergonomics (Responsive Design)", h2_style))
    fe_resp = (
        "A major overhaul was deployed to guarantee optimal usability across all form factors: mobile smartphones, tablets, and desktop workstations.<br/>"
        "&bull; <b>Desktop View (&gt;920px):</b> Expansive horizontal navigation bar, multi-column metric cards, full tabular summaries, and prominent quick-action toolbars.<br/>"
        "&bull; <b>Mobile View (&le;920px &amp; &le;480px):</b> Automated conversion to an off-canvas slide-out hamburger navigation drawer, stacked single-column cards, "
        "horizontal swipe containers for data tables, and enlarged 48px touch targets for thumb ergonomics.<br/>"
        "&bull; <b>Fluid Layout Engine:</b> Responsive typography using viewport clamps, flexible container margins (<code>14px</code> on mobile up to <code>1200px</code> centered on desktop), "
        "and mobile-optimized login shells preventing iOS/Android keyboard zoom."
    )
    story.append(Paragraph(fe_resp, body_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("B. Complete Template Catalog & View Hierarchy", h2_style))

    tmpl_data = [
        [Paragraph("<b>Template File</b>", table_header), Paragraph("<b>Route / URL</b>", table_header), Paragraph("<b>UI Purpose & Key Components</b>", table_header)],
        [
            Paragraph("<code>portal.html</code>", code_style),
            Paragraph("<code>/portal</code>, <code>/</code>", code_style),
            Paragraph("Unified entry landing page. Features role selection cards (Soldier, Officer, Administrator) with explicit signup links.", table_cell)
        ],
        [
            Paragraph("<code>login.html</code>", code_style),
            Paragraph("<code>/login</code>, <code>/*_login</code>", code_style),
            Paragraph("Single Unified Login interface with role selector tabs, remember-me support, password visibility toggle, and error toasts.", table_cell)
        ],
        [
            Paragraph("<code>signup.html</code>", code_style),
            Paragraph("<code>/signup</code>", code_style),
            Paragraph("Personnel registration form. Unit selection, live password strength requirements meter, and officer auto-assignment notice.", table_cell)
        ],
        [
            Paragraph("<code>officer_signup.html</code>", code_style),
            Paragraph("<code>/officer_signup</code>", code_style),
            Paragraph("Officer registration. Requires cryptographically verified one-time officer access code (<code>OFC-XXXXXX</code>).", table_cell)
        ],
        [
            Paragraph("<code>admin_signup.html</code>", code_style),
            Paragraph("<code>/admin_signup</code>", code_style),
            Paragraph("Senior Administrator registration. Requires one-time administrator access code (<code>ADM-XXXXXX</code>).", table_cell)
        ],
        [
            Paragraph("<code>upgrade_password.html</code>", code_style),
            Paragraph("<code>/upgrade_password</code>", code_style),
            Paragraph("Mandatory password upgrade screen for users with legacy/weak passwords. Enforces minimum length and complexity.", table_cell)
        ],
        [
            Paragraph("<code>forgot_password.html</code>", code_style),
            Paragraph("<code>/forgot_password</code>", code_style),
            Paragraph("Self-service credential recovery portal. Role switcher (Personnel/Officer/Admin), Registered Unit verification, and dynamic strength meter.", table_cell)
        ],
        [
            Paragraph("<code>dashboard.html</code>", code_style),
            Paragraph("<code>/dashboard</code>", code_style),
            Paragraph("Personnel daily portal. Daily wellness survey input, 7-day health trend chart, recent check-in history, quick stress relief button.", table_cell)
        ],
        [
            Paragraph("<code>results.html</code>", code_style),
            Paragraph("<code>/results</code>", code_style),
            Paragraph("Immediate post-checkin feedback. Shows AI-predicted stress level badge (Low/Moderate/High) and tailored coping advice.", table_cell)
        ],
        [
            Paragraph("<code>snapshot.html</code>", code_style),
            Paragraph("<code>/snapshot</code>", code_style),
            Paragraph("Personnel historical health telemetry view with PDF export trigger.", table_cell)
        ],
        [
            Paragraph("<code>officer_dashboard.html</code>", code_style),
            Paragraph("<code>/officer_dashboard</code>", code_style),
            Paragraph("Welfare Officer primary station. Summary counters, urgent triage alert queue, going quiet counter, workload indicators.", table_cell)
        ],
        [
            Paragraph("<code>officer_personnel.html</code>", code_style),
            Paragraph("<code>/officer/personnel</code>", code_style),
            Paragraph("Comprehensive roster view of all unit personnel. Displays unit badge, last check-in status, physiological metrics, and direct dossier profile link.", table_cell)
        ],
        [
            Paragraph("<code>personnel_detail.html</code>", code_style),
            Paragraph("<code>/officer/personnel/&lt;id&gt;</code>", code_style),
            Paragraph("Deep-dive soldier dossier. Physiological time-series charts, check-in history, counseling note authoring card, and welfare request logs.", table_cell)
        ],
        [
            Paragraph("<code>officer_going_quiet.html</code>", code_style),
            Paragraph("<code>/officer/going_quiet</code>", code_style),
            Paragraph("Early-warning monitoring page listing personnel who have not submitted a check-in for 3 or more consecutive days.", table_cell)
        ],
        [
            Paragraph("<code>officer_requests.html</code>", code_style),
            Paragraph("<code>/officer/requests</code>", code_style),
            Paragraph("5-stage stress relief request kanban/pipeline. Officers can advance status from New &rarr; Acknowledged &rarr; In Progress &rarr; Resolved.", table_cell)
        ],
        [
            Paragraph("<code>officer_activity_log.html</code>", code_style),
            Paragraph("<code>/officer/activity_log</code>", code_style),
            Paragraph("Immutable audit trail displaying officer actions, counseling notes, case reviews, and request pipeline status transitions.", table_cell)
        ],
        [
            Paragraph("<code>admin_dashboard.html</code>", code_style),
            Paragraph("<code>/admin_dashboard</code>", code_style),
            Paragraph("Senior Command Command Center. Unit stress distribution breakdown, high-risk flags, one-time code generator, batch export buttons.", table_cell)
        ],
        [
            Paragraph("<code>admin_units.html</code>", code_style),
            Paragraph("<code>/admin/units</code>", code_style),
            Paragraph("Comparative unit welfare index table comparing risk percentages across regiments, battalions, and commands.", table_cell)
        ],
    ]
    t_tmpl = Table(tmpl_data, colWidths=[120, 110, printable_width - 230])
    t_tmpl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.5, c_border),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, c_bg_light]),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_tmpl)
    story.append(Spacer(1, 14))

    # =========================================================================
    # SECTION 3: BACKEND ARCHITECTURE & APIS
    # =========================================================================
    story.append(PageBreak())
    story.append(Paragraph("3. Backend Architecture & API Specifications", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_primary, spaceBefore=2, spaceAfter=8))

    be_text = (
        "The backend is powered by <b>Python 3 and Flask</b>, engineered for serverless deployment on Vercel Functions. "
        "It employs modular architectural patterns including custom request decorators, strict security policies, "
        "cryptographic password hashing, rate limiting, and an automated officer assignment load balancer."
    )
    story.append(Paragraph(be_text, body_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("A. Key Architectural Subsystems", h2_style))
    subsystems_text = (
        "&bull; <b>Role-Based Access Control (RBAC):</b> Implemented via <code>@login_required</code> and <code>@role_required('admin'|'officer'|'personnel')</code> decorators. "
        "Unauthorized requests are redirected to role-specific login routes with contextual flash error notices.<br/>"
        "&bull; <b>Automated Officer Load Balancer (<code>pick_officer</code>):</b> When a new soldier signs up, the system executes an atomic subquery calculating the active caseload "
        "of all registered officers and assigns the soldier to the officer with the lowest workload, ensuring fair distribution across medical staff.<br/>"
        "&bull; <b>Universal Officer Request Pipeline:</b> Cross-unit routing guarantees that all stress relief requests dispatched by any soldier are surfaced on the Welfare Officer console, preventing dropped communications.<br/>"
        "&bull; <b>Timezone Localization & Device Synchronization:</b> Backend operations strictly adhere to Indian Standard Time (<code>UTC+05:30</code>) "
        "via <code>now_local()</code> and <code>today_local()</code>, paired with client-side dynamic device clock synchronization for accurate temporal greetings.<br/>"
        "&bull; <b>CSRF & Same-Origin Enforcement:</b> All state-mutating POST requests undergo header origin checks verifying the request origin matches the host domain.<br/>"
        "&bull; <b>Rate Limiting & Brute-Force Backoff:</b> Failed login attempts are throttled to prevent credential stuffing attacks.<br/>"
        "&bull; <b>Formula & CSV Injection Defense:</b> When generating Excel and CSV exports, strings prefixed with arithmetic characters (<code>=, +, -, @</code>) "
        "are sanitized to prevent remote code execution in spreadsheet software."
    )
    story.append(Paragraph(subsystems_text, body_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("B. Complete Route & API Reference", h2_style))

    routes_data = [
        [Paragraph("<b>Route Path</b>", table_header), Paragraph("<b>HTTP</b>", table_header), Paragraph("<b>Role</b>", table_header), Paragraph("<b>Functionality & Behavior</b>", table_header)],
        [Paragraph("<code>/login</code>", code_style), Paragraph("GET/POST", table_cell), Paragraph("Public", badge_pill), Paragraph("Authenticates users, checks password strength; redirects to upgrade or dashboard.", table_cell)],
        [Paragraph("<code>/forgot_password</code>", code_style), Paragraph("GET/POST", table_cell), Paragraph("Public", badge_pill), Paragraph("Identity recovery via Name + Unit + Role match; sets strong password & syncs to cloud.", table_cell)],
        [Paragraph("<code>/signup</code>", code_style), Paragraph("GET/POST", table_cell), Paragraph("Public", badge_pill), Paragraph("Creates soldier account; enforces strong password; calls pick_officer().", table_cell)],
        [Paragraph("<code>/officer_signup</code>", code_style), Paragraph("GET/POST", table_cell), Paragraph("Public", badge_pill), Paragraph("Validates and consumes single-use OFC-* code; creates officer account.", table_cell)],
        [Paragraph("<code>/admin_signup</code>", code_style), Paragraph("GET/POST", table_cell), Paragraph("Public", badge_pill), Paragraph("Validates and consumes single-use ADM-* code; creates admin account.", table_cell)],
        [Paragraph("<code>/upgrade_password</code>", code_style), Paragraph("GET/POST", table_cell), Paragraph("Auth", badge_pill), Paragraph("Upgrades legacy/weak password to PBKDF2 hash and syncs to cloud storage.", table_cell)],
        [Paragraph("<code>/submit_survey</code>", code_style), Paragraph("POST", table_cell), Paragraph("Soldier", badge_pill), Paragraph("Validates inputs; invokes Random Forest ML model; writes DailyData record.", table_cell)],
        [Paragraph("<code>/submit_stress_request</code>", code_style), Paragraph("POST", table_cell), Paragraph("Soldier", badge_pill), Paragraph("Dispatches confidential stress relief request to assigned officer.", table_cell)],
        [Paragraph("<code>/officer/requests/&lt;id&gt;/advance</code>", code_style), Paragraph("POST", table_cell), Paragraph("Officer", badge_pill), Paragraph("Transitions request through stage pipeline (New &rarr; Acknowledged &rarr; Resolved).", table_cell)],
        [Paragraph("<code>/add_note</code>", code_style), Paragraph("POST", table_cell), Paragraph("Officer", badge_pill), Paragraph("Saves confidential clinical counseling note and records audit ActivityLog.", table_cell)],
        [Paragraph("<code>/admin/generate_code</code>", code_style), Paragraph("POST", table_cell), Paragraph("Admin", badge_pill), Paragraph("Generates cryptographically random one-time registration codes for staff.", table_cell)],
        [Paragraph("<code>/admin/export_excel</code>", code_style), Paragraph("GET", table_cell), Paragraph("Admin", badge_pill), Paragraph("Generates styled OpenPyXL workbook containing force-wide health data.", table_cell)],
        [Paragraph("<code>/officer/export_report</code>", code_style), Paragraph("GET", table_cell), Paragraph("Officer", badge_pill), Paragraph("Compiles ReportLab PDF wellness dossier for assigned personnel.", table_cell)],
    ]
    t_routes = Table(routes_data, colWidths=[125, 45, 45, printable_width - 215])
    t_routes.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.5, c_border),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, c_bg_light]),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_routes)
    story.append(Spacer(1, 14))

    # =========================================================================
    # SECTION 4: MACHINE LEARNING & PREDICTION SUBSYSTEM
    # =========================================================================
    story.append(PageBreak())
    story.append(Paragraph("4. Machine Learning & Stress Prediction Subsystem", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_primary, spaceBefore=2, spaceAfter=8))

    ml_text = (
        "The core diagnostic engine of APPS utilizes an offline-trained <b>Random Forest Classifier</b> serialized with Joblib (<code>stress_model.pkl</code>). "
        "During every daily survey submission, the model evaluates multidimensional physiological and operational indicators in real-time "
        "to deliver an immediate risk classification without human latency."
    )
    story.append(Paragraph(ml_text, body_style))
    story.append(Spacer(1, 6))

    ml_features_data = [
        [Paragraph("<b>Feature Column</b>", table_header), Paragraph("<b>Data Type</b>", table_header), Paragraph("<b>Valid Range</b>", table_header), Paragraph("<b>Clinical Correlation & Significance</b>", table_header)],
        [
            Paragraph("<code>sleep_hours</code>", code_style),
            Paragraph("Float (REAL)", table_cell),
            Paragraph("0.0 &ndash; 24.0 hrs", table_cell),
            Paragraph("Restorative rest. Sleep deprivation (&lt;5 hrs) strongly correlates with cognitive fatigue and acute stress.", table_cell)
        ],
        [
            Paragraph("<code>duty_hours</code>", code_style),
            Paragraph("Float (REAL)", table_cell),
            Paragraph("0.0 &ndash; 24.0 hrs", table_cell),
            Paragraph("Operational workload. Extended duty shifts (&gt;14 hrs) increase burn-out potential and cardiovascular strain.", table_cell)
        ],
        [
            Paragraph("<code>heart_rate</code>", code_style),
            Paragraph("Integer", table_cell),
            Paragraph("40 &ndash; 200 bpm", table_cell),
            Paragraph("Resting / baseline pulse. Elevated resting heart rate serves as an autonomic indicator of sympathetic nervous activation.", table_cell)
        ],
        [
            Paragraph("<code>mood_score</code>", code_style),
            Paragraph("Integer", table_cell),
            Paragraph("1 (Low) &ndash; 10 (High)", table_cell),
            Paragraph("Subjective emotional state self-reported by personnel. Strong inverse predictor of depressive/anxious symptoms.", table_cell)
        ],
    ]
    t_ml = Table(ml_features_data, colWidths=[90, 65, 80, printable_width - 235])
    t_ml.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.5, c_border),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, c_bg_light]),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_ml)
    story.append(Spacer(1, 8))

    ml_outcomes = (
        "<b>Model Inference & Actionable Outputs:</b><br/>"
        "&bull; <b>Low Stress:</b> Positive reinforcement, suggestions for maintaining physical training and good sleep hygiene.<br/>"
        "&bull; <b>Moderate Stress:</b> Non-clinical intervention advice, including 4-7-8 breathing exercises, tactical hydration, and micro-breaks.<br/>"
        "&bull; <b>High Stress:</b> Immediate alert badge, priority referral prompts to assigned Welfare Officer, and automated inclusion in officer urgent triage queues."
    )
    story.append(Paragraph(ml_outcomes, body_style))
    story.append(Spacer(1, 14))

    # =========================================================================
    # SECTION 5: DATABASE ARCHITECTURE & SCHEMA
    # =========================================================================
    story.append(Paragraph("5. Database Architecture & Schema Specification", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_primary, spaceBefore=2, spaceAfter=8))

    db_text = (
        "The application utilizes a normalized relational schema executed on <b>SQLite3</b>. "
        "The schema isolates sensitive military medical telemetry, assistance request records, authentication tokens, and user credentials."
    )
    story.append(Paragraph(db_text, body_style))
    story.append(Spacer(1, 6))

    schema_data = [
        [Paragraph("<b>Table Name</b>", table_header), Paragraph("<b>Key Columns & Constraints</b>", table_header), Paragraph("<b>Description & Relationship</b>", table_header)],
        [
            Paragraph("<code>Users</code>", code_style),
            Paragraph("<code>id (PK), name, unit, password, role, assigned_officer_id (FK)</code>", code_style),
            Paragraph("Stores personnel credentials (scrypt/PBKDF2 hashes), unit/station, assigned role, and linked officer. Supports weak password intercept and unit-verified forgot password recovery.", table_cell)
        ],
        [
            Paragraph("<code>DailyData</code>", code_style),
            Paragraph("<code>id (PK), user_id (FK), date, sleep_hours, duty_hours, heart_rate, mood_score, stress_level, recommendation</code>", code_style),
            Paragraph("Daily wellness check-ins. One record per user per day. Stores physiological telemetry and AI predictions.", table_cell)
        ],
        [
            Paragraph("<code>OfficerCodes</code>", code_style),
            Paragraph("<code>id (PK), code (UNIQUE), is_used (0|1)</code>", code_style),
            Paragraph("Cryptographically seeded one-time registration codes (<code>OFC-XXXXXX</code>) for authorized medical officers.", table_cell)
        ],
        [
            Paragraph("<code>AdminCodes</code>", code_style),
            Paragraph("<code>id (PK), code (UNIQUE), is_used (0|1)</code>", code_style),
            Paragraph("Single-use registration tokens (<code>ADM-XXXXXX</code>) for system administrators.", table_cell)
        ],
        [
            Paragraph("<code>StressRequests</code>", code_style),
            Paragraph("<code>id (PK), user_id (FK), message, status, stage, officer_id (FK), remedy, created_date, resolved_date, stage_updated_at</code>", code_style),
            Paragraph("Confidential assistance requests tracking through a 5-stage lifecycle (New &rarr; Acknowledged &rarr; In Progress &rarr; Resolved &rarr; Closed).", table_cell)
        ],
        [
            Paragraph("<code>CounselingNotes</code>", code_style),
            Paragraph("<code>id (PK), user_id (FK), officer_id (FK), date, notes, action_taken, follow_up_date</code>", code_style),
            Paragraph("Privileged clinical case notes maintained by welfare officers, including observations, action taken, and follow-up dates.", table_cell)
        ],
        [
            Paragraph("<code>ActivityLog</code>", code_style),
            Paragraph("<code>id (PK), actor_id (FK), action, target_user_id (FK), details, created_at</code>", code_style),
            Paragraph("Immutable audit trail logging all officer and admin actions for accountability and forensic integrity.", table_cell)
        ],
        [
            Paragraph("<code>PersonnelTags</code>", code_style),
            Paragraph("<code>id (PK), user_id (FK), tag (NOCASE), officer_id (FK), created_date, UNIQUE(user_id, tag)</code>", code_style),
            Paragraph("Archived table: Historical categorization tags preserved in database schema.", table_cell)
        ],
    ]
    t_schema = Table(schema_data, colWidths=[90, 200, printable_width - 290])
    t_schema.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.5, c_border),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, c_bg_light]),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_schema)
    story.append(Spacer(1, 14))

    # =========================================================================
    # SECTION 6: SERVERLESS CLOUD PERSISTENCE & BLOB SYNC
    # =========================================================================
    story.append(PageBreak())
    story.append(Paragraph("6. Serverless Cloud Persistence & Vercel Blob Sync Engine", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_primary, spaceBefore=2, spaceAfter=8))

    blob_intro = (
        "One of the most critical engineering challenges in serverless architectures is <b>filesystem ephemerality</b>. "
        "On Vercel, serverless function containers are destroyed and recreated frequently (container cold boots). "
        "Any SQLite database stored on <code>/tmp/apps.db</code> is wiped back to the git bundle on container restart, "
        "causing newly updated passwords, signups, and check-ins to vanish after idle periods.<br/><br/>"
        "To permanently rectify this, a custom <b>Bidirectional Cloud Object Storage Synchronization Engine</b> was engineered."
    )
    story.append(Paragraph(blob_intro, body_style))
    story.append(Spacer(1, 6))

    story.append(callout(
        "ARCHITECTURAL INNOVATION: SyncedConnection & Cloud Hydration",
        "1. Automatic Cold-Boot Hydration: On container startup, sync_db_from_blob() checks the private Vercel Blob store ('apps-data'). If a newer master apps.db exists, it is securely downloaded and hydrated into the container.\n"
        "2. Transparent Commit Hook: By subclassing sqlite3.Connection into SyncedConnection and injecting it as the connection factory, every commit() automatically triggers sync_db_to_blob(), pushing instantaneous database snapshots to private blob storage.\n"
        "3. Zero-Data-Loss Guarantee: Updates to passwords, new users, and welfare requests survive serverless container recycles indefinitely.",
        "#F0FDF4", "#16A34A"
    ))
    story.append(Spacer(1, 14))

    # =========================================================================
    # SECTION 7: WHAT WE ADDED & RECENT ENHANCEMENTS
    # =========================================================================
    story.append(Paragraph("7. Key Enhancements & Changelog (Recent Additions)", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_primary, spaceBefore=2, spaceAfter=8))

    story.append(Paragraph(
        "The following itemized features and structural enhancements were designed, coded, and deployed to production:",
        body_style
    ))
    story.append(Spacer(1, 4))

    features_added_data = [
        [Paragraph("<b>Category</b>", table_header), Paragraph("<b>Enhancement / Feature Added</b>", table_header), Paragraph("<b>Impact & Operational Benefit</b>", table_header)],
        [
            Paragraph("<b>Clinical Care</b>", table_cell),
            Paragraph("<b>Clinical Counseling Notes & Removal of Tags System:</b><br/>"
                      "Restored dedicated <b>Add Counseling Note</b> clinical module with notes textarea, action taken, and follow-up scheduling. "
                      "Equipped with mobile/desktop floating action button (<code>+ Add Note</code>) and 1-click <code>Schedule Counseling &rarr;</code> triage button on high-risk alerts. "
                      "Permanently decommissioned subjective tags across soldier dossiers and rosters to focus personnel records purely on objective physiological telemetry and clinical observations.", table_cell),
            Paragraph("Empowers officers to record structured clinical interventions while removing tag clutter from personnel listings.", table_cell)
        ],
        [
            Paragraph("<b>Login Reliability</b>", table_cell),
            Paragraph("<b>Login Diagnostics & Whitespace Resilience:</b><br/>"
                      "Engineered intelligent authentication diagnostics distinguishing unregistered accounts from password mismatches. "
                      "Provides 1-click action buttons (<code>[Register Account]</code> or <code>[Reset Password]</code>). "
                      "Added mobile trailing-space tolerance in password hashing and interactive visibility eye toggles (👁️) across all forms.", table_cell),
            Paragraph("Completely prevents false 'Invalid credentials' rejections caused by mobile keyboard autofill or user uncertainty.", table_cell)
        ],
        [
            Paragraph("<b>Security & Auth</b>", table_cell),
            Paragraph("<b>Self-Service Forgot Password & Recovery Engine:</b><br/>"
                      "Engineered dedicated <code>/forgot_password</code> portal with role selector (Personnel/Officer/Admin). "
                      "Verifies user identity via Registered Name, Unit, and Role without external email dependencies. "
                      "Features real-time client-side password strength validation and automatic cloud sync.", table_cell),
            Paragraph("Enables users to safely recover locked accounts instantly without administrator intervention while preserving high security standards.", table_cell)
        ],
        [
            Paragraph("<b>Security & Auth</b>", table_cell),
            Paragraph("<b>Weak Password Policy & Forced Upgrade Flow:</b><br/>"
                      "Enforced minimum 8-character length and complexity rules. Banned trivial passwords (<code>123</code>, <code>123456</code>, <code>password</code>). "
                      "Built mandatory <code>/upgrade_password</code> route that intercepts existing users with legacy passwords upon login.", table_cell),
            Paragraph("Completely eliminated vulnerabilities from default weak credentials; upgraded existing user accounts securely to scrypt hashes.", table_cell)
        ],
        [
            Paragraph("<b>Data Persistence</b>", table_cell),
            Paragraph("<b>Vercel Blob Storage Integration:</b><br/>"
                      "Provisioned private store <code>apps-data</code>, configured access tokens, built <code>SyncedConnection</code> and automatic two-way synchronization in <code>app.py</code>.", table_cell),
            Paragraph("Resolved the serverless container reset bug where password updates disappeared after overnight cold restarts.", table_cell)
        ],
        [
            Paragraph("<b>Workflow & Triage</b>", table_cell),
            Paragraph("<b>Universal 5-Stage Stress Request Pipeline:</b><br/>"
                      "Transformed requests into an omnipresent triage pipeline accessible to all on-duty officers. "
                      "Requests from all personnel dispatch directly to the Officer Dashboard with real-time pending badges and alert banners.", table_cell),
            Paragraph("Ensures no personnel message is missed, enabling collaborative review and transparent status tracking across officers.", table_cell)
        ],
        [
            Paragraph("<b>Timezone & Clock</b>", table_cell),
            Paragraph("<b>Indian Standard Time (IST) & Dynamic Device Greeting:</b><br/>"
                      "Standardized backend datetime calculations to IST (<code>UTC+05:30</code>) via <code>now_local()</code>. "
                      "Paired with client-side JavaScript dynamic clock synchronization for accurate morning, afternoon, and evening greetings.", table_cell),
            Paragraph("Eliminates UTC time-drift errors on check-ins and aligns dashboard greetings with the soldier's actual physical hour of day.", table_cell)
        ],
        [
            Paragraph("<b>Responsive Design</b>", table_cell),
            Paragraph("<b>Full Mobile & Laptop Responsive Layout:</b><br/>"
                      "Engineered adaptive breakpoints (<code>860px</code>, <code>720px</code>, <code>480px</code>). Added mobile slide drawer navigation, stacked stat cards, "
                      "responsive tables with horizontal scrolling, and touch-friendly 48px buttons.", table_cell),
            Paragraph("Seamless user experience across smartphones, iPads/tablets, and high-resolution laptop monitors.", table_cell)
        ],
        [
            Paragraph("<b>System Brand</b>", table_cell),
            Paragraph("<b>Official Full Form & Platform Rebranding:</b><br/>"
                      "Formally established and propagated the project definition: <b>AI Predictive Personnel Stress (Welfare and Monitoring System)</b> "
                      "across all portal headers, login branding, technical architecture docstrings, and generated physical dossiers.", table_cell),
            Paragraph("Establishes unified institutional identity and professional clarity across clinical and administrative tiers.", table_cell)
        ],
        [
            Paragraph("<b>Officer Dashboard</b>", table_cell),
            Paragraph("<b>Personnel Roster Single-Card Redesign:</b><br/>"
                      "Redesigned the soldier list in <code>officer_personnel.html</code> so each person occupies a clean, single card with inline unit badges and last check-in timestamp.", table_cell),
            Paragraph("Eliminated visual clutter and disorganized stacked text, presenting medical staff with a clear clinical roster.", table_cell)
        ],
        [
            Paragraph("<b>Clinical Safety</b>", table_cell),
            Paragraph("<b>'Going Quiet' Early Warning Watchlist:</b><br/>"
                      "Automated identification of soldiers who have missed daily check-ins for 3+ consecutive days, surfaced on a dedicated priority monitoring page.", table_cell),
            Paragraph("Enables welfare officers to proactively reach out to at-risk soldiers before mental distress escalates unnoticed.", table_cell)
        ],
        [
            Paragraph("<b>Reporting & Exports</b>", table_cell),
            Paragraph("<b>Sanitized Multi-Format Exports (Prioritized PDF Dossiers):</b><br/>"
                      "Prioritized official ReportLab PDF clinical dossiers on soldier dashboard while retaining comprehensive administrative Excel and CSV exports for commanders.", table_cell),
            Paragraph("Eliminates raw CSV clutter on soldier views while enabling senior officers and medical boards to produce official physical dossiers.", table_cell)
        ],
    ]
    t_feat = Table(features_added_data, colWidths=[90, 240, printable_width - 330], repeatRows=1)
    t_feat.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.5, c_border),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, c_bg_light]),
        ('PADDING', (0,0), (-1,-1), 4.5),
    ]))
    story.append(t_feat)
    story.append(Spacer(1, 14))

    # =========================================================================
    # SECTION 8: DEPLOYMENT & PRODUCTION SPECIFICATIONS
    # =========================================================================
    story.append(PageBreak())
    story.append(Paragraph("8. Deployment & Production Verification", h1_style))
    story.append(HRFlowable(width="100%", thickness=1, color=c_primary, spaceBefore=2, spaceAfter=8))

    deploy_data = [
        [Paragraph("<b>Parameter</b>", table_header), Paragraph("<b>Production Configuration Value</b>", table_header)],
        [Paragraph("<b>Production Web URL</b>", table_cell), Paragraph("<code>https://apps-nu-henna.vercel.app</code>", code_style)],
        [Paragraph("<b>Hosting Platform</b>", table_cell), Paragraph("Vercel Serverless Platform (Node.js/Python WSGI Bridge)", table_cell)],
        [Paragraph("<b>Object Storage Store</b>", table_cell), Paragraph("Vercel Blob Storage &bull; Store: <code>apps-data</code> (ID: <code>store_eS9r7map43Dx32yX</code>)", table_cell)],
        [Paragraph("<b>Source Control</b>", table_cell), Paragraph("Git (Dual-tracked: <code>origin/main</code> and <code>origin/master</code>)", table_cell)],
        [Paragraph("<b>Credential Recovery</b>", table_cell), Paragraph("Unit-verified self-service recovery (<code>/forgot_password</code>) active in production", table_cell)],
        [Paragraph("<b>Timezone Standard</b>", table_cell), Paragraph("Indian Standard Time (IST - <code>UTC+05:30</code>) with dynamic client clock sync", table_cell)],
        [Paragraph("<b>Secret Key Management</b>", table_cell), Paragraph("Environment-backed session secret key with stable production fallback", table_cell)],
        [Paragraph("<b>Database Verification</b>", table_cell), Paragraph("Verified live: User 'Dharun' upgraded to strong password, persistent across cold boots", table_cell)],
    ]
    t_deploy = Table(deploy_data, colWidths=[150, printable_width - 150])
    t_deploy.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), c_primary),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.5, c_border),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, c_bg_light]),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_deploy)
    story.append(Spacer(1, 12))

    story.append(callout(
        "DOCUMENT CONCLUSION & SYSTEM HEALTH STATUS",
        "The APPS platform is fully operational, verified in production, and hardened against security vulnerabilities, "
        "data loss on cold boots, and layout distortions across mobile and desktop devices. "
        "All features detailed in this specification are live and accessible.",
        "#F8FAFC", "#475569"
    ))

    # Build the document using the dynamic NumberedCanvas
    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"Document successfully created at: {output_path}")

if __name__ == "__main__":
    out_file = sys.argv[1] if len(sys.argv) > 1 else "APPS_System_Documentation.pdf"
    create_docs_pdf(out_file)
