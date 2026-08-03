"""
Resume Builder for LaKeysha Strickland
Generates a professionally designed PDF version of the resume.
"""

from fpdf import FPDF
import os

# --- Hyperlinks ---
LINKEDIN_URL = "https://www.linkedin.com/in/lakeysha-strickland/"
GITHUB_PROJECT_URL = "https://phoenixkey87-cmyk.github.io/Keysha_40th_Vietnam/"

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

# --- Resume Data ---
NAME = "LaKeysha Strickland"
CREDENTIALS = "MBA, CSM, CSPO, AWS AI Practitioner"
TITLE = "SYSTEMS ENGINEER II | AUTOMATION & CHANGE MANAGEMENT"
CONTACT_PREFIX = "St. Louis, MO  |  314-307-4921  |  keystrickland@charter.net  |  "

SUMMARY = (
    "Results-driven Systems Engineer II with 12+ years of experience spanning enterprise network operations, "
    "change governance, and AI-assisted automation. Specializes in building production-ready tools with "
    "Python, Microsoft Power Platform, and Amazon Kiro that prevent service-impacting outages, eliminate "
    "manual processes, and give operations teams faster, data-driven decisions. Architected an enterprise "
    "Deconfliction Dashboard that prevents 10-19 potential outages daily and automated carrier "
    "maintenance workflows across 13+ provider relationships. Recognized for delivering high-impact "
    "solutions, standardizing governance processes, and training cross-functional teams at scale."
)

SKILLS = {
    "Automation & Development": [
        "Python", "Amazon Kiro", "Amazon Q", "Claude Code", "Power Apps",
        "Power Automate", "AI-Assisted Development", "GitHub / GitHub Pages",
        "Outlook COM Automation", "REST APIs", "Regex"
    ],
    "Cloud & Platforms": [
        "AWS", "Azure", "SharePoint", "Microsoft Teams", "Microsoft 365"
    ],
    "Operations & Process": [
        "Change Management", "ITIL Framework", "Risk & Impact Analysis",
        "Method of Procedure (MOP)", "SOP Development", "Agile / Scrum",
        "Six Sigma", "Compliance & Governance"
    ],
    "Tools & Reporting": [
        "BMC Helix", "Jira", "Tableau", "Excel / openpyxl",
        "Technical Training", "Stakeholder Management", "UAT"
    ],
}

EXPERIENCE = [
    {
        "title": "Systems Engineer II - Automation & Change Management",
        "company": "Spectrum (Charter Communications)  |  St. Louis, MO",
        "dates": "Mar 2026 - Present",
        "bullets": [
            "Architected and deployed an enterprise Deconfliction Dashboard and Portal using Amazon Kiro and Microsoft Power Apps that identifies overlapping maintenance activity before implementation, preventing 10-19 potential customer-impacting outages daily.",
            "Built a Python-based carrier maintenance automation system that monitors a shared Outlook mailbox, detects notifications from 13+ carriers via regex parsing, and auto-generates pre-filled CRQ records mapped to CM Portal MOP Mgmt fields -- eliminating manual inbox monitoring entirely.",
            "Coordinate ~50 enterprise change tickets weekly and ~20 third-party carrier tickets daily for Comcast, Lumen, Netflix, and other telecommunications and content partners.",
            "Collaborate with 20+ engineering and operations teams nationwide to validate impacts, correct ticket data, enforce approved MOPs, and reduce operational risk.",
            "Updated and standardized 6 SOPs to strengthen governance, improve consistency, and clarify execution requirements across engineering workflows.",
            "Train employees on impact corrections, carrier ticket workflows, MOP selection, and internal tools; produce video guides and technical documentation to accelerate adoption.",
            "Leverage Amazon Q for code support, documentation, testing, and workflow automation; represent the team in cloud-related meetings and communicate updates to stakeholders.",
        ],
    },
    {
        "title": "Change Control Analyst II",
        "company": "Spectrum (Charter Communications)  |  St. Louis, MO",
        "dates": "Mar 2019 - Mar 2026",
        "bullets": [
            "Led analytic oversight of scheduled network maintenance, identifying resource and cross-service collisions across a converged enterprise network serving millions of subscribers.",
            "Served as primary point of contact for third-party carrier maintenance, coordinating complex changes among engineering, operations, project teams, and external providers.",
            "Prepared operational reporting, analyzed maintenance trends, and recommended process improvements that strengthened compliance and network reliability.",
            "Developed change support documentation, manuals, procedures, and training materials while performing impact assessments and risk reviews in alignment with ITIL practices.",
        ],
    },
    {
        "title": "Change Control Analyst",
        "company": "Spectrum (Charter Communications)  |  St. Louis, MO",
        "dates": "Jan 2017 - Mar 2019",
        "bullets": [
            "Coordinated and facilitated scheduled network maintenance while enforcing change policies designed to protect the Spectrum network and revenue-generating customers.",
            "Analyzed customer and service impacts, escalated conflicts, and supported carrier-class operating practices across enterprise maintenance activity.",
        ],
    },
    {
        "title": "Internet & Voice Repair Technician",
        "company": "Spectrum (Charter Communications)  |  St. Louis, MO",
        "dates": "Jan 2014 - Jan 2017",
        "bullets": [
            "Resolved complex internet and voice service issues, managed customer escalations, and consistently met service-quality and customer-experience targets.",
            "Earned the Spectrum BEST Award for outstanding customer service performance.",
        ],
    },
]

PROJECTS = [
    {
        "title": "Enterprise Deconfliction Dashboard & Portal",
        "tech": "Amazon Kiro, Microsoft Power Apps, SharePoint, AI-Assisted Development",
        "bullets": [
            "Took the solution from business problem through design, development, testing, user adoption, and ongoing enhancement; partnered with a colleague for SQL integration support.",
            "Centralized maintenance visibility so specialists can identify ticket collisions before execution, reducing outage risk and eliminating hours of manual research per day.",
            "Delivers proactive conflict detection used daily by the Change Management team to make informed go/no-go decisions on scheduled maintenance.",
        ],
    },
    {
        "title": "Third-Party Carrier Maintenance Automation",
        "tech": "Python, Outlook COM (pywin32), Microsoft Teams Webhooks, openpyxl, Amazon Kiro",
        "bullets": [
            "Engineered an end-to-end automated workflow that connects to the ChangeManagement-Carrier shared mailbox via COM automation, detects maintenance emails by subject-line keywords, and extracts structured data using configurable regex patterns.",
            "Sends real-time Adaptive Card alerts to Microsoft Teams upon detection, giving the team instant visibility into incoming carrier work without manual inbox monitoring.",
            "Auto-generates pre-filled Carrier CRQ records with impact auto-determination logic, scheduled dates, and carrier-specific assignee group mapping for 13+ providers.",
            "Maintains a persistent tracking spreadsheet with duplicate detection, enabling full ticket lifecycle management from detection through CRQ creation.",
        ],
    },
    {
        "title": "Phoenix Travel Digital Experiences",
        "tech": "Amazon Kiro, GitHub, GitHub Pages, HTML/CSS",
        "bullets": [
            "Built and deployed live customer-facing travel sites featuring interactive itineraries, RSVP pages, and payment summaries for group travel clients.",
            "Used AI-assisted development to translate business requirements into polished web experiences, managing source control and production publishing through GitHub Pages.",
        ],
    },
]

EDUCATION = [
    ("Master of Business Administration (MBA)", "Louisiana State University", "2026"),
    ("Bachelor of Science, Business Administration", "Lindenwood University - Dean's List", "2021"),
]

CERTIFICATIONS = [
    "AWS Certified AI Practitioner",
    "Certified Scrum Product Owner (CSPO), Scrum Alliance, 2023",
    "Certified ScrumMaster (CSM), Scrum Alliance, 2022",
    "Project Management Certificate, NYU (Pathstream), 2024",
    "Six Sigma: Green Belt",
    "Six Sigma: Yellow Belt",
    "Scrum Fundamentals Certified (SFC)",
]

AWARDS = [
    "Spectrum BEST Award - recognized for outstanding customer service performance.",
    "Selected by a director for consideration for the WICT Women to Watch recognition.",
    "Active in NAMIC and WICT professional-development initiatives; Make-A-Wish volunteer.",
]


# =============================================================================
# PDF GENERATION
# =============================================================================

def sanitize(text):
    """Replace Unicode characters that core fonts can't render."""
    replacements = {
        "\u2013": "-",
        "\u2014": "--",
        "\u2018": "'",
        "\u2019": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2022": "-",
        "\u2026": "...",
        "\u00b7": " - ",
    }
    for orig, repl in replacements.items():
        text = text.replace(orig, repl)
    return text


class ResumePDF(FPDF):
    """Custom PDF class for resume."""

    # Colors
    NAVY = (61, 79, 95)
    ACCENT = (184, 115, 51)
    DARK = (45, 45, 45)
    LIGHT = (100, 100, 100)

    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=15)

    def header_section(self):
        """Name, credentials, title, contact."""
        # Name
        self.set_font("Helvetica", "B", 22)
        self.set_text_color(*self.NAVY)
        self.cell(0, 10, sanitize(NAME), align="C", new_x="LMARGIN", new_y="NEXT")

        # Credentials
        self.set_font("Helvetica", "", 10)
        self.set_text_color(*self.ACCENT)
        self.cell(0, 6, sanitize(CREDENTIALS), align="C", new_x="LMARGIN", new_y="NEXT")

        # Title
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*self.DARK)
        self.cell(0, 6, sanitize(TITLE), align="C", new_x="LMARGIN", new_y="NEXT")

        # Contact with links
        self.set_font("Helvetica", "", 8)
        self.set_text_color(*self.DARK)
        prefix_w = self.get_string_width(CONTACT_PREFIX)
        linkedin_w = self.get_string_width("LinkedIn")
        sep_w = self.get_string_width("  |  ")
        project_w = self.get_string_width("Live Project")
        total_w = prefix_w + linkedin_w + sep_w + project_w
        start_x = (self.w - total_w) / 2

        y = self.get_y()
        self.set_xy(start_x, y)
        self.cell(prefix_w, 5, sanitize(CONTACT_PREFIX))
        self.set_text_color(*self.ACCENT)
        self.cell(linkedin_w, 5, "LinkedIn", link=LINKEDIN_URL)
        self.set_text_color(*self.DARK)
        self.cell(sep_w, 5, "  |  ")
        self.set_text_color(*self.ACCENT)
        self.cell(project_w, 5, "Live Project", link=GITHUB_PROJECT_URL)
        self.ln(8)

    def section_heading(self, text):
        """Section heading with accent underline."""
        self.ln(3)
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(*self.NAVY)
        self.cell(0, 7, sanitize(text), new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(*self.ACCENT)
        self.set_line_width(0.5)
        self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
        self.ln(3)

    def body_text(self, text):
        """Body paragraph."""
        self.set_font("Helvetica", "", 9)
        self.set_text_color(*self.DARK)
        self.multi_cell(0, 4.5, sanitize(text))
        self.ln(2)

    def bullet_point(self, text):
        """Indented bullet."""
        self.set_font("Helvetica", "", 9)
        self.set_text_color(*self.DARK)
        x = self.get_x()
        self.set_x(x + 4)
        self.multi_cell(0, 4.5, sanitize(f"  -  {text}"))
        self.ln(1)

    def job_header(self, title, company, dates):
        """Job title with dates on right."""
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*self.NAVY)
        # Title
        title_w = self.get_string_width(sanitize(title))
        self.cell(title_w + 2, 6, sanitize(title))
        # Dates
        self.set_font("Helvetica", "", 8)
        self.set_text_color(*self.ACCENT)
        self.cell(0, 6, sanitize(dates), align="R", new_x="LMARGIN", new_y="NEXT")
        # Company
        self.set_font("Helvetica", "", 8)
        self.set_text_color(*self.LIGHT)
        self.cell(0, 5, sanitize(company), new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def project_header(self, title, tech):
        """Project title and tech stack."""
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(*self.NAVY)
        self.cell(0, 6, sanitize(title), new_x="LMARGIN", new_y="NEXT")
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(*self.LIGHT)
        self.cell(0, 4, sanitize(tech), new_x="LMARGIN", new_y="NEXT")
        self.ln(1)


def build_pdf_resume():
    """Generate the PDF resume."""
    pdf = ResumePDF()
    pdf.add_page()
    pdf.set_margins(18, 15, 18)

    # Header
    pdf.header_section()

    # Professional Summary
    pdf.section_heading("PROFESSIONAL SUMMARY")
    pdf.body_text(SUMMARY)

    # Core Skills
    pdf.section_heading("CORE SKILLS")
    for category, skills in SKILLS.items():
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_text_color(*ResumePDF.NAVY)
        pdf.cell(0, 5, sanitize(category.upper()), new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(*ResumePDF.DARK)
        pdf.multi_cell(0, 4, sanitize("  |  ".join(skills)))
        pdf.ln(2)

    # Professional Experience
    pdf.section_heading("PROFESSIONAL EXPERIENCE")
    for job in EXPERIENCE:
        pdf.job_header(job["title"], job["company"], job["dates"])
        for bullet in job["bullets"]:
            pdf.bullet_point(bullet)
        pdf.ln(2)

    # Key Projects
    pdf.section_heading("KEY PROJECTS")
    for proj in PROJECTS:
        pdf.project_header(proj["title"], proj["tech"])
        for bullet in proj["bullets"]:
            pdf.bullet_point(bullet)
        pdf.ln(2)

    # Education & Certifications
    pdf.section_heading("EDUCATION & CERTIFICATIONS")
    for degree, school, year in EDUCATION:
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_text_color(*ResumePDF.DARK)
        pdf.cell(0, 5, sanitize(f"{degree}  |  {school}  |  {year}"), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(*ResumePDF.DARK)
    for cert in CERTIFICATIONS:
        pdf.cell(0, 4.5, sanitize(f"  -  {cert}"), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    # Awards
    pdf.section_heading("AWARDS & COMMUNITY LEADERSHIP")
    for award in AWARDS:
        pdf.bullet_point(award)

    # Save
    pdf_path = os.path.join(OUTPUT_DIR, "LaKeysha_Strickland_Resume.pdf")
    pdf.output(pdf_path)
    print(f"[OK] PDF saved: {pdf_path}")
    return pdf_path


# =============================================================================
# MAIN
# =============================================================================

if __name__ == "__main__":
    print("Building resume PDF...")
    print("-" * 50)
    build_pdf_resume()
    print("-" * 50)
    print("Done!")
