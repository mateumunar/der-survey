import json
import os
import socket
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_file
import io

app = Flask(__name__)
# /tmp is writable on Render (and most cloud platforms); fall back to local dir
_data_dir = "/tmp" if os.path.isdir("/tmp") else os.path.dirname(__file__)
RESPONSES_FILE = os.path.join(_data_dir, "responses.json")


def load_responses():
    if not os.path.exists(RESPONSES_FILE):
        return []
    with open(RESPONSES_FILE, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def save_response(entry):
    responses = load_responses()
    responses.append(entry)
    with open(RESPONSES_FILE, "w", encoding="utf-8") as f:
        json.dump(responses, f, ensure_ascii=False, indent=2)


@app.route("/")
def index():
    return render_template("survey.html")


@app.route("/submit", methods=["POST"])
def submit():
    data = request.get_json()
    if not data or not data.get("name", "").strip():
        return jsonify({"error": "Name is required"}), 400
    data["submitted_at"] = datetime.now().isoformat()
    save_response(data)
    return jsonify({"ok": True})


@app.route("/responses")
def responses_raw():
    return jsonify(load_responses())


@app.route("/report")
def report():
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    responses = load_responses()
    doc = Document()

    # Title
    title = doc.add_heading("SAP DER Customer Survey – Summary Report", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  "
        f"Total respondents: {len(responses)}"
    ).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    if not responses:
        doc.add_paragraph("No responses have been submitted yet.")
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name="DER_Survey_Report.docx",
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    # ── Section 1: Summary statistics ─────────────────────────────────────
    doc.add_heading("Section 1 – Aggregate Results", level=1)

    questions = _get_questions()
    for q in questions:
        doc.add_heading(f"{q['id']}. {q['title']}", level=2)
        qkey = q["id"]

        if q["type"] == "text":
            doc.add_paragraph("Free-text responses:")
            for r in responses:
                ans = r.get(qkey, "").strip()
                if ans:
                    p = doc.add_paragraph(style="List Bullet")
                    p.add_run(f"{r.get('name', 'Anonymous')}: ").bold = True
                    p.add_run(ans)
            doc.add_paragraph()
            continue

        options = q.get("options", [])
        counts = {}
        if q["type"] in ("radio", "checkbox"):
            for opt in options:
                counts[opt] = 0
            for r in responses:
                ans = r.get(qkey)
                if isinstance(ans, list):
                    for a in ans:
                        if a in counts:
                            counts[a] += 1
                elif ans and ans in counts:
                    counts[ans] += 1

        # Table: option | count | %
        if counts:
            total = len(responses)
            table = doc.add_table(rows=1, cols=3)
            table.style = "Table Grid"
            hdr = table.rows[0].cells
            hdr[0].text = "Option"
            hdr[1].text = "Count"
            hdr[2].text = "%"
            for cell in hdr:
                for run in cell.paragraphs[0].runs:
                    run.bold = True

            sorted_opts = sorted(counts.items(), key=lambda x: -x[1])
            for opt, cnt in sorted_opts:
                row = table.add_row().cells
                row[0].text = opt
                row[1].text = str(cnt)
                row[2].text = f"{cnt/total*100:.0f}%"
            doc.add_paragraph()

        # "Other" free-text entries
        if q.get("other"):
            other_entries = [(r.get("name", "Anonymous"), r.get(f"{qkey}_other", "")) for r in responses]
            other_entries = [(n, v) for n, v in other_entries if v and v.strip()]
            if other_entries:
                doc.add_paragraph("Other (specified):")
                for name, val in other_entries:
                    p = doc.add_paragraph(style="List Bullet")
                    p.add_run(f"{name}: ").bold = True
                    p.add_run(val)

        # Sub-questions (Q2, Q9)
        for sub in q.get("subquestions", []):
            doc.add_paragraph(sub["label"], style="Intense Quote")
            sub_counts = {}
            for opt in sub["options"]:
                sub_counts[opt] = 0
            for r in responses:
                ans = r.get(sub["key"])
                if ans and ans in sub_counts:
                    sub_counts[ans] += 1
            total = len(responses)
            table = doc.add_table(rows=1, cols=3)
            table.style = "Table Grid"
            hdr = table.rows[0].cells
            hdr[0].text = "Option"; hdr[1].text = "Count"; hdr[2].text = "%"
            for cell in hdr:
                for run in cell.paragraphs[0].runs:
                    run.bold = True
            for opt, cnt in sorted(sub_counts.items(), key=lambda x: -x[1]):
                row = table.add_row().cells
                row[0].text = opt; row[1].text = str(cnt)
                row[2].text = f"{cnt/total*100:.0f}%"
            doc.add_paragraph()

    # ── Section 2: Individual responses ───────────────────────────────────
    doc.add_page_break()
    doc.add_heading("Section 2 – Individual Responses", level=1)

    for idx, r in enumerate(responses, 1):
        doc.add_heading(
            f"Respondent {idx}: {r.get('name', 'Anonymous')}  "
            f"({r.get('submitted_at', '')[:16]})",
            level=2,
        )
        for q in questions:
            qkey = q["id"]
            ans = r.get(qkey)
            if q["type"] == "compound":
                for sub in q.get("subquestions", []):
                    val = r.get(sub["key"], "—")
                    p = doc.add_paragraph(style="List Bullet")
                    p.add_run(f"{sub['label']}: ").bold = True
                    p.add_run(val or "—")
            else:
                if isinstance(ans, list):
                    display = "; ".join(ans) if ans else "—"
                else:
                    display = ans or "—"
                other_val = r.get(f"{qkey}_other", "")
                if other_val:
                    display += f" (other: {other_val})"
                p = doc.add_paragraph(style="List Bullet")
                p.add_run(f"Q{qkey}: ").bold = True
                p.add_run(display)
        doc.add_paragraph()

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="DER_Survey_Report.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


def _get_questions():
    cap10 = [
        "Customer enrollment and DER program management, including incentives and onboarding",
        "Device connectivity and protocol handling for EV chargers, batteries, heat pumps and solar",
        "Demand side management and demand response execution",
        "Energy data management (validated interval data, quality status, correction history, versioning, audit trail)",
        "Virtual power plant and flexibility market participation",
        "Energy communities and energy sharing",
        "Modeling of complex prosumer installations (multiple assets, registers and directions behind a single connection point)",
        "Grid capacity, congestion management and hosting capacity",
        "Settlement, billing determinants and flexibility remuneration",
        "Analytics and AI applied to energy data",
    ]
    return [
        {
            "id": "q1",
            "title": "Which best describes your organization's primary business model? Select one.",
            "type": "radio", "other": True,
            "options": [
                "Retail or supply business",
                "Distribution network operator",
                "Integrated utility, both network and retail",
                "Generation, aggregator or flexibility service provider",
                "Municipal utility or cooperative",
                "Other",
            ],
        },
        {
            "id": "q2", "title": "In which market do you primarily operate?",
            "type": "radio",
            "options": ["Europe", "North America", "Asia Pacific and Japan", "Latin America", "Middle East and Africa"],
        },
        {
            "id": "q3", "title": "What is your core system today?",
            "type": "radio",
            "options": ["SAP S/4HANA Utilities", "SAP ECC IS-U", "Non-SAP core", "Ongoing transformation to SAP S/4HANA Utilities"],
        },
        {
            "id": "q4",
            "title": "Thinking about the next 24 months, which capabilities matter most to your business? Select up to three.",
            "type": "checkbox", "max": 3, "options": cap10,
        },
        {
            "id": "q5",
            "title": "Which of those same capabilities do you expect to obtain primarily from an internal solution or a specialist provider rather than directly from SAP? Select up to three.",
            "type": "checkbox", "max": 3, "options": cap10,
        },
        {
            "id": "q6",
            "title": "Where does your validated interval energy data live today, and who controls it? Select all that apply.",
            "type": "checkbox", "max": 7,
            "options": [
                "SAP S/4HANA Utilities or SAP ECC IS-U",
                "SAP Cloud for Energy",
                "A meter data management system from our metering vendor",
                "A specialist third-party energy data platform",
                "Our own data lake or in-house build",
                "Outsourced to a service provider",
                "Not consolidated — it sits across several systems",
            ],
        },
        {
            "id": "q7",
            "title": "Could you reproduce a settlement determination or network investment decision from two years ago, including the original reading, validation status, and correction history — from a system you control and without vendor cooperation?",
            "type": "radio",
            "options": [
                "Yes, confidently and from a system we control",
                "Yes, but it would take significant manual effort",
                "Only with the cooperation of a software vendor or service provider",
                "No",
                "We have not tested this",
            ],
        },
        {
            "id": "q8",
            "title": "How would you characterize the need for a governed enterprise record of energy data (validated interval data, quality status, versioning, audit trail)?",
            "type": "radio",
            "options": [
                "Mandatory foundation — other DER capabilities do not work without it",
                "Important, but it can follow other DER capabilities",
                "Useful to have, not a priority",
                "Already solved for us — not a gap",
                "Not relevant to our business model",
            ],
        },
        {
            "id": "q9",
            "title": "Regardless of what any vendor offers today, which three of the following would be most valuable to you? Select up to three.",
            "type": "checkbox", "max": 3, "other": True,
            "options": [
                "Modeling of energy types at register level (consumed, produced, stored, active and reactive)",
                "Near real-time ingestion of smart meter data",
                "Validation, estimation and interpolation with an explicit quality status on every value",
                "Versioning, correction history and full audit trail",
                "Definition of energy communities and calculation of flows between participants",
                "Aggregation across assets, portfolios and communities",
                "Price, weather and other non-energy series held in the same interval model",
                "Forecasting",
                "Synthetic load profile management",
                "Extensibility to apply our own estimation and forecasting logic",
                "Billing determination",
                "Energy settlement",
                "Analytics and AI data products built on the governed record",
                "Other, please specify",
            ],
        },
        {
            "id": "q10", "title": "Where should this governed energy data record sit, and when do you need it?",
            "type": "compound",
            "subquestions": [
                {"key": "q10_placement", "label": "Placement", "options": [
                    "In the SAP enterprise platform",
                    "In a specialist partner solution integrated with SAP",
                    "In our own platform under our control",
                    "No preference — provided it is open and interoperable",
                ]},
                {"key": "q10_timing", "label": "Timing", "options": [
                    "Already needed or overdue",
                    "Within twelve months",
                    "Twelve to twenty-four months",
                    "Twenty-four to thirty-six months",
                    "No defined need",
                ]},
            ],
        },
        {
            "id": "q11",
            "title": "What is the single DER capability you most need and do not have today?",
            "type": "text",
        },
        {
            "id": "q12",
            "title": "What is the main barrier preventing your organization from obtaining it?",
            "type": "text",
        },
        {
            "id": "q13",
            "title": "How important is it that Energy Data Management capability can be deployed independently of your core-system upgrade timeline?",
            "type": "radio", "optional": True,
            "options": [
                "Essential — we will not take on a core dependency",
                "Very important",
                "Important",
                "Neutral",
                "Not important",
            ],
        },
        {
            "id": "q14", "title": "What is your preferred deployment mode?",
            "type": "radio", "optional": True, "other": True,
            "options": [
                "Public cloud",
                "Private cloud",
                "Within the existing core system",
                "Independent platform integrated with the core system",
                "No preference",
                "Other",
            ],
        },
    ]


if __name__ == "__main__":
    local_ip = socket.gethostbyname(socket.gethostname())
    print(f"\n  Survey running at:")
    print(f"    Local:   http://localhost:5000")
    print(f"    Network: http://{local_ip}:5000")
    print(f"\n  Download report: http://localhost:5000/report\n")
    app.run(host="0.0.0.0", port=5000, debug=False)
