"""
This module defines the blueprint and routes for the Balance Sheet generation feature.
It handles multiple pages (tabs) of the balance sheet form and integrates with the
Excel generation logic to produce the final downloadable report.
"""
from blueprints.excel_gen import generate_balance_sheet  # pylint: disable=import-error
from flask import (  # pylint: disable=import-error
    Blueprint,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

balance_sheet_bp = Blueprint("balance_sheet", __name__, template_folder="templates")


def get_form_data(req_form):
    """
    Helper to extract form data correctly.
    Standard to_dict() flattens lists (lossy for dynamic rows).
    to_dict(flat=False) makes everything a list.
    We mix them: default flat, but keep known lists or keys ending in [] as lists.
    """
    data = req_form.to_dict()
    # Explicitly handle known list fields or pattern matching
    # Check for keys ending in brackets, e.g., "sch3_particulars[]"
    for key in req_form.keys():
        if key.endswith("[]") or key in ["sch3_particulars", "sch3_amount"]:
            data[key] = req_form.getlist(key)
    return data


@balance_sheet_bp.route("/balance-sheet")
def index():
    """Redirects to the home page of the balance sheet."""
    return render_template("balance_sheet/home.html", active_tab="home", page_title="Home")


@balance_sheet_bp.route("/balance-sheet/page1", methods=["GET", "POST"])
def page1():
    """Rank PAGE 1: Schedules 1-4 (Liabilities & Assets)."""
    if request.method == "POST":
        # Save Page 1 data to session or DB using helper to keep lists
        session["bs_page1"] = get_form_data(request.form)
        return redirect(url_for("balance_sheet.page2"))

    data = session.get("bs_page1", {})
    return render_template(
        "balance_sheet/page1.html", active_tab="page1", data=data, page_title="Page 1 (Schedules 1-4)"
    )


@balance_sheet_bp.route("/balance-sheet/page2", methods=["GET", "POST"])
def page2():
    """Rank PAGE 2: Schedules 8-13 (Advances, Income Outstanding, Cash/Bank)."""
    if request.method == "POST":
        session["bs_page2"] = get_form_data(request.form)
        return redirect(url_for("balance_sheet.page3"))

    data = session.get("bs_page2", {})
    return render_template(
        "balance_sheet/page2.html", active_tab="page2", data=data, page_title="Page 2 (Schedules 8-13)"
    )


@balance_sheet_bp.route("/balance-sheet/page3", methods=["GET", "POST"])
def page3():
    """Rank PAGE 3: Schedule 6 (Investments)."""
    if request.method == "POST":
        session["bs_page3"] = get_form_data(request.form)
        return redirect(url_for("balance_sheet.page4"))

    data = session.get("bs_page3", {})
    return render_template("balance_sheet/page3.html", active_tab="page3", data=data, page_title="Page 3 (Investments)")


@balance_sheet_bp.route("/balance-sheet/page4", methods=["GET", "POST"])
def page4():
    """Rank PAGE 4: Schedules 5 & 7 (Immovable Properties & Furniture)."""
    if request.method == "POST":
        session["bs_page4"] = get_form_data(request.form)
        return redirect(url_for("balance_sheet.page4c"))

    data = session.get("bs_page4", {})
    return render_template("balance_sheet/page4.html", active_tab="page4", data=data, page_title="Page 4 (Assets)")


@balance_sheet_bp.route("/balance-sheet/page4c", methods=["GET", "POST"])
def page4c():
    """Rank PAGE 4C: Schedule IXC (Contribution)."""
    if request.method == "POST":
        session["bs_page4c"] = get_form_data(request.form)
        return redirect(url_for("balance_sheet.page4d"))

    data = session.get("bs_page4c", {})
    return render_template(
        "balance_sheet/page4c.html", active_tab="page4c", data=data, page_title="Page 4C (Contribution)"
    )


@balance_sheet_bp.route("/balance-sheet/page4d", methods=["GET", "POST"])
def page4d():
    """Rank PAGE 4D: Schedule IX-D (Auditor Information)."""
    if request.method == "POST":
        session["bs_page4d"] = get_form_data(request.form)
        return redirect(url_for("balance_sheet.page_winman"))

    data = session.get("bs_page4d", {})
    return render_template("balance_sheet/page4d.html", active_tab="page4d", data=data, page_title="Page 4D (History)")


@balance_sheet_bp.route("/balance-sheet/page_winman", methods=["GET", "POST"])
def page_winman():
    """Rank Page Winman: Income Details & Application of Income."""
    if request.method == "POST":
        session["bs_winman"] = get_form_data(request.form)
        return redirect(url_for("balance_sheet.page_corpus"))

    data = session.get("bs_winman", {})
    return render_template("balance_sheet/page_winman.html", active_tab="page_winman", data=data, page_title="Winman")


@balance_sheet_bp.route("/balance-sheet/page_corpus", methods=["GET", "POST"])
def page_corpus():
    """Rank Page Corpus: Schedule J & R details."""
    if request.method == "POST":
        session["bs_corpus"] = get_form_data(request.form)
        return redirect(url_for("balance_sheet.page_accumulation"))

    data = session.get("bs_corpus", {})
    return render_template(
        "balance_sheet/page_corpus.html", active_tab="page_corpus", data=data, page_title="Corpus Fund"
    )


@balance_sheet_bp.route("/balance-sheet/page_accumulation", methods=["GET", "POST"])
def page_accumulation():
    """Rank Page Accumulation: Schedule I Details."""
    if request.method == "POST":
        # Save Accumulation data
        session["bs_accumulation"] = get_form_data(request.form)
        # Finish -> Redirect to Generation
        return redirect(url_for("balance_sheet.generate_excel_route"))

    data = session.get("bs_accumulation", {})
    return render_template(
        "balance_sheet/page_accumulation.html", active_tab="page_accumulation", data=data, page_title="Accumulation"
    )


@balance_sheet_bp.route("/balance-sheet/generate")
def generate_excel_route():
    """Generates and downloads the Balance Sheet Excel."""
    try:
        output = generate_balance_sheet(session)
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="Balance_Sheet.xlsx",
        )
    except Exception as e:  # pylint: disable=broad-except
        return f"Error generating Excel: {str(e)}", 500
