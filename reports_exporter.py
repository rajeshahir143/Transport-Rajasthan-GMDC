"""Separate Excel exports for Rajasthan GMDC reports. Uses only data present in SQLite."""
from pathlib import Path
from datetime import datetime
import sqlite3

REPORTS = [
    "Agent Wise","Balance Sheet","Builty Commission","Daily Quota","DO Report",
    "General Ledger","P&L Account","Party Sales","Pending Bill","Pending DNote",
    "Purchase-Sale Summary","Purchase Summary","Sale Summary","Status DO",
    "Trading Account","Trial Balance","Vehicle Report"
]
ACCOUNTING_ONLY = {"Balance Sheet","General Ledger","P&L Account","Purchase-Sale Summary",
                   "Purchase Summary","Sale Summary","Trading Account","Trial Balance"}

def _tables(con):
    return {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}

def _cols(con, table):
    return {r[1] for r in con.execute(f'PRAGMA table_info("{table}")')}

def _date(value):
    value=str(value or "").strip()
    for fmt in ("%d/%m/%y","%d/%m/%Y","%Y-%m-%d","%d-%m-%y","%d-%m-%Y"):
        try: return datetime.strptime(value,fmt).date()
        except ValueError: pass
    return None

def _within(value,start,end):
    parsed=_date(value)
    return parsed is not None and start <= parsed <= end

def _records(con, report, start, end):
    tables=_tables(con)
    has=lambda t: t in tables
    cols=lambda t: _cols(con,t) if has(t) else set()
    rows=[]
    if report in ("DO Report","Status DO","Daily Quota") and has("dos"):
        for r in con.execute("SELECT * FROM dos ORDER BY id"):
            d=dict(r)
            dt=d.get("do_date") or d.get("po_date") or d.get("schedule_date")
            if _within(dt,start,end): rows.append(d)
        if report=="Status DO":
            for d in rows:
                total=float(d.get("qty") or 0); rem=float(d.get("remaining_qty") or 0)
                d["status"]="Pending" if rem>0 else "Completed"
                d["pending_qty_ton"]=rem
        if report=="Daily Quota":
            rows=[d for d in rows if d.get("schedule_date") or d.get("qty") is not None]
    elif report in ("Agent Wise","Builty Commission","Party Sales","Vehicle Report") and has("builtys"):
        bcols=cols("builtys"); dcols=cols("dos")
        select="b.*"
        join=" LEFT JOIN dos d ON d.id=b.do_id" if has("dos") else ""
        for r in con.execute("SELECT "+select+(", d.customer_name AS do_party, d.do_no AS linked_do_no, d.gmdc_do_no AS gmdc_do_no, d.transporter_name AS do_agent, d.material AS do_material" if join else "")+" FROM builtys b"+join+" ORDER BY b.id"):
            d=dict(r); dt=d.get("date")
            if not _within(dt,start,end): continue
            if report=="Agent Wise": d["report_agent"]=d.get("agent_name") or d.get("do_agent") or d.get("transporter_name") or ""
            if report=="Party Sales": d["report_party"]=d.get("sale_party") or d.get("supply_party") or d.get("do_party") or ""
            if report=="Vehicle Report": d["report_vehicle"]=d.get("vehicle_no") or d.get("truck_no") or ""
            if report=="Builty Commission": d["commission_amount"]=d.get("commission") or 0
            rows.append(d)
    elif report=="Pending Bill" and has("builtys"):
        join=" LEFT JOIN dos d ON d.id=b.do_id" if has("dos") else ""
        extra=", d.customer_name AS party, d.do_no AS do_no" if join else ""
        for r in con.execute("SELECT b.*"+extra+" FROM builtys b"+join+" WHERE COALESCE(b.billed,0)=0 ORDER BY b.date,b.id"):
            d=dict(r)
            if _within(d.get("date"),start,end): rows.append(d)
    elif report=="Pending DNote" and has("debit_notes"):
        for r in con.execute("SELECT * FROM debit_notes ORDER BY dn_date,id"):
            d=dict(r)
            if _within(d.get("dn_date"),start,end): rows.append(d)
    elif report=="DO Report" and has("po_builtys"):
        # Imported B167 records are retained as JSON payloads; export them as source records too.
        import json
        for r in con.execute("SELECT * FROM po_builtys ORDER BY id"):
            d=dict(r)
            if _within(d.get("lr_date"),start,end):
                try: d.update(json.loads(d.get("data_json") or "{}"))
                except Exception: pass
                rows.append(d)
    elif report=="Pending Bill" and has("bills"):
        for r in con.execute("SELECT * FROM bills ORDER BY bill_date,id"):
            d=dict(r)
            if _within(d.get("bill_date"),start,end) and str(d.get("status","")).casefold() in ("pending","unpaid"):
                rows.append(d)
    return rows

def export_report(database_path, report, start_text, end_text, output_path):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
    except ImportError as exc:
        raise RuntimeError("Excel export માટે openpyxl જરૂરી છે. Software update/install કરો.") from exc
    start,end=_date(start_text),_date(end_text)
    if not start or not end: raise ValueError("Date DD/MM/YY અથવા DD/MM/YYYY formatમાં આપો.")
    if start>end: raise ValueError("From Date, To Date કરતાં આગળ ન હોવી જોઈએ.")
    con=sqlite3.connect(str(database_path)); con.row_factory=sqlite3.Row
    try:
        rows=_records(con,report,start,end)
        wb=Workbook(); ws=wb.active; ws.title=report[:31]
        ws.append([report]); ws.append(["From",start.strftime("%d/%m/%Y"),"To",end.strftime("%d/%m/%Y")])
        if report in ACCOUNTING_ONLY:
            ws.append(["STATUS","આ રિપોર્ટ માટે જરૂરી Ledger/Accounting transaction tables હાલના databaseમાં ઉપલબ્ધ નથી. આંકડા બનાવટી રીતે ગણવામાં આવ્યા નથી."])
            ws.append(["આ રિપોર્ટ પૂર્ણ કરવા માટે અલગ accounting entries/ledger data જરૂરી છે."])
        elif not rows:
            ws.append(["No matching records in selected date range"])
        else:
            headers=[]
            for row in rows:
                for key in row:
                    if key not in headers and key!="data_json": headers.append(key)
            ws.append(headers)
            for row in rows: ws.append([row.get(k) for k in headers])
        ws.freeze_panes="A4"
        for cell in ws[1]:
            cell.font=Font(bold=True,size=14,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="1F4E78")
        for cell in ws[3]:
            cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="4472C4")
        for column in ws.columns:
            letter=get_column_letter(column[0].column)
            width=min(42,max(12,max((len(str(c.value or "")) for c in column[:100]),default=12)+2))
            ws.column_dimensions[letter].width=width
            for cell in column: cell.alignment=Alignment(vertical="center")
        wb.save(output_path)
        return len(rows)
    finally: con.close()
