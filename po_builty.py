"""PO Builty form and B167-style PDF, linked to DO quantities in Rajasthan GMDC mode."""
import json, re, tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from db import db, get_active_company

# Fixed B167 consignor / consignee master details for Rajasthan Builty.
FIXED_CONSIGNOR_NAME = "KARAMTARA ENGINEERING LTD"
FIXED_CONSIGNOR_ADDRESS = (
    "15 NO. 112/P18P2, 113/22, 116/2, 118, "
    "119/2, 119/1/P18P2 & P3, 120, VILL-MEGHPAR-KUNJISAR, TAL-BHACHAU"
)
FIXED_CONSIGNOR_GSTIN = "24AABCK1921E125"
FIXED_CONSIGNEE_NAME = "KARAMTARA USA INC"
FIXED_CONSIGNEE_ADDRESS = (
    "95 HILLSIDE PARK DR, NEW HYDE PARK, NEW YORK. "
    "Pin Code: 11040; State Code/State: /"
)
FIXED_CONSIGNEE_GSTIN = "URP"

FIELDS = [
("po_no","PO No."),("po_date","PO Date"),("po_buyer","Buyer / Customer"),("po_vendor","Vendor / Supplier"),
("po_material","PO Material"),("po_qty","PO Quantity"),("po_uom","PO UOM"),("po_rate","PO Rate"),("po_delivery_date","Delivery Date"),
("builty_qty","Builty Qty (TON)"),
("builty_company","Builty Company Name"),
("builty_address","Builty Company Address"),
("lr_no","LR No."),("lr_date","LR Date"),("consignor_name","Consignor Name"),
("consignor_address","Consignor Address"),("consignor_gstin","Consignor GSTIN"),
("consignee_name","Consignee Name"),("consignee_address","Consignee Address"),
("consignee_gstin","Consignee GSTIN"),("invoice_no","Invoice No."),("challan_no","Challan No."),
("eway_bill_no","E-Way Bill No."),("eway_generate_date","E-Way Bill Generate Date"),
("eway_valid_date","E-Way Bill Valid Date"),("packing_list_no","Packing List No."),
("description","Description"),("bundle_1","Bundle 1"),("total_bundle","Total Bundle 2"),("weight_1","Weight 1"),
("weight_2","Weight 2"),("truck_no","Truck No."),("driver_name","Driver Name"),
("driver_address","Address"),("truck_owner","Truck Owner Name"),("mobile_no","Mobile No."),
("cgst","CGST"),("sgst","SGST"),("igst","IGST"),("total","Total"),("notes","Notes / Terms")]

class POBuiltyTab:
    def __init__(self, nb, firm_getter, export_dir, open_file, create_do_callback=None):
        self.firm_getter, self.export_dir, self.open_file = firm_getter, Path(export_dir), open_file
        self.create_do_callback = create_do_callback
        with db() as con:
            con.execute("""CREATE TABLE IF NOT EXISTS po_builtys(
              id INTEGER PRIMARY KEY AUTOINCREMENT, lr_no TEXT NOT NULL, lr_date TEXT,
              data_json TEXT NOT NULL DEFAULT '{}',
              created_at TEXT DEFAULT (datetime('now','localtime')),
              updated_at TEXT DEFAULT (datetime('now','localtime')))""")
        f=ttk.Frame(nb,padding=8); self.frame=f
        is_rajasthan = str(get_active_company()).casefold() == "rajasthan gmdc"
        nb.add(f,text="2 · Builty Entry" if is_rajasthan else "PO Builty")
        ttk.Label(f,text="Builty Entry",font=("Segoe UI",17,"bold"),foreground="#B22222").pack(anchor="w")
        # GMDC-style two-column workspace: entry form at left, order and saved
        # Builty navigation at right.
        split=ttk.Panedwindow(f,orient="horizontal"); split.pack(fill="both",expand=True)
        left=ttk.Frame(split,padding=4); right=ttk.Frame(split,padding=4)
        split.add(left,weight=3); split.add(right,weight=3)

        pending_box=ttk.LabelFrame(right,text="Pending PO — select a PO to auto-fill details",padding=4)
        pending_box.pack(fill="x",pady=(0,5))
        pending_head=ttk.Frame(pending_box); pending_head.pack(fill="x")
        ttk.Button(pending_head,text="Refresh POs",command=self.refresh_all).pack(side="right",padx=2,pady=(0,3))
        pcols=("id","po","party","material","pending")
        self.pending_tree=ttk.Treeview(pending_box,columns=pcols,show="headings",height=5,selectmode="browse")
        for key,label,width in (("id","ID",42),("po","PO No.",115),("party","Party Name",180),("material","Item Name",150),("pending","Pending TON",90)):
            self.pending_tree.heading(key,text=label)
            self.pending_tree.column(key,width=width,anchor="w" if key in ("party","material") else "center")
        self.pending_tree.pack(fill="x",expand=False)
        self.pending_tree.bind("<<TreeviewSelect>>",self.on_pending_po_select)

        details=ttk.LabelFrame(right,text="DO Details (Auto from PO)",padding=6)
        details.pack(fill="x",pady=(0,8))
        self.po_details_label=ttk.Label(details,text="Select a pending PO to auto-fill party, GSTIN, material and available quantity.",wraplength=520,justify="left")
        self.po_details_label.pack(anchor="w")

        saved_head=ttk.Frame(right); saved_head.pack(fill="x",pady=(2,3))
        ttk.Label(saved_head,text="Saved Builty List",font=("Segoe UI",11,"bold")).pack(side="left")
        ttk.Label(saved_head,text="Search:").pack(side="left",padx=(12,3))
        self.search_var=tk.StringVar()
        search=ttk.Entry(saved_head,textvariable=self.search_var,width=18)
        search.pack(side="left")
        self.search_var.trace_add("write",lambda *_:self.refresh())
        ttk.Button(saved_head,text="Selected PO Builty",command=self.show_selected_po_builtys).pack(side="right",padx=2)
        ttk.Button(saved_head,text="All Builty",command=self.show_all_builtys).pack(side="right",padx=2)
        self.builty_filter="all"

        # Keep PO linkage hidden from the form; selection is made in the right list.
        self.do_var=tk.StringVar()
        self.do_combo=ttk.Combobox(f,textvariable=self.do_var,state="readonly",width=1)
        self.pending_label=ttk.Label(f,text="Pending: — TON")
        self.vars={k:tk.StringVar() for k,_ in FIELDS}; self.vars["lr_date"].set(datetime.now().strftime("%d/%m/%Y")); self.vars["builty_company"].set("MANSI COAL CARRIER"); self.vars["builty_address"].set("At Post Kunriya, Tal. Bhuj-Kutch. Corr Add Office No b/h sadguru complex madhapar highway"); self.set_fixed_parties()
        self.selected_id=None
        # Scrollable form: keep every B167 field reachable even on shorter screens.
        form_host=ttk.Frame(left); form_host.pack(fill="both",expand=True)
        form_canvas=tk.Canvas(form_host,highlightthickness=0,borderwidth=0)
        form_scroll=ttk.Scrollbar(form_host,orient="vertical",command=form_canvas.yview)
        form_canvas.configure(yscrollcommand=form_scroll.set)
        form_scroll.pack(side="right",fill="y")
        form_canvas.pack(side="left",fill="both",expand=True)
        form_inner=ttk.Frame(form_canvas)
        form_window=form_canvas.create_window((0,0),window=form_inner,anchor="nw")
        form_inner.bind("<Configure>",lambda _e:form_canvas.configure(scrollregion=form_canvas.bbox("all")))
        form_canvas.bind("<Configure>",lambda e:form_canvas.itemconfigure(form_window,width=e.width))
        ttk.Label(form_inner,text="Transport Details",font=("Segoe UI",10,"bold")).pack(anchor="w",pady=(2,4))
        grid=ttk.Frame(form_inner); grid.pack(fill="both",expand=True)
        for i,(key,label) in enumerate(FIELDS):
            rr,cc=divmod(i,2); ttk.Label(grid,text=label).grid(row=rr,column=cc*2,sticky="w",padx=3,pady=3)
            ttk.Entry(grid,textvariable=self.vars[key],width=25).grid(row=rr,column=cc*2+1,sticky="ew",padx=3,pady=3)
        for cc in range(4): grid.columnconfigure(cc,weight=1)
        bar2=ttk.Frame(left); bar2.pack(fill="x",pady=6)
        ttk.Button(bar2,text="+ Add",command=self.clear).pack(side="left",padx=2)
        ttk.Button(bar2,text="Edit Selected",command=self.edit_selected).pack(side="left",padx=2)
        ttk.Button(bar2,text="◀ Previous",command=lambda:self.navigate(-1)).pack(side="left",padx=2)
        ttk.Button(bar2,text="Next ▶",command=lambda:self.navigate(1)).pack(side="left",padx=2)
        ttk.Button(bar2,text="Save Builty",command=self.save).pack(side="left",padx=2)
        ttk.Button(bar2,text="Print Preview",command=self.print_pdf).pack(side="left",padx=2)
        ttk.Button(bar2,text="Chit Print",command=self.print_pdf).pack(side="left",padx=2)
        ttk.Button(bar2,text="Clear",command=self.clear).pack(side="left",padx=2)

        self.tree=ttk.Treeview(right,columns=("id","lr","date","party","truck","weight","amount"),show="headings",height=15)
        for key,label,width in [("id","ID",35),("lr","Builty No.",85),("date","Date",78),("party","Party Name",140),("truck","Truck No.",90),("weight","Weight",75),("amount","Amount",80)]:
            self.tree.heading(key,text=label); self.tree.column(key,width=width,anchor="w" if key=="party" else "center")
        self.tree.pack(fill="both",expand=True,pady=5); self.tree.bind("<<TreeviewSelect>>",self.load)
        self.refresh()
        self.refresh_do_options()
        self.refresh_pending_po_list()

    def refresh_all(self):
        self.refresh_do_options()
        self.refresh_pending_po_list()
        self.refresh()

    def refresh_pending_po_list(self):
        if not hasattr(self,"pending_tree"):
            return
        for item in self.pending_tree.get_children():
            self.pending_tree.delete(item)
        # A restored backup may contain stale remaining_qty values from an
        # older version. Recalculate from original PO quantity and saved
        # Builty transactions instead of trusting that cached column.
        with db() as con:
            rows=con.execute("""SELECT d.id,d.po_no,d.po_date,d.do_no,d.gmdc_do_no,d.do_date,
                                      d.customer_no,d.customer_name,d.customer_gstin,d.order_type,
                                      d.bill_to,d.ship_to,d.state,d.hsn,d.material,d.schedule_date,
                                      d.qty,d.remaining_qty,d.uom,d.unit_price,d.extended_price,
                                      d.transporter_code,d.transporter_name,d.order_total,
                                      d.from_place,d.to_place
                               FROM dos d WHERE COALESCE(d.qty,0)>0
                               ORDER BY d.id DESC""").fetchall()
        self.do_rows={}
        for row in rows:
            pending=self._do_pending(int(row["id"]))
            with db() as con:
                con.execute("UPDATE dos SET remaining_qty=? WHERE id=?",(pending,int(row["id"])))
            if pending<=0.000001:
                continue
            po_no=row["po_no"] or row["do_no"] or row["gmdc_do_no"] or ""
            label=f"#{row['id']} | PO {po_no} | {row['customer_name'] or ''}"
            data=dict(row)
            data["remaining_qty"]=pending
            self.do_rows[label]=data
            self.pending_tree.insert("","end",values=(
                row["id"],po_no,row["customer_name"] or "",
                row["material"] or "",
                f"{pending:.3f} {row['uom'] or 'TON'}"))
        # Keep the full PO lookup separate from the pending-only list so a
        # restored Builty can still resolve its original linked PO when edited.
        self.refresh_do_options()

    def on_pending_po_select(self,_event=None):
        selection=self.pending_tree.selection()
        if not selection:
            return
        values=self.pending_tree.item(selection[0],"values")
        ident=str(values[0])
        selected_row=None
        for label,row in getattr(self,"do_rows",{}).items():
            if str(row["id"])==ident:
                self.do_var.set(label)
                selected_row=row
                break
        if selected_row is None:
            self.refresh_do_options()
            for label,row in getattr(self,"do_rows",{}).items():
                if str(row["id"])==ident:
                    self.do_var.set(label)
                    selected_row=row
                    break
        self.selected_id=None
        if selected_row:
            self.autofill_from_po(selected_row)
        self.update_pending_label()
        self.update_po_details()
        if getattr(self,"builty_filter","all")=="selected":
            self.refresh()

    def set_fixed_parties(self):
        """Apply the same approved consignor and consignee details on every Builty."""
        fixed={
            "consignor_name":FIXED_CONSIGNOR_NAME,
            "consignor_address":FIXED_CONSIGNOR_ADDRESS,
            "consignor_gstin":FIXED_CONSIGNOR_GSTIN,
            "consignee_name":FIXED_CONSIGNEE_NAME,
            "consignee_address":FIXED_CONSIGNEE_ADDRESS,
            "consignee_gstin":FIXED_CONSIGNEE_GSTIN,
        }
        for key,value in fixed.items():
            if key in self.vars:
                self.vars[key].set(value)

    def autofill_from_po(self, po):
        """Copy the saved PO master data into the corresponding B167 fields."""
        mapping={
            "po_no":"po_no",
            "po_date":"po_date",
            "po_buyer":"customer_name",
            "po_vendor":"transporter_name",
            "po_material":"material",
            "po_qty":"qty",
            "po_uom":"uom",
            "po_rate":"unit_price",
            "po_delivery_date":"schedule_date",
            "consignee_name":"customer_name",
            "consignee_address":"ship_to",
            "consignee_gstin":"customer_gstin",
            "consignor_name":"transporter_name",
            "consignor_address":"bill_to",
        }
        for target,source in mapping.items():
            if target in self.vars:
                value=po.get(source)
                if value is not None and str(value).strip():
                    self.vars[target].set(str(value).strip())
        if "builty_qty" in self.vars and not self.vars["builty_qty"].get().strip():
            remaining=self._do_pending(int(po["id"]),self.selected_id)
            self.vars["builty_qty"].set(f"{remaining:g}" if remaining>0 else "")
        # These two parties are constant for every Rajasthan B167 Builty,
        # regardless of the selected PO's billing / shipping addresses.
        self.set_fixed_parties()
        self.update_po_details()

    def update_po_details(self):
        po=self._selected_do()
        if not hasattr(self,"po_details_label"):
            return
        if not po:
            self.po_details_label.configure(text="Select a pending PO to auto-fill party, GSTIN, material and available quantity.")
            return
        pending=self._do_pending(int(po["id"]),self.selected_id)
        self.po_details_label.configure(text=(
            f"PO No.: {po.get('do_no') or po.get('gmdc_do_no') or ''}    "
            f"Party: {po.get('customer_name') or ''}\n"
            f"Material: {po.get('material') or ''}    "
            f"Available Qty: {pending:g} {po.get('uom') or 'TON'}"
        ))

    def edit_selected(self):
        if not self.tree.selection():
            messagebox.showwarning("Edit Builty","Select a saved Builty from the right-side list.",parent=self.frame)
            return
        self.load()

    def navigate(self,step):
        rows=self.tree.get_children()
        if not rows:
            return
        selected=self.tree.selection()
        index=rows.index(selected[0]) if selected and selected[0] in rows else (0 if step>0 else len(rows)-1)
        index=max(0,min(len(rows)-1,index+step))
        self.tree.selection_set(rows[index])
        self.tree.focus(rows[index])
        self.tree.see(rows[index])
        self.load()

    def refresh_do_options(self):
        try:
            with db() as con:
                rows=con.execute("""SELECT id,po_no,po_date,do_no,gmdc_do_no,do_date,customer_no,
                    customer_name,customer_gstin,order_type,bill_to,ship_to,state,hsn,material,
                    schedule_date,qty,remaining_qty,uom,unit_price,extended_price,transporter_code,
                    transporter_name,order_total,from_place,to_place FROM dos ORDER BY id DESC""").fetchall()
                self.do_rows={}
                labels=[]
                for r in rows:
                    no=r["po_no"] or r["do_no"] or r["gmdc_do_no"] or r["id"]
                    label=f"#{r['id']} | PO {no} | {r['customer_name'] or ''}"
                    self.do_rows[label]=dict(r)
                    labels.append(label)
                self.do_combo.configure(values=labels)
                if self.do_var.get() not in self.do_rows:
                    self.do_var.set("")
            self.update_pending_label()
        except Exception as exc:
            messagebox.showerror("Load DOs",str(exc),parent=self.frame)

    def _selected_do(self):
        return getattr(self,"do_rows",{}).get(self.do_var.get())

    def _do_pending(self, do_id, exclude_id=None):
        with db() as con:
            row=con.execute("SELECT qty FROM dos WHERE id=?",(do_id,)).fetchone()
            if not row:return 0.0
            total=float(row["qty"] or 0)
            # When editing an existing B167 record, exclude its linked standard
            # Builty from the consumed total so the same quantity can be edited.
            current_standard_id=None
            if exclude_id is not None:
                current=con.execute("SELECT data_json FROM po_builtys WHERE id=?",(exclude_id,)).fetchone()
                if current:
                    try:
                        current_standard_id=int(json.loads(current["data_json"] or "{}").get("standard_builty_id") or 0) or None
                    except (TypeError,ValueError):
                        pass
            standard=con.execute(
                "SELECT COALESCE(SUM(qty),0) FROM builtys WHERE do_id=? AND (? IS NULL OR id<>?)",
                (do_id,current_standard_id,current_standard_id)
            ).fetchone()[0]
            # Older B167 records may predate standard Builty integration. Count
            # only those legacy rows that do not yet point to a standard builty.
            po_rows=con.execute("SELECT id,data_json FROM po_builtys").fetchall()
            legacy_total=0.0
            for item in po_rows:
                if exclude_id is not None and int(item["id"])==int(exclude_id):continue
                try:
                    payload=json.loads(item["data_json"] or "{}")
                    if (str(payload.get("linked_do_id",""))==str(do_id)
                            and not payload.get("standard_builty_id")):
                        legacy_total+=max(0.0,float(payload.get("builty_qty") or 0))
                except (TypeError,ValueError):
                    continue
            return max(0.0,total-float(standard or 0)-legacy_total)

    def update_pending_label(self):
        do=self._selected_do()
        if not do:
            self.pending_label.configure(text="Pending: — TON")
            return
        pending=self._do_pending(int(do["id"]),self.selected_id)
        self.pending_label.configure(text=f"Pending: {pending:g} TON")

    def clear(self):
        self.selected_id=None
        for v in self.vars.values(): v.set("")
        self.vars["lr_date"].set(datetime.now().strftime("%d/%m/%Y"))
        self.vars["builty_qty"].set("")
        self.do_var.set("")
        if hasattr(self,"pending_tree"):
            for item in self.pending_tree.selection():
                self.pending_tree.selection_remove(item)
        self.update_pending_label()
        self.vars["builty_company"].set("MANSI COAL CARRIER")
        self.vars["builty_address"].set("At Post Kunriya, Tal. Bhuj-Kutch. Corr Add Office No b/h sadguru complex madhapar highway")
        for x in self.tree.selection(): self.tree.selection_remove(x)

    def create_do(self):
        if not self.create_do_callback:
            messagebox.showerror("PO to DO", "DO workflow callback is not configured.", parent=self.frame)
            return
        data = {k: v.get().strip() for k, v in self.vars.items()}
        if not data.get("po_no"):
            messagebox.showwarning("PO to DO", "Import a PO PDF and confirm the PO No. first.", parent=self.frame)
            return
        self.create_do_callback(data)

    def import_pdf(self):
        path=filedialog.askopenfilename(title="Select PO PDF",filetypes=[("PDF files","*.pdf")])
        if not path: return
        try:
            from pypdf import PdfReader
            text="\n".join(page.extract_text() or "" for page in PdfReader(path).pages)
            if not text.strip():
                # OCR fallback for scanned/image-only supplier PDFs.
                try:
                    import fitz, pytesseract, io
                    from PIL import Image
                    pages=[]
                    pdf_doc=fitz.open(path)
                    for page in pdf_doc:
                        pix=page.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False)
                        pages.append(pytesseract.image_to_string(Image.open(io.BytesIO(pix.tobytes("png")))))
                    text="\n".join(pages)
                except Exception as ocr_error:
                    raise ValueError("Scanned PDF OCR is unavailable. Install PyMuPDF, Pillow, pytesseract and Tesseract OCR, or fill fields manually. "+str(ocr_error))
            if not text.strip(): raise ValueError("PDF text could not be extracted; fill fields manually.")
            patterns={
              "po_no":r"(?i)(?:PURCHASE\s*ORDER\s*(?:NO|NUMBER)|P\.?\s*O\.?\s*(?:NO|NUMBER))\s*[:#-]?\s*([A-Z0-9/-]{5,})",
              "po_date":r"(?i)(?:PURCHASE\s*ORDER\s*DATE|P\.?\s*O\.?\s*DATE)\s*[:#-]?\s*([0-9]{1,2}[./-][0-9]{1,2}[./-][0-9]{2,4})",
              "po_delivery_date":r"(?i)(?:DELIVERY\s*DATE|DELIVERY\s*BY|SCHEDULED\s*DELIVERY)\s*[:#-]?\s*([0-9]{1,2}[./-][0-9]{1,2}[./-][0-9]{2,4})",
              "po_qty":r"(?i)(?:ORDER\s*QUANTITY|QUANTITY|QTY)\s*[:#-]?\s*([0-9,]+(?:\.[0-9]+)?)",
              "po_rate":r"(?i)(?:UNIT\s*PRICE|RATE(?:\s*/\s*(?:KG|MT|TON))?)\s*[:#-]?\s*(?:INR|RS\.?|₹)?\s*([0-9,]+(?:\.[0-9]+)?)",
              "lr_no":r"\bLR\s*NO\.?\s*[:#-]\s*([A-Z0-9/-]+)",
              "lr_date":r"\bLR\s*DATE\s*[:#-]\s*([0-9./-]+)",
              "invoice_no":r"INVOI[CS]E\s*NO\.?\s*[:#-]\s*([^\n]+)",
              "challan_no":r"CHALAN\s*NO\.?\s*[:#-]\s*([^\n]+)",
              "eway_bill_no":r"E[\s-]*WAY\s*BILL\s*NO\.?\s*[:#-]\s*([^\n]+)",
              "eway_generate_date":r"E[\s-]*WAY\s*BILL\s*GENERATE\s*DATE\s*[:#-]?\s*([^\n]+)",
              "eway_valid_date":r"E[\s-]*WAY\s*BILL\s*VALID\s*DATE\s*[:#-]?\s*([^\n]+)",
              "packing_list_no":r"PACKING\s*LIST\s*NO\.?\s*[:#-]\s*([^\n]+)",
              "consignor_gstin":r"CONSIGNOR[\s\S]{0,500}?GSTIN\s*[:#-]?\s*([0-9A-Z]{15})",
              "consignee_gstin":r"CONSIGNEE[\s\S]{0,3000}?GSTIN\s*(?:NO\.?\s*|NO/UNIQUE\s*NO\s*)?[:#-]?\s*([0-9A-Z]{15}|URP)",
              "truck_no":r"TRUCK\s*NO\.?\s*[:#-]\s*([A-Z0-9 -]+)",
              "driver_name":r"DRIVER'?S?\s*NAME\s*[:#-]\s*([^\n]+)",
              "truck_owner":r"TRUCK\s*OWNER\s*NAME\s*[:#-]\s*([^\n]+)",
              "mobile_no":r"MOBILE\s*NO\.?\s*[:#-]\s*([0-9 +()-]+)"}
            n=0
            for key,pat in patterns.items():
                m=re.search(pat,text,re.I)
                if m: self.vars[key].set(re.sub(r"\s+"," ",m.group(1)).strip(" :\t")[:500]); n+=1
            for key,label in [("consignor_name","CONSIGNOR"),("consignee_name","CONSIGNEE")]:
                m=re.search(label+r"\s*:?\s*\n+\s*([^\n]+)",text,re.I)
                if m:
                    value=m.group(1).strip()
                    if value.upper() not in ("GSTIN","GSTIN NO/UNIQUE NO"):
                        self.vars[key].set(value[:500]); n+=1
            # Tabular PO layouts often show the PO number without a field label.
            if not self.vars["po_no"].get():
                m=re.search(r"(?i)\b(4000[0-9]{6})\b", text)
                if m: self.vars["po_no"].set(m.group(1)); n+=1
            m=re.search(r"(?i)(?:VENDOR|SUPPLIER)(?:\s+NAME)?\s*[:#-]?\s*([^\n]+)",text)
            if m: self.vars["po_vendor"].set(m.group(1).strip()[:500]); n+=1
            m=re.search(r"(?i)(?:BUYER|PURCHASER)(?:\s+NAME)?\s*[:#-]?\s*([^\n]+)",text)
            if m: self.vars["po_buyer"].set(m.group(1).strip()[:500]); n+=1
            # Supplier PO template may print company names without explicit Buyer/Vendor labels.
            if re.search(r"ARVIND\s+LIMITED", text, re.I):
                self.vars["po_buyer"].set("ARVIND LIMITED"); n+=1
            if re.search(r"NILAV\s+MINETECH", text, re.I):
                self.vars["po_vendor"].set("NILAV MINETECH"); n+=1
            m=re.search(r"(?i)(?:MATERIAL|ITEM\s*DESCRIPTION|DESCRIPTION)\s*[:#-]?\s*([^\n]+)",text)
            if m: self.vars["po_material"].set(m.group(1).strip()[:500]); n+=1
            m=re.search(r"(?i)\b(KG|KGS|MT|TON|TONNE|TONNES)\b",text)
            if m: self.vars["po_uom"].set(m.group(1).upper()); n+=1
            if re.search(r"KARAMTARA ENGINEERING LTD",text,re.I):
                self.vars["consignor_name"].set("KARAMTARA ENGINEERING LTD"); n+=1
            m=re.search(r"KARAMTARA ENGINEERING LTD\s*\n+([^\n]+)",text,re.I)
            if m: self.vars["consignor_address"].set(m.group(1).strip()[:500]); n+=1
            if re.search(r"KARAMTARA USA INC",text,re.I):
                self.vars["consignee_name"].set("KARAMTARA USA INC"); n+=1
            m=re.search(r"KARAMTARA USA INC\s*\n+([^\n]+)",text,re.I)
            if m: self.vars["consignee_address"].set(m.group(1).strip()[:500]); n+=1
            m=re.search(r"(?m)^\s*([0-9]{1,4})\s+([0-9]+(?:\.[0-9]+)?)\s*$",text)
            if m: self.vars["total_bundle"].set(m.group(1)); self.vars["weight_1"].set(m.group(2)); n+=2
            m=re.search(r"(?m)^\s*(Int_[A-Za-z0-9_.' -]{3,70}|[A-Za-z0-9_.' -]*MILS[A-Za-z0-9_.' -]*)\s*$",text,re.I)
            if m: self.vars["description"].set(m.group(1).strip()); n+=1
            messagebox.showinfo("Import PO PDF",f"Read {n} field(s). Please review fields and fill anything missing.",parent=self.frame)
        except Exception as e: messagebox.showerror("Import PO PDF",str(e))

    def _sync_do_balance(self, do_id):
        with db() as con:
            standard=con.execute("SELECT COALESCE(SUM(qty),0) FROM builtys WHERE do_id=?",(do_id,)).fetchone()[0]
            rows=con.execute("SELECT data_json FROM po_builtys").fetchall()
            legacy_total=0.0
            for item in rows:
                try:
                    payload=json.loads(item["data_json"] or "{}")
                    if str(payload.get("linked_do_id",""))==str(do_id) and not payload.get("standard_builty_id"):
                        legacy_total+=max(0.0,float(payload.get("builty_qty") or 0))
                except (TypeError,ValueError):
                    pass
            con.execute("UPDATE dos SET remaining_qty=MAX(0,COALESCE(qty,0)-?-?) WHERE id=?",(float(standard or 0),legacy_total,do_id))

    def save(self):
        d={k:v.get().strip() for k,v in self.vars.items()}
        if not d["lr_no"]:
            messagebox.showwarning("Save Builty","LR No. is required.",parent=self.frame)
            return False
        po=self._selected_do()
        if not po:
            messagebox.showwarning("Save Builty","Select the linked PO first.",parent=self.frame)
            return False
        try:
            qty=float((d.get("builty_qty") or "").replace(",",""))
            if qty<=0: raise ValueError
        except (TypeError,ValueError):
            messagebox.showwarning("Save Builty","Enter a valid Builty Qty greater than zero.",parent=self.frame)
            return False

        old_do_id=None
        standard_id=None
        old_billed=False
        if self.selected_id:
            with db() as con:
                old=con.execute("SELECT data_json FROM po_builtys WHERE id=?",(self.selected_id,)).fetchone()
                if old:
                    old_data=json.loads(old["data_json"] or "{}")
                    old_do_id=int(old_data.get("linked_do_id") or 0) or None
                    standard_id=int(old_data.get("standard_builty_id") or 0) or None
                    if standard_id:
                        linked=con.execute("SELECT billed,freight_paid FROM builtys WHERE id=?",(standard_id,)).fetchone()
                        if linked and (linked["billed"] or linked["freight_paid"]):
                            messagebox.showwarning("Builty Locked","This Builty is already billed or freight-paid; edit it from the linked transaction first.",parent=self.frame)
                            return False

        pending=self._do_pending(int(po["id"]),self.selected_id)
        if qty>pending+1e-6:
            messagebox.showwarning("PO Quantity",f"Builty quantity exceeds PO pending quantity.\nPending: {pending:.3f}\nEntered: {qty:.3f}",parent=self.frame)
            return False

        d["linked_do_id"]=str(po["id"])
        try:
            with db() as con:
                # The normal Bill and Freight screens read the standard builtys table.
                # Persist the B167 record there, and keep the full B167 payload for printing.
                existing_no=con.execute("SELECT id FROM builtys WHERE builty_no=?",(d["lr_no"],)).fetchone()
                if existing_no and (not standard_id or int(existing_no["id"])!=standard_id):
                    raise ValueError("આ LR No. પહેલેથી Builty No. તરીકે વપરાયેલ છે.")
                if standard_id:
                    con.execute("""UPDATE builtys SET builty_no=?,do_id=?,truck_no=?,date=?,qty=?,
                        from_place=?,to_place=?,consignor=?,consignee=?,driver_name=?,ewaybill=?,
                        vehicle_no=?,despatch_weight=?,receipt_weight=?,owner_name=?,mobile_no=?,
                        address=?,challan_no=? WHERE id=?""",
                        (d["lr_no"],int(po["id"]),d.get("truck_no",""),d.get("lr_date",""),qty,
                         d.get("consignor_address",""),d.get("consignee_address",""),
                         d.get("consignor_name",""),d.get("consignee_name",""),d.get("driver_name",""),
                         d.get("eway_bill_no",""),d.get("truck_no",""),qty,qty,d.get("truck_owner",""),
                         d.get("mobile_no",""),d.get("driver_address",""),d.get("challan_no",""),standard_id))
                else:
                    cur=con.execute("""INSERT INTO builtys
                        (builty_no,do_id,truck_no,date,qty,from_place,to_place,consignor,consignee,
                         driver_name,ewaybill,vehicle_no,despatch_weight,receipt_weight,owner_name,
                         mobile_no,address,challan_no)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (d["lr_no"],int(po["id"]),d.get("truck_no",""),d.get("lr_date",""),qty,
                         d.get("consignor_address",""),d.get("consignee_address",""),
                         d.get("consignor_name",""),d.get("consignee_name",""),d.get("driver_name",""),
                         d.get("eway_bill_no",""),d.get("truck_no",""),qty,qty,d.get("truck_owner",""),
                         d.get("mobile_no",""),d.get("driver_address",""),d.get("challan_no","")))
                    standard_id=cur.lastrowid
                d["standard_builty_id"]=str(standard_id)
                if self.selected_id:
                    con.execute("UPDATE po_builtys SET lr_no=?,lr_date=?,data_json=?,updated_at=datetime('now','localtime') WHERE id=?",
                                (d["lr_no"],d["lr_date"],json.dumps(d,ensure_ascii=False),self.selected_id))
                else:
                    cur=con.execute("INSERT INTO po_builtys(lr_no,lr_date,data_json) VALUES(?,?,?)",
                                    (d["lr_no"],d["lr_date"],json.dumps(d,ensure_ascii=False)))
                    self.selected_id=cur.lastrowid
            self._sync_do_balance(int(po["id"]))
            if old_do_id and old_do_id!=int(po["id"]): self._sync_do_balance(old_do_id)
            self.refresh(); self.refresh_do_options()
            return True
        except Exception as e:
            messagebox.showerror("Save Builty",str(e),parent=self.frame)
            return False

    def refresh(self):
        if not self.tree.winfo_exists():
            return
        for x in self.tree.get_children(): self.tree.delete(x)
        with db() as con:
            rows=con.execute("SELECT id,lr_no,lr_date,data_json FROM po_builtys ORDER BY id DESC").fetchall()
        selected_po=self._selected_do()
        query=self.search_var.get().strip().casefold() if hasattr(self,"search_var") else ""
        for r in rows:
            d=json.loads(r["data_json"] or "{}")
            linked=str(d.get("linked_do_id",""))
            if getattr(self,"builty_filter","all")=="selected" and (not selected_po or linked!=str(selected_po["id"])):
                continue
            searchable=" ".join(str(v or "") for v in (r["id"],r["lr_no"],r["lr_date"],d.get("consignee_name",""),d.get("truck_no",""),d.get("linked_do_id",""))).casefold()
            if query and query not in searchable:
                continue
            self.tree.insert("","end",values=(r["id"],r["lr_no"],r["lr_date"],d.get("consignee_name",""),d.get("truck_no",""),d.get("builty_qty",""),d.get("total","")))
        self.refresh_pending_po_list()

    def show_selected_po_builtys(self):
        if not self._selected_do():
            messagebox.showinfo("Builty List","Select a pending PO first.",parent=self.frame)
            return
        self.builty_filter="selected"
        self.refresh()

    def show_all_builtys(self):
        self.builty_filter="all"
        self.refresh()

    def load(self,_event=None):
        s=self.tree.selection()
        if not s:return
        self.selected_id=int(self.tree.item(s[0])["values"][0])
        with db() as con:r=con.execute("SELECT data_json FROM po_builtys WHERE id=?",(self.selected_id,)).fetchone()
        if r:
            d=json.loads(r["data_json"] or "{}")
            for k,v in self.vars.items():v.set(str(d.get(k,"")))
            linked_id=str(d.get("linked_do_id",""))
            for label,row in getattr(self,"do_rows",{}).items():
                if str(row["id"])==linked_id:
                    self.do_var.set(label)
                    break
            self.update_pending_label()

    def delete(self):
        s=self.tree.selection()
        if not s: messagebox.showwarning("Delete PO Builty","Select a record first."); return
        ident=int(self.tree.item(s[0])["values"][0])
        if not messagebox.askyesno("Delete PO Builty","Delete this PO Builty?"):return
        with db() as con:
            row=con.execute("SELECT data_json FROM po_builtys WHERE id=?",(ident,)).fetchone()
            linked_id=None
            standard_id=None
            if row:
                try:
                    payload=json.loads(row["data_json"] or "{}")
                    linked_id=int(payload.get("linked_do_id") or 0) or None
                    standard_id=int(payload.get("standard_builty_id") or 0) or None
                except (TypeError,ValueError):
                    pass
            if standard_id:
                linked=con.execute("SELECT billed,freight_paid FROM builtys WHERE id=?",(standard_id,)).fetchone()
                if linked and (linked["billed"] or linked["freight_paid"]):
                    messagebox.showwarning("Delete Builty","This Builty is already billed or freight-paid and cannot be deleted.",parent=self.frame)
                    return
                con.execute("DELETE FROM builtys WHERE id=?",(standard_id,))
            con.execute("DELETE FROM po_builtys WHERE id=?",(ident,))
        if linked_id:
            self._sync_do_balance(linked_id)
        self.clear()
        self.refresh()
        self.refresh_do_options()

    def print_pdf(self):
        d={k:v.get().strip() for k,v in self.vars.items()}
        if not d["lr_no"]:messagebox.showwarning("Print PDF","Enter/import LR No. first.");return
        if not self.save():return
        from po_builty import generate_po_builty_pdf
        # Keep the original saved timestamp; previewing/printing later must not change it.
        with db() as con:
            saved=con.execute("SELECT created_at FROM po_builtys WHERE id=?",(self.selected_id,)).fetchone()
        d["created_at"] = saved["created_at"] if saved else ""
        out=self.export_dir/("PO_Builty_"+re.sub(r"[^A-Za-z0-9_-]","_",d["lr_no"])+".pdf")
        try:generate_po_builty_pdf(out,self.firm_getter(),d)
        except Exception as e:messagebox.showerror("PO Builty PDF",str(e));return
        if messagebox.askyesno("PO Builty PDF",f"Saved {out.name}. Open to print?"):self.open_file(out)


def generate_po_builty_pdf(out_path,firm,data):
    """Create a clear, full-width B167 LR print matching the supplied reference."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from xml.sax.saxutils import escape
    from datetime import datetime

    styles=getSampleStyleSheet()
    normal=ParagraphStyle("B167Normal",parent=styles["Normal"],fontName="Helvetica",fontSize=8,leading=9.5,alignment=1,spaceAfter=0)
    bold=ParagraphStyle("B167Bold",parent=normal,fontName="Helvetica-Bold")
    title=ParagraphStyle("B167Title",parent=bold,fontSize=18,leading=20)
    subtitle=ParagraphStyle("B167Subtitle",parent=bold,fontSize=10,leading=12)
    left=ParagraphStyle("B167Left",parent=normal,alignment=0)
    small=ParagraphStyle("B167Small",parent=normal,fontSize=7.2,leading=8.4)
    smallbold=ParagraphStyle("B167SmallBold",parent=small,fontName="Helvetica-Bold")
    centerbold=ParagraphStyle("B167CenterBold",parent=bold,fontSize=8.5,leading=10)

    def p(value="",style=normal):
        raw=str(value or "").replace("\\r","")
        return Paragraph("<br/>".join(escape(line) for line in raw.split("\\n")),style)
    def table(rows,widths,heights=None,pad=3):
        t=Table(rows,colWidths=widths,rowHeights=heights,hAlign="CENTER")
        t.setStyle(TableStyle([
            ("BOX",(0,0),(-1,-1),0.9,colors.black),
            ("INNERGRID",(0,0),(-1,-1),0.65,colors.black),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",(0,0),(-1,-1),pad),("RIGHTPADDING",(0,0),(-1,-1),pad),
            ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
        ]))
        return t
    def val(key,default=""):
        v=data.get(key)
        return str(v if v not in (None,"") else default)

    name=val("builty_company",firm.get("name") or "MANSI COAL CARRIER").upper()
    address=val("builty_address",firm.get("address") or "At Post Kunriya, Tal. Bhuj-Kutch. Corr Add Office No b/h sadguru complex madhapar highway")

    # Fixed party details requested by the user; always print these on Rajasthan B167.
    consignor=[
        ": CONSIGNOR :",
        "KARAMTARA ENGINEERING LTD",
        "15 NO. 112/P18P2, 113/22, 116/2, 118",
        "119/2, 119/1/P18P2 & P3, 120, VILL-MEGHPAR-KUNJISAR TAL-BHACHAU",
        "GSTIN: 24AABCK1921E125",
    ]
    consignee=[
        ": CONSIGNEE :",
        "KARAMTARA USA INC",
        "95 HILLSIDE PARK DR, NEW HYDE PARK",
        "NEW YORK.",
        "Pin Code : 11040",
        "State Code/State : /",
        "GSTIN No/Unique No : URP",
    ]

    doc=SimpleDocTemplate(str(out_path),pagesize=A4,leftMargin=5*mm,rightMargin=5*mm,topMargin=4*mm,bottomMargin=4*mm)
    W=200*mm
    story=[]
    head=Table([
        [p("|| JAY MURLIDHAR ||",smallbold)],
        [p(name,title)],
        [p("TRANSPORT CONTRACTOR & COMMISSION AGENT",subtitle)],
        [p(address,smallbold)],
    ],colWidths=[W],rowHeights=[5*mm,10*mm,7*mm,6*mm])
    head.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),1.0,colors.black),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))
    story.append(head)
    story.append(table([[
        p("LR NO : "+val("lr_no"),smallbold),
        p("LR DATE : "+val("lr_date"),smallbold)
    ]],[100*mm,100*mm],[6*mm]))

    left_col=Table([
        [[p(x,smallbold if i in (0,1) else small) for i,x in enumerate(consignor)]],
        [[p(x,smallbold if i in (0,1) else small) for i,x in enumerate(consignee)]],
    ],colWidths=[100*mm],rowHeights=[31*mm,35*mm])
    left_col.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),0.9,colors.black),("INNERGRID",(0,0),(-1,-1),0.65,colors.black),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
        ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
    ]))
    meta=[
        ("INVOICE NO", "invoice_no"),
        ("CHALAN NO", "challan_no"),
        ("E WAY BILL NO", "eway_bill_no"),
        ("E WAY BILL GENERATE DATE", "eway_generate_date"),
        ("E WAY BILL VALID DATE", "eway_valid_date"),
        ("Packing list no", "packing_list_no"),
    ]
    right_meta=Table([[p(label+" : "+val(key),smallbold)] for label,key in meta],
                     colWidths=[100*mm],rowHeights=[11*mm]*6)
    right_meta.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),0.9,colors.black),("INNERGRID",(0,0),(-1,-1),0.65,colors.black),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),4),
        ("RIGHTPADDING",(0,0),(-1,-1),4),("TOPPADDING",(0,0),(-1,-1),2),
        ("BOTTOMPADDING",(0,0),(-1,-1),2),
    ]))
    parties=Table([[left_col,right_meta]],colWidths=[100*mm,100*mm],rowHeights=[66*mm])
    parties.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.9,colors.black),("INNERGRID",(0,0),(-1,-1),0.65,colors.black),("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
    story.append(parties)

    desc=val("description",val("po_material"))
    weight1=val("weight_1")
    weight2=val("weight_2")
    try:
        total_weight=float(weight1 or 0)+float(weight2 or 0)
        total_weight_text=f"{total_weight:,.3f}" if (weight1 or weight2) else val("builty_qty")
    except (TypeError,ValueError):
        total_weight_text=" / ".join(x for x in (weight1,weight2) if x) or val("builty_qty")
    bundle1=val("bundle_1")
    bundle2=val("total_bundle")
    weight1=val("weight_1")
    weight2=val("weight_2")
    def add_numeric(a,b):
        try:
            total=float((a or "0").replace(",",""))+float((b or "0").replace(",",""))
            return f"{total:,.3f}".rstrip("0").rstrip(".")
        except (TypeError,ValueError,AttributeError):
            return ""
    bundle_total=add_numeric(bundle1,bundle2)
    weight_total=add_numeric(weight1,weight2)
    bundle_lines="\\n".join(x for x in (bundle1,bundle2) if x)
    weight_lines="\\n".join(x for x in (weight1,weight2) if x)
    goods=table([
        [p("NO",centerbold),p("DESCRIPTION",centerbold),p("TOTAL\\nBUNDLE",centerbold),p("TOTAL\\nWEIGHT",centerbold)],
        [p("1",centerbold),p(desc,centerbold),p(bundle_lines,normal),p(weight_lines,normal)],
        [p("SUB TOTAL",centerbold),"",p(bundle_total,bold),p(weight_total,bold)],
    ],[15*mm,103*mm,32*mm,50*mm],[10*mm,32*mm,7*mm])
    story.append(goods)

    truck_text=Table([
        [p("Truck No :",smallbold),p(val("truck_no"),smallbold)],
        [p("Drivers Name",smallbold),p(val("driver_name"),small)],
        [p("Address",smallbold),p(val("driver_address"),small)],
        [p("Truck Owner Name",smallbold),p(val("truck_owner"),small)],
        [p("Mobile No",smallbold),p(val("mobile_no"),small)],
    ],colWidths=[35*mm,75*mm],rowHeights=[7*mm]*5)
    truck_text.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),0.9,colors.black),("INNERGRID",(0,0),(-1,-1),0.65,colors.black),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),4),
        ("RIGHTPADDING",(0,0),(-1,-1),4),("TOPPADDING",(0,0),(-1,-1),1),("BOTTOMPADDING",(0,0),(-1,-1),1),
    ]))
    tax=Table([
        [p("CGST "+val("cgst","0")+"%",smallbold)],
        [p("SGST "+val("sgst","0")+"%",smallbold)],
        [p("IGST "+val("igst","0")+"%",smallbold)],
        [p("TOTAL : "+val("total"),smallbold)],
    ],colWidths=[90*mm],rowHeights=[8.75*mm]*4)
    tax.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),0.9,colors.black),("INNERGRID",(0,0),(-1,-1),0.65,colors.black),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),4),
    ]))
    truck_and_tax=Table([[truck_text,tax]],colWidths=[110*mm,90*mm],rowHeights=[35*mm])
    truck_and_tax.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),0.9,colors.black),("INNERGRID",(0,0),(-1,-1),0.65,colors.black),
        ("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))
    story.append(truck_and_tax)

    # Footer sized and divided like the user's two reference images.
    notice_style=ParagraphStyle("B167Notice",parent=left,fontName="Helvetica-Bold",fontSize=11.5,leading=15)
    notices=Table([
        [p("* 1% SHORTAGE ALLOWED. WE ARE NOT RESPONSIBLE FOR\\nQUALITY MATERIAL DIRECT LOADING FROM COMPANY",notice_style)],
        [p("* GSTIN NO 24AHOPC9866DIZ5",notice_style)],
        [p("* SUBJECT TO BHUJ JURISDICTION",notice_style)],
    ],colWidths=[122*mm],rowHeights=[23*mm,12*mm,12*mm])
    notices.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),3*mm),("RIGHTPADDING",(0,0),(-1,-1),2*mm),
        ("TOPPADDING",(0,0),(-1,-1),1*mm),("BOTTOMPADDING",(0,0),(-1,-1),1*mm),
    ]))
    created_raw=str(data.get("created_at") or "").strip()
    created_at=None
    for fmt in ("%Y-%m-%d %H:%M:%S","%Y-%m-%dT%H:%M:%S"):
        try:
            created_at=datetime.strptime(created_raw[:19],fmt)
            break
        except ValueError:
            pass
    if created_at is None:
        created_at=datetime.now()
    sign_date=created_at.strftime("%Y.%m.%d")
    sign_time=created_at.strftime("%H:%M:%S")
    sign_name=ParagraphStyle("B167SignName",parent=centerbold,fontSize=13,leading=16)
    sign_meta=ParagraphStyle("B167SignMeta",parent=small,fontSize=7.5,leading=9,alignment=0)
    sign_for=ParagraphStyle("B167SignFor",parent=centerbold,fontSize=11,leading=14)
    sign_top=Table([[
        [p("MAVJI NARAN",sign_name),p("CHAD",sign_name)],
        [p("Digitally signed by",sign_meta),p("MAVJI NARAN CHAD",ParagraphStyle("B167SignBold",parent=sign_meta,fontName="Helvetica-Bold")),p("Date: "+sign_date,sign_meta),p(sign_time+" +05'30'",sign_meta)]
    ]],colWidths=[39*mm,39*mm],rowHeights=[35*mm])
    sign_top.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),1.8,colors.black),("INNERGRID",(0,0),(-1,-1),1.2,colors.black),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(0,0),(0,0),"CENTER"),
        ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
        ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2),
    ]))
    sign_bottom=Table([[p("FOR : MANSI COAL CAREER",sign_for)]],colWidths=[78*mm],rowHeights=[12*mm])
    sign_bottom.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),1.8,colors.black),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),("LEFTPADDING",(0,0),(-1,-1),2),
        ("RIGHTPADDING",(0,0),(-1,-1),2),("TOPPADDING",(0,0),(-1,-1),1),
        ("BOTTOMPADDING",(0,0),(-1,-1),1),
    ]))
    right_panel=Table([[sign_top],[sign_bottom]],colWidths=[78*mm],rowHeights=[35*mm,12*mm])
    right_panel.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),0),
        ("RIGHTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),
        ("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))
    footer=Table([[notices,right_panel]],colWidths=[122*mm,78*mm],rowHeights=[47*mm])
    footer.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),1.8,colors.black),("INNERGRID",(0,0),(-1,-1),1.5,colors.black),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
        ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))
    story.append(footer)
    doc.build(story)
    return str(out_path)
