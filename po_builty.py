"""PO Builty form and B167-style PDF, linked to DO quantities in Rajasthan GMDC mode."""
import json, re, tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from db import db, get_active_company

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
("description","Description"),("total_bundle","Total Bundle"),("weight_1","Weight 1"),
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
        ttk.Label(f,text="Builty Entry — B167 Print Format",font=("Segoe UI",17,"bold"),foreground="#B22222").pack(anchor="w")
        ttk.Label(f,text="Select a pending PO above, enter B167 details, then Save / Print. Saved weight reduces PO pending quantity.",wraplength=1000).pack(anchor="w",pady=4)
        # GMDC-style top action bar: controls stay visible without scrolling to the bottom.
        actions=ttk.Frame(f); actions.pack(fill="x",pady=(2,5))
        ttk.Button(actions,text="Save Builty",command=self.save).pack(side="left",padx=2)
        ttk.Button(actions,text="Edit Selected",command=self.edit_selected).pack(side="left",padx=2)
        ttk.Button(actions,text="Delete Selected",command=self.delete).pack(side="left",padx=2)
        ttk.Button(actions,text="Previous",command=lambda:self.navigate(-1)).pack(side="left",padx=2)
        ttk.Button(actions,text="Next",command=lambda:self.navigate(1)).pack(side="left",padx=2)
        ttk.Button(actions,text="New / Cancel",command=self.clear).pack(side="left",padx=2)
        ttk.Button(actions,text="Print Builty",command=self.print_pdf).pack(side="left",padx=2)
        ttk.Button(actions,text="Refresh",command=self.refresh_all).pack(side="right",padx=2)

        # Upper pending-PO list, matching the GMDC DO selector workflow.
        pending_box=ttk.LabelFrame(f,text="Pending PO List — select a PO",padding=4)
        pending_box.pack(fill="x",pady=(0,5))
        pcols=("id","po","date","customer","material","qty","remaining")
        self.pending_tree=ttk.Treeview(pending_box,columns=pcols,show="headings",height=4,selectmode="browse")
        for key,label,width in (("id","ID",42),("po","PO No.",125),("date","PO Date",90),("customer","Customer",190),("material","Material",180),("qty","PO Qty",85),("remaining","Pending",85)):
            self.pending_tree.heading(key,text=label); self.pending_tree.column(key,width=width,anchor="w" if key in ("customer","material") else "center")
        self.pending_tree.pack(fill="x",expand=False)
        self.pending_tree.bind("<<TreeviewSelect>>",self.on_pending_po_select)

        # Keep the linked-order variables internally for existing save/edit logic,
        # but do not display the redundant Linked PO / Refresh POs / Pending TON row.
        self.do_var=tk.StringVar()
        self.do_combo=ttk.Combobox(f,textvariable=self.do_var,state="readonly",width=1)
        self.pending_label=ttk.Label(f,text="Pending: — TON")
        split=ttk.Panedwindow(f,orient="horizontal"); split.pack(fill="both",expand=True)
        left=ttk.Frame(split,padding=4); right=ttk.Frame(split,padding=4)
        split.add(left,weight=3); split.add(right,weight=2)
        bar=ttk.Frame(left); bar.pack(fill="x")
        ttk.Button(bar,text="Import Builty PDF",command=self.import_pdf).pack(side="left",padx=3)
        ttk.Button(bar,text="New / Clear",command=self.clear).pack(side="left",padx=3)
        self.vars={k:tk.StringVar() for k,_ in FIELDS}; self.vars["lr_date"].set(datetime.now().strftime("%d/%m/%Y")); self.vars["builty_company"].set("MANSI COAL CARRIER"); self.vars["builty_address"].set("At Post Kunriya, Tal. Bhuj-Kutch. Corr Add Office No b/h sadguru complex madhapar highway")
        self.selected_id=None
        grid=ttk.Frame(left); grid.pack(fill="both",expand=True)
        for i,(key,label) in enumerate(FIELDS):
            r,c=divmod(i,2); ttk.Label(grid,text=label).grid(row=r,column=c*2,sticky="w",padx=3,pady=3)
            ttk.Entry(grid,textvariable=self.vars[key],width=28).grid(row=r,column=c*2+1,sticky="ew",padx=3,pady=3)
        for c in range(4): grid.columnconfigure(c,weight=1)
        bar2=ttk.Frame(left); bar2.pack(fill="x",pady=6)
        ttk.Button(bar2,text="Save Builty",command=self.save).pack(side="left",padx=3)
        ttk.Button(bar2,text="Print / Preview PDF",command=self.print_pdf).pack(side="left",padx=3)
        ttk.Button(bar2,text="Delete Selected",command=self.delete).pack(side="left",padx=3)
        ttk.Label(right,text="Saved Builty List",font=("Segoe UI",11,"bold")).pack(anchor="w")
        self.tree=ttk.Treeview(right,columns=("id","lr","date","party","truck"),show="headings",height=20)
        for c,l,w in [("id","ID",40),("lr","LR No.",80),("date","LR Date",90),("party","Consignee",160),("truck","Truck No.",95)]:
            self.tree.heading(c,text=l); self.tree.column(c,width=w,anchor="w" if c=="party" else "center")
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
        with db() as con:
            rows=con.execute("""SELECT d.id,d.do_no,d.gmdc_do_no,d.do_date,d.customer_name,d.material,
                                      d.qty,d.remaining_qty,d.uom
                               FROM dos d WHERE COALESCE(d.remaining_qty,0)>0.000001
                               ORDER BY d.id DESC""").fetchall()
        for row in rows:
            label=row["do_no"] or row["gmdc_do_no"] or ""
            self.pending_tree.insert("","end",values=(
                row["id"],label,row["do_date"] or "",row["customer_name"] or "",
                row["material"] or "",f"{float(row['qty'] or 0):g} {row['uom'] or 'TON'}",
                f"{float(row['remaining_qty'] or 0):.3f} TON"))

    def on_pending_po_select(self,_event=None):
        selection=self.pending_tree.selection()
        if not selection:
            return
        values=self.pending_tree.item(selection[0],"values")
        ident=str(values[0])
        for label,row in getattr(self,"do_rows",{}).items():
            if str(row["id"])==ident:
                self.do_var.set(label)
                break
        self.selected_id=None
        self.update_pending_label()

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
                rows=con.execute("SELECT id,do_no,gmdc_do_no,customer_name,qty,uom FROM dos ORDER BY id DESC").fetchall()
                self.do_rows={}
                labels=[]
                for r in rows:
                    label=f"#{r['id']} | {'PO' if str(get_active_company()).casefold() == 'rajasthan gmdc' else 'DO'} {r['do_no'] or r['gmdc_do_no'] or r['id']} | {r['customer_name'] or ''} | {float(r['qty'] or 0):g} {r['uom'] or 'TON'}"
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
        for x in self.tree.get_children(): self.tree.delete(x)
        with db() as con: rows=con.execute("SELECT id,lr_no,lr_date,data_json FROM po_builtys ORDER BY id DESC").fetchall()
        for r in rows:
            d=json.loads(r["data_json"] or "{}")
            self.tree.insert("","end",values=(r["id"],r["lr_no"],r["lr_date"],d.get("consignee_name",""),d.get("truck_no","")))
        self.refresh_pending_po_list()

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
        out=self.export_dir/("PO_Builty_"+re.sub(r"[^A-Za-z0-9_-]","_",d["lr_no"])+".pdf")
        try:generate_po_builty_pdf(out,self.firm_getter(),d)
        except Exception as e:messagebox.showerror("PO Builty PDF",str(e));return
        if messagebox.askyesno("PO Builty PDF",f"Saved {out.name}. Open to print?"):self.open_file(out)


def generate_po_builty_pdf(out_path,firm,data):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate,Table,TableStyle,Paragraph,Spacer
    from xml.sax.saxutils import escape
    styles=getSampleStyleSheet()
    normal=ParagraphStyle("pon",parent=styles["Normal"],fontName="Helvetica",fontSize=7.5,leading=9,alignment=1)
    bold=ParagraphStyle("pob",parent=normal,fontName="Helvetica-Bold")
    title=ParagraphStyle("pot",parent=bold,fontSize=17,leading=19)
    sub=ParagraphStyle("pos",parent=bold,fontSize=10,leading=12)
    left=ParagraphStyle("pol",parent=normal,alignment=0)
    def p(x="",s=normal):return Paragraph(escape(str(x or "").replace("\n"," ")[:1500]),s)
    def t(rows,widths,heights=None):
        q=Table(rows,colWidths=widths,rowHeights=heights,hAlign="CENTER")
        q.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.8,colors.black),("INNERGRID",(0,0),(-1,-1),0.55,colors.black),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2)]));return q
    name=str(data.get("builty_company") or firm.get("name") or "MANSI COAL CARRIER").upper()
    address=data.get("builty_address") or firm.get("address") or "At Post Kunriya, Tal. Bhuj-Kutch. Corr Add Office No b/h sadguru complex madhapar highway"
    story=[t([[p("|| JAY MURLIDHAR ||")],[p(name,title)],[p("TRANSPORT CONTRACTOR & COMMISSION AGENT",sub)],[p(address,bold)]],[166*mm],[5*mm,11*mm,7*mm,7*mm]),
      t([[p("LR NO : "+data.get("lr_no",""),bold),p("LR DATE : "+data.get("lr_date",""),bold)]],[83*mm,83*mm],[7*mm]),
      t([[p("PO NO : "+data.get("po_no",""),bold),p("PO DATE : "+data.get("po_date",""),bold)]],[83*mm,83*mm],[7*mm])]
    leftparty=[p(": CONSIGNOR :",bold),p(data.get("consignor_name",""),bold),p(data.get("consignor_address","")),p("GSTIN: "+data.get("consignor_gstin",""),bold)]
    rightparty=[p(": CONSIGNEE :",bold),p(data.get("consignee_name",""),bold),p(data.get("consignee_address","")),p("GSTIN No/Unique No : "+data.get("consignee_gstin",""),bold)]
    meta=[("INVOICE NO","invoice_no"),("CHALAN NO","challan_no"),("E WAY BILL NO","eway_bill_no"),("E WAY BILL GENERATE DATE","eway_generate_date"),("E WAY BILL VALID DATE","eway_valid_date"),("Packing list no","packing_list_no")]
    rightmeta=t([[p(label+" : "+data.get(key,""),left)] for label,key in meta],[83*mm])
    parties=Table([[t([[x] for x in leftparty],[83*mm]),rightmeta],[t([[x] for x in rightparty],[83*mm]),""]],colWidths=[83*mm,83*mm],rowHeights=[30*mm,28*mm])
    parties.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.8,colors.black),("INNERGRID",(0,0),(-1,-1),0.55,colors.black),("SPAN",(1,0),(1,1)),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
    story.append(parties)
    weights=" / ".join(x for x in (data.get("weight_1",""),data.get("weight_2","")) if x) or ((str(data.get("builty_qty","")).strip()+" TON") if data.get("builty_qty") else "")
    story.append(t([[p("NO",bold),p("DESCRIPTION",bold),p("TOTAL BUNDLE",bold),p("TOTAL WEIGHT",bold)],[p("1",bold),p(data.get("description","")),p(data.get("total_bundle","")),p(weights)],[p("SUB TOTAL",bold),"",p(data.get("total_bundle",""),bold),p(weights,bold)]],[27*mm,76*mm,28*mm,35*mm],[11*mm,31*mm,7*mm]))
    truck=Table([[p("Truck No :",bold),p(data.get("truck_no",""),bold),p("CGST : "+data.get("cgst",""),bold)],[p("Drivers Name",bold),p(data.get("driver_name","")),p("SGST : "+data.get("sgst",""),bold)],[p("Address",bold),p(data.get("driver_address","")),p("IGST : "+data.get("igst",""),bold)],[p("Truck Owner Name",bold),p(data.get("truck_owner","")),p("TOTAL : "+data.get("total",""),bold)],[p("Mobile No",bold),p(data.get("mobile_no","")),p("",bold)]],colWidths=[27*mm,76*mm,63*mm],rowHeights=[7*mm]*5)
    truck.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.8,colors.black),("INNERGRID",(0,0),(-1,-1),0.55,colors.black),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
    story.extend([truck,Spacer(1,2*mm),t([[p(data.get("notes") or "* 1% SHORTAGE ALLOWED. WE ARE NOT RESPONSIBLE FOR QUALITY MATERIAL DIRECT LOADING FROM COMPANY",left),p("For : "+name,bold)]],[105*mm,61*mm],[15*mm])])
    doc=SimpleDocTemplate(str(out_path),pagesize=A4,leftMargin=22*mm,rightMargin=22*mm,topMargin=18*mm,bottomMargin=18*mm)
    doc.build(story);return str(out_path)
