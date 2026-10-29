# Transport — Rajasthan GMDC

Independent desktop application for the Rajasthan GMDC workflow, maintained separately from [Transport](https://github.com/rajeshahir143/Transport).

## Workflow
- Import a Purchase Order (PO) PDF or enter PO details manually.
- Create a DO from the selected PO.
- Create one or more Builtys linked to a DO.
- Builty quantities reduce the DO's remaining quantity; over-allocation must be blocked.
- Continue through Bill, Freight Paid, Dashboard, and exports.
- Use the PO Builty form to print the Rajasthan B167-style document.

## Data safety
- The default active company is **Rajasthan GMDC**.
- Database files are stored under this repository's own application directory in `Companies/Rajasthan GMDC/`.
- This application does not use the original Transport repository's database files.
- Verify backup/restore and the full PO → DO → Builty → Bill workflow before using live data.

## Setup
Install Python dependencies:
```bash
pip install -r requirements.txt
python main.py
```
