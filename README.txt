PENSION PACKET GENERATOR
========================

What it does
------------
This interactive Python script asks for an employee's retirement details and
creates three editable Microsoft Word outputs:

1. Retirement_Certificates_<name>.docx
   - Non-Liability Certificate
   - Service Verification Certificate
   - Specimen Signature Certificate
   - Letter to the Inspector of Schools, MDC Majuli

2. Kritagyata_Form_<name>.docx
   - Data Required for Online Pension File (Kritagyata)
   - Employee data, family members, enclosure checklist and signature spaces

3. Complete_Pension_Packet_<name>.docx
   - The four retirement pages plus the Kritagyata form appended as an extra
     page/section, in one editable Word document.

Files in this package
---------------------
pension_packet_generator.py
Retirement_Certificate_Formats_Blank_Template.docx
Data_Required_for_Online_Pension_File_Kritagyata.docx
requirements.txt
run_pension_generator.bat
pension_form.html (double-click web form, works offline, no Python needed)
vendor/jszip.min.js (needed by the web form, keep beside it)
embed_templates.py (re-run after changing the DOCX templates)
server.py (optional backend, not needed for the form)

Web form (easiest, no installation)
-----------------------------------
1. Double-click pension_form.html (or run_pension_form.bat).
2. Fill the 4 steps. Use the Settings tab to change the school
   name, UDISE line, address, inspector office, district and the
   default Headmaster/DDO name. They print on every document and
   are saved in the browser.
3. Review, then download the ZIP with all three Word files.
   Everything runs offline in the browser; nothing is uploaded.
4. If you edit either DOCX template, run
   "python embed_templates.py" once so the form picks up the change.

Command-line version (needs Python)
------------------------------------
1. Install Python 3.10 or newer from https://www.python.org/downloads/ .
   During installation, enable "Add Python to PATH".
2. Extract all the ZIP contents into one folder. Do not separate the two DOCX
   templates from the Python script.
3. Open Command Prompt in that folder and run:
      python -m pip install -r requirements.txt
4. Start the program by double-clicking run_pension_generator.bat, or run:
      python pension_packet_generator.py
5. Answer the prompts. Press Enter for optional fields you do not know.
6. Generated files are saved in the Generated_Pension_Packets subfolder unless
   you enter another output folder.

Important
---------
- The script runs locally and does not upload employee information anywhere.
- The qualifying-service duration is calculated from the joining and retirement
  dates using calendar years/months/days. Confirm it against the Service Book.
- Service verification dates and page references must be entered from the
  original Service Book. If no periods are entered, the verification table is
  left blank.
- The Kritagyata enclosure checklist marks selected items as "Enclosed" and
  all other items as "Pending". Review each status before printing.
- Optional fields can remain blank; the program does not invent missing values.
- Always proofread the generated documents and verify official records before
  signature or submission.
