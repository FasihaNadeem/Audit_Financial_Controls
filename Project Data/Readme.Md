### Audit of Financial Controls 

Disclaimer
This project is not related to any real company. It was independently created for learning and practice purposes only. The data used is fictional but structured to closely reflect how real financial data looks in industry.

Project Objective:
This is a hands-on practice project that simulates how an Internal Auditor performs an Audit of Financial Controls over vendor payments.
In real organizations, the Finance department processes hundreds of payments every year — to vendors, suppliers, and service providers. An internal auditor's job is to independently review these payments and check whether the company's internal rules and controls are being followed properly.
This project takes that exact process and implements it practically — from receiving raw financial data all the way to documenting audit findings.

The Audit of Financial Controls is a specific phase of internal auditing. In this phase, an auditor does not just look at numbers — they test whether controls are working. For example:

Was every payment approved by the right person?
Does every payment have a supporting invoice?
Were payments made only to approved vendors?
Were any payments duplicated?

To understand this process deeply, this project was built to practically implement each step an auditor follows — using real-world structure, realistic data, and actual audit techniques.

Project Structure
The project is built as an Excel workbook with 5 sheets. Here is what each sheet contains and why:

Sheet 1 — Audit Program
The audit program is the planning document. Before an auditor touches any data, they plan what they are going to test and why. This sheet lists all the audit steps that were performed in this project — what was checked, what the objective was, and whether it was completed.
In real audits, this document is prepared during the planning phase before fieldwork begins.

Sheet 2 — Vendor Master List
This is the approved vendor list of the organization — the list of companies that are officially registered and authorized to receive payments.
In practice, this list is provided to the auditor by the Finance or Procurement department. The auditor uses this list to check whether every payment in the transaction data was made to an approved vendor. Any payment going to a vendor not on this list is a red flag.

Sheet 3 — Raw Transaction Data
This sheet contains 500 vendor payment transactions representing one full financial year (FY 2024-25).
The data is fictional but follows the same structure as real financial data provided by a Finance department to an auditor. It includes transaction IDs, dates, vendor names, departments, amounts, payment modes, approvers, invoice numbers, and status.
The auditor does not process all 500 transactions — instead, they select a sample to test.

Sheet 4 — Sampling
This sheet is where the auditor's testing work happens.
Instead of reviewing all 500 transactions, the auditor selects a representative sample of 50 transactions — 40 chosen randomly and 10 selected because they are the highest value transactions (which carry the most risk).
Each sampled transaction is then tested against the company's internal control rules — checking approvals, invoices, and vendor status — and the auditor records their observations.

Sheet 5 — Audit Findings
This sheet contains the formal findings identified during the audit.
A finding is raised when a control has failed — for example, a payment was made without an invoice, or a payment was approved by someone who did not have the authority to approve it. Each finding includes a description of what went wrong, the risk level, and a recommendation for how it should be fixed.
