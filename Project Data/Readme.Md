# Audit of Financial Controls 

## Disclaimer

This project is not related to any real company. It was independently created for learning and practice purposes only. The data used is fictional but structured to closely reflect how real financial data looks in industry.

## Project Objective

The objective of this project is to practically implement the **Audit of Financial Controls** phase of internal auditing.

In this phase, an auditor reviews vendor payment transactions to test whether the organization's internal controls are working as intended. The focus is not on the accuracy of financial numbers — but on whether the **rules and processes** of the organization are being followed.

This project was built to bridge the gap between audit theory and real-world application — using realistic data structure, actual audit techniques, and professional documentation standards.

## What is an Audit of Financial Controls

Every organization has internal rules about how payments should be processed. For example:

- Payments above a certain amount must be approved by a senior person
- Every payment must have a supporting invoice
- Payments can only be made to vendors who are officially registered and approved

An **Audit of Financial Controls** checks whether these rules are actually being followed in practice. When a rule is found to be broken or bypassed, it is called a **control failure** — and the auditor documents it as a finding.

This type of audit is a standard part of the internal audit function in any organization and is particularly important in finance and fintech companies where large volumes of payments are processed regularly.

## Project Structure

The project is an Excel workbook with 5 sheets, each representing a specific part of the audit process.

| Sheet | Name | Purpose |
|-------|------|---------|
| 1 | Audit Program | Planning document — lists all audit steps |
| 2 | Vendor Master | Approved vendor reference list |
| 3 | Raw Data | 500 vendor payment transactions |
| 4 | Sampling | Auditor's testing workspace |
| 5 | Findings | Documented control failures |

---

## Phase 1 — Planning

### What Happens in Planning

Before touching any data, an auditor prepares an **Audit Program** — a structured checklist of everything that will be tested during the audit. This document serves as the roadmap for the entire audit.

The Audit Program in this project covers the following areas:

- Review of raw transaction data for surface-level issues
- Selection of a representative sample from the full population
- Testing of authorization and approval controls
- Verification of supporting documentation (invoices)
- Matching of payments against the approved vendor master list

### Why Planning Matters

Without a plan, an auditor might miss important areas or spend too much time on low-risk items. Planning ensures that the audit is focused, efficient, and covers all key risk areas.

---

## Phase 2 — Fieldwork

Fieldwork is where the actual audit testing takes place. It is divided into four steps.

---

### Step 1 — Initial Data Review

**Sheet: Raw Data (Sheet 3)**

The first thing an auditor does when receiving transaction data from the Finance department is a quick visual scan — without any formulas — to identify obvious problems.

In this project, the following checks were performed visually and through basic Excel techniques:

---

#### Missing Invoices
Every payment above a certain threshold must have a supporting invoice. The Invoice Number column was filtered to find blank cells.

```
Filter → Invoice No. column → Blanks
```

---

#### Duplicate Transactions
A duplicate payment occurs when the same vendor is paid the same amount more than once. This was detected using the COUNTIFS formula in a helper column.

```excel
=COUNTIFS($C$3:$C$500, C3, $E$3:$E$500, E3)
```

Any result greater than 1 indicates a potential duplicate.

---

#### Round Numbers
Genuine invoices from vendors almost always include GST, delivery charges, and other components — making the total a non-round number like PKR 47,832. An unusually high number of round amounts like exactly PKR 50,000 or PKR 100,000 can indicate fabricated invoices.

Round numbers were flagged using the MOD function:

```excel
=IF(MOD(E3, 10000) = 0, "Round Number", "OK")
```

Conditional formatting was also applied to visually highlight round amounts in red.

---

#### Weekend Transactions
Payments processed on weekends are unusual because offices are closed. Weekend transactions were identified using the WEEKDAY function:

```excel
=IF(WEEKDAY(B3, 2) > 5, "Weekend", "OK")
```

The parameter `2` starts the week on Monday. Any result above 5 means Saturday or Sunday.

Not all weekend transactions are findings. Only those combined with other red flags — such as a suspicious vendor or missing invoice — were escalated.

---

#### Threshold Splitting
This is a technique where a single large payment is intentionally broken into multiple smaller payments to stay below the approval limit. For example, if payments above PKR 50,000 require VP Finance approval, someone might split a PKR 1,50,000 payment into three payments of PKR 49,000 each to avoid getting senior approval.

This was detected by checking for multiple payments to the same vendor in the PKR 45,000–49,999 range:

```excel
=AND(COUNTIFS($C$3:$C$500, C3, $E$3:$E$500, E3) > 1, E3 >= 45000, E3 < 50000)
```

---

### Step 2 — Sampling

**Sheet: Sampling (Sheet 4)**

An auditor cannot test all 500 transactions individually — it would be impractical and is not how audits work. Instead, a representative **sample** is selected for detailed testing.

In this project, two sampling methods were used:

---

#### Random Sample — 40 Transactions

40 transactions were randomly selected from the full population of 500 using the RANDBETWEEN function:

```excel
=RANDBETWEEN(3, 501)
```

This generates a random row number. The transaction from that row in the Raw Data sheet was then copied to the Sampling sheet.

**Important step:** After generating all 40 random numbers, they were converted from formulas to fixed values using Paste Special → Values Only. This prevents the numbers from changing every time the sheet recalculates.

Random sampling ensures that every transaction in the population has an equal chance of being selected — making the sample representative of the full dataset.

---

#### High Value Sample — Top 10 Transactions

The 10 highest-value transactions were always included in the sample regardless of random selection. High-value transactions carry more financial risk and are always prioritized in audit testing.

The top 10 amounts were identified using the LARGE function:

```excel
=LARGE(RAW_DATA!E3:E501, 1)    → Highest amount
=LARGE(RAW_DATA!E3:E501, 2)    → Second highest
=LARGE(RAW_DATA!E3:E501, 3)    → Third highest
...and so on up to 10
```

Each identified amount was then located in the Raw Data sheet using CTRL+F and the transaction details were copied to the Sampling sheet.

---

### Step 3 — Controls Testing

**Sheet: Sampling (Sheet 4)**

Every transaction in the sample was tested against the company's authorization policy. The policy defines who is allowed to approve payments based on the amount:

| Amount | Required Approver |
|--------|------------------|
| Below PKR 50,000 | Finance Manager |
| PKR 50,001 to PKR 200,000 | VP Finance |
| Above PKR 200,000 | CFO |

For each sampled transaction, the Approved By column was checked against the amount to verify the right person had approved it.

An approval violation was flagged using this formula:

```excel
=IF(AND(E3 > 200000, OR(G3 = "Finance Manager", G3 = "Head Ops")), "Violation", "OK")
```

This checks whether a payment exceeding PKR 200,000 was approved by someone with insufficient authority.

**Finding identified:** One transaction of PKR 285,750 was found approved by a Finance Manager instead of the CFO — a clear control bypass.

---

### Step 4 — Vendor Verification

**Sheet: Vendor Master (Sheet 2) + Sampling (Sheet 4)**

Every vendor that received a payment was cross-checked against the approved Vendor Master List using VLOOKUP. The Vendor Master contains the official list of registered and approved vendors along with their NTN numbers and contact persons.

```excel
=VLOOKUP(E7, VENDOR_MASTER!B:F, 5, 0)
```

**How the formula works:**
- `E7` — the vendor name from the Sampling sheet
- `VENDOR_MASTER!B:F` — the range in the Vendor Master sheet to search through
- `5` — return the 5th column which contains the approval status
- `0` — exact match required

**Possible results:**

| Result | Meaning |
|--------|---------|
| ✅ Approved | Vendor is on the approved list — OK |
| ❌ NOT Approved | Vendor exists but is not approved — finding |
| #N/A | Vendor is not in the master list at all — serious finding |

Any result other than Approved requires investigation.

---

**What is NTN?**
NTN stands for National Tax Number — a unique identifier issued by FBR (Federal Board of Revenue) to every registered business in Pakistan. It is used to verify that a vendor is a legitimate, formally registered company. A vendor without an NTN should not be receiving payments.

---

**What is a Contact Person?**
Each vendor in the master list has a designated contact person. In audit, the contact person is used for **third party confirmation** — if a transaction looks suspicious, the auditor can reach out directly to the vendor to verify whether the invoice was genuinely issued and whether the payment was received.

---

## Formulas Used

| Formula | Purpose | Where Used |
|---------|---------|------------|
| `=COUNTIFS()` | Detect duplicate payments by matching vendor and amount | Raw Data — helper column |
| `=MOD()` | Flag amounts that are suspiciously round numbers | Raw Data — helper column |
| `=WEEKDAY()` | Identify transactions processed on weekends | Raw Data — helper column |
| `=AND()` / `=OR()` | Combine multiple conditions in a single check | Raw Data — approval check |
| `=RANDBETWEEN()` | Generate random row numbers for sample selection | Sampling sheet |
| `=LARGE()` | Extract the top 10 highest transaction amounts | Sampling sheet |
| `=VLOOKUP()` | Match vendor names against the approved vendor list | Sampling sheet |
| `=IF()` | Label results as OK or flagged based on conditions | Multiple sheets |

---

## Findings Summary

Three control failures were identified and documented in the Findings sheet.

---

### F-001 — Missing Invoice
**Risk Level:** Medium

A payment of PKR 80,648 was processed and completed without a supporting invoice from the vendor. The remarks field confirmed the invoice was absent at the time of payment.

**Recommendation:** Obtain the invoice immediately. No payment should be processed in future without a valid invoice on file at the time of processing.

---

### F-002 — Duplicate Payment
**Risk Level:** Medium

Two separate payments were made to the same vendor for the exact same amount of PKR 16,735 on the same date. Both transactions referenced the same invoice number — confirming this was a duplicate payment. The remark on the second transaction said "Reprocessed."

**Recommendation:** Verify with the bank whether both payments were debited. If confirmed, recover PKR 16,735 from the vendor. Implement a system control to block reprocessing of already paid invoices.

---

### F-003 — Approval Rule Violation
**Risk Level:** High

A payment of PKR 285,750 was approved by a Finance Manager. Per company policy, any payment above PKR 200,000 must be approved by the CFO. This is a clear bypass of the authorization control.

**Recommendation:** Obtain retrospective CFO approval immediately. Review all transactions above PKR 200,000 to identify similar violations. Implement a system-level restriction to prevent junior staff from approving payments beyond their authorized limit.

---

## Key Learnings

Working through this project practically demonstrates several important audit concepts:

**Sampling is not guessing.** There is a structured logic behind how a sample is selected — combining random selection with risk-based selection of high-value items ensures balanced and meaningful coverage.

**One red flag is not always a finding.** A weekend transaction alone is not a finding. A weekend transaction combined with a suspicious vendor and a missing invoice — that is a finding. Context and combination of red flags matter.

**The auditor's job is to find and report — not to fix.** When a duplicate payment was found, it was documented with a recommendation. Reversing the payment is the Finance team's responsibility — not the auditor's.

**Clean-looking data does not mean clean controls.** Even well-formatted, complete-looking data can hide control failures. Approval violations and duplicate payments are not visible from a quick glance — they require structured testing with the right formulas and techniques.

**Documentation is everything.** Every step taken, every transaction tested, and every finding identified was recorded in working papers. In a real audit, undocumented work is treated as work that was never done.

---

*This project was independently built for practice and learning. All data is fictional.*
`Sheet 5 — Audit Findings`
This sheet contains the formal findings identified during the audit.
A finding is raised when a control has failed — for example, a payment was made without an invoice, or a payment was approved by someone who did not have the authority to approve it. Each finding includes a description of what went wrong, the risk level, and a recommendation for how it should be fixed.
