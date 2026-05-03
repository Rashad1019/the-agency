# Real Estate Lead List Cleaner

**Developed by Rashad Ferguson — Data & AI Operations**
[rashad.io](https://rashad.io) · [github.com/rashadflowers99](https://github.com/rashadflowers99)

---

## The Problem This Solves

Every month, real estate agents export their lead lists from Zillow, Realtor.com, Facebook Ads, and their CRM — and every month, those exports are a mess.

Phone numbers in five different formats. Duplicate contacts. Names in ALL CAPS. Budget values that say `$450k` in one row and `650000` in the next. Leads with a valid phone number but no email — sitting there doing nothing. And the same name appearing four times with no way to know if it's one person or four.

Agents spend hours cleaning this by hand before running a single campaign. **This tool does it in under 10 seconds.**

---

## What It Does

One script. One command. Everything out the other end.

Run `python clean_leads.py --input your_export.xlsx` and you get a **formatted four-tab Excel workbook**:

| Tab | What it is |
|---|---|
| **Master List** | All leads in one view, color-coded by status |
| **✓ Import Ready** | Fully standardized, CRM-ready records |
| **📞 Email Outreach** | Warm leads with a phone but no email — personalized SMS + call scripts included |
| **⚠ Name Conflicts** | Same-name pairs scored by likelihood of being the same person |

Plus a plain-English **processing report** showing exactly what changed.

---

## What Gets Standardized

**Phone numbers → `(XXX) XXX-XXXX`**
Handles every format agents actually export:
`4077241971` · `407-724-1971` · `(407) 724-1971` · `407.724.1971` · `14077241971`

**Email addresses → lowercase, validated**
`MichaelScott31@AOL.com` → `michaelscott31@aol.com`
Malformed emails are flagged rather than silently passed through.

**Budget values → `$XXX,XXX`**
`$450k` → `$450,000` · `650000` → `$650,000` · `Unknown` stays labeled for qualification on first call.

**Lead source labels → canonical names**
`zillow lead` / `ZILLOW` / `Zillow` → `Zillow`
`website` / `Website Contact` → `Website`
`FB_Ad_1` / `FB_Ad_2` → `Facebook Ad`

**Names → Title Case, split into First + Last columns**
`JANE SMITH` / `john doe` / `sARAH cONNOR` → `Jane Smith`

**Duplicates removed by phone OR email match**
Keeps the first occurrence. Exact re-entries are removed before any other processing.

---

## The Email Outreach Queue

This is the feature that matters most.

The old approach flagged leads with missing emails and moved on. That approach silently discards warm leads with verified phone numbers and known budgets — potential commissions left on the table.

This pipeline puts them in a dedicated **Email Outreach** tab with:
- A **personalized SMS script** per lead, ready to copy and send
- A **personalized call script** per lead
- Empty `Email Collected` and `Date Contacted` columns to fill in after contact

**Live demo result (Florida leads file):**
- 11 leads with valid phones but no email on file
- Average budget: $522,727
- Estimated commission per close (~2.5%): ~$13,000
- One 5-minute call per lead. That is the only work required.

---

## Same-Name Conflict Detection

Phone and email deduplication catches exact re-entries. It does not catch the harder problem: the same name appearing multiple times with different contact details.

This pipeline detects every same-name group and scores each pair using a rule-based confidence system — not fuzzy matching, which only catches spelling variation. The scorer looks at:

| Signal | Weight |
|---|---|
| Same area code | +3 (strong — same metro = likely same person) |
| One entry has email, one doesn't | +2 (likely incomplete duplicate) |
| Identical budget | +2 |
| Budgets within $50K | +1 |
| Same lead source | +1 |
| Different Florida metros | -2 (813 vs 904 = Tampa vs Jacksonville) |
| Budgets differ by more than $50K | -1 |

**Three verdicts:**
- 🔴 **LIKELY DUPLICATE** — score ≥ 4. Review and probably merge.
- 🟡 **POSSIBLE DUPLICATE** — score 1–3. Agent makes the call.
- 🟢 **PROBABLY DIFFERENT** — score ≤ 0. Treat as separate leads.

The `⚠ Name Conflicts` tab shows every pair side by side with the reasoning spelled out. The last column — **Your Decision** — is blank for the agent to fill in: `MERGE`, `KEEP BOTH`, or `DISCARD B`.

**Live demo result:** 80 same-name pairs across 10 name groups. 7 flagged as likely duplicates.

---

## Before and After

**Before** (raw CRM export):

| Lead_Name | Phone | Email | Source | Budget |
|---|---|---|---|---|
| JANE SMITH | 4074855131 | | FB_Ad_1 | 650000 |
| john doe | 9049578340 | | website | 500000 |
| john doe | (407) 621-4051 | johndoe25@AOL.com | Realtor.com | 650000 |
| Hal Jordan | 561.417.9649 | HalJordan23@AOL.com | zillow lead | $450k |
| Hal Jordan | 4078536833 | HalJordan21@AOL.com | Realtor.com | 650000 |

**After:**

*Import Ready tab:*
| First Name | Last Name | Phone | Email | Source | Budget |
|---|---|---|---|---|---|
| John | Doe | (407) 621-4051 | johndoe25@aol.com | Realtor.com | $650,000 |
| Hal | Jordan | (561) 417-9649 | haljordan23@aol.com | Zillow | $450,000 |
| Hal | Jordan | (813) 853-6833 | haljordan21@aol.com | Realtor.com | $650,000 |

*Email Outreach tab:*
| First Name | Phone | Budget | SMS Script |
|---|---|---|---|
| Jane | (407) 485-5131 | $650,000 | Hi Jane, this is [AGENT NAME]... |
| John | (904) 957-8340 | $500,000 | Hi John, this is [AGENT NAME]... |

*Name Conflicts tab:*
| Name | Verdict | Why |
|---|---|---|
| Hal Jordan | LIKELY DUPLICATE | same area code (813) \| budgets within $50,000 |
| John Doe | POSSIBLE DUPLICATE | identical budget \| same lead source (Realtor.com) |

---

## Live Demo Numbers (Florida Leads File)

```
Records received:            50
Duplicates removed:          -7
────────────────────────────────
After deduplication:         43

✓  Import-ready:             32
📞 Email outreach queue:     11
⚠  Needs manual review:       0
🔍 Same-name conflicts:      80 pairs (7 likely duplicates)
```

---

## Works With Exports From

Zillow Premier Agent · Realtor.com · Facebook Lead Ads · Kvcore · Follow Up Boss · BoomTown · LionDesk · Any CRM that exports `.xlsx` or `.csv`

Column names are detected and normalized automatically regardless of your CRM's header format.

---

## How to Run

```bash
# Install dependencies (one time)
pip install pandas openpyxl

# Run on your file
python clean_leads.py --input your_export.xlsx

# Specify output folder
python clean_leads.py --input your_export.xlsx --output ./cleaned/
```

---

## The Time Math

| Task | Manual | This tool |
|---|---|---|
| Standardize 50 phone numbers | ~45 min | < 1 sec |
| Find and remove duplicates | ~30 min | < 1 sec |
| Fix name casing across 50 rows | ~15 min | < 1 sec |
| Normalize budget values | ~20 min | < 1 sec |
| Identify and score same-name conflicts | ~60 min | < 1 sec |
| Write personalized outreach scripts | ~30 min | < 1 sec |
| **Total** | **~3.5 hours** | **< 10 seconds** |

At $50/hour, that is **$175 of your time saved every month** — before a single call is made.

---

## About This Project

This is Project 01 in a portfolio of AI-assisted operations tools for real estate and professional service businesses, built by Rashad Ferguson.

Each tool targets one specific painful manual task — not generic AI automation, but a hardened workflow for the exact job you are already doing by hand every month.

**Series:**
- **Project 01** — Real Estate Lead List Cleaner *(this project)*
- **Project 02** — Client Intake Note Summarizer *(prompt architecture for attorneys, brokers, consultants)*
- **Project 03** — Monthly Operations Report Automator *(CSV → narrative report, no spreadsheet skills required)*

→ [rashad.io](https://rashad.io)
