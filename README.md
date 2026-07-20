# TA Dashboard Generator

A Python + Pandas + Streamlit dashboard for Talent Acquisition analytics — **Regular hires only**.

## What's new in this build

- ✅ **STRICT Regular filter** (Option A) — only rows where Col AB (Joined) / Col X (Open Reqs) explicitly contains "regular"
- ✅ **Median TTF removed** everywhere (KPIs, charts, tables)
- ✅ **Data labels** added to every bar chart, line chart, and pie chart
- ✅ **Hiring Source** = Column W (`Job_App_for_Hire_1: Source`)
- ✅ **Hiring Manager** = Column X (`Job_App_for_Hire_1: Hiring Manager`, as text)
- ✅ **Conversion / Quality section removed** (Interview→Offer, Offer→Hire, Interview→Hire, App→Hire)
- ✅ **Workload vs Onboarding Output scatter removed** from Recruiter tab
- 🆕 **Primary Recruiter on Open Reqs = Column M** (position-based, matches the Joined tab's Col H convention)
- 🆕 **Detailed stage funnel on Open Reqs (Col AK)** — Open, Screening/Interview, Reference Check, Background Check, Offer, EA, Ready for Hire (not just the binary Pure Open / Advanced Stage split)
- 🆕 **New Insights tab (Tab 4)**: aging-vs-offer-decline correlation by Department/Region/Hiring Manager/Recruiter/Source, QoQ narrative insights gated by a materiality filter (sample size ≥ 5, shift ≥ 25%, or aged-share ≥ 35% combined with declines), and recruiter TTF outlier detection (≥10 days worse than team average)

## Column mappings (locked)

### Candidate Joined detail report

| Column | Meaning | Used for |
|---|---|---|
| **H** | Primary Recruiter | Recruiter cuts |
| **W** | Job_App_for_Hire_1: Source | Hiring Source cuts |
| **X** | Job_App_for_Hire_1: Hiring Manager | Hiring Manager cuts (text) |
| **AA** | Department Name | Department cuts |
| **AB** | Regular Hire flag | **STRICT Regular filter** (must contain "regular") |
| **AD** | Head of Department | Business Owner cuts |
| **AK** | Total Days to Onboard | **TTF numerator** |
| **AM** | Region | Region cuts |
| B | Hire Date | Month / Quarter grouping |

### Job Req & open Pos

| Column | Meaning | Used for |
|---|---|---|
| **A** | Open Job Requisition: Reference ID | Open Reqs (unique) |
| **B** | Position ID | Positions open without Req |
| **M** | Primary Recruiter | Recruiter cuts / ownership |
| **P** | Business Head | Ownership |
| **X** | Worker Type | STRICT Regular filter |
| **AK** | Candidate stage | Pure Open vs Advanced Stage, and detailed 7-stage funnel |
| **AM** | Days bracket | Aging bucket |
| **AO** | Region | Region cuts |

### Disposition Reasons Report

| Column | Meaning | Used for |
|---|---|---|
| **H** | Disposition reason | Did-not-join-after-offer flag |
| (header-detected) | Department / Region / Hiring Manager / Recruiter / Source | Correlating declines against aging cuts on the Insights tab |

## TTF formula

**Portfolio TTF** = SUM(Col AK Total Days to Onboard for Regular hires) ÷ COUNT(Regular Onboards with valid TTF)

- Outliers excluded: blank, negative, or > 365 days
- Only Portfolio Average shown (Median removed)

TTF can be sliced by: **Region, Recruiter, Department, Head of Department, Job Family**.

## Three dashboard sections

### 1. Filled Roles
- Hiring Outcomes: Onboards by Recruiter, Head of Dept, Department, Region, Source (Col W), Hiring Manager (Col X), Job Family
- Efficiency: Portfolio TTF + Total Days + Regular Onboards used, cuts by all dimensions, QoQ trend
- (Conversion / Quality removed)

### 2. Open Roles
- Requisition Health: Open Reqs, Pure Open, Advanced Stage, Positions without Req
- Aging & Risk: Aging buckets, Aging × Stage Class heatmap
- Ownership: By Recruiter, Hiring Manager, Function, Business Owner, Region

### 3. Recruiter & Team Performance
- Recruiter summary (Onboards, Open Req Load, Active Candidates, Portfolio TTF)
- Onboards / Load / Active Candidate charts
- Portfolio TTF by Recruiter
- (Workload vs Output scatter removed)
- Portfolio TTF by Recruiting Manager

### 4. Insights (new)
- **Narrative insights**: QoQ shifts in hiring volume and TTF by Department / Region / Hiring Manager / Recruiter / Source, only reported when they clear a materiality bar (>=5 sample size, >=25% shift) - no low-signal activity metrics
- **Aging vs. offer-decline correlation**: for any cut, how many open reqs are aged 90+ days alongside how many post-offer declines fall in the same cut - flags where aging and pipeline breakage are compounding
- **Recruiter TTF outliers**: recruiters >=10 days slower than team average Portfolio TTF, with enough onboards (>=5) to be a real signal

## Cross-cutting filters

Region • Head of Dept / Business Owner • Department • Job Family • Hiring Manager • Recruiter • Recruiting Manager • Time Period (Year, Quarter, Month, YTD) • Hiring Source

## Run

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Then upload `Recruiting Dashboard.xlsx` from the left sidebar. Strict Regular filter is auto-applied via Col AB (Joined) and Col X (Open Reqs).

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI, filters, 3 tabs, charts |
| `ta_engine.py` | Data ingestion, business rules, TTF engine |
| `requirements.txt` | Python dependencies |
| `README.md` | This file |
