# Custom Excel Template — Named Range Guide

This document is for whoever sets up a client's (e.g. Colliers') custom Excel
export. It explains exactly how to prepare a client-provided master
workbook so our platform can fill it in automatically, with their native
formulas, macros, and branding left completely untouched.

## How it works

1. The client opens their own master valuation workbook (their "master
   template") in Excel.
2. Using Excel's **Name Manager** (Formulas tab → Name Manager), they give
   specific input cells one of the reserved names listed below — for
   example, naming the cell where the property name goes `IU_PropertyName`.
3. They upload that workbook once, through the admin panel (per
   organization).
4. On every future analysis for that organization, our engine opens their
   file, finds those names, and writes the parsed values directly into
   those cells — nothing else on the sheet is touched. Their formulas,
   macros, number formats, and logos are preserved exactly as they were.
5. If a different organization logs in, their own distinct template is
   used instead. Templates are fully isolated per organization — one
   organization's proprietary model is never visible to, or reused for,
   another.

A client only needs to name the cells they actually want filled in. A name
we don't recognize is ignored; a name we recognize that they didn't define
is simply never written to — nothing is guessed or forced into a cell they
didn't explicitly mark.

## Two kinds of named ranges

### Scalar fields — one value per analysis

Name a single cell with one of these. Each one is written once, straight
from that analysis's parsed metrics (or, for the three debt fields, from
the buyer-entered financing assumptions on the download request):

| Name | What goes there |
|---|---|
| `IU_PropertyName` | Property name |
| `IU_PropertyAddress` | Property address |
| `IU_AssetClass` | Asset class (e.g. Office, Retail, Industrial, Multifamily) |
| `IU_TotalUnitCount` | Total unit count |
| `IU_TotalBuildingAreaSF` | Total building area (SF) |
| `IU_PhysicalOccupancyPct` | Physical occupancy % |
| `IU_EconomicOccupancyPct` | Economic occupancy % |
| `IU_TotalVacantUnits` | Total vacant units |
| `IU_AvgInPlaceRent` | Average in-place monthly rent |
| `IU_AvgMarketRent` | Average market monthly rent |
| `IU_AnnualizedGPR` | Annualized Gross Potential Rent |
| `IU_TotalAnnualConcessions` | Total annual concessions |
| `IU_NetRentalIncome` | Net rental income (from T12) |
| `IU_DebtLTVPct` | Buyer-entered LTV % |
| `IU_DebtInterestRatePct` | Buyer-entered interest rate % |
| `IU_DebtAmortizationYears` | Buyer-entered amortization (years) |

### Row fields — one row per tenant (the rent roll table)

A rent roll has a variable number of tenants, but a named range is always
a single cell — so instead of naming a range, **name the FIRST DATA CELL
of each column**. For example, if tenant names start in cell `B10` of the
rent roll table, name that one cell `IU_Row_TenantName` — not the whole
column, and not the header above it.

Our engine writes one row per tenant **downward** from that cell. All of
the `IU_Row_*` names in a template must sit on the **same row** (they mark
the top-left corner of the same table) — if they don't, the upload is
rejected with an explanation of which names disagree.

| Name | Column |
|---|---|
| `IU_Row_TenantName` | Tenant name |
| `IU_Row_SuiteID` | Suite ID |
| `IU_Row_SpaceType` | Space type |
| `IU_Row_SF` | Square footage |
| `IU_Row_LeaseStart` | Lease start date |
| `IU_Row_LeaseEnd` | Lease end date |
| `IU_Row_BaseRent` | Base rent |
| `IU_Row_AnnualRent` | Annualized rent |
| `IU_Row_Concessions` | Concessions |
| `IU_Row_Escalation` | Escalation |
| `IU_Row_RecoveryStructure` | Recovery structure |
| `IU_Row_BaseYear` | Base year |
| `IU_Row_OccupancyStatus` | Occupancy status |
| `IU_Row_MLAProfile` | Market Leasing Assumptions profile |

## Safety behavior (what to expect)

- **Nothing is ever fabricated.** If a value genuinely isn't present in the
  uploaded source documents for a given deal, that cell is simply left
  untouched rather than filled with a guess or a placeholder.
- **A template with zero recognized names is rejected at upload**, with the
  full list of supported names, so a broken setup is caught immediately —
  not on the client's first real analysis.
- **If the `IU_Row_*` names don't all anchor to the same row**, the upload
  is rejected with the exact conflict, rather than silently guessing which
  one is "right."
- **If an analysis has more tenants than the template's row table can hold
  without overwriting a cell that already contains one of the client's own
  formulas**, the ENTIRE row write is refused (not just the overflow rows) —
  the download fails with a clear error instead of silently truncating
  data or corrupting part of the client's model. This means a template's
  row table should be built with enough pre-formatted rows for the
  client's typical rent roll size.
- **Re-uploading a replacement template overwrites the previous one** for
  that organization — there's no need to delete the old one first, and no
  old file is left behind on our servers.
- **Macros are preserved** automatically for `.xlsm` uploads.

## Admin endpoints

| Action | Endpoint |
|---|---|
| Upload / replace a template | `POST /api/admin/organizations/<org_id>/excel-template/upload/` (multipart, field `file`) |
| Inspect a template | `GET /api/admin/organizations/<org_id>/excel-template/` |
| Remove a template | `DELETE /api/admin/organizations/<org_id>/excel-template/delete/` |

All three require an authenticated platform admin. The inspect endpoint
reports which of the names above the uploaded template actually defines,
and which known names it's missing — useful for telling a client exactly
what else they could name if they want more fields auto-filled.

Removing a template immediately falls the organization back to our
standard generated workbook on their very next export — there's no
separate toggle to flip.
