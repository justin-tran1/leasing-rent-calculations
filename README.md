# leasing-rent-calculations

A CBRE-branded lease rent calculator. Enter up to six lease options, compare them side by side, chart them and export the full month-by-month schedule to Excel.

## Open the calculator

No install or build step. Open `index.html` in a browser (double-click it, or drag it into Chrome or Edge). Everything it needs ships in this repository, so it also works offline.

To serve it locally instead:

```bash
npm run serve   # http://localhost:8080
```

The page also runs as-is from any static host, such as GitHub Pages or SharePoint.

## What you can enter for each option

| Input | Notes |
|---|---|
| Option name | Shown on tabs, charts and Excel sheets |
| Size of the space | SF or SM (set once for all options) |
| Commencement date, length of term, ending date | Term in months or years. The ending date follows the term; edit the ending date directly for a specific date and the term updates (a partial final month is prorated by days) |
| Base rent | Per SF per year, per SF per month, total per month or total per year |
| Rate type | NNN, Modified Gross or Full Service |
| Annual escalation | % per year (default **3.0%**) or a fixed amount per year |
| OpEx | Per SF per year, with an annual increase (default **3.5%**) |
| Free rent | Number of months at the start or end of the term, or custom lease months (for example `1-3, 13`); covers base rent only, base rent and OpEx, or base rent, OpEx and parking |
| Parking | Number of spaces, cost per space per month and annual increase |
| Security deposit | First month's base rent, last month's base rent or none, times a number of months |

Global settings: currency symbol, area unit and the discount rate used for NPV.

## How the numbers work

- **Lease months** run from the commencement day to the day before it in the next month. **Lease years** are 12-month blocks from commencement. Escalations, OpEx increases and parking increases apply on each lease anniversary.
- **NNN:** the tenant pays all OpEx on top of base rent.
- **Full Service:** OpEx sits inside the rent. With "Tenant pays OpEx increases over the base year" ticked (the default), the tenant pays only the increase over lease year 1.
- **Modified Gross:** the tenant pays its share of OpEx directly (default 50%). The landlord's share sits inside the rent, and increases on that share over the base year pass through when the base-year box is ticked.
- **Free rent** abates the chosen charges in full for each free month; half months are allowed (2.5 months abates months 1 and 2 in full and month 3 by half).
- **Security deposit** uses the full, unabated base rent of the first or last lease month.
- **Effective rent** is the total lease cost divided by size and by term in years.
- **NPV** discounts each monthly payment, paid in advance, at the annual discount rate converted to a monthly rate.

Results are estimates for comparing proposals. Confirm all figures against the lease documents.

## Outputs

- **Comparison table** of every option: terms, total base rent, free rent, net base rent, OpEx, parking, total lease cost, difference from the lowest option, average monthly and annual cost, effective rent, NPV and security deposit.
- **Charts:** total monthly cost, net base rent per month, base rent rate, cumulative cost, annual cost, total cost by option, effective rent by option and a cost breakdown. Show all options on one overlapping chart or as separate charts (with an optional shared scale). Pick line, stepped line, area, column or horizontal bar; set the title, legend, height, line width, data labels, gridlines and markers; and colour each option from the CBRE palette or any custom colour. Download any chart as a PNG.
- **Lease schedule:** monthly or annual view for any option.
- **Excel export** (`Export to Excel`):
  - `Comparison` sheet with every option side by side and an editable NPV discount rate
  - `Monthly comparison` sheet with each option's rate, net base rent and total cost by lease month
  - `Charts` sheet with the charts as configured on screen
  - one sheet per option with the monthly schedule, totals, an annual summary, the inputs and the results

  Totals, annual rows, averages, effective rent and NPV are live Excel formulas, so the workbook stays consistent if a month is edited.
- **Save scenario / Open scenario** stores every option and chart setting in a `.json` file to reopen or share. The browser also remembers the last session automatically.

## Brand

Colours, type and logo follow the CBRE brand guidelines: CBRE Green `#003F2D` with Accent Green `#17E88F`, Financier Display headings and Calibre body text (with Georgia and Arial as fallbacks where the brand fonts are not installed). The default chart order is the CBRE Charts & Graphs palette with accent green and wheat skipped, so neighbouring series stay distinguishable for colour-blind readers. The "CBRE high contrast" theme keeps every series above 3:1 contrast on white.

## Project layout

```
index.html          page markup
css/styles.css      CBRE styling
js/calc.js          calculation engine (no DOM; runs in Node for tests)
js/format.js        number, date and label formatting
js/charts.js        Chart.js configuration and rendering
js/excel.js         Excel workbook builder (ExcelJS)
js/app.js           UI state and wiring
js/logos.js         CBRE logos embedded for the Excel export (generated)
assets/             CBRE logos
vendor/             Chart.js and ExcelJS browser builds with licences
tests/              Node tests for the engine and the Excel export
```

## Development

```bash
npm install
npm test          # engine and Excel export tests
npm run vendor    # refresh vendor/ and js/logos.js after upgrading Chart.js or ExcelJS
```
