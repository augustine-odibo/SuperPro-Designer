# SuperPro Designer: Protocol & User Guide

A practical quick-start guide to modeling processes in **SuperPro Designer** (Intelligen, Inc.), the process simulator for batch, continuous, and mixed batch/continuous processes in the pharmaceutical, biotech, specialty chemical, food, consumer goods, mineral processing, and water/wastewater industries.

> This is an unofficial community summary. For authoritative details, use the official documentation linked below.

---

## Table of Contents

1. [Official Resources](#official-resources)
2. [Key Concepts](#key-concepts)
3. [Modeling Protocol (Step by Step)](#modeling-protocol-step-by-step)
4. [Reports](#reports)
5. [Advanced Features](#advanced-features)
6. [Troubleshooting Tips](#troubleshooting-tips)
7. [References](#references)

---

## Official Resources

| Resource | Location |
|---|---|
| Online manual / help | https://docs.intelligen.com/spd/SPDWebHelp/ |
| Tutorial chapter | https://docs.intelligen.com/spd/SPDWebHelp/Bk7/Chapter2/Chapter2.htm |
| Training videos | https://www.intelligen.com/?p=3824 |
| Product features | https://www.intelligen.com/products/superpro-product-features/ |
| Papers & literature | https://www.intelligen.com/download/literature/ |
| In-app help | `Help` menu (the PDF manual ships with the installation) |
| Example files | `C:\Users\Public\Public Documents\Intelligen\SuperPro Designer\<version>\Examples` |

The software ships with 60+ example files (e.g., monoclonal antibody production, citric acid, lysine, biodiesel). Studying them is the fastest way to learn.

---

## Key Concepts

- **Unit Procedure**: a processing step carried out in one piece of equipment (e.g., a fermentor or a vessel procedure).
- **Operation**: an individual task within a batch unit procedure (Charge, Heat, React, Agitate, Transfer Out, CIP, and so on). Batch processes use this two-level procedure/operation hierarchy, consistent with ISA-95/S88 concepts.
- **Streams**: material connections between procedures. Feed (input) streams require composition and flow specs.
- **Design vs. Rating mode**: in *Design* mode the software sizes the equipment; in *Rating* mode you specify the size and it checks feasibility.
- **Pure Components & Stock Mixtures**: the materials registered in the process file (from the built-in database, DIPPR/PPDS, or user-defined).
- **Scheduling**: in batch mode, timing of operations determines cycle time, the Gantt chart, and the number of batches per year.

---

## Modeling Protocol (Step by Step)

### 1. Create a New Process File
- `File → New`
- Select the **operating mode**: **Batch** or **Continuous**.
- For batch processes, set the **annual operating time** (e.g., 7920 h/yr) and the **scheduling** basis via `Tasks → Recipe Scheduling Information`.

### 2. Register Materials
- `Tasks → Pure Components → Register, Edit/View Properties`
  - Add components from the Designer database (water, oxygen, glucose, biomass, etc.).
  - Create user-defined components if needed (supply MW, density, Cp, etc.).
- `Tasks → Stock Mixtures → Register, Edit/View Properties`
  - Register mixtures such as media and buffers.

### 3. Build the Flowsheet
- `Unit Procedures` menu → choose a procedure (e.g., *Vessel Procedure → in a Fermentor*).
- Click on the canvas to place it.
- Repeat for all steps of the process.
- Connect procedures using **Connect Mode** (stream drawing tool). Draw input and output streams for each procedure.

### 4. Add Operations (Batch Mode)
- Right-click a procedure → `Add/Remove Operations`.
- Add operations in chronological order, for example:
  1. Charge (media)
  2. Heat / Sterilize
  3. Cool
  4. Ferment / React
  5. Transfer Out
  6. CIP (cleaning)

### 5. Specify Operating Data
- Double-click each operation (or right-click → `Operation Data`) and set:
  - Amounts / flowrates and source streams
  - Duration, or process time calculated from rate
  - Temperature, pressure, agitation power
  - Reaction stoichiometry, extent/conversion, kinetics
  - Scheduling links (start relative to the beginning of the batch, or to another operation)
- Specify **feed stream** composition, temperature, and flow by double-clicking the stream.
- Set **equipment data**: right-click → `Equipment Data` → choose Design or Rating mode, the maximum allowable volume or size, and the material of construction.

### 6. Solve Material & Energy Balances
- `Tasks → Solve M&E Balances` (or click the calculator toolbar icon).
- Review the error/warning output window and correct the input specifications.
- Check stream results (double-click a stream, or use `View → Stream Table`).

### 7. Review Scheduling & Throughput
- `Charts → Gantt Charts → Operations GC` to see operation timing.
- `Charts → Equipment Occupancy` to identify the scheduling bottleneck.
- `Tasks → Throughput Analysis & Debottlenecking` for cycle-time and capacity analysis.
- Resource charts: labor, utilities, and material demand/inventory.

### 8. Perform Economic Evaluation
- Set prices: `Tasks → Pure Components/Stock Mixtures → Edit Properties → Economics` (purchase/selling prices).
- Define revenue streams (main product) and waste streams.
- Specify labor, utility, and consumable costs.
- `Tasks → Perform Economic Calculations`.
- Review the capital investment, operating cost, unit production cost, NPV, IRR, and payback time.

### 9. Generate Reports
- `Reports` menu (see the next section). Reports can be exported to Word, Excel, or HTML.

### 10. Save & Document
- `File → Save` (`.spf` file).
- Use the comment feature on procedures, streams, and operations to document assumptions.

---

## Reports

| Report | Abbrev. | Content |
|---|---|---|
| Stream Report | SR | Stream flows, compositions, properties |
| Materials & Streams Report | — | Raw material requirements and stream details |
| Economic Evaluation Report | EER | CAPEX, OPEX, profitability |
| Itemized Cost Report | ICR | Detailed cost breakdown |
| Cash Flow Analysis | CFR | Year-by-year cash flows |
| Environmental Impact Report | EIR | Waste/emissions characterization |
| Throughput Analysis Report | THR | Cycle time, bottlenecks |
| Emissions Report | EMR | Air emissions by component |

---

## Advanced Features

- **Recycle loops**: SuperPro converges processes with any number of recycle loops.
- **Thermodynamics**: VLE models for flash and distillation operations.
- **Independently cycling procedures**: procedures that cycle multiple times per batch (see the Citric Acid and Lysine examples).
- **Heat integration**: virtual heat recovery between streams (see the BioDiesel example).
- **COM/OLE Automation**: control SuperPro from Excel VBA, Python (`win32com`), etc., for parametric studies, optimization, sensitivity and risk analysis (e.g., with Crystal Ball), and custom reports. Documentation: `Help → COM Interface & Library`.
- **SchedulePro export**: send batch recipes to SchedulePro for production planning and scheduling of multi-product facilities.

### Example: Python Automation (Windows)

```python
import win32com.client

spd = win32com.client.Dispatch("SuperPro.Designer")
doc = spd.OpenDoc(r"C:\path\to\process.spf")
doc.DoMEBalances()
doc.DoEconomicCalculations()
doc.SaveDoc()
spd.CloseDoc(False, doc)
```

> Method names vary by version. Check `Help → COM Interface & Library` for the exact API of your installation.

---

## Troubleshooting Tips

- **"Equipment volume exceeded"**: increase the maximum vessel size, allow multiple units in parallel, or switch to Design mode.
- **Negative or unassigned flows**: check Transfer Out operations and stream connectivity.
- **Recycle not converging**: provide better initial guesses for tear streams, or adjust the convergence tolerance (`Tasks → Recycle Loop(s) Options`).
- **Long cycle time**: check the Equipment Occupancy chart for the bottleneck procedure, then add staggered equipment or parallel units.
- **Economics show zero revenue**: make sure the product stream is flagged as a *revenue stream* and the component has a selling price.

---

## References

- Intelligen, Inc. *SuperPro Designer User's Guide*. https://docs.intelligen.com/spd/SPDWebHelp/
- Intelligen, Inc. *Overview and Training Videos*. https://www.intelligen.com/?p=3824
- Intelligen, Inc. *SuperPro Designer Product Features*. https://www.intelligen.com/products/superpro-product-features/
- Petrides, D. et al. Numerous application papers on process simulation using SuperPro Designer. https://www.intelligen.com/download/literature/

---

*SuperPro Designer is a trademark of Intelligen, Inc. This document is an independent guide and is not affiliated with or endorsed by Intelligen.*
