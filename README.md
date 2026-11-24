# 220 kV Overhead Transmission Line Corona, Sag & Environmental Performance Study

Independent Python reproduction of an existing graduate electrical engineering study. The objective is to check workbook calculations and produce portfolio figures while preserving the workbook's selected models and seven environmental scenarios.

## Run

Requires Python 3.10 or later. From this project folder:

```sh
python -m pip install -r requirements.txt
python analysis.py "C:/path/to/Project_4_220kV_Corona_Sag_Environmental_Study.xlsx"
```

Alternatively, place the original workbook beside `analysis.py` using that filename and run `python analysis.py`. In PyCharm, select the interpreter with these dependencies and supply the workbook path as a script parameter. The provided source filename has no `(3)` suffix; its actual contents were inspected. No source workbook is included or rewritten.

Uses only NumPy, pandas, Matplotlib, openpyxl and standard Python libraries. The script reads raw input cells, independently calculates engineering results, and compares against Excel's saved numeric caches. It does not evaluate formulas with Excel or claim to refresh those caches. Read-only loading also avoids an existing chart XML definition that openpyxl cannot load in normal mode.

## Calculations and outputs

Topics include dry/wet Peek-type corona thresholds and losses, visual inception, voltage/radius sensitivity, twin-bundle fields and electrical parameters, parabolic sag and support forces, span/tension/vertical-load sensitivity, simplified thermal expansion, wind loading and the seven combined conditions.

Running creates one `outputs/results_summary.csv` and exactly nine 300 DPI PNG figures in `outputs/figures`. CSV columns follow the requested format. Each Scenario includes the source sheet/cell, and Difference means Python minus saved Excel value. CSV numbers retain 15 significant digits so small differences are visible; round only for presentation. No extra graphs, workbooks or PDF reports are generated.

## Corrected workbook formulas

- `05_Temperature!B31` uses scenario temperature minus reference temperature. At 55°C, the 30°C increase gives **10.9995 m** sag and **14.0005 m** ground clearance, reflected in Graphs 6 and 7.
- `05_Temperature!B23` uses attachment height minus hot sag, giving **13.2322 m** ground clearance at 75°C.
- `07_Environmental!D18:D22` use attachment height minus vertical sag. Graph 9 applies this definition to all seven scenarios, and clearance reductions use those same values.

## Numerical verification

Comparison uses full working precision and **absolute tolerances only**. MATCH requires `abs(Python − Excel) <= tolerance`; otherwise DIFFERENCE is retained. A missing/non-numeric cache gives NO EXCEL CACHE, never a zero. The tolerances allow floating-point and saved-cache precision noise, not modelling discrepancies.

| Reported unit | Absolute tolerance |
|---|---:|
| m, N/m, kV/cm, kN, nF/km, kW/km/phase, °C | 1e-8 |
| kV, kV LL, %, A, Pa, degrees, ohm/phase | 1e-7 |
| N, mm | 1e-5 |
| mH/km, uF, dimensionless | 1e-9 |
| H | 1e-10 |
| kW | 1e-6 |

The CSV includes all twelve requested minimum checks and broader sensitivities. Units distinguish phase and line-to-line kV, metres and centimetres, H and mH/km, nF/km and total uF. Charging current uses total capacitance and RMS phase voltage. The script prints verification counts and checks the source SHA-256 before/after each run.

### Verification of the supplied workbook

The updated script completed successfully. All **297 comparisons matched** the corrected workbook's saved values within the stated tolerances. All nine figures were regenerated at 300 DPI.

| Selected result | Python and saved Excel result (rounded) |
|---|---:|
| Dry disruptive voltage | 278.4646 kV LL |
| Wet disruptive voltage | 212.9435 kV LL |
| Twin-bundle disruptive voltage | 394.8735 kV LL |
| Inductance reduction | 27.4620% |
| Capacitance increase | 36.1718% |
| Single charging current | 53.1048 A |
| Twin charging current | 72.3138 A |
| Baseline sag | 9.7340 m |
| Baseline clearance | 15.2660 m |
| Sag at 75°C | 11.7678 m |
| Wind pressure at 15 m/s | 137.8125 Pa |
| Hot + extra loading + wind clearance | 10.2903 m |

## Assumptions and limitations

GMD uses the workbook's literal phase spacings 6, 6 and 12 m. Air correction uses its 76 cmHg pressure, 273 temperature offset and 3.92 coefficient. These formula literals are documented in the script and guarded against source formula changes. Remaining selected inputs and sensitivity/scenario literals are read from their workbook cells. The standalone temperature block retains its local coefficient and 25-to-75°C definition.

Corona uses the workbook's empirical Peek-type equations, including the zero-loss condition below inception. Single-conductor GMR is 0.7788r; twin-bundle capacitance radius and inductance GMR remain distinct. Sag is `wL²/(8H)` with horizontal tension H. Suspended length is `L + 8d²/(3L)`, and thermal expansion is converted entirely into additional sag. This is not a conductor stress-strain, creep or temperature-dependent tension solution.

Combined hot cases share the horizontal tension back-calculated from the standalone 75°C sag; extra loading and wind do not trigger a new tension solver. Additional vertical load remains the existing selected percentage, without an invented ice model. Ground clearance uses level-ground geometry in all scenarios. Wind produces lateral displacement without reducing vertical clearance at fixed H.

This educational study is not certified or construction-ready design software. Calculated clearance is not a regulatory compliance determination. Resultant support tension is not a breaking-strength or safety-margin assessment.
