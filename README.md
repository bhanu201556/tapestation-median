# Agilent Tapestation-median
Adds a median fragment-size column to TapeStation / Fragment Analyzer sample tables
# add_median — TapeStation Median Fragment-Size Tool

A small Python script that reads a TapeStation / Fragment Analyzer run (the
electropherogram trace + the sample table), computes the **median fragment size**
for every sample, and writes a clean Excel file with an extra
`Median Size [bp]` column.

The median is taken as the 50%-cumulative-area point of each sample's size
distribution, calibrated against the 10-peak ladder.

---

## Why median?

For library QC, the **median** is often a more robust summary of fragment size
than the peak/mode, because it is less sensitive to a single dominant peak or to
tailing in the distribution. This tool gives you median plus (optionally) mode,
mean, Q25 and Q75 so you can sanity-check the distribution shape.

---

## Requirements

- Python 3.8+
- The following packages:

```bash
pip install pandas numpy scipy openpyxl
```

---

## Input files

You need two CSVs exported from your TapeStation software:

1. The **electropherogram** export (the raw trace) — filename contains
   `Electropherogram`.
2. The **sample table** — filename contains `SampleTable` (optional; if missing,
   the script builds a table from scratch).

The script auto-detects the ladder lane and the 25 bp / 1500 bp markers.

---

## How to run

**Option A — run it inside the folder that has your CSVs:**

```bash
python add_median.py
```

**Option B — point it at a folder:**

```bash
python add_median.py "C:/path/to/exported/csvs"
```

**Option C — pass the two files explicitly:**

```bash
python add_median.py run_Electropherogram.csv run_SampleTable.csv
```

**Want extra columns** (mode / mean / Q25 / Q75)? Add `--full`:

```bash
python add_median.py --full
```

---

## Output

A styled Excel file is written next to your inputs:

```
<run-prefix>_withMedian.xlsx
```

It is your original sample table plus a highlighted `Median Size [bp]` column
(and the extra statistics columns if you used `--full`).

---

## Notes / limits

- The median window is currently fixed to **200–1200 bp** (`LOWER_BP` / `UPPER_BP`
  near the top of the script). Adjust these if your libraries sit outside that
  range.
- Calibration expects the standard 10-peak ladder
  (25, 50, 100, 200, 300, 400, 500, 700, 1000, 1500 bp). If your ladder differs,
  edit `LADDER_SIZES`.

---

## License

Released under the MIT License — see the `LICENSE` file. You are free to use,
modify and share it; attribution is appreciated.

---

## Author

Created by **Bhanu Pratap Gurjar, PhD**,
 Bioinformatician, NGS Core Facility — National Institute of Immunology, New Delhi
📧 bhanu@nii.a.in

Feel free to open an issue or reach out if something breaks or you want a feature
added.
