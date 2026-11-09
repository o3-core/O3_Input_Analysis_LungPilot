# o3_input_report_lungpilot

Summarise an O3 radiotherapy record set (JSON, lung pilot) into a Word (.docx) report, and optionally into a
results JSON. The tool runs entirely on your own computer: it makes no network calls, uploads nothing, and
only uses postal-code data (for patient-to-facility distances) that you have already installed locally.

The report contains a patient table, a curation check (missing data and outliers), a summary table, statistical
associations, distributions, one histogram per page, and fitted scatter plots. It contains no street addresses or birth dates;
postal codes appear only in the optional source-file table (`--analyze-sources`).

## Requirements

- Python 3.9 or later
- `pip install -r requirements.txt`

## Quick start

```
python o3_input_report_lungpilot.py data.json
```

For offline distances, install a GeoNames postal-code file once (see "Order of steps" below):

```
python o3_input_report_lungpilot.py data.json --analyze-sources
python o3_input_report_lungpilot.py --install-source US.zip
```

## Usage (output of `--help`)

```text
usage: o3_input_report_lungpilot.py [-h] [-o OUTPUT] [--analyze-sources]
                                    [--install-source FILE]
                                    [--country COUNTRY] [--units {mi,km}]
                                    [--fill-distances OUT.json] [--savewmods]
                                    [--results-json [FILE]] [--no-report]
                                    [--data-dir DATA_DIR]
                                    [json]

Create a Word summary report (and optional results JSON) from an O3 JSON record set, entirely offline.

positional arguments:
  json                  O3 record set .json

options:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        output .docx file, or a folder (default:
                        AnalysisSummary_YYYYMMDDHHMM.docx next to the input
                        JSON)
  --analyze-sources     list the pgeocode source file(s) needed to compute
                        distances offline, and whether they are installed.
                        With no -o, no report is written.
  --install-source FILE
                        install a downloaded GeoNames <CC>.zip/.txt into the
                        pgeocode cache and exit
  --country COUNTRY     default country for ambiguous postal codes (default
                        US)
  --units {mi,km}       distance units (default mi)
  --fill-distances OUT.json
                        write a copy of the JSON with RTCourse direct-distance
                        value/method/units filled in (only where the value is
                        blank/-9999), and exit
  --savewmods           save a copy of the input JSON with the distance fields
                        filled in (blank values only), named
                        <input>_YYYYMMDDHHMM.json in the input file's folder,
                        and exit
  --results-json [FILE]
                        also write the analysis results (dataset counts,
                        summary, associations, distributions, histograms,
                        analysis-plot fits) to JSON in report order, without
                        the patients/curation tables or figures. FILE or
                        folder; default AnalysisResults_YYYYMMDDHHMM.json next
                        to the input
  --no-report           do not write the Word report (use with --results-json)
  --data-dir DATA_DIR   pgeocode data folder (sets PGEOCODE_DATA_DIR)

WHAT IT DOES
  Reads an O3 record set (.json) and writes a Word report (AnalysisSummary_<YYYYMMDDHHMM>.docx) with a
  patient table, curation check, summary table, associations, distributions, histograms and fitted plots.
  Everything runs on this computer: no network calls, no uploads. pgeocode is used only when its postal-code
  data file is already installed on disk, and nothing is ever downloaded.

ORDER OF STEPS
  1. (One time) Find out which postal-code file(s) the data needs:
         python o3_input_report_lungpilot.py data.json --analyze-sources
     This lists the country file (for example US.zip), where to get it, where it must be installed and
     whether it is already present. It writes no report unless you also give -o.
  2. (One time, on a computer that may go online) download the file named in step 1 (GeoNames US.zip),
     copy it to this computer, then install it:
         python o3_input_report_lungpilot.py --install-source C:\downloads\US.zip
     Use --data-dir to install to, and read from, a different folder.
  3. (Optional) Save a copy of the JSON with the RTCourse distance fields filled in
     (value, method "Postal Code", units). Only blank or -9999 values are changed; the original is untouched:
         python o3_input_report_lungpilot.py data.json --savewmods
  4. Write the report (and optionally a JSON of the results):
         python o3_input_report_lungpilot.py data.json
         python o3_input_report_lungpilot.py data.json --results-json

OPTIONS IN DETAIL
  json                    The O3 record set to analyse. Required except with --install-source.
  -o, --output PATH       Where the Word report goes. A .docx file name is used exactly as given; a folder
                          gets AnalysisSummary_<YYYYMMDDHHMM>.docx inside it. Default: that name next to the
                          input JSON.
  --analyze-sources       Scan every postal-code field, work out the country, and list the pgeocode source
                          files needed for offline distances, with download URL, install command, install
                          path, present/missing, and any codes not found in an installed file. Adds the same
                          table to the report if a report is written.
  --install-source FILE   Install a downloaded GeoNames <CC>.zip or <CC>.txt into the pgeocode cache
                          (converted to the layout pgeocode reads), then exit.
  --country CC            Country for postal codes that could be several countries (a bare 5-digit code).
                          Default US. Codes with a US state in the address, or in Canadian/UK format, are
                          detected automatically.
  --units mi|km           Distance units for the report and for filled-in distances. Default mi.
  --savewmods             Save a copy of the input JSON, with the distance fields filled in, as
                          <input>_<YYYYMMDDHHMM>.json next to the input, then exit. Fields filled:
                          RTCourse_DistanceToPatientPrimaryResidenceDirect (value),
                          ...DirectMethod ("Postal Code") and RTCourse_DistanceToPatientDomicileDirectUnits.
                          Courses that already have a value are left alone. Distances are straight-line
                          between postal-code centroids, not street-address or driving distance.
  --fill-distances OUT    Same as --savewmods but you choose the output file. Refuses to overwrite the input.
  --results-json [FILE]   Also write the analysis results to JSON in the same order as the report
                          (dataset, summary, associations, distributions, histograms, analysis-plot fits).
                          Excludes the patients and curation tables and all figures. FILE or a folder;
                          default AnalysisResults_<YYYYMMDDHHMM>.json next to the input.
  --no-report             Do not write the Word report (use with --results-json for JSON only).
  --data-dir DIR          pgeocode data folder (sets PGEOCODE_DATA_DIR). Default ~/.cache/pgeocode.
  -h, --help              Show this guide.

EXAMPLES
  Check what postal-code data is needed, without writing a report:
      python o3_input_report_lungpilot.py data.json --analyze-sources
  Install the downloaded US postal-code file:
      python o3_input_report_lungpilot.py --install-source C:\downloads\US.zip
  Write the report next to the data file:
      python o3_input_report_lungpilot.py data.json
  Write the report to a folder, in kilometres, with a results JSON beside the input:
      python o3_input_report_lungpilot.py data.json -o C:\reports --units km --results-json
  Report plus the source-file table:
      python o3_input_report_lungpilot.py data.json --analyze-sources -o C:\reports
  Results JSON only:
      python o3_input_report_lungpilot.py data.json --results-json results.json --no-report
  Save a copy of the data with distances filled in:
      python o3_input_report_lungpilot.py data.json --savewmods

NOTES
  - Run with plain "python" (not "python -I"): python-docx may live in your user site-packages.
  - Without the postal-code file installed, the report is still written but distances are left out and
    the distance rows say so.
  - Statistics are exploratory and unadjusted; the report's Notes columns explain each test and caveat.
  - Requires: python-docx, matplotlib, scipy, numpy, pandas, pgeocode.
```

## Privacy

Do not commit patient data. The included `.gitignore` excludes `*.json` and generated reports. Check any data
file you analyse for fields such as dates of birth and street addresses before sharing it.

## Limitations

- Statistics are exploratory and unadjusted; many data sets will be small.
- Distances are straight-line between postal-code centroids, not street-address or driving distances.
- Non-US postal codes have not been tested.
- Per-patient fractions, dose, technique, age and distance use each patient's first RT course; the curation check flags patients with more than one.
- `--savewmods` / `--fill-distances` write a full copy of the input JSON, which still contains any date-of-birth or address fields. Treat it as PHI.

## License

Apache License 2.0. See [LICENSE](LICENSE).
