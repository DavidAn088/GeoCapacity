# GeoCapacity

**GeoCapacity** is a Streamlit-based geotechnical foundation capacity calculator developed by **David An**.

GeoCapacity has expanded from a pile-capacity calculator into a broader **foundation-capacity platform** covering both **deep foundations** and **shallow foundations**.

## Scope

### Deep foundation
- Bored and driven pile capacity checking
- Borehole-layer based calculation
- Supported sections: circular, square, hexagonal, hollow circular, hollow square, I-section, and custom geometry
- Editable equation library
- Capacity plots, calculation tables, and PDF reports

### Shallow foundation
- Strip and rectangular footing capacity checking
- Total stress and effective stress options
- Excel-style factor table and footing geometry sketch
- PDF report export

## Required files for deployment

Keep these files/folders in the GitHub repository root:

```text
streamlit_app.py
requirements.txt
README.md
templates/
samples/
```

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy with Streamlit Community Cloud

1. Create a GitHub repository named `GeoCapacity`.
2. Upload all files in this package.
3. Go to Streamlit Community Cloud.
4. Click **New app**.
5. Select the GitHub repository.
6. Set the main file path to:

```text
streamlit_app.py
```

7. Click **Deploy**.

## Note

Do not upload confidential borehole or project data to a public repository. Use sample/template CSV files for public demos.


## Scope expansion

GeoCapacity v1.0 includes both:

- **Deep Foundation**: bored and driven pile capacity based on borehole layers and editable equations.
- **Shallow Foundation**: strip and rectangular footing bearing capacity with total/effective stress analysis, eccentricity, contact stress, bearing factors, factor table, geometry sketch, and PDF export.


## Default interface note

The left sidebar theme is set to **Bright** by default in GeoCapacity v1.0.


## Run-label and shallow-equation update

GeoCapacity v1.0 now shows shallow foundation LaTeX equations in the Shallow Foundation results page, keeps run dropdown labels bold, and removes the repeated deep Equation Guide block from the bottom of Equation Lab.


## Homepage developer text removed

The homepage now shows only the GeoCapacity app name and subtitle. The version remains GeoCapacity v1.0.

## Calculation core note

GeoCapacity v1.0 uses the verified v9.10.2 deep-foundation calculation core for pile capacity. Other UI/branding/reporting parts are kept unchanged.

## Deep foundation Excel verification

The deep-foundation default equation `su2_t_m2` was corrected to match the uploaded Excel file: blank SPT values now produce blank/NaN `su2`, not zero. This makes Excel `AVERAGE(su1, su2)` behavior match exactly.
A verification summary is included in `deep_excel_verification_summary.csv`.


## Latest upgrade

- Deep and shallow result pages now show the Equation Lab option used for the run.
- Result pages now show the LaTeX formula library used for the run, linked from the active Equation Lab table.
- Shallow custom equation libraries now affect shallow calculation, sensitivity, and shallow design charts.
- Strip footing remains protected: Fcs = Fqs = Fγs = 1.0.
- Shallow footing CSV now supports eccentricity_mode:
  - MOMENT: ex = Mty / Qt and ey = Mtx / Qt.
  - DIRECT: use ex_direct_m and ey_direct_m.
- Shallow Data Format page includes a brief shape, beta angle, and eccentricity guide.


## Eccentricity input simplification

Shallow footing input now uses only `Ox_m` and `Oy_m` with `eccentricity_mode`.

- `eccentricity_mode = MOMENT`: `Ox_m` and `Oy_m` are load offsets; the app calculates `ex = Mty/Qt` and `ey = Mtx/Qt`.
- `eccentricity_mode = DIRECT`: `Ox_m` and `Oy_m` are final eccentricities; the app uses `ex = Ox_m` and `ey = Oy_m`.
- For strip footing, `Oy`, `ey`, `Mcx`, and `Hy` are kept zero, and all shape factors are forced to 1.
