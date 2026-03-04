# ForamEcoQS

**ForamEcoQS** (Foraminiferal Ecological Quality Status) is a comprehensive software tool designed to calculate biotic indices based on benthic foraminifera for assessing the ecological quality of marine environments.

This tool integrates multiple indices providing a unified platform for advanced ecological analyses, data management, visualization, and statistical agreement analysis.

---

## Table of Contents

1. [Key Features](#key-features)
2. [Biotic Indices](#biotic-indices)
3. [Data Management](#data-management)
4. [Visualization and Plotting](#visualization-and-plotting)
5. [EQS Agreement Analysis](#eqs-agreement-analysis)
6. [Databank Management](#databank-management)
7. [Geographic Areas Database](#geographic-areas-database)
8. [Installation and Usage](#installation-and-usage)
9. [References](#references)
10. [Authors and Contacts](#authors-and-contacts)
11. [Citation](#citation)
12. [License](#license)

---

## Key Features

### Multi-Index Calculations

ForamEcoQS calculates a comprehensive suite of biotic indices for ecological quality assessment:

**Sensitivity-Based Indices:**
- **Foram-AMBI** - Foram-AMBI sensitivity index (Alve et al. 2016; Jorissen et al. 2018)
- **Foram-M-AMBI** - Multivariate Foram-AMBI combining Foram-AMBI, Shannon diversity, and species richness
- **FSI** - Foram Stress Index (Dimiza et al. 2016)
- **TSI-Med** - Tolerant Species Index for the Mediterranean with grain-size correction (Barras et al. 2014)
- **NQIf** - Norwegian Quality Index for foraminifera (Alve et al. 2019)
- **BENTIX** - Benthic Index using simplified ecological groups (Simboura and Zenetos 2002)
- **BQI** - Benthic Quality Index adapted for foraminifera (Rosenberg et al. 2004)
- **FIEI** - Foraminiferal Index of Environmental Impact (Mojtahid et al. 2006)
- **FoRAM Index** - For tropical/subtropical coral reef environments (Hallock et al. 2003; Prazeres et al. 2020)

**Diversity Indices:**
- **exp(H'bc)** - Effective number of species with bias correction (Chao and Shen 2003; Bouchet et al. 2012)
- **H'log2** - Shannon-Wiener Index (base 2)
- **H'ln** - Shannon-Wiener Index (natural logarithm)
- **Simpson (1-D)** - Simpson's diversity index
- **Pielou's J** - Evenness index
- **ES100** - Hurlbert's expected species for 100 individuals

**Additional Metrics:**
- Species Richness (S)
- Total Abundance (N)
- Ecological Group percentages (Eco1-5)
- FoRAM functional group percentages (Symbiont-bearing, Stress-tolerant, Heterotrophic)

### Calculation & Workflow Controls

- **Selective Index Calculation:** Choose which indices to compute with dedicated dialogs and default presets tuned to available datasets.
- **Threshold System Selection:** Switch Foram-AMBI classification between Borja (2003) and Parent (2021) systems and view EQS class labels alongside EQR values.
- **Interactive Sample Pipeline:** Add, rename, or remove samples, undo recent edits, and clean/normalize imported matrices before calculating indices.
- **Databank Templates:** Build ready-to-use spreadsheets populated from reference databanks (Foram-AMBI, FSI, TSI-Med) or user-defined species lists to streamline data entry.
- **EQS Agreement Suite:** Generate full agreement workups—pairwise K matrix, confusion matrices, and heatmap dashboards—with exportable summaries for reporting.

### Configurable Threshold Systems

- **Foram-AMBI thresholds:** Choice between Borja (2003) and Parent (2021) classification systems
- **Index-specific EQS classifications:** Each index provides Ecological Quality Status (High, Good, Moderate, Poor, Bad)

---

## Data Management

### Databank Exploration & Curation

- **Read-Only Viewer:** Browse ecological group databanks with search, ecological group filters, quick statistics, and export to CSV/Excel.
- **Foram-AMBI Databank Manager:** Import/export EG1-EG5 classifications, edit species assignments, reset to published baselines, or maintain custom overrides that persist across sessions.
- **FSI Databank Manager:** Manage sensitive/tolerant classifications with CSV/Excel import/export, filtering, and toggle tools for batch edits and custom databank activation.
- **Geographic Areas Database:** Maintain geographic/environmental reference records with search/filter tools and CSV/Excel interchange for bibliographic tracking.
- **User Custom Lists:** Create per-index species lists (e.g., Foram-AMBI, FSI, TSI-Med, FoRAM Index) stored in the `user_lists` directory for reuse in templates and calculations.

### File Operations

- **Open:** Load data from Excel (.xls, .xlsx) or CSV files
- **Save Samples:** Export current dataset
- **Load Indices for Graphs:** Load previously exported indices data for visualization without recalculation
- **Create Templates:** Generate species templates from selected databanks for data entry

### Data Editing

- **New Sample:** Add new sample columns
- **Remove Selected Sample:** Delete samples from dataset
- **Rename Selected Sample:** Modify sample names
- **Undo:** Revert recent changes
- **Clean and Normalize:** Data preprocessing and normalization
- **Species Name Matching:** Resolve names against reference databanks with synonym handling and confidence scoring to minimize manual corrections.

### Template Creation

- **Create Template from Databank:** Generate Excel templates based on Foram-AMBI databanks
- **Create FSI Template:** Generate templates for FSI analysis
- **Create TSI-Med Template:** Generate templates for TSI-Med analysis

### Specialized Databank Support

- **Custom Databank Detection:** Automatic fallback between bundled and user-supplied databanks for FSI and TSI-Med calculations with at-a-glance source summaries.
- **Reference Selection:** Choose among multiple published Foram-AMBI databanks (Jorissen, Alve, Bouchet variants, O'Malley) when building templates or matching species.

---

## Visualization and Plotting

### Plot Types

ForamEcoQS provides comprehensive visualization options:

- **Bar Chart:** Individual sample bar plots for selected indices
- **Grouped Bar:** Side-by-side comparison of indices per sample
- **Line Plot:** Trend visualization across samples
- **Box Plot:** Statistical distribution of indices across samples
- **Scatter Plot:** Index values plotted by sample
- **Heatmap:** Normalized indices values displayed as a color-coded matrix
- **K Matrix Heatmap:** Color-coded Cohen's Kappa matrix ("K matrix") summarizing pairwise agreement between indices
- **Eco Groups Distribution:** Stacked area and box plots of ecological group percentages
- **Composite Panel:** Multi-panel dashboard with multiple plot types
- **EQS Classification:** Ecological Quality Status visualization

### Plot Customization

- **DPI Selection:** 150, 300, or 600 DPI for export quality
- **Font Options:** Arial, Calibri, Times New Roman, Segoe UI
- **Color Schemes:** Default, Ocean, Earth, Grayscale, Colorblind-friendly
- **Grid Toggle:** Show/hide grid lines
- **Sample Selection:** Choose specific samples to include in plots
- **Index Selection:** Select which indices to display

### Export Options

- Export individual plots as PNG images
- Export all plots in batch
- Export results to Excel with formatting
- Export composite dashboards

---

## EQS Agreement Analysis

A comprehensive module for analyzing agreement between different biotic indices in their ecological quality classifications.

### Cohen's Kappa Analysis

- Pairwise Cohen's Kappa calculation between all indices with EQS classifications
- Color-coded Kappa matrix (K matrix) visualization
- Interpretation guide (Poor to Almost Perfect agreement)
- Export to Excel with conditional formatting

### Confusion Matrix

- Generate pairwise confusion matrices between any two indices
- Visual highlighting of agreement (diagonal cells)
- Agreement statistics including:
  - Total samples analyzed
  - Exact agreement count and percentage
  - Cohen's Kappa value with interpretation

### Agreement Heatmap

- Visual heatmap of Cohen's Kappa values between all index pairs (K matrix heatmap)
- Color gradient from poor (red) to perfect (blue) agreement
- Interactive visualization using OxyPlot

### Summary Statistics

- EQS class distribution by index with visual bars
- Ranked pairwise Kappa values
- Mean Kappa across all comparisons
- Best and worst agreeing index pairs
- Consensus analysis:
  - Full consensus (all indices agree)
  - Majority consensus (more than 50% agree)
  - No clear consensus

### Kappa Interpretation Scale

| Kappa Range | Interpretation |
|-------------|----------------|
| 0.81 - 1.00 | Almost perfect agreement |
| 0.61 - 0.80 | Substantial agreement |
| 0.41 - 0.60 | Moderate agreement |
| 0.21 - 0.40 | Fair agreement |
| 0.00 - 0.20 | Slight agreement |
| < 0.00 | Poor agreement (worse than chance) |

---

## Databank Management

### Foram-AMBI Databank Manager

Manage species ecological group assignments for Foram-AMBI calculations:

- **Available Databanks:**
  - Jorissen (Mediterranean, Foram-AMBI)
  - Alve (NE Atlantic/Arctic, Foram-AMBI)
  - Bouchet Mediterranean (Foram-AMBI)
  - Bouchet Atlantic (Foram-AMBI)
  - Bouchet South Atlantic (Foram-AMBI)
  - O'Malley (Gulf of Mexico, Foram-AMBI)

- **Features:**
  - View and edit ecological group assignments (EG1-EG5)
  - Search and filter species
  - Add new species
  - Import from CSV or Excel
  - Export databanks
  - Compare databanks

### FSI Databank Manager

Manage species classifications for the Foram Stress Index:

- Edit Sensitive (S) and Tolerant (T) classifications
- Import/export functionality
- Add new species with classification
- Toggle classifications

### User Custom Lists Manager

Create and manage custom species lists for calculations:

- Create lists for specific indices (Foram-AMBI, FSI, TSI-Med, FoRAM Index, BQI, BENTIX, NQIf, Custom)
- Import species from CSV or Excel
- Export custom lists
- Specify which index each list is used for

---

## Geographic Areas Database

A reference database of geographic regions with associated databanks and applicable indices:

| Region | Environment | Applicable Indices | Reference |
|--------|-------------|-------------------|-----------|
| Mediterranean Sea | Marine Shelf | Foram-AMBI, TSI-Med, FSI | Jorissen et al. (2018) |
| Mediterranean Lagoons | Transitional Waters | Foram-AMBI, FSI | Bouchet et al. (2021) |
| Atlantic European Coast | Marine Shelf | Foram-AMBI | Bouchet et al. (2012) |
| North Sea | Marine Shelf | Foram-AMBI, NQIf | Alve et al. (2016) |
| Gulf of Mexico | Marine Shelf | Foram-AMBI | O'Malley et al. (2021) |
| South Atlantic | Marine Shelf | Foram-AMBI | Bouchet et al. (2025) |
| Norwegian Fjords | Fjord | Foram-AMBI, NQIf | Alve et al. (2019) |
| Tropical Coral Reefs | Coral Reef | FoRAM Index | Hallock et al. (2003) |
| Deep-Sea Bathyal | Deep-Sea | Foram-AMBI | Jorissen et al. (1995) |
| Adriatic Sea | Marine Shelf | TSI-Med, Foram-AMBI | Barras et al. (2014) |

### Features

- View geographic areas with associated references and DOIs
- Filter and search regions
- Export database to CSV or Excel
- Add custom geographic areas

---

## Installation and Usage

### System Requirements

- Windows operating system
- .NET Framework 4.7.2 or later

### Installation

1. Download the latest release from the repository
2. Extract the archive to your preferred location
3. Ensure the following databank files are in the same folder as the executable:
   - `fsi_databank.csv` - For FSI calculation
   - `tsimed_databank.csv` - For TSI-Med calculation
   - `foram_index_databank.csv` - For FoRAM Index calculation
   - `geographic_areas_databank.csv` - Geographic areas reference

### Basic Workflow

1. **Start the application:** Run `ForamEcoQS.exe`
2. **Load data:** Use File > Open to load your Excel or CSV file with species abundance data
3. **Select databank:** Choose an appropriate reference databank from the dropdown
4. **Validate taxa:** Click "Compare" to match your species against the databank
5. **Calculate indices:** Click "Calculate Indices" or use the Advanced Indices menu
6. **Visualize results:** Use the Plot Options tab to create visualizations
7. **Analyze agreement:** Use Plots > EQS Agreement Analysis to compare index classifications
8. **Export results:** Save results to Excel or export plots as images

### Data Format

Input data should be organized with:
- First column: Species names
- Subsequent columns: Sample abundances (counts or percentages)
- First row: Sample names

---

## References

### Primary Index References

- **Alve E. et al. (2016)** Foram-AMBI: A sensitivity index based on benthic foraminiferal faunas from NE Atlantic and Arctic fjords, continental shelves and slopes. *Marine Micropaleontology 122*:1-12. https://doi.org/10.1016/j.marmicro.2015.11.001

- **Alve E. et al. (2019)** Intercalibration of benthic foraminiferal and macrofaunal biotic indices: An example from the Norwegian Skagerrak coast (NE North Sea). *Ecological Indicators 96*:107-115. https://doi.org/10.1016/j.ecolind.2018.08.037

- **Barras C. et al. (2014)** Live benthic foraminiferal faunas from the French Mediterranean coast: Towards a new biotic index of environmental quality. *Ecological Indicators 36*:719-743. https://doi.org/10.1016/j.ecolind.2013.09.028

- **Bouchet V.M.P. et al. (2012)** Benthic foraminifera provide a promising tool for ecological quality assessment of marine waters. *Ecological Indicators 23*:66-75. https://doi.org/10.1016/j.ecolind.2012.03.011

- **Bouchet V.M.P. et al. (2021)** Benthic foraminifera to assess Ecological Quality Status of Italian transitional waters. *Marine Pollution Bulletin 163*:112071. https://doi.org/10.1016/j.marpolbul.2021.112071

- **Bouchet V.M.P. et al. (2025)** Foram-AMBI for South Atlantic coastal and transitional waters. *Journal of Micropalaeontology 44*:237. https://doi.org/10.5194/jm-44-237-2025

- **Chao A. and Shen T.J. (2003)** Nonparametric estimation of Shannon's index of diversity when there are unseen species in sample. *Environmental and Ecological Statistics 10*:429-443.

- **Dimiza M.D. et al. (2016)** The Foram Stress Index: A new tool for environmental assessment of soft-bottom environments using benthic foraminifera. A case study from the Saronikos Gulf, Greece, Eastern Mediterranean. *Ecological Indicators 60*:611-621.

- **Hallock P. et al. (2003)** Foraminifera as bioindicators in coral reef assessment and monitoring: The FORAM Index. *Environmental Monitoring and Assessment 81*:221-238. https://doi.org/10.1023/A:1021337310386

- **Jorissen F.J. et al. (2018)** Developing Foram-AMBI for biomonitoring in the Mediterranean: Species assignments to ecological categories. *Marine Micropaleontology 140*:33-45. https://doi.org/10.1016/j.marmicro.2017.12.006

- **Jorissen F.J. et al. (1995)** A conceptual model explaining benthic foraminiferal microhabitats. *Marine Micropaleontology 26*:3-15. https://doi.org/10.1016/0377-8398(95)00047-X

- **Mojtahid M. et al. (2006)** Benthic foraminifera as bio-indicators of drill cutting disposal in tropical east Atlantic outer shelf environments. *Marine Micropaleontology 61*:58-75.

- **Muxika I., Borja A., Bald J. (2007)** Using historical data, expert judgement and multivariate analysis in assessing reference conditions and benthic ecological status, according to the European Water Framework Directive. *Marine Pollution Bulletin 55*(1-6):16-29. https://doi.org/10.1016/j.marpolbul.2006.05.025

- **O'Malley B.J. et al. (2021)** Development of a benthic foraminifera based marine biotic index (Foram-AMBI) for the Gulf of Mexico: A decision support tool. *Ecological Indicators 120*:107049. https://doi.org/10.1016/j.ecolind.2020.107049

- **Prazeres M., Martinez-Colon M., Hallock P. (2020)** Foraminifera as bioindicators of water quality: The FoRAM Index revisited. *Environmental Pollution 257*:113612. https://doi.org/10.1016/j.envpol.2019.113612

- **Rosenberg R. et al. (2004)** Marine quality assessment by use of benthic species-abundance distributions: a proposed new protocol within the European Union Water Framework Directive. *Marine Pollution Bulletin 49*(9-10):728-739. https://doi.org/10.1016/j.marpolbul.2004.05.013

- **Simboura N. and Zenetos A. (2002)** Benthic indicators to use in Ecological Quality classification of Mediterranean soft bottom marine ecosystems, including a new Biotic Index. *Mediterranean Marine Science 3/2*:77-111. https://doi.org/10.12681/mms.266

---

## Index Classification Details

### Foram-AMBI Classification

Based on five ecological groups (EG1-EG5) weighted by sensitivity to organic enrichment.

**Borja (2003) thresholds:**

| Foram-AMBI | EQS Class |
|------------|-----------|
| 0.0 - 1.2 | High |
| 1.2 - 3.3 | Good |
| 3.3 - 4.3 | Moderate |
| 4.3 - 5.5 | Poor |
| > 5.5 | Bad |

**Parent (2021) thresholds:**

| Foram-AMBI | EQS Class |
|------------|-----------|
| 0.0 - 1.4 | High |
| 1.4 - 2.4 | Good |
| 2.4 - 3.4 | Moderate |
| 3.4 - 4.4 | Poor |
| > 4.4 | Bad |

### FSI Classification (Dimiza et al. 2016)

Formula: FSI = (10 x Sensitive + Tolerant) / (Sensitive + Tolerant)

| FSI | EQS Class |
|-----|-----------|
| >= 9 | High |
| 6 - 9 | Good |
| 4 - 6 | Moderate |
| 2 - 4 | Poor |
| < 2 | Bad |

### TSI-Med Classification (Barras et al. 2014)

| TSI-Med | EQS Class |
|---------|-----------|
| 0 - 4 | High |
| 4 - 16 | Good |
| 16 - 36 | Moderate |
| 36 - 64 | Poor |
| > 64 | Bad |

### NQIf Classification

| NQIf | EQS Class |
|------|-----------|
| >= 0.82 | High |
| 0.63 - 0.82 | Good |
| 0.49 - 0.63 | Moderate |
| 0.31 - 0.49 | Poor |
| < 0.31 | Bad |

### exp(H'bc) Classification (Bouchet et al. 2012)

| exp(H'bc) | EQS Class |
|-----------|-----------|
| >= 18 | High |
| 11 - 18 | Good |
| 6 - 11 | Moderate |
| 3 - 6 | Poor |
| < 3 | Bad |

### BENTIX Classification (Simboura and Zenetos 2002)

Formula: BENTIX = (6 x %GS + 2 x %GT) / 100
- GS = EG1 + EG2 (Sensitive group)
- GT = EG3 + EG4 + EG5 (Tolerant group)

| BENTIX | EQS Class |
|--------|-----------|
| >= 4.5 | High |
| 3.5 - 4.5 | Good |
| 2.5 - 3.5 | Moderate |
| 2.0 - 2.5 | Poor |
| < 2.0 | Bad |

### BQI Classification (Rosenberg et al. 2004, adapted)

Formula: BQI = log10(S + 1) x mean_sensitivity

Sensitivity values: EG1=15, EG2=12, EG3=8, EG4=4, EG5=1

| BQI | EQS Class |
|-----|-----------|
| >= 12 | High |
| 8 - 12 | Good |
| 5 - 8 | Moderate |
| 2 - 5 | Poor |
| < 2 | Bad |

### Foram-M-AMBI Classification

Multivariate index combining Foram-AMBI, Shannon H', and Species Richness.

| EQR | EQS Class |
|-----|-----------|
| >= 0.82 | High |
| 0.62 - 0.82 | Good |
| 0.41 - 0.62 | Moderate |
| 0.20 - 0.41 | Poor |
| < 0.20 | Bad |

### FoRAM Index (Hallock et al. 2003)

Formula: FI = (10 x Ps) + Po + (2 x Ph)
- Ps = proportion of symbiont-bearing foraminifera
- Po = proportion of stress-tolerant foraminifera
- Ph = proportion of other heterotrophic foraminifera

| FoRAM Index | Interpretation |
|-------------|----------------|
| > 4 | Suitable for coral growth |
| 2 - 4 | Marginal conditions |
| < 2 | Unsuitable for coral growth |

---

## Authors and Contacts

- **Matteo Mangiagalli**
  - University of Urbino Carlo Bo, Italy
  - Email: m.mangiagalli@campus.uniurb.it

- **Fabrizio Frontalini**
  - University of Urbino Carlo Bo, Italy
  - Email: fabrizio.frontalini@uniurb.it

- **Carla Cristallo**
  - University of Urbino Carlo Bo, Italy
  - Email: c.cristallo1@campus.uniurb.it

- **Fabio Francescangeli**
  - University of Fribourg, Switzerland
  - Email: fabio.francescangeli@unifr.ch

---

## Citation

If you use this software in your publications, please cite it as follows:

> Mangiagalli M., Frontalini F., Cristallo C., Francescangeli F. (2024). ForamEcoQS - Foraminiferal Ecological Quality Status. Software version 1.0.

---

## Used Libraries

- **ClosedXML** - Library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files
- **ExcelDataReader** - Lightweight and fast library for reading Microsoft Excel files
- **OxyPlot.WindowsForms** - Cross-platform plotting library for .NET

---

## License

This project is distributed under the MIT License. See the `LICENSE.txt` file for details.
