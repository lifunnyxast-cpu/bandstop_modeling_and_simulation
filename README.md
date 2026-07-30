================================================================================
STOP Bandstop FSS Dataset
================================================================================
Date: 2026-07-24

This folder contains the complete bandstop (stop-type) FSS dataset and simulation pipeline accompanying the paper:

"Spatial Dispersion Compensated Equivalent Circuit Model for Wide-Angle Frequency-Selective Surfaces"

The accompanying **Spatial Dispersion-Compensated Equivalent Circuit Model Extraction Tool** (`SDC-ECM.exe`) is available for download from this release.

================================================================================
FILE INVENTORY
================================================================================

1. Index_L_C_alpha.csv  (48.9 KB, 506 rows + header)
   ───────────────────────────────────────────
   Primary dataset: 506 screened bandstop topologies.
   Each row = one topology at normal incidence (0 deg), described by the
   five-dimensional SDC-ECM feature vector.

   Columns:
     Unique_ID         — topology identifier (matches raw S-parameter filename)
     f0_TE_GHz         — TE resonant frequency at 0 deg (GHz)
     C0_TE_fF          — TE equivalent capacitance at 0 deg (fF)
     L0_TE_nH          — TE equivalent inductance at 0 deg (nH)
     R0_TE_ohm         — TE equivalent series resistance at 0 deg (ohm)
     alpha_TE_eff      — TE spatial-dispersion compensation factor (dimensionless)
     f0_TM_GHz
     C0_TM_fF
     L0_TM_nH
     R0_TM_ohm
     alpha_TM_eff      — TM compensation factor

   The angle-dependent capacitance follows the closed-form SDC-ECM:
     C(theta) = C0 * (1 - sin^2(theta) / alpha_eff)
     f_r(theta) = f0 / sqrt(1 - sin^2(theta) / alpha_eff)

   Calibration anchor angles: 0 deg (C0, L0) and 60 deg (alpha_eff).
   Validated angular range: 0–60 deg.
   80 deg data are available in Index_LUT.csv for extrapolation analysis.


2. Index_LUT.csv  (328.8 KB, 3013 rows + header)
   ────────────────────────────────────────────
   Multi-angle Look-Up Table: independently extracted Foster LC parameters
   at every simulated angle for 608 topologies.

   Columns:
     Unique_ID         — row identifier (TopologyName + _Ang{deg})
     Topology_Name     — links to raw S-parameter file and Index_L_C_alpha.csv
     Angle_deg         — incidence angle: 0, 20, 40, 60, 80 deg
     f_TE_GHz          — TE resonant frequency at this angle (GHz)
     L_TE_nH           — TE Foster-extracted inductance at this angle (nH)
     C_TE_fF           — TE Foster-extracted capacitance at this angle (fF)
     R_TE_ohm          — TE Foster-extracted resistance at this angle (ohm)
     f_TM_GHz, L_TM_nH, C_TM_fF, R_TM_ohm  — same for TM polarization

   NOTE: In this LUT, L, C, R are extracted independently per angle
   (no constant-L assumption). The SDC-ECM constant-L approximation
   is validated by the small variance of L across angles observed here.

   Angle coverage: 608 topologies × 5 angles = 3040 possible;
   3013 rows = 27 missing entries where Foster extraction failed
   (typically weak resonance at 80 deg).


3. cst_sweep_data.csv  (657.9 KB, 2100 rows + header)
   ────────────────────────────────────────────
   Parameter-sweep configuration table for batch CST simulation.
   Each row defines one full-wave simulation case.

   Key columns:
     Source_ID          — topology family identifier (.png image label)
     Sym_Type           — rotational symmetry order (1, 2, 3, 4, 6, 12, ...)
     Var_Radius_mm       — period R (mm), hexagonal unit-cell circumradius
     Var_Width_mm         — line width / strip width (mm)
     Path_0, Path_1, ... — geometry trace coordinates (x,y pairs)
     Aux_Hex_Apothem     — derived hexagon apothem
     Width_Ratio         — ww/x0 ratio (horizontal step ratio)

   This sweep table covers 2100 geometry variants across multiple
   symmetry families, from which the 506 valid topologies were screened.


4. Modeling_for_TO_old.py  (95.3 KB)
   ────────────────────────────────────
   CST COM automation module.

   Classes:
     CSTInterface      — COM connection to CST Studio Suite
     SimulationBuilder — parametric geometry construction (build_geometry,
                         create_materials, boolean operations)
     SweepManager      — frequency-domain solver execution
     ResultExtractor   — S-parameter export (Magnitude_dB / Phase_deg)

   Key method: SimulationBuilder.build(row, points)
     row    — one row of cst_sweep_data.csv as a pandas Series
     points — Path_* columns reshaped into (N,2) coordinate array

   Simulation setup:
     - Solver: Frequency Domain (F-solver)
     - Floquet ports, unit-cell boundary
     - Substrate: polyimide (eps_r=3.2, tan_d=0.002, h=0.025 mm)
     - Metal: copper (Ohmic sheet, t=0.018 mm, 0-thickness model)
     - Frequency: 0–20 GHz
     - Angles: 0, 20, 40, 60, 80 deg
     - Polarization: TE and TM

   NOTE: This module imports from the newer Modeling_for_TO.py in the
   parent directory. Place both files in the same directory when running.


5. run_cst_pld.py  (6.2 KB)
   ─────────────────────────
   Batch automation driver script.

   Usage:
     python run_cst_pld.py

   Workflow:
     1. Reads cst_sweep_data.csv
     2. Iterates all 2100 rows sequentially
     3. For each row:
        a. Extracts Path_* coordinates
        b. Builds 3D geometry via SimulationBuilder.build()
        c. Runs CST frequency sweep via SweepManager.run_sweep_task()
        d. Exports S-parameters (Magnitude_dB, Phase_deg) at 0 and 80 deg
        e. Merges results into S_Parameter_Results/ID{src}_{sym}_R{r}_W{w}.csv
     4. Skips already-existing output files (checkpoint-resume friendly)

   Requirements:
     pip install pywin32 numpy pandas
     Microsoft CST Studio Suite (installed, with COM license)


================================================================================
FILTERING PIPELINE (from raw CST to final 506)
================================================================================

  CST raw simulations:          747 files (S_Parameter_Results_for_stop/)
    ├── 48 with no detectable S21 notch (S21_min > -3 dB) → excluded
    ├── 20 with weak resonance (-3 to -10 dB) → excluded
    └── 172 with strong resonance (< -10 dB) but Foster L/C/alpha
          extraction failed (numerically unstable) → excluded
  Raw Foster extraction:        593 entries with valid resonance
    └── TE/TM symmetry filter:  |f0_TE - f0_TM| < 0.10 GHz → 535 entries
         └── Alpha outlier filter: |alpha_TE| <= 500 AND |alpha_TM| <= 500
             → FINAL: 506 bandstop topologies
