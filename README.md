# Multimodal CBAM Transportation Optimization

Cost-optimal routing of Turkish steel/cement/fertiliser/aluminium/hydrogen exports to European demand nodes under the EU Carbon Border Adjustment Mechanism (CBAM), solved with two complementary approaches:

- **MILP** — exact solver via Gurobi (`cbam_multimodal_milp_gurobi.py`)
- **ALNS** — Adaptive Large Neighbourhood Search metaheuristic (`alns_cbam.py`)

---

## Problem Description

Orders originate from Turkish port/industrial source nodes and must be delivered to European demand cities by a deadline. Transportation can use sea, rail, road, or air legs through intermediate transit nodes. The objective minimises total cost:

```
min  Σ (transport cost + emission cost + handling cost + CBAM certificate cost)
```

**CBAM cost** is calculated per EU Regulation 2023/956. The phased factor φ_r scales up from 2.5 % (2026) to 100 % (2034):

| Year | φ_r |
|------|-----|
| 2026 | 2.5 % |
| 2027 | 5.0 % |
| 2028 | 10.0 % |
| 2030 | 48.5 % |
| 2034 | 100.0 % |

Only imports above the de minimis threshold (50 t/year) are subject to CBAM certificates.

---

## Repository Structure

```
multimodal-cbam-alns/
├── alns_cbam.py                    # ALNS metaheuristic solver (v13)
├── cbam_multimodal_milp_gurobi.py  # MILP exact solver (v6, requires Gurobi)
└── 50node_150order_inputt.xlsx     # Benchmark instance (50 nodes, 150 orders)
└── 30node_20order_input.xlsx     # Benchmark instance (30 nodes, 20 orders)
└── 20node_12order_input.xlsx     # Benchmark instance (20 nodes, 12 orders)
```

---

## Input File Format

Both solvers read the same `.xlsx` workbook. The file must contain the following sheets:

| Sheet | Key Columns | Description |
|-------|-------------|-------------|
| `Nodes` | `node_id`, `node_type` (`source`/`trans`/`demand`), `lat`, `lon`, `handling_day`, `handling_cost` | All network nodes |
| `Sources` | `source_node`, `capacity_ton`, `paid_carbon_price_per_tco2` | Supply capacities and pre-paid carbon prices |
| `Orders` | `order_id`, `destination_node`, `tons`, `deadline_date`, `product_type`, `embedded_emission_tco2_per_ton`, `covered_emission_tco2_per_ton`, `benchmark_emission_tco2_per_ton` | Demand orders |
| `Modes` | `mode`, `cost_per_tkm`, `speed_km_day`, `emission_kg_per_tkm`, `arc_capacity` | Transport modes (`sea`, `rail`, `road`, `air`) |
| `Arcs` | `from_node`, `to_node`, `mode` | Allowed directed connections |
| `Parameters` | `parameter`, `value` | Solver settings including `planning_year` (2026–2034) |
| `CBAM_Prices` | `cbam_period` (ISO week, e.g. `2027-W03`), `price_per_tco2` | Weekly CBAM certificate prices |
| `Mode_Schedules` | `mode`, `schedule_type`, `interval_days`, `allowed_weekday` | Departure frequency constraints |

The provided benchmark instance (`50node_150order_inputt.xlsx`) contains 50 nodes (12 Turkish source ports, 18 European transit hubs, 20 European demand cities) and 150 orders.

---

## Installation

```bash
pip install pandas openpyxl
```

The MILP solver additionally requires a valid **Gurobi licence**:

```bash
pip install gurobipy
```

---

## Usage

### ALNS Solver

```bash
python alns_cbam.py --input 50node_150order_inputt.xlsx
```

Key options:

| Argument | Default | Description |
|----------|---------|-------------|
| `--input` | *(required)* | Path to input `.xlsx` file |
| `--max_iter` | 2000 | Maximum ALNS iterations |
| `--time_limit` | 300.0 | Wall-clock time limit (seconds) |
| `--alpha` | 0.9975 | Simulated annealing cooling rate |
| `--seed` | random | Random seed for reproducibility |
| `--out_prefix` | `alns_v13` | Output file prefix |
| `--soft_arc_cap` | off | Relax arc capacity to soft constraint |
| `--arc_overflow_penalty` | 5000.0 | Penalty per ton of arc overflow (soft mode) |
| `--deadline_buffer` | 0 | Extra buffer days added to all deadlines |
| `--no_xlsx` | off | Skip Excel report generation |
| `--silent` | off | Suppress console output |

Outputs: `<out_prefix>_report_<timestamp>.xlsx` and `<out_prefix>_log_<timestamp>.xlsx`

### MILP Solver (Gurobi)

```bash
python cbam_multimodal_milp_gurobi.py 50node_150order_inputt.xlsx
```

Output: `cbam_milp_solution_<timestamp>.xlsx`

---

## ALNS Design

The metaheuristic uses adaptive operator selection with simulated annealing acceptance.

**Destroy operators**

| Operator | Description |
|----------|-------------|
| Random | Removes a random subset of orders |
| Worst Cost | Targets the most expensive assignments |
| CBAM | Targets orders with highest CBAM exposure |
| Deadline | Targets orders closest to deadline violation |
| Source Overload | Removes orders from over-capacity sources |
| Related | Removes geographically/structurally similar orders |

**Repair operators**

| Operator | Description |
|----------|-------------|
| Greedy | Inserts orders by lowest marginal cost |
| Noisy Greedy | Greedy with randomised cost perturbation |
| Regret-2 | Inserts by maximum regret between best and second-best |
| CBAM Priority | Prioritises high-CBAM-exposure orders |

Initial solutions are generated by a portfolio of constructive heuristics (best-fit, regret-based, criticality-sequenced, backtracking), with the best feasible solution used as the starting point.

---

## Output

Both solvers produce a styled Excel report containing:

- **KPI summary** — total cost breakdown (transport / emission / handling / CBAM)
- **Order summary** — per-order route, cost, CBAM period, on-time status
- **Route legs** — mode-colour-coded leg-by-leg breakdown with dates and distances
- **Solver metadata** — planning year, φ_r, objective value, runtime

