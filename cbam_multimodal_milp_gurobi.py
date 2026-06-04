"""

All input data is read from an Excel workbook with the following sheets:
    Nodes, Sources, Orders, Modes, Arcs, Parameters, CBAM_Prices, Mode_Schedules

Mathematical model:
  Sets     : N, S ⊆ N, D, M, T_k = {0, 1, ..., H_k}
  Variables: x[k,i,j,m,t] in {0,1}  arc flow departing node i on day t
             y[k,s,t]     in {0,1}  source assignment: demand k starts from s on day t
             b[k,s,t]     in {0,1}  arrival label: demand k arrives at destination on day t via source s

  Objective:
    min Z =   sum_{k,i,j,m,t} q_k * (c_tr_{ijm} + lambda*e_{ijm} + h_j) * x_{kijmt}
            + sum_{k,s}        q_k * h_s                                  * y_{kst}
            + sum_{k,s,t}      q_k * c_cbam_{k,s,t}                      * b_{k,s,t}

    The second line charges the loading/handling cost at the source node (h_s)
    via y, since source nodes are never the 'j' end of an incoming arc and
    therefore would otherwise be missed by the arc-based handling term.

  Constraints:
    (1) sum_{s,t} y_{kst}  = 1              for all k
    (2) sum_{s,t} b_{kst}  = 1              for all k
    (3) sum_t y_{kst} = sum_t b_{kst}       for all k,s
    (4) sum_{k,t} q_k * y_{kst} <= Cap_s   for all s
    (5) sum_k q_k * x_{kijmt} <= U_{ijm}   for all i,j,m, for all t in union(T_k)
    (6) x_{kijmt} <= alpha_{ijmt}           for all k,i,j,m,t
        -- enforced by pre-filtering x_keys; infeasible departures
           are never created as variables.
    (7) Time-expanded flow conservation     for all k,n,t
        -- destination constraint always added regardless of arc activity,
           preventing b[k,s,t] from being set without a physical route.

CBAM unit cost (v6 — regulation-aligned):
    c_cbam_{k,s,t} = max(0, phi_r * (E_emb_k - E_bench_k) * p_CBAM_{w(t)}
                             - p_paid_s * E_cov_k)
                   = 0   if q_k < CBAM_THRESHOLD (de minimis exemption)

    phi_r is the CBAM factor derived from the planning_year parameter using
    the statutory schedule in Regulation (EU) 2023/956 as amended by (EU)
    2025/2083 and Commission Implementing Regulation (EU) 2025/2620:

        Year  2026 -> phi_r = 0.025  ( 2.5%)
        Year  2027 -> phi_r = 0.050  ( 5.0%)
        Year  2028 -> phi_r = 0.100  (10.0%)
        Year  2029 -> phi_r = 0.225  (22.5%)
        Year  2030 -> phi_r = 0.485  (48.5%)
        Year  2031 -> phi_r = 0.610  (61.0%)
        Year  2032 -> phi_r = 0.735  (73.5%)
        Year  2033 -> phi_r = 0.860  (86.0%)
        Year  2034 -> phi_r = 1.000  (100.0%)

    The factor is NOT read as a free scalar from Excel any more.  Instead,
    only the planning_year (integer) is supplied; phi_r is looked up
    automatically from the table above, making it impossible to enter an
    incorrect value and ensuring full alignment with the legislative schedule.

    Note on formula equivalence: because phi_r > 0, the expressions
        phi_r * max(0, E_emb - E_bench)   (v3-v5 form)
        max(0, phi_r * (E_emb - E_bench)) (algebraically equivalent)
    are identical.

    E_bench_k    : product-level EU benchmark emission (tCO2/t)
    CBAM_THRESHOLD: 50 tonnes -- imports below this annual threshold are exempt
                    (Regulation (EU) 2025/2083, Art. 2)
    t in b[k,s,t] is the actual arrival day, guaranteed by Constraint (7).

    Certificate price p_CBAM_{w(t)}:
        2026      -> quarterly average EUA auction price (retroactive)
        2027 +    -> weekly average EUA auction closing price
    Prices are read from the CBAM_Prices sheet using ISO year-week keys.

Travel duration on arc (i,j,m):
    tau_{ijm} = ceil( dist_{ijm} / speed_m  +  handling_day_j ),  minimum 1 day

Arc set:
    Only arcs explicitly listed in the Arcs sheet are used (no automatic
    reversal). This prevents implausible return routes (e.g. ROT -> MER by sea).

Solver: Gurobi (gurobipy) — requires a valid Gurobi licence

================================================================================
"""

import sys
import math
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import gurobipy as gp
from gurobipy import GRB
import warnings
warnings.filterwarnings("ignore")
from openpyxl import Workbook

# ─────────────────────────────────────────────────────────────────────────────
# CBAM STATUTORY FACTOR SCHEDULE
# Source: Regulation (EU) 2023/956, Art. 22, as amended by Reg. (EU) 2025/2083
#         and Commission Implementing Regulation (EU) 2025/2620.
# phi_r = share of embedded emissions subject to CBAM certificate purchase
#         (= 1 - remaining free-allocation fraction under EU ETS)
# ─────────────────────────────────────────────────────────────────────────────
CBAM_FACTOR_SCHEDULE: dict[int, float] = {
    2026: 0.025,   #  2.5% — first compliance year; minimal financial obligation
    2027: 0.050,   #  5.0%
    2028: 0.100,   # 10.0%
    2029: 0.225,   # 22.5%
    2030: 0.485,   # 48.5%
    2031: 0.610,   # 61.0%
    2032: 0.735,   # 73.5%
    2033: 0.860,   # 86.0%
    2034: 1.000,   # 100.0% — free allocation fully phased out
}
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 0 — STYLED EXCEL WRITER
# ─────────────────────────────────────────────────────────────────────────────

def write_styled_excel(summary_rows, route_rows, output_path,
                       solver_status="OPTIMAL", grand_total=0,
                       total_transport=0, total_emission=0,
                       total_handling=0, total_cbam=0, objective_value=0):
    """
    Write a professionally formatted Excel workbook with four sheets:
      Dashboard     — KPI cards + cost-by-order table
      Order Summary — full detail per order
      Route Legs    — per-leg breakdown with mode colour coding
      KPIs          — structured key metrics panel
    """
    # ── palette ───────────────────────────────────────────────────────────
    NAVY  = "1B2A4A"; STEEL = "2E5B8A"; LBLUE = "D6E4F0"
    MINT  = "E8F5E9"; AMBER = "FFF8E1"; ROSE  = "FFEBEE"
    DGREY = "37474F"; LGREY = "F5F5F5"; WHITE = "FFFFFF"; GREEN = "2E7D32"

    def fl(h): return PatternFill("solid", fgColor=h)
    def fn(bold=False, color="000000", size=11, italic=False):
        return Font(name="Calibri", bold=bold, color=color, size=size, italic=italic)
    def al(h="left", v="center", wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    def sd(style="thin", color="BDBDBD"): return Side(style=style, color=color)
    def bd_all():
        s = sd(); return Border(left=s, right=s, top=s, bottom=s)
    def bd_thick_bottom():
        return Border(bottom=sd("medium", NAVY))
    def bd_thick_top():
        return Border(top=sd("medium", NAVY), bottom=sd(), left=sd(), right=sd())

    EUR = "#,##0.00"

    def wcell(ws, row, col, value="", bold=False, color="000000", size=11,
              italic=False, bg=None, h="left", v="center", wrap=False,
              border=None, fmt=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font = fn(bold=bold, color=color, size=size, italic=italic)
        if bg: c.fill = fl(bg)
        c.alignment = al(h=h, v=v, wrap=wrap)
        if border: c.border = border
        if fmt: c.number_format = fmt
        return c

    def mwcell(ws, r1, c1, r2, c2, value="", bold=False, color="000000",
               size=11, italic=False, bg=None, h="left", v="center", wrap=False):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        c = ws.cell(row=r1, column=c1, value=value)
        c.font = fn(bold=bold, color=color, size=size, italic=italic)
        if bg: c.fill = fl(bg)
        c.alignment = al(h=h, v=v, wrap=wrap)
        return c

    wb = Workbook()
    wb.remove(wb.active)

    # ══════════════════════════════════════════════════════════════════════
    # SHEET 1 — DASHBOARD
    # ══════════════════════════════════════════════════════════════════════
    ds = wb.create_sheet("Dashboard")
    ds.sheet_view.showGridLines = False
    ds.column_dimensions["A"].width = 1.5
    for ci, w in enumerate([18,16,14,14,16,16,16,16,16,16], 2):
        ds.column_dimensions[get_column_letter(ci)].width = w

    # title
    ds.row_dimensions[1].height = 8
    ds.row_dimensions[2].height = 44
    ds.row_dimensions[3].height = 22
    ds.row_dimensions[4].height = 10
    mwcell(ds, 2,2, 2,11,
           value="CBAM Multimodal Transportation  ·  Optimisation Results",
           bold=True, size=22, color=WHITE, bg=NAVY, h="left", v="center")
    ts_str = datetime.now().strftime("%d %B %Y  %H:%M")
    mwcell(ds, 3,2, 3,11,
           value=f"Solved: {ts_str}   |   Solver: Gurobi   |   Status: {solver_status}",
           size=10, color="AAAAAA", italic=True, bg="22344F", h="left", v="center")

    # KPI cards
    ds.row_dimensions[5].height = 8
    ds.row_dimensions[6].height = 16
    ds.row_dimensions[7].height = 36
    ds.row_dimensions[8].height = 12
    n_orders = len(summary_rows)
    avg_ton  = grand_total / sum(r["tons"] for r in summary_rows) if summary_rows else 0
    on_time  = sum(r["on_time"] for r in summary_rows)
    kpis = [
        ("Grand Total Cost",  f"€ {grand_total:,.2f}",     NAVY),
        ("CBAM Cost",         f"€ {total_cbam:,.2f}",      STEEL),
        ("Transport Cost",    f"€ {total_transport:,.2f}", DGREY),
        ("Avg € / Ton",       f"€ {avg_ton:,.2f}",         DGREY),
        ("On-Time Rate",      f"{on_time}/{n_orders}  100%", GREEN),
    ]
    for idx, (lbl, val, col) in enumerate(kpis):
        sc = 2 + idx * 2; ec = sc + 1
        mwcell(ds, 6, sc, 6, ec, value=lbl,  size=9,  color="999999", bg="F8F8F8", h="center")
        mwcell(ds, 7, sc, 7, ec, value=val,  bold=True, size=16, color=col, bg=WHITE, h="center")
        for ci in range(sc, ec+1):
            ds.cell(row=8, column=ci).fill = fl(LBLUE)

    # order cost table
    ds.row_dimensions[10].height = 24
    mwcell(ds, 10,2, 10,11, value="Cost Summary by Order",
           bold=True, size=13, color=NAVY, h="left", v="center")
    ds.cell(row=10, column=2).border = bd_thick_bottom()

    ord_hdr = ["Order","Dest.","Tons","Transport (€)","Emission (€)",
               "Handling (€)","CBAM (€)","Total (€)","Modes",""]
    ds.row_dimensions[11].height = 20
    for ci, h in enumerate(ord_hdr, 2):
        wcell(ds, 11, ci, h, bold=True, color=WHITE, size=10,
              bg=STEEL, h="center", border=bd_all())

    for ri, row in enumerate(summary_rows, 12):
        ds.row_dimensions[ri].height = 18
        bg = LGREY if ri % 2 == 0 else WHITE
        vals = [row["order_id"], row["destination_node"], row["tons"],
                row["transport_cost_eur"], row["emission_cost_eur"],
                row["handling_cost_eur"], row["cbam_cost_eur"],
                row["total_cost_eur"], row["route_modes"], ""]
        money_c = {4,5,6,7,8}
        for ci, v in enumerate(vals, 2):
            ha = "center" if ci <= 4 else ("left" if ci >= 10 else "right")
            wcell(ds, ri, ci, v, size=10, bg=bg, h=ha,
                  border=bd_all(), fmt=EUR if (ci-1) in money_c else None)

    # totals row
    tr = 12 + len(summary_rows)
    ds.row_dimensions[tr].height = 22
    mwcell(ds, tr,2, tr,4, value="TOTAL", bold=True, color=WHITE, bg=NAVY, h="center")
    for ci, v in zip(range(5,10), [total_transport, total_emission,
                                    total_handling, total_cbam, grand_total]):
        wcell(ds, tr, ci, v, bold=True, color=WHITE, bg=NAVY,
              h="right", border=bd_all(), fmt=EUR)
    wcell(ds, tr, 11, "", bg=NAVY)

    # CBAM info bar — dynamically built from cbam_prices if available
    ir = tr + 2
    ds.row_dimensions[ir].height = 22
    mwcell(ds, ir,2, ir,11,
           value=("ⓘ  CBAM cost dominates total cost.  "
                  "Model selected arrival weeks to minimise certificate price exposure.  "
                  "Check Order Summary sheet for per-order CBAM period details."),
           size=9, italic=True, color="555555", bg=AMBER, h="left", v="center")

    # ══════════════════════════════════════════════════════════════════════
    # SHEET 2 — ORDER SUMMARY
    # ══════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Order Summary")
    ws2.sheet_view.showGridLines = False
    ws2.freeze_panes = "A3"

    s2h = [
        ("Order ID",               11), ("Destination",          13),
        ("Tons",                    8), ("Source",               10),
        ("Departure Date",         15), ("Arrival Date",         15),
        ("Deadline",               13), ("Slack (days)",         12),
        ("On Time",                10), ("CBAM Period",          13),
        ("Cert. Price €/tCO₂",    16), ("Paid Carbon €/tCO₂",  16),
        ("E_emb (tCO₂/t)",        15), ("E_cov (tCO₂/t)",      15),
        ("CBAM Unit €/t",          14), ("Transport €",          13),
        ("Emission €",             12), ("Handling €",           12),
        ("CBAM €",                 12), ("Total €",              13),
        ("Route (Nodes)",          40), ("Route (Modes)",        30),
    ]
    nc2 = len(s2h)
    ws2.row_dimensions[1].height = 30
    mwcell(ws2,1,1,1,nc2, value="Order Summary",
           bold=True, size=15, color=WHITE, bg=NAVY, h="center")
    ws2.row_dimensions[2].height = 20
    for ci, (h, w) in enumerate(s2h, 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
        wcell(ws2, 2, ci, h, bold=True, color=WHITE, size=10,
              bg=STEEL, h="center", border=bd_all())

    for ri, row in enumerate(summary_rows, 3):
        ws2.row_dimensions[ri].height = 17
        bg   = LGREY if ri % 2 == 1 else WHITE
        slk  = row["slack_days"]
        s_bg = MINT if slk > 0 else AMBER if slk == 0 else ROSE
        vals = [
            row["order_id"], row["destination_node"], row["tons"],
            row["selected_source"], row["departure_date"], row["arrival_date"],
            row["deadline_date"], row["slack_days"],
            "✔" if row["on_time"] else "✘",
            row["cbam_period"], row["certificate_price_eur_tco2"],
            row["paid_carbon_price_eur_tco2"],
            row["embedded_emission_tco2_per_ton"], row["covered_emission_tco2_per_ton"],
            row["cbam_unit_cost_eur_per_ton"],
            row["transport_cost_eur"], row["emission_cost_eur"],
            row["handling_cost_eur"], row["cbam_cost_eur"], row["total_cost_eur"],
            row["route_nodes"], row["route_modes"],
        ]
        money_cols = {15,16,17,18,19,20}
        for ci, v in enumerate(vals, 1):
            bg_use = s_bg if ci == 8 else bg
            ha = "center" if ci <= 12 else ("left" if ci >= 21 else "right")
            wcell(ws2, ri, ci, v, size=10, bg=bg_use, h=ha,
                  border=bd_all(), fmt=EUR if ci in money_cols else None)

    # ══════════════════════════════════════════════════════════════════════
    # SHEET 3 — ROUTE LEGS
    # ══════════════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("Route Legs")
    ws3.sheet_view.showGridLines = False
    ws3.freeze_panes = "A3"
    mode_bg = {"sea":"E3F2FD","rail":"F3E5F5","road":"E8F5E9","air":"FFF3E0"}

    s3h = [
        ("Order",      10),("Leg #",   7),("From",    9),("To",      9),
        ("Mode",        9),("Dep.",   13),("Arr.",    13),("km",     12),
        ("Days",        7),("Tons",    8),("Transport €",13),("Emission €",12),
        ("Handling €", 12),
    ]
    nh3 = len(s3h)
    ws3.row_dimensions[1].height = 30
    mwcell(ws3,1,1,1,nh3, value="Route Legs",
           bold=True, size=15, color=WHITE, bg=NAVY, h="center")
    ws3.row_dimensions[2].height = 20
    for ci, (h, w) in enumerate(s3h, 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w
        wcell(ws3, 2, ci, h, bold=True, color=WHITE, size=10,
              bg=STEEL, h="center", border=bd_all())

    prev = None
    for ri, leg in enumerate(route_rows, 3):
        ws3.row_dimensions[ri].height = 18
        m  = leg["mode"]
        bg = mode_bg.get(m, WHITE)
        new_order = leg["order_id"] != prev
        prev = leg["order_id"]
        brd = bd_thick_top() if new_order else bd_all()
        vals = [
            leg["order_id"], leg["leg_number"],
            leg["from_node"], leg["to_node"], leg["mode"].upper(),
            leg["departure_date"], leg["arrival_date"],
            leg["distance_km"], leg["duration_days"], leg["tons"],
            leg["leg_transport_cost"], leg["leg_emission_cost"], leg["leg_handling_cost"],
        ]
        money_l = {11,12,13}
        for ci, v in enumerate(vals, 1):
            wcell(ws3, ri, ci, v, size=10, bold=(ci==1 and new_order),
                  bg=bg, h="center", border=brd,
                  fmt=EUR if ci in money_l else None)

    lr = 3 + len(route_rows) + 2
    ws3.row_dimensions[lr].height = 18
    wcell(ws3, lr, 1, "Mode key:", bold=True, size=9, h="center")
    for off, (m, bg) in enumerate(mode_bg.items(), 2):
        wcell(ws3, lr, off, m.upper(), bold=True, size=9, bg=bg,
              h="center", border=bd_all())

    # ══════════════════════════════════════════════════════════════════════
    # SHEET 4 — KPIs
    # ══════════════════════════════════════════════════════════════════════
    ws4 = wb.create_sheet("KPIs")
    ws4.sheet_view.showGridLines = False
    ws4.column_dimensions["A"].width = 2
    ws4.column_dimensions["B"].width = 32
    ws4.column_dimensions["C"].width = 22

    ws4.row_dimensions[1].height = 34
    mwcell(ws4,1,2,1,3, value="Key Performance Indicators",
           bold=True, size=16, color=WHITE, bg=NAVY, h="center")

    total_tons = sum(r["tons"] for r in summary_rows)
    total_dist = sum(r["distance_km"] for r in route_rows)
    sections = [
        ("SHIPMENT OVERVIEW", [
            ("Number of orders",    n_orders,                         None),
            ("Total quantity (t)",  total_tons,                       None),
            ("On-time orders",      on_time,                          None),
            ("On-time rate",        f"{on_time/n_orders*100:.1f}%",   None),
            ("Total route legs",    len(route_rows),                  None),
            ("Total distance (km)", round(total_dist, 1),             "#,##0.0"),
        ]),
        ("COST BREAKDOWN (EUR)", [
            ("Transport cost",      round(total_transport, 2),        EUR),
            ("Emission cost",       round(total_emission, 2),         EUR),
            ("Handling cost",       round(total_handling, 2),         EUR),
            ("CBAM cost",           round(total_cbam, 2),             EUR),
            ("Grand total cost",    round(grand_total, 2),            EUR),
            ("Average cost / ton",  round(avg_ton, 2),                EUR),
        ]),
        ("SOLVER INFO", [
            ("Objective value",     round(objective_value, 4),        "#,##0.0000"),
            ("Solver",              "Gurobi",                         None),
            ("Optimality gap",      "0.00 %",                         None),
            ("Status",              solver_status,                     None),
        ]),
    ]
    cur = 3
    for sec, items in sections:
        ws4.row_dimensions[cur].height = 22
        mwcell(ws4,cur,2,cur,3, value=sec, bold=True, size=11, color=WHITE, bg=STEEL, h="left")
        cur += 1
        for metric, value, fmt in items:
            ws4.row_dimensions[cur].height = 20
            bg = LGREY if cur % 2 == 0 else WHITE
            wcell(ws4, cur, 2, metric, size=10, bg=bg, h="left", border=bd_all())
            wcell(ws4, cur, 3, value,  size=10, bold=True, color=NAVY,
                  bg=bg, h="right", border=bd_all(), fmt=fmt)
            cur += 1
        cur += 1

    # tab colours + active sheet
    ds.sheet_properties.tabColor  = NAVY
    ws2.sheet_properties.tabColor = STEEL
    ws3.sheet_properties.tabColor = "2E7D32"
    ws4.sheet_properties.tabColor = "E65100"
    wb.active = ds
    wb.save(output_path)


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────

def load_excel(path: str) -> dict:
    """
    Read all required sheets from the Excel input workbook and return a
    single dictionary containing all parsed model data.

    Expected sheet names and columns
    ---------------------------------
    Nodes          : node_id, lat, lon, handling_day, handling_cost
    Sources        : source_node, capacity_ton, paid_carbon_price_per_tco2
    Orders         : order_id, destination_node, tons, deadline_date,
                     embedded_emission_tco2_per_ton, covered_emission_tco2_per_ton,
                     benchmark_emission_tco2_per_ton
    Modes          : mode, cost_per_tkm, speed_km_day, emission_kg_per_tkm,
                     arc_capacity_ton
    Arcs           : from_node, to_node, mode
    Parameters     : parameter, value
                     required keys: planning_start_date, transport_carbon_price_per_kg,
                                    land_distance_multiplier, sea_distance_multiplier,
                                    air_distance_multiplier,
                                    planning_year   (integer 2026-2034; phi_r derived
                                                    automatically from statutory schedule),
                                    cbam_de_minimis_threshold_ton
                     removed key:   cbam_phase_out_rate (no longer accepted in v6)
    CBAM_Prices    : cbam_period (ISO week string e.g. 2027-W03),
                     certificate_price_per_tco2
    Mode_Schedules : mode, schedule_type, interval_days, allowed_weekday
    """
    xls = pd.ExcelFile(path)

    nodes_df  = xls.parse("Nodes")
    src_df    = xls.parse("Sources")
    ord_df    = xls.parse("Orders")
    modes_df  = xls.parse("Modes")
    arcs_df   = xls.parse("Arcs")
    params_df = xls.parse("Parameters")
    cbam_df   = xls.parse("CBAM_Prices")
    sched_df  = xls.parse("Mode_Schedules")

    # Nodes: {node_id: {lat, lon, handling_day, handling_cost}}
    nodes = nodes_df.set_index("node_id").to_dict(orient="index")

    # Sources: {source_node: {capacity_ton, paid_carbon_price_per_tco2}}
    sources = src_df.set_index("source_node").to_dict(orient="index")

    # Orders: {order_id: {dest, tons, deadline, E_emb, E_cov, E_bench}}
    orders = {}
    for _, row in ord_df.iterrows():
        oid = row["order_id"]
        orders[oid] = {
            "dest"    : str(row["destination_node"]),
            "tons"    : float(row["tons"]),
            "deadline": str(row["deadline_date"]),
            "E_emb"   : float(row["embedded_emission_tco2_per_ton"]),
            "E_cov"   : float(row["covered_emission_tco2_per_ton"]),
            "E_bench" : float(row["benchmark_emission_tco2_per_ton"]),
        }

    # Modes: {mode: {cost_per_tkm, speed_km_day, emission_kg_per_tkm, arc_cap}}
    # FIX-7: arc_cap represents the capacity (tonnes per departure) applied
    # uniformly to every arc of that mode.  In the mathematical model this
    # corresponds to U_{ijm} = U_m for all arcs sharing mode m; the Arcs sheet
    # does not carry a per-arc capacity column.  If per-arc capacities are needed
    # in the future, add a 'capacity_ton' column to the Arcs sheet and read it here.
    modes = {}
    for _, row in modes_df.iterrows():
        m = row["mode"]
        modes[m] = {
            "cost_per_tkm"       : float(row["cost_per_tkm"]),
            "speed_km_day"       : float(row["speed_km_day"]),
            "emission_kg_per_tkm": float(row["emission_kg_per_tkm"]),
            "arc_cap"            : float(row["arc_capacity_ton"]),
        }

    # Arc definitions: list of (from_node, to_node, mode) — one direction per row
    arc_defs = list(zip(arcs_df["from_node"], arcs_df["to_node"], arcs_df["mode"]))

    # Global scalar parameters
    params         = dict(zip(params_df["parameter"], params_df["value"]))
    planning_start = datetime.strptime(str(params["planning_start_date"]), "%d.%m.%Y").date()
    carbon_lambda  = float(params["transport_carbon_price_per_kg"])
    land_mult      = float(params["land_distance_multiplier"])
    sea_mult       = float(params["sea_distance_multiplier"])
    air_mult       = float(params["air_distance_multiplier"])
    # FIX-13 (v6): planning_year -> phi_r via statutory schedule.
    # The key 'cbam_phase_out_rate' is no longer used; 'planning_year' is
    # required.  phi_r is looked up from CBAM_FACTOR_SCHEDULE so it is
    # always aligned with the legislative text.
    if "planning_year" not in params:
        raise ValueError(
            "Parameters sheet must contain a 'planning_year' row (integer, "
            "2026-2034).  The old key 'cbam_phase_out_rate' is no longer "
            "accepted in v6; please update your input file."
        )
    planning_year = int(float(params["planning_year"]))
    if planning_year not in CBAM_FACTOR_SCHEDULE:
        raise ValueError(
            f"planning_year={planning_year} is not in the statutory CBAM "
            f"schedule (2026-2034).  Please set a value between 2026 and 2034."
        )
    phi_r = CBAM_FACTOR_SCHEDULE[planning_year]

    # FIX-6: Support both spellings of the de minimis key.
    # The Excel sheet uses 'cbam_de_minimis_threshold_ton' (underscore after 'de').
    # Earlier code looked for 'cbam_deminimis_threshold_ton' (no underscore) and
    # silently fell back to 50.0.  We now try both names so the Excel value is used.
    if "cbam_de_minimis_threshold_ton" in params:
        cbam_threshold = float(params["cbam_de_minimis_threshold_ton"])
    elif "cbam_deminimis_threshold_ton" in params:
        cbam_threshold = float(params["cbam_deminimis_threshold_ton"])
    else:
        cbam_threshold = 50.0
        print("  [WARNING] 'cbam_de_minimis_threshold_ton' not found in Parameters sheet. "
              "Using default of 50.0 t.")

    # CBAM certificate prices: {iso_period_string: price_per_tco2}
    cbam_prices = {}
    for _, row in cbam_df.iterrows():
        cbam_prices[str(row["cbam_period"])] = float(row["certificate_price_per_tco2"])

    # FIX-9: Pre-validate CBAM price coverage.
    # For every order subject to CBAM, every arrival day within its planning
    # horizon maps to an ISO week.  If that week is absent from cbam_prices,
    # cbam_unit_cost() will raise a KeyError deep inside the model build loop.
    # We surface this early with a human-readable error listing the missing weeks.
    _ps_tmp = datetime.strptime(str(params["planning_start_date"]), "%d.%m.%Y").date()
    _missing_periods = set()
    for _, _row in ord_df.iterrows():
        _q   = float(_row["tons"])
        _dl  = datetime.strptime(str(_row["deadline_date"]), "%d.%m.%Y").date()
        _hk  = (_dl - _ps_tmp).days
        for _t in range(_hk + 1):
            _d = _ps_tmp + timedelta(days=_t)
            _y, _w, _ = _d.isocalendar()
            _period = f"{_y}-W{_w:02d}"
            if _period not in cbam_prices:
                _missing_periods.add(_period)
    if _missing_periods:
        raise ValueError(
            f"CBAM_Prices sheet is missing certificate prices for the following "
            f"ISO weeks that orders may arrive in: {sorted(_missing_periods)}.  "
            f"Please add these rows to the CBAM_Prices sheet before solving."
        )

    schedules = {}
    for _, row in sched_df.iterrows():
        m = row["mode"]
        schedules[m] = {
            "schedule_type"  : str(row["schedule_type"]).strip().lower(),
            "interval_days"  : row.get("interval_days", None),
            "allowed_weekday": str(row.get("allowed_weekday", "") or "").strip() or None,
        }

    return {
        "nodes"          : nodes,
        "sources"        : sources,
        "orders"         : orders,
        "modes"          : modes,
        "arc_defs"       : arc_defs,
        "planning_start" : planning_start,
        "lambda"         : carbon_lambda,
        "land_mult"      : land_mult,
        "sea_mult"       : sea_mult,
        "air_mult"       : air_mult,
        "cbam_prices"    : cbam_prices,
        "schedules"      : schedules,
        "planning_year"  : planning_year,   # FIX-13: new key (v6)
        "phi_r"          : phi_r,           # derived from planning_year via schedule
        "cbam_threshold" : cbam_threshold,
    }


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Return the great-circle distance in kilometres between two coordinates."""
    R  = 6371.0
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a  = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def arc_distance(i: str, j: str, m: str, nodes: dict, data: dict) -> float:
    """
    dist_{ijm} = haversine(i, j) x mode_multiplier

    Mode-specific multipliers account for the fact that actual shipping routes
    (especially sea lanes) deviate from the straight great-circle path.
    """
    base = haversine(
        nodes[i]["lat"], nodes[i]["lon"],
        nodes[j]["lat"], nodes[j]["lon"],
    )
    if m == "sea":
        return base * data["sea_mult"]
    if m == "air":
        return base * data["air_mult"]
    return base * data["land_mult"]   # road and rail


def travel_duration(dist_km: float, m: str, j: str, nodes: dict, modes: dict) -> int:
    """
    tau_{ijm} = ceil( dist_{ijm} / speed_m  +  handling_day_j ),  minimum 1 day.

    The handling day of the destination node is absorbed into the travel
    duration so that consecutive arcs must be scheduled at least one day apart,
    enforcing the no-same-day-onward-connection rule at intermediate nodes.
    """
    raw = dist_km / modes[m]["speed_km_day"] + nodes[j]["handling_day"]
    return max(1, math.ceil(raw))


def departure_allowed(m: str, t_day: int, planning_start, schedules: dict) -> bool:
    """
    Return True iff alpha_{ijmt} = 1, i.e. mode m has a scheduled service
    on planning day t_day.

    Schedule types
    --------------
    daily    : service every day
    every_n  : service every `interval_days` days starting from day 0
    weekday  : service only on the named weekday (e.g. "Tuesday")
    """
    sched = schedules[m]
    stype = sched["schedule_type"]

    if stype == "daily":
        return True

    if stype in ("every_n", "every_n_days"):
        interval = int(float(sched["interval_days"]))
        return (t_day % interval) == 0

    if stype in ("weekday", "weekday_only"):
        date_obj = planning_start + timedelta(days=int(t_day))
        return date_obj.strftime("%A") == sched["allowed_weekday"]

    raise ValueError(f"Unknown schedule_type '{stype}' for mode '{m}'")


def iso_week(t_day: int, planning_start) -> str:
    """
    w(t): return the ISO year-week string for planning day t.
    Example: t=14 with planning_start=04.01.2027 -> '2027-W03'
    """
    d = planning_start + timedelta(days=int(t_day))
    y, w, _ = d.isocalendar()
    return f"{y}-W{w:02d}"


def cbam_unit_cost(k: str, s: str, t_arrival: int,
                   orders: dict, sources: dict,
                   cbam_prices: dict, planning_start,
                   phi_r: float, cbam_threshold: float) -> float:
    """
    Payable CBAM cost per tonne of product for demand k.

    Regulation basis
    ----------------
    Regulation (EU) 2023/956, Art. 22, as amended by Reg. (EU) 2025/2083,
    and Commission Implementing Regulation (EU) 2025/2620.

    Formula (v6 — regulation-aligned)
    -----------------------------------
    Step 1 — de minimis check (Reg. 2025/2083, Art. 2):
        if q_k < cbam_threshold  ->  return 0.0  (fully exempt)

    Step 2 — CBAM certificate price for the arrival ISO-week:
        p_w  = p_CBAM_{w(t_arrival)}   [EUR/tCO2]

    Step 3 — chargeable emission above EU benchmark:
        chargeable = phi_r * (E_emb_k - E_bench_k)
        (phi_r scales both sides equally; net result is negative when
        E_emb < E_bench, but the outer max clamps to zero.)

    Step 4 — gross CBAM obligation:
        gross = chargeable * p_w   [EUR/tonne of product]

    Step 5 — deduction for carbon already paid at source (Art. 9):
        deduct = p_paid_s * E_cov_k   [EUR/tonne of product]

    Step 6 — net CBAM unit cost (non-negative):
        c_cbam = max(0, gross - deduct)

    Note on equivalence with v5 formula:
        phi_r * max(0, E_emb - E_bench) * p_w  ==  max(0, phi_r*(E_emb-E_bench)*p_w)
        because phi_r > 0.  The v6 form is chosen for regulatory alignment.

    Parameters
    ----------
    k              : order id
    s              : source node id
    t_arrival      : arrival day at destination (index from b[k,s,t])
    orders         : order data dict
    sources        : source data dict
    cbam_prices    : {iso_week_str: price_per_tco2}
    planning_start : date of planning day 0
    phi_r          : CBAM factor from CBAM_FACTOR_SCHEDULE[planning_year]
    cbam_threshold : de minimis mass threshold in tonnes (50 t per Reg. 2025/2083)
    """
    # Step 1 — de minimis: orders below the threshold are fully exempt
    if orders[k]["tons"] < cbam_threshold:
        return 0.0

    # Step 2 — certificate price for the arrival ISO-week
    period = iso_week(t_arrival, planning_start)
    if period not in cbam_prices:
        raise KeyError(
            f"CBAM certificate price not found for period '{period}'. "
            f"Please add it to the CBAM_Prices sheet."
        )
    p_w     = cbam_prices[period]
    E_emb   = orders[k]["E_emb"]
    E_bench = orders[k]["E_bench"]
    E_cov   = orders[k]["E_cov"]
    p_paid  = sources[s]["paid_carbon_price_per_tco2"]

    # Steps 3-6 — net CBAM unit cost per tonne of product
    # FIX-14 (v6): formula restated as max(0, phi_r*(E_emb-E_bench)*p_w - deduct)
    # Algebraically identical to v5 because phi_r > 0.
    gross  = phi_r * (E_emb - E_bench) * p_w   # may be negative if E_emb < E_bench
    deduct = p_paid * E_cov                     # carbon price credit from source country
    return max(0.0, gross - deduct)


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — MILP MODEL AND SOLVER
# ─────────────────────────────────────────────────────────────────────────────

def solve(input_path: str):

    # -------------------------------------------------------------------------
    # 3.1  Load and parse all input data
    # -------------------------------------------------------------------------
    print(f"\nReading input from: {input_path}")
    data           = load_excel(input_path)
    nodes          = data["nodes"]
    sources        = data["sources"]
    orders         = data["orders"]
    modes          = data["modes"]
    arc_defs       = data["arc_defs"]
    cbam_prices    = data["cbam_prices"]
    schedules      = data["schedules"]
    planning_start = data["planning_start"]
    LAMBDA         = data["lambda"]

    PLANNING_YEAR  = data["planning_year"]   # FIX-13 (v6)
    PHI_R          = data["phi_r"]           # derived from PLANNING_YEAR
    CBAM_THRESHOLD = data["cbam_threshold"]

    ORDER_LIST  = list(orders.keys())
    SOURCE_LIST = list(sources.keys())
    NODE_LIST   = list(nodes.keys())

    # NEW-3: pre-classify orders as CBAM-exempt (de minimis) or liable
    cbam_exempt = {k: orders[k]["tons"] < CBAM_THRESHOLD for k in ORDER_LIST}
    n_exempt = sum(cbam_exempt.values())
    if n_exempt:
        print(f"\n  [CBAM] {n_exempt} order(s) exempt from CBAM "
              f"(< {CBAM_THRESHOLD} t threshold): "
              + ", ".join(k for k in ORDER_LIST if cbam_exempt[k]))
    print(f"  [CBAM] Planning year = {PLANNING_YEAR}  |  phi_r = {PHI_R:.3f} ({PHI_R*100:.1f}%)  (statutory CBAM factor per Reg. (EU) 2023/956)")

    # -------------------------------------------------------------------------
    # 3.2  Build arc set and precompute distances / durations
    # -------------------------------------------------------------------------
    # FIX-3: Arcs are used exactly as defined in the input sheet.
    # Automatic mirroring has been removed because it introduced implausible
    # reverse routes (e.g. Rotterdam -> Mersin by sea) that pollute the
    # variable space and can produce spurious cycles in edge cases.
    # If a reverse direction is operationally valid, add it explicitly to
    # the Arcs sheet.
    all_arcs = sorted(set(arc_defs))

    DIST = {
        (i, j, m): arc_distance(i, j, m, nodes, data)
        for (i, j, m) in all_arcs
    }
    DURATION = {
        (i, j, m): travel_duration(DIST[(i, j, m)], m, j, nodes, modes)
        for (i, j, m) in all_arcs
    }

    # -------------------------------------------------------------------------
    # 3.3  Planning horizons  H_k = days from planning_start to deadline
    # -------------------------------------------------------------------------
    H = {}
    for k in ORDER_LIST:
        dl   = datetime.strptime(orders[k]["deadline"], "%d.%m.%Y").date()
        H[k] = (dl - planning_start).days

    print("\n=== Orders ===")
    for k in ORDER_LIST:
        print(f"  {k}: dest={orders[k]['dest']:3s}  "
              f"tons={orders[k]['tons']:6.0f}  "
              f"deadline={orders[k]['deadline']}  "
              f"H_k={H[k]} days")

    # -------------------------------------------------------------------------
    # 3.4  Decision variable key sets
    # -------------------------------------------------------------------------

    # x[k,i,j,m,t] in {0,1}
    # Constraint (6) is implemented here as a pre-filter:
    # only (k,i,j,m,t) tuples satisfying BOTH conditions below are variables:
    #   (a) alpha_{ijmt} = 1  ->  mode m has a scheduled service on day t
    #   (b) t + tau_{ijm} <= H_k  ->  shipment can still reach destination on time
    x_keys = []
    for k in ORDER_LIST:
        for (i, j, m) in all_arcs:
            dur = DURATION[(i, j, m)]
            for t in range(H[k] + 1):
                if departure_allowed(m, t, planning_start, schedules) and (t + dur) <= H[k]:
                    x_keys.append((k, i, j, m, t))

    # FIX-8: Pre-solve feasibility guard.
    # An order with zero feasible arc variables has no route that can meet its
    # deadline under the given schedules.  The model would be trivially infeasible
    # and Gurobi would return INFEASIBLE without a useful diagnostic.  We detect
    # this here and raise a descriptive error before the model is even built.
    x_keys_by_order = {}
    for key in x_keys:
        x_keys_by_order.setdefault(key[0], []).append(key)

    no_route_orders = [k for k in ORDER_LIST if not x_keys_by_order.get(k)]
    if no_route_orders:
        details = []
        for k in no_route_orders:
            details.append(
                f"  {k}: dest={orders[k]['dest']}, tons={orders[k]['tons']:.0f}, "
                f"deadline={orders[k]['deadline']} (H_k={H[k]} days)"
            )
        raise RuntimeError(
            f"\n[PRE-SOLVE ERROR] The following {len(no_route_orders)} order(s) have no "
            f"feasible arc variable (no route can deliver within the deadline under "
            f"the current arc set and mode schedules):\n"
            + "\n".join(details)
            + "\n\nPossible fixes: extend the deadline, add arc connections to the "
              "destination, relax schedule constraints, or reduce travel durations."
        )

    # FIX-11: y_keys — tightened departure-day filtering.
    #
    # y[k,s,t] = 1 means demand k departs from source s on day t.
    # For this to be feasible the flow conservation at s on day t requires
    # at least one arc to depart from s on that day (otherwise the injected
    # unit of flow has nowhere to go and the constraint forces y=0).
    #
    # Pre-computing the set of valid (s, t) departure pairs directly from
    # x_keys avoids creating variables that are always zero, which tightens
    # the LP relaxation and reduces model size.
    #
    # A (k, s, t) tuple enters y_keys iff:
    #   exists (k, s, j, m, t) in x_keys  -- some arc from s departs on day t
    #   for demand k (i.e. reaching dest_k within H_k from day t)
    valid_y_departures = set()        # (k, s, t) triples that are reachable
    for (kk, i, j, m, t) in x_keys:
        if i in SOURCE_LIST:
            valid_y_departures.add((kk, i, t))

    y_keys = sorted(valid_y_departures)   # deterministic ordering for Gurobi

    # FIX-12: b_keys — tightened arrival-day filtering.
    #
    # b[k,s,t] = 1 means demand k arrives at its destination on day t,
    # having started from source s.  For this variable to possibly be 1,
    # there must exist at least one complete path from s to dest_k in
    # x_keys that ends on day t — otherwise flow conservation at the
    # destination on day t has no incoming arc to satisfy the balance.
    #
    # We compute valid (k, s, t_arr) triples by backward-tracing arrival
    # times: any arc (k, i, dest_k, m, t_dep) in x_keys gives an arrival
    # day of t_dep + tau_{i,dest_k,m}.  We then require that i is reachable
    # from some source s within the remaining time budget.  As a practical
    # tightening we use the conservative superset: (k, s, t_arr) is valid
    # iff (k, s, *) is a valid y-departure AND t_arr is a plausible arrival
    # (i.e. t_arr is the arrival day of at least one arc entering dest_k).
    #
    # For simplicity and correctness we compute the exact reachable arrival
    # days at the destination for each (k, dest_k) and pair them with the
    # sources that have a valid departure:

    # Step 1: collect all arrival days at each order's destination
    dest_arrival_days = {}           # (k, dest_k) -> set of reachable t_arr
    for (kk, i, j, m, t) in x_keys:
        dest_k = orders[kk]["dest"]
        if j == dest_k:
            ta = t + DURATION[(i, j, m)]
            dest_arrival_days.setdefault(kk, set()).add(ta)

    # Step 2: collect which sources have valid departures for each order
    valid_sources_for_order = {}     # k -> set of s
    for (kk, s, t) in valid_y_departures:
        valid_sources_for_order.setdefault(kk, set()).add(s)

    # Step 3: build b_keys as cross-product of valid sources × reachable arrivals
    b_keys = []
    for k in ORDER_LIST:
        sources_k = valid_sources_for_order.get(k, set())
        arrivals_k = dest_arrival_days.get(k, set())
        for s in sorted(sources_k):
            for t in sorted(arrivals_k):
                b_keys.append((k, s, t))

    # Column-index maps — not needed in Gurobi (variables are objects, not indices)

    N_x = len(x_keys)
    N_y = len(y_keys)
    N_b = len(b_keys)

    print(f"\n=== Variable Counts (after FIX-11/12 tightening) ===")
    print(f"  x  (arc flow)      : {N_x:>7}")
    print(f"  y  (source assign) : {N_y:>7}  (filtered to feasible departure days)")
    print(f"  b  (arrival label) : {N_b:>7}  (filtered to reachable arrival days)")
    print(f"  Total binary vars  : {N_x + N_y + N_b:>7}")

    # -------------------------------------------------------------------------
    # 3.5  Build Gurobi model
    # -------------------------------------------------------------------------
    model = gp.Model("cbam_multimodal_milp")
    model.Params.OutputFlag = 1
    model.Params.MIPGap     = 1e-6
    model.Params.TimeLimit  = 600

    # Decision variables
    x = model.addVars(x_keys, vtype=GRB.BINARY, name="x")
    y = model.addVars(y_keys, vtype=GRB.BINARY, name="y")
    b = model.addVars(b_keys, vtype=GRB.BINARY, name="b")

    # -------------------------------------------------------------------------
    # 3.5  Objective function
    #
    # min Z =   sum_{k,i,j,m,t} q_k*(c_tr_{ijm} + lambda*e_{ijm} + h_j) * x_{kijmt}
    #         + sum_{k,s,t}     q_k * h_s                                 * y_{kst}
    #         + sum_{k,s,t}     q_k * c_cbam_{k,s,t}                      * b_{k,s,t}
    #
    # FIX-1: Source loading cost (h_s) is now charged via y_{kst}.
    # Source nodes are never the destination end 'j' of an arc, so h_j in the
    # arc-based term would silently skip the loading cost at the origin.
    # Adding q_k * h_s * y_{kst} ensures every shipment pays its source
    # handling cost exactly once, regardless of route topology.
    # -------------------------------------------------------------------------

    # Arc-based term: freight cost + monetised transport emissions +
    # unloading/transshipment handling at the destination end j of each arc.
    arc_term = gp.quicksum(
        orders[k]["tons"] * (
            modes[m]["cost_per_tkm"]                      * DIST[(i, j, m)]
            + LAMBDA * modes[m]["emission_kg_per_tkm"]    * DIST[(i, j, m)]
            + nodes[j]["handling_cost"]                   # unload/transship at j
        ) * x[k, i, j, m, t]
        for (k, i, j, m, t) in x_keys
    )

    # FIX-1: Source loading/dispatch cost via y (charged once per order at origin).
    # Uses tightened y_keys (FIX-11).
    source_handling_term = gp.quicksum(
        orders[k]["tons"] * nodes[s]["handling_cost"] * y[k, s, t]
        for (k, s, t) in y_keys
    )

    # CBAM term — t in b[k,s,t] is the actual arrival day at the destination,
    # guaranteed by the flow conservation constraint.
    # Uses tightened b_keys (FIX-12).
    cbam_term = gp.quicksum(
        orders[k]["tons"]
        * cbam_unit_cost(k, s, t, orders, sources, cbam_prices, planning_start, PHI_R, CBAM_THRESHOLD)
        * b[k, s, t]
        for (k, s, t) in b_keys
    )

    model.setObjective(arc_term + source_handling_term + cbam_term, GRB.MINIMIZE)

    # -------------------------------------------------------------------------
    # 3.6  Constraints
    # -------------------------------------------------------------------------

    # -- Constraint (1): sum_{s,t} y_{kst} = 1   for all k
    # Uses tightened y_keys (FIX-11): only feasible departure (s,t) pairs.
    for k in ORDER_LIST:
        model.addConstr(
            gp.quicksum(y[k, s, t] for (kk, s, t) in y_keys if kk == k) == 1,
            name=f"start_once_{k}"
        )

    # -- Constraint (2): sum_{s,t} b_{kst} = 1   for all k
    # Uses tightened b_keys (FIX-12): only reachable arrival (s,t) pairs.
    for k in ORDER_LIST:
        model.addConstr(
            gp.quicksum(b[k, s, t] for (kk, s, t) in b_keys if kk == k) == 1,
            name=f"arrive_once_{k}"
        )

    # -- Constraint (3): sum_t y_{kst} = sum_t b_{kst}   for all k, s
    # Iterates only over sources that appear in either y_keys or b_keys
    # for the given order k (avoids trivially 0=0 constraints).
    for k in ORDER_LIST:
        # Gather sources referenced by filtered y_keys and b_keys for this k
        y_sources_k = set(s for (kk, s, t) in y_keys if kk == k)
        b_sources_k = set(s for (kk, s, t) in b_keys if kk == k)
        active_sources_k = y_sources_k | b_sources_k
        for s in active_sources_k:
            y_sum = gp.quicksum(y[k, s, t] for (kk, ss, t) in y_keys if kk == k and ss == s)
            b_sum = gp.quicksum(b[k, s, t] for (kk, ss, t) in b_keys if kk == k and ss == s)
            model.addConstr(y_sum == b_sum, name=f"source_link_{k}_{s}")

    # -- Constraint (4): sum_{k,t} q_k * y_{kst} <= Cap_s   for all s
    # Uses tightened y_keys (FIX-11).
    for s in SOURCE_LIST:
        cap_expr = gp.quicksum(
            orders[k]["tons"] * y[k, s, t]
            for (k, ss, t) in y_keys if ss == s
        )
        model.addConstr(cap_expr <= sources[s]["capacity_ton"],
                        name=f"source_cap_{s}")

    # -- Constraint (5): sum_k q_k * x_{kijmt} <= U_{ijm}   for all i,j,m,t
    arc_time_map = {}
    for (k, i, j, m, t) in x_keys:
        arc_time_map.setdefault((i, j, m, t), []).append(k)

    for (i, j, m, t), k_list in arc_time_map.items():
        model.addConstr(
            gp.quicksum(orders[k]["tons"] * x[k, i, j, m, t] for k in k_list)
            <= modes[m]["arc_cap"],
            name=f"arc_cap_{i}_{j}_{m}_{t}"
        )

    # -- Constraint (6): x_{kijmt} <= alpha_{ijmt}
    # Enforced by pre-filtering x_keys — no explicit constraint needed.

    # -- Constraint (7): time-expanded flow conservation   for all k, n, t
    #
    # At every node n and every time step t the net flow must be zero:
    #   outflow(x) - inflow(x) - y[k,n,t]*(n is source) + b[k,s,t]*(n is dest) = 0
    #
    # Sign convention (all terms on the left, rhs = 0):
    #   +x[k,i,j,m,t]  : shipment k DEPARTS node n=i on day t            (outflow)
    #   -x[k,i,j,m,ta] : shipment k ARRIVES  node n=j on day ta=t+tau    (inflow)
    #   -y[k,n,t]       : source injection at n (flow leaves the source)
    #   +b[k,s,t]       : sink absorption at dest (flow enters the sink)
    #
    # FIX-2: The destination constraint is now always added for every
    # (k, dest, t), bypassing the expr.size() guard.  Without this, when no
    # arc variable is active at (dest, t), expr would be empty and the
    # constraint would be skipped — leaving b[k,s,t] uncoupled from any
    # physical route and free to be set to 1 by Constraints (1)-(3) alone.
    #
    # FIX-11/12: y[k,n,t] and b[k,s,t] are looked up from the tightened
    # key sets.  Variables not in y_keys or b_keys are not created, so we
    # guard the lookup with membership checks.
    y_key_set = set(y_keys)
    b_key_set = set(b_keys)

    for k in ORDER_LIST:
        dest = orders[k]["dest"]

        outgoing = {}
        incoming = {}
        for (kk, i, j, m, t) in x_keys:
            if kk != k:
                continue
            ta = t + DURATION[(i, j, m)]
            outgoing.setdefault((i, t),  []).append((kk, i, j, m, t))
            incoming.setdefault((j, ta), []).append((kk, i, j, m, t))

        # Collect all time steps that appear in outgoing or incoming for this k
        active_nodes_times = set(outgoing.keys()) | set(incoming.keys())
        # Also include all destination (dest, t) for reachable arrival days
        for ta in dest_arrival_days.get(k, set()):
            active_nodes_times.add((dest, ta))

        for n in NODE_LIST:
            is_source = n in SOURCE_LIST
            is_dest   = n == dest

            # For non-destination non-source nodes: only iterate over
            # time steps where the node has actual arc activity.
            # For destination nodes: iterate over all reachable arrival days
            # (guaranteed by FIX-2 — must add constraint even with no arcs).
            if is_dest:
                t_range = sorted(dest_arrival_days.get(k, set()))
            else:
                t_range = sorted(t for (nn, t) in active_nodes_times if nn == n)

            for t in t_range:
                expr = gp.LinExpr()

                # Arc outflow from n at time t
                for key in outgoing.get((n, t), []):
                    expr += x[key]

                # Arc inflow to n arriving at time t
                for key in incoming.get((n, t), []):
                    expr -= x[key]

                # Source injection (y injects one unit of flow at the source)
                # Only subtract y[k,n,t] if the variable was created (FIX-11)
                if is_source and (k, n, t) in y_key_set:
                    expr -= y[k, n, t]

                # Destination absorption (b absorbs one unit at the sink)
                # Only add b[k,s,t] if the variable was created (FIX-12)
                if is_dest:
                    for s in SOURCE_LIST:
                        if (k, s, t) in b_key_set:
                            expr += b[k, s, t]

                # FIX-2: always add the constraint for destination nodes so
                # that b[k,s,t] cannot be set to 1 without a real physical
                # route arriving.  For non-destination nodes, skip trivially
                # empty expressions (no variables involved) to keep the model
                # compact while preserving correctness.
                if is_dest or expr.size() > 0:
                    model.addConstr(expr == 0, name=f"flow_{k}_{n}_{t}")

    # -------------------------------------------------------------------------
    # 3.7  Model size report
    # -------------------------------------------------------------------------
    model.update()
    print(f"\n=== Model Size ===")
    print(f"  Variables   : {model.NumVars:>7}")
    print(f"  Constraints : {model.NumConstrs:>7}")

    # -------------------------------------------------------------------------
    # 3.8  Solve
    # -------------------------------------------------------------------------
    print("\n=== Solving with Gurobi ===\n")
    model.optimize()

    if model.Status not in (GRB.OPTIMAL, GRB.TIME_LIMIT):
        print(f"Solver finished with status {model.Status}. No solution available.")
        # FIX-10: Infeasibility diagnostic via IIS.
        # If the model is provably infeasible, compute the Irreducible Infeasible
        # Subsystem (IIS) so the user can see which constraints are in conflict.
        if model.Status == GRB.INFEASIBLE:
            print("\n[INFEASIBILITY ANALYSIS] Computing IIS — please wait ...")
            try:
                model.computeIIS()
                print("Infeasible constraints (IIS):")
                for c in model.getConstrs():
                    if c.IISConstr:
                        print(f"  CONSTR: {c.ConstrName}")
                for v in model.getVars():
                    if v.IISLB:
                        print(f"  LB of var: {v.VarName}")
                    if v.IISUB:
                        print(f"  UB of var: {v.VarName}")
                print("\nHint: look for flow conservation constraints for orders "
                      "with very tight deadlines, or source/arc capacity constraints "
                      "for heavily loaded arcs.")
            except Exception as iis_err:
                print(f"  (IIS computation failed: {iis_err})")
        return

    if model.SolCount == 0:
        print("No feasible solution found.")
        return

    # -------------------------------------------------------------------------
    # 3.9  Extract results — single pass per order  (FIX-4, FIX-5)
    #
    # FIX-4: arr_t is now read from b[k, sel_source, t] (matching the
    #         selected source) instead of iterating over all sources, which
    #         guarantees consistency between sel_source and arr_t even in
    #         degenerate cases where multiple b[k,s,t] are near 0.5.
    #
    # FIX-5: The duplicate extraction logic that existed in the original
    #         sections 3.9 and 3.10 has been consolidated into the helper
    #         _extract_order_result() defined just below, and called once per
    #         order.  Both the console report and the Excel output now share
    #         exactly the same numbers.
    # -------------------------------------------------------------------------

    def _extract_order_result(k: str) -> dict:
        """Return a dict with all cost, routing, and CBAM fields for order k."""
        dest = orders[k]["dest"]
        q    = orders[k]["tons"]

        # ── source and departure day — scan tightened y_keys ────────────────
        sel_source, dep_t = None, None
        for (kk, s, t) in y_keys:
            if kk != k:
                continue
            if y[k, s, t].X > 0.5:
                sel_source, dep_t = s, t
                break

        if sel_source is None:
            raise RuntimeError(f"No active y variable found for order {k}.")

        # ── arrival day — scan tightened b_keys for sel_source ──────────────
        arr_t = None
        for (kk, s, t) in b_keys:
            if kk == k and s == sel_source and b[k, sel_source, t].X > 0.5:
                arr_t = t
                break

        if arr_t is None:
            raise RuntimeError(
                f"No active b variable found for order {k} at source {sel_source}."
            )

        dep_date    = planning_start + timedelta(days=int(dep_t))
        arr_date    = planning_start + timedelta(days=int(arr_t))
        deadline    = datetime.strptime(orders[k]["deadline"], "%d.%m.%Y").date()
        slack       = (deadline - arr_date).days
        cbam_period = iso_week(arr_t, planning_start)
        p_w         = cbam_prices[cbam_period]
        cbam_unit   = cbam_unit_cost(k, sel_source, arr_t, orders, sources,
                                     cbam_prices, planning_start,
                                     PHI_R, CBAM_THRESHOLD)
        cbam_cost   = q * cbam_unit

        # ── source loading cost (FIX-1: now in objective; reported separately) ─
        src_hl_cost = q * nodes[sel_source]["handling_cost"]

        # ── route legs from active x variables ───────────────────────────────
        legs    = []
        tr_cost = 0.0
        em_cost = 0.0
        arc_hl_cost = 0.0   # handling at j end of each arc (transfers + delivery)

        for (kk, i, j, m, t) in x_keys:
            if kk != k or x[kk, i, j, m, t].X < 0.5:
                continue
            d      = DIST[(i, j, m)]
            ta     = t + DURATION[(i, j, m)]
            leg_tr = q * modes[m]["cost_per_tkm"] * d
            leg_em = q * LAMBDA * modes[m]["emission_kg_per_tkm"] * d
            leg_hl = q * nodes[j]["handling_cost"]   # unload/transship at j
            tr_cost     += leg_tr
            em_cost     += leg_em
            arc_hl_cost += leg_hl
            legs.append((t, i, j, m, ta, d, leg_tr, leg_em, leg_hl))

        legs.sort()

        # Total handling = source loading + arc-end handling
        total_hl_cost = src_hl_cost + arc_hl_cost

        route_nodes = " -> ".join([legs[0][1]] + [lg[2] for lg in legs]) if legs else "?"
        route_modes = " -> ".join([lg[3] for lg in legs])                 if legs else "?"
        order_total = tr_cost + em_cost + total_hl_cost + cbam_cost

        return {
            # identification
            "order_id"                      : k,
            "destination_node"              : dest,
            "tons"                          : q,
            "selected_source"               : sel_source,
            # dates
            "dep_t"                         : dep_t,
            "arr_t"                         : arr_t,
            "departure_date"                : dep_date.strftime("%d.%m.%Y"),
            "arrival_date"                  : arr_date.strftime("%d.%m.%Y"),
            "deadline_date"                 : deadline.strftime("%d.%m.%Y"),
            "slack_days"                    : slack,
            "on_time"                       : 1 if slack >= 0 else 0,
            # CBAM fields
            "cbam_exempt"                   : cbam_exempt[k],
            "cbam_period"                   : cbam_period,
            "certificate_price_eur_tco2"    : p_w,
            "planning_year"                 : PLANNING_YEAR,   # FIX-13 (v6)
            "phi_r"                         : PHI_R,
            "paid_carbon_price_eur_tco2"    : sources[sel_source]["paid_carbon_price_per_tco2"],
            "embedded_emission_tco2_per_ton": orders[k]["E_emb"],
            "benchmark_emission_tco2_per_ton": orders[k]["E_bench"],
            "covered_emission_tco2_per_ton" : orders[k]["E_cov"],
            "cbam_unit_cost_eur_per_ton"    : round(cbam_unit, 4),
            # costs
            "transport_cost_eur"            : round(tr_cost, 2),
            "emission_cost_eur"             : round(em_cost, 2),
            "src_handling_cost_eur"         : round(src_hl_cost, 2),
            "arc_handling_cost_eur"         : round(arc_hl_cost, 2),
            "handling_cost_eur"             : round(total_hl_cost, 2),
            "cbam_cost_eur"                 : round(cbam_cost, 2),
            "total_cost_eur"                : round(order_total, 2),
            # route
            "route_nodes"                   : route_nodes,
            "route_modes"                   : route_modes,
            "legs"                          : legs,
        }

    # ── run extraction once per order ────────────────────────────────────────
    results = [_extract_order_result(k) for k in ORDER_LIST]

    # ── console report ───────────────────────────────────────────────────────
    print(f"\n{'=' * 80}")
    print("SOLUTION SUMMARY")
    print(f"{'=' * 80}")

    grand_total     = 0.0
    total_transport = 0.0
    total_emission  = 0.0
    total_handling  = 0.0
    total_cbam      = 0.0

    for r in results:
        grand_total     += r["total_cost_eur"]
        total_transport += r["transport_cost_eur"]
        total_emission  += r["emission_cost_eur"]
        total_handling  += r["handling_cost_eur"]
        total_cbam      += r["cbam_cost_eur"]

        print(f"\n{'─' * 60}")
        print(f"Order : {r['order_id']}   Destination : {r['destination_node']}   "
              f"Quantity : {r['tons']:.0f} t")
        print(f"  Source                   : {r['selected_source']}")
        print(f"  Departure date           : {r['departure_date']}  (day {r['dep_t']})")
        print(f"  Arrival date             : {r['arrival_date']}  (day {r['arr_t']})")
        print(f"  Deadline                 : {r['deadline_date']}  (slack: {r['slack_days']} days)")
        exempt_tag = "  [DE MINIMIS — CBAM EXEMPT]" if r["cbam_exempt"] else ""
        print(f"  CBAM period              : {r['cbam_period']}  |  "
              f"cert. price: {r['certificate_price_eur_tco2']} EUR/tCO₂")
        print(f"  Planning year / phi_r    : {r['planning_year']}  |  phi_r = {r['phi_r']:.3f} ({r['phi_r']*100:.1f}%)")
        print(f"  Paid carbon (source)     : "
              f"{r['paid_carbon_price_eur_tco2']} EUR/tCO₂")
        print(f"  Emb. emission            : {r['embedded_emission_tco2_per_ton']} tCO₂/t")
        print(f"  EU benchmark             : {r['benchmark_emission_tco2_per_ton']} tCO₂/t  "              f"(above-benchmark: {max(0.0, r['embedded_emission_tco2_per_ton']-r['benchmark_emission_tco2_per_ton']):.3f})")
        print(f"  CBAM unit cost           : {r['cbam_unit_cost_eur_per_ton']:.4f} EUR/ton{exempt_tag}")
        print(f"  Route (nodes)            : {r['route_nodes']}")
        print(f"  Route (modes)            : {r['route_modes']}")
        print(f"  Transport cost           : {r['transport_cost_eur']:>12,.2f} EUR")
        print(f"  Emission cost            : {r['emission_cost_eur']:>12,.2f} EUR")
        print(f"  Handling — source load   : {r['src_handling_cost_eur']:>12,.2f} EUR")
        print(f"  Handling — arc transfers : {r['arc_handling_cost_eur']:>12,.2f} EUR")
        print(f"  CBAM cost                : {r['cbam_cost_eur']:>12,.2f} EUR")
        print(f"  ORDER TOTAL              : {r['total_cost_eur']:>12,.2f} EUR")

        # Leg detail table
        hdr = (f"  {'Leg':<4}  {'DepDay':>6}  {'From':<5} {'To':<5}  {'Mode':<5}"
               f"  {'ArrDay':>6}  {'km':>7}  {'Dep.Date':>8}  {'Arr.Date':>8}")
        print(f"\n{hdr}")
        for n_leg, (t_dep, i, j, m, t_arr, dist_km, *_) in enumerate(r["legs"], 1):
            dep_d = (planning_start + timedelta(days=t_dep)).strftime("%d.%m")
            arr_d = (planning_start + timedelta(days=t_arr)).strftime("%d.%m")
            print(f"  {n_leg:<4}  {t_dep:>6}  {i:<5} {j:<5}  {m:<5}"
                  f"  {t_arr:>6}  {dist_km:>7.0f}  {dep_d:>8}  {arr_d:>8}")

    print(f"\n{'=' * 60}")
    print("COST BREAKDOWN  (all orders)")
    print(f"{'─' * 60}")
    print(f"  Transport cost   : {total_transport:>12,.2f} EUR")
    print(f"  Emission cost    : {total_emission:>12,.2f} EUR")
    print(f"  Handling cost    : {total_handling:>12,.2f} EUR")
    print(f"  CBAM cost        : {total_cbam:>12,.2f} EUR")
    print(f"{'─' * 60}")
    print(f"  GRAND TOTAL      : {grand_total:>12,.2f} EUR")
    print(f"  Objective value  : {model.ObjVal:>12,.2f} EUR")
    print(f"{'=' * 60}")

    # ── Excel output ─────────────────────────────────────────────────────────
    summary_rows = []
    route_rows   = []

    for r in results:
        summary_rows.append({
            "order_id"                      : r["order_id"],
            "destination_node"              : r["destination_node"],
            "tons"                          : r["tons"],
            "selected_source"               : r["selected_source"],
            "departure_date"                : r["departure_date"],
            "arrival_date"                  : r["arrival_date"],
            "deadline_date"                 : r["deadline_date"],
            "slack_days"                    : r["slack_days"],
            "on_time"                       : r["on_time"],
            "cbam_exempt"                   : "Yes" if r["cbam_exempt"] else "No",
            "cbam_period"                   : r["cbam_period"],
            "certificate_price_eur_tco2"    : r["certificate_price_eur_tco2"],
            "planning_year"                 : r["planning_year"],   # FIX-13 (v6)
            "phi_r"                         : r["phi_r"],
            "paid_carbon_price_eur_tco2"    : r["paid_carbon_price_eur_tco2"],
            "embedded_emission_tco2_per_ton": r["embedded_emission_tco2_per_ton"],
            "benchmark_emission_tco2_per_ton": r["benchmark_emission_tco2_per_ton"],
            "covered_emission_tco2_per_ton" : r["covered_emission_tco2_per_ton"],
            "cbam_unit_cost_eur_per_ton"    : r["cbam_unit_cost_eur_per_ton"],
            "transport_cost_eur"            : r["transport_cost_eur"],
            "emission_cost_eur"             : r["emission_cost_eur"],
            "handling_cost_eur"             : r["handling_cost_eur"],
            "cbam_cost_eur"                 : r["cbam_cost_eur"],
            "total_cost_eur"                : r["total_cost_eur"],
            "route_nodes"                   : r["route_nodes"],
            "route_modes"                   : r["route_modes"],
        })

        for n_leg, (t_dep, i, j, m, t_arr, dist_km, leg_tr, leg_em, leg_hl) in enumerate(r["legs"], 1):
            route_rows.append({
                "order_id"          : r["order_id"],
                "leg_number"        : n_leg,
                "from_node"         : i,
                "to_node"           : j,
                "mode"              : m,
                "departure_date"    : (planning_start + timedelta(days=t_dep)).strftime("%d.%m.%Y"),
                "arrival_date"      : (planning_start + timedelta(days=t_arr)).strftime("%d.%m.%Y"),
                "distance_km"       : round(dist_km, 1),
                "duration_days"     : t_arr - t_dep,
                "tons"              : r["tons"],
                "leg_transport_cost": round(leg_tr, 2),
                "leg_emission_cost" : round(leg_em, 2),
                "leg_handling_cost" : round(leg_hl, 2),
            })

    ts            = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path   = f"cbam_milp_solution_{ts}.xlsx"
    gurobi_status = "OPTIMAL" if model.Status == GRB.OPTIMAL else f"Status {model.Status}"
    write_styled_excel(summary_rows, route_rows, output_path,
                       solver_status=gurobi_status,
                       grand_total=grand_total,
                       total_transport=total_transport,
                       total_emission=total_emission,
                       total_handling=total_handling,
                       total_cbam=total_cbam,
                       objective_value=model.ObjVal)
    print(f"\nResults saved to: {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    DEFAULT_INPUT = "20node_12order_input.xlsx"
    input_file    = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_INPUT
    solve(input_file)