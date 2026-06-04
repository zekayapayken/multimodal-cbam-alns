
from __future__ import annotations

import argparse
import bisect
import copy
import heapq
import itertools
import math
import os
import random
import sys
import time
from collections import defaultdict
from dataclasses import dataclass, field, replace
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1

# ─────────────────────────────────────────────────────────────────────────────
# SABITLER
# ─────────────────────────────────────────────────────────────────────────────
VERSION     = "13"           # Tüm versiyon referansları bu sabite bağlıdır
EARTH_R     = 6371.0
INF         = float("inf")
EPS         = 1e-9           # Kapasite/maliyet karşılaştırma toleransı (genel)
EPS_REPORT  = 1e-7           # Raporlama toleransı (arc overflow, lateness görünürlük eşiği)

# CBAM aşamalı devreye giriş faktörü φ_r (Tüzük AB 2023/956, Mad. 22)
CBAM_PHI: Dict[int, float] = {
    2026: 0.025, 2027: 0.050, 2028: 0.100, 2029: 0.225,
    2030: 0.485, 2031: 0.610, 2032: 0.735, 2033: 0.860, 2034: 1.000,
}


def cbam_phi(year: int) -> float:
    """
    Planlama yılına göre CBAM φ_r değerini döndürür.
    2026 öncesi → 0.0 (CBAM finansal yükümlülük yok)
    2034 sonrası → 1.0 (tam yükümlülük)
    """
    if year < 2026:
        return 0.0
    if year > 2034:
        return 1.0
    return CBAM_PHI[year]


# ─────────────────────────────────────────────────────────────────────────────
# VERİ YAPILARI
# ─────────────────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Node:
    node_id:       str
    name:          str
    country:       str
    node_type:     str
    lat:           float
    lon:           float
    handling_day:  float   # δ_j (gün)
    handling_cost: float   # h_j (€/ton)


@dataclass(frozen=True)
class Source:
    source_id:   str
    capacity:    float   # Cap_s (ton)
    paid_carbon: float   # p^paid_s (€/tCO2)


@dataclass(frozen=True)
class Mode:
    mode_id:             str
    cost_per_tkm:        float   # c^tr_m (€/ton·km)
    speed_km_day:        float   # v_m (km/gün)
    emission_kg_per_tkm: float   # e_m (kg CO2/ton·km)
    arc_capacity:        float   # U_m (ton/hareket)


@dataclass(frozen=True)
class Arc:
    frm:         str
    to:          str
    mode_id:     str
    dist_km:     float   # δ_ijm (km)
    travel_days: int     # τ_ijm (gün)
    unit_cost:   float   # önhesaplı: taşıma + emisyon + h_j (€/ton)


@dataclass(frozen=True)
class Order:
    order_id:     str
    dest:         str
    qty:          float   # q_k (ton)
    deadline:     int     # H_k (planlama başından gün sayısı)
    emb_em:       float   # E^emb_k (tCO2/ton)
    cov_em:       float   # E^cov_k (tCO2/ton)
    bench_em:     float   # Ē_k (tCO2/ton)
    cbam_subject: bool    # q_k >= Θ mi?
    product:      str = ""  # ürün tipi (steel, aluminium, cement, vb.)


@dataclass
class Leg:
    """Tek bir yay üzerindeki taşıma hareketi."""
    frm:        str
    to:         str
    mode_id:    str
    depart_day: int
    arrive_day: int
    dist_km:    float
    unit_cost:  float   # €/ton


@dataclass
class Assignment:
    """Bir siparişin tam atama kararı: kaynak + güzergâh + zamanlama."""
    order_id:   str
    source:     str
    depart_day: int
    arrive_day: int
    cbam_week:  str
    legs:       List[Leg]
    move_cost:  float   # taşıma + elleçleme (€)
    cbam_cost:  float   # CBAM yükümlülüğü (€)
    total_cost: float   # move_cost + cbam_cost (€)


@dataclass
class Solution:
    assignments:          Dict[str, Assignment] = field(default_factory=dict)
    obj:                  float = INF
    feasible:             bool  = False  # ALNS açısından kabul edilebilir çözüm (soft capacity açıkken penalty'li olabilir)
    strict_feasible:      bool  = False  # kaynak + yay kapasitesi + deadline tam fizibil mi?
    base_obj:             float = INF    # taşıma + CBAM maliyeti (ceza hariç)
    arc_overflow_penalty: float = 0.0    # soft arc kapasite aşım cezası
    overflow_arc_count:   int   = 0      # kapasitesi aşılmış yay-gün sayısı
    total_overflow_ton:   float = 0.0    # toplam yay kapasite aşımı (ton)
    deadline_violation_count: int = 0    # orijinal deadline aşan sipariş sayısı
    total_lateness_days: float = 0.0     # toplam deadline aşım günü
    max_lateness_days:   float = 0.0     # maksimum deadline aşım günü


# ─────────────────────────────────────────────────────────────────────────────
# PROBLEM INSTANCE
# ─────────────────────────────────────────────────────────────────────────────

class Instance:
    """Excel dosyasından tüm problem verilerini yükler ve türetilmiş yapıları kurar."""

    def __init__(self, xlsx_path: str | Path):
        self.path         = Path(xlsx_path)
        self.nodes:       Dict[str, Node]        = {}
        self.sources:     Dict[str, Source]      = {}
        self.orders:      Dict[str, Order]       = {}
        self.modes:       Dict[str, Mode]        = {}
        self.arcs:        List[Arc]              = []
        self.out_arcs:    Dict[str, List[Arc]]   = defaultdict(list)
        self.arc_index:   Dict[Tuple[str, str, str], Arc] = {}   # (frm,to,mode)→Arc O(1) lookup
        self.schedules:   Dict[str, dict]        = {}
        self.cbam_prices: Dict[str, float]       = {}
        self.params:      Dict[str, object]      = {}
        self._load()

    # ── Yardımcılar ──────────────────────────────────────────────────────────

    @staticmethod
    def _haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
        """İki koordinat arasındaki Haversine (büyük çember) mesafesi (km)."""
        phi1 = math.radians(lat1)
        phi2 = math.radians(lat2)
        dphi = math.radians(lat2 - lat1)
        dlam = math.radians(lon2 - lon1)
        a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlam / 2) ** 2
        return 2 * EARTH_R * math.asin(math.sqrt(a))

    @staticmethod
    def _parse_date(x) -> datetime:
        if isinstance(x, datetime):
            return x
        return pd.to_datetime(str(x), dayfirst=True).to_pydatetime()

    @staticmethod
    def _iso_week(dt: datetime) -> str:
        y, w, _ = dt.isocalendar()
        return f"{y}-W{w:02d}"

    @staticmethod
    def _safe_str_col(row, col: str, default: str = "") -> str:
        """pandas Series'ten sütun değerini güvenli oku; sütun yoksa default döndür."""
        if col in row.index:
            val = row[col]
            return str(val).strip() if not pd.isna(val) else default
        return default

    # ── Yükleme ──────────────────────────────────────────────────────────────

    def _load(self):
        xl = pd.ExcelFile(self.path)

        # ── Nodes ────────────────────────────────────────────────────────────
        for _, r in pd.read_excel(xl, "Nodes").iterrows():
            n = Node(
                str(r.node_id).strip(), str(r.node_name).strip(),
                str(r.country).strip(),  str(r.node_type).strip(),
                float(r.lat), float(r.lon),
                float(r.handling_day), float(r.handling_cost),
            )
            self.nodes[n.node_id] = n

        # ── Parameters ───────────────────────────────────────────────────────
        for _, r in pd.read_excel(xl, "Parameters").iterrows():
            self.params[str(r.parameter).strip()] = r.value

        self.plan_start = self._parse_date(self.params["planning_start_date"])
        self.plan_year  = int(self.params.get("planning_year", self.plan_start.year))
        self.phi_r      = cbam_phi(self.plan_year)   # FIX: yıl < 2026 → 0.0
        self.lambda_kg  = float(self.params.get("transport_carbon_price_per_kg", 0))
        self.de_minimis = float(self.params.get("cbam_de_minimis_threshold_ton", 50))

        mu_land = float(self.params.get("land_distance_multiplier", 1.15))
        mu_sea  = float(self.params.get("sea_distance_multiplier",  1.08))
        mu_air  = float(self.params.get("air_distance_multiplier",  1.00))
        self._mu = {"road": mu_land, "rail": mu_land, "sea": mu_sea, "air": mu_air}

        # ── Sources ───────────────────────────────────────────────────────────
        for _, r in pd.read_excel(xl, "Sources").iterrows():
            sid = str(r.source_node).strip()
            self.sources[sid] = Source(
                sid,
                float(r.capacity_ton),
                float(r.paid_carbon_price_per_tco2),
            )

        # ── Orders ───────────────────────────────────────────────────────────
        for _, r in pd.read_excel(xl, "Orders").iterrows():
            oid = str(r.order_id).strip()
            dd  = self._parse_date(r.deadline_date)
            H   = (dd.date() - self.plan_start.date()).days
            qty = float(r.tons)
            # FIX: product_type güvenli okuma ('product_type' in r.index ile)
            product = self._safe_str_col(r, "product_type", default="")
            self.orders[oid] = Order(
                oid,
                str(r.destination_node).strip(),
                qty, H,
                float(r.embedded_emission_tco2_per_ton),
                float(r.covered_emission_tco2_per_ton),
                float(r.benchmark_emission_tco2_per_ton),
                qty >= self.de_minimis,
                product,
            )

        # ── Modes ─────────────────────────────────────────────────────────────
        for _, r in pd.read_excel(xl, "Modes").iterrows():
            mid = str(r["mode"]).strip()
            self.modes[mid] = Mode(
                mid,
                float(r.cost_per_tkm),
                float(r.speed_km_day),
                float(r.emission_kg_per_tkm),
                float(r.arc_capacity_ton),
            )

        # ── Mode Schedules ────────────────────────────────────────────────────
        for _, r in pd.read_excel(xl, "Mode_Schedules").iterrows():
            mid = str(r["mode"]).strip()
            self.schedules[mid] = {
                "type":    str(r.schedule_type).strip(),
                "n":       None if pd.isna(r.interval_days) else int(r.interval_days),
                "weekday": None if pd.isna(r.allowed_weekday)
                           else str(r.allowed_weekday).strip().lower(),
            }

        # ── CBAM Prices ───────────────────────────────────────────────────────
        for _, r in pd.read_excel(xl, "CBAM_Prices").iterrows():
            self.cbam_prices[str(r.cbam_period).strip()] = float(r.certificate_price_per_tco2)
        # Listede olmayan haftalar için son bilinen fiyatı kullan
        self._last_cbam_price = self.cbam_prices[sorted(self.cbam_prices)[-1]]

        # ── Arcs — birim maliyet önhesaplı ───────────────────────────────────
        for _, r in pd.read_excel(xl, "Arcs").iterrows():
            frm = str(r.from_node).strip()
            to  = str(r.to_node).strip()
            mid = str(r["mode"]).strip()
            if frm not in self.nodes or to not in self.nodes or mid not in self.modes:
                continue
            ni   = self.nodes[frm]
            nj   = self.nodes[to]
            mo   = self.modes[mid]
            dist = self._haversine(ni.lat, ni.lon, nj.lat, nj.lon) * self._mu[mid]
            tau  = max(1, math.ceil(dist / mo.speed_km_day + nj.handling_day))
            # Önhesaplı: taşıma maliyeti + emisyon maliyeti + varış elleçleme
            unit = (mo.cost_per_tkm * dist
                    + self.lambda_kg * mo.emission_kg_per_tkm * dist
                    + nj.handling_cost)
            arc = Arc(frm, to, mid, dist, int(tau), unit)
            self.arcs.append(arc)
            self.out_arcs[frm].append(arc)
            self.arc_index[(frm, to, mid)] = arc   # FIX-v13-4: O(1) lookup

        self.source_ids = list(self.sources.keys())
        self.original_deadlines = {oid: o.deadline for oid, o in self.orders.items()}
        self.max_H      = max(o.deadline for o in self.orders.values())

    # ── Sefer takvimi ─────────────────────────────────────────────────────────

    def available(self, mode_id: str, day: int) -> bool:
        """Mod m, planlama günü day'de sefere açık mı?"""
        s   = self.schedules.get(mode_id, {"type": "daily"})
        typ = s["type"]
        if typ == "daily":
            return True
        if typ in ("every_n_days", "every_n"):
            return day % (s.get("n") or 1) == 0
        if typ in ("weekday_only", "weekday"):
            return (self.plan_start + timedelta(days=day)).strftime("%A").lower() \
                   == (s.get("weekday") or "")
        return True

    # ── CBAM birim maliyeti ───────────────────────────────────────────────────

    def cbam_unit(self, order: Order, source_id: str, arrive_day: int
                  ) -> Tuple[float, str]:
        """
        Sipariş k'nın source_id kaynağından gönderilip arrive_day'de varması
        durumundaki CBAM birim maliyetini ve ISO hafta etiketini döndürür.
        Denklem (2): c^CBAM_kst = max(0, φ_r·max(0,E^emb-Ē)·p^CBAM_w - p^paid·E^cov)
        """
        week = self._iso_week(self.plan_start + timedelta(days=arrive_day))
        pw   = self.cbam_prices.get(week, self._last_cbam_price)
        if not order.cbam_subject:
            return 0.0, week
        ps    = self.sources[source_id].paid_carbon
        gross = self.phi_r * max(0.0, order.emb_em - order.bench_em) * pw
        ded   = ps * order.cov_em
        return max(0.0, gross - ded), week


# ─────────────────────────────────────────────────────────────────────────────
# KAPASİTE TAKİBİ
# ─────────────────────────────────────────────────────────────────────────────

def rebuild_capacity(inst: Instance, sol: Solution
                     ) -> Tuple[Dict[str, float], Dict[Tuple, float]]:
    """Mevcut çözümden kaynak ve yay kapasitesi kullanımını yeniden hesapla."""
    src_used: Dict[str, float]  = defaultdict(float)
    arc_used: Dict[Tuple, float] = defaultdict(float)
    for oid, asgn in sol.assignments.items():
        q = inst.orders[oid].qty
        src_used[asgn.source] += q
        for leg in asgn.legs:
            arc_used[(leg.frm, leg.to, leg.mode_id, leg.depart_day)] += q
    return src_used, arc_used


def arc_overflow_penalty_rate(inst: Instance) -> float:
    """Soft arc capacity modunda 1 ton kapasite aşımı için ceza katsayısı (€ / ton)."""
    return float(getattr(inst, "arc_overflow_penalty_eur_per_ton", 5000.0))


def compute_arc_overflow_penalty(inst: Instance, arc_used: Dict[Tuple, float]) -> float:
    """Çözüm seviyesinde toplam yay kapasite aşımı cezası."""
    if not getattr(inst, "soft_arc_capacity", False):
        return 0.0
    rate = arc_overflow_penalty_rate(inst)
    penalty = 0.0
    for (_, _, mid, _), used in arc_used.items():
        cap = inst.modes[mid].arc_capacity
        if used > cap + EPS:
            penalty += rate * (used - cap)
    return penalty


def arc_overflow_rows(inst: Instance, arc_used: Dict[Tuple, float], tol: float = EPS_REPORT) -> List[dict]:
    """Kapasitesi aşılmış yay-gün kayıtlarını rapor satırı olarak döndür."""
    rows = []
    for (frm, to, mid, day), used in sorted(arc_used.items()):
        cap = inst.modes[mid].arc_capacity
        overflow = used - cap
        if overflow > tol:
            rows.append({
                "from_node": frm,
                "to_node": to,
                "mode": mid,
                "depart_day": day,
                "used_ton": used,
                "capacity_ton": cap,
                "overflow_ton": overflow,
            })
    return rows


def update_strict_feasibility_metrics(inst: Instance, sol: Solution,
                                      arc_used: Dict[Tuple, float]) -> None:
    """Strict fizibilite metriklerini Solution üzerine güvenli şekilde yaz.
    """
    rows = arc_overflow_rows(inst, arc_used)
    sol.overflow_arc_count = len(rows)
    sol.total_overflow_ton = sum(r["overflow_ton"] for r in rows)

    original_deadlines = getattr(inst, "original_deadlines", {})
    lateness_values = []
    for oid, asgn in sol.assignments.items():
        dl = original_deadlines.get(oid, inst.orders[oid].deadline)
        late = max(0, asgn.arrive_day - dl)
        if late > EPS_REPORT:
            lateness_values.append(late)
    sol.deadline_violation_count = len(lateness_values)
    sol.total_lateness_days = float(sum(lateness_values))
    sol.max_lateness_days = float(max(lateness_values) if lateness_values else 0.0)

    sol.strict_feasible = (sol.feasible
                           and sol.overflow_arc_count == 0
                           and sol.total_overflow_ton <= EPS_REPORT
                           and sol.deadline_violation_count == 0)


def marginal_arc_overflow_penalty(inst: Instance, order: Order, arc: Arc,
                                  day: int, arc_used: Dict[Tuple, float]) -> float:
    """
    Bir siparişin ilgili yay/gün kombinasyonuna eklenmesiyle oluşan marjinal
    kapasite aşımı cezası. Soft modda Dijkstra'nın kapasiteyi görmesini sağlar.
    """
    if not getattr(inst, "soft_arc_capacity", False):
        return 0.0
    key = (arc.frm, arc.to, arc.mode_id, day)
    used = arc_used.get(key, 0.0)
    cap = inst.modes[arc.mode_id].arc_capacity
    before = max(0.0, used - cap)
    after = max(0.0, used + order.qty - cap)
    return arc_overflow_penalty_rate(inst) * (after - before)


def arc_has_capacity(inst: Instance, order: Order, arc: Arc,
                     day: int, arc_used: Dict) -> bool:
    if getattr(inst, "soft_arc_capacity", False):
        return True
    key = (arc.frm, arc.to, arc.mode_id, day)
    return arc_used.get(key, 0.0) + order.qty <= inst.modes[arc.mode_id].arc_capacity + EPS


def apply_assignment(order: Order, asgn: Assignment,
                     src_used: Dict, arc_used: Dict):
    """Atamayı kapasite izleme yapılarına uygula."""
    src_used[asgn.source] += order.qty
    for leg in asgn.legs:
        arc_used[(leg.frm, leg.to, leg.mode_id, leg.depart_day)] += order.qty


def assignment_capacity_feasible(inst: Instance, order: Order, asgn: Assignment,
                                 src_used: Dict[str, float],
                                 arc_used: Dict[Tuple, float]) -> bool:
    """Önhesaplı Assignment mevcut kapasite durumunda hâlâ uygulanabilir mi?"""
    if src_used.get(asgn.source, 0.0) + order.qty > inst.sources[asgn.source].capacity + EPS:
        return False
    if not getattr(inst, "soft_arc_capacity", False):
        for leg in asgn.legs:
            key = (leg.frm, leg.to, leg.mode_id, leg.depart_day)
            if arc_used.get(key, 0.0) + order.qty > inst.modes[leg.mode_id].arc_capacity + EPS:
                return False
    return True


# ─────────────────────────────────────────────────────────────────────────────
# ROTA BULMA — TEK DİJKSTRA (tüm departure günleri)
# ─────────────────────────────────────────────────────────────────────────────

def find_best_assignment(
    inst:      Instance,
    order:     Order,
    source_id: str,
    src_used:  Dict[str, float],
    arc_used:  Dict[Tuple, float],
    noise:     float = 0.0,
    rng:       random.Random = None,
) -> Optional[Assignment]:
    """
    Tek bir Dijkstra çalıştırarak tüm departure günleri × tüm rotalar üzerinde
    en ucuz geçerli Assignment'ı döndürür.

    Durum uzayı = (node, day) → O(H) kat daha az tekrarlı hesap.
    Kaynakta beklemeye izin var (her gün ayrı başlangıç durumu).
    Ara düğümlerde bekleme yok (no-waiting — Kısıt 6).
    """
    if src_used.get(source_id, 0.0) + order.qty > inst.sources[source_id].capacity + EPS:
        return None

    rng = rng or random.Random()

    counter  = 0
    pq:       list = []
    dist_map: Dict[Tuple[str, int], float] = {}
    pred_map: Dict[Tuple[str, int], Tuple] = {}  # (node,day) → ((prev_node,prev_day), arc)

    # Tüm uygun departure günleri için başlangıç durumları ekle
    for dep in range(order.deadline + 1):
        state = (source_id, dep)
        dist_map[state] = 0.0
        heapq.heappush(pq, (0.0, counter, source_id, dep))
        counter += 1

    best_asgn:      Optional[Assignment] = None
    best_cost:      float = INF
    source_loading: float = order.qty * inst.nodes[source_id].handling_cost

    # FIX-1: time.time() döngü içinden çıkarıldı (pahalı sistem çağrısı).
    # FIX-2: max_expansions artırıldı; büyük ağda erken kesme optimal rotayı kaçırır.
    #        Pruning (cost + source_loading >= best_cost) zaten etkin bir erken çıkış sağlar.
    expansions = 0
    max_expansions = 200_000
    while pq:
        expansions += 1
        if expansions > max_expansions:
            return best_asgn  # güvenli kesme; varış bulunduysa döndür
        cost, _, node, day = heapq.heappop(pq)

        # Pruning: yay maliyeti + yükleme ≥ şimdiki en iyi toplam maliyet ise kes.
        # CBAM ≥ 0 olduğu için bu kesim kesinlikle güvenlidir.
        if cost + source_loading >= best_cost - EPS:
            break

        state = (node, day)
        if dist_map.get(state, INF) < cost - EPS:
            continue  # eski etiket, atla

        if node == order.dest:
            # Yolu geri izle
            legs: List[Leg] = []
            st = state
            while st in pred_map:
                (prev_node, prev_day), arc = pred_map[st]
                legs.append(Leg(arc.frm, arc.to, arc.mode_id,
                                prev_day, st[1], arc.dist_km, arc.unit_cost))
                st = (prev_node, prev_day)
            legs.reverse()

            dep_day = legs[0].depart_day if legs else day
            arr_day = day
            move_c  = source_loading + sum(order.qty * lg.unit_cost for lg in legs)
            cbam_u, week = inst.cbam_unit(order, source_id, arr_day)
            cbam_c  = order.qty * cbam_u
            total_c = move_c + cbam_c

            if total_c < best_cost:
                best_cost = total_c
                best_asgn = Assignment(
                    order.order_id, source_id, dep_day, arr_day, week,
                    legs, move_c, cbam_c, total_c,
                )
            continue  # varışta duruyoruz, genişletme yok

        # Komşu yayları genişlet
        for arc in inst.out_arcs.get(node, []):
            if not inst.available(arc.mode_id, day):
                continue
            if not arc_has_capacity(inst, order, arc, day, arc_used):
                continue
            arr = day + arc.travel_days
            if arr > order.deadline:
                continue

            # FIX: noise_factor alt sınır eklendi — negatif maliyet önlenir
            if noise > 0:
                noise_factor = max(0.01, 1.0 + rng.uniform(-noise, noise))
            else:
                noise_factor = 1.0

            step      = order.qty * arc.unit_cost * noise_factor
            step     += marginal_arc_overflow_penalty(inst, order, arc, day, arc_used)
            new_cost  = cost + step
            nxt_state = (arc.to, arr)

            if new_cost < dist_map.get(nxt_state, INF) - EPS:
                dist_map[nxt_state] = new_cost
                pred_map[nxt_state] = ((node, day), arc)
                counter += 1
                heapq.heappush(pq, (new_cost, counter, arc.to, arr))

    return best_asgn


def best_insert(inst: Instance, order: Order,
                src_used: Dict, arc_used: Dict,
                sources: Optional[List[str]] = None,
                noise: float = 0.0,
                rng: random.Random = None) -> Optional[Assignment]:
    """Tüm (veya verilen) kaynaklar arasından en ucuz Assignment'ı bul."""
    best: Optional[Assignment] = None
    for sid in (sources or inst.source_ids):
        a = find_best_assignment(inst, order, sid, src_used, arc_used, noise, rng)
        if a and (best is None or a.total_cost < best.total_cost):
            best = a
    return best


# ─────────────────────────────────────────────────────────────────────────────
# ÇÖZÜM DEĞERLENDİRME VE DOĞRULAMA
# ─────────────────────────────────────────────────────────────────────────────

def evaluate(inst: Instance, sol: Solution) -> Solution:
    """Tüm kısıtları kontrol et, objective ve strict/soft fizibilite metriklerini güncelle."""
    sol.strict_feasible = False
    sol.overflow_arc_count = 0
    sol.total_overflow_ton = 0.0
    sol.deadline_violation_count = 0
    sol.total_lateness_days = 0.0
    sol.max_lateness_days = 0.0
    sol.arc_overflow_penalty = 0.0

    if set(sol.assignments) != set(inst.orders):
        sol.obj = INF
        sol.base_obj = INF
        sol.feasible = False
        return sol

    src_used, arc_used = rebuild_capacity(inst, sol)

    for sid, src in inst.sources.items():
        if src_used.get(sid, 0.0) > src.capacity + EPS_REPORT:
            sol.obj = INF
            sol.base_obj = INF
            sol.feasible = False
            update_strict_feasibility_metrics(inst, sol, arc_used)
            return sol

    overflow_rows = arc_overflow_rows(inst, arc_used)

    if overflow_rows and not getattr(inst, "soft_arc_capacity", False):
        sol.obj = INF
        sol.base_obj = INF
        sol.feasible = False
        sol.overflow_arc_count = len(overflow_rows)
        sol.total_overflow_ton = sum(r["overflow_ton"] for r in overflow_rows)
        sol.strict_feasible = False
        return sol

    base_obj = sum(a.total_cost for a in sol.assignments.values())
    penalty = compute_arc_overflow_penalty(inst, arc_used)
    sol.base_obj = base_obj
    sol.arc_overflow_penalty = penalty
    sol.obj = base_obj + penalty
    sol.feasible = True
    update_strict_feasibility_metrics(inst, sol, arc_used)
    return sol

def validate(inst: Instance, sol: Solution) -> bool:
    """
    Çözümü bağımsız doğrula — her kısıtı açıkça assert et.
    Hata varsa AssertionError ile hangi sipariş ve kısıt olduğunu raporlar.
    """
    evaluate(inst, sol)
    assert sol.feasible, "Çözüm infeasible (evaluate başarısız)"

    src_used, arc_used = rebuild_capacity(inst, sol)

    # Kaynak kapasitesi
    for sid, src in inst.sources.items():
        assert src_used.get(sid, 0.0) <= src.capacity + EPS_REPORT, \
            f"Kaynak {sid} kapasitesi aşıldı: {src_used[sid]:.1f} > {src.capacity}"

    # Yay kapasitesi: soft_arc_capacity=True ise kapasite aşımı raporlanır ama
    # başlangıç çözümünü engellemez. Strict modda assertion çalışır.
    if not getattr(inst, "soft_arc_capacity", False):
        for (frm, to, mid, day), used in arc_used.items():
            cap = inst.modes[mid].arc_capacity
            assert used <= cap + EPS_REPORT, \
                f"Yay ({frm}→{to}, {mid}, gün {day}) kapasitesi aşıldı: {used:.1f} > {cap}"

    for oid, asgn in sol.assignments.items():
        order = inst.orders[oid]

        # Deadline
        assert asgn.arrive_day <= order.deadline, \
            f"{oid}: deadline aşıldı ({asgn.arrive_day} > {order.deadline})"

        # Kaynak geçerli
        assert asgn.source in inst.sources, \
            f"{oid}: geçersiz kaynak '{asgn.source}'"

        # Maliyet değerleri mantıklı  (FIX: eklendi)
        assert asgn.move_cost >= 0, f"{oid}: move_cost negatif ({asgn.move_cost:.4f})"
        assert asgn.cbam_cost >= 0, f"{oid}: cbam_cost negatif ({asgn.cbam_cost:.4f})"

        # Leg zinciri tutarlılığı
        prev_node = asgn.source
        prev_day  = asgn.depart_day
        for leg in asgn.legs:
            assert leg.frm == prev_node, \
                f"{oid}: yol kopukluğu ({leg.frm} != {prev_node})"
            assert leg.depart_day == prev_day, \
                f"{oid}: gün süreksizliği (leg.depart={leg.depart_day}, beklenen={prev_day})"
            assert inst.available(leg.mode_id, leg.depart_day), \
                f"{oid}: takvim ihlali (mod={leg.mode_id}, gün={leg.depart_day})"
            prev_node = leg.to
            prev_day  = leg.arrive_day

        assert prev_node == order.dest, \
            f"{oid}: varış düğümü yanlış ('{prev_node}' != '{order.dest}')"

    return True



# ─────────────────────────────────────────────────────────────────────────────
# ÖN FİZİBİLİTE TEŞHİSİ
# ─────────────────────────────────────────────────────────────────────────────

def _min_feasible_arrival_relaxed(inst: Instance, order: Order,
                                  max_extra_days: int = 90) -> Optional[Assignment]:
    """
    Siparişi mevcut deadline yerine deadline + max_extra_days ile dener.
    Amaç çözmek değil; 'bu sipariş en erken kaçıncı günde fiziksel olarak gelebilir?'
    sorusunu cevaplamak. Kapasite boş varsayılır, fakat unsplittable arc kapasitesi,
    sefer takvimi ve no-waiting korunur.
    """
    relaxed = replace(order, deadline=order.deadline + max_extra_days)
    zero_src: Dict[str, float] = defaultdict(float)
    zero_arc: Dict[Tuple, float] = defaultdict(float)
    best: Optional[Assignment] = None
    for sid in inst.source_ids:
        a = find_best_assignment(inst, relaxed, sid, zero_src, zero_arc, 0.0, random.Random(0))
        if a and (best is None or a.arrive_day < best.arrive_day or
                  (a.arrive_day == best.arrive_day and a.total_cost < best.total_cost)):
            best = a
    return best


def diagnose_instance_feasibility(inst: Instance,
                                  write_csv: bool = True,
                                  out_csv: str = "infeasible_orders_diagnosis.csv"
                                  ) -> Tuple[bool, pd.DataFrame]:
    """
    Başlangıç çözümü üretmeden önce hızlı fizibilite teşhisi yapar.

    Kontrol 1: Toplam talep toplam kaynak kapasitesini aşıyor mu?
    Kontrol 2: Her sipariş, kapasite boşken en az bir kaynaktan deadline içinde
               fiziksel olarak teslim edilebiliyor mu?
    Kontrol 3: Feasible değilse mevcut mod/yay kapasiteleri altında en erken
               kaçıncı günde gelebileceğini tahmin eder.

    Not: Bu kontrol global feasibility kanıtı değildir; ancak tekil sipariş infeasible ise
    tüm problem kesinlikle infeasible'dır.
    """
    rows = []
    total_qty = sum(o.qty for o in inst.orders.values())
    total_cap = sum(s.capacity for s in inst.sources.values())

    zero_src: Dict[str, float] = defaultdict(float)
    zero_arc: Dict[Tuple, float] = defaultdict(float)

    for oid, order in inst.orders.items():
        feasible_sources = []
        for sid in inst.source_ids:
            a = find_best_assignment(inst, order, sid, zero_src, zero_arc, 0.0, random.Random(0))
            if a is not None:
                feasible_sources.append((sid, a.arrive_day, a.total_cost))

        if feasible_sources:
            best_src, best_arr, best_cost = sorted(feasible_sources, key=lambda x: (x[1], x[2]))[0]
            rows.append({
                "order_id": oid,
                "destination": order.dest,
                "tons": order.qty,
                "deadline_day": order.deadline,
                "status": "individually_feasible",
                "feasible_source_count": len(feasible_sources),
                "earliest_arrival_day": best_arr,
                "min_extra_days_needed": max(0, best_arr - order.deadline),
                "suggested_source": best_src,
                "note": "OK",
            })
        else:
            relaxed_best = _min_feasible_arrival_relaxed(inst, order, max_extra_days=90)
            if relaxed_best is None:
                note = "No route even with +90 days; check network connectivity/arcs."
                earliest = None
                extra = None
                src = None
            else:
                earliest = relaxed_best.arrive_day
                extra = max(0, earliest - order.deadline)
                src = relaxed_best.source
                note = "Deadline too tight under unsplittable arc capacity and schedules."
            rows.append({
                "order_id": oid,
                "destination": order.dest,
                "tons": order.qty,
                "deadline_day": order.deadline,
                "status": "individually_infeasible",
                "feasible_source_count": 0,
                "earliest_arrival_day": earliest,
                "min_extra_days_needed": extra,
                "suggested_source": src,
                "note": note,
            })

    df = pd.DataFrame(rows)
    infeas = df[df["status"] == "individually_infeasible"].copy()

    if write_csv:
        df.to_csv(out_csv, index=False)

    ok = (total_qty <= total_cap + EPS) and infeas.empty

    # v5: Bu fonksiyon artık hata fırlatmaz.
    # Strict instance infeasible ise main() bu tabloyu kullanarak güvenli bir
    # deadline-relaxation uygular ve ALNS'in gerçekten çalışmasını sağlar.
    return ok, df

# ─────────────────────────────────────────────────────────────────────────────
# BAŞLANGIÇ ÇÖZÜMÜ
# ─────────────────────────────────────────────────────────────────────────────

def _order_cbam_exposure(order: Order) -> float:
    """Siparişin potansiyel CBAM riski: tonaj × benchmark üstü gömülü emisyon."""
    return order.qty * max(0.0, order.emb_em - order.bench_em)


def _count_feasible_sources(inst: Instance, order: Order) -> int:
    """
    Kapasite boşken kaç kaynaktan fiziksel/takvimsel rota bulunabilir?
    Maliyet hesaplamaz; yalnızca sefer takvimi + süre + no-waiting kontrol eder.
    BFS ile hızlıca çalışır.
    """
    cnt = 0
    for sid in inst.source_ids:
        seen  = set()
        # FIX: 'q' → 'bfs_q' (builtin çakışmasını önle)
        bfs_q = [(sid, d) for d in range(order.deadline + 1)]
        head  = 0
        found = False
        while head < len(bfs_q) and not found:
            node, day = bfs_q[head]
            head += 1
            if (node, day) in seen:
                continue
            seen.add((node, day))
            if node == order.dest:
                found = True
                break
            for arc in inst.out_arcs.get(node, []):
                if not inst.available(arc.mode_id, day):
                    continue
                arr = day + arc.travel_days
                if arr <= order.deadline:
                    bfs_q.append((arc.to, arr))
        if found:
            cnt += 1
    return cnt


def _criticality_sequences(inst: Instance) -> List[List[Order]]:
    """Farklı önceliklendirme kriterlerine göre deterministik sipariş sıralamaları üret."""
    orders = list(inst.orders.values())
    fsc    = {o.order_id: _count_feasible_sources(inst, o) for o in orders}

    def critical_key(o: Order):
        # Az alternatif kaynak + erken deadline + yüksek tonaj + yüksek CBAM riski
        return (fsc[o.order_id], o.deadline, -o.qty, -_order_cbam_exposure(o))

    return [
        sorted(orders, key=critical_key),
        sorted(orders, key=lambda o: (o.deadline, fsc[o.order_id], -o.qty)),
        sorted(orders, key=lambda o: (-o.qty, o.deadline)),
        sorted(orders, key=lambda o: (-_order_cbam_exposure(o), o.deadline)),
        sorted(orders, key=lambda o: (o.dest, o.deadline, -o.qty)),
    ]


def _precompute_empty_candidates(inst: Instance, rng: random.Random
                                 ) -> Dict[str, List[Assignment]]:
    """
    Kapasite boşken her (sipariş, kaynak) çiftinin en iyi rotasını önhesapla.
    Başlangıç inşası sırasında tekrarlı Dijkstra çağrılarını önler.
    """
    cache:    Dict[str, List[Assignment]] = {}
    zero_src: Dict[str, float]  = defaultdict(float)
    zero_arc: Dict[Tuple, float] = defaultdict(float)

    for oid, order in inst.orders.items():
        cands = []
        for sid in inst.source_ids:
            a = find_best_assignment(inst, order, sid, zero_src, zero_arc, 0.0, rng)
            if a is not None:
                cands.append(a)
        cands.sort(key=lambda x: x.total_cost)
        cache[oid] = cands
    return cache


def _candidate_from_cache_or_search(inst: Instance, order: Order,
                                    src_used: Dict[str, float],
                                    arc_used: Dict[Tuple, float],
                                    cache: Dict[str, List[Assignment]],
                                    rng: random.Random,
                                    noise: float = 0.0) -> List[Assignment]:
    """
    Önce cache'den feasible adayları al.
    Cache boş kalırsa (kapasite çakışması) gerçek zamanlı Dijkstra çalıştır.
    """
    feasible = [
        a for a in cache.get(order.order_id, [])
        if assignment_capacity_feasible(inst, order, a, src_used, arc_used)
    ]
    if not feasible:
        for sid in inst.source_ids:
            a = find_best_assignment(inst, order, sid, src_used, arc_used, noise, rng)
            if a is not None:
                feasible.append(a)
    feasible.sort(key=lambda x: x.total_cost)
    return feasible


def _best_fit_construct(inst: Instance, rng: random.Random,
                        cache: Dict[str, List[Assignment]],
                        alpha: float = 8.0) -> Optional[Solution]:
    """
    Best-fit decreasing construction:
    Büyük siparişleri önce yerleştirir (azalan tonaj sırası).
    Kaynak seçiminde rem_frac (kalan kapasite oranı) yüksek olan kaynaklar
    cezalandırılır; böylece model yarı dolu kaynakları tercih eder
    ve kapasite daha dengeli dolar. alpha arttıkça bu tercih kuvvetlenir.
    """
    sol      = Solution()
    src_used: Dict[str, float]  = defaultdict(float)
    arc_used: Dict[Tuple, float] = defaultdict(float)

    for order in sorted(inst.orders.values(), key=lambda o: (-o.qty, o.deadline)):
        opts = []
        for a in cache.get(order.order_id, []):
            if not assignment_capacity_feasible(inst, order, a, src_used, arc_used):
                continue
            rem      = inst.sources[a.source].capacity - src_used.get(a.source, 0.0) - order.qty
            rem_frac = rem / max(inst.sources[a.source].capacity, EPS)
            # Çok boş kaynak (rem_frac yüksek) → score yüksek → tercih edilmez
            score = a.total_cost * (1.0 + alpha * rem_frac)
            opts.append((score, a))

        if not opts:
            a = best_insert(inst, order, src_used, arc_used, noise=0.0, rng=rng)
            if a is None:
                return None
        else:
            opts.sort(key=lambda x: x[0])
            a = opts[0][1]

        sol.assignments[order.order_id] = a
        apply_assignment(order, a, src_used, arc_used)

    return evaluate(inst, sol)


def _regret_construct(inst: Instance, rng: random.Random,
                      cache: Dict[str, List[Assignment]],
                      noise: float = 0.0) -> Optional[Solution]:
    """
    Regret-2 başlangıç inşası:
    Her adımda kalan siparişler için en iyi iki atama hesaplanır.
    Regret = (2. en iyi - 1. en iyi): bu fark en yüksek olan önce atanır.
    Sıkı deadline'lı siparişler eşitlikte öncele.
    """
    sol       = Solution()
    src_used: Dict[str, float]  = defaultdict(float)
    arc_used: Dict[Tuple, float] = defaultdict(float)
    remaining = set(inst.orders)

    while remaining:
        best_oid    = None
        best_asgn   = None
        best_regret = -INF

        for oid in list(remaining):
            order = inst.orders[oid]
            cands = _candidate_from_cache_or_search(
                inst, order, src_used, arc_used, cache, rng, noise)
            if not cands:
                continue
            regret = (cands[1].total_cost - cands[0].total_cost) if len(cands) > 1 else 1e9
            # Sıkı deadline → küçük tiebreaker eklenir
            regret += 1e-3 / max(1, order.deadline + 1)
            if regret > best_regret:
                best_regret = regret
                best_oid    = oid
                best_asgn   = cands[0]

        if best_oid is None or best_asgn is None:
            return None

        order = inst.orders[best_oid]
        sol.assignments[best_oid] = best_asgn
        apply_assignment(order, best_asgn, src_used, arc_used)
        remaining.remove(best_oid)

    return evaluate(inst, sol)


def _construct_with_sequence(inst: Instance, seq: List[Order],
                              rng: random.Random,
                              cache: Dict[str, List[Assignment]],
                              noise: float = 0.0,
                              restricted_choice: bool = False) -> Optional[Solution]:
    """
    Verilen sırayla greedy inşa yapar.
    restricted_choice=True ise ilk 3 aday arasından ağırlıklı rastgele seçer.
    Yük dengeleme, deadline slack ve CBAM görünürlüğünü skorlamaya katar.
    """
    sol      = Solution()
    src_used: Dict[str, float]  = defaultdict(float)
    arc_used: Dict[Tuple, float] = defaultdict(float)

    for order in seq:
        raw_cands = _candidate_from_cache_or_search(
            inst, order, src_used, arc_used, cache, rng, noise)
        if not raw_cands:
            return None

        cands = []
        for a in raw_cands:
            util         = (src_used.get(a.source, 0.0) + order.qty) \
                           / max(inst.sources[a.source].capacity, EPS)
            slack        = order.deadline - a.arrive_day
            slack_pen    = 1.0 + 0.015 / max(1, slack + 1)
            util_pen     = 1.0 + 0.20 * util ** 3
            score        = a.total_cost * util_pen * slack_pen
            cands.append((score, a))

        cands.sort(key=lambda x: x[0])

        if restricted_choice and len(cands) > 1:
            top     = cands[:min(3, len(cands))]
            weights = [1.0 / (rank + 1) for rank in range(len(top))]
            asgn    = rng.choices([a for _, a in top], weights=weights, k=1)[0]
        else:
            asgn = cands[0][1]

        sol.assignments[order.order_id] = asgn
        apply_assignment(order, asgn, src_used, arc_used)

    return evaluate(inst, sol)


def _source_priority(inst: Instance, order: Order, src_used: Dict[str, float]) -> List[str]:
    """
    Kaynakları hızlı başlangıç çözümü için sıralar.
    Önce kapasitesi yeten ve destination'a coğrafi olarak daha yakın/tight-fit olan
    kaynaklar denenir. Bu, 150 siparişlik inputta tüm kaynakları gereksiz yere
    Dijkstra ile taramayı engeller.
    """
    dest_node = inst.nodes[order.dest]
    rows = []
    for sid, src in inst.sources.items():
        rem = src.capacity - src_used.get(sid, 0.0)
        if rem + EPS < order.qty:
            continue
        sn = inst.nodes[sid]
        geo = inst._haversine(sn.lat, sn.lon, dest_node.lat, dest_node.lon)
        tight_fit = (rem - order.qty) / max(src.capacity, EPS)
        rows.append((geo + 250.0 * tight_fit, sid))
    rows.sort()
    return [sid for _, sid in rows]


def _fast_sequences(inst: Instance) -> List[List[Order]]:
    """
    Büyük inputta hızlı sıralamalar. inst._feasible_source_count'tan gelen
    gerçek teşhis verisini kullanır; tek/az kaynak alternatifi olan siparişler
    önce yerleştirilir.
    """
    orders = list(inst.orders.values())
    fsc = getattr(inst, "_feasible_source_count", {})
    def scarce(o: Order) -> int:
        return int(fsc.get(o.order_id, 999))
    return [
        # v10: Sıkı deadline'lar önce. O028/O037 gibi erken deadline'lı
        # siparişlerin sona kalıp kapasite/takvim çakışmasına düşmesini engeller.
        sorted(orders, key=lambda o: (o.deadline, scarce(o), -o.qty, -_order_cbam_exposure(o))),
        sorted(orders, key=lambda o: (scarce(o), o.deadline, -o.qty, -_order_cbam_exposure(o))),
        sorted(orders, key=lambda o: (scarce(o), -o.qty, o.deadline)),
        sorted(orders, key=lambda o: (scarce(o), -_order_cbam_exposure(o), o.deadline, -o.qty)),
        sorted(orders, key=lambda o: (o.dest, scarce(o), o.deadline, -o.qty)),
    ]


def _find_with_deadline_recovery(inst: Instance, order: Order,
                                 src_used: Dict[str, float],
                                 arc_used: Dict[Tuple, float],
                                 rng: random.Random,
                                 max_extra: int = 30) -> Optional[Assignment]:
    """
    Başlangıç inşasında deadline dolayısıyla atanamayan siparişler için
    deadline'ı geçici olarak +max_extra gün açarak fiziksel rota arar.
    """
    old_deadline = order.deadline
    relaxed = replace(order, deadline=old_deadline + max_extra)
    candidates = []
    for sid in _source_priority(inst, relaxed, src_used):
        a = find_best_assignment(inst, relaxed, sid, src_used, arc_used, 0.0, rng)
        if a is not None:
            candidates.append(a)
    if not candidates:
        return None
    candidates.sort(key=lambda x: x.total_cost)
    best = candidates[0]
    # FIX-3: new_deadline hesaplanıp sadece gerekli olduğunda (arrive_day > old_deadline)
    # inst.orders güncelleniyor. Aynı sipariş için tekrar çağrıda çift mutasyon önlenir.
    new_deadline = max(old_deadline, best.arrive_day)
    if new_deadline != old_deadline:
        inst.orders[order.order_id] = replace(order, deadline=new_deadline)
        if not hasattr(inst, "_dynamic_deadline_relaxations"):
            inst._dynamic_deadline_relaxations = []
        inst._dynamic_deadline_relaxations.append({
            "order_id": order.order_id,
            "old_deadline_day": old_deadline,
            "new_deadline_day": new_deadline,
            "added_days": new_deadline - old_deadline,
            "reason": "global_capacity_or_schedule_conflict_during_initial_construction",
        })
    return best


def _fast_probe(inst: Instance, order: "Order", src_used: Dict[str, float],
                arc_used: Dict[Tuple, float], source_list: List[str],
                cands: List[Tuple[float, "Assignment"]], seen: set,
                noise: float, rng: random.Random) -> None:

    for sid in source_list:
        if sid in seen:
            continue
        seen.add(sid)
        a = find_best_assignment(inst, order, sid, src_used, arc_used, noise, rng)
        if a is None:
            continue
        util  = (src_used.get(a.source, 0.0) + order.qty) / max(inst.sources[a.source].capacity, EPS)
        slack = order.deadline - a.arrive_day
        score = a.total_cost * (1.0 + 0.15 * util ** 3 + 0.01 / max(1, slack + 1))
        cands.append((score, a))


def _fast_construct(inst: Instance, seq: List[Order], rng: random.Random,
                    noise: float = 0.0, restricted_choice: bool = False,
                    source_probe: int = 5) -> Optional[Solution]:
    """
    Büyük instance için hızlı başlangıç kurucu. Her siparişte önce en umut verici
    kaynakları dener; gerekirse tüm kaynaklara genişler. Feasible bir aday bulunca
    en düşük skorla yerleştirir.
    """
    sol = Solution()
    src_used: Dict[str, float] = defaultdict(float)
    arc_used: Dict[Tuple, float] = defaultdict(float)

    for order in seq:
        srcs = _source_priority(inst, order, src_used)
        if not srcs:
            return None

        candidates: List[Tuple[float, Assignment]] = []
        tried: set = set()

        # FIX-v13-8: _probe artık modül-seviyesi _fast_probe; döngüde closure yaratılmıyor.
        _fast_probe(inst, order, src_used, arc_used,
                    srcs[:source_probe], candidates, tried, noise, rng)
        if not candidates:
            _fast_probe(inst, order, src_used, arc_used,
                        srcs[source_probe:], candidates, tried, noise, rng)
        if not candidates:
            # Orijinal deadline'ları kesin korumak için dinamik deadline
            # recovery kapalıdır. Aday yoksa bu kurucu başarısız olur ve
            # initial_solution farklı sıralama/denemeye geçer.
            return None

        candidates.sort(key=lambda x: x[0])
        if restricted_choice and len(candidates) > 1:
            top = candidates[:min(3, len(candidates))]
            weights = [1.0/(i+1) for i in range(len(top))]
            asgn = rng.choices([a for _, a in top], weights=weights, k=1)[0]
        else:
            asgn = candidates[0][1]

        sol.assignments[order.order_id] = asgn
        apply_assignment(order, asgn, src_used, arc_used)

    return evaluate(inst, sol)



def _backtracking_construct(inst: Instance, rng: random.Random,
                            time_limit: float = 10.0,
                            node_limit: int = 200_000) -> Optional[Solution]:
    """
    Robust fallback başlangıç kurucusu.

    Küçük veya kaynak kapasitesi çok sıkı instance'larda greedy/best-fit sıralama
    hatalı paketleme yapabilir. Bu fallback, sipariş-kaynak atamalarını sınırlı
    backtracking ile dener; her denemede mevcut kaynak ve yay kapasitesine göre
    route'u tekrar doğrular. Büyük instance'larda yalnızca son çare olarak kısa
    süre çalıştırılmalıdır.
    """
    t0 = time.time()
    sol = Solution()
    src_used: Dict[str, float] = defaultdict(float)
    arc_used: Dict[Tuple, float] = defaultdict(float)

    # Kapasite sıkılığına göre önce büyük ve az esnek siparişleri yerleştir.
    empty_src: Dict[str, float] = defaultdict(float)
    empty_arc: Dict[Tuple, float] = defaultdict(float)
    cand_cache: Dict[str, List[Assignment]] = {}
    for oid, order in inst.orders.items():
        cands: List[Assignment] = []
        for sid in inst.source_ids:
            a = find_best_assignment(inst, order, sid, empty_src, empty_arc, 0.0, rng)
            if a is not None:
                cands.append(a)
        cands.sort(key=lambda a: a.total_cost)
        cand_cache[oid] = cands

    if any(len(v) == 0 for v in cand_cache.values()):
        return None

    remaining = set(inst.orders.keys())
    best_sol: Optional[Solution] = None
    best_obj = INF
    nodes_seen = 0

    def _remove_assignment(order: Order, asgn: Assignment):
        src_used[asgn.source] -= order.qty
        if abs(src_used[asgn.source]) < EPS:
            src_used.pop(asgn.source, None)
        for leg in asgn.legs:
            key = (leg.frm, leg.to, leg.mode_id, leg.depart_day)
            arc_used[key] -= order.qty
            if abs(arc_used[key]) < EPS:
                arc_used.pop(key, None)

    def _candidate_options(oid: str) -> List[Assignment]:
        order = inst.orders[oid]
        opts: List[Assignment] = []
        seen_sources = set()
        # Önce empty-capacity adaylarının kaynak sırasını kullan.
        for cached in cand_cache.get(oid, []):
            sid = cached.source
            if sid in seen_sources:
                continue
            seen_sources.add(sid)
            if src_used.get(sid, 0.0) + order.qty > inst.sources[sid].capacity + EPS:
                continue
            if assignment_capacity_feasible(inst, order, cached, src_used, arc_used):
                opts.append(cached)
            else:
                # Aynı kaynak için mevcut yay kapasitesine göre alternatif rota ara.
                fresh = find_best_assignment(inst, order, sid, src_used, arc_used, 0.0, rng)
                if fresh is not None and assignment_capacity_feasible(inst, order, fresh, src_used, arc_used):
                    opts.append(fresh)
        opts.sort(key=lambda a: a.total_cost)
        return opts

    def _select_next_oid() -> Tuple[Optional[str], List[Assignment]]:
        best_oid = None
        best_opts: List[Assignment] = []
        best_key = None
        for oid in remaining:
            opts = _candidate_options(oid)
            if not opts:
                return oid, []
            order = inst.orders[oid]
            key = (len(opts), order.deadline, -order.qty)
            if best_key is None or key < best_key:
                best_key = key
                best_oid = oid
                best_opts = opts
        return best_oid, best_opts

    def _dfs(partial_cost: float) -> bool:
        nonlocal best_sol, best_obj, nodes_seen
        nodes_seen += 1
        if nodes_seen > node_limit or time.time() - t0 > time_limit:
            return False
        if partial_cost >= best_obj - EPS:
            return False
        if not remaining:
            cand = evaluate(inst, copy.deepcopy(sol))
            if cand.feasible and cand.obj < best_obj:
                best_sol = cand
                best_obj = cand.obj
                return True
            return False

        oid, opts = _select_next_oid()
        if oid is None or not opts:
            return False

        order = inst.orders[oid]
        remaining.remove(oid)
        found_any = False
        # En ucuz birkaç seçeneği dene; küçük instance'ta tüm kaynaklar denenmiş olur.
        for asgn in opts:
            if not assignment_capacity_feasible(inst, order, asgn, src_used, arc_used):
                continue
            sol.assignments[oid] = asgn
            apply_assignment(order, asgn, src_used, arc_used)
            if _dfs(partial_cost + asgn.total_cost):
                found_any = True
            _remove_assignment(order, asgn)
            sol.assignments.pop(oid, None)
        remaining.add(oid)
        return found_any

    _dfs(0.0)
    return best_sol

def initial_solution(inst: Instance, rng: random.Random) -> Solution:
    """
    Büyük instance'lar için önce hızlı kurucular denenir. Eğer bu kurucular
    kaynak kapasitesi sıkı olan küçük/orta instance'larda çözüm bulamazsa,
    kodda zaten bulunan fakat v10 başlangıç akışında çağrılmayan daha sağlam
    candidate-cache tabanlı best-fit/regret fallback devreye girer.

    Bu değişiklik çözüm modelini gevşetmez:
      - deadline relaxation yapmaz,
      - kaynak kapasitesini gevşetmez,
      - yay kapasitesini gevşetmez.
    Sadece başlangıç çözümü kurma stratejisini güçlendirir.
    """
    base_sequences = _fast_sequences(inst)
    best_sol: Optional[Solution] = None

    def _try(sol: Optional[Solution]):
        nonlocal best_sol
        if sol and sol.feasible and (best_sol is None or sol.obj < best_sol.obj):
            best_sol = copy.deepcopy(sol)

    # 1) Büyük instance için hızlı kurucular
    for seq in base_sequences:
        _try(_fast_construct(
            inst, seq, rng,
            noise=0.0,
            restricted_choice=False,
            source_probe=int(getattr(inst, "source_probe", 2)),
        ))

    # 2) Küçük blok shuffle ile çeşitlilik
    N_TRIES = int(getattr(inst, "initial_random_tries", 2))
    seed_seq = base_sequences[0]
    for _ in range(N_TRIES):
        seq = seed_seq[:]
        window = max(3, min(8, len(seq)))
        for start in range(0, len(seq), window):
            block = seq[start:start + window]
            rng.shuffle(block)
            seq[start:start + window] = block
        if rng.random() < 0.3:
            rng.shuffle(seq)
        noise = rng.uniform(0.0, 0.03)
        _try(_fast_construct(
            inst, seq, rng,
            noise=noise,
            restricted_choice=True,
            source_probe=int(getattr(inst, "source_probe", 2)),
        ))

    # 3) v10.1 fallback: küçük/sıkı kapasite instance'ları için sağlam kurucular.
    #    30node_20order_input.xlsx'te hızlı greedy kaynak kapasitesini kötü paketleyip
    #    başlangıç çözümü bulamıyordu; best-fit fallback feasible çözümü buluyor.
    if best_sol is None:
        cache = _precompute_empty_candidates(inst, rng)

        _try(_best_fit_construct(inst, rng, cache, alpha=8.0))
        _try(_best_fit_construct(inst, rng, cache, alpha=2.0))
        _try(_regret_construct(inst, rng, cache, noise=0.0))

        for seq in _criticality_sequences(inst):
            _try(_construct_with_sequence(
                inst, seq, rng, cache,
                noise=0.0,
                restricted_choice=False,
            ))

        # Son çare: az gürültülü birkaç randomized construction.
        for _ in range(max(5, min(30, len(inst.orders)))):
            seq = list(inst.orders.values())
            rng.shuffle(seq)
            _try(_construct_with_sequence(
                inst, seq, rng, cache,
                noise=rng.uniform(0.0, 0.03),
                restricted_choice=True,
            ))

    # 4) Son ve robust fallback: küçük/sıkı kaynak kapasiteli instance'lar için
    #    sınırlı backtracking. 20node/30node gibi örneklerde greedy kaynak
    #    paketleme hatasını çözer; büyük instance'ta kısa süre sonunda bırakır.
    if best_sol is None:
        bt_time = float(getattr(inst, "initial_backtracking_time_limit", 15.0 if len(inst.orders) <= 30 else 5.0))
        bt_nodes = int(getattr(inst, "initial_backtracking_node_limit", 300_000 if len(inst.orders) <= 30 else 50_000))
        _try(_backtracking_construct(inst, rng, time_limit=bt_time, node_limit=bt_nodes))

    if best_sol is None:
        raise RuntimeError(
            "Başlangıç feasible çözümü oluşturulamadı. Strict deadline/kaynak/yay kapasitesi hâlâ çakışıyor olabilir. "
            "Not: Bu durum inputun global infeasible olmasından veya çok sıkı kaynak/yay kapasitesi paketlemesinden kaynaklanabilir."
        )
    return best_sol

# ─────────────────────────────────────────────────────────────────────────────
# DESTROY OPERATÖRLERİ
# ─────────────────────────────────────────────────────────────────────────────

def destroy_random(inst, sol, q, rng) -> List[str]:
    """Rastgele q sipariş kaldır."""
    return rng.sample(list(sol.assignments), min(q, len(sol.assignments)))


def destroy_worst_cost(inst, sol, q, rng) -> List[str]:
    """
    En yüksek maliyetli q siparişi kaldır.
    FIX-v11: Tamamen deterministik sıralamayı kırmak için ilk 2q aday arasından
    rastgele q tanesini seç. Böylece aynı siparişler sürekli kaldırılmaz.
    """
    ranked = sorted(sol.assignments,
                    key=lambda k: sol.assignments[k].total_cost,
                    reverse=True)
    pool = ranked[:max(q, min(2 * q, len(ranked)))]
    return rng.sample(pool, min(q, len(pool)))


def destroy_cbam(inst, sol, q, rng) -> List[str]:
    """
    Pozitif CBAM yükümlülüğü en yüksek q siparişi kaldır (yeniden zamanlama için).
    FIX-v11: İlk 2q aday havuzundan rastgele örnekleme eklendi.
    """
    ranked = sorted(sol.assignments,
                    key=lambda k: (sol.assignments[k].cbam_cost,
                                   sol.assignments[k].total_cost),
                    reverse=True)
    pool = ranked[:max(q, min(2 * q, len(ranked)))]
    return rng.sample(pool, min(q, len(pool)))


def destroy_deadline(inst, sol, q, rng) -> List[str]:
    """
    Deadline'a en yakın q siparişi kaldır (varış/deadline oranı yüksek).
    FIX-v11: İlk 2q aday havuzundan rastgele örnekleme eklendi.
    """
    def slack_ratio(oid):
        a = sol.assignments[oid]
        return a.arrive_day / max(1, inst.orders[oid].deadline)
    ranked = sorted(sol.assignments, key=slack_ratio, reverse=True)
    pool = ranked[:max(q, min(2 * q, len(ranked)))]
    return rng.sample(pool, min(q, len(pool)))


def destroy_source_overload(inst, sol, q, rng) -> List[str]:
    """Kapasiteye en yakın kaynağın siparişlerinden q tanesini kaldır."""
    src_used: Dict[str, float] = defaultdict(float)
    for oid, a in sol.assignments.items():
        src_used[a.source] += inst.orders[oid].qty

    if not src_used:
        return destroy_random(inst, sol, q, rng)

    overloaded = max(src_used, key=lambda s: src_used[s] / inst.sources[s].capacity)
    pool       = [oid for oid, a in sol.assignments.items() if a.source == overloaded]
    rng.shuffle(pool)
    return pool[:q]


def destroy_related(inst, sol, q, rng) -> List[str]:
    """
    İlişkisel yıkım operatörü (Shaw, 1998).
    Rastgele bir tohum sipariş seçilir; normalize edilmiş ağırlıklı skalar
    ile benzer siparişler birlikte kaldırılır.

    FIX-v13-5: Önceki lexicographic 4-tuple sıralama → normalize+ağırlıklı skalar.
    Shaw (1998)'deki relatedness: R(i,j) = w1*loc + w2*time + w3*cap
    Burada:
      dest_diff  = hedef aynıysa 0, farklıysa 1              (ağırlık w_dest=3.0)
      src_diff   = kaynak aynıysa 0, farklıysa 1             (ağırlık w_src=2.0)
      mode_diff  = normalize edilmiş simetrik mod farkı      (ağırlık w_mode=1.5)
      dl_diff    = normalize edilmiş |deadline farkı|        (ağırlık w_dl=1.0)
    Skalar küçüldükçe sipariş çifti daha benzer → önce seçilir.
    """
    if not sol.assignments:
        return []

    seed_id    = rng.choice(list(sol.assignments))
    seed_a     = sol.assignments[seed_id]
    seed_o     = inst.orders[seed_id]
    seed_modes = {lg.mode_id for lg in seed_a.legs}

    # Normalizasyon referansları
    max_dl   = max(1, inst.max_H)                # deadline farkı normalize edeni
    n_modes  = max(1, len(inst.modes))           # mod farkı normalize edeni
    W_DEST, W_SRC, W_MODE, W_DL = 3.0, 2.0, 1.5, 1.0

    def relatedness(oid: str) -> float:
        a = sol.assignments[oid]
        o = inst.orders[oid]
        dest_d = 0.0 if o.dest   == seed_o.dest   else 1.0
        src_d  = 0.0 if a.source == seed_a.source  else 1.0
        mode_d = len(seed_modes.symmetric_difference(
                     {lg.mode_id for lg in a.legs})) / n_modes
        dl_d   = abs(o.deadline - seed_o.deadline) / max_dl
        return W_DEST * dest_d + W_SRC * src_d + W_MODE * mode_d + W_DL * dl_d

    return sorted(sol.assignments, key=relatedness)[:q]


DESTROY_OPS = {
    "random":          destroy_random,
    "worst_cost":      destroy_worst_cost,
    "cbam":            destroy_cbam,
    "deadline":        destroy_deadline,
    "source_overload": destroy_source_overload,
    "related":         destroy_related,
}


# ─────────────────────────────────────────────────────────────────────────────
# REPAIR OPERATÖRLERİ
# ─────────────────────────────────────────────────────────────────────────────

def _repair_core(inst: Instance, sol: Solution, removed: List[str],
                 order_fn, noise: float, rng: random.Random) -> Optional[Solution]:
    """
    Ortak repair iskeleti.
    order_fn(removed, inst) → hangi sırayla atama yapılacağını belirler.
    """
    sol      = copy.deepcopy(sol)
    src_used, arc_used = rebuild_capacity(inst, sol)
    rem      = order_fn(removed, inst)

    for oid in rem:
        order = inst.orders[oid]
        asgn  = best_insert(inst, order, src_used, arc_used, noise=noise, rng=rng)
        if asgn is None:
            return None
        sol.assignments[oid] = asgn
        apply_assignment(order, asgn, src_used, arc_used)

    return evaluate(inst, sol)


def repair_greedy(inst, sol, removed, rng) -> Optional[Solution]:
    """Deadline'a göre sıralı, gürültüsüz greedy repair."""
    return _repair_core(inst, sol, removed,
                        lambda rem, i: sorted(rem, key=lambda k: i.orders[k].deadline),
                        noise=0.0, rng=rng)


def repair_noisy(inst, sol, removed, rng) -> Optional[Solution]:
    """
    Maliyet fonksiyonuna gürültü ekleyerek çeşitlilik sağlayan repair.
    FIX-v11: noise 0.06 → 0.15; daha geniş komşuluk keşfi.
    Ek olarak insertion sırası %30 olasılıkla hafifçe karıştırılır.
    """
    def noisy_order(rem, i):
        seq = sorted(rem, key=lambda k: i.orders[k].deadline)
        if rng.random() < 0.30:
            # Küçük window shuffle ile insertion sırasına çeşitlilik kat
            window = max(2, len(seq) // 4)
            start  = rng.randint(0, max(0, len(seq) - window))
            block  = seq[start:start + window]
            rng.shuffle(block)
            seq[start:start + window] = block
        return seq
    return _repair_core(inst, sol, removed, noisy_order, noise=0.15, rng=rng)


def repair_regret2(inst, sol, removed, rng) -> Optional[Solution]:
    """
    Regret-2 repair: her adımda 'en çok kaybedecek' siparişi önce ata.
    En iyi iki seçenek arasındaki maliyet farkı regret skorunu belirler.

    HIZLANDIRMA: Kapasite değişmemiş siparişler için aday listesi cache'den okunur.
    Bir sipariş atandıktan sonra yalnızca etkilenen yayları kullanan diğer siparişlerin
    cache'i temizlenir — bu O(|removed|²×|sources|×Dijkstra) karmaşıklığını pratikte
    çok daha düşük bir sabite indirger.
    """
    sol      = copy.deepcopy(sol)
    src_used, arc_used = rebuild_capacity(inst, sol)
    rem      = list(removed)

    # Cache: oid → (best_asgn, second_best_cost, regret)
    # None → yeniden hesaplanmalı
    cache: Dict[str, Optional[Tuple]] = {oid: None for oid in rem}

    def compute_entry(oid: str):
        """(cands[0], regret) veya None döndürür."""
        order = inst.orders[oid]
        cands = []
        for sid in inst.source_ids:
            a = find_best_assignment(inst, order, sid, src_used, arc_used, 0.0, rng)
            if a:
                cands.append(a)
        if not cands:
            return None
        cands.sort(key=lambda x: x.total_cost)
        regret = (cands[1].total_cost - cands[0].total_cost) if len(cands) > 1 else 1e9
        return (cands[0], regret)

    while rem:
        # Cache'i doldurmayan girişleri hesapla
        for oid in rem:
            if cache[oid] is None:
                cache[oid] = compute_entry(oid)

        best_regret = -INF
        best_asgn   = None
        best_oid    = None

        for oid in rem:
            entry = cache[oid]
            if entry is None:
                continue
            asgn, regret = entry
            if regret > best_regret:
                best_regret = regret
                best_asgn   = asgn
                best_oid    = oid

        if best_oid is None or best_asgn is None:
            return None

        order = inst.orders[best_oid]
        # v9 safety: cache'ten gelen assignment uygulanmadan önce mevcut
        # kapasite durumuna göre tekrar doğrulanır. Geçersizse yeniden hesaplatılır.
        if not assignment_capacity_feasible(inst, order, best_asgn, src_used, arc_used):
            cache[best_oid] = None
            continue
        sol.assignments[best_oid] = best_asgn
        apply_assignment(order, best_asgn, src_used, arc_used)
        rem.remove(best_oid)
        del cache[best_oid]

        # Yalnızca aynı kaynağı veya aynı (yay, gün) kombinasyonunu kullanan
        # siparişlerin cache'ini temizle — diğerleri hâlâ geçerli.
        used_arcs = {(lg.frm, lg.to, lg.mode_id, lg.depart_day) for lg in best_asgn.legs}
        for oid in rem:
            if cache[oid] is None:
                continue
            prev_asgn, _ = cache[oid]
            if prev_asgn.source == best_asgn.source:
                cache[oid] = None  # kaynak kapasitesi değişti
                continue
            for lg in prev_asgn.legs:
                if (lg.frm, lg.to, lg.mode_id, lg.depart_day) in used_arcs:
                    cache[oid] = None  # yay kapasitesi değişti
                    break

    return evaluate(inst, sol)


def repair_cbam_priority(inst, sol, removed, rng) -> Optional[Solution]:
    """
    CBAM yükümlülüğü yüksek siparişleri önce ata.
    Bu siparişlerin daha erken atanması daha iyi CBAM haftası seçimini kolaylaştırır.
    """
    # FIX: closure yerine order_fn'in i parametresini kullanan lambda
    return _repair_core(
        inst, sol, removed,
        lambda rem, i: sorted(
            rem,
            key=lambda k: i.orders[k].qty * max(0.0, i.orders[k].emb_em - i.orders[k].bench_em),
            reverse=True,
        ),
        noise=0.0, rng=rng,
    )


REPAIR_OPS = {
    "greedy":        repair_greedy,
    "noisy":         repair_noisy,
    "regret2":       repair_regret2,
    "cbam_priority": repair_cbam_priority,
}


# ─────────────────────────────────────────────────────────────────────────────
# ADAPTİF AĞIRLIK YÖNETİCİSİ
# ─────────────────────────────────────────────────────────────────────────────

class AdaptiveWeights:
    """
    Roulette-wheel seçim + üstel hareketli ortalama ağırlık güncelleme.
    Segment (varsayılan 100 iter) sonunda flush() ile ağırlıklar güncellenir.
    """

    def __init__(self, names: List[str], decay: float = 0.8, min_w: float = 0.05):
        self.names   = names
        self.decay   = decay
        self.min_w   = min_w
        self.weights = {n: 1.0 for n in names}
        self._scores = {n: 0.0 for n in names}
        self._counts = {n: 0   for n in names}

    def select(self, rng: random.Random) -> str:

        weights_list = list(self.weights.values())
        cumulative   = list(itertools.accumulate(weights_list))
        total        = cumulative[-1]
        r            = rng.random() * total
        idx          = bisect.bisect_left(cumulative, r)
        idx          = min(idx, len(self.names) - 1)  # sınır koruması
        return self.names[idx]

    def record(self, name: str, reward: float):
        self._scores[name] += reward
        self._counts[name] += 1

    def flush(self):
        """Segment sonunda ağırlıkları güncelle (üstel hareketli ortalama)."""
        for n in self.names:
            if self._counts[n] > 0:
                avg = self._scores[n] / self._counts[n]
                self.weights[n] = max(
                    self.min_w,
                    self.decay * self.weights[n] + (1 - self.decay) * avg,
                )
        self._scores = {n: 0.0 for n in self.names}
        self._counts = {n: 0   for n in self.names}


# ─────────────────────────────────────────────────────────────────────────────
# ALNS ANA DÖNGÜSÜ
# ─────────────────────────────────────────────────────────────────────────────

# Ödül seviyeleri — Algorithm 1 (rapor)
R_NEW_BEST = 8.0
R_IMPROVE  = 4.0
R_ACCEPT   = 1.0
R_REJECT   = 0.2


def alns(
    inst:       Instance,
    max_iter:   int   = 2000,
    time_limit: float = 300.0,
    alpha:      float = 0.9975,  # SA soğuma faktörü (yavaş soğuma)
    segment:    int   = 100,     # ağırlık güncelleme sıklığı
    decay:      float = 0.8,     # üstel ortalama bozunma oranı
    seed:       Optional[int] = None,  # FIX-v11: None → her run'da farklı seed
    verbose:    bool  = True,
) -> Tuple[Solution, List[dict]]:
    """
    Adaptive Large Neighbourhood Search ana döngüsü.

    Dönüş değerleri:
        best (Solution): bulunan en iyi çözüm
        log  (List[dict]): iterasyon bazında convergence kaydı

    seed=None (varsayılan) ise her çalıştırmada farklı bir rastgele tohum
    kullanılır; bu sayede ALNS her run'da farklı bir arama yolunu izler.
    Tekrarlanabilir sonuç için açıkça bir tam sayı veriniz (örn. seed=42).
    """
    # FIX-v13-2: import os dosya başına taşındı; buradaki yerel import kaldırıldı.
    if seed is None:
        seed = int.from_bytes(os.urandom(4), "big")
    if verbose:
        print(f"  Rastgele tohum (seed): {seed}")
    rng = random.Random(seed)

    # ── 1. Başlangıç çözümü ──────────────────────────────────────────────────
    if verbose:
        print("\nBaşlangıç çözümü kuruluyor…")
    current = initial_solution(inst, rng)
    best    = copy.deepcopy(current)

    if verbose:
        print(f"  Başlangıç maliyeti: {current.obj:,.2f} €  "
              f"({len(current.assignments)}/{len(inst.orders)} sipariş)")

    # ── 2. SA parametreleri ───────────────────────────────────────────────────
    # FIX-v13-6: Instance-aware T0 (delta örnekleme).
    # Sabit katsayı (0.08) yerine: başlangıçta 50 rastgele hamle yapılır,
    # gözlemlenen delta değerlerinin medyanından exp(-delta_median/T0)=0.5
    # koşulu sağlanır → T0 = -delta_median / ln(0.5) = delta_median / ln(2).
    # Bu yaklaşım SA literatüründe standart "calibration by sampling" yöntemidir
    # (Ropke & Pisinger 2006; Kirkpatrick et al. 1983).
    # Küçük instance'ta T0 düşük, büyük/zor instance'ta yüksek otomatik ayarlanır.
    _N_CALIB    = 50
    _deltas: List[float] = []
    _calib_rng  = random.Random(rng.randint(0, 2**31))
    _n_calib_orders = len(inst.orders)
    _calib_src_used, _calib_arc_used = rebuild_capacity(inst, current)
    for _ in range(_N_CALIB):
        _q_c = max(1, min(8, round(_n_calib_orders * _calib_rng.uniform(0.03, 0.09))))
        _rem  = destroy_random(inst, current, _q_c, _calib_rng)
        _part = Solution(
            assignments={k: v for k, v in current.assignments.items() if k not in _rem},
            obj=INF, feasible=False,
        )
        _cand = repair_greedy(inst, _part, _rem, _calib_rng)
        if _cand and _cand.feasible and _cand.obj > current.obj:
            _deltas.append(_cand.obj - current.obj)
    if _deltas:
        _delta_median = sorted(_deltas)[len(_deltas) // 2]
        T0 = max(1.0, _delta_median / math.log(2.0))
    else:
        # Fallback: hiç kötü çözüm bulunamadıysa başlangıç maliyetinin küçük bir oranı
        T0 = max(1.0, 0.05 * current.obj)
    T  = T0
    if verbose:
        print(f"  SA T0 (delta-sampling, {len(_deltas)} örnek): {T0:,.2f} €")

    # ── 3. Adaptif ağırlıklar ─────────────────────────────────────────────────
    d_weights = AdaptiveWeights(list(DESTROY_OPS), decay)
    r_weights = AdaptiveWeights(list(REPAIR_OPS),  decay)
    # v10: Önceki loglarda regret2 en fazla NEW_BEST üreten repair idi.
    # Başlangıç ağırlığını artırmak aramayı kalite yönünde hızlandırır; adaptif
    # mekanizma başarısızsa ağırlığı yine düşürür.
    if "regret2" in r_weights.weights:
        r_weights.weights["regret2"] = 2.5
    if "greedy" in r_weights.weights:
        r_weights.weights["greedy"] = 1.2

    log:     List[dict] = []
    t_start: float      = time.time()
    n_orders: int       = len(inst.orders)
    _last_best_check    = best.obj  # reheating için iyileşme takibi

    # ── 4. Ana döngü ──────────────────────────────────────────────────────────
    for it in range(1, max_iter + 1):
        if time.time() - t_start > time_limit:
            break
        t_iter = time.time()

        # FIX-v12: q parametresi dengelendi.
        # v11'de max(8, min(30)) ile q ortalama ~17'ye çıkmıştı; regret2 repair bu
        # durumda iterasyon başına süreyi 5× artırdı (410ms → 2077ms) ve 300s'de
        # sadece 88 iterasyon yapılabildi. v10'da 828 iterasyon yapılmıştı.
        # v12: büyük instance'ta %3-9 → 4-12 arası; v10 hızını korur, v11
        # çeşitliliğini de (2-8 aralığına geri dönmüyor).
        if n_orders <= 30:
            q = max(1, min(10, round(n_orders * rng.uniform(0.10, 0.30))))
        else:
            q = max(4, min(12, round(n_orders * rng.uniform(0.03, 0.09))))

        # Operatör seçimi
        d_name = d_weights.select(rng)
        r_name = r_weights.select(rng)

        # Destroy
        removed = DESTROY_OPS[d_name](inst, current, q, rng)
        # FIX: partial için deepcopy gereksiz; repair zaten deepcopy yapıyor.
        # dict.pop ile kaldırılan atamaları geçici bir kopya üzerinde işle.
        partial = Solution(
            assignments={k: v for k, v in current.assignments.items() if k not in removed},
            obj=INF,
            feasible=False,
        )

        # Repair
        candidate = REPAIR_OPS[r_name](inst, partial, removed, rng)

        # Değerlendirme ve kabul kararı
        reward   = R_REJECT
        accepted = False
        cand_obj = INF

        if candidate and candidate.feasible:
            cand_obj = candidate.obj
            delta    = cand_obj - current.obj

            if cand_obj < best.obj - EPS:
                best     = copy.deepcopy(candidate)
                current  = candidate
                accepted = True
                reward   = R_NEW_BEST
            elif cand_obj < current.obj - EPS:
                current  = candidate
                accepted = True
                reward   = R_IMPROVE
            elif rng.random() < math.exp(-delta / max(T, EPS)):
                current  = candidate
                accepted = True
                reward   = R_ACCEPT

        # Ağırlık güncelleme
        d_weights.record(d_name, reward)
        r_weights.record(r_name, reward)

        # SA soğuma
        T *= alpha

        # SA Reheating: FIX-v12: T0*0.20 → T0*0.10.
        # T0 artık 0.08*obj (daha küçük); 0.20 ile reheating çok agresif olurdu.
        # 0.10 ile 100 iterasyonda iyileşme yoksa sıcaklık T0'ın %10'una döner;
        # yerel optimumdan kaçmak için yeterli ama intensification'ı bozmaz.
        if it % 100 == 0:
            if best.obj >= _last_best_check - 1e-6:
                T = max(T, T0 * 0.10)
            _last_best_check = best.obj

        # Segment sonu ağırlık güncelleme
        if it % segment == 0:
            d_weights.flush()
            r_weights.flush()

        # FIX: elapsed_s döngü bitiş zamanını, iter_ms bu iterasyonun süresini gösterir
        now = time.time()
        elapsed_now = now - t_start
        iter_ms_now = round((now - t_iter) * 1000, 1)

        # Karar tipi (log okunabilirliği için)
        if reward == R_NEW_BEST:
            decision = "NEW_BEST"
        elif reward == R_IMPROVE:
            decision = "IMPROVE"
        elif accepted:
            decision = "SA_ACCEPT"
        else:
            decision = "REJECT"

        # Improvement % relative to initial
        if log and log[0]["best"] > 0:
            pct_improv = 100.0 * (log[0]["best"] - best.obj) / log[0]["best"]
        else:
            pct_improv = 0.0

        if it == 1 or it % 25 == 0 or reward in (R_NEW_BEST, R_IMPROVE):
            log.append({
                "iter":           it,
                "best":           round(best.obj, 2),
                "current":        round(current.obj, 2),
                "candidate":      round(cand_obj, 2) if cand_obj < INF else None,
                "decision":       decision,
                "accepted":       accepted,
                "destroy_op":     d_name,
                "repair_op":      r_name,
                "q_removed":      q,
                "temperature":    round(T, 4),
                "pct_improvement":round(pct_improv, 4),
                "elapsed_s":      round(elapsed_now, 3),
                "iter_ms":        iter_ms_now,
            })

        if verbose and (it % 200 == 0 or reward == R_NEW_BEST):
            tag = " ◄ YENİ EN İYİ" if reward == R_NEW_BEST else ""
            pct_str = f"{pct_improv:.2f}%" if pct_improv > 0 else "—"
            print(f"  iter {it:5d} | best {best.obj:>12,.2f} € | cur {current.obj:>12,.2f} € | "
                  f"T={T:7.3f} | {elapsed_now:6.1f}s | d={d_name:<15} r={r_name:<14} | "
                  f"iyileşme={pct_str}{tag}")

    return best, log


# ─────────────────────────────────────────────────────────────────────────────
# ÇIKTI ÜRETME
# ─────────────────────────────────────────────────────────────────────────────

def solution_to_frames(inst: Instance, sol: Solution
                       ) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Sipariş ve leg seviyesinde DataFrame döndür."""
    order_rows = []
    leg_rows   = []

    for oid in sorted(sol.assignments):
        asgn  = sol.assignments[oid]
        order = inst.orders[oid]
        path  = [asgn.source] + [lg.to for lg in asgn.legs]
        modes = " → ".join(dict.fromkeys(lg.mode_id for lg in asgn.legs))
        slack = order.deadline - asgn.arrive_day

        # Emisyon ve elleçleme maliyetlerini ayrıştır
        transport_cost = 0.0
        emission_cost  = 0.0
        handling_cost_val = 0.0
        for leg in asgn.legs:
            mo = inst.modes[leg.mode_id]
            arc_obj = inst.arc_index.get((leg.frm, leg.to, leg.mode_id))  # FIX-v13-4: O(1)
            if arc_obj:
                tr_cost  = mo.cost_per_tkm * arc_obj.dist_km * order.qty
                em_cost  = inst.lambda_kg * mo.emission_kg_per_tkm * arc_obj.dist_km * order.qty
                hdl_cost = inst.nodes[leg.to].handling_cost * order.qty
            else:
                tr_cost  = em_cost = hdl_cost = 0.0
            transport_cost  += tr_cost
            emission_cost   += em_cost
            handling_cost_val += hdl_cost
        # kaynak yükleme handling
        handling_cost_val += inst.nodes[asgn.source].handling_cost * order.qty

        order_rows.append({
            "order_id":       oid,
            "tons":           order.qty,
            "product":        order.product,
            "source":         asgn.source,
            "destination":    order.dest,
            "depart_day":     asgn.depart_day,
            "arrive_day":     asgn.arrive_day,
            "deadline_day":   order.deadline,
            "slack_days":     slack,
            "on_time":        "✔" if slack >= 0 else "✘",
            "arrival_date":   (inst.plan_start + timedelta(days=asgn.arrive_day)).strftime("%d.%m.%Y"),
            "deadline_date":  (inst.plan_start + timedelta(days=order.deadline)).strftime("%d.%m.%Y"),
            "cbam_week":      asgn.cbam_week,
            "cbam_subject":   "CBAM" if order.cbam_subject else "muaf",
            "path":           " → ".join(path),
            "modes":          modes,
            "transport_cost": round(transport_cost,    2),
            "emission_cost":  round(emission_cost,     2),
            "handling_cost":  round(handling_cost_val, 2),
            "move_cost":      round(asgn.move_cost,    2),
            "cbam_cost":      round(asgn.cbam_cost,    2),
            "total_cost":     round(asgn.total_cost,   2),
        })

        for k, leg in enumerate(asgn.legs, 1):
            mo = inst.modes[leg.mode_id]
            arc_obj = inst.arc_index.get((leg.frm, leg.to, leg.mode_id))  # FIX-v13-4: O(1)
            dist = arc_obj.dist_km if arc_obj else leg.dist_km
            tr_c  = mo.cost_per_tkm * dist * order.qty
            em_c  = inst.lambda_kg * mo.emission_kg_per_tkm * dist * order.qty
            hdl_c = inst.nodes[leg.to].handling_cost * order.qty
            leg_rows.append({
                "order_id":    oid,
                "leg_seq":     k,
                "from_node":   leg.frm,
                "to_node":     leg.to,
                "mode":        leg.mode_id,
                "depart_day":  leg.depart_day,
                "arrive_day":  leg.arrive_day,
                "depart_date": (inst.plan_start + timedelta(days=leg.depart_day)).strftime("%d.%m.%Y"),
                "arrive_date": (inst.plan_start + timedelta(days=leg.arrive_day)).strftime("%d.%m.%Y"),
                "distance_km":    round(dist,    2),
                "tons":           order.qty,
                "transport_cost": round(tr_c,    2),
                "emission_cost":  round(em_c,    2),
                "handling_cost":  round(hdl_c,   2),
                "leg_cost":       round(tr_c + em_c + hdl_c, 2),
            })

    return pd.DataFrame(order_rows), pd.DataFrame(leg_rows)


def print_report(inst: Instance, sol: Solution, elapsed: float, n_iter: int):
    """Konsola biçimlendirilmiş çözüm raporu yaz."""
    orders_df, _ = solution_to_frames(inst, sol)
    sep = "─" * 120

    print(f"\n{'═'*120}")
    print(f"  ALNS v{VERSION} ÇÖZÜM RAPORU — CBAM Çok Modlu Taşımacılık")
    print(f"{'═'*120}")

    cols = ["order_id", "tons", "product", "source", "destination", "modes",
            "depart_day", "arrive_day", "deadline_day", "slack_days", "on_time",
            "cbam_week", "cbam_subject", "transport_cost", "emission_cost",
            "handling_cost", "cbam_cost", "total_cost"]
    print("\n" + orders_df[cols].to_string(index=False))
    print(sep)

    total_tr   = orders_df["transport_cost"].sum()
    total_em   = orders_df["emission_cost"].sum()
    total_hdl  = orders_df["handling_cost"].sum()
    total_cbam = orders_df["cbam_cost"].sum()
    total_all  = orders_df["total_cost"].sum()
    n_on_time  = (orders_df["on_time"] == "✔").sum()
    n_orders   = len(orders_df)

    print(f"\n  Taşıma maliyeti    : {total_tr:>13,.2f} €")
    print(f"  Emisyon maliyeti   : {total_em:>13,.2f} €")
    print(f"  Elleçleme maliyeti : {total_hdl:>13,.2f} €")
    print(f"  CBAM yükümlülüğü   : {total_cbam:>13,.2f} €")
    print(f"  {'─'*38}")
    print(f"  BASE OBJECTIVE     : {sol.base_obj:>13,.2f} €")
    print(f"  ARC OVERFLOW PEN.  : {sol.arc_overflow_penalty:>13,.2f} €")
    print(f"  TOTAL OBJECTIVE    : {sol.obj:>13,.2f} €")
    print(f"  Strict feasible    : {'YES' if sol.strict_feasible else 'NO'}")
    print(f"  Overflow arc count : {sol.overflow_arc_count:>13}")
    print(f"  Total overflow ton : {sol.total_overflow_ton:>13,.2f}")
    print(f"  Deadline violations: {sol.deadline_violation_count:>13}")
    print(f"  Max lateness days  : {sol.max_lateness_days:>13,.2f}")
    print(f"  Ort. maliyet/ton   : {total_all/max(1,orders_df['tons'].sum()):>13,.2f} €/ton")
    print(f"  Zamanında teslim   : {n_on_time}/{n_orders}  ({100*n_on_time/max(1,n_orders):.1f}%)")

    src_used: Dict[str, float] = defaultdict(float)
    for oid, a in sol.assignments.items():
        src_used[a.source] += inst.orders[oid].qty

    print(f"\n  KAYNAK KULLANIMI:")
    for sid in sorted(inst.sources):
        cap  = inst.sources[sid].capacity
        used = src_used.get(sid, 0.0)
        pct  = 100 * used / cap
        bar  = "█" * int(pct / 5)
        print(f"    {sid:<8}  {used:>7.0f}/{cap:>7.0f} ton  ({pct:5.1f}%)  {bar}")

    unassigned = [oid for oid in inst.orders if oid not in sol.assignments]
    if unassigned:
        print(f"\n  ⚠ ATANAMAYAN SİPARİŞLER: {', '.join(unassigned)}")
    else:
        print(f"\n  ✓ Tüm {len(inst.orders)} sipariş başarıyla atandı.")

    print(f"\n  Çalışma süresi : {elapsed:.2f}s  |  Log satırı : {n_iter}")
    print(f"{'═'*120}\n")


# ─────────────────────────────────────────────────────────────────────────────
# XLSX RAPOR YAZICI
# ─────────────────────────────────────────────────────────────────────────────

# Renk paleti (template ile uyumlu)
_C_DARK_NAVY  = "1B2A4A"
_C_NAVY       = "2E5B8A"
_C_MID_NAVY   = "22344F"
_C_LIGHT_GREY = "F5F5F5"
_C_HEADER_BG  = "2E5B8A"
_C_WHITE      = "FFFFFF"
_C_LABEL      = "999999"
_C_TEXT_DK    = "1B2A4A"
_C_GREEN      = "2E7D32"
_C_WARN       = "C62828"
_C_AMBER      = "E65100"
_C_SUBHDR     = "F8F8F8"
_C_BLUE_VAL   = "2E5B8A"
_C_SLATE      = "37474F"
_C_ALT_ROW    = "EEF4FB"


def _font(name="Calibri", size=10, bold=False, color="000000", italic=False):
    return Font(name=name, size=size, bold=bold,
                color=color, italic=italic)


def _fill(color):
    return PatternFill("solid", fgColor=color)


def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _border_thin(sides="bottom"):
    thin = Side(style="thin", color="CCCCCC")
    none = Side(style=None)
    b = Border(
        left   = thin if "left"   in sides else none,
        right  = thin if "right"  in sides else none,
        top    = thin if "top"    in sides else none,
        bottom = thin if "bottom" in sides else none,
    )
    return b


def _border_all():
    thin = Side(style="thin", color="DDDDDD")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _write_header_row(ws, row, col_start, headers, bg=_C_HEADER_BG):
    """Başlık satırı yaz; tüm hücreleri renklendir."""
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=col_start + i, value=h)
        c.font      = _font(bold=True, color=_C_WHITE, size=10)
        c.fill      = _fill(bg)
        c.alignment = _align("center")
        c.border    = _border_all()


def _fmt_eur(val):
    """€ prefix ile formatlı string."""
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return "—"
    return f"€ {val:,.2f}"


def _write_kpi_box(ws, row, col, label, value, val_color=_C_TEXT_DK):
    """Küçük KPI kutusu: label üstte, değer altında."""
    lc = ws.cell(row=row, column=col, value=label)
    lc.font      = _font(size=9, color=_C_LABEL)
    lc.fill      = _fill(_C_SUBHDR)
    lc.alignment = _align("center")

    vc = ws.cell(row=row + 1, column=col, value=value)
    vc.font      = _font(size=15, bold=True, color=val_color)
    vc.fill      = _fill(_C_WHITE)
    vc.alignment = _align("center")


def write_xlsx_report(
    inst:       Instance,
    sol:        Solution,
    elapsed:    float,
    out_path:   str | Path,
    solver_tag: str = f"ALNS v{VERSION}",
):
    """
    Rapor XLSX dosyasına 3 sekme yazar:
      1. Dashboard    — KPI ozeti + siparis maliyet tablosu
      2. Order Detail — tum siparis detaylari
      3. Route Legs   — leg seviyesi guzergah tablosu
    Log ayri dosyaya (write_xlsx_log) yazilir.
    """
    orders_df, legs_df = solution_to_frames(inst, sol)
    wb = Workbook()

    # ── Renk yardımcıları ────────────────────────────────────────────────────
    def stripe(row_idx):
        return _fill(_C_ALT_ROW) if row_idx % 2 == 0 else _fill(_C_WHITE)

    # ─────────────────────────────────────────────────────────────────────────
    # SEKMELERİ OLUŞTUR
    # ─────────────────────────────────────────────────────────────────────────
    ws_dash  = wb.active
    ws_dash.title = "Dashboard"
    ws_ord   = wb.create_sheet("Order Detail")
    ws_legs  = wb.create_sheet("Route Legs")

    # =========================================================================
    # 1. DASHBOARD
    # =========================================================================
    ws = ws_dash
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    for col_letter, w in zip("BCDEFGHIJ", [24, 18, 18, 18, 18, 18, 18, 18, 18]):
        ws.column_dimensions[col_letter].width = w

    # — Başlık bandı —
    ws.row_dimensions[1].height = 8
    ws.row_dimensions[2].height = 36
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 8

    now_str  = datetime.now().strftime("%d %b %Y  %H:%M")
    title_c  = ws["B2"]
    title_c.value     = "CBAM Multimodal Transportation  ·  Optimisation Results"
    title_c.font      = _font(size=22, bold=True, color=_C_WHITE)
    title_c.fill      = _fill(_C_DARK_NAVY)
    title_c.alignment = _align("left", "center")
    ws.merge_cells("B2:J2")

    sub_c = ws["B3"]
    sub_c.value     = (f"Çözüm: {now_str}   |   Çözücü: {solver_tag}   |   "
                       f"Çalışma süresi: {elapsed:.1f}s   |   "
                       f"Siparişler: {len(inst.orders)}   |   "
                       f"CBAM φ_r: {inst.phi_r}  ({inst.plan_year})")
    sub_c.font      = _font(size=10, color="AAAAAA")
    sub_c.fill      = _fill(_C_MID_NAVY)
    sub_c.alignment = _align("left", "center")
    ws.merge_cells("B3:J3")

    # — KPI kutuları (satır 6-7) —
    ws.row_dimensions[5].height = 10
    ws.row_dimensions[6].height = 22
    ws.row_dimensions[7].height = 30
    ws.row_dimensions[8].height = 12

    total_tr   = orders_df["transport_cost"].sum()
    total_em   = orders_df["emission_cost"].sum()
    total_hdl  = orders_df["handling_cost"].sum()
    total_cbam = orders_df["cbam_cost"].sum()
    total_all  = orders_df["total_cost"].sum()
    total_tons = orders_df["tons"].sum()
    n_on_time  = (orders_df["on_time"] == "✔").sum()
    n_orders_  = len(orders_df)
    avg_per_t  = total_all / max(1, total_tons)

    kpi_cols = [
        ("B", "Grand Total Cost",    _fmt_eur(total_all),  _C_TEXT_DK),
        ("D", "CBAM Cost",           _fmt_eur(total_cbam), _C_BLUE_VAL),
        ("F", "Transport + Handling",_fmt_eur(total_tr + total_hdl), _C_SLATE),
        ("H", "Avg € / Ton",         _fmt_eur(avg_per_t),  _C_SLATE),
        ("J", "On-Time Rate",
         f"{n_on_time}/{n_orders_}  {100*n_on_time/max(1,n_orders_):.0f}%",
         _C_GREEN),
    ]
    for col_ltr, lbl, val, col_v in kpi_cols:
        col_idx = ord(col_ltr) - ord("A") + 1
        _write_kpi_box(ws, 6, col_idx, lbl, val, col_v)
        # birleştir 2 sütun
        nxt = get_column_letter(col_idx + 1)
        ws.merge_cells(f"{col_ltr}6:{nxt}6")
        ws.merge_cells(f"{col_ltr}7:{nxt}7")

    # — Maliyet kırılımı başlık —
    ws.row_dimensions[9].height  = 22
    ws.row_dimensions[10].height = 22

    sec_c = ws["B9"]
    sec_c.value     = "Sipariş Bazlı Maliyet Özeti"
    sec_c.font      = _font(size=13, bold=True, color=_C_TEXT_DK)
    sec_c.alignment = _align("left", "center")

    dash_hdrs = ["Sipariş", "Varış", "Ton", "Kaynak",
                 "Taşıma (€)", "Emisyon (€)", "Elleçleme (€)",
                 "CBAM (€)", "Toplam (€)", "Mod"]
    _write_header_row(ws, 10, 2, dash_hdrs)

    for i, (_, r) in enumerate(orders_df.iterrows()):
        row_idx = 11 + i
        ws.row_dimensions[row_idx].height = 16
        vals = [
            r["order_id"], r["destination"], r["tons"], r["source"],
            r["transport_cost"], r["emission_cost"], r["handling_cost"],
            r["cbam_cost"], r["total_cost"], r["modes"],
        ]
        fill = stripe(i)
        for j, v in enumerate(vals):
            c = ws.cell(row=row_idx, column=2 + j, value=v)
            c.font      = _font(size=10)
            c.fill      = fill
            c.alignment = _align("center" if j in (0,1,3,9) else "right")
            c.border    = _border_thin("bottom")
            if j in (4, 5, 6, 7, 8) and isinstance(v, (int, float)):
                c.number_format = '#,##0.00'

    # Toplam satırı
    total_row = 11 + len(orders_df)
    ws.row_dimensions[total_row].height = 18
    total_vals = ["", "TOPLAM", total_tons, "",
                  total_tr, total_em, total_hdl,
                  total_cbam, total_all, ""]
    for j, v in enumerate(total_vals):
        c = ws.cell(row=total_row, column=2 + j, value=v)
        c.font      = _font(size=10, bold=True, color=_C_WHITE)
        c.fill      = _fill(_C_DARK_NAVY)
        c.alignment = _align("center" if j in (0,1,3,9) else "right")
        if j in (4, 5, 6, 7, 8) and isinstance(v, (int, float)):
            c.number_format = '#,##0.00'

    # =========================================================================
    # 2. ORDER DETAIL
    # =========================================================================
    ws = ws_ord
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2

    col_widths_ord = {
        "B": 10, "C": 10, "D": 8,  "E": 12, "F": 12, "G": 12,
        "H": 12, "I": 12, "J": 12, "K": 6,  "L": 7,  "M": 13,
        "N": 13, "O": 11, "P": 10, "Q": 30, "R": 20,
        "S": 13, "T": 13, "U": 13, "V": 13, "W": 13,
    }
    for col_ltr, w in col_widths_ord.items():
        ws.column_dimensions[col_ltr].width = w

    # Başlık bandı
    ws["B1"].value     = "Order Detail — Sipariş Detayları"
    ws["B1"].font      = _font(size=14, bold=True, color=_C_WHITE)
    ws["B1"].fill      = _fill(_C_DARK_NAVY)
    ws["B1"].alignment = _align("left", "center")
    ws.merge_cells("B1:W1")
    ws.row_dimensions[1].height = 28

    ord_hdrs = [
        "Sipariş ID", "Varış", "Ton", "Ürün", "Kaynak",
        "Kalkış Günü", "Varış Günü", "Son Tarih (Gün)", "Esneklik (Gün)", "Zamanında",
        "Kalkış Tarihi", "Varış Tarihi", "Son Teslim Tarihi",
        "CBAM Haftası", "CBAM Durumu",
        "Güzergâh (Düğümler)", "Güzergâh (Modlar)",
        "Taşıma (€)", "Emisyon (€)", "Elleçleme (€)", "CBAM (€)", "Toplam (€)",
    ]
    _write_header_row(ws, 2, 2, ord_hdrs)
    ws.row_dimensions[2].height = 18

    col_map_ord = [
        "order_id", "destination", "tons", "product", "source",
        "depart_day", "arrive_day", "deadline_day", "slack_days", "on_time",
        "depart_day", "arrival_date", "deadline_date",
        "cbam_week", "cbam_subject",
        "path", "modes",
        "transport_cost", "emission_cost", "handling_cost", "cbam_cost", "total_cost",
    ]
    # depart_day yerine depart_date (computed above from arrival_date already)
    # Override col 11 (kalkış tarihi) → compute from plan_start + depart_day
    for i, (_, r) in enumerate(orders_df.iterrows()):
        row_idx = 3 + i
        ws.row_dimensions[row_idx].height = 15
        fill = stripe(i)
        for j, col_key in enumerate(col_map_ord):
            col_num = 2 + j
            if col_key == "depart_day" and j == 10:
                # Kalkış tarihi
                val = (inst.plan_start + timedelta(days=int(r["depart_day"]))).strftime("%d.%m.%Y")
            else:
                val = r[col_key]

            c = ws.cell(row=row_idx, column=col_num, value=val)
            c.font   = _font(size=10)
            c.fill   = fill
            c.border = _border_thin("bottom")

            # Hizalama
            if j in (0,1,3,4,9,13,14,16):
                c.alignment = _align("center")
            elif j in (15,):
                c.alignment = _align("left", wrap=True)
            else:
                c.alignment = _align("right")

            # Sayı formatı
            if j in (17,18,19,20,21) and isinstance(val, (int, float)):
                c.number_format = '#,##0.00'

            # On-time renk
            if j == 9:
                c.font = _font(size=10, bold=True,
                               color=_C_GREEN if val == "✔" else _C_WARN)

            # Slack renk: 0 = amber, <0 = red
            if j == 8 and isinstance(val, (int, float)):
                if val < 0:
                    c.font = _font(size=10, color=_C_WARN)
                elif val == 0:
                    c.font = _font(size=10, color=_C_AMBER)

            # CBAM cost renk
            if j == 20 and isinstance(val, (int, float)) and val > 0:
                c.font = _font(size=10, color=_C_BLUE_VAL)

    # Toplam satırı
    total_row_ord = 3 + len(orders_df)
    ws.row_dimensions[total_row_ord].height = 18
    totals_ord = {17: total_tr, 18: total_em, 19: total_hdl,
                  20: total_cbam, 21: total_all}
    c_lbl = ws.cell(row=total_row_ord, column=2, value="GENEL TOPLAM")
    c_lbl.font = _font(bold=True, color=_C_WHITE, size=10)
    c_lbl.fill = _fill(_C_DARK_NAVY)
    c_lbl.alignment = _align("center")
    ws.merge_cells(f"B{total_row_ord}:Q{total_row_ord}")
    for j in range(17, 22):
        c = ws.cell(row=total_row_ord, column=2 + j, value=totals_ord[j])
        c.font = _font(bold=True, color=_C_WHITE, size=10)
        c.fill = _fill(_C_DARK_NAVY)
        c.alignment = _align("right")
        c.number_format = '#,##0.00'

    ws.freeze_panes = "B3"

    # =========================================================================
    # 3. ROUTE LEGS
    # =========================================================================
    ws = ws_legs
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2

    for col_ltr, w in zip("BCDEFGHIJKLMNOPQ",
                           [10,6,8,8,8,12,12,12,12,8,8,13,13,13,13,13]):
        ws.column_dimensions[col_ltr].width = w

    ws["B1"].value     = "Route Legs — Güzergâh Bacakları"
    ws["B1"].font      = _font(size=14, bold=True, color=_C_WHITE)
    ws["B1"].fill      = _fill(_C_DARK_NAVY)
    ws["B1"].alignment = _align("left", "center")
    ws.merge_cells("B1:Q1")
    ws.row_dimensions[1].height = 28

    leg_hdrs = ["Sipariş", "Bacak #", "Çıkış", "Varış", "Mod",
                "Kalkış Günü", "Varış Günü", "Kalkış Tarihi", "Varış Tarihi",
                "Mesafe (km)", "Ton",
                "Taşıma (€)", "Emisyon (€)", "Elleçleme (€)", "Bacak Toplam (€)"]
    _write_header_row(ws, 2, 2, leg_hdrs)
    ws.row_dimensions[2].height = 18

    leg_col_keys = [
        "order_id","leg_seq","from_node","to_node","mode",
        "depart_day","arrive_day","depart_date","arrive_date",
        "distance_km","tons",
        "transport_cost","emission_cost","handling_cost","leg_cost",
    ]
    for i, (_, r) in enumerate(legs_df.iterrows()):
        row_idx = 3 + i
        ws.row_dimensions[row_idx].height = 15
        fill = stripe(i)
        for j, col_key in enumerate(leg_col_keys):
            val = r[col_key]
            c = ws.cell(row=row_idx, column=2 + j, value=val)
            c.font   = _font(size=10)
            c.fill   = fill
            c.border = _border_thin("bottom")
            if j in (0,2,3,4):
                c.alignment = _align("center")
            elif j in (7,8):
                c.alignment = _align("center")
            else:
                c.alignment = _align("right")
            if j in (11,12,13,14) and isinstance(val, (int, float)):
                c.number_format = '#,##0.00'
            if j == 9 and isinstance(val, (int, float)):
                c.number_format = '#,##0.0'

    # Leg toplam satırı
    total_row_leg = 3 + len(legs_df)
    ws.row_dimensions[total_row_leg].height = 18
    c_lbl = ws.cell(row=total_row_leg, column=2, value="TOPLAM")
    c_lbl.font = _font(bold=True, color=_C_WHITE, size=10)
    c_lbl.fill = _fill(_C_DARK_NAVY)
    c_lbl.alignment = _align("center")
    ws.merge_cells(f"B{total_row_leg}:K{total_row_leg}")
    for j, col_key in enumerate(["transport_cost","emission_cost","handling_cost","leg_cost"]):
        c = ws.cell(row=total_row_leg, column=13 + j,
                    value=round(legs_df[col_key].sum(), 2))
        c.font = _font(bold=True, color=_C_WHITE, size=10)
        c.fill = _fill(_C_DARK_NAVY)
        c.alignment = _align("right")
        c.number_format = '#,##0.00'

    ws.freeze_panes = "B3"

    # ── Kaydet ───────────────────────────────────────────────────────────────
    wb.save(out_path)


def write_xlsx_log(
    log:      List[dict],
    out_path: str | Path,
    solver_tag: str = f"ALNS v{VERSION}",
):
    """
    ALNS iterasyon logunu ayrı bir XLSX dosyasına yazar.
    Satırlar: her kayıtlı iterasyon.
    NEW_BEST satırları sarı vurgu + yeşil font ile işaretlenir.
    """
    wb  = Workbook()
    ws  = wb.active
    ws.title = "ALNS Log"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2

    for col_ltr, w in zip("BCDEFGHIJKLMN",
                           [8, 14, 14, 14, 12, 8, 15, 15, 10, 12, 14, 12, 10]):
        ws.column_dimensions[col_ltr].width = w

    # Başlık bandı
    ws["B1"].value     = f"ALNS Convergence Log — İterasyon Günlüğü  |  {solver_tag}"
    ws["B1"].font      = _font(size=14, bold=True, color=_C_WHITE)
    ws["B1"].fill      = _fill(_C_DARK_NAVY)
    ws["B1"].alignment = _align("left", "center")
    ws.merge_cells("B1:N1")
    ws.row_dimensions[1].height = 28

    log_hdrs = ["İter.", "En İyi (€)", "Mevcut (€)", "Aday (€)",
                "Karar", "q", "Destroy Op.", "Repair Op.",
                "Sıcaklık (T)", "İyileşme %", "Geçen (s)", "İter. (ms)"]
    _write_header_row(ws, 2, 2, log_hdrs)
    ws.row_dimensions[2].height = 18

    log_col_keys = [
        "iter", "best", "current", "candidate",
        "decision", "q_removed", "destroy_op", "repair_op",
        "temperature", "pct_improvement", "elapsed_s", "iter_ms",
    ]

    def stripe(row_idx):
        return _fill(_C_ALT_ROW) if row_idx % 2 == 0 else _fill(_C_WHITE)

    for i, row_d in enumerate(log):
        row_idx = 3 + i
        ws.row_dimensions[row_idx].height = 15
        fill = stripe(i)
        is_new_best = (row_d.get("decision") == "NEW_BEST")

        for j, col_key in enumerate(log_col_keys):
            val = row_d.get(col_key)
            if val is None:
                val = "—"
            c = ws.cell(row=row_idx, column=2 + j, value=val)
            c.border = _border_thin("bottom")

            if is_new_best:
                c.fill = _fill("FFF9C4")
                c.font = _font(size=10, bold=True,
                               color=_C_GREEN if j == 1 else "000000")
            else:
                c.fill = fill
                if j == 4 and isinstance(val, str):
                    col_v = {
                        "NEW_BEST":  _C_GREEN,
                        "IMPROVE":   _C_BLUE_VAL,
                        "SA_ACCEPT": _C_AMBER,
                        "REJECT":    "888888",
                    }.get(val, "000000")
                    c.font = _font(size=10, bold=(val == "NEW_BEST"), color=col_v)
                else:
                    c.font = _font(size=10)

            if j == 0:
                c.alignment = _align("center")
            elif j in (4, 5, 6, 7):
                c.alignment = _align("center")
            else:
                c.alignment = _align("right")

            if j in (1, 2, 3) and isinstance(val, (int, float)):
                c.number_format = '#,##0.00'
            if j == 8 and isinstance(val, (int, float)):
                c.number_format = '0.0000'
            if j == 9 and isinstance(val, (int, float)):
                c.number_format = '0.0000'

    # Son satır özeti (istatistik)
    n_new_best = sum(1 for r in log if r.get("decision") == "NEW_BEST")
    n_improve  = sum(1 for r in log if r.get("decision") == "IMPROVE")
    n_accept   = sum(1 for r in log if r.get("decision") == "SA_ACCEPT")
    n_reject   = sum(1 for r in log if r.get("decision") == "REJECT")
    summary_row = 3 + len(log) + 1
    ws.row_dimensions[summary_row].height = 18

    summary_vals = [
        ("B", "ÖZET"),
        ("C", f"Kayıtlı iter: {len(log)}"),
        ("D", f"NEW_BEST: {n_new_best}"),
        ("E", f"IMPROVE: {n_improve}"),
        ("F", f"SA_ACCEPT: {n_accept}"),
        ("G", f"REJECT: {n_reject}"),
    ]
    for col_ltr, txt in summary_vals:
        col_idx = ord(col_ltr) - ord("A") + 1
        c = ws.cell(row=summary_row, column=col_idx, value=txt)
        c.font      = _font(size=10, bold=True, color=_C_WHITE)
        c.fill      = _fill(_C_DARK_NAVY)
        c.alignment = _align("center")

    ws.freeze_panes = "B3"
    wb.save(out_path)




def relax_deadlines_from_diagnosis(inst: Instance, diag_df: pd.DataFrame,
                                   buffer_days: int = 0,
                                   out_csv: str | Path = "deadline_relaxation_report.csv") -> int:
    """
    Strict inputta tekil olarak infeasible olan siparişleri en erken fiziksel
    varış gününe kadar gevşetir. Bu, orijinal modeli değiştiren kontrollü bir
    feasibility-recovery katmanıdır; yapılan tüm değişiklikler CSV'ye yazılır.
    """
    changes = []
    for _, row in diag_df.iterrows():
        if row.get("status") != "individually_infeasible":
            continue
        oid = str(row["order_id"])
        earliest = row.get("earliest_arrival_day")
        if pd.isna(earliest) or oid not in inst.orders:
            continue
        old_deadline = inst.orders[oid].deadline
        new_deadline = max(old_deadline, int(math.ceil(float(earliest))) + buffer_days)
        if new_deadline > old_deadline:
            inst.orders[oid] = replace(inst.orders[oid], deadline=new_deadline)
            changes.append({
                "order_id": oid,
                "old_deadline_day": old_deadline,
                "new_deadline_day": new_deadline,
                "added_days": new_deadline - old_deadline,
                "earliest_arrival_day": int(math.ceil(float(earliest))),
                "destination": row.get("destination"),
                "tons": row.get("tons"),
                "suggested_source": row.get("suggested_source"),
            })
    pd.DataFrame(changes).to_csv(out_csv, index=False)
    return len(changes)

# ─────────────────────────────────────────────────────────────────────────────
# GİRİŞ NOKTASI
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # FIX-7+8: Hardcoded parametreler argparse'a taşındı.
    # FIX-9: Test amaçlı global +30 gün buffer kaldırıldı;
    #         gerçek teşhis bazlı deadline gevşetme uygulanıyor.
    ap = argparse.ArgumentParser(description=f"ALNS v{VERSION} — CBAM Multimodal Transportation")
    ap.add_argument("--input",            required=True,
                    help="Excel girdi dosyası (.xlsx)")
    ap.add_argument("--max_iter",         type=int,   default=2000,
                    help="Maksimum iterasyon (varsayılan: 2000)")
    ap.add_argument("--time_limit",       type=float, default=300.0,
                    help="Süre limiti saniye (varsayılan: 300)")
    ap.add_argument("--alpha",            type=float, default=0.9975,
                    help="SA soğuma faktörü (varsayılan: 0.9975)")
    ap.add_argument("--seed",             type=int,   default=None,
                    help="Rastgele tohum (varsayılan: None → her run'da farklı). "
                         "Tekrarlanabilir sonuç için açıkça bir tam sayı verin, örn. --seed 42")
    ap.add_argument("--out_prefix",       default=f"alns_v{VERSION}",
                    help=f"CSV çıktı dosya öneki (varsayılan: alns_v{VERSION})")
    ap.add_argument("--soft_arc_cap",     action="store_true",
                    help="Yay kapasitesini soft (yumuşak) olarak uygula; "
                         "aşım raporlanır ama çözüm reddedilmez")
    ap.add_argument("--arc_overflow_penalty", type=float, default=5000.0,
                    help="Soft arc kapasitede her 1 ton aşım için ceza (€ / ton). Varsayılan: 5000")
    ap.add_argument("--deadline_buffer",  type=int, default=0,
                    help="Tüm siparişlere eklenecek operasyonel buffer gün (varsayılan: 0)")
    ap.add_argument("--allow_deadline_relaxation", action="store_true",
                    help="Varsayılan kapalıdır. Açılırsa teşhis bazlı deadline gevşetmeye izin verir.")
    ap.add_argument("--skip_diagnosis", action="store_true", default=True,
                    help=f"Tekil fizibilite teşhis taramasını atla. v{VERSION}'da varsayılan açık.")
    ap.add_argument("--run_diagnosis", dest="skip_diagnosis", action="store_false",
                    help="Tekil fizibilite teşhis taramasını çalıştır. Büyük instance'ta yavaş olabilir.")
    ap.add_argument("--source_probe", type=int, default=2,
                    help="Başlangıç inşasında önce denenecek kaynak sayısı. Varsayılan: 2")
    ap.add_argument("--initial_random_tries", type=int, default=2,
                    help="Başlangıç çözümü için rastgele blok-shuffle deneme sayısı. Varsayılan: 2")
    ap.add_argument("--no_xlsx", action="store_true",
                    help="Hızlı test için xlsx rapor/log yazma.")
    ap.add_argument("--silent",           action="store_true",
                    help="Konsol çıktısını kapat")
    args = ap.parse_args()

    MAX_ITER   = args.max_iter
    TIME_LIMIT = args.time_limit
    ALPHA      = args.alpha
    SEED       = args.seed
    OUT_PREFIX = args.out_prefix
    VERBOSE    = not args.silent

    input_path = Path(args.input)
    # FIX: Çıktılar her zaman çalışma dizinine yazılır.
    # args.input read-only bir klasörde (örn. /mnt/user-data/uploads) olabilir;
    # .parent kullanmak PermissionError'a yol açıyordu.
    script_dir = Path.cwd()

    if VERBOSE:
        print(f"\n{'═'*60}")
        print(f"  CBAM Multimodal ALNS v{VERSION} — EMU676")
        print(f"{'═'*60}")
        print(f"  Girdi       : {input_path}")
        print(f"  Maks. iter  : {MAX_ITER}")
        print(f"  Süre limiti : {TIME_LIMIT}s")
        print(f"  Tohum       : {SEED}")
        print(f"  Soft arc    : {args.soft_arc_cap}")
        if args.soft_arc_cap:
            print(f"  Arc cezası  : {args.arc_overflow_penalty:,.0f} €/ton overflow")
        print(f"  DL buffer   : +{args.deadline_buffer} gün")

    t0   = time.time()
    inst = Instance(input_path)

    # Soft arc capacity: komut satırı argümanıyla kontrol edilir
    inst.soft_arc_capacity = args.soft_arc_cap
    inst.arc_overflow_penalty_eur_per_ton = args.arc_overflow_penalty
    inst.source_probe = args.source_probe
    inst.initial_random_tries = args.initial_random_tries

    if VERBOSE:
        print(f"\n  Düğümler  : {len(inst.nodes)}")
        print(f"  Yaylar    : {len(inst.arcs)}")
        print(f"  Siparişler: {len(inst.orders)}")
        print(f"  Kaynaklar : {len(inst.sources)}")
        print(f"  CBAM φ_r  : {inst.phi_r} ({inst.plan_year})")

    # ── Fizibilite teşhisi ──────────────────────────────────────────────────
    # Büyük inputta bu tarama pahalıdır. --skip_diagnosis hızlı deney için atlar.
    if args.skip_diagnosis:
        ok = True
        diag_df = pd.DataFrame()
        inst._feasible_source_count = {oid: 999 for oid in inst.orders}
        if VERBOSE:
            print("\n  Fizibilite teşhisi atlandı (--skip_diagnosis).")
    else:
        if VERBOSE:
            print("\n  Fizibilite teşhisi yapılıyor…")
        ok, diag_df = diagnose_instance_feasibility(
            inst, write_csv=True,
            out_csv=str(script_dir / "infeasible_orders_diagnosis.csv"),
        )
        inst._feasible_source_count = {
            str(row["order_id"]): int(row["feasible_source_count"])
            for _, row in diag_df.iterrows()
        }

    # Operasyonel deadline buffer (istenirse tüm siparişlere eklenir)
    if args.deadline_buffer > 0:
        dl_rows = []
        for oid, o in list(inst.orders.items()):
            inst.orders[oid] = replace(o, deadline=o.deadline + args.deadline_buffer)
            dl_rows.append({
                "order_id": oid,
                "old_deadline_day": o.deadline,
                "new_deadline_day": o.deadline + args.deadline_buffer,
                "added_days": args.deadline_buffer,
                "reason": "operational_buffer",
            })
        pd.DataFrame(dl_rows).to_csv(
            script_dir / "deadline_relaxation_report.csv", index=False)
        if VERBOSE:
            print(f"  Operasyonel buffer: tüm deadline'lara +{args.deadline_buffer} gün eklendi.")
    elif not ok:
        if args.allow_deadline_relaxation:
            # v10: Varsayılan olarak gevşetme yapılmaz. Kullanıcı özellikle isterse
            # yalnızca tekil infeasible siparişler için minimum gevşetme uygulanır.
            n_relaxed = relax_deadlines_from_diagnosis(
                inst, diag_df, buffer_days=0,
                out_csv=script_dir / "deadline_relaxation_report.csv",
            )
            if VERBOSE and n_relaxed > 0:
                print(f"  Teşhis bazlı deadline gevşetme: {n_relaxed} sipariş → "
                      "deadline_relaxation_report.csv")
        else:
            raise RuntimeError(
                "Orijinal deadline seti altında tekil fizibilite teşhisi başarısız. "
                "v10 varsayılan olarak deadline relaxation yapmaz. Girdi verisini düzeltin "
                "veya bilinçli olarak --allow_deadline_relaxation kullanın."
            )

    best, log = alns(
        inst,
        max_iter   = MAX_ITER,
        time_limit = TIME_LIMIT,
        alpha      = ALPHA,
        seed       = SEED,   # None → her run'da farklı; alns() içinde gerçek seed yazdırılır
        verbose    = VERBOSE,
    )

    # FIX-v11: Gerçekte kullanılan seed'i log'dan oku (alns() verbose modda yazdırır).
    # XLSX solver_tag için SEED'i güncelle.
    if SEED is None and log:
        pass  # seed alns() içinde zaten ekrana yazdırıldı

    elapsed = time.time() - t0

    # Başlangıç sırasında global kapasite/takvim çakışması için yapılan ek deadline
    # gevşetmelerini kaydet.
    dyn_relax = getattr(inst, "_dynamic_deadline_relaxations", [])
    if dyn_relax:
        pd.DataFrame(dyn_relax).to_csv(script_dir / "dynamic_deadline_relaxation_report.csv", index=False)
        if VERBOSE:
            print(f"\n  Ek global gevşetme: {len(dyn_relax)} kayıt → dynamic_deadline_relaxation_report.csv")

    # Bağımsız doğrulama
    assert validate(inst, best), "Çözüm doğrulama başarısız!"

    if VERBOSE:
        print_report(inst, best, elapsed, len(log))

    # ── İki ayrı XLSX çıktısı ────────────────────────────────────────────────
    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    solver_tag = f"ALNS v{VERSION}  |  seed={SEED}  |  iter={len(log)}"

    if not args.no_xlsx:
        # 1) Ana rapor (Dashboard + Order Detail + Route Legs)
        xlsx_report = script_dir / f"{OUT_PREFIX}_report_{timestamp}.xlsx"
        write_xlsx_report(
            inst       = inst,
            sol        = best,
            elapsed    = elapsed,
            out_path   = xlsx_report,
            solver_tag = solver_tag,
        )

        # 2) ALNS log (ayrı dosya)
        xlsx_log = script_dir / f"{OUT_PREFIX}_log_{timestamp}.xlsx"
        write_xlsx_log(
            log        = log,
            out_path   = xlsx_log,
            solver_tag = solver_tag,
        )

        if VERBOSE:
            print(f"\n  ✓ Rapor    : {xlsx_report}")
            print(f"  ✓ ALNS Log : {xlsx_log}")
    elif VERBOSE:
        print("\n  XLSX rapor/log yazımı atlandı (--no_xlsx).")

    # v9: Final fizibilite ve kapasite aşımı raporları.
    # Overflow olmasa bile CSV üretilir; böylece çıktı paketi her run'da aynı kalır.
    _, arc_used_report = rebuild_capacity(inst, best)
    overflow_rows = arc_overflow_rows(inst, arc_used_report)
    ov_path = script_dir / f"{OUT_PREFIX}_arc_overflow_{timestamp}.csv"
    ov_cols = ["from_node", "to_node", "mode", "depart_day",
               "used_ton", "capacity_ton", "overflow_ton"]
    pd.DataFrame(overflow_rows, columns=ov_cols).to_csv(ov_path, index=False)

    summary_path = script_dir / f"{OUT_PREFIX}_solution_summary_{timestamp}.csv"
    summary_rows = [{
        "solution_status": "strict_feasible" if best.strict_feasible else "penalized_relaxed_with_arc_overflow",
        "feasible_for_alns": best.feasible,
        "strict_feasible": best.strict_feasible,
        "base_objective_eur": best.base_obj,
        "arc_overflow_penalty_eur": best.arc_overflow_penalty,
        "total_objective_eur": best.obj,
        "overflow_arc_count": best.overflow_arc_count,
        "total_overflow_ton": best.total_overflow_ton,
        "deadline_violation_count": best.deadline_violation_count,
        "total_lateness_days": best.total_lateness_days,
        "max_lateness_days": best.max_lateness_days,
        "orders_assigned": len(best.assignments),
        "orders_total": len(inst.orders),
        "elapsed_s": elapsed,
        "iterations_logged": len(log),
        "soft_arc_capacity_enabled": bool(getattr(inst, "soft_arc_capacity", False)),
    }]
    pd.DataFrame(summary_rows).to_csv(summary_path, index=False)

    if VERBOSE:
        if best.strict_feasible:
            print("  ✓ Strict feasible solution found: yay kapasite aşımı yok.")
        else:
            print("  ⚠ Penalized relaxed solution found: yay kapasite aşımı içeriyor.")
        print(f"  ✓ Çözüm özeti          : {summary_path}")
        print(f"  ✓ Yay kapasite raporu  : {ov_path}")

    return best


if __name__ == "__main__":
    if len(sys.argv) == 1:
        # ── DOSYA ADINI BURAYA YAZ ──────────────────────────────────────────
        # FIX-v11: seed argümanı kaldırıldı → her çalıştırmada farklı seed.
        # Tekrarlanabilir sonuç için "--seed", "42" ekleyin.
        sys.argv = [sys.argv[0],
                    "--input",      "30node_20order_input.xlsx",
                    "--max_iter",   "800",
                    "--time_limit", "320",
                    "--out_prefix", f"alns_v{VERSION}",
                    "--soft_arc_cap"]
        # ────────────────────────────────────────────────────────────────────
    main()
