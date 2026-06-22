from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


class IBKStore:
    """SQLite arxiv: yuklangan fayllar snapshoti va nazoratdan yechilgan qatorlar."""

    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.ensure_schema()

    def connect(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def ensure_schema(self):
        with self.connect() as conn:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS snapshots (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    report_id TEXT UNIQUE,
                    file_name TEXT,
                    file_path TEXT,
                    deposit_path TEXT,
                    report_date TEXT,
                    uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS active_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    snapshot_id INTEGER,
                    item_key TEXT,
                    decl TEXT,
                    item_no TEXT,
                    regime TEXT,
                    post TEXT,
                    stir TEXT,
                    company TEXT,
                    gtd_date TEXT,
                    hs_code TEXT,
                    goods TEXT,
                    weight REAL DEFAULT 0,
                    value REAL DEFAULT 0,
                    payment REAL DEFAULT 0,
                    partiya INTEGER DEFAULT 0,
                    country TEXT,
                    warehouse TEXT,
                    reason TEXT,
                    FOREIGN KEY(snapshot_id) REFERENCES snapshots(id)
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS released_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    base_snapshot_id INTEGER,
                    final_snapshot_id INTEGER,
                    item_key TEXT,
                    decl TEXT,
                    item_no TEXT,
                    stir TEXT,
                    company TEXT,
                    regime TEXT,
                    post TEXT,
                    release_type TEXT,
                    released_partiya INTEGER DEFAULT 0,
                    released_weight REAL DEFAULT 0,
                    released_value REAL DEFAULT 0,
                    released_payment REAL DEFAULT 0,
                    released_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            cur.execute("CREATE INDEX IF NOT EXISTS idx_active_snapshot ON active_items(snapshot_id)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_active_key ON active_items(item_key)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_release_pair ON released_items(base_snapshot_id, final_snapshot_id)")
            for col, typ in [("transport", "TEXT"), ("source_post", "TEXT")]:
                try:
                    cur.execute(f"ALTER TABLE active_items ADD COLUMN {col} {typ}")
                except Exception:
                    pass

    def register_snapshot(self, report_id: str, file_name: str, file_path: str, deposit_path: str | None, report_date: str) -> int:
        with self.connect() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT OR REPLACE INTO snapshots(report_id, file_name, file_path, deposit_path, report_date)
                VALUES (?, ?, ?, ?, ?)
            """, (report_id, file_name, file_path, deposit_path, report_date))
            row = cur.execute("SELECT id FROM snapshots WHERE report_id=?", (report_id,)).fetchone()
            return int(row["id"])

    def insert_active_items(self, snapshot_id: int, rows: list[dict[str, Any]]):
        with self.connect() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM active_items WHERE snapshot_id=?", (snapshot_id,))
            cur.executemany("""
                INSERT INTO active_items(
                    snapshot_id, item_key, decl, item_no, regime, post, stir, company, gtd_date,
                    hs_code, goods, weight, value, payment, partiya, country, warehouse, reason,
                    transport, source_post
                ) VALUES (
                    :snapshot_id, :item_key, :decl, :item_no, :regime, :post, :stir, :company, :gtd_date,
                    :hs_code, :goods, :weight, :value, :payment, :partiya, :country, :warehouse, :reason,
                    :transport, :source_post
                )
            """, [{**r, "snapshot_id": snapshot_id, "transport": r.get("transport",""), "source_post": r.get("source_post","")} for r in rows])

    def snapshot_id_by_date(self, date_text: str) -> int | None:
        with self.connect() as conn:
            row = conn.execute("SELECT id FROM snapshots WHERE report_date=? ORDER BY id DESC LIMIT 1", (date_text,)).fetchone()
            return int(row["id"]) if row else None

    def list_snapshots(self) -> list[dict[str, Any]]:
        with self.connect() as conn:
            rows = conn.execute("SELECT * FROM snapshots ORDER BY report_date DESC, id DESC").fetchall()
            return [dict(r) for r in rows]

    @staticmethod
    def _normalize_periods(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
        """report_date ustuni 'kk.oo.yyyy' (dd.mm.yyyy) formatda saqlanadi -
        bu satr sifatida noto'g'ri tartiblanadi (masalan '01.05.2026' <
        '15.04.2026'). Shu sabab haqiqiy sanaga o'tkazib, ISO ('yyyy-mm-dd')
        ko'rinishda qaytaramiz va xronologik tartiblaymiz."""
        df = df.copy()
        df["date"] = df["date"].astype(str)
        dt = pd.to_datetime(df["date"], format="%d.%m.%Y", errors="coerce")
        # ehtiyot uchun, agar boshqa formatda kelsa ham urinib ko'ramiz
        mask = dt.isna()
        if mask.any():
            dt2 = pd.to_datetime(df.loc[mask, "date"], errors="coerce")
            dt.loc[mask] = dt2
        df["_dt"] = dt
        df = df.dropna(subset=["_dt"])
        df["date"] = df["_dt"].dt.strftime("%Y-%m-%d")
        df = df.sort_values("_dt")
        periods = df.drop_duplicates("date")["date"].tolist()
        return df, periods

    def available_dates(self) -> list[str]:
        """Arxivdagi barcha snapshot sanalarini dd.mm.yyyy formatda, xronologik
        kamayish tartibida qaytaradi."""
        with self.connect() as conn:
            rows = conn.execute(
                "SELECT DISTINCT report_date FROM snapshots ORDER BY report_date DESC"
            ).fetchall()
        raw = [r[0] for r in rows if r[0]]
        def _dt(s):
            try:
                return datetime.strptime(s, "%d.%m.%Y")
            except Exception:
                return datetime.min
        return sorted(raw, key=_dt, reverse=True)

    def country_flow_by_date(self, date_text: str) -> list[dict[str, Any]]:
        """Berilgan sana (dd.mm.yyyy) snapshot uchun davlatlar kesimida
        jami qiymat/vazn/partiya. Qiymat bo'yicha kamayish tartibida qaytariladi."""
        with self.connect() as conn:
            df = pd.read_sql_query(
                """
                SELECT a.country AS name,
                       SUM(a.value)   AS qiymat,
                       SUM(a.weight)  AS vazn,
                       SUM(a.partiya) AS partiya
                FROM active_items a
                JOIN snapshots s ON s.id = a.snapshot_id
                WHERE s.report_date = ?
                  AND a.country IS NOT NULL AND a.country != ''
                GROUP BY a.country
                ORDER BY qiymat DESC
                """,
                conn,
                params=(date_text,),
            )
        if df.empty:
            return []
        return [
            {
                "name": str(r["name"]),
                "qiymat": float(r["qiymat"]),
                "vazn": float(r["vazn"]),
                "partiya": int(r["partiya"]),
            }
            for _, r in df.iterrows()
        ]

    def company_series_all(self) -> dict[str, Any]:
        """Har bir korxona (STIR) bo'yicha har bir snapshot (davr) kesimida
        jami qiymat/vazn/partiya. Davrlar bo'yicha tendensiya tahlili uchun.
        Davrlar 'yyyy-mm-dd' ko'rinishda, xronologik tartibda qaytariladi."""
        with self.connect() as conn:
            df = pd.read_sql_query(
                """
                SELECT s.report_date as date, a.stir as stir, a.company as company,
                       SUM(a.value) as value, SUM(a.weight) as weight, SUM(a.partiya) as partiya
                FROM active_items a
                JOIN snapshots s ON s.id = a.snapshot_id
                WHERE a.stir IS NOT NULL AND a.stir != ''
                GROUP BY s.report_date, a.stir
                """,
                conn,
            )
        if df.empty:
            return {"periods": [], "companies": []}
        df, periods = self._normalize_periods(df)
        name_map = df.groupby("stir")["company"].last().to_dict()
        companies = []
        for stir, g in df.groupby("stir"):
            by_date = {row["date"]: row for _, row in g.iterrows()}
            series = []
            for p in periods:
                row = by_date.get(p)
                series.append({
                    "date": p,
                    "value": float(row["value"]) if row is not None else 0.0,
                    "weight": float(row["weight"]) if row is not None else 0.0,
                    "partiya": int(row["partiya"]) if row is not None else 0,
                })
            companies.append({"stir": str(stir), "company": str(name_map.get(stir, "")), "series": series})
        return {"periods": periods, "companies": companies}

    def goods_series_all(self) -> dict[str, Any]:
        """Har bir tovar (HS kod) bo'yicha har bir davrda jami qiymat/vazn/partiya.
        Tovarlar bo'yicha yangi qo'shilgan/chiqib ketgan tahlili uchun."""
        with self.connect() as conn:
            df = pd.read_sql_query(
                """
                SELECT s.report_date as date, a.hs_code as hs_code, a.goods as goods,
                       SUM(a.value) as value, SUM(a.weight) as weight, SUM(a.partiya) as partiya
                FROM active_items a
                JOIN snapshots s ON s.id = a.snapshot_id
                WHERE a.hs_code IS NOT NULL AND a.hs_code != ''
                GROUP BY s.report_date, a.hs_code
                """,
                conn,
            )
        if df.empty:
            return {"periods": [], "goods": []}
        df, periods = self._normalize_periods(df)
        name_map = df.groupby("hs_code")["goods"].last().to_dict()
        items = []
        for hs, g in df.groupby("hs_code"):
            by_date = {row["date"]: row for _, row in g.iterrows()}
            series = []
            for p in periods:
                row = by_date.get(p)
                series.append({
                    "date": p,
                    "value": float(row["value"]) if row is not None else 0.0,
                    "weight": float(row["weight"]) if row is not None else 0.0,
                    "partiya": int(row["partiya"]) if row is not None else 0,
                })
            name = str(name_map.get(hs, "") or "")
            short = (name[:42].rstrip() + "\u2026") if len(name) > 42 else name
            items.append({"hs_code": str(hs), "goods": name, "goods_short": short or str(hs), "series": series})
        return {"periods": periods, "goods": items}

    def warehouse_series_all(self) -> dict[str, Any]:
        """Har bir ombor bo'yicha har bir davr kesimida jami qiymat/vazn/partiya."""
        with self.connect() as conn:
            df = pd.read_sql_query(
                """
                SELECT s.report_date as date, a.warehouse as warehouse,
                       SUM(a.value) as value, SUM(a.weight) as weight, SUM(a.partiya) as partiya
                FROM active_items a
                JOIN snapshots s ON s.id = a.snapshot_id
                WHERE a.warehouse IS NOT NULL AND a.warehouse != ''
                GROUP BY s.report_date, a.warehouse
                """,
                conn,
            )
        if df.empty:
            return {"periods": [], "warehouses": []}
        df, periods = self._normalize_periods(df)
        items = []
        for wh, g in df.groupby("warehouse"):
            by_date = {row["date"]: row for _, row in g.iterrows()}
            series = []
            for p in periods:
                row = by_date.get(p)
                series.append({
                    "date": p,
                    "value": float(row["value"]) if row is not None else 0.0,
                    "weight": float(row["weight"]) if row is not None else 0.0,
                    "partiya": int(row["partiya"]) if row is not None else 0,
                })
            items.append({"warehouse": str(wh), "series": series})
        items.sort(key=lambda x: sum(s["value"] for s in x["series"]), reverse=True)
        return {"periods": periods, "warehouses": items}

    def transport_series_all(self) -> dict[str, Any]:
        """Har bir transport turi bo'yicha har bir davr kesimida jami qiymat/vazn/partiya."""
        with self.connect() as conn:
            df = pd.read_sql_query(
                """
                SELECT s.report_date as date, a.transport as transport,
                       SUM(a.value) as value, SUM(a.weight) as weight, SUM(a.partiya) as partiya
                FROM active_items a
                JOIN snapshots s ON s.id = a.snapshot_id
                WHERE a.transport IS NOT NULL AND a.transport != ''
                GROUP BY s.report_date, a.transport
                """,
                conn,
            )
        if df.empty:
            return {"periods": [], "transports": []}
        df, periods = self._normalize_periods(df)
        items = []
        for tr, g in df.groupby("transport"):
            by_date = {row["date"]: row for _, row in g.iterrows()}
            series = []
            for p in periods:
                row = by_date.get(p)
                series.append({
                    "date": p,
                    "value": float(row["value"]) if row is not None else 0.0,
                    "weight": float(row["weight"]) if row is not None else 0.0,
                    "partiya": int(row["partiya"]) if row is not None else 0,
                })
            items.append({"transport": str(tr), "series": series})
        items.sort(key=lambda x: sum(s["value"] for s in x["series"]), reverse=True)
        return {"periods": periods, "transports": items}

    def transport_company_series_all(self) -> dict[str, Any]:
        """Har bir transport turi kesimida korxonalarning davriy aylanmasi."""
        with self.connect() as conn:
            df = pd.read_sql_query(
                """
                SELECT s.report_date as date, a.transport as transport,
                       a.stir as stir, a.company as company,
                       SUM(a.value) as value, SUM(a.weight) as weight, SUM(a.partiya) as partiya
                FROM active_items a
                JOIN snapshots s ON s.id = a.snapshot_id
                WHERE a.transport IS NOT NULL AND a.transport != ''
                  AND a.stir IS NOT NULL AND a.stir != ''
                GROUP BY s.report_date, a.transport, a.stir
                """,
                conn,
            )
        if df.empty:
            return {"periods": [], "transports": []}
        df, periods = self._normalize_periods(df)
        name_map = df.groupby("stir")["company"].last().to_dict()
        result = {}
        for (tr, stir), g in df.groupby(["transport", "stir"]):
            if tr not in result:
                result[tr] = {}
            by_date = {row["date"]: row for _, row in g.iterrows()}
            series = []
            for p in periods:
                row = by_date.get(p)
                series.append({
                    "date": p,
                    "value": float(row["value"]) if row is not None else 0.0,
                    "weight": float(row["weight"]) if row is not None else 0.0,
                    "partiya": int(row["partiya"]) if row is not None else 0,
                })
            result[tr][str(stir)] = {
                "stir": str(stir),
                "company": str(name_map.get(stir, "")),
                "series": series,
                "total_value": sum(s["value"] for s in series),
            }
        transports = []
        for tr, companies in result.items():
            clist = sorted(companies.values(), key=lambda x: x["total_value"], reverse=True)
            transports.append({"transport": str(tr), "companies": clist})
        transports.sort(key=lambda x: sum(c["total_value"] for c in x["companies"]), reverse=True)
        return {"periods": periods, "transports": transports}

    def country_transport_summary(self) -> dict[str, Any]:
        """Har bir davlat uchun dominant transport turi (post ustunidan olinadi)."""
        with self.connect() as conn:
            df = pd.read_sql_query(
                """
                SELECT a.country, a.post,
                       SUM(a.weight) as weight
                FROM active_items a
                WHERE a.country IS NOT NULL AND a.country != ''
                  AND a.post    IS NOT NULL AND a.post    != ''
                GROUP BY a.country, a.post
                """,
                conn,
            )
        if df.empty:
            return {}

        def _post_transport(post: str) -> str:
            p = str(post).lower()
            if "temir" in p or "railroad" in p:
                return "Temir yo'l"
            # Toshkent-AERO aeroport bojxonasi: barcha postlar havo transporti
            return "Avia"

        df["transport"] = df["post"].apply(_post_transport)
        result: dict[str, Any] = {}
        for country, group in df.groupby("country"):
            transports: dict[str, float] = {}
            for _, row in group.iterrows():
                t = row["transport"]
                transports[t] = transports.get(t, 0.0) + float(row["weight"])
            dominant = max(transports, key=lambda k: transports[k])
            result[str(country)] = {"dominant": dominant, "transports": transports}
        return result

    def compute_released(self, base_snapshot_id: int, final_snapshot_id: int) -> dict[str, Any]:
        with self.connect() as conn:
            base = pd.read_sql_query("SELECT * FROM active_items WHERE snapshot_id=?", conn, params=(base_snapshot_id,))
            final = pd.read_sql_query("SELECT * FROM active_items WHERE snapshot_id=?", conn, params=(final_snapshot_id,))
        if base.empty:
            return {"total": {}, "rows": [], "partial": [], "unreleased": []}
        return compute_released_frames(base, final)


def compute_released_frames(base: pd.DataFrame, final: pd.DataFrame) -> dict[str, Any]:
    if final.empty:
        final = pd.DataFrame(columns=base.columns)
    base_g = base.groupby("item_key", dropna=False).agg(
        decl=("decl", "first"), item_no=("item_no", "first"), stir=("stir", "first"), company=("company", "first"),
        regime=("regime", "first"), post=("post", "first"), warehouse=("warehouse", "first"),
        partiya=("partiya", "sum"), weight=("weight", "sum"), value=("value", "sum"), payment=("payment", "sum"),
    )
    final_g = final.groupby("item_key", dropna=False).agg(
        partiya=("partiya", "sum"), weight=("weight", "sum"), value=("value", "sum"), payment=("payment", "sum"),
    ) if not final.empty else pd.DataFrame()

    rows = []
    unreleased = []
    for key, b in base_g.iterrows():
        has_final = key in final_g.index if not final_g.empty else False
        f = final_g.loc[key] if has_final else None
        remain_partiya = int(f.partiya) if f is not None else 0
        remain_weight = float(f.weight) if f is not None else 0.0
        remain_value = float(f.value) if f is not None else 0.0
        remain_payment = float(f.payment) if f is not None else 0.0
        released = {
            "key": {"decl": str(b.decl), "item_no": str(b.item_no)},
            "decl": str(b.decl),
            "item_no": str(b.item_no),
            "korxona": str(b.company),
            "stir": str(b.stir),
            "regime": str(b.regime),
            "post": str(b.post),
            "warehouse": str(b.warehouse),
            "base_partiya": int(b.partiya),
            "base_vazn": float(b.weight),
            "base_qiymat": float(b.value),
            "base_tolov": float(b.payment),
            "remain_partiya": remain_partiya,
            "remain_vazn": remain_weight,
            "remain_qiymat": remain_value,
            "remain_tolov": remain_payment,
            "released_partiya": max(0, int(b.partiya) - remain_partiya),
            "released_vazn": max(0.0, float(b.weight) - remain_weight),
            "released_qiymat": max(0.0, float(b.value) - remain_value),
            "released_tolov": max(0.0, float(b.payment) - remain_payment),
        }
        released["current_vazn"] = released["remain_vazn"]
        released["current_qiymat"] = released["remain_qiymat"]
        released["current_partiya"] = released["remain_partiya"]
        released["released_pct"] = (released["released_qiymat"] / released["base_qiymat"] * 100) if released["base_qiymat"] else 0.0
        if released["released_vazn"] <= 0.005 and released["released_qiymat"] <= 0.005 and released["released_partiya"] <= 0 and released["released_tolov"] <= 0.005:
            unreleased.append(released)
        else:
            released["release_type"] = "qisman" if has_final and released["released_partiya"] == 0 else "to'liq"
            rows.append(released)
    rows.sort(key=lambda x: (x["released_qiymat"], x["released_vazn"], x["released_partiya"]), reverse=True)
    total = {
        "korxona": "Jami",
        "stir": "",
        "base_vazn": sum(r["base_vazn"] for r in rows),
        "base_qiymat": sum(r["base_qiymat"] for r in rows),
        "base_partiya": sum(r["base_partiya"] for r in rows),
        "remain_vazn": sum(r["remain_vazn"] for r in rows),
        "remain_qiymat": sum(r["remain_qiymat"] for r in rows),
        "remain_partiya": sum(r["remain_partiya"] for r in rows),
        "released_vazn": sum(r["released_vazn"] for r in rows),
        "released_qiymat": sum(r["released_qiymat"] for r in rows),
        "released_partiya": sum(r["released_partiya"] for r in rows),
        "released_tolov": sum(r["released_tolov"] for r in rows),
    }
    total["released_pct"] = (total["released_qiymat"] / total["base_qiymat"] * 100) if total["base_qiymat"] else 0.0
    return {
        "total": total,
        "rows": rows[:500],
        "partial": [r for r in rows if r.get("release_type") == "qisman"][:200],
        "unreleased": sorted(unreleased, key=lambda x: (x["base_qiymat"], x["base_vazn"]), reverse=True)[:200],
        "top_released": rows[:20],
    }
