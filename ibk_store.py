from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

# (kalit, o'zbekcha nom, bugundan orqaga necha kun ichida) - kumulyativ
RELEASE_BUCKETS = [
    ("kun1", "1 kun", 1),
    ("hafta1", "1 hafta", 7),
    ("kun15", "15 kun", 15),
    ("oy1", "1 oy", 30),
    ("oy3", "3 oy", 90),
    ("oy6", "6 oy", 180),
    ("yil1", "1 yil", 365),
]


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

    def country_flow_by_date(self, date_text: str, post_filter: str = "") -> list[dict[str, Any]]:
        """Berilgan sana (dd.mm.yyyy) snapshot uchun davlatlar kesimida
        jami qiymat/vazn/partiya. Qiymat bo'yicha kamayish tartibida qaytariladi."""
        post_clause = "AND a.source_post = ?" if post_filter else ""
        params = (date_text, post_filter) if post_filter else (date_text,)
        with self.connect() as conn:
            df = pd.read_sql_query(
                f"""
                SELECT a.country AS name,
                       SUM(a.value)   AS qiymat,
                       SUM(a.weight)  AS vazn,
                       SUM(a.partiya) AS partiya
                FROM active_items a
                JOIN snapshots s ON s.id = a.snapshot_id
                WHERE s.report_date = ?
                  AND a.country IS NOT NULL AND a.country != '' {post_clause}
                GROUP BY a.country
                ORDER BY qiymat DESC
                """,
                conn,
                params=params,
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

    def company_series_all(self, post_filter: str = "") -> dict[str, Any]:
        """Har bir korxona (STIR) bo'yicha har bir snapshot (davr) kesimida
        jami qiymat/vazn/partiya. Davrlar bo'yicha tendensiya tahlili uchun.
        Davrlar 'yyyy-mm-dd' ko'rinishda, xronologik tartibda qaytariladi.
        post_filter berilsa, faqat shu source_post'ga tegishli qatorlar."""
        post_clause, params = ("AND a.source_post = ?", (post_filter,)) if post_filter else ("", ())
        with self.connect() as conn:
            df = pd.read_sql_query(
                f"""
                SELECT s.report_date as date, a.stir as stir, a.company as company,
                       SUM(a.value) as value, SUM(a.weight) as weight, SUM(a.partiya) as partiya
                FROM active_items a
                JOIN snapshots s ON s.id = a.snapshot_id
                WHERE a.stir IS NOT NULL AND a.stir != '' {post_clause}
                GROUP BY s.report_date, a.stir
                """,
                conn,
                params=params,
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

    def goods_series_all(self, post_filter: str = "") -> dict[str, Any]:
        """Har bir tovar (HS kod) bo'yicha har bir davrda jami qiymat/vazn/partiya.
        Tovarlar bo'yicha yangi qo'shilgan/chiqib ketgan tahlili uchun."""
        post_clause, params = ("AND a.source_post = ?", (post_filter,)) if post_filter else ("", ())
        with self.connect() as conn:
            df = pd.read_sql_query(
                f"""
                SELECT s.report_date as date, a.hs_code as hs_code, a.goods as goods,
                       SUM(a.value) as value, SUM(a.weight) as weight, SUM(a.partiya) as partiya
                FROM active_items a
                JOIN snapshots s ON s.id = a.snapshot_id
                WHERE a.hs_code IS NOT NULL AND a.hs_code != '' {post_clause}
                GROUP BY s.report_date, a.hs_code
                """,
                conn,
                params=params,
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

    def warehouse_series_all(self, post_filter: str = "") -> dict[str, Any]:
        """Har bir ombor bo'yicha har bir davr kesimida jami qiymat/vazn/partiya."""
        post_clause, params = ("AND a.source_post = ?", (post_filter,)) if post_filter else ("", ())
        with self.connect() as conn:
            df = pd.read_sql_query(
                f"""
                SELECT s.report_date as date, a.warehouse as warehouse,
                       SUM(a.value) as value, SUM(a.weight) as weight, SUM(a.partiya) as partiya
                FROM active_items a
                JOIN snapshots s ON s.id = a.snapshot_id
                WHERE a.warehouse IS NOT NULL AND a.warehouse != '' {post_clause}
                GROUP BY s.report_date, a.warehouse
                """,
                conn,
                params=params,
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

    def transport_series_all(self, post_filter: str = "") -> dict[str, Any]:
        """Har bir transport turi bo'yicha har bir davr kesimida jami qiymat/vazn/partiya."""
        post_clause, params = ("AND a.source_post = ?", (post_filter,)) if post_filter else ("", ())
        with self.connect() as conn:
            df = pd.read_sql_query(
                f"""
                SELECT s.report_date as date, a.transport as transport,
                       SUM(a.value) as value, SUM(a.weight) as weight, SUM(a.partiya) as partiya
                FROM active_items a
                JOIN snapshots s ON s.id = a.snapshot_id
                WHERE a.transport IS NOT NULL AND a.transport != '' {post_clause}
                GROUP BY s.report_date, a.transport
                """,
                conn,
                params=params,
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

    def transport_company_series_all(self, post_filter: str = "") -> dict[str, Any]:
        """Har bir transport turi kesimida korxonalarning davriy aylanmasi."""
        post_clause, params = ("AND a.source_post = ?", (post_filter,)) if post_filter else ("", ())
        with self.connect() as conn:
            df = pd.read_sql_query(
                f"""
                SELECT s.report_date as date, a.transport as transport,
                       a.stir as stir, a.company as company,
                       SUM(a.value) as value, SUM(a.weight) as weight, SUM(a.partiya) as partiya
                FROM active_items a
                JOIN snapshots s ON s.id = a.snapshot_id
                WHERE a.transport IS NOT NULL AND a.transport != ''
                  AND a.stir IS NOT NULL AND a.stir != '' {post_clause}
                GROUP BY s.report_date, a.transport, a.stir
                """,
                conn,
                params=params,
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

    def country_transport_summary(self, post_filter: str = "") -> dict[str, Any]:
        """Har bir davlat uchun transport turi breakdown (active_items.transport ustunidan)."""
        post_clause, params = ("AND a.source_post = ?", (post_filter,)) if post_filter else ("", ())
        with self.connect() as conn:
            df = pd.read_sql_query(
                f"""
                SELECT a.country, a.transport,
                       SUM(a.weight) as weight, SUM(a.value) as value
                FROM active_items a
                WHERE a.country   IS NOT NULL AND a.country   != ''
                  AND a.transport IS NOT NULL AND a.transport != '' {post_clause}
                GROUP BY a.country, a.transport
                """,
                conn,
                params=params,
            )
        if df.empty:
            return {}

        result: dict[str, Any] = {}
        for country, group in df.groupby("country"):
            transports: dict[str, float] = {}
            for _, row in group.iterrows():
                t = str(row["transport"])
                transports[t] = transports.get(t, 0.0) + float(row["weight"] or 0)
            dominant = max(transports, key=lambda k: transports[k])
            result[str(country)] = {"dominant": dominant, "transports": transports}
        return result

    def avia_db_stats(self, post_filter: str = "") -> dict[str, Any]:
        """active_items WHERE transport='Avia' dan qiymat statistikasi (ming $)."""
        post_clause, params = ("AND source_post = ?", (post_filter,)) if post_filter else ("", ())
        with self.connect() as conn:
            df_sum = pd.read_sql_query(
                "SELECT COUNT(DISTINCT decl) AS decl_soni, SUM(value) AS jami_qiymat,"
                " SUM(weight) AS jami_vazn, SUM(partiya) AS jami_partiya"
                f" FROM active_items WHERE transport='Avia' {post_clause}",
                conn, params=params,
            )
            df_oy = pd.read_sql_query(
                "SELECT strftime('%Y-%m', gtd_date) AS oy,"
                " COUNT(DISTINCT decl) AS decl_soni, SUM(value) AS qiymat,"
                " SUM(weight) AS vazn, SUM(partiya) AS partiya"
                f" FROM active_items WHERE transport='Avia' AND gtd_date IS NOT NULL {post_clause}"
                " GROUP BY oy ORDER BY oy DESC LIMIT 24",
                conn, params=params,
            )
            df_comp = pd.read_sql_query(
                "SELECT company, stir, COUNT(DISTINCT decl) AS decl_soni,"
                " SUM(value) AS qiymat, SUM(weight) AS vazn, SUM(partiya) AS partiya"
                f" FROM active_items WHERE transport='Avia' AND company IS NOT NULL AND company!='' {post_clause}"
                " GROUP BY stir, company ORDER BY qiymat DESC LIMIT 20",
                conn, params=params,
            )
            df_cnt = pd.read_sql_query(
                "SELECT country, COUNT(DISTINCT decl) AS decl_soni,"
                " SUM(value) AS qiymat, SUM(weight) AS vazn"
                f" FROM active_items WHERE transport='Avia' AND country IS NOT NULL AND country!='' {post_clause}"
                " GROUP BY country ORDER BY qiymat DESC LIMIT 15",
                conn, params=params,
            )
        s = df_sum.iloc[0]
        return {
            "decl_soni": int(s["decl_soni"] or 0),
            "jami_qiymat_k": round(float(s["jami_qiymat"] or 0) / 1000, 1),
            "jami_vazn_tn": round(float(s["jami_vazn"] or 0) / 1000, 3),
            "jami_partiya": int(s["jami_partiya"] or 0),
            "by_month": [
                {
                    "oy": r["oy"],
                    "decl_soni": int(r["decl_soni"] or 0),
                    "qiymat_k": round(float(r["qiymat"] or 0) / 1000, 1),
                    "vazn_tn": round(float(r["vazn"] or 0) / 1000, 3),
                    "partiya": int(r["partiya"] or 0),
                }
                for r in df_oy.to_dict("records")
            ],
            "by_company": [
                {
                    "company": str(r["company"] or ""),
                    "stir": str(r["stir"] or ""),
                    "decl_soni": int(r["decl_soni"] or 0),
                    "qiymat_k": round(float(r["qiymat"] or 0) / 1000, 1),
                    "vazn_tn": round(float(r["vazn"] or 0) / 1000, 3),
                    "partiya": int(r["partiya"] or 0),
                }
                for r in df_comp.to_dict("records")
            ],
            "by_country": [
                {
                    "country": str(r["country"] or ""),
                    "decl_soni": int(r["decl_soni"] or 0),
                    "qiymat_k": round(float(r["qiymat"] or 0) / 1000, 1),
                    "vazn_tn": round(float(r["vazn"] or 0) / 1000, 3),
                }
                for r in df_cnt.to_dict("records")
            ],
        }

    def compute_released(self, base_snapshot_id: int, final_snapshot_id: int) -> dict[str, Any]:
        with self.connect() as conn:
            base = pd.read_sql_query("SELECT * FROM active_items WHERE snapshot_id=?", conn, params=(base_snapshot_id,))
            final = pd.read_sql_query("SELECT * FROM active_items WHERE snapshot_id=?", conn, params=(final_snapshot_id,))
        if base.empty:
            return {"total": {}, "rows": [], "partial": [], "unreleased": []}
        return compute_released_frames(base, final)

    def item_timelines(self) -> pd.DataFrame:
        """Har bir item_key (deklaratsiya + tovar qatori) uchun BUTUN arxiv
        tarixi bo'yicha birinchi va oxirgi ko'rinish sanasini, hamda (agar
        keyinchalik hech qachon qayta ko'rinmagan bo'lsa) taxminiy yechilish
        sanasini hisoblaydi. Faqat bitta sananing (bitta yuklangan fayl)
        holatiga emas, balki hamma arxivlangan fayllarga qarab baholanadi -
        agar biror oraliq fayl to'liqsiz bo'lib item vaqtincha ko'rinmay
        qolsa-yu, undan keyingi faylda yana paydo bo'lsa, bu vaqtinchalik
        uzilish (bitta faylning kamchiligi) deb hisoblanadi, yechilish emas."""
        with self.connect() as conn:
            df = pd.read_sql_query("""
                SELECT ai.item_key, ai.decl, ai.item_no, ai.regime, ai.post, ai.warehouse,
                       ai.stir, ai.company, ai.weight, ai.value, ai.payment, ai.partiya,
                       ai.gtd_date, s.report_date
                FROM active_items ai JOIN snapshots s ON ai.snapshot_id = s.id
            """, conn)
            # Snapshots jadvalidan ALOHIDA olinadi - agar biror sana (masalan
            # hamma narsa yechilib, active_items bo'sh qolgan yuklama) uchun
            # active_items'da mutlaqo qator qolmagan bo'lsa, yuqoridagi JOIN
            # bu sanani butunlay yashirib qo'yardi va "keyingi sana" sifatida
            # hech qachon hisobga olinmasdi.
            all_snap_dates_raw = pd.read_sql_query("SELECT report_date FROM snapshots", conn)["report_date"]
        cols = ["item_key", "decl", "item_no", "regime", "post", "warehouse", "stir", "company",
                "weight", "value", "payment", "partiya", "gtd_date",
                "first_seen_date", "last_seen_date", "release_date", "is_released"]
        all_snap_dates = sorted(pd.to_datetime(all_snap_dates_raw, format="%d.%m.%Y", errors="coerce").dropna().unique())
        if df.empty or not all_snap_dates:
            return pd.DataFrame(columns=cols)
        df["_rd"] = pd.to_datetime(df["report_date"], format="%d.%m.%Y", errors="coerce")
        df = df.dropna(subset=["_rd"])
        if df.empty:
            return pd.DataFrame(columns=cols)
        all_dates = all_snap_dates
        latest_date = all_dates[-1]
        grp = df.groupby("item_key")["_rd"]
        first_seen = grp.min()
        last_seen = grp.max()

        def _release_date(d):
            if d == latest_date:
                return pd.NaT
            later = [x for x in all_dates if x > d]
            return later[0] if later else pd.NaT

        release_dates = last_seen.map(_release_date)
        last_rows = df.sort_values("_rd").groupby("item_key").tail(1).set_index("item_key")
        out = last_rows.drop(columns=["_rd", "report_date"])
        out["first_seen_date"] = first_seen
        out["last_seen_date"] = last_seen
        out["release_date"] = release_dates
        out["is_released"] = out["release_date"].notna()
        return out.reset_index()[cols]

    def items_active_as_of(self, date_text: str, timeline: pd.DataFrame | None = None) -> pd.DataFrame:
        """Berilgan sanaga 'holatiga' haqiqatda faol (nazoratda) bo'lgan
        item'larni butun tarix asosida qaytaradi - yolg'iz shu sanadagi
        (agar mavjud bo'lsa ham) bitta faylga tayanmaydi."""
        df = timeline if timeline is not None else self.item_timelines()
        if df.empty:
            return df
        target = pd.to_datetime(date_text, format="%d.%m.%Y", errors="coerce")
        if pd.isna(target):
            return df.iloc[0:0]
        mask = (df["first_seen_date"] <= target) & (df["release_date"].isna() | (df["release_date"] > target))
        return df[mask].copy()

    def compute_released_robust(self, base_date: str, final_date: str) -> dict[str, Any]:
        """/api/release uchun asosiy hisoblash - ikkita alohida sana snapshot
        emas, item'larning butun tarix bo'yicha faollik oralig'idan foydalanadi."""
        timeline = self.item_timelines()
        base = self.items_active_as_of(base_date, timeline)
        final = self.items_active_as_of(final_date, timeline)
        if base.empty:
            return {"total": {}, "rows": [], "partial": [], "unreleased": []}
        return compute_released_frames(base, final)

    def released_year_bucket_table(self, as_of: "datetime | None" = None) -> dict[str, Any]:
        """Deklaratsiya yili kesimida, bugundan orqaga 1 kun/1 hafta/15 kun/
        1 oy/3 oy/6 oy/1 yil ICHIDA nazoratdan yechilgan tovarlarning
        partiya/vazn/qiymat yig'indisi (kumulyativ - '1 oy' ustuniga '1
        hafta' va '15 kun'dagilar ham kiradi, chunki ular ham 1 oy ichida)."""
        timeline = self.item_timelines()
        empty = {"years": [], "buckets": [[b[0], b[1]] for b in RELEASE_BUCKETS], "rows": [], "total": {}}
        if timeline.empty:
            return empty
        released = timeline[timeline["is_released"]].copy()
        if released.empty:
            return empty
        now = pd.Timestamp(as_of) if as_of is not None else pd.Timestamp.now().normalize()
        released["_decl_year"] = pd.to_datetime(released["gtd_date"], errors="coerce").dt.year
        released = released.dropna(subset=["_decl_year"])
        if released.empty:
            return empty
        released["_decl_year"] = released["_decl_year"].astype(int)
        released["_days_since_release"] = (now - released["release_date"]).dt.days
        years = sorted(released["_decl_year"].unique().tolist())

        def _bucket_sums(frame: pd.DataFrame) -> dict[str, dict[str, float]]:
            out = {}
            for bkey, _blabel, bdays in RELEASE_BUCKETS:
                sub = frame[(frame["_days_since_release"] >= 0) & (frame["_days_since_release"] <= bdays)]
                out[bkey] = {
                    "partiya": int(sub["partiya"].sum()),
                    "vazn": float(sub["weight"].sum()),
                    "qiymat": float(sub["value"].sum()),
                }
            return out

        rows = [{"year": year, **_bucket_sums(released[released["_decl_year"] == year])} for year in years]
        total = {"year": "Jami", **_bucket_sums(released)}
        return {"years": years, "buckets": [[b[0], b[1]] for b in RELEASE_BUCKETS], "rows": rows, "total": total}

    def released_bucket_detail(self, year: int, bucket_key: str, as_of: "datetime | None" = None) -> pd.DataFrame:
        """Berilgan yil + vaqt oralig'i katagiga tegishli aniq deklaratsiya
        qatorlarini qaytaradi (Excel eksporti uchun)."""
        bucket_days = {b[0]: b[2] for b in RELEASE_BUCKETS}
        if bucket_key not in bucket_days:
            return pd.DataFrame()
        timeline = self.item_timelines()
        if timeline.empty:
            return timeline
        released = timeline[timeline["is_released"]].copy()
        if released.empty:
            return released
        now = pd.Timestamp(as_of) if as_of is not None else pd.Timestamp.now().normalize()
        released["_decl_year"] = pd.to_datetime(released["gtd_date"], errors="coerce").dt.year
        released = released.dropna(subset=["_decl_year"])
        if released.empty:
            return released
        released["_decl_year"] = released["_decl_year"].astype(int)
        released["_days_since_release"] = (now - released["release_date"]).dt.days
        mask = (
            (released["_decl_year"] == year)
            & (released["_days_since_release"] >= 0)
            & (released["_days_since_release"] <= bucket_days[bucket_key])
        )
        return released[mask].copy()


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
    # Jami (total) BUTUN base davri bo'yicha hisoblanadi (rows + unreleased),
    # aks holda hali umuman yechilmagan (0% released) kompaniyalar remain_*
    # yig'indisidan tushib qolib, "jami qoldiq" haqiqiy sondan kam chiqadi -
    # frontend releaseCompanyRows()/releaseTotalRow() ham shu ikkalasini
    # birlashtirib hisoblaydi, total shu bilan mos bo'lishi kerak.
    all_rows = rows + unreleased
    total = {
        "korxona": "Jami",
        "stir": "",
        "base_vazn": sum(r["base_vazn"] for r in all_rows),
        "base_qiymat": sum(r["base_qiymat"] for r in all_rows),
        "base_partiya": sum(r["base_partiya"] for r in all_rows),
        "remain_vazn": sum(r["remain_vazn"] for r in all_rows),
        "remain_qiymat": sum(r["remain_qiymat"] for r in all_rows),
        "remain_partiya": sum(r["remain_partiya"] for r in all_rows),
        "released_vazn": sum(r["released_vazn"] for r in all_rows),
        "released_qiymat": sum(r["released_qiymat"] for r in all_rows),
        "released_partiya": sum(r["released_partiya"] for r in all_rows),
        "released_tolov": sum(r["released_tolov"] for r in all_rows),
    }
    total["released_pct"] = (total["released_qiymat"] / total["base_qiymat"] * 100) if total["base_qiymat"] else 0.0
    return {
        "total": total,
        "rows": rows[:500],
        "partial": [r for r in rows if r.get("release_type") == "qisman"][:200],
        "unreleased": sorted(unreleased, key=lambda x: (x["base_qiymat"], x["base_vazn"]), reverse=True)[:200],
        "top_released": rows[:20],
    }
