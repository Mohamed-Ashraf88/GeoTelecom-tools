"""
Telecom LinkMaster Pro
====================================
Version: V4.7.1
Architecture: MVC Pattern

Features:
    - Adaptive Resolution & Centering (Fixed High-DPI logical scaling logic)
    - Proactive UX Data Template & Downloadable Mock Data Generator
    - Fluid Scrollable UI & 8-Point Grid System
    - Elevated Card Design with Modern Typography System
    - Dynamic Path Truncation (Prevents UI stretching)
    - Animated Pulsing States (Tactile Feedback during processing)
    - Native Top Menu Bar (File / Workspaces / Help).
    - Dedicated Scrollable 'About/Documentation' Window
    - Tabbed UI (Bulk File Processing, Single Point Calculator, Live Console)
    - Global Status Bar (Always-visible progress & status across all tabs)
    - Unified Command Bar (Consolidated Header & Pill-shaped Ribbon Actions)
    - Pre-Validation Indicators (✅/❌ with Hover missing column details)
    - Advanced Params: Top-N Targets, Radius Filtering, ML Anomaly Filtering
    - 3D Slant Distance Support (Auto-detects Height/Elevation columns)
    - Multiprocessing Optimization (Shared global arrays for workers)
    - Line-of-Sight (LOS) & Fresnel Zone Clearance Checker (Open-Elevation API)
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Menu
from typing import Dict, List, Optional, Tuple, Any, Union, Callable
from datetime import datetime
import os
import sys
import json
import logging
import logging.handlers
import threading
import multiprocessing as mp
from concurrent.futures import ProcessPoolExecutor
import traceback
import tempfile
import subprocess
import requests
import time
import webbrowser  # Added to open LinkedIn links natively in the browser

# Third-party libraries
import pandas as pd
import numpy as np
from pyproj import Geod
from pydantic import BaseModel, Field
from sklearn.ensemble import IsolationForest
from PIL import Image, ImageDraw, ImageTk

# Dynamic Pydantic Version Handling
try:
    from pydantic import field_validator

    PYDANTIC_V2 = True
except ImportError:
    from pydantic import validator

    PYDANTIC_V2 = False

__title__ = "Telecom LinkMaster Pro"
__version__ = "4.8"
__author__ = "Mohamed Ashraf"

PROGRAM_NAME: str = __title__
VERSION_NUMBER: str = f"V{__version__}"
AUTHOR_NAME: str = __author__
LOG_FILE: str = "app_log.txt"
CONFIG_FILE: str = os.path.join(os.path.expanduser("~"), "GeoDistanceApp_config.json")

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

COLORS = {
    "bg_main": ("#f0f0f0", "#1a1a1a"),
    "bg_card": ("#ffffff", "#2b2b2b"),
    "text_primary": ("#000000", "#ffffff"),
    "text_secondary": ("#555555", "#a0a0a0"),
    "accent": ("#0078d7", "#3298fe"),
    "btn_bg": ("#0078d7", "#1f6aa5"),
    "btn_active": ("#005a9e", "#144870"),
    "separator": ("#e0e0e0", "#333333"),
}

# Typography System
APP_FONT_FAMILY = "San Francisco" if sys.platform == "darwin" else "Segoe UI"
FONTS = {
    "H1": (APP_FONT_FAMILY, 20, "bold"),
    "H2": (APP_FONT_FAMILY, 16, "bold"),
    "Body": (APP_FONT_FAMILY, 13),
    "BodyBold": (APP_FONT_FAMILY, 13, "bold"),
    "Meta": (APP_FONT_FAMILY, 11),
    "Console": ("Consolas", 12),
}


def _truncate_path(path: str, max_chars: int = 45) -> str:
    """Intelligently truncates a file path to prevent layout stretching."""
    if not path or len(path) <= max_chars:
        return path
    return "..." + path[-(max_chars - 3) :]


class ToolTip:
    """Creates a hover tooltip for any CustomTkinter/Tkinter widget."""

    def __init__(self, widget: tk.Widget, text: str):
        self.widget = widget
        self.text = text
        self.tw = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)

    def enter(self, event: Any = None) -> None:
        if self.tw or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + 20

        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry(f"+{x}+{y}")
        self.tw.attributes("-topmost", True)

        is_dark = ctk.get_appearance_mode() == "Dark"
        bg = "#2b2b2b" if is_dark else "#ffffff"
        fg = "#ffffff" if is_dark else "#000000"

        label = tk.Label(
            self.tw,
            text=self.text,
            justify="left",
            background=bg,
            foreground=fg,
            relief="solid",
            borderwidth=1,
            font=FONTS["Meta"],
            padx=8,
            pady=4,
        )
        label.pack()

    def leave(self, event: Any = None) -> None:
        if self.tw:
            self.tw.destroy()
            self.tw = None

    def update_text(self, text: str) -> None:
        self.text = text


class TextboxHandler(logging.Handler):
    """Custom logging handler to post log messages safely to a Tkinter Textbox."""

    def __init__(self, textbox: ctk.CTkTextbox):
        super().__init__()
        self.textbox = textbox
        self.setFormatter(
            logging.Formatter(
                "%(asctime)s - [%(levelname)s] - %(message)s", datefmt="%H:%M:%S"
            )
        )

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.textbox.configure(state="normal")
            self.textbox.insert("end", msg + "\n")
            self.textbox.see("end")
            self.textbox.configure(state="disabled")

        self.textbox.after(0, append)


def setup_logging() -> logging.Logger:
    logger = logging.getLogger("GeoDistanceApp")
    logger.setLevel(logging.INFO)
    log_path = os.path.join(os.path.expanduser("~"), LOG_FILE)
    fh = logging.handlers.RotatingFileHandler(
        log_path, encoding="utf-8", maxBytes=5 * 1024 * 1024, backupCount=3
    )
    fh.setFormatter(
        logging.Formatter(
            "%(asctime)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
        )
    )
    if not logger.handlers:
        logger.addHandler(fh)
    return logger


logger = setup_logging()


def open_file_auto(path: str) -> bool:
    """Safely auto-opens a file using the OS default application."""
    if not os.path.exists(path):
        return False
    try:
        if os.name == "nt":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.call(["open", path])
        else:
            subprocess.call(["xdg-open", path])
        return True
    except Exception as e:
        logger.error(f"Failed to auto-open {path}: {e}")
        return False


class WorkspaceConfig(BaseModel):
    scenario: str = Field(default="Allow 0m (Identical Locations)")
    calculation_method: str = Field(default="Fast (Haversine)")
    top_n_targets: int = Field(default=1, ge=1, le=10)
    max_radius_m: float = Field(default=0.0)
    chunk_size: int = Field(default=500)
    enable_anomaly_detection: bool = Field(default=True)
    anomaly_contamination: float = Field(default=0.05)
    enable_los_check: bool = Field(default=False)


class AppRootConfig(BaseModel):
    active_workspace: str = Field(default="Default")
    workspaces: Dict[str, WorkspaceConfig] = Field(
        default_factory=lambda: {"Default": WorkspaceConfig()}
    )
    ui_theme: str = Field(default="System")
    ui_scaling: str = Field(default="100%")
    last_export_dir: str = Field(default="")

    def to_dict(self) -> Dict[str, Any]:
        return self.model_dump() if PYDANTIC_V2 else self.dict()

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AppRootConfig":
        try:
            return cls(**data)
        except Exception:
            return cls()


class ConfigManager:
    @staticmethod
    def load() -> AppRootConfig:
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    data = json.load(f)
                    if "workspaces" not in data:
                        return AppRootConfig()
                    return AppRootConfig.from_dict(data)
            except Exception:
                pass
        return AppRootConfig()

    @staticmethod
    def save(config: AppRootConfig) -> None:
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(config.to_dict(), f, indent=4)
        except Exception as e:
            logger.error(f"Failed to save config: {e}")


class ColumnMapper:
    STANDARD_COLUMNS = {
        "Point name": [
            "point name",
            "pointname",
            "site_name",
            "sitename",
            "site",
            "id",
            "name",
            "station",
            "cell",
            "node",
        ],
        "Latitude": ["latitude", "lat", "y", "northing", "lat_deg"],
        "Longitude": ["longitude", "lon", "long", "lng", "x", "easting", "lon_deg"],
        "Height": [
            "height",
            "elevation",
            "alt",
            "altitude",
            "z",
            "antenna_height",
            "ant_height",
        ],
    }

    @staticmethod
    def rename(df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
        mapping: Dict[str, str] = {}
        renamed = 0
        for col in df.columns:
            norm = str(col).strip().lower().replace("_", " ").replace("-", " ")
            for std, vars in ColumnMapper.STANDARD_COLUMNS.items():
                if (
                    any(v in norm for v in vars)
                    and col != std
                    and std not in df.columns
                ):
                    mapping[col] = std
                    renamed += 1
                    break
        if mapping:
            df = df.rename(columns=mapping)
        return df, renamed


_T_LATS: np.ndarray = np.array([])
_T_LONS: np.ndarray = np.array([])
_T_HEIGHTS: np.ndarray = np.array([])
_T_NAMES: np.ndarray = np.array([])


def _init_worker(
    t_lats: np.ndarray, t_lons: np.ndarray, t_heights: np.ndarray, t_names: np.ndarray
) -> None:
    """Initializes global arrays inside each worker process for memory efficiency."""
    global _T_LATS, _T_LONS, _T_HEIGHTS, _T_NAMES
    _T_LATS = t_lats
    _T_LONS = t_lons
    _T_HEIGHTS = t_heights
    _T_NAMES = t_names


class DistanceCalculator:
    EARTH_RADIUS = 6371000.0

    @staticmethod
    def parse_file(
        path: str, crs_mode: str, config: WorkspaceConfig
    ) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
        ext = os.path.splitext(path)[1].lower()
        logger.info(f"Parsing file: {os.path.basename(path)}")

        try:
            if ext in [".xlsx", ".xls"]:
                original_df = pd.read_excel(
                    path, engine="openpyxl" if ext == ".xlsx" else None
                )
            elif ext in [".csv", ".txt"]:
                original_df = pd.read_csv(path, sep=None, engine="python")
            elif ext == ".geojson":
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                records = []
                for feat in data.get("features", []):
                    geom = feat.get("geometry", {})
                    props = feat.get("properties", {})
                    if geom and geom.get("type") == "Point":
                        coords = geom.get("coordinates", [None, None])
                        if len(coords) >= 2:
                            props["Longitude"] = coords[0]
                            props["Latitude"] = coords[1]
                    records.append(props)
                original_df = pd.DataFrame(records)
            else:
                raise ValueError(f"Unsupported file format: {ext}")
        except Exception as e:
            logger.error(f"File Parsing Failed: {str(e)}")
            raise ValueError(f"Failed to read file {os.path.basename(path)}: {str(e)}")

        df, renamed = ColumnMapper.rename(original_df)
        if renamed > 0:
            logger.info(f"Auto-renamed {renamed} column(s) to standard headers.")

        req_cols = ["Point name", "Latitude", "Longitude"]
        if not all(c in df.columns for c in req_cols):
            missing = [c for c in req_cols if c not in df.columns]
            logger.error(f"Missing required columns: {missing}")
            raise ValueError(f"Missing required columns in {os.path.basename(path)}")

        df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
        df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
        if "Height" in df.columns:
            df["Height"] = pd.to_numeric(df["Height"], errors="coerce").fillna(0)

        # Basic validity checking
        valid_mask = (
            df["Latitude"].notna()
            & df["Longitude"].notna()
            & (df["Latitude"].between(-90, 90))
            & (df["Longitude"].between(-180, 180))
        )

        dropped_df = original_df[~valid_mask].copy()
        if not dropped_df.empty:
            dropped_df["Drop_Reason"] = "Invalid or Missing Coordinates"

        df = df[valid_mask].copy()

        # Deduplication
        dups = df.duplicated(subset=["Point name"], keep="first")
        if dups.any():
            dup_df = original_df.loc[df[dups].index].copy()
            dup_df["Drop_Reason"] = "Duplicate Point Name"
            dropped_df = pd.concat([dropped_df, dup_df])
            df = df[~dups]

        dropped_invalid_count = len(dropped_df)
        dropped_anomalies_count = 0

        # ML Anomaly Detection Pipeline
        if config.enable_anomaly_detection and len(df) > 10:
            logger.info(
                f"Running Isolation Forest ML on {len(df)} coordinate points..."
            )
            clf = IsolationForest(
                contamination=config.anomaly_contamination, random_state=42
            )
            preds = clf.fit_predict(df[["Latitude", "Longitude"]].values)
            anomaly_mask = preds == -1

            if anomaly_mask.any():
                anom_df = original_df.loc[df[anomaly_mask].index].copy()
                anom_df["Drop_Reason"] = "Statistical Anomaly (ML)"
                dropped_df = pd.concat([dropped_df, anom_df])
                df = df[~anomaly_mask]
                dropped_anomalies_count = anomaly_mask.sum()
                logger.warning(
                    f"ML Anomaly filter isolated and dropped {dropped_anomalies_count} outlier points."
                )

        stats = {
            "total": len(original_df),
            "valid": len(df),
            "dropped_invalid": dropped_invalid_count,
            "dropped_anomalies": dropped_anomalies_count,
            "renamed": renamed,
            "has_height": "Height" in df.columns,
        }
        logger.info(
            f"Loaded '{os.path.basename(path)}' - Valid: {stats['valid']}, Dropped: {len(dropped_df)}"
        )
        return df, dropped_df, stats

    @staticmethod
    def calc_3d_slant(dist_2d: np.ndarray, h1: float, h2_arr: np.ndarray) -> np.ndarray:
        return np.sqrt(dist_2d**2 + (h2_arr - h1) ** 2)

    @staticmethod
    def calculate_geodesic_vectorized(
        slat: float, slon: float, tlats: np.ndarray, tlons: np.ndarray
    ) -> Tuple[np.ndarray, np.ndarray]:
        geod = Geod(ellps="WGS84")
        fwd, _, d2d = geod.inv(
            np.full(len(tlons), slon), np.full(len(tlats), slat), tlons, tlats
        )
        bearings = (fwd + 360) % 360
        return d2d, bearings

    @staticmethod
    def check_los_clearance(
        lon1: float,
        lat1: float,
        lon2: float,
        lat2: float,
        h1_agl: float,
        h2_agl: float,
        dist_m: float,
    ) -> str:
        """Hits Open-Elevation API, calculates 60% Fresnel Zone @ 5GHz, and factors in Earth Bulge (K=4/3)"""
        if dist_m < 10:
            return "CLEAR (Too Close)"

        num_points = max(
            3, min(int(dist_m / 10.0), 500)
        )  # Capped at 500 to prevent API payload explosion
        geod = Geod(ellps="WGS84")

        inner_pts = geod.npts(lon1, lat1, lon2, lat2, num_points - 2)
        all_pts = [(lon1, lat1)] + inner_pts + [(lon2, lat2)]

        payload = {
            "locations": [{"latitude": p[1], "longitude": p[0]} for p in all_pts]
        }

        try:
            resp = requests.post(
                "https://api.open-elevation.com/api/v1/lookup", json=payload, timeout=15
            )
            resp.raise_for_status()
            elevations = [loc["elevation"] for loc in resp.json()["results"]]
        except Exception:
            return "API Error / Timeout"

        if len(elevations) != len(all_pts):
            return "Data Mismatch"

        terr_start = elevations[0]
        terr_end = elevations[-1]

        abs_h1 = terr_start + h1_agl
        abs_h2 = terr_end + h2_agl
        f_ghz = 5.0

        for i, (lon, lat) in enumerate(all_pts):
            if i == 0 or i == len(all_pts) - 1:
                continue

            _, _, d1 = geod.inv(lon1, lat1, lon, lat)
            d2 = dist_m - d1
            if d1 <= 0 or d2 <= 0:
                continue

            f1 = 17.32 * np.sqrt(
                (d1 / 1000.0) * (d2 / 1000.0) / (f_ghz * (dist_m / 1000.0))
            )
            req_clearance = 0.6 * f1

            bulge = (d1 / 1000.0) * (d2 / 1000.0) / 17.0
            los_alt = abs_h1 + (d1 / dist_m) * (abs_h2 - abs_h1)
            obstacle_h = elevations[i] + bulge

            if (los_alt - obstacle_h) < req_clearance:
                return "BLOCKED"

        return "CLEAR"

    @staticmethod
    def process_chunk(
        args: Tuple[pd.DataFrame, Dict[str, Any]],
    ) -> List[Dict[str, Any]]:
        source_chunk, cfg = args
        global _T_LATS, _T_LONS, _T_HEIGHTS, _T_NAMES

        results: List[Dict[str, Any]] = []
        geod = Geod(ellps="WGS84")

        # Refactored iterrows to itertuples for massive performance gain
        for row in source_chunk.itertuples():
            s_lat = float(row.Latitude)
            s_lon = float(row.Longitude)

            # Using getattr to safely fetch optional columns handled by pandas namedtuples
            s_name = str(getattr(row, "_1", getattr(row, "Point_name", "Unknown")))

            s_h = 0.0
            if hasattr(row, "Height"):
                s_h = float(row.Height)

            if cfg["method"] == "High Precision (Geodesic)":
                fwd, _, d2d = geod.inv(
                    np.full(len(_T_LONS), s_lon),
                    np.full(len(_T_LATS), s_lat),
                    _T_LONS,
                    _T_LATS,
                )
                bearings = (fwd + 360) % 360
            else:
                l1, l2 = np.radians(s_lat), np.radians(_T_LATS)
                dl = l2 - l1
                dlo = np.radians(_T_LONS) - np.radians(s_lon)
                a = np.clip(
                    np.sin(dl / 2) ** 2
                    + np.cos(l1) * np.cos(l2) * np.sin(dlo / 2) ** 2,
                    0,
                    1,
                )
                d2d = (
                    DistanceCalculator.EARTH_RADIUS
                    * 2
                    * np.arctan2(np.sqrt(a), np.sqrt(1 - a))
                )
                y = np.sin(dlo) * np.cos(l2)
                x = np.cos(l1) * np.sin(l2) - np.sin(l1) * np.cos(l2) * np.cos(dlo)
                bearings = (np.degrees(np.arctan2(y, x)) + 360) % 360

            dists = (
                DistanceCalculator.calc_3d_slant(d2d, s_h, _T_HEIGHTS)
                if cfg["use_3d"]
                else d2d
            )

            if cfg["ignore_0"]:
                valid = dists > 0.01
                dists = np.where(valid, dists, np.inf)

            top_n = min(cfg["top_n"], len(dists))
            sorted_idx = np.argsort(dists)[:top_n]

            for rank in range(top_n):
                idx = sorted_idx[rank]
                dist_val = dists[idx]
                if dist_val == np.inf or (
                    cfg["max_rad"] > 0 and dist_val > cfg["max_rad"]
                ):
                    continue

                los_status = "N/A"
                if cfg.get("enable_los") and rank == 0:
                    los_status = DistanceCalculator.check_los_clearance(
                        s_lon,
                        s_lat,
                        _T_LONS[idx],
                        _T_LATS[idx],
                        s_h,
                        _T_HEIGHTS[idx],
                        dist_val,
                    )

                results.append(
                    {
                        "Source Point": s_name,
                        "Target Point": _T_NAMES[idx],
                        "Rank": rank + 1,
                        "Distance (m)": float(dist_val),
                        "Bearing (deg)": float(bearings[idx]),
                        "Calculation Type": (
                            "3D Slant" if cfg["use_3d"] else "2D Ground"
                        ),
                        "LOS Status (Top 1)": los_status,
                    }
                )
        return results


class SanityCheckPopup(ctk.CTkToplevel):
    def __init__(
        self,
        parent: ctk.CTk,
        s_stats: Dict[str, Any],
        t_stats: Dict[str, Any],
        on_confirm: Callable[[], None],
        on_cancel: Callable[[], None],
    ) -> None:
        super().__init__(parent)
        self.title("Data Sanity Check")
        self.geometry("450x450")
        self.grab_set()

        txt = f"""
        📊 SOURCE DATA:
        - Valid Points: {s_stats['valid']}
        - Dropped (Invalid/Dup): {s_stats['dropped_invalid']}
        - Dropped (ML Anomalies): {s_stats['dropped_anomalies']}
        - 3D Height Data: {'Yes' if s_stats['has_height'] else 'No'}

        🎯 TARGET DATA:
        - Valid Points: {t_stats['valid']}
        - Dropped (Invalid/Dup): {t_stats['dropped_invalid']}
        - Dropped (ML Anomalies): {t_stats['dropped_anomalies']}
        - 3D Height Data: {'Yes' if t_stats['has_height'] else 'No'}
        """

        if (
            s_stats["dropped_invalid"] > 0
            or t_stats["dropped_invalid"] > 0
            or s_stats["dropped_anomalies"] > 0
            or t_stats["dropped_anomalies"] > 0
        ):
            txt += "\n⚠️ Notice: All dropped data will be exported to a report."

        ctk.CTkLabel(
            self,
            text="Ready to process with the following data:",
            font=FONTS["H2"],
        ).pack(pady=(24, 8))
        ctk.CTkLabel(self, text=txt, justify="left", font=FONTS["Body"]).pack(
            pady=8, padx=24, fill="x"
        )

        btn_frm = ctk.CTkFrame(self, fg_color="transparent")
        btn_frm.pack(pady=24)

        def cancel_action():
            self.destroy()
            on_cancel()

        ctk.CTkButton(
            btn_frm,
            text="Cancel",
            fg_color="gray",
            font=FONTS["BodyBold"],
            command=cancel_action,
        ).pack(side="left", padx=16)

        def confirm_action() -> None:
            self.destroy()
            on_confirm()

        ctk.CTkButton(
            btn_frm,
            text="Proceed Calculation",
            fg_color=COLORS["accent"],
            font=FONTS["BodyBold"],
            command=confirm_action,
        ).pack(side="right", padx=16)


class GeoDistanceApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(PROGRAM_NAME)

        # Fixed Logical Scaling UI Resolution System
        self._center_and_scale_window()
        self.minsize(850, 600)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.root_config = ConfigManager.load()
        if self.root_config.active_workspace not in self.root_config.workspaces:
            self.root_config.active_workspace = "Default"
            self.root_config.workspaces["Default"] = WorkspaceConfig()

        ctk.set_appearance_mode(self.root_config.ui_theme)
        try:
            scale_val = float(self.root_config.ui_scaling.strip("%")) / 100.0
            ctk.set_widget_scaling(scale_val)
            ctk.set_window_scaling(scale_val)
        except Exception:
            pass

        self.s_df: pd.DataFrame = pd.DataFrame()
        self.t_df: pd.DataFrame = pd.DataFrame()
        self.s_drop: pd.DataFrame = pd.DataFrame()
        self.t_drop: pd.DataFrame = pd.DataFrame()
        self.src_path: str = ""
        self.tgt_path: str = ""
        self.advanced_visible: bool = False

        # Thread safe events for processing
        self.abort_event = threading.Event()
        self.is_processing = False

        self.temp_ico_path: Optional[str] = None
        self._build_ui()
        self._set_dynamic_icon()

        self._apply_workspace_to_ui(
            self.root_config.workspaces[self.root_config.active_workspace]
        )

        logger.info(
            f"Application started. Active Workspace: {self.root_config.active_workspace}"
        )

    def _center_and_scale_window(self) -> None:
        """Calculates dimensions logically to prevent High-DPI double-scaling explosion."""
        self.update_idletasks()

        # Logical base dimensions. CustomTkinter will auto-scale these for 4K/High-DPI.
        w = 1050
        h = 700

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # Prevent overflow on extremely small screens (sub-1080p)
        if screen_width > 0 and screen_height > 0:
            w = min(w, int(screen_width * 0.95))
            h = min(h, int(screen_height * 0.95))

        # Enforce minimum boundaries
        w = max(w, 850)
        h = max(h, 600)

        x = int((screen_width - w) / 2)
        y = int((screen_height - h) / 2)

        self.geometry(f"{w}x{h}+{x}+{y}")

    def _set_dynamic_icon(self) -> None:
        try:
            icon_img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
            draw = ImageDraw.Draw(icon_img)
            accent_color = (
                COLORS["accent"][0]
                if isinstance(COLORS["accent"], tuple)
                else COLORS["accent"]
            )

            draw.rounded_rectangle([(4, 4), (60, 60)], radius=12, fill=accent_color)
            draw.line([(32, 50), (32, 24)], fill="white", width=4)
            draw.ellipse([(28, 20), (36, 28)], fill="white")
            draw.arc([(18, 10), (46, 38)], start=210, end=330, fill="white", width=3)
            draw.arc([(10, 2), (54, 46)], start=210, end=330, fill="white", width=3)

            icon_photo = ImageTk.PhotoImage(icon_img)
            self.iconphoto(True, icon_photo)

            if os.name == "nt":
                fd, self.temp_ico_path = tempfile.mkstemp(suffix=".ico")
                os.close(fd)
                icon_img.save(self.temp_ico_path, format="ICO", sizes=[(64, 64)])
                self.iconbitmap(self.temp_ico_path)
        except Exception as e:
            logger.error(f"Failed to generate window icon: {e}")

    def _build_ui(self) -> None:
        menubar = Menu(self)

        # File Menu
        file_menu = Menu(menubar, tearoff=0)
        file_menu.add_command(
            label="Import Source", command=lambda: self._browse("src")
        )
        file_menu.add_command(
            label="Import Target", command=lambda: self._browse("tgt")
        )
        file_menu.add_separator()
        file_menu.add_command(label="Open Log File", command=self._open_log_file)
        file_menu.add_command(
            label="Open Config Folder", command=self._open_config_folder
        )
        file_menu.add_command(
            label="Open Last Export Directory", command=self._open_last_export_dir
        )
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_close)

        # Workspace Menu
        self.workspace_menu = Menu(menubar, tearoff=0)
        self._update_workspace_menu()

        # View / Appearance Menu
        view_menu = Menu(menubar, tearoff=0)

        theme_menu = Menu(view_menu, tearoff=0)
        theme_menu.add_command(
            label="System Default", command=lambda: self._set_theme("System")
        )
        theme_menu.add_command(label="Light", command=lambda: self._set_theme("Light"))
        theme_menu.add_command(label="Dark", command=lambda: self._set_theme("Dark"))

        scale_menu = Menu(view_menu, tearoff=0)
        scale_menu.add_command(label="100%", command=lambda: self._set_scaling("100%"))
        scale_menu.add_command(label="125%", command=lambda: self._set_scaling("125%"))
        scale_menu.add_command(label="150%", command=lambda: self._set_scaling("150%"))

        view_menu.add_cascade(label="Theme", menu=theme_menu)
        view_menu.add_cascade(label="Scaling", menu=scale_menu)

        # Help Menu
        help_menu = Menu(menubar, tearoff=0)
        help_menu.add_command(label="About the Tool", command=self._show_about)
        help_menu.add_command(label="Contact Us", command=self._show_contact)

        menubar.add_cascade(label="File", menu=file_menu)
        menubar.add_cascade(label="View", menu=view_menu)
        menubar.add_cascade(label="Workspaces", menu=self.workspace_menu)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.configure(menu=menubar)

        # Global Status Bar
        self.status_frame = ctk.CTkFrame(self, height=32, corner_radius=0)
        self.status_frame.pack(side="bottom", fill="x")

        self.lbl_stat = ctk.CTkLabel(
            self.status_frame, text="Ready.", width=250, anchor="w", font=FONTS["Body"]
        )
        self.lbl_stat.pack(side="left", padx=(16, 8), pady=8)

        self.prog = ctk.CTkProgressBar(self.status_frame)
        self.prog.pack(side="left", fill="x", expand=True, padx=(8, 24), pady=8)
        self.prog.set(0)

        # Unified Command Bar Header
        hdr = ctk.CTkFrame(
            self,
            fg_color=COLORS["bg_card"],
            height=56,
            corner_radius=12,
            border_width=1,
            border_color=COLORS["separator"],
        )
        hdr.pack(fill="x", padx=16, pady=(16, 8))

        frm_title = ctk.CTkFrame(hdr, fg_color="transparent")
        frm_title.pack(side="left", padx=16, pady=8)
        ctk.CTkLabel(frm_title, text=PROGRAM_NAME, font=FONTS["H1"]).pack(anchor="w")
        ctk.CTkLabel(
            frm_title,
            text=f"{VERSION_NUMBER} | Author: {AUTHOR_NAME}",
            font=FONTS["Meta"],
            text_color=COLORS["text_secondary"],
        ).pack(anchor="w")

        frm_actions = ctk.CTkFrame(hdr, fg_color="transparent")
        frm_actions.pack(side="right", padx=16, pady=8)

        btn_template = ctk.CTkButton(
            frm_actions,
            text="📋 Data Template",
            width=120,
            corner_radius=20,
            command=self._show_data_template,
            font=FONTS["BodyBold"],
            fg_color="#107C41",
            hover_color="#0b5c30",
        )
        btn_template.pack(side="left", padx=(0, 8))

        btn_save_ws = ctk.CTkButton(
            frm_actions,
            text="💾 Save Workspace",
            width=120,
            corner_radius=20,
            command=self._save_workspace_as,
            font=FONTS["BodyBold"],
        )
        btn_save_ws.pack(side="left", padx=8)

        btn_clear = ctk.CTkButton(
            frm_actions,
            text="🗑️ Clear Console",
            width=120,
            corner_radius=20,
            command=self._console_clear,
            fg_color=COLORS["separator"],
            text_color=COLORS["text_primary"],
            font=FONTS["BodyBold"],
        )
        btn_clear.pack(side="left", padx=(8, 0))

        # Core Tabs Content Area
        self.tabs = ctk.CTkTabview(self, fg_color="transparent")
        self.tabs.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.tab_bulk = self.tabs.add("Bulk File Processing")
        self.tab_single = self.tabs.add("Single Point Calculator")
        self.tab_console = self.tabs.add("Console / History")

        self._build_bulk_tab()
        self._build_single_tab()
        self._build_console_tab()

    def _build_bulk_tab(self) -> None:
        scroll_container = ctk.CTkScrollableFrame(self.tab_bulk, fg_color="transparent")
        scroll_container.pack(fill="both", expand=True)

        frm_files = ctk.CTkFrame(
            scroll_container,
            fg_color=COLORS["bg_card"],
            corner_radius=12,
            border_width=1,
            border_color=COLORS["separator"],
        )
        frm_files.pack(fill="x", pady=(8, 16), padx=8)

        # Source Row
        ctk.CTkLabel(frm_files, text="📂 Source File:", font=FONTS["BodyBold"]).grid(
            row=0, column=0, padx=(16, 8), pady=16, sticky="w"
        )
        self.entry_src = ctk.CTkEntry(
            frm_files,
            state="readonly",
            placeholder_text="No file selected...",
            font=FONTS["Body"],
        )
        self.entry_src.grid(row=0, column=1, padx=8, pady=16, sticky="ew")

        self.lbl_src_status = ctk.CTkLabel(frm_files, text="", width=24)
        self.lbl_src_status.grid(row=0, column=2, padx=(8, 8))
        self.src_tooltip = ToolTip(self.lbl_src_status, "")

        ctk.CTkButton(
            frm_files,
            text="Browse Source",
            command=lambda: self._browse("src"),
            width=100,
            font=FONTS["BodyBold"],
        ).grid(row=0, column=3, padx=(8, 16), pady=16, sticky="e")

        # Target Row
        ctk.CTkLabel(frm_files, text="🎯 Target File:", font=FONTS["BodyBold"]).grid(
            row=1, column=0, padx=(16, 8), pady=(0, 16), sticky="w"
        )
        self.entry_tgt = ctk.CTkEntry(
            frm_files,
            state="readonly",
            placeholder_text="No file selected...",
            font=FONTS["Body"],
        )
        self.entry_tgt.grid(row=1, column=1, padx=8, pady=(0, 16), sticky="ew")

        self.lbl_tgt_status = ctk.CTkLabel(frm_files, text="", width=24)
        self.lbl_tgt_status.grid(row=1, column=2, padx=(8, 8))
        self.tgt_tooltip = ToolTip(self.lbl_tgt_status, "")

        ctk.CTkButton(
            frm_files,
            text="Browse Target",
            command=lambda: self._browse("tgt"),
            width=100,
            font=FONTS["BodyBold"],
        ).grid(row=1, column=3, padx=(8, 16), pady=(0, 16), sticky="e")

        frm_files.grid_columnconfigure(1, weight=1)

        # Core Settings
        frm_set = ctk.CTkFrame(scroll_container, fg_color="transparent")
        frm_set.pack(fill="x", pady=8, padx=8)

        self.var_scen = ctk.StringVar()
        ctk.CTkComboBox(
            frm_set,
            variable=self.var_scen,
            font=FONTS["Body"],
            dropdown_font=FONTS["Body"],
            values=["Allow 0m (Identical Locations)", "Ignore 0m (Find Next Closest)"],
        ).grid(row=0, column=0, padx=(0, 8), pady=8, sticky="ew")

        self.var_calc = ctk.StringVar()
        ctk.CTkComboBox(
            frm_set,
            variable=self.var_calc,
            font=FONTS["Body"],
            dropdown_font=FONTS["Body"],
            values=["Fast (Haversine)", "High Precision (Geodesic)"],
        ).grid(row=0, column=1, padx=(8, 0), pady=8, sticky="ew")
        frm_set.grid_columnconfigure((0, 1), weight=1)

        # Advanced Settings Accordion
        self.btn_adv_toggle = ctk.CTkButton(
            frm_set,
            text="⚙️ Advanced Settings ▼",
            fg_color="transparent",
            font=FONTS["BodyBold"],
            text_color=COLORS["text_secondary"],
            hover_color=COLORS["separator"],
            command=self._toggle_advanced,
        )
        self.btn_adv_toggle.grid(
            row=1, column=0, columnspan=2, sticky="w", padx=0, pady=(16, 8)
        )

        self.frm_adv_container = ctk.CTkFrame(
            frm_set,
            fg_color=COLORS["bg_card"],
            corner_radius=12,
            border_width=1,
            border_color=COLORS["separator"],
        )

        frm_adv = ctk.CTkFrame(self.frm_adv_container, fg_color="transparent")
        frm_adv.pack(fill="x", pady=(16, 8), padx=16)

        ctk.CTkLabel(frm_adv, text="Top N Targets:", font=FONTS["BodyBold"]).pack(
            side="left", padx=(0, 8)
        )
        self.var_topn = ctk.IntVar()

        spinbox_frame = ctk.CTkFrame(frm_adv, fg_color="transparent")
        spinbox_frame.pack(side="left", padx=(0, 24))

        def dec_topn():
            if self.var_topn.get() > 1:
                self.var_topn.set(self.var_topn.get() - 1)

        def inc_topn():
            if self.var_topn.get() < 10:
                self.var_topn.set(self.var_topn.get() + 1)

        ctk.CTkButton(
            spinbox_frame, text="-", width=32, font=FONTS["BodyBold"], command=dec_topn
        ).pack(side="left", padx=(0, 4))
        self.entry_topn = ctk.CTkEntry(
            spinbox_frame,
            width=48,
            textvariable=self.var_topn,
            justify="center",
            state="readonly",
            font=FONTS["BodyBold"],
        )
        self.entry_topn.pack(side="left")
        ctk.CTkButton(
            spinbox_frame, text="+", width=32, font=FONTS["BodyBold"], command=inc_topn
        ).pack(side="left", padx=(4, 0))

        ctk.CTkLabel(
            frm_adv, text="Max Radius (m, 0=Any):", font=FONTS["BodyBold"]
        ).pack(side="left")
        self.var_rad = ctk.DoubleVar()
        ctk.CTkEntry(
            frm_adv, textvariable=self.var_rad, width=80, font=FONTS["Body"]
        ).pack(side="left", padx=(8, 24))

        frm_cb = ctk.CTkFrame(self.frm_adv_container, fg_color="transparent")
        frm_cb.pack(fill="x", pady=(0, 16), padx=16)

        self.var_anomaly = ctk.BooleanVar()
        ctk.CTkCheckBox(
            frm_cb,
            text="ML Anomaly Filter",
            variable=self.var_anomaly,
            font=FONTS["Body"],
        ).pack(side="left", padx=(0, 4))
        lbl_anom_tip = ctk.CTkLabel(
            frm_cb,
            text="(?)",
            text_color=COLORS["text_secondary"],
            cursor="hand2",
            font=FONTS["Meta"],
        )
        lbl_anom_tip.pack(side="left", padx=(0, 32))
        ToolTip(
            lbl_anom_tip,
            "Uses Isolation Forest Machine Learning to automatically\nidentify and drop coordinate pairs that are statistical outliers.",
        )

        self.var_los = ctk.BooleanVar()
        ctk.CTkCheckBox(
            frm_cb,
            text="Enable 5GHz LOS Terrain Check (Top 1)",
            variable=self.var_los,
            font=FONTS["Body"],
        ).pack(side="left", padx=(0, 4))
        lbl_los_tip = ctk.CTkLabel(
            frm_cb,
            text="(?)",
            text_color=COLORS["text_secondary"],
            cursor="hand2",
            font=FONTS["Meta"],
        )
        lbl_los_tip.pack(side="left")
        ToolTip(
            lbl_los_tip,
            "Queries Open-Elevation API for line-of-sight terrain.\nChecks 60% of the 1st Fresnel Zone at 5GHz with K=4/3 Earth Curvature.\nNote: Only evaluates #1 closest target.",
        )

        # Action Buttons Area
        frm_actions = ctk.CTkFrame(scroll_container, fg_color="transparent")
        frm_actions.pack(fill="x", pady=24, padx=8)
        frm_actions.grid_columnconfigure((0, 1), weight=1)

        self.btn_proc = ctk.CTkButton(
            frm_actions,
            text="Validate & Process",
            font=FONTS["H2"],
            height=48,
            corner_radius=12,
            command=self._start_phase1,
        )
        self.btn_proc.grid(row=0, column=0, padx=(0, 8), sticky="ew")

        self.btn_abort = ctk.CTkButton(
            frm_actions,
            text="🛑 Abort Process",
            font=FONTS["H2"],
            height=48,
            corner_radius=12,
            fg_color="#d13438",
            hover_color="#a80000",
            state="disabled",
            command=self._abort_process,
        )
        self.btn_abort.grid(row=0, column=1, padx=(8, 0), sticky="ew")

    def _build_single_tab(self) -> None:
        scroll_container = ctk.CTkScrollableFrame(
            self.tab_single, fg_color="transparent"
        )
        scroll_container.pack(fill="both", expand=True)

        frm_cards = ctk.CTkFrame(scroll_container, fg_color="transparent")
        frm_cards.pack(fill="x", pady=16)
        frm_cards.grid_columnconfigure((0, 1), weight=1)

        card_src = ctk.CTkFrame(
            frm_cards,
            fg_color=COLORS["bg_card"],
            corner_radius=12,
            border_width=1,
            border_color=COLORS["separator"],
        )
        card_src.grid(row=0, column=0, padx=(0, 16), sticky="ew")
        ctk.CTkLabel(card_src, text="📡 Source Point", font=FONTS["H2"]).pack(
            pady=(24, 16)
        )

        frm_src_in = ctk.CTkFrame(card_src, fg_color="transparent")
        frm_src_in.pack(pady=(0, 24), padx=24)
        ctk.CTkLabel(frm_src_in, text="📍 Latitude:", font=FONTS["BodyBold"]).grid(
            row=0, column=0, padx=8, pady=8, sticky="e"
        )
        self.e_slat = ctk.CTkEntry(
            frm_src_in, placeholder_text="e.g. 30.1234", font=FONTS["Body"]
        )
        self.e_slat.grid(row=0, column=1, padx=8, pady=8, sticky="ew")
        ctk.CTkLabel(frm_src_in, text="📍 Longitude:", font=FONTS["BodyBold"]).grid(
            row=1, column=0, padx=8, pady=8, sticky="e"
        )
        self.e_slon = ctk.CTkEntry(
            frm_src_in, placeholder_text="e.g. -90.5678", font=FONTS["Body"]
        )
        self.e_slon.grid(row=1, column=1, padx=8, pady=8, sticky="ew")
        frm_src_in.grid_columnconfigure(1, weight=1)

        card_tgt = ctk.CTkFrame(
            frm_cards,
            fg_color=COLORS["bg_card"],
            corner_radius=12,
            border_width=1,
            border_color=COLORS["separator"],
        )
        card_tgt.grid(row=0, column=1, padx=(16, 0), sticky="ew")
        ctk.CTkLabel(card_tgt, text="🎯 Target Point", font=FONTS["H2"]).pack(
            pady=(24, 16)
        )

        frm_tgt_in = ctk.CTkFrame(card_tgt, fg_color="transparent")
        frm_tgt_in.pack(pady=(0, 24), padx=24)
        ctk.CTkLabel(frm_tgt_in, text="📍 Latitude:", font=FONTS["BodyBold"]).grid(
            row=0, column=0, padx=8, pady=8, sticky="e"
        )
        self.e_tlat = ctk.CTkEntry(
            frm_tgt_in, placeholder_text="e.g. 30.1245", font=FONTS["Body"]
        )
        self.e_tlat.grid(row=0, column=1, padx=8, pady=8, sticky="ew")
        ctk.CTkLabel(frm_tgt_in, text="📍 Longitude:", font=FONTS["BodyBold"]).grid(
            row=1, column=0, padx=8, pady=8, sticky="e"
        )
        self.e_tlon = ctk.CTkEntry(
            frm_tgt_in, placeholder_text="e.g. -90.5689", font=FONTS["Body"]
        )
        self.e_tlon.grid(row=1, column=1, padx=8, pady=8, sticky="ew")
        frm_tgt_in.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            scroll_container,
            text="Calculate Distance & Bearing",
            font=FONTS["H2"],
            height=48,
            corner_radius=12,
            command=self._calc_single,
        ).pack(pady=(16, 32))

        self.res_panel = ctk.CTkFrame(
            scroll_container,
            fg_color=COLORS["bg_card"],
            corner_radius=12,
            border_width=1,
            border_color=COLORS["separator"],
        )
        self.res_panel.pack(fill="x", padx=32)
        self.lbl_sres = ctk.CTkLabel(
            self.res_panel,
            text="Enter coordinates and click Calculate.",
            font=FONTS["Body"],
            text_color=COLORS["text_secondary"],
        )
        self.lbl_sres.pack(pady=32)

    def _build_console_tab(self) -> None:
        self.console_textbox = ctk.CTkTextbox(
            self.tab_console,
            font=FONTS["Console"],
            state="disabled",
            wrap="word",
            border_width=1,
            border_color=COLORS["separator"],
        )
        self.console_textbox.pack(fill="both", expand=True, padx=8, pady=8)

        txt_handler = TextboxHandler(self.console_textbox)
        logger.addHandler(txt_handler)

        self.console_menu = Menu(self.console_textbox, tearoff=0)
        self.console_menu.add_command(label="Copy", command=self._console_copy)
        self.console_menu.add_command(
            label="Select All", command=self._console_select_all
        )
        self.console_menu.add_separator()
        self.console_menu.add_command(
            label="Clear Console", command=self._console_clear
        )

        def do_popup(event):
            try:
                self.console_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.console_menu.grab_release()

        if sys.platform == "darwin":
            self.console_textbox.bind("<Button-2>", do_popup)
            self.console_textbox.bind("<Button-3>", do_popup)
        else:
            self.console_textbox.bind("<Button-3>", do_popup)

    def _show_data_template(self) -> None:
        template_win = ctk.CTkToplevel(self)
        template_win.title("Data Format Template")
        template_win.geometry("650x450")
        template_win.transient(self)
        template_win.grab_set()

        ctk.CTkLabel(
            template_win, text="Required Input File Format", font=FONTS["H1"]
        ).pack(pady=(32, 8))

        info_text = (
            "Ensure your uploaded Excel or CSV files match this structure.\n"
            "The app automatically detects alternative column names \n(like 'Lat', 'Lon', 'Site', 'Elevation')."
        )
        ctk.CTkLabel(
            template_win,
            text=info_text,
            justify="center",
            text_color=COLORS["text_secondary"],
            font=FONTS["Body"],
        ).pack(pady=(0, 24))

        table_frame = ctk.CTkFrame(
            template_win,
            fg_color=COLORS["bg_card"],
            corner_radius=12,
            border_width=1,
            border_color=COLORS["separator"],
        )
        table_frame.pack(fill="x", padx=40, pady=8)

        headers = ["Point Name", "Latitude", "Longitude", "Height (Optional)"]
        for col_idx, h in enumerate(headers):
            ctk.CTkLabel(table_frame, text=h, font=FONTS["BodyBold"]).grid(
                row=0, column=col_idx, padx=16, pady=16, sticky="w"
            )

        mock_data = [
            ["Site_A", "34.0522", "-118.2437", "30"],
            ["Site_B", "34.0528", "-118.2451", "15"],
            ["Site_C", "34.0535", "-118.2465", "45"],
        ]

        for row_idx, row in enumerate(mock_data, start=1):
            for col_idx, cell in enumerate(row):
                ctk.CTkLabel(table_frame, text=cell, font=FONTS["Body"]).grid(
                    row=row_idx, column=col_idx, padx=16, pady=8, sticky="w"
                )

        for i in range(4):
            table_frame.grid_columnconfigure(i, weight=1)

        ctk.CTkFrame(template_win, fg_color="transparent", height=16).pack()

        dl_btn = ctk.CTkButton(
            template_win,
            text="⬇️ Download Sample Format",
            font=FONTS["BodyBold"],
            height=48,
            corner_radius=12,
            command=self._download_sample_file,
        )
        dl_btn.pack(pady=16)

    def _download_sample_file(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV File", "*.csv"), ("Excel File", "*.xlsx")],
            initialfile="Data_Format_Template.csv",
            title="Save Sample Template",
        )
        if not path:
            return

        mock_df = pd.DataFrame(
            [
                {
                    "Point Name": "Site_A",
                    "Latitude": 34.0522,
                    "Longitude": -118.2437,
                    "Height": 30,
                },
                {
                    "Point Name": "Site_B",
                    "Latitude": 34.0528,
                    "Longitude": -118.2451,
                    "Height": 15,
                },
                {
                    "Point Name": "Site_C",
                    "Latitude": 34.0535,
                    "Longitude": -118.2465,
                    "Height": 45,
                },
            ]
        )

        try:
            if path.endswith(".xlsx"):
                mock_df.to_excel(path, index=False)
            else:
                mock_df.to_csv(path, index=False)

            logger.info(f"Sample data template exported to: {path}")
            messagebox.showinfo(
                "Success",
                f"Sample format template saved successfully!\n\nFile is located at:\n{path}",
            )
            open_file_auto(path)
        except Exception as e:
            logger.error(f"Failed to export sample template: {e}")
            messagebox.showerror("Export Error", f"Failed to save file:\n{str(e)}")

    def _update_workspace_menu(self) -> None:
        self.workspace_menu.delete(0, tk.END)
        self.workspace_menu.add_command(
            label="Save Current Workspace As...", command=self._save_workspace_as
        )
        self.workspace_menu.add_separator()

        for w_name in self.root_config.workspaces.keys():

            def make_cmd(name: str):
                return lambda: self._load_workspace(name)

            mark = "✓  " if w_name == self.root_config.active_workspace else "     "
            self.workspace_menu.add_command(
                label=f"{mark}{w_name}", command=make_cmd(w_name)
            )

    def _get_current_ui_workspace(self) -> WorkspaceConfig:
        return WorkspaceConfig(
            scenario=self.var_scen.get(),
            calculation_method=self.var_calc.get(),
            top_n_targets=self.var_topn.get(),
            max_radius_m=self.var_rad.get(),
            enable_anomaly_detection=self.var_anomaly.get(),
            enable_los_check=self.var_los.get(),
        )

    def _apply_workspace_to_ui(self, config: WorkspaceConfig) -> None:
        self.var_scen.set(config.scenario)
        self.var_calc.set(config.calculation_method)
        self.var_topn.set(config.top_n_targets)
        self.var_rad.set(config.max_radius_m)
        self.var_anomaly.set(config.enable_anomaly_detection)
        self.var_los.set(config.enable_los_check)

        if (
            config.top_n_targets != 1
            or config.max_radius_m != 0.0
            or not config.enable_anomaly_detection
            or config.enable_los_check
        ):
            self._toggle_advanced(show=True)
        else:
            self._toggle_advanced(show=False)

    def _save_workspace_as(self) -> None:
        dialog = ctk.CTkInputDialog(
            text="Enter a name for this Workspace (e.g. 'Urban 5GHz Check'):",
            title="Save Workspace",
        )
        name = dialog.get_input()
        if not name:
            return

        self.root_config.workspaces[name] = self._get_current_ui_workspace()
        self.root_config.active_workspace = name
        ConfigManager.save(self.root_config)
        self._update_workspace_menu()
        logger.info(f"Workspace saved and activated: {name}")

    def _load_workspace(self, name: str) -> None:
        if name in self.root_config.workspaces:
            self._apply_workspace_to_ui(self.root_config.workspaces[name])
            self.root_config.active_workspace = name
            ConfigManager.save(self.root_config)
            self._update_workspace_menu()
            logger.info(f"Switched Workspace: {name}")

    def _set_theme(self, theme: str) -> None:
        ctk.set_appearance_mode(theme)
        self.root_config.ui_theme = theme
        ConfigManager.save(self.root_config)

    def _set_scaling(self, scale_str: str) -> None:
        try:
            scale_val = float(scale_str.strip("%")) / 100.0
            ctk.set_widget_scaling(scale_val)
            ctk.set_window_scaling(scale_val)
            self.root_config.ui_scaling = scale_str
            ConfigManager.save(self.root_config)
        except Exception as e:
            logger.error(f"Failed to change scaling: {e}")

    def _open_log_file(self) -> None:
        if not open_file_auto(os.path.join(os.path.expanduser("~"), LOG_FILE)):
            messagebox.showinfo("Not Found", "Log file does not exist yet.")

    def _open_config_folder(self) -> None:
        open_file_auto(os.path.dirname(os.path.abspath(CONFIG_FILE)))

    def _open_last_export_dir(self) -> None:
        if self.root_config.last_export_dir and os.path.exists(
            self.root_config.last_export_dir
        ):
            open_file_auto(self.root_config.last_export_dir)
        else:
            messagebox.showinfo(
                "Not Found", "Last export directory is not set or no longer exists."
            )

    def _console_clear(self) -> None:
        self.console_textbox.configure(state="normal")
        self.console_textbox.delete("1.0", "end")
        self.console_textbox.configure(state="disabled")

    def _console_copy(self) -> None:
        try:
            selected = self.console_textbox.selection_get()
            if selected:
                self.clipboard_clear()
                self.clipboard_append(selected)
        except Exception:
            pass

    def _console_select_all(self) -> None:
        self.console_textbox.focus_set()
        try:
            self.console_textbox._textbox.tag_add("sel", "1.0", "end")
        except Exception:
            pass

    def _pulse_abort_btn(self, step=0) -> None:
        """Creates a smooth animated red pulse effect on the Abort button during active processing."""
        if not self.is_processing:
            self.btn_abort.configure(fg_color="#d13438")
            return
        colors = ["#d13438", "#ff4444"]
        self.btn_abort.configure(fg_color=colors[step % 2])
        self.after(600, lambda: self._pulse_abort_btn(step + 1))

    def _abort_process(self) -> None:
        if not self.abort_event.is_set():
            self.abort_event.set()
            self.btn_abort.configure(state="disabled")
            self.lbl_stat.configure(text="Aborting process... please wait.")
            logger.warning("Abort signal received! Canceling remaining tasks...")

    def _handle_abort(self) -> None:
        self.prog.set(0)
        self.is_processing = False
        self.lbl_stat.configure(text="Process Aborted by User.")
        self.btn_proc.configure(state="normal")
        self.btn_abort.configure(state="disabled")
        logger.warning("Process successfully aborted. Remaining tasks canceled.")

    # ==========================================
    # MODIFICATION 2: About the Tool Window
    # ==========================================
    def _show_about(self) -> None:
        about_win = ctk.CTkToplevel(self)
        about_win.title("About & Documentation")
        about_win.geometry("700x650")
        about_win.transient(self)
        about_win.grab_set()

        # Beautiful scrollable layout for the about content
        scroll_container = ctk.CTkScrollableFrame(about_win, fg_color="transparent")
        scroll_container.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(
            scroll_container,
            text=f"{PROGRAM_NAME} {VERSION_NUMBER}",
            font=FONTS["H1"],
            text_color=COLORS["accent"],
        ).pack(pady=(16, 8))
        ctk.CTkLabel(
            scroll_container,
            text=f"Developed by {AUTHOR_NAME}",
            font=FONTS["Body"],
            text_color=COLORS["text_secondary"],
        ).pack(pady=(0, 24))

        # Comprehensive Summary
        intro_text = (
            "Telecom Geo-Distance Calculator Pro is a premium, high-performance desktop application engineered specifically "
            "for RF and Telecom professionals. It streamlines the complex task of calculating highly accurate spatial distances, "
            "bearings, and line-of-sight clearances across massive geographic datasets."
        )
        ctk.CTkLabel(
            scroll_container,
            text=intro_text,
            font=FONTS["Body"],
            justify="left",
            wraplength=600,
        ).pack(fill="x", padx=16, pady=(0, 24))

        # Beautifully formatted feature cards
        features = {
            "🚀 Core Processing Capabilities": "Dual modes support both bulk file processing (XLSX, CSV, TXT, GeoJSON) for massive networks, and a Single Point Calculator for quick, on-the-fly spatial checks.",
            "📐 Advanced Math Engines": "Choose between Fast Haversine (spherical model) for rapid approximations or High-Precision Geodesic calculations utilizing the WGS84 ellipsoid for absolute accuracy over long distances.",
            "🏔️ 3D Slant Distance Integration": "Automatically detects 'Height' or 'Elevation' columns to convert standard 2D ground distances into true 3D spatial volumetric distances.",
            "📡 Line-of-Sight (LOS) Checking": "Interfaces seamlessly with the Open-Elevation API to sample terrain paths. Evaluates 60% of the 1st Fresnel zone at 5GHz, automatically incorporating K=4/3 Earth Curvature to detect blockages.",
            "🤖 Machine Learning Anomaly Filtering": "Utilizes an Isolation Forest ML algorithm to automatically analyze coordinate pairs, identifying and dropping statistical outliers and typos before they ruin your data model.",
            "🧹 Automated Data Auditing": "Ensures pristine output by pre-validating files, dropping corrupted points, and seamlessly exporting an audit-ready 'Dropped Data' report alongside your results.",
        }

        for title, desc in features.items():
            frm = ctk.CTkFrame(
                scroll_container,
                fg_color=COLORS["bg_card"],
                corner_radius=8,
                border_width=1,
                border_color=COLORS["separator"],
            )
            frm.pack(fill="x", padx=16, pady=8)
            ctk.CTkLabel(frm, text=title, font=FONTS["H2"]).pack(
                anchor="w", padx=16, pady=(16, 4)
            )
            ctk.CTkLabel(
                frm,
                text=desc,
                font=FONTS["Body"],
                justify="left",
                wraplength=560,
                text_color=COLORS["text_secondary"],
            ).pack(anchor="w", padx=16, pady=(0, 16))

        # Prominent Close Button
        ctk.CTkButton(
            about_win,
            text="Close Documentation",
            font=FONTS["BodyBold"],
            height=48,
            corner_radius=12,
            fg_color=COLORS["separator"],
            text_color=COLORS["text_primary"],
            hover_color=COLORS["bg_main"],
            command=about_win.destroy,
        ).pack(pady=16)

    # ==========================================
    # MODIFICATION 1: Modernize Contact Us Menu
    # ==========================================
    def _show_contact(self) -> None:
        contact_win = ctk.CTkToplevel(self)
        contact_win.title("Contact & Support")
        contact_win.geometry("450x380")
        contact_win.transient(self)
        contact_win.grab_set()

        ctk.CTkLabel(
            contact_win,
            text="Get in Touch",
            font=FONTS["H1"],
            text_color=COLORS["accent"],
        ).pack(pady=(32, 8))
        ctk.CTkLabel(
            contact_win,
            text="For support, bug reports, or feature requests.",
            font=FONTS["Body"],
            text_color=COLORS["text_secondary"],
        ).pack(pady=(0, 24))

        # Info Card Frame
        info_frm = ctk.CTkFrame(
            contact_win,
            fg_color=COLORS["bg_card"],
            corner_radius=12,
            border_width=1,
            border_color=COLORS["separator"],
        )
        info_frm.pack(fill="x", padx=40, pady=8)

        ctk.CTkLabel(info_frm, text="Developer:", font=FONTS["BodyBold"]).grid(
            row=0, column=0, sticky="e", padx=(24, 8), pady=(24, 8)
        )
        ctk.CTkLabel(info_frm, text=AUTHOR_NAME, font=FONTS["Body"]).grid(
            row=0, column=1, sticky="w", padx=(8, 24), pady=(24, 8)
        )

        ctk.CTkLabel(info_frm, text="Email:", font=FONTS["BodyBold"]).grid(
            row=1, column=0, sticky="e", padx=(24, 8), pady=(8, 24)
        )
        ctk.CTkLabel(
            info_frm, text="Mohamed--Ashraf@outlook.com", font=FONTS["Body"]
        ).grid(row=1, column=1, sticky="w", padx=(8, 24), pady=(8, 24))

        # Styled LinkedIn Button
        btn_linkedin = ctk.CTkButton(
            contact_win,
            text="🔗 Connect on LinkedIn",
            font=FONTS["BodyBold"],
            fg_color="#0A66C2",
            hover_color="#004182",
            height=48,
            corner_radius=12,
            command=lambda: webbrowser.open(
                "https://www.linkedin.com/in/mohamed---ashraf/"
            ),
        )
        btn_linkedin.pack(pady=(24, 16))

        # Close button
        ctk.CTkButton(
            contact_win,
            text="Close",
            fg_color="transparent",
            text_color=COLORS["text_secondary"],
            font=FONTS["BodyBold"],
            hover_color=COLORS["separator"],
            command=contact_win.destroy,
        ).pack()

    def _calc_single(self) -> None:
        try:
            slat, slon = float(self.e_slat.get()), float(self.e_slon.get())
            tlat, tlon = float(self.e_tlat.get()), float(self.e_tlon.get())

            d2d, b = DistanceCalculator.calculate_geodesic_vectorized(
                slat, slon, np.array([tlat]), np.array([tlon])
            )
            res_str = f"Distance: {d2d[0]:,.2f} m  |  Bearing: {b[0]:.2f}°"
            self.lbl_sres.configure(
                text=res_str, text_color=COLORS["accent"], font=FONTS["H2"]
            )
            logger.info(f"Single Calc Result -> {res_str}")
        except ValueError:
            self.lbl_sres.configure(
                text="Error: Enter valid numeric coordinates.",
                text_color="#d13438",
                font=FONTS["BodyBold"],
            )
        except Exception as e:
            self.lbl_sres.configure(
                text=f"Error: {str(e)}", text_color="#d13438", font=FONTS["BodyBold"]
            )

    def _toggle_advanced(self, show: bool = None) -> None:
        if show is not None:
            self.advanced_visible = show
        else:
            self.advanced_visible = not self.advanced_visible

        if self.advanced_visible:
            self.btn_adv_toggle.configure(text="⚙️ Advanced Settings ▲")
            self.frm_adv_container.grid(
                row=2, column=0, columnspan=2, sticky="ew", padx=8
            )
        else:
            self.btn_adv_toggle.configure(text="⚙️ Advanced Settings ▼")
            self.frm_adv_container.grid_forget()

    def _peek_headers(self, path: str) -> Tuple[bool, List[str]]:
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in [".xlsx", ".xls"]:
                df = pd.read_excel(
                    path, nrows=0, engine="openpyxl" if ext == ".xlsx" else None
                )
            elif ext in [".csv", ".txt"]:
                df = pd.read_csv(path, sep=None, engine="python", nrows=0)
            elif ext == ".geojson":
                return True, []
            else:
                return False, ["Unsupported Format"]

            mapping = {}
            for col in df.columns:
                norm = str(col).strip().lower().replace("_", " ").replace("-", " ")
                for std, vars in ColumnMapper.STANDARD_COLUMNS.items():
                    if (
                        any(v in norm for v in vars)
                        and col != std
                        and std not in df.columns
                    ):
                        mapping[col] = std
                        break
            if mapping:
                df = df.rename(columns=mapping)

            req_cols = ["Point name", "Latitude", "Longitude"]
            missing = [c for c in req_cols if c not in df.columns]
            return len(missing) == 0, missing
        except Exception:
            return False, ["Unreadable file format or empty data"]

    def _browse(self, mode: str) -> None:
        path = filedialog.askopenfilename(
            filetypes=[
                ("All Supported", "*.xlsx *.xls *.csv *.txt *.geojson"),
                ("Excel", "*.xlsx *.xls"),
                ("CSV", "*.csv"),
                ("Text", "*.txt"),
                ("GeoJSON", "*.geojson"),
            ]
        )
        if not path:
            return

        is_valid, missing = self._peek_headers(path)
        status_text, status_color = "✅", "green" if is_valid else "red"
        if not is_valid:
            status_text = "❌"
        tooltip_msg = (
            "Valid file headers detected."
            if is_valid
            else f"Missing required columns:\n{', '.join(missing)}\n\nClick '📋 Data Template' above for an example."
        )

        trunc_path = _truncate_path(path)

        if mode == "src":
            self.src_path = path
            self.entry_src.configure(state="normal")
            self.entry_src.delete(0, "end")
            self.entry_src.insert(0, trunc_path)
            self.entry_src.configure(state="readonly")
            self.lbl_src_status.configure(text=status_text, text_color=status_color)
            self.src_tooltip.update_text(tooltip_msg)
            logger.info(f"Source file selected: {path} | Valid: {is_valid}")
        else:
            self.tgt_path = path
            self.entry_tgt.configure(state="normal")
            self.entry_tgt.delete(0, "end")
            self.entry_tgt.insert(0, trunc_path)
            self.entry_tgt.configure(state="readonly")
            self.lbl_tgt_status.configure(text=status_text, text_color=status_color)
            self.tgt_tooltip.update_text(tooltip_msg)
            logger.info(f"Target file selected: {path} | Valid: {is_valid}")

    def _start_phase1(self) -> None:
        if not self.src_path or not self.tgt_path:
            messagebox.showerror(
                "Missing Files", "Select both Source and Target files."
            )
            return

        active_name = self.root_config.active_workspace
        self.root_config.workspaces[active_name] = self._get_current_ui_workspace()
        ConfigManager.save(self.root_config)
        current_cfg = self.root_config.workspaces[active_name]

        if current_cfg.enable_los_check:
            if not messagebox.askyesno(
                "API Limits Warning",
                "⚠️ LINE OF SIGHT API WARNING ⚠️\n\nLOS checking uses a free external API (Open-Elevation).\nIt can be extremely SLOW and may time-out on massive datasets.\n\nTo prevent full locking, it will ONLY evaluate the #1 closest target per source point.\n\nDo you want to proceed?",
            ):
                logger.info("LOS Validation aborted by user due to warning.")
                return

        self.btn_proc.configure(state="disabled")
        self.btn_abort.configure(state="normal")
        self.lbl_stat.configure(text="Phase 1: Validating Data...")
        self.prog.set(0.2)

        # UI pulse animation activation
        self.is_processing = True
        self._pulse_abort_btn()

        logger.info("--- PHASE 1: Data Parsing & Validation Started ---")
        threading.Thread(target=self._task_phase1, daemon=True).start()

    def _task_phase1(self) -> None:
        try:
            current_cfg = self.root_config.workspaces[self.root_config.active_workspace]
            self.s_df, self.s_drop, s_stat = DistanceCalculator.parse_file(
                self.src_path, "Auto", current_cfg
            )
            self.t_df, self.t_drop, t_stat = DistanceCalculator.parse_file(
                self.tgt_path, "Auto", current_cfg
            )

            def on_cancel():
                self.is_processing = False
                self.btn_proc.configure(state="normal")
                self.btn_abort.configure(state="disabled")
                self.lbl_stat.configure(text="Ready.")
                self.prog.set(0)

            self.after(
                0,
                lambda: SanityCheckPopup(
                    self, s_stat, t_stat, self._start_phase2, on_cancel
                ),
            )

        except Exception as e:
            logger.error(f"Validation Phase Error: {str(e)}")
            self.after(0, lambda: messagebox.showerror("Validation Error", str(e)))
            self.after(0, lambda: setattr(self, "is_processing", False))
            self.after(0, lambda: self.btn_proc.configure(state="normal"))
            self.after(0, lambda: self.btn_abort.configure(state="disabled"))
            self.after(0, lambda: self.prog.set(0))
            self.after(0, lambda: self.lbl_stat.configure(text="Ready."))

    def _start_phase2(self) -> None:
        self.lbl_stat.configure(text="Phase 2: Calculating Matrix...")
        self.prog.set(0.5)
        logger.info("--- PHASE 2: Calculation Engine Started ---")
        threading.Thread(target=self._task_phase2, daemon=True).start()

    def _task_phase2(self) -> None:
        start_time = time.time()
        try:
            current_cfg = self.root_config.workspaces[self.root_config.active_workspace]
            cfg_dict = {
                "method": current_cfg.calculation_method,
                "ignore_0": "Ignore" in current_cfg.scenario,
                "top_n": current_cfg.top_n_targets,
                "max_rad": current_cfg.max_radius_m,
                "use_3d": "Height" in self.s_df.columns
                and "Height" in self.t_df.columns,
                "enable_los": current_cfg.enable_los_check,
            }

            t_lats, t_lons = self.t_df["Latitude"].values, self.t_df["Longitude"].values
            t_names = self.t_df["Point name"].values.astype(str)
            t_heights = (
                self.t_df["Height"].values
                if cfg_dict["use_3d"]
                else np.zeros_like(t_lats)
            )

            results: List[Dict[str, Any]] = []
            total = len(self.s_df)
            chunk_size = current_cfg.chunk_size
            chunks = [
                self.s_df.iloc[i : i + chunk_size] for i in range(0, total, chunk_size)
            ]
            args_list = [(chk, cfg_dict) for chk in chunks]

            logger.info(
                f"Processing matrix: {total} Source points -> {len(self.t_df)} Target points."
            )
            self.abort_event.clear()

            with ProcessPoolExecutor(
                initializer=_init_worker, initargs=(t_lats, t_lons, t_heights, t_names)
            ) as executor:
                self.active_futures = [
                    executor.submit(DistanceCalculator.process_chunk, arg)
                    for arg in args_list
                ]
                for i, future in enumerate(self.active_futures):
                    while not future.done():
                        if self.abort_event.is_set():
                            for f in self.active_futures:
                                f.cancel()
                            self.after(0, self._handle_abort)
                            return
                        time.sleep(0.1)

                    if self.abort_event.is_set():
                        self.after(0, self._handle_abort)
                        return

                    res = future.result()
                    results.extend(res)
                    self.after(0, self.prog.set, 0.5 + (0.4 * ((i + 1) / len(chunks))))

            res_df = pd.DataFrame(results)
            logger.info(
                f"Calculation Complete. Found {len(res_df)} connections in {time.time() - start_time:.2f} seconds."
            )
            self.after(0, lambda: self._save_and_finish(res_df))

        except Exception as e:
            logger.error(f"Calculation Failure:\n{traceback.format_exc()}")
            self.after(0, lambda: messagebox.showerror("Calculation Error", str(e)))
            self.after(0, lambda: setattr(self, "is_processing", False))
            self.after(0, lambda: self.btn_proc.configure(state="normal"))
            self.after(0, lambda: self.btn_abort.configure(state="disabled"))

    def _save_and_finish(self, res_df: pd.DataFrame) -> None:
        self.prog.set(1.0)
        self.is_processing = False
        self.lbl_stat.configure(text="Saving reports...")

        base_dir = os.path.dirname(self.src_path)
        timestamp = datetime.now().strftime("%H%M%S")
        res_path = os.path.join(base_dir, f"Distance_Results_{timestamp}.xlsx")

        try:
            res_df.to_excel(res_path, index=False)
            open_file_auto(res_path)
            logger.info(f"Success! Saved Matrix to: {res_path}")

            self.root_config.last_export_dir = base_dir
            ConfigManager.save(self.root_config)

            if not self.s_drop.empty or not self.t_drop.empty:
                drop_path = os.path.join(
                    base_dir, f"Dropped_Data_Report_{timestamp}.xlsx"
                )
                with pd.ExcelWriter(drop_path) as writer:
                    if not self.s_drop.empty:
                        self.s_drop.to_excel(
                            writer, sheet_name="Dropped_Source", index=False
                        )
                    if not self.t_drop.empty:
                        self.t_drop.to_excel(
                            writer, sheet_name="Dropped_Target", index=False
                        )
                open_file_auto(drop_path)

            messagebox.showinfo(
                "Success", "Processing complete! Files have been saved and opened."
            )
        except Exception as e:
            logger.error(f"Error during file save: {str(e)}")
            messagebox.showerror("Save Error", f"Failed to save results: {e}")

        self.btn_proc.configure(state="normal")
        self.btn_abort.configure(state="disabled")
        self.lbl_stat.configure(text="Ready.")

    def _on_close(self) -> None:
        if (
            hasattr(self, "temp_ico_path")
            and self.temp_ico_path
            and os.path.exists(self.temp_ico_path)
        ):
            try:
                os.remove(self.temp_ico_path)
            except Exception:
                pass
        self.destroy()


if __name__ == "__main__":
    mp.freeze_support()
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    GeoDistanceApp().mainloop()
