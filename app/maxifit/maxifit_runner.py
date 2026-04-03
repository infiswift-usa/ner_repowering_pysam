"""
Maxifit automation: restart app, select area/point, create plant from panel placement permutations.

Run from project root:
  python -m app.maxifit.maxifit_runner
  python -m app.maxifit.maxifit_runner --specs specs/plant_specs.json
"""
from __future__ import annotations

import argparse
import csv
import json
import logging
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any
from uuid import UUID

import pandas as pd


logger = logging.getLogger(__name__)

# Project root
PROJECT_ROOT = Path(__file__).resolve().parents[1]
MASTER_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(PROJECT_ROOT))
DEFAULT_MANIFEST = PROJECT_ROOT/"specs"/ "manifest.csv"
SIMULATION_RUNS_DIR = PROJECT_ROOT / "output" / "simulation_runs"
MAXIFIT_EXE =r"C:\Program Files (x86)\MaxiFitPointVer5\過積載シミュレーション.exe"
MAIN_WINDOW_TITLE_RE = r".*MAXIFIT.*simulator.*"
UI_SLEEP = 0.5
DIALOG_WAIT = 10
from maxifit.run_results import maxifit_run_to_dataframe


def _close_maxifit_if_running() -> bool:
    """Close Maxifit if running. Return True if it was running."""
    try:
        from pywinauto import Application

        app = Application(backend="uia").connect(path=MAXIFIT_EXE)
        for w in app.windows():
            try:
                w.close()
                time.sleep(0.3)
            except Exception:
                pass
        time.sleep(1.0)
        return True
    except Exception:
        return False


def _start_maxifit() -> Any:
    """Start Maxifit and return connected Application."""
    from pywinauto import Application
    from pywinauto.timings import Timings

    Timings.window_find_timeout = 30
    app = Application(backend="uia").start(cmd_line=MAXIFIT_EXE)
    Timings.fast()
    main = app.window(title_re=MAIN_WINDOW_TITLE_RE)
    main.wait("ready", timeout=30)
    try:
        main.maximize()
        time.sleep(0.5)
    except Exception:
        pass
    return app


def _load_area_point_indices(manifest_path: Path) -> tuple[list[str], dict[str, list[str]]]:
    """Load ordered area list and points-by-area from manifest CSV."""
    areas: list[str] = []
    points_by_area: dict[str, list[str]] = {}
    seen: set[tuple[str, str]] = set()

    with manifest_path.open(encoding="utf-8") as f:
        for row in csv.DictReader(f):
            a, p = row.get("area", ""), row.get("point", "")
            if not a or not p:
                continue
            key = (a, p)
            if key not in seen:
                seen.add(key)
                if a not in points_by_area:
                    areas.append(a)
                    points_by_area[a] = []
                points_by_area[a].append(p)

    return areas, points_by_area


def _select_combo_by_index(main: Any, combo: Any, index: int) -> None:
    """Select ComboBox item by index (HOME + DOWN*n + ENTER). No GUI reads needed."""
    from pywinauto.keyboard import send_keys

    try:
        main.set_focus()
    except Exception:
        pass
    time.sleep(0.25)
    combo.click_input()
    time.sleep(0.25)
    combo.expand()
    time.sleep(0.3)

    send_keys("{HOME}")
    time.sleep(0.1)
    for _ in range(index):
        send_keys("{DOWN}")
        time.sleep(0.03)
    send_keys("{ENTER}")
    time.sleep(0.1)

    try:
        combo.collapse()
    except Exception:
        pass
    time.sleep(UI_SLEEP)


def _select_combo_by_text(main: Any, combo: Any, item_text: str) -> None:
    """Select ComboBox item by text (for PCS type, panel name - short lists)."""
    from pywinauto.keyboard import send_keys

    try:
        main.set_focus()
    except Exception:
        pass
    time.sleep(0.25)
    combo.click_input()
    time.sleep(0.25)
    combo.expand()
    time.sleep(0.3)

    list_controls = combo.children(control_type="List")
    if list_controls:
        items = list_controls[0].children(control_type="ListItem")
        for item in items:
            text = (item.window_text() or "").strip()
            if text == item_text or item_text in text:
                item.click_input()
                break
        else:
            send_keys("{ESC}")
            raise RuntimeError(f"Combo item '{item_text}' not found")
    else:
        send_keys("{ESC}")
        raise RuntimeError(f"Combo has no List; cannot select '{item_text}'")

    try:
        combo.collapse()
    except Exception:
        pass
    time.sleep(UI_SLEEP)


def _select_area_and_point(
    app: Any,
    area: str,
    point: str,
    manifest_path: Path,
    main: Any = None,
) -> None:
    """Select area and point using indices from manifest (avoids scrolling whole list)."""
    if main is None:
        main = app.window(title_re=MAIN_WINDOW_TITLE_RE)
    areas, points_by_area = _load_area_point_indices(manifest_path)

    # Resolve area: exact match or first that contains
    area_index = None
    resolved_area = None
    for i, a in enumerate(areas):
        if a == area or area in a:
            area_index = i
            resolved_area = a
            break
    if area_index is None:
        raise RuntimeError(f"Area '{area}' not found in manifest")

    # Resolve point within that area
    points = points_by_area.get(resolved_area, [])
    point_index = None
    for i, p in enumerate(points):
        if p == point or point in p:
            point_index = i
            break
    if point_index is None:
        raise RuntimeError(f"Point '{point}' not found for area '{resolved_area}' in manifest")

    area_combo = main.child_window(auto_id="areaSelectComboBox", control_type="ComboBox")
    point_combo = main.child_window(auto_id="pointSelectComboBox", control_type="ComboBox")

    _select_combo_by_index(main, area_combo, area_index)
    time.sleep(2.0)  # Point dropdown updates after area change
    _select_combo_by_index(main, point_combo, point_index)


def _set_numeric_updown(combo: Any, value: int | float) -> None:
    """Set value in a NumericUpDown ComboBox via its Edit child."""
    edit = combo.child_window(control_type="Edit")
    edit.set_focus()
    edit.set_edit_text("")
    time.sleep(0.1)
    edit.set_edit_text(str(int(value)))
    time.sleep(0.1)


def _configure_pcs_panel(
    app: Any,
    pcs_config: dict[str, Any],
    modules_per_string: int,
    main: Any = None,
) -> None:
    """Open panel dialog for selected PCS and set strings, tilt, azimuth, module."""
    if main is None:
        main = app.window(title_re=MAIN_WINDOW_TITLE_RE)
    panel_btn = main.child_window(auto_id="panelSettingButton01", control_type="Button")
    panel_btn.click_input()

    panel_win = app.window(title_re=".*パネル入力.*")
    panel_win.wait("ready", timeout=DIALOG_WAIT)

    series_combo = panel_win.child_window(auto_id="panelSeriesNumericUpDown", control_type="ComboBox")
    _set_numeric_updown(series_combo, modules_per_string)

    parallel_combo = panel_win.child_window(auto_id="panelParallelNumericUpDown", control_type="ComboBox")
    _set_numeric_updown(parallel_combo, pcs_config["strings"])

    tilt_combo = panel_win.child_window(auto_id="installationAngleNumericUpDown", control_type="ComboBox")
    _set_numeric_updown(tilt_combo, pcs_config["tilt"])

    azimuth_combo = panel_win.child_window(auto_id="azimuthNumericUpDown", control_type="ComboBox")
    _set_numeric_updown(azimuth_combo, pcs_config.get("azimuth", -5))

    panel_name_combo = panel_win.child_window(auto_id="panelSelectComboBoxSub", control_type="ComboBox")
    module_type = pcs_config.get("module_type", "NER132M625E-NGD")
    _select_combo_by_text(panel_win, panel_name_combo, module_type)

    enter_btn = panel_win.child_window(auto_id="enterButton", control_type="Button")
    enter_btn.click_input()
    panel_win.wait_not("exists", timeout=5)


def _configure_pcs_panel_strings_only(app: Any, strings: int, main: Any = None) -> None:
    """Open panel dialog and update only 並列数 (strings). Used after copying a PCS."""
    if main is None:
        main = app.window(title_re=MAIN_WINDOW_TITLE_RE)
    panel_btn = main.child_window(auto_id="panelSettingButton01", control_type="Button")
    panel_btn.click_input()

    panel_win = app.window(title_re=".*パネル入力.*")
    panel_win.wait("ready", timeout=DIALOG_WAIT)

    parallel_combo = panel_win.child_window(auto_id="panelParallelNumericUpDown", control_type="ComboBox")
    _set_numeric_updown(parallel_combo, strings)

    enter_btn = panel_win.child_window(auto_id="enterButton", control_type="Button")
    enter_btn.click_input()
    panel_win.wait_not("exists", timeout=5)


def _get_pcs_indices_to_update(
    config_prev: dict[str, Any],
    config_next: dict[str, Any],
) -> list[int]:
    """Return 1-based PCS indices requiring panel updates when transitioning config_prev -> config_next."""
    prev_pcs = config_prev["pcs_config"]
    next_pcs = config_next["pcs_config"]
    n = len(prev_pcs)
    if config_prev["tilt"] != config_next["tilt"]:
        return list(range(1, n + 1))
    return [i + 1 for i in range(n) if prev_pcs[i]["strings"] != next_pcs[i]["strings"]]


def _update_pcs_panel_strings_and_tilt(
    app: Any, strings: int, tilt: int | float, main: Any = None
) -> None:
    """Open panel dialog and update only 並列数 (strings) and 設置角度 (tilt)."""
    if main is None:
        main = app.window(title_re=MAIN_WINDOW_TITLE_RE)
    panel_btn = main.child_window(auto_id="panelSettingButton01", control_type="Button")
    panel_btn.click_input()

    panel_win = app.window(title_re=".*パネル入力.*")
    panel_win.wait("ready", timeout=DIALOG_WAIT)

    parallel_combo = panel_win.child_window(auto_id="panelParallelNumericUpDown", control_type="ComboBox")
    _set_numeric_updown(parallel_combo, strings)
    tilt_combo = panel_win.child_window(auto_id="installationAngleNumericUpDown", control_type="ComboBox")
    _set_numeric_updown(tilt_combo, tilt)

    enter_btn = panel_win.child_window(auto_id="enterButton", control_type="Button")
    enter_btn.click_input()
    panel_win.wait_not("exists", timeout=5)


def _transition_to_config(
    app: Any,
    config_prev: dict[str, Any],
    config_next: dict[str, Any],
    main: Any = None,
) -> None:
    """Transition from config_prev to config_next by updating only PCSs that changed."""
    if main is None:
        main = app.window(title_re=MAIN_WINDOW_TITLE_RE)
    indices = _get_pcs_indices_to_update(config_prev, config_next)
    if not indices:
        return
    next_pcs = config_next["pcs_config"]
    for pcs_id in indices:
        _select_pcs_by_index(app, pcs_id, main=main)
        pcs_cfg = next_pcs[pcs_id - 1]
        _update_pcs_panel_strings_and_tilt(
            app, pcs_cfg["strings"], pcs_cfg["tilt"], main=main
        )


def _copy_pcs(app: Any, quantity: int = 1, main: Any = None) -> None:
    """Copy the selected PCS. Set quantity first, then click PCSコピー."""
    if main is None:
        main = app.window(title_re=MAIN_WINDOW_TITLE_RE)
    qty_combo = main.child_window(auto_id="pcsCopyNumericUpDown", control_type="ComboBox")
    _set_numeric_updown(qty_combo, quantity)
    time.sleep(0.1)
    copy_btn = main.child_window(auto_id="pcsCopyButton", control_type="Button")
    copy_btn.click_input()
    time.sleep(UI_SLEEP)


def _select_pcs_by_index(app: Any, pcs_id: int, main: Any = None) -> None:
    """Select PCS in listbox by arrow-key navigation (avoids off-screen items)."""
    from pywinauto.keyboard import send_keys

    if main is None:
        main = app.window(title_re=MAIN_WINDOW_TITLE_RE)
    listbox = main.child_window(auto_id="pcsListBox", control_type="List")
    listbox.set_focus()
    time.sleep(0.2)
    send_keys("{HOME}")  # Go to first item
    time.sleep(0.1)
    for _ in range(pcs_id - 1):
        send_keys("{DOWN}")
        time.sleep(0.05)
    time.sleep(UI_SLEEP)


def _add_first_pcs(app: Any, pcs_type: str, main: Any = None) -> None:
    """Add exactly one PCS of the given type."""
    if main is None:
        main = app.window(title_re=MAIN_WINDOW_TITLE_RE)
    listbox = main.child_window(auto_id="pcsListBox", control_type="List")
    items = listbox.children(control_type="ListItem")
    current = len([i for i in items if i.window_text()])
    if current >= 1:
        return
    display_type = f"SunGrow {pcs_type}" if not pcs_type.startswith("SunGrow") else pcs_type
    pcs_combo = main.child_window(auto_id="pcsSelectComboBox", control_type="ComboBox")
    add_btn = main.child_window(auto_id="pcsAddButton", control_type="Button")
    _select_combo_by_text(main, pcs_combo, display_type)
    add_btn.click_input()
    time.sleep(UI_SLEEP)


def _setup_logger(run_dir: Path) -> None:
    """Configure logger to write to run_dir/run.log and stdout."""
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    file_handler = logging.FileHandler(run_dir / "run.log", encoding="utf-8")
    file_handler.setFormatter(fmt)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(fmt)
    logger.addHandler(console_handler)


def _setup_run_folder(
    run_id: str,
    specs: dict[str, Any],
    ordered_configs: list[dict[str, Any]],
    specs_path: Path,
) -> Path:
    """Create run folder, write run.json and permutations.json. Returns run_dir."""
    run_dir = SIMULATION_RUNS_DIR / run_id
    run_dir.mkdir(parents=True, exist_ok=True)

    location = specs.get("location", {})
    run_meta = {
        "run_id": run_id,
        "started_at": datetime.now().isoformat(),
        "specs_source": specs.get("source", ""),
        "location": location,
        "permutation_count": len(ordered_configs),
        "execution_order": list(range(1, len(ordered_configs) + 1)),
        "specs_path": str(specs_path),
    }
    with (run_dir / "run.json").open("w", encoding="utf-8") as f:
        json.dump(run_meta, f, ensure_ascii=False, indent=2)

    permutations_data = {
        "permutations": [{"id": i + 1, **cfg} for i, cfg in enumerate(ordered_configs)]
    }
    with (run_dir / "permutations.json").open("w", encoding="utf-8") as f:
        json.dump(permutations_data, f, ensure_ascii=False, indent=2)

    return run_dir


def _reencode_csv_to_utf8(csv_path: Path) -> None:
    """Read CSV as CP932 (Maxifit default on Japanese Windows) and overwrite with UTF-8."""
    if not csv_path.exists():
        return
    for enc in ("cp932", "shift_jis"):
        try:
            text = csv_path.read_text(encoding=enc)
            csv_path.write_text(text, encoding="utf-8")
            return
        except (UnicodeDecodeError, UnicodeEncodeError):
            continue


def _export_csv_to_total_chart(app: Any, output_csv_path: Path, main: Any = None) -> None:
    """Open total chart, click CSV export, save to output_csv_path via Save As dialog."""
    from pywinauto import Desktop
    from pywinauto.keyboard import send_keys

    if main is None:
        main = app.window(title_re=MAIN_WINDOW_TITLE_RE)
    chart_btn = main.child_window(auto_id="totalChartViewButton", control_type="Button")
    chart_btn.click_input()

    desktop = Desktop(backend="uia")
    chart_win = desktop.window(title_re=".*チャート.*|.*Chart.*")
    chart_win.wait("ready", timeout=DIALOG_WAIT)
    csv_btn = chart_win.child_window(auto_id="CsvOutButton", control_type="Button")
    csv_btn.click_input()

    # Wait for Save As dialog and its filename field to be ready (faster than full dialog ready)
    save_dlg = desktop.window(title_re=".*(名前を付けて|Save As|保存).*")
    try:
        save_dlg.wait("exists", timeout=5)
        # Wait for filename Edit - the control we paste into (faster than dialog "ready")
        try:
            filename_edit = save_dlg.child_window(control_type="ComboBox").child_window(
                control_type="Edit"
            )
        except Exception:
            filename_edit = save_dlg.child_window(control_type="Edit", found_index=0)
        filename_edit.wait("enabled", timeout=3)
    except Exception:
        time.sleep(0.5)  # Fallback if dialog structure differs

    path_str = str(output_csv_path)
    try:
        import win32clipboard
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, path_str)
        finally:
            win32clipboard.CloseClipboard()
    except ImportError:
        send_keys(path_str)
    else:
        send_keys("^v")
    time.sleep(0.2)
    send_keys("{ENTER}")

    try:
        save_dlg.wait_not("exists", timeout=5)
    except Exception:
        time.sleep(0.5)

    chart_win.close()
    try:
        chart_win.wait_not("exists", timeout=5)
    except Exception:
        time.sleep(0.3)
    try:
        main.set_focus()
    except Exception:
        pass

    _reencode_csv_to_utf8(output_csv_path)


class MaxifitRunner:
    """Drive Maxifit GUI to run PV simulations for each panel placement permutation.

    Expects a payload dict from :meth:`SimulationPlanner.automation_payload`, with keys
    ``specs`` and ``ordered_permutations``.

    After a successful :meth:`run`, :attr:`pv_generation_results_df` holds a combined
    dataframe (see :func:`app.maxifit.run_results.maxifit_run_to_dataframe`) aligned with
    ``dbo.pv_generation_results``.
    """

    def __init__(
        self,
        payload: dict[str, Any],
        *,
        manifest_path: Path,
        specs_path: Path,
        simulation_id: UUID | str | None = None,
        save_results_dataframe: bool = True,
    ) -> None:
        try:
            self._specs = payload["specs"]
            self._ordered = payload["ordered_permutations"]
        except KeyError as e:
            raise KeyError(
                "payload must include 'specs' and 'ordered_permutations' "
                "(use SimulationPlanner.automation_payload())"
            ) from e
        self._manifest_path = manifest_path
        self._specs_path = specs_path
        self._simulation_id = simulation_id
        self._save_results_dataframe = save_results_dataframe
        self._run_dir: Path | None = None
        self.pv_generation_results_df: pd.DataFrame | None = None

    @property
    def run_dir(self) -> Path | None:
        """Output folder from the last :meth:`run`, or ``None`` if not started or failed early."""
        return self._run_dir

    def load_results_dataframe(
        self,
        run_dir: Path | str | None = None,
        *,
        include_permutation_params: bool = True,
    ) -> pd.DataFrame:
        """Build the PV generation dataframe from a completed run folder (default: last ``run_dir``)."""
        root = Path(run_dir) if run_dir is not None else self._run_dir
        if root is None:
            raise RuntimeError("No run_dir; call run() first or pass run_dir explicitly.")
        return maxifit_run_to_dataframe(
            root,
            simulation_id=self._simulation_id,
            include_permutation_params=include_permutation_params,
        )

    def run(self) -> int:
        """Start Maxifit, select area/point, run each permutation and export CSVs. Returns 0 on success."""
        specs = self._specs
        ordered = self._ordered
        manifest_path = self._manifest_path
        specs_path = self._specs_path

        location = specs.get("location", {})
        area = location.get("area", "")
        point = location.get("point", "")
        if not area or not point:
            print("Error: specs must have location.area and location.point", file=sys.stderr)
            return 1

        run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        run_dir = _setup_run_folder(run_id, specs, ordered, specs_path)
        self._run_dir = run_dir
        self.pv_generation_results_df = None
        _setup_logger(run_dir)
        logger.info("Run folder: %s", run_dir)

        logger.info("Restarting Maxifit...")
        _close_maxifit_if_running()
        time.sleep(1.0)
        app = _start_maxifit()
        time.sleep(1.0)
        main = app.window(title_re=MAIN_WINDOW_TITLE_RE)
        t_boot = time.perf_counter()

        logger.info("Selecting area=%r, point=%r", area, point)
        _select_area_and_point(app, area, point, manifest_path, main=main)
        time.sleep(1.0)

        for i in range(len(ordered)):
            cfg = ordered[i]
            t0 = time.perf_counter()
            if i == 0:
                _create_plant_from_config(app, cfg, main=main)
                time.sleep(1.0)
            else:
                prev_config = ordered[i - 1]
                _transition_to_config(app, prev_config, cfg, main=main)
                time.sleep(1.0)
            csv_path = run_dir / f"perm_{i + 1:03d}.csv"
            _export_csv_to_total_chart(app, csv_path, main=main)
            elapsed = time.perf_counter() - t0
            logger.info(
                "Perm %d/%d: overload=%.0f%%, tilt=%.0f° -> %s (%.1fs)",
                i + 1,
                len(ordered),
                cfg["overload_rate_pct"],
                cfg["tilt"],
                f"perm_{i + 1:03d}.csv",
                elapsed,
            )

        total_elapsed = time.perf_counter() - t_boot
        logger.info("Total time from Maxifit boot: %.1fs (%.1f min)", total_elapsed, total_elapsed / 60)

        self.pv_generation_results_df = maxifit_run_to_dataframe(
            run_dir,
            simulation_id=self._simulation_id,
            include_permutation_params=True,
        )
        if self._save_results_dataframe and self.pv_generation_results_df is not None:
            out_parquet = run_dir / "pv_generation_results.parquet"
            out_csv = run_dir / "pv_generation_results.csv"
            try:
                self.pv_generation_results_df.to_parquet(out_parquet, index=False)
                logger.info(
                    "Saved %d rows to %s",
                    len(self.pv_generation_results_df),
                    out_parquet.name,
                )
            except Exception as e:
                logger.warning("Parquet export failed (%s); writing CSV instead.", e)
                self.pv_generation_results_df.to_csv(out_csv, index=False)
                logger.info(
                    "Saved %d rows to %s",
                    len(self.pv_generation_results_df),
                    out_csv.name,
                )

        logger.info("Done.")
        return 0


def _create_plant_from_config(app: Any, config: dict[str, Any], main: Any = None) -> None:
    """Create plant: add first PCS, configure it, then copy and adjust as needed."""
    if main is None:
        main = app.window(title_re=MAIN_WINDOW_TITLE_RE)
    pcs_configs = config["pcs_config"]
    if not pcs_configs:
        return
    modules_per_string = pcs_configs[0].get("modules_per_string", 16)
    pcs_type = pcs_configs[0].get("pcs_type", "SG100CX-JP")
    first_config = pcs_configs[0]

    _add_first_pcs(app, pcs_type, main=main)
    _select_pcs_by_index(app, 1, main=main)
    _configure_pcs_panel(app, first_config, modules_per_string, main=main)

    # Copy remaining PCSs in one operation using the copy amount field
    extra_pcs = len(pcs_configs) - 1
    if extra_pcs > 0:
        _select_pcs_by_index(app, 1, main=main)  # Select template (PCS 1)
        _copy_pcs(app, quantity=extra_pcs, main=main)
        for i in range(1, len(pcs_configs)):
            pcs_cfg = pcs_configs[i]
            if pcs_cfg["strings"] != first_config["strings"]:
                _select_pcs_by_index(app, i + 1, main=main)
                _configure_pcs_panel_strings_only(app, pcs_cfg["strings"], main=main)


def run(specs_path: Path, manifest_path: Path) -> int:
    """Run automation: parse specs via SimulationPlanner, then MaxifitRunner."""
    if not specs_path.exists():
        print(f"Error: specs file not found: {specs_path}", file=sys.stderr)
        return 1
    if not manifest_path.exists():
        print(f"Error: manifest file not found: {manifest_path}", file=sys.stderr)
        return 1

    from simulation_planning.simulation_planner import SimulationPlanner

    planner = SimulationPlanner.from_json_path(specs_path)
    payload = planner.automation_payload()
    runner = MaxifitRunner(payload, manifest_path=manifest_path, specs_path=specs_path)
    return runner.run()


def main() -> int:
    parser = argparse.ArgumentParser(description="Maxifit automation: restart, select area/point, create plant.")
    parser.add_argument(
        "--specs",
        type=Path,
        default=PROJECT_ROOT /"specs"/ "mie tsu_extracted.json",
        help="Path to plant specs JSON.",
    )
    parser.add_argument(
        "--manifest",
        type=Path,
        default=DEFAULT_MANIFEST,
        help="Path to area/point manifest CSV (from run_maxifit_simulation).",
    )
    args = parser.parse_args()
    return run(args.specs, args.manifest)


if __name__ == "__main__":
    sys.exit(main())