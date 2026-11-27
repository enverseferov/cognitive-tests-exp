"""Cognitive load plots from a Muse recording (before / after physical exercise).

Usage (in the activated environment):
    uv run plot-cognitive data/recording.txt                 # single recording
    uv run plot-cognitive data/before.txt data/after.txt     # before / after comparison
    uv run plot-cognitive data/before.txt data/after.txt --skip 60   # drop the first 60 s

Computed over 4-second windows with a 1 s step, clean segments only:
  * Engagement / concentration = beta / (alpha + theta)  — rises with active thinking
  * Mental workload            = theta / alpha           — rises with working-memory load
  * Relaxation                 = relative alpha          — rises in a calm state
The indices are relative: compare only the same person under the same conditions
(same task, same headset placement).
"""

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import OpenMuse
from scipy.signal import butter, filtfilt, iirnotch, welch

# project root: data and figures are always saved to <project>/data and <project>/figures
ROOT = Path(__file__).resolve().parents[2]
FIG_DIR = ROOT / "figures"

FS = 256
WIN, STEP = 4.0, 1.0  # seconds
RAIL_LO, RAIL_HI = 0.5, 1449.5  # ADC rail values = no contact
MAX_PTP = 500.0  # µV after filtering: above = artifact (blink, movement)
MAX_GYRO = 20.0  # °/s: above = head was moving
CHANNELS = ["EEG_AF7", "EEG_AF8", "EEG_TP9", "EEG_TP10"]
BANDS = {"theta": (4, 8), "alpha": (8, 13), "beta": (13, 30)}


def load(path: Path):
    with open(path, encoding="utf-8") as f:
        data = OpenMuse.decode_rawdata(f.readlines())
    return data["EEG"], data.get("ACCGYRO")


def analyze(path: Path, skip: float) -> pd.DataFrame:
    eeg, imu = load(path)
    t = eeg["time"].to_numpy()
    b_bp, a_bp = butter(4, [1, 40], btype="band", fs=FS)
    b_n, a_n = iirnotch(50, 30, fs=FS)

    raw = {ch: eeg[ch].to_numpy(dtype=float) for ch in CHANNELS}
    filt = {ch: filtfilt(b_n, a_n, filtfilt(b_bp, a_bp, x - np.median(x))) for ch, x in raw.items()}

    if imu is not None and not imu.empty:
        gyro = np.sqrt(imu.GYRO_X**2 + imu.GYRO_Y**2 + imu.GYRO_Z**2).to_numpy()
        imu_t = imu["time"].to_numpy()
    else:
        gyro, imu_t = np.zeros(1), np.array([t[0]])

    n = int(WIN * FS)
    rows = []
    for start in np.arange(t[0] + skip, t[-1] - WIN, STEP):
        i0 = np.searchsorted(t, start)
        sl = slice(i0, i0 + n)
        g = gyro[(imu_t >= start) & (imu_t < start + WIN)]
        moving = g.size > 0 and g.mean() > MAX_GYRO
        row = {"time": start + WIN / 2 - t[0], "moving": moving}
        for ch in CHANNELS:
            r, x = raw[ch][sl], filt[ch][sl]
            bad = moving or ((r <= RAIL_LO) | (r >= RAIL_HI)).any() or np.ptp(x) > MAX_PTP
            row[f"{ch}_ok"] = not bad
            if bad:
                continue
            f, p = welch(x, FS, nperseg=FS * 2)
            bp = {k: p[(f >= lo) & (f < hi)].sum() for k, (lo, hi) in BANDS.items()}
            tot = p[(f >= 1) & (f < 40)].sum()
            row[f"{ch}_engage"] = bp["beta"] / (bp["alpha"] + bp["theta"])
            row[f"{ch}_load"] = bp["theta"] / bp["alpha"]
            row[f"{ch}_relax"] = bp["alpha"] / tot
        rows.append(row)
    df = pd.DataFrame(rows)

    # use only channels with at least 40 % clean windows
    good = [ch for ch in CHANNELS if df[f"{ch}_ok"].mean() >= 0.4]
    print(f"\n{path.name}: длительность {t[-1] - t[0]:.0f} с, окон {len(df)}")
    for ch in CHANNELS:
        mark = "используется" if ch in good else "отброшен (плохой контакт)"
        print(f"  {ch:9s} чистых окон {100 * df[f'{ch}_ok'].mean():5.1f} %  -> {mark}")
    if not good:
        print("  Нет ни одного пригодного канала — запись нужно повторить.")
        return None
    df.attrs["good"] = good
    return aggregate(df, good)


def aggregate(df: pd.DataFrame, channels: list) -> pd.DataFrame:
    """Average metrics over the selected channels."""
    for k in ("engage", "load", "relax"):
        df[k] = df[[f"{ch}_{k}" for ch in channels]].mean(axis=1)
    df["clean"] = df[[f"{ch}_ok" for ch in channels]].any(axis=1)
    df.attrs["channels"] = channels
    return df


METRICS = [
    ("engage", "Engagement / concentration\nβ / (α + θ)", "#d1495b"),
    ("load", "Mental workload\nθ / α", "#edae49"),
    ("relax", "Relaxation\nrelative α", "#00798c"),
]


def plot_timeline(df: pd.DataFrame, title: str, out: Path):
    fig, axes = plt.subplots(len(METRICS) + 1, 1, figsize=(11, 9), sharex=True)
    for ax, (key, label, color) in zip(axes, METRICS):
        s = df[key]
        ax.plot(df.time, s, color=color, alpha=0.25, lw=0.8)
        ax.plot(df.time, s.rolling(15, min_periods=5, center=True).median(), color=color, lw=2)
        ax.axhline(s.median(), color="gray", ls="--", lw=0.8)
        ax.set_ylabel(label, fontsize=9)
        ax.grid(alpha=0.3)
    ax = axes[-1]
    ax.fill_between(df.time, 0, 1, where=~df.clean, color="red", alpha=0.3, step="mid", label="artifact / no contact")
    ax.fill_between(df.time, 0, 1, where=df.moving, color="orange", alpha=0.5, step="mid", label="head movement")
    ax.set_yticks([])
    ax.set_ylabel("Signal quality", fontsize=9)
    ax.legend(loc="upper right", fontsize=8)
    ax.set_xlabel("Time, s")
    fig.suptitle(f"{title}  (channels: {', '.join(c[4:] for c in df.attrs['channels'])})")
    fig.tight_layout()
    fig.savefig(out, dpi=120, bbox_inches="tight")
    print(f"  график -> {out}")


def plot_compare(dfs: dict, out: Path):
    fig, axes = plt.subplots(1, len(METRICS), figsize=(12, 4.5))
    names = list(dfs)
    for ax, (key, label, color) in zip(axes, METRICS):
        vals = [dfs[n][key].dropna().to_numpy() for n in names]
        bp = ax.boxplot(vals, tick_labels=names, showfliers=False, patch_artist=True, widths=0.5)
        for patch in bp["boxes"]:
            patch.set_facecolor(color)
            patch.set_alpha(0.5)
        meds = [np.median(v) for v in vals]
        for i, m in enumerate(meds, 1):
            ax.text(i, m, f" {m:.2f}", va="bottom", fontsize=9)
        if len(meds) == 2 and meds[0]:
            ax.set_xlabel(f"change: {100 * (meds[1] / meds[0] - 1):+.0f} %")
        ax.set_title(label, fontsize=10)
        ax.grid(alpha=0.3, axis="y")
    fig.suptitle("Cognitive metrics comparison (median over clean windows)")
    fig.tight_layout()
    fig.savefig(out, dpi=120, bbox_inches="tight")
    print(f"\nСравнение -> {out}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+", type=Path, help="1 запись или 2 (до и после) / 1 recording or 2 (before and after)")
    ap.add_argument("--skip", type=float, default=30, help="отбросить первые N секунд (стабилизация), по умолчанию 30 / drop the first N seconds (settling), default 30")
    ap.add_argument("--labels", nargs="+", help="подписи для графиков, по умолчанию before и after / plot labels, default before and after")
    ap.add_argument("--no-show", action="store_true", help="не открывать окна / do not open plot windows")
    args = ap.parse_args()

    labels = args.labels or (["before", "after"] if len(args.files) == 2 else [f.stem for f in args.files])
    outdir = FIG_DIR
    outdir.mkdir(exist_ok=True)

    dfs, stems = {}, {}
    for f, lab in zip(args.files, labels):
        df = analyze(f, args.skip)
        if df is not None:
            dfs[lab], stems[lab] = df, f

    # when comparing — use only channels that are good in ALL recordings, otherwise the comparison is invalid
    if len(dfs) > 1:
        common = [ch for ch in CHANNELS if all(ch in d.attrs["good"] for d in dfs.values())]
        if not common:
            raise SystemExit("\nНет общего рабочего канала во всех записях — сравнение невозможно.")
        print(f"\nОбщие каналы для сравнения: {', '.join(c[4:] for c in common)}")
        for lab in dfs:
            dfs[lab] = aggregate(dfs[lab], common)

    for lab, df in dfs.items():
        f = stems[lab]
        plot_timeline(df, f"{lab} — {f.name}", outdir / f"{f.stem}_cognitive.png")
    if len(dfs) > 1:
        plot_compare(dfs, outdir / f"compare_{'_vs_'.join(f.stem for f in args.files)}.png")
    if not args.no_show:
        plt.show()


if __name__ == "__main__":
    main()
