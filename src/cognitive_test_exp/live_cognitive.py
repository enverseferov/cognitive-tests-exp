"""Live cognitive load plot from Muse (concentration and workload only).

Usage (in the activated environment):
    uv run live-cognitive                 # plot; starts the Muse stream itself if it is not running
    uv run live-cognitive --record        # + raw Muse recording to data/muse_<date_time>.txt
    uv run live-cognitive --label before  # label (shown on the plot and used in file names)

While running:
    space  — add a marker (e.g. start/end of a test)
    close the window or press Ctrl+C — stop

Results are saved to data/cognitive_<label>_<date_time>.csv (one value per second)
and figures/cognitive_<label>_<date_time>.png (final plot).

Metrics are computed over a 4-second signal window (with ~2 s delay), using only the
forehead electrodes AF7/AF8 (the most reliable ones on Muse). Segments without contact
or with head movement are skipped.

For final before/after numbers use plot-cognitive on the raw recordings (--record):
the live values can differ from the offline analysis.
"""

import argparse
import csv
import shutil
import subprocess
import sys
import time
from collections import deque
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.animation import FuncAnimation
from scipy.signal import butter, iirnotch, sosfiltfilt, tf2sos, welch

# project root: data and figures are always saved to <project>/data and <project>/figures
ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = ROOT / "data"
FIG_DIR = ROOT / "figures"

MUSE_ADDRESS = "DC6EB9D3-0CE2-B804-8FE1-9EE1DDEA309A"  # MuseS-FF61
FS = 256
WIN = 4.0  # analysis window, s
BUF = 10.0  # filtering buffer, s
EDGE = 2.0  # margin from buffer edges (the filter distorts the signal there), s = display delay
CH_NAMES = ["TP9", "AF7", "AF8", "TP10"]  # order of the first Muse-EEG stream channels
USE = [1, 2]  # AF7, AF8
RAIL_LO, RAIL_HI = 0.5, 1449.5
MAX_PTP = 500.0  # µV
MAX_GYRO = 20.0  # °/s
SMOOTH = 10  # line smoothing, points (seconds)

SOS = np.vstack([butter(4, [1, 40], btype="band", fs=FS, output="sos"), tf2sos(*iirnotch(50, 30, fs=FS))])


def band_metrics(x):
    f, p = welch(x, FS, nperseg=FS * 2)
    th = p[(f >= 4) & (f < 8)].sum()
    al = p[(f >= 8) & (f < 13)].sum()
    be = p[(f >= 13) & (f < 30)].sum()
    return be / (al + th), th / al


def find_stream(stype, timeout):
    from mne_lsl.lsl import resolve_streams

    s = resolve_streams(timeout=timeout, stype=stype)
    return s[0] if s else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--record", action="store_true", help="также писать сырую запись Muse в data/ / also save the raw Muse recording to data/")
    ap.add_argument("--label", default="", help="подпись записи (на графике), например before или after / recording label (on the plot), e.g. before or after")
    ap.add_argument("--address", default=MUSE_ADDRESS)
    args = ap.parse_args()

    from mne_lsl.lsl import StreamInlet

    stamp = f"{datetime.now():%Y%m%d_%H%M%S}"
    name = f"cognitive_{args.label + '_' if args.label else ''}{stamp}"
    DATA_DIR.mkdir(exist_ok=True)
    FIG_DIR.mkdir(exist_ok=True)

    # connect to the stream (or start our own) ---
    stream_proc = None
    other = subprocess.run(["pgrep", "-f", "OpenMuse stream"], capture_output=True, text=True).stdout.strip()
    print("Ищу поток Muse...")
    sinfo = find_stream("EEG", 10)
    if sinfo is None and other:
        sys.exit("OpenMuse stream уже запущен, но поток ЭЭГ не найден. Подождите немного и запустите снова\n"
                 "(второе подключение к гарнитуре не запускаю, чтобы не оборвать текущую запись).")
    if sinfo is not None and args.record:
        print("Поток Muse уже запущен другим процессом — сырая запись (--record) не будет сделана этим скриптом.")
    if sinfo is None:
        cmd = [shutil.which("OpenMuse") or "OpenMuse", "stream", "--address", args.address]
        if args.record:
            raw_file = DATA_DIR / f"muse_{stamp}.txt"
            cmd += ["--record", str(raw_file)]
            print(f"Сырая запись в {raw_file}")
        print("Подключаюсь к Muse...")
        stream_proc = subprocess.Popen(cmd)
        deadline = time.time() + 45
        while sinfo is None:
            if stream_proc.poll() is not None:
                sys.exit("Поток завершился — проверьте, что Muse включена, рядом и не подключена к телефону.")
            if time.time() > deadline:
                stream_proc.terminate()
                sys.exit("Muse не подключилась за 45 с.")
            sinfo = find_stream("EEG", 1)

    eeg_in = StreamInlet(sinfo)
    eeg_in.open_stream(timeout=10)
    imu_info = find_stream("ACCGYRO", 2)
    imu_in = None
    if imu_info is not None:
        imu_in = StreamInlet(imu_info)
        imu_in.open_stream(timeout=10)

    buf = deque(maxlen=int(BUF * FS))
    gyro = deque(maxlen=int(WIN * 52))

    t0 = time.time()
    T, ENG, LOAD, OK = [], [], [], []
    marks = []  # (time, number)
    status = {"text": "waiting for signal..."}

    csv_path = DATA_DIR / f"{name}.csv"
    fcsv = open(csv_path, "w", newline="", encoding="utf-8")
    w = csv.writer(fcsv)
    w.writerow(["time_s", "clock", "concentration", "load", "clean", "AF7_ok", "AF8_ok", "marker"])

    # plot window ---
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 6.5), sharex=True)
    fig.canvas.manager.set_window_title(f"Cognitive load {args.label}")
    l1r, = ax1.plot([], [], color="#d1495b", alpha=0.25, lw=0.8)
    l1, = ax1.plot([], [], color="#d1495b", lw=2.2)
    l2r, = ax2.plot([], [], color="#edae49", alpha=0.3, lw=0.8)
    l2, = ax2.plot([], [], color="#c47f00", lw=2.2)
    ax1.set_ylabel("Concentration\nβ / (α + θ)")
    ax2.set_ylabel("Mental workload\nθ / α")
    ax2.set_xlabel("Time since start, s   (space = marker)")
    for ax in (ax1, ax2):
        ax.grid(alpha=0.3)
    title = fig.suptitle("")

    def smooth(v):
        a = np.array(v, dtype=float)
        if len(a) < 2:
            return a
        out = np.full_like(a, np.nan)
        for i in range(len(a)):
            seg = a[max(0, i - SMOOTH + 1): i + 1]
            seg = seg[~np.isnan(seg)]
            if seg.size >= 3:
                out[i] = np.median(seg)
        return out

    def on_key(ev):
        if ev.key == " ":
            n = len(marks) + 1
            tm = time.time() - t0
            marks.append((tm, n))
            for ax in (ax1, ax2):
                ax.axvline(tm, color="k", ls=":", lw=1)
            ax1.text(tm, 1.0, f" {n}", transform=ax1.get_xaxis_transform(), va="top", fontsize=9)
            w.writerow([f"{tm:.1f}", f"{datetime.now():%H:%M:%S}", "", "", "", "", "", n])
            print(f"Метка {n} на {tm:.0f} с")

    fig.canvas.mpl_connect("key_press_event", on_key)

    def update(_):
        data, _ts = eeg_in.pull_chunk(timeout=0.0)
        if len(data):
            buf.extend(np.asarray(data)[:, :4])
        if imu_in is not None:
            g, _ = imu_in.pull_chunk(timeout=0.0)
            if len(g):
                g = np.asarray(g)
                gyro.extend(np.sqrt((g[:, 3:6] ** 2).sum(axis=1)))
        if len(buf) < (WIN + 2 * EDGE) * FS:
            title.set_text(status["text"])
            return

        arr = np.asarray(buf, dtype=float)
        n, m = int(WIN * FS), int(EDGE * FS)
        sl = slice(len(arr) - m - n, len(arr) - m)  # window from the middle of the buffer
        moving = len(gyro) > 0 and float(np.mean(gyro)) > MAX_GYRO
        engs, loads, ch_ok = [], [], {}
        for i in range(4):
            raw = arr[sl, i]
            x = sosfiltfilt(SOS, arr[:, i] - np.median(arr[:, i]))[sl]
            ok = not moving and not ((raw <= RAIL_LO) | (raw >= RAIL_HI)).any() and np.ptp(x) <= MAX_PTP
            ch_ok[CH_NAMES[i]] = ok
            if ok and i in USE:
                e, l = band_metrics(x)
                engs.append(e)
                loads.append(l)
        tm = time.time() - t0
        eng = float(np.mean(engs)) if engs else np.nan
        load = float(np.mean(loads)) if loads else np.nan
        T.append(tm)
        ENG.append(eng)
        LOAD.append(load)
        OK.append(bool(engs))
        w.writerow([f"{tm:.1f}", f"{datetime.now():%H:%M:%S}",
                    "" if np.isnan(eng) else f"{eng:.3f}", "" if np.isnan(load) else f"{load:.3f}",
                    int(bool(engs)), int(ch_ok["AF7"]), int(ch_ok["AF8"]), ""])
        fcsv.flush()

        se, sl = smooth(ENG), smooth(LOAD)
        l1r.set_data(T, ENG)
        l1.set_data(T, se)
        l2r.set_data(T, LOAD)
        l2.set_data(T, sl)
        for ax, s in ((ax1, se), (ax2, sl)):
            ax.set_xlim(0, max(60, tm + 5))
            vals = np.array(s)[~np.isnan(s)]
            if vals.size:
                ax.set_ylim(0, max(1.0, np.percentile(vals, 98) * 1.3))

        contact = "  ".join(f"{k}:{'✓' if v else '✗'}" for k, v in ch_ok.items())
        cur_e = f"{se[-1]:.2f}" if not np.isnan(se[-1]) else "—"
        cur_l = f"{sl[-1]:.2f}" if not np.isnan(sl[-1]) else "—"
        warn = "  ⚠ movement" if moving else ("" if engs else "  ⚠ no forehead contact")
        title.set_text(f"{args.label}  {tm:5.0f} s   concentration {cur_e}   workload {cur_l}\n"
                       f"contact  {contact}{warn}")

    ani = FuncAnimation(fig, update, interval=1000, cache_frame_data=False)  # noqa: F841
    try:
        plt.show()
    except KeyboardInterrupt:
        pass
    finally:
        fcsv.close()
        if T:
            fig.savefig(FIG_DIR / f"{name}.png", dpi=120, bbox_inches="tight")
            ok = [i for i, v in enumerate(OK) if v]
            if ok:
                print(f"\nИтог ({len(ok)} из {len(T)} с с чистым сигналом):")
                print(f"  концентрация (медиана) {np.nanmedian(ENG):.2f}")
                print(f"  нагрузка     (медиана) {np.nanmedian(LOAD):.2f}")
            print(f"Значения -> {csv_path}\nГрафик   -> {FIG_DIR / (name + '.png')}")
        eeg_in.close_stream()
        if imu_in is not None:
            imu_in.close_stream()
        if stream_proc is not None:
            stream_proc.terminate()
            try:
                stream_proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                stream_proc.kill()
            print("Поток Muse остановлен.")


if __name__ == "__main__":
    main()
