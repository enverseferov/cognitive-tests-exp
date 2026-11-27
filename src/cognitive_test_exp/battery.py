"""Check the Muse headset battery level.

Usage (from the project folder):
    uv run muse-battery

If a recording is already running (live-cognitive), the level is read from its stream —
the recording is not interrupted. Otherwise the script briefly connects to the headset
and disconnects. Muse sends the battery level about every 5 seconds, so the answer
takes 5–20 s.
"""

import argparse
import shutil
import subprocess
import sys
import time

from cognitive_test_exp.live_cognitive import MUSE_ADDRESS, find_stream

WAIT_S = 20  # how long to wait for a battery value, s


def main() -> None:
    ap = argparse.ArgumentParser(description="Заряд гарнитуры Muse / Muse headset battery level")
    ap.add_argument("--address", default=MUSE_ADDRESS)
    args = ap.parse_args()

    from mne_lsl.lsl import StreamInlet

    # if a stream is already running (recording in progress) — just read it, do not open a second connection
    running = subprocess.run(["pgrep", "-f", "OpenMuse stream"], capture_output=True, text=True).stdout.strip()
    stream_proc = None
    print("Ищу гарнитуру...")
    sinfo = find_stream("Battery", 10 if running else 2)
    if sinfo is None and running:
        sys.exit("Запись идёт, но поток заряда не найден. Попробуйте ещё раз через несколько секунд.")
    try:
        if sinfo is None:
            cmd = [shutil.which("OpenMuse") or "OpenMuse", "stream", "--address", args.address]
            stream_proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            deadline = time.time() + 45
            while sinfo is None:
                if stream_proc.poll() is not None:
                    sys.exit("Не удалось подключиться — проверьте, что Muse включена, рядом "
                             "и не подключена к телефону.")
                if time.time() > deadline:
                    sys.exit("Muse не подключилась за 45 с.")
                sinfo = find_stream("Battery", 1)

        inlet = StreamInlet(sinfo)
        inlet.open_stream(timeout=10)
        deadline = time.time() + WAIT_S
        level = None
        while time.time() < deadline:
            data, _ = inlet.pull_chunk(timeout=1.0)
            if len(data):
                level = float(data[-1][0])
                break
        inlet.close_stream()

        if level is None:
            sys.exit(f"Гарнитура подключена, но заряд не пришёл за {WAIT_S} с. Попробуйте ещё раз.")
        if level >= 50:
            note = "достаточно для записи"
        elif level >= 20:
            note = "хватит на короткую запись, лучше зарядить"
        else:
            note = "мало — зарядите перед записью"
        print(f"\nЗаряд Muse: {level:.0f} %  ({note})")
    finally:
        if stream_proc is not None:
            stream_proc.terminate()
            try:
                stream_proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                stream_proc.kill()


if __name__ == "__main__":
    main()
