import argparse
import ctypes
import sys
import threading
import time

from process_audio_capture import ProcessAudioCapture, ProcessAudioCaptureError


try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except AttributeError:
    pass


def qpc_stamp():
    counter = ctypes.c_longlong()
    frequency = ctypes.c_longlong()
    ctypes.windll.kernel32.QueryPerformanceCounter(ctypes.byref(counter))
    ctypes.windll.kernel32.QueryPerformanceFrequency(ctypes.byref(frequency))
    return counter.value, frequency.value


def print_qpc(label):
    counter, frequency = qpc_stamp()
    print(f"{label}_qpc={counter} qpc_frequency={frequency} unix={time.time():.6f}", flush=True)


def watch_stdin(stop_event):
    for line in sys.stdin:
        command = line.strip().lower()
        if command in {"q", "quit", "stop"}:
            stop_event.set()
            return


def main():
    parser = argparse.ArgumentParser(description="Record one Windows process tree audio to WAV until stdin receives q.")
    parser.add_argument("output", help="Output WAV path")
    parser.add_argument("--pid", type=int, required=True, help="Target process id")
    args = parser.parse_args()

    if not ProcessAudioCapture.is_supported():
        raise RuntimeError("Windows process audio capture is not supported on this PC.")

    stop_event = threading.Event()
    stdin_thread = threading.Thread(target=watch_stdin, args=(stop_event,), daemon=True)
    stdin_thread.start()

    print(f"process audio capture pid: {args.pid}", flush=True)
    with ProcessAudioCapture(pid=args.pid, output_path=args.output) as capture:
        print_qpc("process audio capture starting")
        capture.start()
        print_qpc("process audio capture started")
        print("process audio capture started", flush=True)
        while not stop_event.is_set():
            time.sleep(0.1)
        capture.stop()

    print("process audio capture stopped", flush=True)


if __name__ == "__main__":
    try:
        main()
    except ProcessAudioCaptureError as exc:
        print(f"process audio capture error: {exc}", file=sys.stderr, flush=True)
        raise
