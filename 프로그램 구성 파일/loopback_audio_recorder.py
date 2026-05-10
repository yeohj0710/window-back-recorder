import argparse
import ctypes
import json
import sys
import threading
import time
import wave

import numpy as np
import soundcard as sc


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


def watch_stdin(stop_event, monitor_event):
    for line in sys.stdin:
        command = line.strip().lower()
        if command in {"q", "quit", "stop"}:
            stop_event.set()
            return
        if command in {"monitor on", "listen on", "unmute", "sound on"}:
            monitor_event.set()
            print("monitor on", flush=True)
        elif command in {"monitor off", "listen off", "mute", "sound off"}:
            monitor_event.clear()
            print("monitor off", flush=True)


def loopback_microphones():
    microphones = []
    for microphone in sc.all_microphones(include_loopback=True):
        if str(microphone).startswith("<Loopback "):
            microphones.append(microphone)
    return microphones


def list_loopbacks():
    devices = []
    default_name = None
    try:
        default_speaker = sc.default_speaker()
        if default_speaker is not None:
            default_name = default_speaker.name
    except Exception:
        default_name = None

    for microphone in loopback_microphones():
        devices.append(
            {
                "name": microphone.name,
                "description": str(microphone),
                "default": microphone.name == default_name,
            }
        )
    print(json.dumps(devices, ensure_ascii=False), flush=True)


def pick_loopback_microphone(source_name=None):
    if source_name:
        for microphone in loopback_microphones():
            if microphone.name == source_name:
                return microphone, None
        raise RuntimeError(f"Loopback source not found: {source_name}")

    speaker = sc.default_speaker()
    if speaker is None:
        raise RuntimeError("No default speaker found.")

    # soundcard exposes WASAPI output capture as loopback microphones.
    return sc.get_microphone(speaker.name, include_loopback=True), speaker


def float_to_int16(samples):
    samples = np.nan_to_num(samples, copy=False)
    samples = np.clip(samples, -1.0, 1.0)
    return (samples * 32767.0).astype("<i2", copy=False)


def record_loopback(output_path, source_name, samplerate, channels, block_frames, monitor_initially):
    stop_event = threading.Event()
    monitor_event = threading.Event()
    if monitor_initially:
        monitor_event.set()

    stdin_thread = threading.Thread(target=watch_stdin, args=(stop_event, monitor_event), daemon=True)
    stdin_thread.start()

    microphone, speaker = pick_loopback_microphone(source_name)
    if speaker:
        print(f"loopback speaker: {speaker.name}", flush=True)
    print(f"loopback microphone: {microphone.name}", flush=True)
    print("monitor on" if monitor_event.is_set() else "monitor off", flush=True)

    with wave.open(output_path, "wb") as wav_file:
        wav_file.setnchannels(channels)
        wav_file.setsampwidth(2)
        wav_file.setframerate(samplerate)

        output_speaker = sc.default_speaker()
        can_monitor = output_speaker is not None and microphone.name != output_speaker.name
        if not can_monitor:
            print("monitor playback disabled for this source to avoid recording feedback", flush=True)
        with microphone.recorder(samplerate=samplerate, channels=channels) as recorder:
            with output_speaker.player(samplerate=samplerate, channels=channels) as player:
                print_qpc("loopback audio capture started")
                print("loopback audio capture started", flush=True)
                while not stop_event.is_set():
                    data = recorder.record(numframes=block_frames)
                    if data.ndim == 1:
                        data = data.reshape((-1, 1))

                    if data.shape[1] < channels:
                        data = np.repeat(data, channels, axis=1)
                    elif data.shape[1] > channels:
                        data = data[:, :channels]

                    wav_file.writeframes(float_to_int16(data).tobytes())

                    if monitor_event.is_set() and can_monitor:
                        player.play(data)

    print("loopback audio stopped", flush=True)


def main():
    parser = argparse.ArgumentParser(description="Record default speaker loopback audio to WAV until stdin receives q.")
    parser.add_argument("--list-json", action="store_true", help="List loopback sources as JSON and exit")
    parser.add_argument("output", nargs="?", help="Output WAV path")
    parser.add_argument("--source", default=None, help="Exact soundcard loopback source name")
    parser.add_argument("--samplerate", type=int, default=48000)
    parser.add_argument("--channels", type=int, default=2)
    parser.add_argument("--block-frames", type=int, default=2048)
    parser.add_argument("--monitor-on", action="store_true", help="Play captured audio to the default speaker while recording")
    args = parser.parse_args()

    if args.list_json:
        list_loopbacks()
        return

    if not args.output:
        parser.error("output is required unless --list-json is used")

    record_loopback(args.output, args.source, args.samplerate, args.channels, args.block_frames, args.monitor_on)


if __name__ == "__main__":
    main()
