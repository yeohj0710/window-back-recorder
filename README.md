# Window Back Recorder

A small Windows recorder/controller for a selected visible window.

## What it does

- Lists visible top-level windows.
- Records the selected window with FFmpeg `gfxcapture` when available, with `gdigrab` as a fallback.
- Sends the target window behind other windows without minimizing it.
- Lets you bring the target window back to the front.
- Can compact the target window to the lower-right corner and restore its size.
- Can record loopback audio from a selected output device and mux it into the MP4 after stopping.
- Lets you toggle listening to captured loopback audio while recording.
- Lets you choose the save folder.
- Saves recordings with date/time names such as `2026-05-09_12-30-05.mp4`.

## Requirements

- Windows 10/11.
- Python 3.
- FFmpeg in `PATH`.
- Python dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Run

Double-click:

```bat
Start-WindowBackRecorder.cmd
```

Or from PowerShell:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WindowBackRecorder.ps1
```

## Important limits

- Do not minimize, hide, or close the target window while recording.
- The target window must keep rendering frames. Covering it with other windows is usually okay with `gfxcapture`; minimizing it can still stop or freeze frames.
- Some apps block capture or render protected content as black. Use built-in/host recording when an app deliberately blocks capture.
- Loopback audio records whatever Windows is actually rendering to the selected output device. If you mute the app internally or mute its Windows audio session, the recording may also be silent.
- To avoid hearing the app while still recording it, route the app to an unused/virtual output device in Windows Sound Mixer, choose that device's `Loopback:` source in this app, and leave `Listen to captured audio` off. Turn it on when you want to hear the content.

## Notes

Use this only where you have permission to record. The app does not join meetings or bypass account permissions; it records pixels/audio that the local Windows session is already rendering.
