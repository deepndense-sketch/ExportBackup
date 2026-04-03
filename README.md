# ExportBackup

A separate Adobe Premiere Pro CEP extension project for exporting the active sequence.

## Goal

This extension is intended to:

- work on the active sequence
- require sequence In and Out points before queueing
- queue an H.264 1080p MP4 backup to Adobe Media Encoder
- unmute all audio tracks for the MP4 queue job
- optionally queue each audio track as a separate MP3 and/or WAV export

## Project Structure

```text
ExportBackup/
  CSXS/manifest.xml
  js/CSInterface.js
  js/main.js
  jsx/export.jsx
  index.html
  .gitignore
```

## Current Status

This project now uses Premiere Pro's `app.encoder.encodeSequence()` workflow to send jobs to Adobe Media Encoder.

## Presets Used

- MP4 video defaults to: `D:\Work\Tools\ExportBackup\presets\1080 AIR.epr`
- The panel can be pointed at a different video `.epr` preset, and it remembers that choice until changed.
- MP3 audio: `D:\Work\Tools\ExportBackup\presets\mp3.epr`
- WAV audio: `C:\Program Files\Adobe\Adobe Media Encoder 2026\MediaIO\systempresets\3F3F3F3F_57415645\Waveform Audio 48kHz 16-bit.epr`

## Important Limitation

Premiere Pro's official scripting API supports muting audio tracks, but does not document track solo control. This extension can clear mutes for the MP4 queue job, but any solo state should be cleared manually in Premiere before running the tool.
