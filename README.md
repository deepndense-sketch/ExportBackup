# ExportBackup

A separate Adobe Premiere Pro CEP extension project for exporting the active sequence.

## Goal

This extension is intended to:

- work on the active sequence
- require sequence In and Out points before queueing
- queue an H.264 1080p MP4 backup to Adobe Media Encoder
- unmute all audio tracks for the MP4 queue job
- optionally queue each audio track as a separate MP3 and/or WAV export
- save files using `SequenceName_BACKUP` and `SequenceName_TrackN` naming
- align matching exported files back onto the active sequence from sequence start

## Project Structure

```text
ExportBackup/
  CSXS/manifest.xml
  deploy_extension.bat
  js/CSInterface.js
  js/main.js
  jsx/export.jsx
  index.html
  version.json
  .gitignore
```

## Current Status

This project now uses Premiere Pro's `app.encoder.encodeSequence()` workflow to send jobs to Adobe Media Encoder.

The panel now also includes:

- local version display from `version.json`
- GitHub update check against the repository `main` branch
- an `Update From GitHub` action which downloads the latest `main` branch ZIP and installs it into Adobe's CEP extensions folder
- manual folder-based alignment for files named like `SequenceName_BACKUP.mp4` and `SequenceName_Track1.wav`
- a deployment batch file for copying the extension into Adobe's CEP extensions folder

The current alignment workflow is intentionally manual:

- choose a folder
- match `SequenceName_BACKUP.*`
- match `SequenceName_TrackN.*`
- place the backup video at sequence start on a chosen video track
- place the backup video's own audio at sequence start on a chosen audio track
- place the other audio files from a chosen start track upward automatically, one track at a time

## Presets Used

- MP4 video defaults to: `D:\Work\Tools\ExportBackup\presets\1080 AIR.epr`
- The panel can be pointed at a different video `.epr` preset, and it remembers that choice until changed.
- MP3 audio: `D:\Work\Tools\ExportBackup\presets\mp3.epr`
- WAV audio: `C:\Program Files\Adobe\Adobe Media Encoder 2026\MediaIO\systempresets\3F3F3F3F_57415645\Waveform Audio 48kHz 16-bit.epr`

## Important Limitation

Premiere Pro's official scripting API supports muting audio tracks, but does not document track solo control. This extension can clear mutes for the MP4 queue job, but any solo state should be cleared manually in Premiere before running the tool.

Track deletion is also not supported reliably through Premiere's official scripting API, so alignment works by placing files onto user-selected target tracks rather than deleting existing tracks.

## Updating Installed Extensions

The intended installed location is:

- `C:\Program Files (x86)\Common Files\Adobe\CEP\extensions\ExportBackup`

The panel checks GitHub on load by comparing local `version.json` against:

- `https://raw.githubusercontent.com/deepndense-sketch/ExportBackup/main/version.json`

When `Update From GitHub` is used, the updater script downloads:

- `https://github.com/deepndense-sketch/ExportBackup/archive/refs/heads/main.zip`

It then extracts that ZIP and mirrors the contents into the installed CEP extension folder:

- `C:\Program Files (x86)\Common Files\Adobe\CEP\extensions\ExportBackup`

Because that destination is under `Program Files`, Windows may prompt for administrator permission.
