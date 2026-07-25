# ExportBackup User Manual

ExportBackup helps you make a backup MP4 and separate audio exports from the active Premiere Pro sequence, then brings them back into the project and lines them up.

The `Queue Backup Exports` section has a visual `Hide`/`Show` control and always opens shown when the panel loads.

## Install

Run `deploy_extension.bat`.

It installs the panel here:

`C:\Program Files (x86)\Common Files\Adobe\CEP\extensions\ExportBackup`

Restart Premiere Pro after installing.

## Before Backup

1. Open the Premiere Pro sequence you want to backup.
2. Set the sequence In and Out points.
3. If any audio track has Solo enabled, turn Solo off before exporting.
4. Open the ExportBackup panel.
5. Click `Choose Folder` and select the export folder.

## Backup

1. Choose `BACKUP TRACK`.
2. Choose `Premiere` or `Media Encoder`.
3. Choose `MP3` or `WAV`.
4. Click `Refresh` to load the track list.
5. Check the backup video and audio tracks you want to export.
6. Click `Backup`.

If backup files already exist, use `Re-backup`.

## Re-backup

Use `Re-backup` when the backup files are already in the project.

Re-backup exports new backup files, replaces the old backup files, keeps the correct names, and aligns the new backup media back into the sequence.

For re-backup, the queue checkboxes and merge selections are ignored. ExportBackup records the existing backup clips before export, exports audio from the original source tracks, and restores every video track to its exact pre-export visibility state. After replacement, the backup video, its audio, and separate audio files return to their recorded tracks and timeline start positions. The backup MP4 audio is left as the only audible audio track.

Before any `_REBKP_TEMP` file is renamed, ExportBackup removes old final-file clips and automatically imported TEMP-file clips from every sequence. Because Premiere exposes deletion for bins rather than ordinary footage items, the targeted backup items are moved into a temporary cleanup bin and that bin is deleted. ExportBackup then verifies by project-item identity that no online or offline target remains anywhere in the project. If Premiere Pro keeps an item or file locked, the completed TEMP renders are left untouched. Run `Align Existing` on the export folder to retry the verified cleanup, rename, import, and alignment transaction.

## Align Existing

Use `Align Existing` when the files were already exported and you only want to import and align them.

Click `Choose Folder` first, then click `Align Existing`.

## Merge Audio Tracks

Use the Merge checkboxes on the right side of the track list.

1. Tick the audio tracks you want to merge together.
2. Click `Merge Selection`.
3. Merged tracks fade and stay locked as one group.

Other audio tracks can still be selected normally.

## Presets

The panel uses presets from the `presets` folder inside the extension.

Click `Show Presets` only if you need to change the preset files.

## Extra Options

- `Remove sequence markers`: removes sequence markers during backup.
- `Copy project file`: copies the Premiere project file to the backup folder.
- `Sort project files`: organizes imported backup files.

## Updates

The version button shows `Latest Version` when the installed release is current.

When a new update is available, the version button turns blue. Click it to update from GitHub.
