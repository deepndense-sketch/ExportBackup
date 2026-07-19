# ExportBackup User Manual

ExportBackup helps you make a backup MP4 and separate audio exports from the active Premiere Pro sequence, then brings them back into the project and lines them up.

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

For re-backup, the queue checkboxes and merge selections are ignored. ExportBackup records the existing backup clips before export, uses the `_BACKUP` clip only as a video layer, keeps its attached audio muted, and exports audio only from the original source tracks. After replacement, the backup video, muted video audio, and separate audio files return to their recorded tracks and timeline start positions.

If Premiere Pro keeps an old backup file locked, the completed `_REBKP_TEMP` renders are left untouched. Run `Align Existing` on the export folder: it verifies that the TEMP files are finished, releases and removes every old backup reference from Premiere, renames the TEMP files to their correct names, and completes alignment.

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

The version button shows the installed version.

When a new update is available, the button changes color. Click it to update from GitHub.
