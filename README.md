# ExportBackup User Manual

ExportBackup helps you make a backup MP4 and separate audio exports from the active Premiere Pro sequence, then brings them back into the project and lines them up.

The `Queue Backup Exports` section has a visual `Hide`/`Show` control and always opens shown when the panel loads.

## Install

Run `deploy_extension.bat`.

It installs the panel here:

`%APPDATA%\Adobe\CEP\extensions\ExportBackup`

Restart Premiere Pro after installing.

## Before Backup

1. Open the Premiere Pro sequence you want to backup.
2. Set the sequence In and Out points.
3. If any audio track has Solo enabled, turn Solo off before exporting.
4. Open the ExportBackup panel.
5. Click `Choose Folder` and select the export folder.

## Backup

1. Choose `BACKUP TO`, or select `TO EMPTY TRACK`.
2. Choose `Premiere` or `Media Encoder`.
3. Choose `MP3` or `WAV`.
4. Click `Refresh` to load the track list.
5. Check the backup video and audio tracks you want to export.
6. Click `Backup`.

If backup files already exist, use `Re-backup`.

## Re-backup

Use `Re-backup` when the backup files are already in the project.

Re-backup exports only the checked backup video and audio items, replaces their old files, keeps the correct names, and aligns the new backup media back into the sequence. Unchecked backup items remain untouched.

Before export, each selected existing backup file is preserved under a unique `_REBKP_OLD_...` name and Premiere is relinked to that preserved file. The visible project-item name stays unchanged, so the sequence still shows the normal backup name. ExportBackup then frees the real backup filename and exports the replacement directly to it.

The selected backup video layer stays visible and is included in the MP4 render. Video tracks above the selected backup track are temporarily hidden, then every video track returns to its exact previous visibility state.

After every selected export is complete, ExportBackup releases the selected old Premiere references, imports the replacements, and restores their recorded tracks and timeline positions before deleting preserved old files. If Premiere or Windows still reports a genuine pending item, ExportBackup runs a full Align Existing pass every 3 seconds: it removes the currently aligned backup clips, imports the files again, restores their recorded positions, and checks cleanup again. A successful zero-remaining response stops retries and removes the JSON recovery map. The retry window also offers Align Existing for a manual pass.

Imported backup MP4 media uses an Orange label. Imported backup MP3, WAV, and merged audio media uses a Brown label.

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

Click `Change Export Presets` only if you need to change the preset files.

## Extra Options

- `Remove sequence markers`: removes sequence markers during backup.
- `Copy project file`: copies the Premiere project file to the backup folder.
- `Sort project files`: organizes imported backup files.

## Updates

The version button shows `Latest Version` when the installed release is current.

When a new update is available, the version button turns blue. Click it to update from GitHub.
