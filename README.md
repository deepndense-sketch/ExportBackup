# Backup Project User Manual

**Current release 5.0.1:** The audio selection list follows the marked In/Out range. Tracks shared by an existing main-section backup and teaser music only appear when source audio overlaps the selected section. Reopen the panel or restart Premiere after updating.

**Since 5.0.0:** Includes the range-aware Backup/Re-backup fixes, checked-track-only re-exports, verified file renaming and relinking, readable status dialogs, and Premiere Pro as the default exporter. Update dialogs show **What's new** automatically from the upcoming release's `version.json.notes`, with scrollable text and an empty-notes fallback. This notes display becomes available after installing v5; older installed versions keep their existing update dialog until then.

Release authors: update `version.json.notes` with short, plain-language changes for every release. The update dialog reads these notes from GitHub automatically; no dialog code change is needed for each new version. Automatic detection of short/incomplete MP4 exports is not included in v5.

**Since 4.10.1:** Premiere Pro is selected every time the panel opens, even if Media Encoder was used in the previous session. Media Encoder remains available as a manual choice. Includes the range and partial Re-backup changes below. Reopen the panel or restart Premiere after installing.

**Since 4.10.0:** New Backup captures the sequence In/Out range and imports MP4 and audio at that In point. Re-backup uses the existing backup's original range and position, even if the current marks now surround a teaser; current marks are restored afterward. Align Existing keeps captured timeline positions. Per-file export ranges are stored in the recovery manifest and local ownership history. If an older backup exists only in the Project panel and its original range cannot be established, put it back at the intended position before Re-backup. Restart Premiere after installing.

**Since 4.9.4:** Partial Re-backup exports only checked items. Video-only keeps existing audio backups; selecting one audio track exports only that track. Unchecked merged groups are excluded, and selecting one member exports it separately while retaining the old merged backup containing unchecked tracks. Unchecked backups keep their files and track placements. A track-list refresh before exporting preserves existing checkbox choices and leaves newly discovered tracks unchecked. Full Re-backup continues to rebuild the complete chosen layout.

**Since 4.9.3:** Successful completion shows a green checkmark and a short heading; use **Show details** to read the log. Errors use a red indicator and warnings use yellow, with details expanded. Alignment logs show native filenames with normal spaces instead of `%20`.

**Since 4.9.2:** Align Existing renames verified backup files in Explorer to the current sequence name, preserving `_BACKUP` / `_TrackN` suffixes, and relinks Premiere project items and timeline clips. Older unrecorded MP4, WAV, MP3 and merged audio backups require one-time ownership confirmation; **Keep names** continues alignment without renaming them. Media recorded as another sequence's backup is excluded. Ownership uses the project path, stable sequence ID, and output path. Only aligned clips starting at zero and matching the full sequence duration within 0.04 seconds qualify. Destination files are never overwritten. Relink failures trigger filename rollback; locked files and any recovery problems are reported. Successful renames update ownership and recovery paths. Ownership history is local and does not transfer with copied media. Confirmation and completion dialogs wrap long names, scroll, and can be resized inside the panel.

The Solo reminder sits slightly closer to Tracks to back up, without moving the heading or other controls. The Path dropdown includes **Parent Folder**: `D:\Show\Project\Edit.prproj` exports into `D:\Show\BACKUP`, creating the folder if needed or reusing it. Update availability appears as a notice and a popup with **Later** and **Update now**. Update now installs directly into CEP; save your project first and restart Premiere after installation. Updates are blocked while exporting or collecting media.

Re-backup accepts the selected track containing a single existing backup MP4 even if its duration differs from the current sequence. It keeps that video track visible to reuse the rendered backup and hides only higher video tracks. Recorded backup audio is muted so source audio is exported. The old clip remains until the replacement is ready, and original track visibility is restored afterward. Normal Backup still requires an empty target.

Since **4.7.4**, a one-time repair is available for older installations with the broken updater: close Premiere, extract the repair ZIP, and double-click `repair-update.cmd`. Keep `repair-update.ps1` beside it. The repair downloads the latest GitHub package, verifies copied files, and moves confirmed duplicate installations outside CEP. Open Premiere afterward and use the plugin's Update button for future updates. Includes the 4.7.3 updater path fix and 4.7.2 backup recognition fix. Unrelated occupied video tracks remain protected.

Backup Project includes Export Backup and Project Collector. Export Backup makes a backup MP4 and separate audio exports from the active Premiere Pro sequence, then brings them back into the project and lines them up.

Version **4.7.0** keeps Backup, Re-backup, and Align Existing visible, with More settings expanded by default. Project Root exports go into a `BACKUP` folder. Re-backup and Align Existing use discovered backup locations independently of the selected new-backup destination. Align Existing also cleans up recognized obsolete backup files when they are unused and a replacement exists.

## Project Collector tab

The panel includes separate **Export Backup** and **Project Collector** tabs. Each keeps its own settings and actions; switching tabs does not export or copy anything. Collector is bundled in `collector/` and is deployed with ExportBackup. Use the Export Backup version button to update the combined package.

Collector uses **Premiere track locks**: lock a track to exclude its media. While Collector is visible, a lightweight check updates the track display about once per second; copying pauses these checks and reads fresh locks before starting. Added sequences appear as full-name tabs with individual remove buttons. Switching tabs only changes the visible track panel; all added sequences remain selected for collection. File counts include media inside nested sequences. Use Refresh after other sequence edits.

Choose **Show / Category** from Collector's top Path dropdown to preview the destination from the saved project filename. Collector uses `Z:\2017\_SMTV2 PROJECT FOLDER\@ SHOWS`, then the built-in category (such as `WOW`), then the project filename with the category moved to the front, without `.prproj`. For example, `3279 3280 WOW Socrates_Is virtue known or learned.prproj` collects into `@ SHOWS\WOW\WOW 3279 3280 Socrates_Is virtue known or learned`. MAIN/INTRO follow this same rule; there is no episode-range lookup. Preview does not access the network. The copy action creates the project folder and media subfolders, and reports actual filesystem errors if the destination cannot be written. Choose **Manual Path** to select a destination yourself.

Both tabs use the outer panel scrollbar. Collector expands to fit its content and has only a Refresh button above its controls.

Collector also offers **Project root**, which copies directly into the directory containing the saved `.prproj`, without category validation or an extra project-name folder. **Manual folder** keeps the chosen destination plus project-name folder. The destination mode is saved, and destination controls are locked during copying.

## Install

Run `deploy_extension.bat`.

It installs the panel here:

`%APPDATA%\Adobe\CEP\extensions\Backup Project`

Restart Premiere Pro after installing.

## Before Backup

1. Save the Premiere project with one configured category and at least one number somewhere in its filename. They may appear in either order, so `PE 3226.prproj`, `3226 PE.prproj`, and `Pierre 3226 WOW Gassendi.prproj` are valid.
2. Open the Premiere Pro sequence you want to backup.
3. Set the sequence In and Out points.
4. If any audio track has Solo enabled, turn Solo off before exporting.
5. Open the Backup Project panel and choose `FTP / Category`, `Project Root`, or `Manual Path`.

`FTP / Category` finds the category anywhere in the Premiere project filename, matches it to a folder under `Y:\@ Backup` while ignoring the folder's leading `@` and trailing `BACKUP`, and uses one reusable child folder with a normalized `CATEGORY NUMBER(S) optional name` format. For example, `3226 WOW Pierre Gassendi.prproj` is routed to `Y:\@ Backup\@ WOW BACKUP\WOW 3226 Pierre Gassendi`. The category-and-number naming requirement is checked only in this mode. `Project Root` and `Manual Path` accept any saved Premiere project name. Choosing a path switches to Manual Path; selecting FTP / Category again overrides that manual path with the category-resolved path.

Projects containing `MAIN` or `INTRO` use grouped episode routing for any episode number. ExportBackup finds an existing range container whose category and inclusive number range match, such as `WOW 3240-3251_Italy_19990522_SM PL`, then creates or reuses a number-only episode folder inside it. Thus `3250 WOW MAIN.prproj` and `3250 WOW INTRO.prproj` both use `...\WOW 3240-3251_Italy_19990522_SM PL\3250`. ExportBackup never creates the range container; if none matches, it asks the user to create the category project root folder first.

Use `Copy Existing Backups` to automatically find already-exported backup media used by the active sequence and copy it into the currently resolved rule-based destination. No source-folder selection is required. The plugin recognizes active-sequence filenames and duration-matched legacy names. It preserves the originals, does not change Premiere links, verifies copied file sizes, and refuses to overwrite a different existing destination file.

`Project Root` exports beside the saved `.prproj` file.

Use the Category Manager in the panel to add a category or delete the selected category. Changes are saved locally and immediately control both project-name validation and FTP category-folder matching. `SM URGENT MESSAGES`, `TRIBUTE`, and `QYP` are not included in the default category list.

## Backup

1. Choose `BACKUP TO`, or select `TO EMPTY TRACK`.
2. Choose `Premiere` or `Media Encoder`.
3. Choose `MP3` or `WAV`.
4. Click `Refresh` to load the track list.
5. Check the backup video and audio tracks you want to export.
6. Click `Backup`.

If backup files already exist, use `Re-backup`.

Backup-like filenames alone do not mean this sequence has been backed up. A candidate must be the only clip on its track and its timeline duration must match the full sequence or current In/Out range within 0.2 seconds. Short excerpts, mixed source tracks, and empty renamed tracks stay ordinary source tracks.

Before a new Backup, matching candidates are listed by their actual names, paths, and tracks. This includes old names after a sequence rename. You can continue with a new Backup or cancel to review them. Different-name candidates remain selectable as source media until you choose how to handle them. A Premiere reference to a missing output is reported separately; only an actual existing destination file blocks a new Backup. Choose an empty video track for the new backup. Alignment cleanup targets the replacement paths and preserves unrelated backup-named source clips.

## Re-backup

Use `Re-backup` when the backup files are already in the project. Existing merge groups load automatically, and Re-backup always makes the current queue checkboxes and merge groups authoritative. Use `Clear Merges` only when you want to rebuild the grouping. Replaced, renamed, unselected, or otherwise obsolete backup audio is removed from Premiere and deleted after the new files are aligned.

Re-backup exports only the checked backup video and audio items, replaces their old files, keeps the correct names, and aligns the new backup media back into the sequence. Unchecked backup items remain untouched.

Before export, each selected existing backup file is preserved under a unique `_REBKP_OLD_...` name and Premiere is relinked to that preserved file. The visible project-item name stays unchanged, so the sequence still shows the normal backup name. ExportBackup then frees the real backup filename and exports the replacement directly to it.

The selected backup video layer stays visible and is included in the MP4 render. Video tracks above the selected backup track are temporarily hidden, then every video track returns to its exact previous visibility state.

After every selected export is complete, ExportBackup releases the selected old Premiere references, imports the replacements, and restores their recorded tracks and timeline positions before deleting preserved old files. If Premiere or Windows still reports a genuine pending item, ExportBackup runs a full Align Existing pass every 3 seconds: it removes the currently aligned backup clips, imports the files again, restores their recorded positions, and checks cleanup again. A successful zero-remaining response stops retries and removes the JSON recovery map. The retry window also offers Align Existing for a manual pass.

Imported backup MP4 media uses an Orange label. Imported backup MP3, WAV, and merged audio media uses a Brown label.

## Align Existing

Use `Align Existing` when the files were already exported and you only want to import and align them.

Select the same backup location mode, then click `Align Existing`.

If a sequence was renamed after its backups were imported, Align Existing can rename old backup files to the active sequence name. It only considers media already used in the active sequence, requires the standard `_BACKUP` or `_Track...` filename rule, checks the clip duration against the active sequence, and refuses to guess when multiple old basenames match. Recorded timeline positions are restored after the rename.

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
