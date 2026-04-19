var exportBackup = exportBackup || {};
var EB_ENCODER_LAUNCH_WAIT_MS = 20000;
var EB_ENCODER_QUEUE_SETTLE_WAIT_MS = 10000;

function ebEscape(value) {
    if (value === null || value === undefined) {
        return "";
    }

    return String(value)
        .split("\\").join("\\\\")
        .split('"').join('\\"')
        .split("\r").join("\\r")
        .split("\n").join("\\n");
}

function ebStringify(payload) {
    try {
        return JSON.stringify(payload);
    } catch (e) {
        return '{"ok":false,"message":"' + ebEscape(e.toString()) + '"}';
    }
}

function ebResult(ok, message, extra) {
    var payload = { ok: !!ok, message: message || "" };
    var key;

    if (extra) {
        for (key in extra) {
            if (extra.hasOwnProperty(key)) {
                payload[key] = extra[key];
            }
        }
    }

    return ebStringify(payload);
}

function ebGetActiveSequence() {
    if (!app || !app.project || !app.project.activeSequence) {
        return null;
    }

    return app.project.activeSequence;
}

function ebGetSequenceName(sequence) {
    try {
        if (sequence && sequence.name) {
            return sequence.name;
        }
    } catch (e) {}

    return "Active_Sequence";
}

function ebSanitizeName(name) {
    var value = String(name || "Active_Sequence");
    value = value.replace(/[\\\/:\*\?"<>\|]/g, "_");
    value = value.replace(/^\s+|\s+$/g, "");
    return value || "Active_Sequence";
}

function ebGetSequenceExportBaseName(sequence) {
    return ebSanitizeName(ebGetSequenceName(sequence));
}

function ebEnsureFolder(path) {
    var folder = new Folder(path);
    if (folder.exists) {
        return true;
    }

    if (folder.parent && !folder.parent.exists) {
        ebEnsureFolder(folder.parent.fsName);
    }

    return folder.create();
}

function ebToFsPath(path) {
    if (!path) {
        return path;
    }

    try {
        return new File(path).fsName;
    } catch (e) {}

    return String(path).split("/").join("\\");
}

function ebSequenceHasInOut(sequence) {
    try {
        var inPoint = parseFloat(sequence.getInPoint());
        var outPoint = parseFloat(sequence.getOutPoint());
        return !isNaN(inPoint) && !isNaN(outPoint) && outPoint > inPoint;
    } catch (e) {
        return false;
    }
}

function ebGetExportExtension(sequence, presetPath, fallback) {
    try {
        var ext = sequence.getExportFileExtension(presetPath);
        if (ext && ext !== "") {
            return ext.charAt(0) === "." ? ext : "." + ext;
        }
    } catch (e) {}

    return fallback;
}

function ebTrackHasClips(track) {
    try {
        return track && track.clips && track.clips.numItems > 0;
    } catch (e) {
        return false;
    }
}

function ebGetTrackCount(trackCollection) {
    if (!trackCollection || trackCollection.numTracks === undefined) {
        return 0;
    }

    return trackCollection.numTracks;
}

function ebCaptureMuteStates(sequence) {
    var states = [];
    var i;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return states;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        try {
            states.push(sequence.audioTracks[i].isMuted && sequence.audioTracks[i].isMuted() ? 1 : 0);
        } catch (e) {
            states.push(0);
        }
    }

    return states;
}

function ebCaptureVideoMuteStates(sequence) {
    var states = [];
    var i;

    if (!sequence.videoTracks || sequence.videoTracks.numTracks === undefined) {
        return states;
    }

    for (i = 0; i < sequence.videoTracks.numTracks; i++) {
        try {
            states.push(sequence.videoTracks[i].isMuted && sequence.videoTracks[i].isMuted() ? 1 : 0);
        } catch (e) {
            states.push(0);
        }
    }

    return states;
}

function ebRestoreMuteStates(sequence, states) {
    var i;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        if (sequence.audioTracks[i] && sequence.audioTracks[i].setMute && i < states.length) {
            sequence.audioTracks[i].setMute(states[i]);
        }
    }
}

function ebRestoreVideoMuteStates(sequence, states) {
    var i;

    if (!sequence.videoTracks || sequence.videoTracks.numTracks === undefined) {
        return;
    }

    for (i = 0; i < sequence.videoTracks.numTracks; i++) {
        if (sequence.videoTracks[i] && sequence.videoTracks[i].setMute && i < states.length) {
            sequence.videoTracks[i].setMute(states[i]);
        }
    }
}

function ebSetAllTrackMutes(sequence, muteValue) {
    var i;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        if (sequence.audioTracks[i] && sequence.audioTracks[i].setMute) {
            sequence.audioTracks[i].setMute(muteValue ? 1 : 0);
        }
    }
}

function ebSetOnlyTrackAudible(sequence, targetIndex) {
    var i;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        if (sequence.audioTracks[i] && sequence.audioTracks[i].setMute) {
            sequence.audioTracks[i].setMute(i === targetIndex ? 0 : 1);
        }
    }
}

function ebHideVideoTracksAbove(sequence, visibleThroughTrackNumber) {
    var maxVisible = Math.max(1, parseInt(visibleThroughTrackNumber, 10) || 1);
    var hiddenCount = 0;
    var i;

    if (!sequence.videoTracks || sequence.videoTracks.numTracks === undefined) {
        return hiddenCount;
    }

    for (i = 0; i < sequence.videoTracks.numTracks; i++) {
        if ((i + 1) > maxVisible && sequence.videoTracks[i] && sequence.videoTracks[i].setMute) {
            sequence.videoTracks[i].setMute(1);
            hiddenCount += 1;
        }
    }

    return hiddenCount;
}

function ebQueueSequence(sequence, outputPath, presetPath, workAreaType) {
    return app.encoder.encodeSequence(sequence, ebToFsPath(outputPath), ebToFsPath(presetPath), workAreaType, 0);
}

function ebWaitForEncoderQueueSettle() {
    $.sleep(EB_ENCODER_QUEUE_SETTLE_WAIT_MS);
}

function ebQueueSequenceWithSettle(sequence, outputPath, presetPath, workAreaType) {
    var jobId = ebQueueSequence(sequence, outputPath, presetPath, workAreaType);
    if (jobId && jobId !== "0") {
        ebWaitForEncoderQueueSettle();
    }

    return jobId;
}

function ebRemoveAllSequenceMarkers(sequence) {
    var removedCount = 0;
    var markers = null;
    var marker = null;
    var nextMarker = null;

    if (!sequence || !sequence.markers || !sequence.markers.getFirstMarker || !sequence.markers.deleteMarker) {
        return removedCount;
    }

    markers = sequence.markers;
    marker = markers.getFirstMarker();

    while (marker) {
        nextMarker = markers.getNextMarker ? markers.getNextMarker(marker) : null;
        markers.deleteMarker(marker);
        removedCount += 1;
        marker = nextMarker;
    }

    return removedCount;
}

function ebCheckPreset(path, label) {
    var file = new File(ebToFsPath(path));
    if (!file.exists) {
        throw new Error(label + " preset was not found: " + path);
    }
}

function ebFindProjectItemByMediaPath(rootItem, mediaPath) {
    var i;

    if (!rootItem || !rootItem.children) {
        return null;
    }

    for (i = 0; i < rootItem.children.numItems; i++) {
        var child = rootItem.children[i];
        if (!child) {
            continue;
        }

        if (child.type === ProjectItemType.BIN) {
            var nested = ebFindProjectItemByMediaPath(child, mediaPath);
            if (nested) {
                return nested;
            }
        } else {
            try {
                if (child.getMediaPath && child.getMediaPath() === mediaPath) {
                    return child;
                }
            } catch (e) {}
        }
    }

    return null;
}

function ebFindChildBinByName(parentItem, name) {
    var i;

    if (!parentItem || !parentItem.children) {
        return null;
    }

    for (i = 0; i < parentItem.children.numItems; i++) {
        var child = parentItem.children[i];
        if (child && child.type === ProjectItemType.BIN && child.name === name) {
            return child;
        }
    }

    return null;
}

function ebEnsureBin(parentItem, name) {
    var existing = ebFindChildBinByName(parentItem, name);
    if (existing) {
        return existing;
    }

    if (parentItem && parentItem.createBin) {
        parentItem.createBin(name);
        return ebFindChildBinByName(parentItem, name);
    }

    return null;
}

function ebGetImportBin(sequence) {
    var root = app.project && app.project.rootItem ? app.project.rootItem : null;
    if (!root) {
        return null;
    }

    return ebEnsureBin(root, "BACKUP") || root;
}

function ebImportProjectItem(mediaPath, targetBin) {
    var fsPath = ebToFsPath(mediaPath);
    app.project.importFiles([fsPath], false, targetBin || app.project.rootItem, false);
    return ebFindProjectItemByMediaPath(targetBin || app.project.rootItem, fsPath) || ebFindProjectItemByMediaPath(app.project.rootItem, fsPath);
}

function ebCreateTimeAtZero() {
    var when = new Time();
    when.seconds = 0;
    return when;
}

function ebGetHighestUsedAudioTrackNumber(sequence) {
    var highest = 0;
    var i;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return 0;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        if (ebTrackHasClips(sequence.audioTracks[i])) {
            highest = i + 1;
        }
    }

    return highest;
}

function ebEnsureAudioTrackCount(sequence, requiredCount) {
    var currentCount = ebGetTrackCount(sequence.audioTracks);
    var tracksToAdd = requiredCount - currentCount;

    if (tracksToAdd <= 0) {
        return ebGetTrackCount(sequence.audioTracks);
    }

    app.enableQE();

    if (typeof qe === "undefined" || !qe.project || !qe.project.getActiveSequence) {
        throw new Error("QE DOM is not available, so audio tracks could not be created automatically.");
    }

    var qeSequence = qe.project.getActiveSequence();
    if (!qeSequence || !qeSequence.addTracks) {
        throw new Error("The active QE sequence could not be accessed for automatic audio-track creation.");
    }

    qeSequence.addTracks(0, 0, tracksToAdd, 3, currentCount);
    return ebGetTrackCount(sequence.audioTracks);
}

function ebValidateBackupTrack(sequence, backupVideoTrackNumber) {
    var resolved = Math.max(1, parseInt(backupVideoTrackNumber, 10) || 1);
    var currentVideoTracks = ebGetTrackCount(sequence.videoTracks);

    if (resolved > currentVideoTracks) {
        throw new Error("V" + resolved + " does not exist in the active sequence.");
    }

    if (ebTrackHasClips(sequence.videoTracks[resolved - 1])) {
        throw new Error("V" + resolved + " is not empty.");
    }
}

function ebBuildQueuedFile(kind, path, trackNumber) {
    return {
        kind: kind,
        path: ebToFsPath(path),
        trackNumber: trackNumber || 0,
        name: new File(ebToFsPath(path)).name
    };
}

function ebGetPathExtension(mediaPath) {
    var resolved = String(mediaPath || "").toLowerCase();
    var dotIndex = resolved.lastIndexOf(".");
    if (dotIndex < 0) {
        return "";
    }

    return resolved.substring(dotIndex);
}

function ebGetOrganizerBinNameForItem(item) {
    var mediaPath = "";
    var extension = "";
    var videoExtensions = {
        ".mp4": true, ".mov": true, ".mxf": true, ".avi": true, ".m4v": true, ".mpg": true,
        ".mpeg": true, ".wmv": true, ".webm": true, ".mts": true, ".m2ts": true
    };
    var audioExtensions = {
        ".wav": true, ".mp3": true, ".aac": true, ".m4a": true, ".aif": true, ".aiff": true,
        ".flac": true, ".ogg": true
    };
    var imageExtensions = {
        ".png": true, ".jpg": true, ".jpeg": true, ".tif": true, ".tiff": true, ".bmp": true,
        ".gif": true, ".webp": true, ".psd": true, ".exr": true, ".dpx": true
    };
    var graphicExtensions = {
        ".mogrt": true, ".ai": true, ".eps": true, ".svg": true, ".pdf": true
    };

    try {
        mediaPath = item && item.getMediaPath ? item.getMediaPath() : "";
    } catch (e2) {
        mediaPath = "";
    }

    extension = ebGetPathExtension(mediaPath);
    if (videoExtensions[extension]) {
        return "VIDEO";
    }
    if (audioExtensions[extension]) {
        return "AUDIO";
    }
    if (imageExtensions[extension]) {
        return "IMAGES";
    }
    if (graphicExtensions[extension]) {
        return "GRAPHICS";
    }

    return "OTHER";
}

function ebOrganizeLooseRootItems() {
    var root = app.project && app.project.rootItem ? app.project.rootItem : null;
    var organizerRoot;
    var itemsToMove = [];
    var movedCounts = {};
    var notes = [];
    var i;

    if (!root || !root.children) {
        return notes;
    }

    organizerRoot = ebEnsureBin(root, "ORGANIZED");
    if (!organizerRoot) {
        notes.push("Could not create the ORGANIZED bin.");
        return notes;
    }

    for (i = 0; i < root.children.numItems; i++) {
        var child = root.children[i];
        var isSequence = false;
        if (!child || child.type === ProjectItemType.BIN) {
            continue;
        }

        try {
            isSequence = child.isSequence && child.isSequence();
        } catch (e) {
            isSequence = false;
        }

        if (isSequence) {
            continue;
        }

        itemsToMove.push(child);
    }

    for (i = 0; i < itemsToMove.length; i++) {
        var item = itemsToMove[i];
        var binName = ebGetOrganizerBinNameForItem(item);
        var targetBin = ebEnsureBin(organizerRoot, binName);

        if (!targetBin || !item.moveBin) {
            continue;
        }

        try {
            item.moveBin(targetBin);
            movedCounts[binName] = (movedCounts[binName] || 0) + 1;
        } catch (moveError) {}
    }

    for (var key in movedCounts) {
        if (movedCounts.hasOwnProperty(key)) {
            notes.push("Organized " + movedCounts[key] + " loose item(s) into " + key + ".");
        }
    }

    if (!notes.length) {
        notes.push("No loose root items needed organizing.");
    }

    return notes;
}

exportBackup.getActiveSequenceName = function () {
    var sequence = ebGetActiveSequence();
    return sequence ? ebGetSequenceName(sequence) : "";
};

exportBackup.getExportSelectionInfo = function () {
    try {
        var sequence = ebGetActiveSequence();
        var items = [{
            kind: "video",
            label: "Backup MP4",
            selected: true,
            locked: false,
            trackNumber: 0
        }];
        var i;

        if (!sequence) {
            return ebResult(false, "No active sequence is open in Premiere Pro.");
        }

        for (i = 0; i < ebGetTrackCount(sequence.audioTracks); i++) {
            if (!ebTrackHasClips(sequence.audioTracks[i])) {
                continue;
            }

            items.push({
                kind: "audio",
                label: "Track " + (i + 1),
                selected: true,
                locked: false,
                trackNumber: i + 1
            });
        }

        return ebResult(true, items.length > 1 ? "Choose which backup files should be queued." : "Choose which backup files should be queued.", {
            sequenceName: ebGetSequenceName(sequence),
            items: items
        });
    } catch (e) {
        return ebResult(false, e.toString());
    }
};

exportBackup.validateBackupExportSettings = function (backupVideoTrackNumber) {
    try {
        var sequence = ebGetActiveSequence();
        if (!sequence) {
            return ebResult(false, "No active sequence is open in Premiere Pro.");
        }

        ebValidateBackupTrack(sequence, backupVideoTrackNumber);
        return ebResult(true, "OK");
    } catch (e) {
        return ebResult(false, e.toString());
    }
};

exportBackup.getAlignmentDefaults = function () {
    try {
        var sequence = ebGetActiveSequence();
        var currentVideoTracks = sequence ? ebGetTrackCount(sequence.videoTracks) : 0;
        var suggestedVideoTrack = currentVideoTracks >= 5 ? 5 : Math.max(1, currentVideoTracks);

        return ebResult(true, "OK", {
            hasActiveSequence: !!sequence,
            sequenceName: sequence ? ebGetSequenceName(sequence) : "",
            currentVideoTracks: currentVideoTracks,
            currentAudioTracks: sequence ? ebGetTrackCount(sequence.audioTracks) : 0,
            suggestedVideoTrack: suggestedVideoTrack || 1
        });
    } catch (e) {
        return ebResult(false, e.toString());
    }
};

exportBackup.runBackupQueue = function (folderPath, videoPresetPath, mp3PresetPath, wavPresetPath, audioFormat, backupVideoTrackNumber, removeSequenceMarkers, selectedItemsJson) {
    try {
        var sequence = ebGetActiveSequence();
        if (!sequence) {
            return ebResult(false, "No active sequence is open in Premiere Pro.");
        }

        if (!folderPath || !ebEnsureFolder(folderPath)) {
            return ebResult(false, "Could not create or access the export folder.");
        }

        if (!ebSequenceHasInOut(sequence)) {
            return ebResult(false, "Set sequence In and Out points first, then run ExportBackup again.");
        }

        ebValidateBackupTrack(sequence, backupVideoTrackNumber);

        var resolvedAudioFormat = String(audioFormat || "mp3").toLowerCase() === "wav" ? "wav" : "mp3";
        var audioPresetPath = resolvedAudioFormat === "wav" ? wavPresetPath : mp3PresetPath;
        var audioLabel = resolvedAudioFormat.toUpperCase();
        var resolvedBackupVideoTrackNumber = Math.max(1, parseInt(backupVideoTrackNumber, 10) || 5);
        var sequenceName = ebGetSequenceExportBaseName(sequence);
        var originalMuteStates = ebCaptureMuteStates(sequence);
        var workAreaType = 1;
        var notes = [];
        var queuedFiles = [];
        var selectedAudioTracks = [];
        var includeBackupVideo = true;
        var shouldRemoveSequenceMarkers = removeSequenceMarkers !== false && String(removeSequenceMarkers).toLowerCase() !== "false";
        var i;

        if (selectedItemsJson) {
            try {
                var selectedItems = JSON.parse(selectedItemsJson);
                if (selectedItems && selectedItems.audioTracks && selectedItems.audioTracks.join) {
                    selectedAudioTracks = selectedItems.audioTracks;
                }
                if (selectedItems && selectedItems.includeVideo === false) {
                    includeBackupVideo = false;
                }
            } catch (e) {
                selectedAudioTracks = [];
                includeBackupVideo = true;
            }
        }

        ebCheckPreset(videoPresetPath, "Video");
        ebCheckPreset(audioPresetPath, audioLabel);

        if (shouldRemoveSequenceMarkers) {
            var removedMarkerCount = ebRemoveAllSequenceMarkers(sequence);
            notes.push("Removed " + removedMarkerCount + " sequence marker" + (removedMarkerCount === 1 ? "" : "s") + " before export.");
        }

        app.encoder.launchEncoder();
        $.sleep(EB_ENCODER_LAUNCH_WAIT_MS);

        ebSetAllTrackMutes(sequence, 0);

        if (includeBackupVideo) {
            var videoExtension = ebGetExportExtension(sequence, videoPresetPath, ".mp4");
            var videoPath = ebToFsPath(folderPath + "\\" + sequenceName + "_BACKUP" + videoExtension);
            var hiddenVideoTrackCount = ebHideVideoTracksAbove(sequence, resolvedBackupVideoTrackNumber);
            var videoJobId = ebQueueSequenceWithSettle(sequence, videoPath, videoPresetPath, workAreaType);
            if (!videoJobId || videoJobId === "0") {
                ebRestoreMuteStates(sequence, originalMuteStates);
                return ebResult(false, "Could not queue the MP4 export in Adobe Media Encoder.");
            }

            queuedFiles.push(ebBuildQueuedFile("video", videoPath, 0));
            notes.push("Queued MP4 backup: " + videoPath);
            if (hiddenVideoTrackCount > 0) {
                notes.push("Video tracks above V" + resolvedBackupVideoTrackNumber + " were hidden while the backup MP4 queue item was created.");
            }
        } else {
            notes.push("Skipped backup MP4 export.");
        }

        for (i = 0; i < ebGetTrackCount(sequence.audioTracks); i++) {
            if (!ebTrackHasClips(sequence.audioTracks[i])) {
                continue;
            }

            if (selectedAudioTracks.length && selectedAudioTracks.join) {
                var selected = false;
                var j;
                for (j = 0; j < selectedAudioTracks.length; j++) {
                    if ((parseInt(selectedAudioTracks[j], 10) || 0) === (i + 1)) {
                        selected = true;
                        break;
                    }
                }

                if (!selected) {
                    notes.push("Skipped " + audioLabel + " export for Track" + (i + 1) + ".");
                    continue;
                }
            }

            ebSetOnlyTrackAudible(sequence, i);
            var exportTrackNumber = i + 1;
            var audioExtension = ebGetExportExtension(sequence, audioPresetPath, "." + resolvedAudioFormat);
            var audioPath = ebToFsPath(folderPath + "\\" + sequenceName + "_Track" + exportTrackNumber + audioExtension);
            var audioJobId = ebQueueSequenceWithSettle(sequence, audioPath, audioPresetPath, workAreaType);

            if (!audioJobId || audioJobId === "0") {
                notes.push("Failed to queue " + audioLabel + " for Track" + exportTrackNumber + ".");
                continue;
            }

            queuedFiles.push(ebBuildQueuedFile("audio", audioPath, exportTrackNumber));
            notes.push("Queued " + audioLabel + " track export: " + audioPath);
        }

        ebRestoreMuteStates(sequence, originalMuteStates);

        try {
            if (app.encoder.startBatch) {
                app.encoder.startBatch();
            }
        } catch (e) {}

        return ebResult(true, "Queued jobs: " + queuedFiles.length + ".\n" + notes.join("\n"), {
            sequenceName: ebGetSequenceName(sequence),
            baseName: sequenceName,
            backupVideoTrackNumber: resolvedBackupVideoTrackNumber,
            audioFormat: resolvedAudioFormat,
            queuedFiles: queuedFiles,
            projectName: app.project && app.project.name ? app.project.name : "",
            projectPath: app.project && app.project.path ? app.project.path : ""
        });
    } catch (e) {
        return ebResult(false, e.toString());
    }
};

exportBackup.alignMappedFiles = function (videoPath, audioJson, backupVideoTrackNumber, sortProjectFiles) {
    try {
        var sequence = ebGetActiveSequence();
        if (!sequence) {
            return ebResult(false, "No active sequence is open in Premiere Pro.");
        }

        var audioEntries = [];
        if (audioJson) {
            audioEntries = JSON.parse(audioJson);
        }

        if (!videoPath && (!audioEntries || !audioEntries.length)) {
            return ebResult(false, "No matched backup video or audio files were provided.");
        }

        var resolvedBackupTrack = Math.max(1, parseInt(backupVideoTrackNumber, 10) || 1);
        if (resolvedBackupTrack > ebGetTrackCount(sequence.videoTracks)) {
            return ebResult(false, "V" + resolvedBackupTrack + " does not exist in the active sequence.");
        }
        if (videoPath && ebTrackHasClips(sequence.videoTracks[resolvedBackupTrack - 1])) {
            return ebResult(false, "V" + resolvedBackupTrack + " is not empty.");
        }

        var highestUsedAudioTrack = ebGetHighestUsedAudioTrackNumber(sequence);
        var backupVideoAudioTrackNumber = videoPath ? highestUsedAudioTrack + 1 : 0;
        var firstOtherAudioTrack = videoPath ? backupVideoAudioTrackNumber + 1 : highestUsedAudioTrack + 1;
        var finalRequiredAudioTrack = firstOtherAudioTrack + (audioEntries && audioEntries.length ? audioEntries.length : 0) - 1;
        var when = ebCreateTimeAtZero();
        var notes = [];
        var importBin = ebGetImportBin(sequence);
        var organizerNotes;
        var i;

        if (videoPath) {
            finalRequiredAudioTrack = Math.max(finalRequiredAudioTrack, backupVideoAudioTrackNumber);
        }

        if (finalRequiredAudioTrack > 0) {
            ebEnsureAudioTrackCount(sequence, finalRequiredAudioTrack);
        }

        if (videoPath) {
            var videoItem = ebImportProjectItem(videoPath, importBin);
            if (!videoItem) {
                return ebResult(false, "Could not import backup video: " + videoPath);
            }

            if (sequence.overwriteClip) {
                sequence.overwriteClip(videoItem, when.seconds, resolvedBackupTrack - 1, backupVideoAudioTrackNumber - 1);
            } else {
                sequence.videoTracks[resolvedBackupTrack - 1].overwriteClip(videoItem, when);
            }

            notes.push("Aligned backup MP4 to V" + resolvedBackupTrack + " and its audio to A" + backupVideoAudioTrackNumber + ".");
        }

        for (i = 0; i < audioEntries.length; i++) {
            var targetTrackNumber = firstOtherAudioTrack + i;
            var audioItem = ebImportProjectItem(audioEntries[i].path, importBin);
            if (!audioItem) {
                return ebResult(false, "Could not import audio file: " + audioEntries[i].path);
            }

            sequence.audioTracks[targetTrackNumber - 1].overwriteClip(audioItem, when);
            notes.push("Aligned " + (audioEntries[i].name || ("Track" + audioEntries[i].trackNumber)) + " to A" + targetTrackNumber + ".");
        }

        if (sortProjectFiles) {
            organizerNotes = ebOrganizeLooseRootItems();
            for (i = 0; i < organizerNotes.length; i++) {
                notes.push(organizerNotes[i]);
            }
        } else {
            notes.push("Project file sorting was skipped.");
        }

        try {
            if (app.project && app.project.save) {
                app.project.save();
            }
        } catch (e) {}

        return ebResult(true, "Alignment completed at sequence start.\n" + notes.join("\n"), {
            projectPath: app.project && app.project.path ? app.project.path : "",
            projectName: app.project && app.project.name ? app.project.name : "",
            backupVideoTrackNumber: resolvedBackupTrack,
            backupVideoAudioTrackNumber: backupVideoAudioTrackNumber,
            firstOtherAudioTrack: firstOtherAudioTrack,
            importBinName: importBin && importBin.name ? importBin.name : ""
        });
    } catch (e) {
        return ebResult(false, e.toString());
    }
};
