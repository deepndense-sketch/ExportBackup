var exportBackup = exportBackup || {};

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

function ebResult(ok, message) {
    return '{"ok":' + (ok ? 'true' : 'false') + ',"message":"' + ebEscape(message) + '"}';
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

exportBackup.getActiveSequenceName = function () {
    var sequence = ebGetActiveSequence();
    if (!sequence) {
        return "";
    }

    return ebGetSequenceName(sequence);
};

exportBackup.validateBackupExportSettings = function (backupVideoTrackNumber) {
    try {
        var sequence = ebGetActiveSequence();
        if (!sequence) {
            return ebResult(false, "No active sequence is open in Premiere Pro.");
        }

        var resolvedBackupVideoTrackNumber = Math.max(1, parseInt(backupVideoTrackNumber, 10) || 1);
        var currentVideoTracks = ebGetTrackCount(sequence.videoTracks);
        if (resolvedBackupVideoTrackNumber > currentVideoTracks) {
            return ebResult(false, "V" + resolvedBackupVideoTrackNumber + " does not exist in the active sequence.");
        }

        var targetVideoTrack = sequence.videoTracks[resolvedBackupVideoTrackNumber - 1];
        if (ebTrackHasClips(targetVideoTrack)) {
            return ebResult(false, "V" + resolvedBackupVideoTrackNumber + " is not empty.");
        }

        return ebResult(true, "OK");
    } catch (e) {
        return ebResult(false, e.toString());
    }
};

exportBackup.getAlignmentDefaults = function () {
    try {
        var sequence = ebGetActiveSequence();
        var currentVideoTracks = sequence ? ebGetTrackCount(sequence.videoTracks) : 0;
        var currentAudioTracks = sequence ? ebGetTrackCount(sequence.audioTracks) : 0;
        var suggestedVideoTrack = 5;
        var suggestedVideoAudioTrack = currentAudioTracks + 1;
        var suggestedAudioStartTrack = suggestedVideoAudioTrack + 1;

        return '{' +
            '"ok":true,' +
            '"hasActiveSequence":' + (sequence ? 'true' : 'false') + ',' +
            '"sequenceName":"' + ebEscape(sequence ? ebGetSequenceName(sequence) : "") + '",' +
            '"currentVideoTracks":' + currentVideoTracks + ',' +
            '"currentAudioTracks":' + currentAudioTracks + ',' +
            '"suggestedVideoTrack":' + suggestedVideoTrack + ',' +
            '"suggestedVideoAudioTrack":' + suggestedVideoAudioTrack + ',' +
            '"suggestedAudioStartTrack":' + suggestedAudioStartTrack +
        '}';
    } catch (e) {
        return '{"ok":false,"message":"' + ebEscape(e.toString()) + '"}';
    }
};

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
            if (ext.charAt(0) !== ".") {
                return "." + ext;
            }
            return ext;
        }
    } catch (e) {}

    return fallback;
}

function ebCaptureMuteStates(sequence) {
    var states = [];

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return states;
    }

    for (var i = 0; i < sequence.audioTracks.numTracks; i++) {
        var track = sequence.audioTracks[i];
        var muted = false;

        try {
            if (track && track.isMuted) {
                muted = track.isMuted();
            }
        } catch (e) {}

        states.push(muted ? 1 : 0);
    }

    return states;
}

function ebCaptureVideoMuteStates(sequence) {
    var states = [];

    if (!sequence.videoTracks || sequence.videoTracks.numTracks === undefined) {
        return states;
    }

    for (var i = 0; i < sequence.videoTracks.numTracks; i++) {
        var track = sequence.videoTracks[i];
        var muted = false;

        try {
            if (track && track.isMuted) {
                muted = track.isMuted();
            }
        } catch (e) {}

        states.push(muted ? 1 : 0);
    }

    return states;
}

function ebRestoreMuteStates(sequence, states) {
    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return;
    }

    for (var i = 0; i < sequence.audioTracks.numTracks; i++) {
        var track = sequence.audioTracks[i];
        if (track && track.setMute && i < states.length) {
            track.setMute(states[i]);
        }
    }
}

function ebRestoreVideoMuteStates(sequence, states) {
    if (!sequence.videoTracks || sequence.videoTracks.numTracks === undefined) {
        return;
    }

    for (var i = 0; i < sequence.videoTracks.numTracks; i++) {
        var track = sequence.videoTracks[i];
        if (track && track.setMute && i < states.length) {
            track.setMute(states[i]);
        }
    }
}

function ebSetAllTrackMutes(sequence, muteValue) {
    var changed = 0;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return changed;
    }

    for (var i = 0; i < sequence.audioTracks.numTracks; i++) {
        var track = sequence.audioTracks[i];
        if (track && track.setMute) {
            track.setMute(muteValue ? 1 : 0);
            changed += 1;
        }
    }

    return changed;
}

function ebSetOnlyTrackAudible(sequence, targetIndex) {
    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return;
    }

    for (var i = 0; i < sequence.audioTracks.numTracks; i++) {
        var track = sequence.audioTracks[i];
        if (track && track.setMute) {
            track.setMute(i === targetIndex ? 0 : 1);
        }
    }
}

function ebHideVideoTracksAbove(sequence, visibleThroughTrackNumber) {
    var hiddenCount = 0;
    var maxVisible = Math.max(1, parseInt(visibleThroughTrackNumber, 10) || 1);

    if (!sequence.videoTracks || sequence.videoTracks.numTracks === undefined) {
        return hiddenCount;
    }

    for (var i = 0; i < sequence.videoTracks.numTracks; i++) {
        if ((i + 1) > maxVisible) {
            var track = sequence.videoTracks[i];
            if (track && track.setMute) {
                track.setMute(1);
                hiddenCount += 1;
            }
        }
    }

    return hiddenCount;
}

function ebTrackHasClips(track) {
    try {
        return track && track.clips && track.clips.numItems > 0;
    } catch (e) {
        return false;
    }
}

function ebQueueSequence(sequence, outputPath, presetPath, workAreaType) {
    return app.encoder.encodeSequence(sequence, ebToFsPath(outputPath), ebToFsPath(presetPath), workAreaType, 0);
}

function ebCheckPreset(path, label) {
    var file = new File(ebToFsPath(path));
    if (!file.exists) {
        throw new Error(label + " preset was not found: " + path);
    }
}

function ebWriteTextFile(filePath, contents) {
    var file = new File(ebToFsPath(filePath));
    file.encoding = "UTF-8";
    if (!file.open("w")) {
        throw new Error("Could not write file: " + filePath);
    }
    file.write(contents);
    file.close();
}

function ebReadTextFile(filePath) {
    var file = new File(ebToFsPath(filePath));
    if (!file.exists) {
        return null;
    }

    file.encoding = "UTF-8";
    if (!file.open("r")) {
        return null;
    }

    var contents = file.read();
    file.close();
    return contents;
}

function ebReadFolderEntries(folderPath) {
    var folder = new Folder(ebToFsPath(folderPath));
    if (!folder.exists) {
        return [];
    }

    var entries = folder.getFiles();
    var files = [];

    for (var i = 0; i < entries.length; i++) {
        if (entries[i] instanceof File) {
            files.push(entries[i]);
        }
    }

    return files;
}

function ebFindProjectItemByMediaPath(rootItem, mediaPath) {
    if (!rootItem || !rootItem.children) {
        return null;
    }

    for (var i = 0; i < rootItem.children.numItems; i++) {
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

function ebImportProjectItem(mediaPath) {
    var fsPath = ebToFsPath(mediaPath);
    app.project.importFiles([fsPath], false, app.project.rootItem, false);
    return ebFindProjectItemByMediaPath(app.project.rootItem, fsPath);
}

function ebGetTrackCount(trackCollection) {
    if (!trackCollection || trackCollection.numTracks === undefined) {
        return 0;
    }

    return trackCollection.numTracks;
}

function ebValidateAvailableTracks(sequence, hasVideo, audioCount, videoTrackNumber, videoAudioTrackNumber, audioStartTrackNumber) {
    var currentVideoCount = ebGetTrackCount(sequence.videoTracks);
    var currentAudioCount = ebGetTrackCount(sequence.audioTracks);
    var requiredVideoCount = hasVideo ? Math.max(1, parseInt(videoTrackNumber, 10) || 1) : 0;
    var requiredAudioCount = 0;

    if (hasVideo && videoAudioTrackNumber && videoAudioTrackNumber > 0) {
        requiredAudioCount = Math.max(requiredAudioCount, parseInt(videoAudioTrackNumber, 10) || 1);
    }

    if (audioCount > 0) {
        requiredAudioCount = Math.max(
            requiredAudioCount,
            (parseInt(audioStartTrackNumber, 10) || 1) + audioCount - 1
        );
    }

    if (currentVideoCount >= requiredVideoCount && currentAudioCount >= requiredAudioCount) {
        return;
    }

    var lines = ["Not enough tracks in the active sequence."];

    if (currentVideoCount < requiredVideoCount) {
        lines.push("Video tracks needed: " + requiredVideoCount + ". Current video tracks: " + currentVideoCount + ".");
    }

    if (currentAudioCount < requiredAudioCount) {
        lines.push("Audio tracks needed: " + requiredAudioCount + ". Current audio tracks: " + currentAudioCount + ".");
    }

    if (hasVideo) {
        lines.push("Backup video target: V" + videoTrackNumber + ".");
        if (videoAudioTrackNumber && videoAudioTrackNumber > 0) {
            lines.push("Backup video audio target: A" + videoAudioTrackNumber + ".");
        }
    }

    if (audioCount > 0) {
        lines.push(
            "Other audio files: " + audioCount +
            ". Start track: A" + audioStartTrackNumber +
            ". End track needed: A" + ((parseInt(audioStartTrackNumber, 10) || 1) + audioCount - 1) + "."
        );
    }

    throw new Error(lines.join("\n"));
}

function ebValidateEmptyTargetTracks(sequence, hasVideo, audioCount, videoTrackNumber, videoAudioTrackNumber, audioStartTrackNumber) {
    var issues = [];

    if (hasVideo) {
        var targetVideoTrack = sequence.videoTracks[videoTrackNumber - 1];
        if (ebTrackHasClips(targetVideoTrack)) {
            issues.push("V" + videoTrackNumber + " is not empty.");
        }

        if (videoAudioTrackNumber && videoAudioTrackNumber > 0) {
            var targetVideoAudioTrack = sequence.audioTracks[videoAudioTrackNumber - 1];
            if (ebTrackHasClips(targetVideoAudioTrack)) {
                issues.push("A" + videoAudioTrackNumber + " is not empty.");
            }
        }
    }

    for (var i = 0; i < audioCount; i++) {
        var targetAudioTrackNumber = audioStartTrackNumber + i;
        var targetAudioTrack = sequence.audioTracks[targetAudioTrackNumber - 1];
        if (ebTrackHasClips(targetAudioTrack)) {
            issues.push("A" + targetAudioTrackNumber + " is not empty.");
        }
    }

    if (issues.length > 0) {
        issues.push("Choose empty destination tracks and try again.");
        throw new Error(issues.join("\n"));
    }
}

function ebCreateTimeAtZero() {
    var when = new Time();
    when.seconds = 0;
    return when;
}


function ebAlignFilesToSequence(sequence, videoPath, audioEntries, videoTrackNumber, videoAudioTrackNumber, audioStartTrackNumber) {
    if (!sequence) {
        throw new Error("No active sequence is open in Premiere Pro.");
    }

    var when = ebCreateTimeAtZero();
    var notes = [];
    var hasVideo = !!videoPath;

    ebValidateAvailableTracks(
        sequence,
        hasVideo,
        audioEntries.length,
        videoTrackNumber,
        videoAudioTrackNumber,
        audioStartTrackNumber
    );
    ebValidateEmptyTargetTracks(
        sequence,
        hasVideo,
        audioEntries.length,
        videoTrackNumber,
        videoAudioTrackNumber,
        audioStartTrackNumber
    );

    if (videoPath) {
        var videoTrack = sequence.videoTracks[videoTrackNumber - 1];
        var videoItem = ebImportProjectItem(videoPath);
        if (!videoItem) {
            throw new Error("Could not import backup video: " + videoPath);
        }
        videoTrack.overwriteClip(videoItem, when);
        notes.push("Aligned backup video to V" + videoTrackNumber + ": " + videoPath);
    } else {
        notes.push("No matching BACKUP video file was found.");
    }

    for (var i = 0; i < audioEntries.length; i++) {
        var targetTrackNumber = audioStartTrackNumber + i;
        var audioTrack = sequence.audioTracks[targetTrackNumber - 1];
        var audioItem = ebImportProjectItem(audioEntries[i].path);
        if (!audioItem) {
            throw new Error("Could not import audio track file: " + audioEntries[i].path);
        }
        audioTrack.overwriteClip(audioItem, when);
        notes.push("Aligned audio to A" + targetTrackNumber + ": " + audioEntries[i].name);
    }

    return notes;
}

exportBackup.runBackupQueue = function (folderPath, videoPresetPath, mp3PresetPath, wavPresetPath, exportMp3, exportWav, backupVideoTrackNumber) {
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

        ebCheckPreset(videoPresetPath, "Video");
        if (exportMp3) {
            ebCheckPreset(mp3PresetPath, "MP3");
        }
        if (exportWav) {
            ebCheckPreset(wavPresetPath, "WAV");
        }

        app.encoder.launchEncoder();
        $.sleep(2500);

        var sequenceName = ebGetSequenceExportBaseName(sequence);
        var originalMuteStates = ebCaptureMuteStates(sequence);
        var notes = [];
        var queuedCount = 0;
        var workAreaType = 1;
        var resolvedBackupVideoTrackNumber = Math.max(1, parseInt(backupVideoTrackNumber, 10) || 5);

        ebSetAllTrackMutes(sequence, 0);

        var videoExtension = ebGetExportExtension(sequence, videoPresetPath, ".mp4");
        var videoPath = ebToFsPath(folderPath + "\\" + sequenceName + "_BACKUP" + videoExtension);
        var hiddenVideoTrackCount = 0;
        var videoJobId = 0;
        hiddenVideoTrackCount = ebHideVideoTracksAbove(sequence, resolvedBackupVideoTrackNumber);
        videoJobId = ebQueueSequence(sequence, videoPath, videoPresetPath, workAreaType);
        if (!videoJobId || videoJobId === "0") {
            ebRestoreMuteStates(sequence, originalMuteStates);
            return ebResult(false, "Could not queue the MP4 export in Adobe Media Encoder.\nPreset: " + ebToFsPath(videoPresetPath) + "\nOutput: " + videoPath + "\nJob ID: " + videoJobId);
        }

        queuedCount += 1;
        notes.push("Queued MP4 backup: " + videoPath);
        notes.push("MP4 job ID: " + videoJobId);
        notes.push("MP4 backup visible through V" + resolvedBackupVideoTrackNumber + ".");
        if (hiddenVideoTrackCount > 0) {
            notes.push("Left " + hiddenVideoTrackCount + " video track(s) above V" + resolvedBackupVideoTrackNumber + " hidden after queueing the MP4 backup.");
        } else {
            notes.push("No video tracks above V" + resolvedBackupVideoTrackNumber + " needed to be hidden for the MP4 backup.");
        }

        var trackCount = sequence.audioTracks && sequence.audioTracks.numTracks !== undefined ? sequence.audioTracks.numTracks : 0;

        for (var i = 0; i < trackCount; i++) {
            var track = sequence.audioTracks[i];
            if (!ebTrackHasClips(track)) {
                continue;
            }

            var exportTrackNumber = i + 1;

            if (exportMp3) {
                ebSetOnlyTrackAudible(sequence, i);
                var mp3Extension = ebGetExportExtension(sequence, mp3PresetPath, ".mp3");
                var mp3Path = ebToFsPath(folderPath + "\\" + sequenceName + "_Track" + exportTrackNumber + mp3Extension);
                var mp3JobId = ebQueueSequence(sequence, mp3Path, mp3PresetPath, workAreaType);
                if (mp3JobId && mp3JobId !== "0") {
                    queuedCount += 1;
                    notes.push("Queued MP3 track export: " + mp3Path);
                } else {
                    notes.push("Failed to queue MP3 for Track" + exportTrackNumber + ". Preset: " + ebToFsPath(mp3PresetPath));
                }
            }

            if (exportWav) {
                ebSetOnlyTrackAudible(sequence, i);
                var wavExtension = ebGetExportExtension(sequence, wavPresetPath, ".wav");
                var wavPath = ebToFsPath(folderPath + "\\" + sequenceName + "_Track" + exportTrackNumber + wavExtension);
                var wavJobId = ebQueueSequence(sequence, wavPath, wavPresetPath, workAreaType);
                if (wavJobId && wavJobId !== "0") {
                    queuedCount += 1;
                    notes.push("Queued WAV track export: " + wavPath);
                } else {
                    notes.push("Failed to queue WAV for Track" + exportTrackNumber + ". Preset: " + ebToFsPath(wavPresetPath));
                }
            }
        }

        ebRestoreMuteStates(sequence, originalMuteStates);

        var manifestPath = ebToFsPath(folderPath + "\\" + sequenceName + "_ALIGN.json");
        var manifest = '{' +
            '"sequenceName":"' + ebEscape(sequenceName) + '",' +
            '"folderPath":"' + ebEscape(ebToFsPath(folderPath)) + '",' +
            '"videoFile":"' + ebEscape(videoPath) + '",' +
            '"backupVideoTrackNumber":' + resolvedBackupVideoTrackNumber +
        '}';
        ebWriteTextFile(manifestPath, manifest);
        notes.push("Saved alignment manifest: " + manifestPath);
        notes.push("Saved alignment default: V" + resolvedBackupVideoTrackNumber + ".");

        notes.unshift("Queued jobs: " + queuedCount + ".");
        notes.push("All exports were sent to Adobe Media Encoder queue using sequence In/Out.");
        notes.push("Expected names: " + sequenceName + "_BACKUP" + videoExtension + " and " + sequenceName + "_TrackN audio files.");
        notes.push("Muted tracks were cleared for the MP4 queue job.");
        notes.push("Track solo is not exposed by Premiere's official scripting API, so clear any solo buttons manually before queuing.");

        return ebResult(true, notes.join("\n"));
    } catch (e) {
        return ebResult(false, e.toString());
    }
};

exportBackup.alignMatchedFiles = function (videoPath, audioJson, videoTrackNumber, videoAudioTrackNumber, audioStartTrackNumber) {
    try {
        var sequence = ebGetActiveSequence();
        if (!sequence) {
            return ebResult(false, "No active sequence is open in Premiere Pro.");
        }
        var parsedAudio = [];
        if (audioJson) {
            try {
                parsedAudio = JSON.parse(audioJson);
            } catch (e) {
                return ebResult(false, "Could not parse audio file list.");
            }
        }

        if (!videoPath && (!parsedAudio || !parsedAudio.length)) {
            return ebResult(false, "No matched video or audio files were provided.");
        }

        var notes = ebAlignFilesToSequence(
            sequence,
            videoPath,
            parsedAudio || [],
            parseInt(videoTrackNumber, 10) || 1,
            parseInt(videoAudioTrackNumber, 10) || 1,
            parseInt(audioStartTrackNumber, 10) || 1
        );

        notes.unshift("Alignment completed at sequence start.");
        return ebResult(true, notes.join("\n"));
    } catch (e) {
        return ebResult(false, e.toString());
    }
};
