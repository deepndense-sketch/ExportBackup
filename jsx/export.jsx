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

function ebSanitizeName(name) {
    var value = String(name || "Active_Sequence");
    value = value.replace(/[\\\/:\*\?"<>\|]/g, "_");
    value = value.replace(/^\s+|\s+$/g, "");
    return value || "Active_Sequence";
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

function ebTrackHasClips(track) {
    try {
        return track && track.clips && track.clips.numItems > 0;
    } catch (e) {
        return false;
    }
}

function ebQueueSequence(sequence, outputPath, presetPath, workAreaType) {
    var jobId = app.encoder.encodeSequence(sequence, ebToFsPath(outputPath), ebToFsPath(presetPath), workAreaType, 0);
    return jobId;
}

function ebCheckPreset(path, label) {
    var file = new File(ebToFsPath(path));
    if (!file.exists) {
        throw new Error(label + " preset was not found: " + path);
    }
}

exportBackup.runBackupQueue = function (folderPath, videoPresetPath, mp3PresetPath, wavPresetPath, exportMp3, exportWav) {
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

        var sequenceName = ebSanitizeName(ebGetSequenceName(sequence));
        var originalMuteStates = ebCaptureMuteStates(sequence);
        var notes = [];
        var queuedCount = 0;
        var workAreaType = 1;

        ebSetAllTrackMutes(sequence, 0);

        var videoExtension = ebGetExportExtension(sequence, videoPresetPath, ".mp4");
        var videoPath = ebToFsPath(folderPath + "\\" + sequenceName + "_backup" + videoExtension);
        var videoJobId = ebQueueSequence(sequence, videoPath, videoPresetPath, workAreaType);
        if (!videoJobId || videoJobId === "0") {
            ebRestoreMuteStates(sequence, originalMuteStates);
            return ebResult(false, "Could not queue the MP4 export in Adobe Media Encoder.\nPreset: " + ebToFsPath(videoPresetPath) + "\nOutput: " + videoPath + "\nJob ID: " + videoJobId);
        }

        queuedCount += 1;
        notes.push("Queued MP4 backup: " + videoPath);
        notes.push("MP4 job ID: " + videoJobId);

        var trackCount = sequence.audioTracks && sequence.audioTracks.numTracks !== undefined ? sequence.audioTracks.numTracks : 0;

        for (var i = 0; i < trackCount; i++) {
            var track = sequence.audioTracks[i];
            if (!ebTrackHasClips(track)) {
                continue;
            }

            var trackLabel = "A" + (i + 1);
            var safeTrackName = ebSanitizeName(track.name || trackLabel);

            if (exportMp3) {
                ebSetOnlyTrackAudible(sequence, i);
                var mp3Extension = ebGetExportExtension(sequence, mp3PresetPath, ".mp3");
                var mp3Path = ebToFsPath(folderPath + "\\" + sequenceName + "_" + trackLabel + "_" + safeTrackName + mp3Extension);
                var mp3JobId = ebQueueSequence(sequence, mp3Path, mp3PresetPath, workAreaType);
                if (mp3JobId && mp3JobId !== "0") {
                    queuedCount += 1;
                    notes.push("Queued MP3 track export: " + mp3Path);
                } else {
                    notes.push("Failed to queue MP3 for " + trackLabel + ". Preset: " + ebToFsPath(mp3PresetPath));
                }
            }

            if (exportWav) {
                ebSetOnlyTrackAudible(sequence, i);
                var wavExtension = ebGetExportExtension(sequence, wavPresetPath, ".wav");
                var wavPath = ebToFsPath(folderPath + "\\" + sequenceName + "_" + trackLabel + "_" + safeTrackName + wavExtension);
                var wavJobId = ebQueueSequence(sequence, wavPath, wavPresetPath, workAreaType);
                if (wavJobId && wavJobId !== "0") {
                    queuedCount += 1;
                    notes.push("Queued WAV track export: " + wavPath);
                } else {
                    notes.push("Failed to queue WAV for " + trackLabel + ". Preset: " + ebToFsPath(wavPresetPath));
                }
            }
        }

        ebRestoreMuteStates(sequence, originalMuteStates);

        notes.unshift("Queued jobs: " + queuedCount + ".");
        notes.push("All exports were sent to Adobe Media Encoder queue using sequence In/Out.");
        notes.push("Muted tracks were cleared for the MP4 queue job.");
        notes.push("Track solo is not exposed by Premiere's official scripting API, so clear any solo buttons manually before queuing.");

        return ebResult(true, notes.join("\n"));
    } catch (e) {
        return ebResult(false, e.toString());
    }
};
