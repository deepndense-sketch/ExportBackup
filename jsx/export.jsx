var EB_GLOBAL = this;

if (typeof EB_GLOBAL.JSON !== "object") {
    EB_GLOBAL.JSON = {};
}

if (typeof EB_GLOBAL.JSON.stringify !== "function") {
    EB_GLOBAL.JSON.stringify = function (value) {
        function quote(string) {
            var escapable = /[\\\"\x00-\x1f\x7f-\x9f]/g;
            var meta = {
                "\b": "\\b",
                "\t": "\\t",
                "\n": "\\n",
                "\f": "\\f",
                "\r": "\\r",
                "\"": "\\\"",
                "\\": "\\\\"
            };

            return "\"" + String(string).replace(escapable, function (character) {
                var escaped = meta[character];
                if (typeof escaped === "string") {
                    return escaped;
                }
                return "\\u" + ("0000" + character.charCodeAt(0).toString(16)).slice(-4);
            }) + "\"";
        }

        function stringify(item) {
            var type = typeof item;
            var parts;
            var key;
            var index;

            if (item === null) {
                return "null";
            }
            if (type === "string") {
                return quote(item);
            }
            if (type === "number") {
                return isFinite(item) ? String(item) : "null";
            }
            if (type === "boolean") {
                return String(item);
            }
            if (type === "object") {
                if (item && typeof item.length === "number" && !(item.propertyIsEnumerable && item.propertyIsEnumerable("length"))) {
                    parts = [];
                    for (index = 0; index < item.length; index += 1) {
                        parts[index] = stringify(item[index]) || "null";
                    }
                    return "[" + parts.join(",") + "]";
                }

                parts = [];
                for (key in item) {
                    if (item.hasOwnProperty(key)) {
                        var value = stringify(item[key]);
                        if (value) {
                            parts.push(quote(key) + ":" + value);
                        }
                    }
                }
                return "{" + parts.join(",") + "}";
            }

            return undefined;
        }

        return stringify(value);
    };
}

if (typeof EB_GLOBAL.JSON.parse !== "function") {
    EB_GLOBAL.JSON.parse = function (text) {
        var source = String(text);
        var safe = /^[\],:{}\s]*$/.test(
            source
                .replace(/\\(?:["\\\/bfnrt]|u[0-9a-fA-F]{4})/g, "@")
                .replace(/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g, "]")
                .replace(/(?:^|:|,)(?:\s*\[)+/g, "")
        );

        if (!safe) {
            throw new SyntaxError("JSON.parse");
        }

        return eval("(" + source + ")");
    };
}

var JSON = EB_GLOBAL.JSON;

var exportBackup = exportBackup || {};
var EB_ENCODER_LAUNCH_WAIT_MS = 20000;
var EB_ENCODER_QUEUE_SETTLE_WAIT_MS = 10000;
var EB_MEDIA_RELEASE_WAIT_MS = 2500;

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

function ebGetTrackName(track) {
    try {
        if (track && track.name !== undefined && track.name !== null) {
            return String(track.name);
        }
    } catch (e) {}

    try {
        if (track && track.getName) {
            return String(track.getName());
        }
    } catch (e2) {}

    return "";
}

function ebSetTrackName(track, name) {
    if (!track) {
        return false;
    }

    try {
        track.name = name;
        return true;
    } catch (e) {}

    try {
        if (track.setName) {
            track.setName(name);
            return true;
        }
    } catch (e2) {}

    return false;
}

function ebNormalizeName(value) {
    return String(value || "").toLowerCase();
}

function ebStripExtension(value) {
    var text = String(value || "");
    var lastSlash = Math.max(text.lastIndexOf("\\"), text.lastIndexOf("/"));
    var fileName = lastSlash >= 0 ? text.substring(lastSlash + 1) : text;
    var dotIndex = fileName.lastIndexOf(".");

    if (dotIndex > 0) {
        return fileName.substring(0, dotIndex);
    }

    return fileName;
}

function ebNormalizeManagedName(value) {
    return ebNormalizeName(ebStripExtension(value));
}

function ebGetManagedAudioTrackRole(trackName, baseName) {
    var normalizedTrackName = ebNormalizeManagedName(trackName);
    var normalizedBaseName = ebNormalizeName(baseName);

    if (!normalizedTrackName) {
        return "";
    }

    if (
        normalizedBaseName &&
        (
            normalizedTrackName === normalizedBaseName + "_backup" ||
            normalizedTrackName === normalizedBaseName + "_backup_rebkp_temp"
        )
    ) {
        return "backup";
    }

    if (/_backup(?:_rebkp_temp)?$/.test(normalizedTrackName)) {
        return "backup";
    }

    return "";
}

function ebGetSequenceManagedAudioTrackRole(trackName, baseName) {
    var normalizedTrackName = ebNormalizeManagedName(trackName);
    var normalizedBaseName = ebNormalizeName(baseName);

    if (!normalizedTrackName || !normalizedBaseName) {
        return "";
    }

    if (
        normalizedTrackName === normalizedBaseName + "_backup" ||
        normalizedTrackName === normalizedBaseName + "_backup_rebkp_temp"
    ) {
        return "backup";
    }

    return "";
}

function ebGetSequenceManagedAudioTrackNumbers(trackName, baseName) {
    var result = [];
    var seen = {};
    var normalizedTrackName = ebNormalizeManagedName(trackName);
    var normalizedBaseName = ebNormalizeName(baseName);
    var prefix = normalizedBaseName ? normalizedBaseName + "_track" : "";
    var suffix = "";
    var parts;
    var i;

    if (!normalizedTrackName || !prefix || normalizedTrackName.indexOf(prefix) !== 0) {
        return result;
    }

    suffix = normalizedTrackName.substring(prefix.length);
    suffix = suffix.replace(/_rebkp_temp$/, "");
    parts = suffix.split("-");

    for (i = 0; i < parts.length; i++) {
        var trackNumber = parseInt(parts[i], 10) || 0;
        if (trackNumber > 0 && !seen[trackNumber]) {
            seen[trackNumber] = true;
            result.push(trackNumber);
        }
    }

    result.sort(function (a, b) { return a - b; });
    return result;
}

function ebIsSequenceManagedBackupTrack(trackName, baseName) {
    return ebGetSequenceManagedAudioTrackRole(trackName, baseName) === "backup";
}

function ebGetTrackSequenceManagedInfo(track, baseName) {
    var info = {
        hasBackup: false,
        hasManagedAudio: false
    };
    var i;

    if (track) {
        var trackName = ebGetTrackName(track);
        if (ebIsSequenceManagedBackupTrack(trackName, baseName)) {
            info.hasBackup = true;
        }
        if (ebGetSequenceManagedAudioTrackNumbers(trackName, baseName).length) {
            info.hasManagedAudio = true;
        }
    }

    if (!track || !track.clips || track.clips.numItems === undefined) {
        return info;
    }

    for (i = 0; i < track.clips.numItems; i++) {
        var clipName = ebGetClipDisplayName(track.clips[i]);
        if (ebIsSequenceManagedBackupTrack(clipName, baseName)) {
            info.hasBackup = true;
        }
        if (ebGetSequenceManagedAudioTrackNumbers(clipName, baseName).length) {
            info.hasManagedAudio = true;
        }
    }

    return info;
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

function ebTrySetSoloMethod(track, methodName) {
    var attempts = [0, false, "0"];
    var i;

    if (!track || !methodName || typeof track[methodName] !== "function") {
        return 0;
    }

    for (i = 0; i < attempts.length; i++) {
        try {
            track[methodName](attempts[i]);
            return 1;
        } catch (e) {}
    }

    return 0;
}

function ebClearSoloOnTrack(track) {
    var changed = 0;
    var key;
    var lowerKey;

    if (!track) {
        return changed;
    }

    changed += ebTrySetSoloMethod(track, "setSolo");
    changed += ebTrySetSoloMethod(track, "setSoloed");
    changed += ebTrySetSoloMethod(track, "setSoloState");
    changed += ebTrySetSoloMethod(track, "setSoloValue");
    changed += ebTrySetSoloMethod(track, "setSoloStatus");

    for (key in track) {
        lowerKey = String(key).toLowerCase();
        if (lowerKey.indexOf("solo") < 0) {
            continue;
        }

        try {
            if (typeof track[key] === "function" && lowerKey.indexOf("set") === 0) {
                changed += ebTrySetSoloMethod(track, key);
            }
        } catch (methodError) {}

        try {
            if (typeof track[key] !== "function") {
                track[key] = 0;
                changed += 1;
            }
        } catch (propertyError) {}
    }

    try {
        if (track.solo !== undefined) {
            track.solo = 0;
            changed += 1;
        }
    } catch (e) {}

    try {
        if (track.soloed !== undefined) {
            track.soloed = 0;
            changed += 1;
        }
    } catch (e2) {}

    try {
        if (track.isSoloed && track.isSoloed() && track.toggleSolo) {
            track.toggleSolo();
            changed += 1;
        }
    } catch (e3) {}

    try {
        if (track.isSolo && track.isSolo() && track.toggleSolo) {
            track.toggleSolo();
            changed += 1;
        }
    } catch (e4) {}

    return changed;
}

function ebClearAllAudioSoloStates(sequence) {
    var cleared = 0;
    var i;

    if (sequence.audioTracks && sequence.audioTracks.numTracks !== undefined) {
        for (i = 0; i < sequence.audioTracks.numTracks; i++) {
            cleared += ebClearSoloOnTrack(sequence.audioTracks[i]);
        }
    }

    try {
        app.enableQE();
        var qeSequence = qe.project.getActiveSequence();
        var count = sequence.audioTracks && sequence.audioTracks.numTracks !== undefined ? sequence.audioTracks.numTracks : 0;
        for (i = 0; i < count; i++) {
            var qeTrack = null;
            try {
                if (qeSequence && qeSequence.getAudioTrackAt) {
                    qeTrack = qeSequence.getAudioTrackAt(i);
                }
            } catch (qeGetError) {}

            cleared += ebClearSoloOnTrack(qeTrack);

            qeTrack = null;
            try {
                if (qeSequence && qeSequence.getAudioTrackAt) {
                    qeTrack = qeSequence.getAudioTrackAt(i + 1);
                }
            } catch (qeGetError2) {}

            cleared += ebClearSoloOnTrack(qeTrack);

            try {
                if (qeSequence && qeSequence.audioTracks && qeSequence.audioTracks[i]) {
                    cleared += ebClearSoloOnTrack(qeSequence.audioTracks[i]);
                }
            } catch (qeCollectionError) {}

            try {
                if (qeSequence && qeSequence.audioTracks && qeSequence.audioTracks[i + 1]) {
                    cleared += ebClearSoloOnTrack(qeSequence.audioTracks[i + 1]);
                }
            } catch (qeCollectionError2) {}

            try {
                if (qeSequence && qeSequence.audioTracks && qeSequence.audioTracks.numItems && qeSequence.audioTracks.numItems > i) {
                    cleared += ebClearSoloOnTrack(qeSequence.audioTracks[i]);
                }
            } catch (qeCollectionError3) {}

            try {
                if (qeSequence && qeSequence.audioTracks && qeSequence.audioTracks.numTracks && qeSequence.audioTracks.numTracks > i) {
                    cleared += ebClearSoloOnTrack(qeSequence.audioTracks[i]);
                }
            } catch (qeCollectionError4) {}

            try {
                if (qeSequence && qeSequence.getAudioTrack) {
                    cleared += ebClearSoloOnTrack(qeSequence.getAudioTrack(i));
                }
            } catch (qeGetTrackError) {}

            try {
                if (qeSequence && qeSequence.getAudioTrack) {
                    cleared += ebClearSoloOnTrack(qeSequence.getAudioTrack(i + 1));
                }
            } catch (qeGetTrackError2) {}
        }

        try {
            if (qeSequence && qeSequence.clearSolo) {
                qeSequence.clearSolo();
                cleared += 1;
            }
        } catch (clearError) {}

        try {
            if (qeSequence && qeSequence.clearAllSolo) {
                qeSequence.clearAllSolo();
                cleared += 1;
            }
        } catch (clearAllError) {}
    } catch (qeError) {
        // QE solo controls are unavailable in some Premiere versions.
    }

    try {
        if (app.project && app.project.activeSequence) {
            app.project.activeSequence = sequence;
        }
    } catch (refreshError) {}

    return cleared;
}

function ebGetManagedAudioTrackNumbers(trackName, baseName) {
    var result = [];
    var seen = {};
    var normalizedTrackName = ebNormalizeManagedName(trackName);
    var normalizedBaseName = ebNormalizeName(baseName);
    var prefix = normalizedBaseName ? normalizedBaseName + "_track" : "";
    var suffix = "";
    var genericMatch;
    var parts;
    var i;

    if (!normalizedTrackName) {
        return result;
    }

    if (prefix && normalizedTrackName.indexOf(prefix) === 0) {
        suffix = normalizedTrackName.substring(prefix.length);
        suffix = suffix.replace(/_rebkp_temp$/, "");
    } else {
        genericMatch = normalizedTrackName.match(/_track(\d+(?:-\d+)*)(?:_rebkp_temp)?$/);
        if (!genericMatch) {
            return result;
        }
        suffix = genericMatch[1];
    }

    parts = suffix.split("-");

    for (i = 0; i < parts.length; i++) {
        var trackNumber = parseInt(parts[i], 10) || 0;
        if (trackNumber > 0 && !seen[trackNumber]) {
            seen[trackNumber] = true;
            result.push(trackNumber);
        }
    }

    result.sort(function (a, b) { return a - b; });
    return result;
}

function ebGetManagedAudioTrackNumber(trackName, baseName) {
    var trackNumbers = ebGetManagedAudioTrackNumbers(trackName, baseName);
    return trackNumbers.length ? trackNumbers[0] : 0;
}

function ebIsManagedAudioTrack(trackName, baseName) {
    return ebGetManagedAudioTrackNumber(trackName, baseName) > 0;
}

function ebIsManagedBackupTrack(trackName, baseName) {
    return ebGetManagedAudioTrackRole(trackName, baseName) === "backup";
}

function ebGetClipDisplayName(clip) {
    try {
        if (clip && clip.projectItem && clip.projectItem.name) {
            return String(clip.projectItem.name);
        }
    } catch (e) {}

    try {
        if (clip && clip.name) {
            return String(clip.name);
        }
    } catch (e2) {}

    return "";
}

function ebGetTrackManagedInfo(track, baseName) {
    var info = {
        hasBackup: false,
        hasManagedAudio: false
    };
    var i;

    if (track) {
        var trackName = ebGetTrackName(track);
        if (ebIsManagedBackupTrack(trackName, baseName)) {
            info.hasBackup = true;
        }
        if (ebGetManagedAudioTrackNumber(trackName, baseName) > 0) {
            info.hasManagedAudio = true;
        }
    }

    if (!track || !track.clips || track.clips.numItems === undefined) {
        return info;
    }

    for (i = 0; i < track.clips.numItems; i++) {
        var clipName = ebGetClipDisplayName(track.clips[i]);
        var managedTrackNumber = ebGetManagedAudioTrackNumber(clipName, baseName);

        if (ebIsManagedBackupTrack(clipName, baseName)) {
            info.hasBackup = true;
        }

        if (managedTrackNumber > 0) {
            info.hasManagedAudio = true;
        }
    }

    return info;
}

function ebGetClipStartInfo(clip) {
    var info = {
        startSeconds: 0,
        startTicks: ""
    };

    try {
        if (clip && clip.start) {
            var seconds = parseFloat(clip.start.seconds);
            if (!isNaN(seconds)) {
                info.startSeconds = seconds;
            }

            if (clip.start.ticks !== undefined && clip.start.ticks !== null) {
                info.startTicks = String(clip.start.ticks);
            }
        }
    } catch (e) {}

    return info;
}

function ebAddClipStartInfo(target, clip) {
    var startInfo = ebGetClipStartInfo(clip);
    target.startSeconds = startInfo.startSeconds;
    target.startTicks = startInfo.startTicks;
    return target;
}

function ebCaptureRebackupLayout(sequence, baseName) {
    var layout = {
        video: null,
        backupAudio: null,
        audioOutputs: []
    };
    var seenAudioOutputs = {};
    var i;
    var j;

    if (sequence.videoTracks && sequence.videoTracks.numTracks !== undefined) {
        for (i = 0; i < sequence.videoTracks.numTracks && !layout.video; i++) {
            var videoTrack = sequence.videoTracks[i];
            if (!videoTrack || !videoTrack.clips || videoTrack.clips.numItems === undefined) {
                continue;
            }

            for (j = 0; j < videoTrack.clips.numItems; j++) {
                if (ebIsSequenceManagedBackupTrack(ebGetClipDisplayName(videoTrack.clips[j]), baseName) || ebIsSequenceManagedBackupTrack(ebGetTrackName(videoTrack), baseName)) {
                    layout.video = ebAddClipStartInfo({
                        targetTrackNumber: i + 1,
                        mediaPath: ebGetManagedClipFinalMediaPath(videoTrack.clips[j])
                    }, videoTrack.clips[j]);
                    break;
                }
            }
        }
    }

    if (sequence.audioTracks && sequence.audioTracks.numTracks !== undefined) {
        for (i = 0; i < sequence.audioTracks.numTracks; i++) {
            var audioTrack = sequence.audioTracks[i];
            if (!audioTrack || !audioTrack.clips || audioTrack.clips.numItems === undefined) {
                continue;
            }

            for (j = 0; j < audioTrack.clips.numItems; j++) {
                var audioClip = audioTrack.clips[j];
                var clipName = ebGetClipDisplayName(audioClip);

                if (!layout.backupAudio && (ebIsSequenceManagedBackupTrack(clipName, baseName) || ebIsSequenceManagedBackupTrack(ebGetTrackName(audioTrack), baseName))) {
                    layout.backupAudio = ebAddClipStartInfo({
                        targetTrackNumber: i + 1,
                        mediaPath: ebGetManagedClipFinalMediaPath(audioClip)
                    }, audioClip);
                }

                var sourceTrackNumbers = ebGetSequenceManagedAudioTrackNumbers(clipName, baseName);
                if (sourceTrackNumbers.length) {
                    var outputKey = sourceTrackNumbers.join("-");
                    if (!seenAudioOutputs[outputKey]) {
                        seenAudioOutputs[outputKey] = true;
                        layout.audioOutputs.push(ebAddClipStartInfo({
                            sourceTrackNumber: sourceTrackNumbers[0],
                            sourceTrackNumbers: sourceTrackNumbers,
                            targetTrackNumber: i + 1,
                            mediaPath: ebGetManagedClipFinalMediaPath(audioClip)
                        }, audioClip));
                    }
                }
            }
        }
    }

    layout.audioOutputs.sort(function (a, b) {
        return (a.sourceTrackNumber || 0) - (b.sourceTrackNumber || 0);
    });

    return layout;
}

function ebIsOriginalAudioTrack(sequence, baseName, trackNumber, requireClips) {
    var resolvedTrackNumber = parseInt(trackNumber, 10) || 0;
    if (
        resolvedTrackNumber < 1 ||
        !sequence.audioTracks ||
        sequence.audioTracks.numTracks === undefined ||
        resolvedTrackNumber > sequence.audioTracks.numTracks
    ) {
        return false;
    }

    var track = sequence.audioTracks[resolvedTrackNumber - 1];
    var trackInfo = ebGetTrackManagedInfo(track, baseName);
    if (trackInfo.hasBackup || trackInfo.hasManagedAudio) {
        return false;
    }

    return requireClips ? ebTrackHasClips(track) : true;
}

function ebGetRebackupAudioDefinitions(sequence, baseName, rebackupLayout) {
    var definitions = [];
    var coveredTrackNumbers = {};
    var seenDefinitions = {};
    var layout = rebackupLayout || ebCaptureRebackupLayout(sequence, baseName);
    var existingOutputs = layout && layout.audioOutputs ? layout.audioOutputs : [];
    var i;
    var j;

    for (i = 0; i < existingOutputs.length; i++) {
        var existingTrackNumbers = existingOutputs[i].sourceTrackNumbers && existingOutputs[i].sourceTrackNumbers.length
            ? existingOutputs[i].sourceTrackNumbers
            : [existingOutputs[i].sourceTrackNumber];
        var validTrackNumbers = [];

        for (j = 0; j < existingTrackNumbers.length; j++) {
            var existingTrackNumber = parseInt(existingTrackNumbers[j], 10) || 0;
            if (
                existingTrackNumber > 0 &&
                !coveredTrackNumbers[existingTrackNumber] &&
                ebIsOriginalAudioTrack(sequence, baseName, existingTrackNumber, false)
            ) {
                coveredTrackNumbers[existingTrackNumber] = true;
                validTrackNumbers.push(existingTrackNumber);
            }
        }

        validTrackNumbers = ebNormalizeTrackGroup(validTrackNumbers);
        if (!validTrackNumbers.length) {
            continue;
        }

        var existingKey = validTrackNumbers.join("-");
        if (!seenDefinitions[existingKey]) {
            seenDefinitions[existingKey] = true;
            definitions.push({
                trackNumber: validTrackNumbers[0],
                trackNumbers: validTrackNumbers,
                mediaPath: existingOutputs[i].mediaPath || ""
            });
        }
    }

    for (i = 0; i < ebGetTrackCount(sequence.audioTracks); i++) {
        var sourceTrackNumber = i + 1;
        if (
            coveredTrackNumbers[sourceTrackNumber] ||
            !ebIsOriginalAudioTrack(sequence, baseName, sourceTrackNumber, true)
        ) {
            continue;
        }

        coveredTrackNumbers[sourceTrackNumber] = true;
        definitions.push({
            trackNumber: sourceTrackNumber,
            trackNumbers: [sourceTrackNumber]
        });
    }

    definitions.sort(function (a, b) {
        return (a.trackNumber || 0) - (b.trackNumber || 0);
    });

    return definitions;
}

function ebFindManagedAudioTrackNumber(sequence, baseName, sourceTrackNumber) {
    var i;
    var expected = Math.max(1, parseInt(sourceTrackNumber, 10) || 0);

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined || expected < 1) {
        return 0;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        if ((ebGetSequenceManagedAudioTrackNumbers(ebGetTrackName(sequence.audioTracks[i]), baseName)[0] || 0) === expected) {
            return i + 1;
        }

        var track = sequence.audioTracks[i];
        if (track && track.clips && track.clips.numItems !== undefined) {
            for (var j = 0; j < track.clips.numItems; j++) {
                if ((ebGetSequenceManagedAudioTrackNumbers(ebGetClipDisplayName(track.clips[j]), baseName)[0] || 0) === expected) {
                    return i + 1;
                }
            }
        }
    }

    return 0;
}

function ebFindManagedBackupAudioTrackNumber(sequence, baseName) {
    var managedSelection = ebGetSequenceManagedSelection(sequence, baseName);
    var keys = ebGetTrackNumberKeys(managedSelection.backupTrackNumbers);
    return keys.length ? keys[0] : 0;
}

function ebFindManagedBackupVideoTrackNumber(sequence, baseName) {
    var i;

    if (!sequence.videoTracks || sequence.videoTracks.numTracks === undefined) {
        return 0;
    }

    for (i = 0; i < sequence.videoTracks.numTracks; i++) {
        if (ebGetTrackSequenceManagedInfo(sequence.videoTracks[i], baseName).hasBackup) {
            return i + 1;
        }
    }

    return 0;
}

function ebRemoveManagedClipsFromTrack(track, baseName, mode, sourceTrackNumber) {
    var removed = 0;
    var i;

    if (!track || !track.clips || track.clips.numItems === undefined) {
        return removed;
    }

    for (i = track.clips.numItems - 1; i >= 0; i--) {
        var clip = track.clips[i];
        var clipName = ebGetClipDisplayName(clip);
        var shouldRemove = false;

        if (mode === "backup") {
            shouldRemove = ebIsManagedBackupTrack(clipName, baseName) || ebIsSequenceManagedBackupTrack(clipName, baseName);
        } else if (mode === "audio") {
            shouldRemove = ebGetManagedAudioTrackNumber(clipName, baseName) === (parseInt(sourceTrackNumber, 10) || 0);
        }

        if (shouldRemove && clip && clip.remove) {
            try {
                clip.remove(0, 0);
                removed += 1;
            } catch (e) {}
        }
    }

    return removed;
}

function ebRemoveManagedClipsFromAllAudioTracks(sequence, baseName, mode, sourceTrackNumber) {
    var removed = 0;
    var i;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return removed;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        removed += ebRemoveManagedClipsFromTrack(sequence.audioTracks[i], baseName, mode, sourceTrackNumber);
    }

    return removed;
}

function ebRemoveAllManagedClipsFromAllAudioTracks(sequence, baseName) {
    var removed = 0;
    var i;
    var j;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return removed;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        var track = sequence.audioTracks[i];
        if (!track || !track.clips || track.clips.numItems === undefined) {
            continue;
        }

        for (j = track.clips.numItems - 1; j >= 0; j--) {
            var clip = track.clips[j];
            var clipName = ebGetClipDisplayName(clip);
            if (
                (ebIsManagedBackupTrack(clipName, baseName) || ebGetManagedAudioTrackNumber(clipName, baseName) > 0) &&
                clip &&
                clip.remove
            ) {
                try {
                    clip.remove(0, 0);
                    removed += 1;
                } catch (e) {}
            }
        }
    }

    return removed;
}

function ebGetSequenceManagedSelection(sequence, baseName) {
    var selection = {
        hasBackup: false,
        backupTrackNumbers: {},
        trackNumbers: {}
    };
    var i;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return selection;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        var trackInfo = ebGetTrackManagedInfo(sequence.audioTracks[i], baseName);

        if (trackInfo.hasBackup) {
            selection.hasBackup = true;
            selection.backupTrackNumbers[i + 1] = true;
        }

        if (trackInfo.hasManagedAudio) {
            selection.trackNumbers[i + 1] = true;
        }
    }

    return selection;
}

function ebGetTrackNumberKeys(map) {
    var result = [];
    var key;

    for (key in map) {
        if (map.hasOwnProperty(key)) {
            result.push(parseInt(key, 10) || 0);
        }
    }

    result.sort(function (a, b) { return a - b; });
    return result;
}

function ebApplyManagedTrackMutePolicy(sequence, baseName) {
    var i;
    var trackInfo;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        if (!sequence.audioTracks[i] || !sequence.audioTracks[i].setMute) {
            continue;
        }

        trackInfo = ebGetTrackManagedInfo(sequence.audioTracks[i], baseName);
        if (trackInfo.hasBackup || trackInfo.hasManagedAudio) {
            sequence.audioTracks[i].setMute(1);
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

function ebExportSequenceDirect(sequence, outputPath, presetPath, workAreaType) {
    if (!sequence || !sequence.exportAsMediaDirect) {
        throw new Error("Premiere Pro direct export is not available in this version.");
    }

    return sequence.exportAsMediaDirect(ebToFsPath(outputPath), ebToFsPath(presetPath), workAreaType);
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

function ebNormalizeMediaPathForComparison(mediaPath) {
    return String(ebToFsPath(mediaPath) || "")
        .split("/").join("\\")
        .toLowerCase();
}

function ebCollectProjectItemsByMediaPath(rootItem, mediaPath, result) {
    var matches = result || [];
    var normalizedMediaPath = ebNormalizeMediaPathForComparison(mediaPath);
    var i;

    if (!rootItem || !rootItem.children || !normalizedMediaPath) {
        return matches;
    }

    for (i = 0; i < rootItem.children.numItems; i++) {
        var child = rootItem.children[i];
        if (!child) {
            continue;
        }

        if (child.type === ProjectItemType.BIN) {
            ebCollectProjectItemsByMediaPath(child, mediaPath, matches);
        } else {
            try {
                if (
                    child.getMediaPath &&
                    ebNormalizeMediaPathForComparison(child.getMediaPath()) === normalizedMediaPath
                ) {
                    matches.push(child);
                }
            } catch (e) {}
        }
    }

    return matches;
}

function ebFindProjectItemByMediaPath(rootItem, mediaPath) {
    var matches = ebCollectProjectItemsByMediaPath(rootItem, mediaPath, []);
    return matches.length ? matches[0] : null;
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

function ebIsProjectItemOffline(item) {
    try {
        if (item && item.isOffline) {
            return !!item.isOffline();
        }
    } catch (e) {}

    return false;
}

function ebSetProjectItemOffline(item) {
    if (!item || !item.setOffline) {
        return false;
    }

    try {
        var result = item.setOffline();
        if (result === true || result === 0 || ebIsProjectItemOffline(item)) {
            return true;
        }
    } catch (e) {}

    return false;
}

function ebDeleteProjectItem(item) {
    if (!item) {
        return false;
    }

    try {
        if (item.deleteBin) {
            var deleteResult = item.deleteBin();
            if (deleteResult === 0 || deleteResult === true || deleteResult === undefined) {
                return true;
            }
        }
    } catch (e) {}

    try {
        if (item.remove) {
            var removeResult = item.remove();
            return removeResult !== false;
        }
    } catch (e2) {}

    return false;
}

function ebGetProjectItemNodeId(item) {
    try {
        if (item && item.nodeId !== undefined && item.nodeId !== null) {
            return String(item.nodeId);
        }
    } catch (e) {}

    return "";
}

function ebProjectTreeContainsItem(rootItem, targetItem) {
    if (!rootItem || !targetItem) {
        return false;
    }

    var targetNodeId = ebGetProjectItemNodeId(targetItem);
    var i;

    if (rootItem === targetItem) {
        return true;
    }
    if (targetNodeId && ebGetProjectItemNodeId(rootItem) === targetNodeId) {
        return true;
    }
    if (!rootItem.children || rootItem.children.numItems === undefined) {
        return false;
    }

    for (i = 0; i < rootItem.children.numItems; i++) {
        if (ebProjectTreeContainsItem(rootItem.children[i], targetItem)) {
            return true;
        }
    }

    return false;
}

function ebGetRemainingProjectItems(items) {
    var remaining = [];
    var rootItem = app.project && app.project.rootItem ? app.project.rootItem : null;
    var i;

    for (i = 0; i < items.length; i++) {
        if (ebProjectTreeContainsItem(rootItem, items[i])) {
            remaining.push(items[i]);
        }
    }

    return remaining;
}

function ebCollectProjectItemsForRelease(rootItem, mediaPath, result) {
    var matches = result || [];
    var fsPath = ebToFsPath(mediaPath);
    var normalizedMediaPath = ebNormalizeMediaPathForComparison(fsPath);
    var targetFileName = "";
    var i;

    try {
        targetFileName = String((new File(fsPath)).name || "").toLowerCase();
    } catch (fileError) {}

    if (!rootItem || !rootItem.children) {
        return matches;
    }

    for (i = 0; i < rootItem.children.numItems; i++) {
        var child = rootItem.children[i];
        if (!child) {
            continue;
        }

        if (child.type === ProjectItemType.BIN) {
            ebCollectProjectItemsForRelease(child, fsPath, matches);
            continue;
        }

        var childPath = "";
        var childName = "";
        try {
            childPath = child.getMediaPath ? ebNormalizeMediaPathForComparison(child.getMediaPath()) : "";
        } catch (pathError) {}
        try {
            childName = String(child.name || "").toLowerCase();
        } catch (nameError) {}

        if (
            (childPath && childPath === normalizedMediaPath) ||
            ((!childPath || ebIsProjectItemOffline(child)) && targetFileName && childName === targetFileName)
        ) {
            matches.push(child);
        }
    }

    return matches;
}

function ebDeleteProjectItemsThroughTemporaryBin(items) {
    var rootItem = app.project && app.project.rootItem ? app.project.rootItem : null;
    var result = {
        moved: 0,
        binDeleted: false
    };
    var cleanupBin = null;
    var cleanupBinName = "__ExportBackup_Cleanup_" + (new Date()).getTime();
    var i;

    if (!rootItem || !rootItem.createBin || !items.length) {
        return result;
    }

    try {
        var createdBin = rootItem.createBin(cleanupBinName);
        cleanupBin = createdBin && createdBin !== 0
            ? createdBin
            : ebFindChildBinByName(rootItem, cleanupBinName);
    } catch (createError) {
        cleanupBin = null;
    }

    if (!cleanupBin) {
        return result;
    }

    for (i = 0; i < items.length; i++) {
        try {
            if (items[i] && items[i].moveBin) {
                var moveResult = items[i].moveBin(cleanupBin);
                if (moveResult === 0 || moveResult === true || moveResult === undefined) {
                    result.moved += 1;
                }
            }
        } catch (moveError) {}
    }

    try {
        if (cleanupBin.deleteBin) {
            var deleteResult = cleanupBin.deleteBin();
            result.binDeleted = deleteResult === 0 || deleteResult === true || deleteResult === undefined;
        }
    } catch (deleteError) {}

    return result;
}

function ebReleaseProjectItemsByMediaPath(mediaPath) {
    var fsPath = ebToFsPath(mediaPath);
    var items = ebCollectProjectItemsForRelease(app.project.rootItem, fsPath, []);
    var result = {
        path: fsPath,
        found: items.length,
        offlined: 0,
        removed: 0,
        remaining: 0,
        remainingOnline: 0
    };
    var i;

    // Premiere exposes deletion for bins, not ordinary footage ProjectItems.
    // Move only these targeted items into a temporary bin, then delete the bin
    // and all of its contents using the documented API.
    var cleanupResult = ebDeleteProjectItemsThroughTemporaryBin(items);
    result.movedToCleanupBin = cleanupResult.moved;
    result.cleanupBinDeleted = cleanupResult.binDeleted;

    ebRemoveUnusedMedia();
    var remainingItems = ebGetRemainingProjectItems(items);

    // If Premiere still holds a file lock, offline only the items that survived
    // the first deletion attempt, then delete them again.
    for (i = 0; i < remainingItems.length; i++) {
        if (ebSetProjectItemOffline(remainingItems[i])) {
            result.offlined += 1;
        }
    }

    var offlineCleanupResult = ebDeleteProjectItemsThroughTemporaryBin(remainingItems);
    result.movedToCleanupBin += offlineCleanupResult.moved;
    result.cleanupBinDeleted = result.cleanupBinDeleted || offlineCleanupResult.binDeleted;

    // Retain legacy fallbacks for Premiere builds that expose extra,
    // undocumented deletion methods.
    remainingItems = ebGetRemainingProjectItems(items);
    for (i = 0; i < remainingItems.length; i++) {
        ebDeleteProjectItem(remainingItems[i]);
    }

    ebRemoveUnusedMedia();
    remainingItems = ebGetRemainingProjectItems(items);
    result.removed = items.length - remainingItems.length;
    result.remaining = remainingItems.length;
    for (i = 0; i < remainingItems.length; i++) {
        if (!ebIsProjectItemOffline(remainingItems[i])) {
            result.remainingOnline += 1;
        }
    }

    return result;
}

function ebGetProjectSequenceCount(sequenceCollection) {
    if (!sequenceCollection) {
        return 0;
    }

    if (sequenceCollection.numSequences !== undefined) {
        return sequenceCollection.numSequences;
    }
    if (sequenceCollection.numItems !== undefined) {
        return sequenceCollection.numItems;
    }

    return 0;
}

function ebGetClipMediaPath(clip) {
    try {
        if (clip && clip.projectItem && clip.projectItem.getMediaPath) {
            return ebToFsPath(clip.projectItem.getMediaPath());
        }
    } catch (e) {}

    return "";
}

function ebStripRebackupTempMarkerFromPath(mediaPath) {
    var fsPath = ebToFsPath(mediaPath);
    var dotIndex = fsPath.lastIndexOf(".");
    var slashIndex = Math.max(fsPath.lastIndexOf("\\"), fsPath.lastIndexOf("/"));
    var marker = "_REBKP_TEMP";
    var markerIndex;

    if (dotIndex <= slashIndex) {
        dotIndex = fsPath.length;
    }

    markerIndex = fsPath.toUpperCase().lastIndexOf(marker, dotIndex);
    if (markerIndex < 0 || markerIndex + marker.length !== dotIndex) {
        return fsPath;
    }

    return fsPath.substring(0, markerIndex) + fsPath.substring(dotIndex);
}

function ebReplacePathExtension(filePath, extension) {
    var fsPath = ebToFsPath(filePath);
    var dotIndex = fsPath.lastIndexOf(".");
    var slashIndex = Math.max(fsPath.lastIndexOf("\\"), fsPath.lastIndexOf("/"));
    var resolvedExtension = String(extension || "");

    if (!fsPath) {
        return "";
    }

    if (resolvedExtension && resolvedExtension.charAt(0) !== ".") {
        resolvedExtension = "." + resolvedExtension;
    }

    if (dotIndex <= slashIndex) {
        return fsPath + resolvedExtension;
    }

    return fsPath.substring(0, dotIndex) + resolvedExtension;
}
function ebBuildRebackupTempPathFromFinalPath(finalPath) {
    var fsPath = ebToFsPath(finalPath);
    var dotIndex = fsPath.lastIndexOf(".");
    var slashIndex = Math.max(fsPath.lastIndexOf("\\"), fsPath.lastIndexOf("/"));

    if (!fsPath) {
        return "";
    }

    if (dotIndex <= slashIndex) {
        return fsPath + "_REBKP_TEMP";
    }

    return fsPath.substring(0, dotIndex) + "_REBKP_TEMP" + fsPath.substring(dotIndex);
}

function ebGetManagedClipFinalMediaPath(clip) {
    return ebStripRebackupTempMarkerFromPath(ebGetClipMediaPath(clip));
}

function ebRemoveProjectClipsByMediaPaths(mediaPaths) {
    var normalizedPaths = {};
    var fileNames = {};
    var sequences = app.project && app.project.sequences ? app.project.sequences : null;
    var sequenceCount = ebGetProjectSequenceCount(sequences);
    var removed = 0;
    var i;
    var j;
    var k;

    for (i = 0; i < mediaPaths.length; i++) {
        var fsPath = ebToFsPath(mediaPaths[i]);
        var normalizedPath = ebNormalizeMediaPathForComparison(fsPath);
        if (normalizedPath) {
            normalizedPaths[normalizedPath] = true;
        }
        try {
            fileNames[String((new File(fsPath)).name || "").toLowerCase()] = true;
        } catch (fileError) {}
    }

    function removeFromTracks(trackCollection) {
        if (!trackCollection || trackCollection.numTracks === undefined) {
            return;
        }

        for (j = 0; j < trackCollection.numTracks; j++) {
            var track = trackCollection[j];
            if (!track || !track.clips || track.clips.numItems === undefined) {
                continue;
            }

            for (k = track.clips.numItems - 1; k >= 0; k--) {
                var clip = track.clips[k];
                var clipPath = ebGetClipMediaPath(clip);
                var normalizedClipPath = ebNormalizeMediaPathForComparison(clipPath);
                var clipName = String(ebGetClipDisplayName(clip) || "").toLowerCase();
                var matchesPath = normalizedClipPath && normalizedPaths[normalizedClipPath];
                var matchesOfflineName = !normalizedClipPath && clipName && fileNames[clipName];

                if ((matchesPath || matchesOfflineName) && clip && clip.remove) {
                    try {
                        clip.remove(0, 0);
                        removed += 1;
                    } catch (removeError) {}
                }
            }
        }
    }

    for (i = 0; i < sequenceCount; i++) {
        var sequence = sequences[i];
        if (!sequence) {
            continue;
        }
        removeFromTracks(sequence.videoTracks);
        removeFromTracks(sequence.audioTracks);
    }

    return removed;
}

function ebGetOnlineProjectItemCountByMediaPath(mediaPath) {
    var items = ebCollectProjectItemsByMediaPath(app.project.rootItem, ebToFsPath(mediaPath), []);
    var onlineCount = 0;
    var i;

    for (i = 0; i < items.length; i++) {
        if (!ebIsProjectItemOffline(items[i])) {
            onlineCount += 1;
        }
    }

    return onlineCount;
}

function ebRemoveUnusedMedia() {
    var methods = [
        { owner: app.project, name: "deleteUnusedProjectItems" },
        { owner: app.project, name: "deleteUnused" },
        { owner: app.project, name: "removeUnused" }
    ];
    var commandNames = [
        "Remove Unused",
        "Remove Unused Media",
        "Delete Unused"
    ];
    var attempted = false;
    var i;

    try {
        app.enableQE();
        if (typeof qe !== "undefined" && qe.project) {
            methods.push({ owner: qe.project, name: "deleteUnused" });
            methods.push({ owner: qe.project, name: "removeUnused" });
            methods.push({ owner: qe.project, name: "deleteUnusedProjectItems" });
        }
    } catch (qeError) {}

    for (i = 0; i < methods.length; i++) {
        var method = methods[i];
        try {
            if (method.owner && method.owner[method.name]) {
                method.owner[method.name]();
                attempted = true;
            }
        } catch (e) {}
    }

    for (i = 0; i < commandNames.length; i++) {
        try {
            var commandId = app.findMenuCommandId(commandNames[i]);
            if (commandId) {
                app.executeCommand(commandId);
                attempted = true;
            }
        } catch (commandError) {}
    }

    return attempted;
}

function ebCreateTimeAtZero() {
    var when = new Time();
    when.seconds = 0;
    return when;
}

function ebCreateTimeFromPlacement(placement) {
    var when = new Time();
    var resolvedPlacement = placement || {};

    try {
        if (resolvedPlacement.startTicks !== undefined && resolvedPlacement.startTicks !== null && String(resolvedPlacement.startTicks)) {
            when.ticks = String(resolvedPlacement.startTicks);
            return when;
        }
    } catch (e) {}

    var seconds = parseFloat(resolvedPlacement.startSeconds);
    when.seconds = isNaN(seconds) ? 0 : seconds;
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

function ebGetHighestSourceAudioTrackNumber(sequence, baseName) {
    var highest = 0;
    var i;
    var trackInfo;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return 0;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        if (!ebTrackHasClips(sequence.audioTracks[i])) {
            continue;
        }

        trackInfo = ebGetTrackManagedInfo(sequence.audioTracks[i], baseName);
        if (trackInfo.hasBackup || trackInfo.hasManagedAudio) {
            continue;
        }

        highest = i + 1;
    }

    return highest;
}

function ebGetFirstEmptyAudioTrackNumber(sequence) {
    var i;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return 0;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        if (!ebTrackHasClips(sequence.audioTracks[i])) {
            return i + 1;
        }
    }

    return 0;
}

function ebGetEmptyAudioTrackNumbers(sequence, excludedTrackNumbers) {
    var numbers = [];
    var excluded = excludedTrackNumbers || {};
    var i;

    if (!sequence.audioTracks || sequence.audioTracks.numTracks === undefined) {
        return numbers;
    }

    for (i = 0; i < sequence.audioTracks.numTracks; i++) {
        if (excluded[i + 1]) {
            continue;
        }

        if (!ebTrackHasClips(sequence.audioTracks[i])) {
            numbers.push(i + 1);
        }
    }

    return numbers;
}

function ebBuildAudioImportTrackNumbers(sequence, count, excludedTrackNumbers) {
    var targets = [];
    var emptyTracks = ebGetEmptyAudioTrackNumbers(sequence, excludedTrackNumbers);
    var nextTrackNumber = ebGetTrackCount(sequence.audioTracks) + 1;
    var i;

    for (i = 0; i < count; i++) {
        if (i < emptyTracks.length) {
            targets.push(emptyTracks[i]);
        } else {
            while (excludedTrackNumbers && excludedTrackNumbers[nextTrackNumber]) {
                nextTrackNumber += 1;
            }
            targets.push(nextTrackNumber);
            nextTrackNumber += 1;
        }
    }

    return targets;
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

function ebFindLowestEmptyVideoTrackNumber(sequence) {
    var i;

    if (!sequence.videoTracks || sequence.videoTracks.numTracks === undefined) {
        return 0;
    }

    for (i = 0; i < sequence.videoTracks.numTracks; i++) {
        if (!ebTrackHasClips(sequence.videoTracks[i])) {
            return i + 1;
        }
    }

    return 0;
}

function ebCreateTopVideoTrack(sequence) {
    var currentCount = ebGetTrackCount(sequence.videoTracks);

    try {
        app.enableQE();
    } catch (e) {}

    if (typeof qe === "undefined" || !qe.project || !qe.project.getActiveSequence) {
        throw new Error("QE DOM is not available, so an empty video track could not be created automatically.");
    }

    var qeSequence = qe.project.getActiveSequence();
    if (!qeSequence || !qeSequence.addTracks) {
        throw new Error("The active QE sequence could not be accessed for automatic video-track creation.");
    }

    qeSequence.addTracks(1, currentCount, 0, 3, 0);
    return ebGetTrackCount(sequence.videoTracks) > currentCount ? currentCount + 1 : 0;
}

function ebResolveBackupVideoTrackNumber(sequence, backupVideoTrackNumber, autoEmptyTrack, allowCreate) {
    var useAuto = autoEmptyTrack === true || String(autoEmptyTrack).toLowerCase() === "true";
    var resolved = 0;

    if (useAuto) {
        resolved = ebFindLowestEmptyVideoTrackNumber(sequence);
        if (resolved > 0) {
            return resolved;
        }

        if (allowCreate !== false && String(allowCreate).toLowerCase() !== "false") {
            resolved = ebCreateTopVideoTrack(sequence);
            if (resolved > 0 && !ebTrackHasClips(sequence.videoTracks[resolved - 1])) {
                return resolved;
            }
        }

        throw new Error("No empty video track was found, and a new empty video track could not be created automatically.");
    }

    return Math.max(1, parseInt(backupVideoTrackNumber, 10) || 1);
}
function ebValidateBackupTrack(sequence, backupVideoTrackNumber, allowManagedBackup, baseName) {
    var resolved = Math.max(1, parseInt(backupVideoTrackNumber, 10) || 1);
    var currentVideoTracks = ebGetTrackCount(sequence.videoTracks);

    if (resolved > currentVideoTracks) {
        throw new Error("V" + resolved + " does not exist in the active sequence.");
    }

    if (ebTrackHasClips(sequence.videoTracks[resolved - 1])) {
        if (allowManagedBackup && ebGetTrackSequenceManagedInfo(sequence.videoTracks[resolved - 1], baseName).hasBackup) {
            return;
        }

        throw new Error("V" + resolved + " is not empty.");
    }
}

function ebFileExists(path) {
    if (!path) {
        return false;
    }

    try {
        return new File(ebToFsPath(path)).exists;
    } catch (e) {
        return false;
    }
}

function ebGetSelectedExportItems(selectedItemsJson) {
    var result = {
        includeVideo: true,
        audioTracks: [],
        audioGroups: [],
        hasAudioSelection: false
    };

    if (!selectedItemsJson) {
        return result;
    }

    try {
        var selectedItems = JSON.parse(selectedItemsJson);
        if (selectedItems && selectedItems.includeVideo === false) {
            result.includeVideo = false;
        }
        if (selectedItems && selectedItems.audioTracks && selectedItems.audioTracks.join) {
            result.audioTracks = selectedItems.audioTracks;
            result.hasAudioSelection = true;
        }
        if (selectedItems && selectedItems.audioGroups && selectedItems.audioGroups.join) {
            result.audioGroups = selectedItems.audioGroups;
            result.hasAudioSelection = true;
        }
    } catch (e) {}

    return result;
}

function ebTrackIsInList(trackNumber, list) {
    var i;

    if (!list) {
        return false;
    }

    for (i = 0; i < list.length; i++) {
        if ((parseInt(list[i], 10) || 0) === trackNumber) {
            return true;
        }
    }

    return false;
}

function ebTrackIsInGroups(trackNumber, groups) {
    var i;

    if (!groups) {
        return false;
    }

    for (i = 0; i < groups.length; i++) {
        if (ebTrackIsInList(trackNumber, groups[i])) {
            return true;
        }
    }

    return false;
}

function ebRebackupDefinitionIsSelected(definition, selection) {
    var trackNumbers = definition && definition.trackNumbers ? definition.trackNumbers : [];
    var i;

    for (i = 0; i < trackNumbers.length; i++) {
        var trackNumber = parseInt(trackNumbers[i], 10) || 0;
        if (trackNumber > 0 && (ebTrackIsInList(trackNumber, selection.audioTracks) || ebTrackIsInGroups(trackNumber, selection.audioGroups))) {
            return true;
        }
    }

    return false;
}
function ebNormalizeTrackGroup(group) {
    var result = [];
    var seen = {};
    var i;

    if (!group) {
        return result;
    }

    for (i = 0; i < group.length; i++) {
        var trackNumber = parseInt(group[i], 10) || 0;
        if (trackNumber > 0 && !seen[trackNumber]) {
            seen[trackNumber] = true;
            result.push(trackNumber);
        }
    }

    result.sort(function (a, b) { return a - b; });
    return result;
}

function ebBuildTrackGroupLabel(group) {
    var normalized = ebNormalizeTrackGroup(group);
    return normalized.join("-");
}

function ebBuildOutputPath(folderPath, sequenceName, suffix, extension, isRebackup) {
    var tempPart = isRebackup ? "_REBKP_TEMP" : "";
    return ebToFsPath(folderPath + "\\" + sequenceName + suffix + tempPart + extension);
}

function ebBuildRequestedOutputFiles(sequence, folderPath, videoPresetPath, audioPresetPath, audioFormat, selectedItemsJson, isRebackup, rebackupLayout) {
    var requested = [];
    var selection = ebGetSelectedExportItems(selectedItemsJson);
    var sequenceName = ebGetSequenceExportBaseName(sequence);
    var i;
    var audioExtension = ebGetExportExtension(sequence, audioPresetPath, "." + audioFormat);
    var shouldIncludeVideo = selection.includeVideo;

    if (shouldIncludeVideo) {
        var videoExtension = ebGetExportExtension(sequence, videoPresetPath, ".mp4");
        var existingVideoFinalPath = isRebackup && rebackupLayout && rebackupLayout.video && rebackupLayout.video.mediaPath
            ? ebStripRebackupTempMarkerFromPath(rebackupLayout.video.mediaPath)
            : "";
        var resolvedVideoFinalPath = existingVideoFinalPath || ebToFsPath(folderPath + "\\" + sequenceName + "_BACKUP" + videoExtension);
        requested.push({
            kind: "video",
            trackNumber: 0,
            path: isRebackup && existingVideoFinalPath
                ? ebBuildRebackupTempPathFromFinalPath(resolvedVideoFinalPath)
                : ebBuildOutputPath(folderPath, sequenceName, "_BACKUP", videoExtension, isRebackup),
            finalPath: resolvedVideoFinalPath
        });
    }

    if (isRebackup) {
        var rebackupAudioDefinitions = ebGetRebackupAudioDefinitions(sequence, sequenceName, rebackupLayout);
        for (i = 0; i < rebackupAudioDefinitions.length; i++) {
            if (!ebRebackupDefinitionIsSelected(rebackupAudioDefinitions[i], selection)) {
                continue;
            }

            var rebackupTrackNumbers = rebackupAudioDefinitions[i].trackNumbers;
            var rebackupGroupLabel = ebBuildTrackGroupLabel(rebackupTrackNumbers);
            var existingAudioFinalPath = rebackupAudioDefinitions[i].mediaPath
                ? ebStripRebackupTempMarkerFromPath(rebackupAudioDefinitions[i].mediaPath)
                : "";
            var resolvedAudioFinalPath = existingAudioFinalPath
                ? ebReplacePathExtension(existingAudioFinalPath, audioExtension)
                : ebToFsPath(folderPath + "\\" + sequenceName + "_Track" + rebackupGroupLabel + audioExtension);
            var oldAudioFinalPath = existingAudioFinalPath && ebNormalizeMediaPathForComparison(existingAudioFinalPath) !== ebNormalizeMediaPathForComparison(resolvedAudioFinalPath)
                ? existingAudioFinalPath
                : "";
            requested.push({
                kind: "audio",
                trackNumber: rebackupTrackNumbers[0],
                trackNumbers: rebackupTrackNumbers,
                path: ebBuildRebackupTempPathFromFinalPath(resolvedAudioFinalPath),
                finalPath: resolvedAudioFinalPath,
                oldFinalPath: oldAudioFinalPath
            });
        }

        return requested;
    }

    for (i = 0; i < selection.audioGroups.length; i++) {
        var group = ebNormalizeTrackGroup(selection.audioGroups[i]);
        if (group.length < 2) {
            continue;
        }

        var groupLabel = ebBuildTrackGroupLabel(group);
        requested.push({
            kind: "audio",
            trackNumber: group[0],
            trackNumbers: group,
            path: ebBuildOutputPath(folderPath, sequenceName, "_Track" + groupLabel, audioExtension, isRebackup),
            finalPath: ebToFsPath(folderPath + "\\" + sequenceName + "_Track" + groupLabel + audioExtension)
        });
    }

    for (i = 0; i < ebGetTrackCount(sequence.audioTracks); i++) {
        var trackNumber = i + 1;
        if (!ebTrackHasClips(sequence.audioTracks[i])) {
            continue;
        }

        if (ebTrackIsInGroups(trackNumber, selection.audioGroups)) {
            continue;
        }

        if (selection.hasAudioSelection) {
            if (!ebTrackIsInList(trackNumber, selection.audioTracks)) {
                continue;
            }
        }

        requested.push({
            kind: "audio",
            trackNumber: trackNumber,
            trackNumbers: [trackNumber],
            path: ebBuildOutputPath(folderPath, sequenceName, "_Track" + trackNumber, audioExtension, isRebackup),
            finalPath: ebToFsPath(folderPath + "\\" + sequenceName + "_Track" + trackNumber + audioExtension)
        });
    }

    return requested;
}

function ebFindExistingOutputConflicts(requestedFiles) {
    var conflicts = [];
    var i;

    for (i = 0; i < requestedFiles.length; i++) {
        if (requestedFiles[i] && requestedFiles[i].path && ebFileExists(requestedFiles[i].path)) {
            conflicts.push(requestedFiles[i]);
        }
    }

    return conflicts;
}

function ebFindExistingProjectConflicts(sequence, requestedFiles, baseName) {
    var conflicts = [];
    var managedSelection = ebGetSequenceManagedSelection(sequence, baseName);
    var i;

    for (i = 0; i < requestedFiles.length; i++) {
        var requested = requestedFiles[i];
        if (!requested) {
            continue;
        }

        if (requested.path && ebFindProjectItemByMediaPath(app.project.rootItem, ebToFsPath(requested.finalPath || requested.path))) {
            conflicts.push({
                kind: requested.kind,
                path: requested.finalPath || requested.path,
                reason: "media"
            });
            continue;
        }

        if (requested.kind === "video" && managedSelection.hasBackup) {
            conflicts.push({
                kind: "video",
                path: requested.finalPath || requested.path,
                reason: "track"
            });
            continue;
        }

        if (requested.kind === "audio") {
            var trackNumbers = requested.trackNumbers && requested.trackNumbers.length ? requested.trackNumbers : [requested.trackNumber];
            for (var j = 0; j < trackNumbers.length; j++) {
                var trackNumber = parseInt(trackNumbers[j], 10) || 0;
                if (managedSelection.trackNumbers[trackNumber] || managedSelection.backupTrackNumbers[trackNumber]) {
                    conflicts.push({
                        kind: "audio",
                        trackNumber: trackNumber,
                        path: requested.finalPath || requested.path,
                        reason: "track"
                    });
                    break;
                }
            }
        }
    }

    return conflicts;
}

function ebBuildQueuedFile(kind, path, trackNumber) {
    return {
        kind: kind,
        path: ebToFsPath(path),
        trackNumber: trackNumber || 0,
        name: new File(ebToFsPath(path)).name
    };
}

function ebFormatConflictMessage(conflicts) {
    var lines = ["Media already exists."];
    var seen = {};
    var i;
    var conflictPath;

    for (i = 0; i < conflicts.length; i++) {
        conflictPath = conflicts[i] && conflicts[i].path ? String(conflicts[i].path) : "";
        if (conflictPath && !seen[conflictPath]) {
            seen[conflictPath] = true;
            lines.push(conflictPath);
        }
    }

    return lines.join("\n");
}

function ebBuildQueuedFileFromRequested(requested) {
    return {
        kind: requested.kind,
        path: ebToFsPath(requested.path),
        finalPath: ebToFsPath(requested.finalPath || requested.path),
        trackNumber: requested.trackNumber || 0,
        trackNumbers: requested.trackNumbers || [],
        name: new File(ebToFsPath(requested.path)).name
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
        var sequenceBaseName;
        var managedSelection;
        var items;
        var i;

        if (!sequence) {
            return ebResult(false, "No active sequence is open in Premiere Pro.");
        }

        sequenceBaseName = ebGetSequenceExportBaseName(sequence);
        managedSelection = ebGetSequenceManagedSelection(sequence, sequenceBaseName);

        items = [{
            kind: "video",
            label: "Backup MP4",
            selected: !managedSelection.hasBackup,
            locked: false,
            trackNumber: 0
        }];

        for (i = 0; i < ebGetTrackCount(sequence.audioTracks); i++) {
            if (!ebTrackHasClips(sequence.audioTracks[i])) {
                continue;
            }

            var currentTrackName = ebGetTrackName(sequence.audioTracks[i]);
            var isManagedTrack = !!managedSelection.trackNumbers[i + 1] || !!managedSelection.backupTrackNumbers[i + 1];
            if (isManagedTrack) {
                continue;
            }

            items.push({
                kind: "audio",
                label: "Track " + (i + 1),
                selected: true,
                locked: false,
                trackNumber: i + 1,
                trackName: currentTrackName
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

exportBackup.validateBackupExportSettings = function (backupVideoTrackNumber, folderPath, videoPresetPath, mp3PresetPath, wavPresetPath, audioFormat, selectedItemsJson, allowExistingFiles, autoEmptyTrack) {
    try {
        var sequence = ebGetActiveSequence();
        var resolvedAudioFormat;
        var audioPresetPath;
        var requestedFiles;
        var conflicts;
        var names;
        if (!sequence) {
            return ebResult(false, "No active sequence is open in Premiere Pro.");
        }

        var sequenceBaseName = ebGetSequenceExportBaseName(sequence);
        var isRebackupMode = allowExistingFiles === true || String(allowExistingFiles).toLowerCase() === "true";
        if (isRebackupMode) {
            if (ebFindManagedBackupVideoTrackNumber(sequence, sequenceBaseName) < 1) {
                return ebResult(false, "No existing backup MP4 track was found for re-backup.");
            }
        } else {
            if (ebFindManagedBackupVideoTrackNumber(sequence, sequenceBaseName) > 0) {
                return ebResult(false, "Backup files are already there. Use Re-backup.");
            }

            backupVideoTrackNumber = ebResolveBackupVideoTrackNumber(sequence, backupVideoTrackNumber, autoEmptyTrack, true);
            ebValidateBackupTrack(sequence, backupVideoTrackNumber, true, sequenceBaseName);
        }

        if (folderPath) {
            resolvedAudioFormat = String(audioFormat || "mp3").toLowerCase() === "wav" ? "wav" : "mp3";
            audioPresetPath = resolvedAudioFormat === "wav" ? wavPresetPath : mp3PresetPath;

            ebCheckPreset(videoPresetPath, "Video");
            ebCheckPreset(audioPresetPath, resolvedAudioFormat.toUpperCase());

            requestedFiles = ebBuildRequestedOutputFiles(
                sequence,
                folderPath,
                videoPresetPath,
                audioPresetPath,
                resolvedAudioFormat,
                selectedItemsJson,
                isRebackupMode,
                isRebackupMode ? ebCaptureRebackupLayout(sequence, sequenceBaseName) : null
            );
            conflicts = ebFindExistingOutputConflicts(requestedFiles);
            conflicts = conflicts.concat(ebFindExistingProjectConflicts(sequence, requestedFiles, sequenceBaseName));

            if (conflicts.length && !isRebackupMode) {
                names = [];
                for (var i = 0; i < conflicts.length; i++) {
                    names.push(new File(conflicts[i].path).name + (conflicts[i].reason ? " (" + conflicts[i].reason + ")" : ""));
                }

                return ebResult(false, "Media already exists.\n" + names.join("\n"), {
                    conflicts: conflicts,
                    hasConflicts: true
                });
            }
        }

        return ebResult(true, "OK", {
            backupVideoTrackNumber: parseInt(backupVideoTrackNumber, 10) || 0
        });
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
            suggestedVideoTrack: sequence ? (ebFindLowestEmptyVideoTrackNumber(sequence) || suggestedVideoTrack || 1) : 1
        });
    } catch (e) {
        return ebResult(false, e.toString());
    }
};

exportBackup.runBackupQueue = function (folderPath, videoPresetPath, mp3PresetPath, wavPresetPath, audioFormat, backupVideoTrackNumber, removeSequenceMarkers, selectedItemsJson, exportMode, isRebackup, autoEmptyTrack) {
    var sequence = null;
    var originalMuteStates = [];
    var originalVideoMuteStates = [];
    var audioStateFinalized = false;

    try {
        sequence = ebGetActiveSequence();
        if (!sequence) {
            return ebResult(false, "No active sequence is open in Premiere Pro.");
        }

        if (!folderPath || !ebEnsureFolder(folderPath)) {
            return ebResult(false, "Could not create or access the export folder.");
        }

        if (!ebSequenceHasInOut(sequence)) {
            return ebResult(false, "Set sequence In and Out points first, then run ExportBackup again.");
        }

        var sequenceName = ebGetSequenceExportBaseName(sequence);
        var shouldRebackup = isRebackup === true || String(isRebackup).toLowerCase() === "true";
        var existingBackupVideoTrackNumber = shouldRebackup ? ebFindManagedBackupVideoTrackNumber(sequence, sequenceName) : 0;
        if (shouldRebackup && existingBackupVideoTrackNumber < 1) {
            return ebResult(false, "No existing backup MP4 track was found for re-backup.");
        }
        var rebackupLayout = shouldRebackup ? ebCaptureRebackupLayout(sequence, sequenceName) : null;

        var resolvedAudioFormat = String(audioFormat || "mp3").toLowerCase() === "wav" ? "wav" : "mp3";
        var audioPresetPath = resolvedAudioFormat === "wav" ? wavPresetPath : mp3PresetPath;
        var audioLabel = resolvedAudioFormat.toUpperCase();
        var resolvedBackupVideoTrackNumber = shouldRebackup
            ? existingBackupVideoTrackNumber
            : ebResolveBackupVideoTrackNumber(sequence, backupVideoTrackNumber, autoEmptyTrack, true);
        originalMuteStates = ebCaptureMuteStates(sequence);
        originalVideoMuteStates = ebCaptureVideoMuteStates(sequence);
        var workAreaType = 1;
        var notes = [];
        var queuedFiles = [];
        var selectedItems = ebGetSelectedExportItems(selectedItemsJson);
        var selectedAudioTracks = selectedItems.audioTracks;
        var includeBackupVideo = selectedItems.includeVideo;
        var shouldRemoveSequenceMarkers = removeSequenceMarkers !== false && String(removeSequenceMarkers).toLowerCase() !== "false";
        var resolvedExportMode = String(exportMode || "").toLowerCase() === "premiere" ? "premiere" : "mediaEncoder";
        var requestedFiles;
        var conflicts;
        var i;

        if (!shouldRebackup) {
            ebValidateBackupTrack(sequence, resolvedBackupVideoTrackNumber, true, sequenceName);
        }

        ebCheckPreset(videoPresetPath, "Video");
        ebCheckPreset(audioPresetPath, audioLabel);
        requestedFiles = ebBuildRequestedOutputFiles(
            sequence,
            folderPath,
            videoPresetPath,
            audioPresetPath,
            resolvedAudioFormat,
            selectedItemsJson,
            shouldRebackup,
            rebackupLayout
        );
        if (shouldRebackup) {
            selectedAudioTracks = [];
            var selectedAudioTrackMap = {};
            for (i = 0; i < requestedFiles.length; i++) {
                if (!requestedFiles[i] || requestedFiles[i].kind !== "audio") {
                    continue;
                }

                var requestedTrackNumbers = requestedFiles[i].trackNumbers && requestedFiles[i].trackNumbers.length
                    ? requestedFiles[i].trackNumbers
                    : [requestedFiles[i].trackNumber];
                for (var requestedTrackIndex = 0; requestedTrackIndex < requestedTrackNumbers.length; requestedTrackIndex++) {
                    var requestedTrackNumber = parseInt(requestedTrackNumbers[requestedTrackIndex], 10) || 0;
                    if (requestedTrackNumber > 0 && !selectedAudioTrackMap[requestedTrackNumber]) {
                        selectedAudioTrackMap[requestedTrackNumber] = true;
                        selectedAudioTracks.push(requestedTrackNumber);
                    }
                }
            }

            selectedAudioTracks.sort(function (a, b) { return a - b; });
            notes.push(
                selectedAudioTracks.length
                    ? "Re-backup will export checked original source audio from A" + selectedAudioTracks.join(", A") + "."
                    : "Re-backup has no checked audio tracks to export."
            );
            if (rebackupLayout.video) {
                notes.push(
                    "Recorded backup video position: V" + rebackupLayout.video.targetTrackNumber +
                    " at " + rebackupLayout.video.startSeconds + "s."
                );
            }
            if (rebackupLayout.backupAudio) {
                notes.push(
                    "Recorded backup-video audio position: A" + rebackupLayout.backupAudio.targetTrackNumber +
                    " at " + rebackupLayout.backupAudio.startSeconds + "s."
                );
            }
            for (i = 0; i < rebackupLayout.audioOutputs.length; i++) {
                var mappedOutput = rebackupLayout.audioOutputs[i];
                notes.push(
                    "Recorded Track" + mappedOutput.sourceTrackNumbers.join("-") +
                    " position: A" + mappedOutput.targetTrackNumber +
                    " at " + mappedOutput.startSeconds + "s."
                );
            }
        }
        conflicts = ebFindExistingOutputConflicts(requestedFiles);
        conflicts = conflicts.concat(ebFindExistingProjectConflicts(sequence, requestedFiles, sequenceName));
        if (conflicts.length && !shouldRebackup) {
            return ebResult(false, ebFormatConflictMessage(conflicts), {
                conflicts: conflicts,
                hasConflicts: true
            });
        }

        if (shouldRemoveSequenceMarkers) {
            var removedMarkerCount = ebRemoveAllSequenceMarkers(sequence);
            notes.push("Removed " + removedMarkerCount + " sequence marker" + (removedMarkerCount === 1 ? "" : "s") + " before export.");
        }

        if (ebRemoveUnusedMedia()) {
            notes.push("Removed unused media before export.");
        }

        if (resolvedExportMode === "mediaEncoder") {
            app.encoder.launchEncoder();
            $.sleep(EB_ENCODER_LAUNCH_WAIT_MS);
        }

        var clearedSoloCount = ebClearAllAudioSoloStates(sequence);
        if (clearedSoloCount > 0) {
            notes.push("Cleared audio solo states before export.");
        }

        ebSetAllTrackMutes(sequence, 0);
        ebApplyManagedTrackMutePolicy(sequence, sequenceName);

        if (includeBackupVideo) {
            var videoRequest = null;
            for (i = 0; i < requestedFiles.length; i++) {
                if (requestedFiles[i].kind === "video") {
                    videoRequest = requestedFiles[i];
                    break;
                }
            }

            var videoPath = videoRequest ? videoRequest.path : "";
            var hiddenVideoTrackCount = ebHideVideoTracksAbove(sequence, resolvedBackupVideoTrackNumber);
            ebClearAllAudioSoloStates(sequence);
            if (resolvedExportMode === "premiere") {
                ebExportSequenceDirect(sequence, videoPath, videoPresetPath, workAreaType);
            } else {
                var videoJobId = ebQueueSequenceWithSettle(sequence, videoPath, videoPresetPath, workAreaType);
                if (!videoJobId || videoJobId === "0") {
                    ebRestoreMuteStates(sequence, originalMuteStates);
                    return ebResult(false, "Could not queue the MP4 export in Adobe Media Encoder.");
                }
            }

            queuedFiles.push(ebBuildQueuedFileFromRequested(videoRequest));
            notes.push((resolvedExportMode === "premiere" ? "Rendered" : "Queued") + " MP4 backup: " + videoPath);
            if (hiddenVideoTrackCount > 0) {
                notes.push("Video tracks above V" + resolvedBackupVideoTrackNumber + " were hidden while the backup MP4 queue item was created.");
            }
        } else {
            notes.push("Skipped backup MP4 export.");
        }

        for (i = 0; i < requestedFiles.length; i++) {
            var audioRequest = requestedFiles[i];
            var trackNumbers;
            var firstTrackNumber;
            var k;

            if (!audioRequest || audioRequest.kind !== "audio") {
                continue;
            }

            trackNumbers = audioRequest.trackNumbers && audioRequest.trackNumbers.length
                ? audioRequest.trackNumbers
                : [audioRequest.trackNumber];
            firstTrackNumber = parseInt(trackNumbers[0], 10) || parseInt(audioRequest.trackNumber, 10) || 1;

            ebSetAllTrackMutes(sequence, 1);
            for (k = 0; k < trackNumbers.length; k++) {
                var audibleTrackNumber = parseInt(trackNumbers[k], 10) || 0;
                if (
                    audibleTrackNumber > 0 &&
                    ebIsOriginalAudioTrack(sequence, sequenceName, audibleTrackNumber, false) &&
                    sequence.audioTracks[audibleTrackNumber - 1] &&
                    sequence.audioTracks[audibleTrackNumber - 1].setMute
                ) {
                    sequence.audioTracks[audibleTrackNumber - 1].setMute(0);
                }
            }

            var audioPath = audioRequest.path;
            ebClearAllAudioSoloStates(sequence);
            if (resolvedExportMode === "premiere") {
                ebExportSequenceDirect(sequence, audioPath, audioPresetPath, workAreaType);
            } else {
                var audioJobId = ebQueueSequenceWithSettle(sequence, audioPath, audioPresetPath, workAreaType);

                if (!audioJobId || audioJobId === "0") {
                    notes.push("Failed to queue " + audioLabel + " for Track" + firstTrackNumber + ".");
                    continue;
                }
            }

            queuedFiles.push(ebBuildQueuedFileFromRequested(audioRequest));
            notes.push((resolvedExportMode === "premiere" ? "Rendered" : "Queued") + " " + audioLabel + " track export: " + audioPath);
        }

        ebRestoreMuteStates(sequence, originalMuteStates);
        ebApplyManagedTrackMutePolicy(sequence, sequenceName);
        audioStateFinalized = true;

        if (resolvedExportMode === "mediaEncoder") {
            try {
                if (app.encoder.startBatch) {
                    app.encoder.startBatch();
                }
            } catch (e) {}
        }

        return ebResult(true, (resolvedExportMode === "premiere" ? "Rendered files: " : "Queued jobs: ") + queuedFiles.length + ".\n" + notes.join("\n"), {
            sequenceName: ebGetSequenceName(sequence),
            baseName: sequenceName,
            backupVideoTrackNumber: resolvedBackupVideoTrackNumber,
            audioFormat: resolvedAudioFormat,
            exportMode: resolvedExportMode,
            rebackup: shouldRebackup,
            rebackupLayout: rebackupLayout,
            queuedFiles: queuedFiles,
            projectName: app.project && app.project.name ? app.project.name : "",
            projectPath: app.project && app.project.path ? app.project.path : ""
        });
    } catch (e) {
        return ebResult(false, e.toString());
    } finally {
        // Video layers are hidden only while export jobs are created. Always
        // restore every layer to its exact pre-export state, including layers
        // that were already hidden before ExportBackup started.
        if (sequence && originalVideoMuteStates.length) {
            ebRestoreVideoMuteStates(sequence, originalVideoMuteStates);
        }
        if (sequence && originalMuteStates.length && !audioStateFinalized) {
            ebRestoreMuteStates(sequence, originalMuteStates);
        }
    }
};

function ebGetRebackupReleasePaths(expectedFiles) {
    var releasePaths = [];
    var seenReleasePaths = {};
    var files = expectedFiles || [];
    var i;

    function addPath(pathValue) {
        if (!pathValue) {
            return;
        }

        var fsPath = ebToFsPath(pathValue);
        var pathKey = ebNormalizeMediaPathForComparison(fsPath);
        if (pathKey && !seenReleasePaths[pathKey]) {
            seenReleasePaths[pathKey] = true;
            releasePaths.push(fsPath);
        }
    }

    for (i = 0; i < files.length; i++) {
        if (!files[i]) {
            continue;
        }

        // A render preset may automatically import the TEMP output into
        // Premiere. Remove that project item before the filesystem rename or it
        // becomes an offline orphan. The old final-path item must also go.
        addPath(files[i].path);
        addPath(files[i].finalPath);
        addPath(files[i].oldFinalPath);
    }

    return releasePaths;
}

exportBackup.prepareRebackupReplacement = function (expectedFilesJson) {
    try {
        var sequence = ebGetActiveSequence();
        if (!sequence) {
            return ebResult(false, "No active sequence is open in Premiere Pro.");
        }

        var expectedFiles = expectedFilesJson ? JSON.parse(expectedFilesJson) : [];
        var sequenceBaseName = ebGetSequenceExportBaseName(sequence);
        var backupVideoTrackNumber = ebFindManagedBackupVideoTrackNumber(sequence, sequenceBaseName);
        var releasePaths = ebGetRebackupReleasePaths(expectedFiles);
        var releaseResults = [];
        var offlinedItems = 0;
        var removedItems = 0;
        var removedTimelineClips = 0;
        var sourceMonitorClosed = false;
        var remainingOnlinePaths = [];
        var remainingProjectPaths = [];
        var i;

        try {
            if (app.sourceMonitor && app.sourceMonitor.closeAllClips) {
                sourceMonitorClosed = app.sourceMonitor.closeAllClips() === 0;
            }
        } catch (sourceMonitorError) {}

        if (backupVideoTrackNumber > 0 && sequence.videoTracks[backupVideoTrackNumber - 1]) {
            removedTimelineClips += ebRemoveManagedClipsFromTrack(sequence.videoTracks[backupVideoTrackNumber - 1], sequenceBaseName, "backup", 0);
        }

        removedTimelineClips += ebRemoveAllManagedClipsFromAllAudioTracks(sequence, sequenceBaseName);

        // Remove every old backup reference from every sequence before touching
        // the source files. This includes clips outside the active sequence.
        removedTimelineClips += ebRemoveProjectClipsByMediaPaths(releasePaths);

        for (i = 0; i < releasePaths.length; i++) {
            var releaseResult = ebReleaseProjectItemsByMediaPath(releasePaths[i]);
            releaseResults.push(releaseResult);
            offlinedItems += releaseResult.offlined;
            removedItems += releaseResult.removed;
        }

        try {
            ebRemoveUnusedMedia();
        } catch (cleanupError) {}

        $.sleep(EB_MEDIA_RELEASE_WAIT_MS);

        for (i = 0; i < releasePaths.length; i++) {
            if (ebGetOnlineProjectItemCountByMediaPath(releasePaths[i]) > 0) {
                remainingOnlinePaths.push(releasePaths[i]);
            }
            if (releaseResults[i] && releaseResults[i].remaining > 0) {
                remainingProjectPaths.push(releasePaths[i]);
            }
        }

        if (remainingProjectPaths.length) {
            return ebResult(false, "Premiere Pro still contains old backup project items, so local files were not replaced.", {
                sourceMonitorClosed: sourceMonitorClosed,
                removedTimelineClips: removedTimelineClips,
                offlinedProjectItems: offlinedItems,
                removedProjectItems: removedItems,
                remainingOnlinePaths: remainingOnlinePaths,
                remainingProjectPaths: remainingProjectPaths,
                releaseResults: releaseResults
            });
        }

        return ebResult(true, "Released old backup media for replacement.", {
            sourceMonitorClosed: sourceMonitorClosed,
            removedTimelineClips: removedTimelineClips,
            offlinedProjectItems: offlinedItems,
            removedProjectItems: removedItems,
            releaseResults: releaseResults
        });
    } catch (e) {
        return ebResult(false, e.toString());
    }
};

function ebGetAudioEntryTrackNumbers(audioEntry) {
    if (!audioEntry) {
        return [];
    }

    var trackNumbers = audioEntry.trackNumbers && audioEntry.trackNumbers.length
        ? audioEntry.trackNumbers
        : [audioEntry.trackNumber];
    return ebNormalizeTrackGroup(trackNumbers);
}

function ebFindRebackupAudioPlacement(rebackupLayout, audioEntry) {
    var outputs = rebackupLayout && rebackupLayout.audioOutputs ? rebackupLayout.audioOutputs : [];
    var entryTrackNumbers = ebGetAudioEntryTrackNumbers(audioEntry);
    var entryKey = entryTrackNumbers.join("-");
    var i;

    for (i = 0; i < outputs.length; i++) {
        var outputTrackNumbers = outputs[i].sourceTrackNumbers && outputs[i].sourceTrackNumbers.length
            ? outputs[i].sourceTrackNumbers
            : [outputs[i].sourceTrackNumber];
        if (ebNormalizeTrackGroup(outputTrackNumbers).join("-") === entryKey) {
            return outputs[i];
        }
    }

    for (i = 0; i < outputs.length; i++) {
        if ((parseInt(outputs[i].sourceTrackNumber, 10) || 0) === (entryTrackNumbers[0] || 0)) {
            return outputs[i];
        }
    }

    return null;
}

function ebGetNextUnreservedTrackNumber(firstTrackNumber, reservedTrackNumbers) {
    var candidate = Math.max(1, parseInt(firstTrackNumber, 10) || 1);
    while (reservedTrackNumbers[candidate]) {
        candidate += 1;
    }
    return candidate;
}

exportBackup.alignMappedFiles = function (videoPath, audioJson, backupVideoTrackNumber, sortProjectFiles, rebackupLayoutJson) {
    try {
        var sequence = ebGetActiveSequence();
        if (!sequence) {
            return ebResult(false, "No active sequence is open in Premiere Pro.");
        }

        var audioEntries = [];
        var rebackupLayout = null;
        if (audioJson) {
            audioEntries = JSON.parse(audioJson);
        }
        if (rebackupLayoutJson) {
            rebackupLayout = JSON.parse(rebackupLayoutJson);
        }

        if (!videoPath && (!audioEntries || !audioEntries.length)) {
            return ebResult(false, "No matched backup video or audio files were provided.");
        }

        var recordedBackupVideoTrack = rebackupLayout && rebackupLayout.video
            ? parseInt(rebackupLayout.video.targetTrackNumber, 10) || 0
            : 0;
        var resolvedBackupTrack = recordedBackupVideoTrack > 0
            ? recordedBackupVideoTrack
            : Math.max(1, parseInt(backupVideoTrackNumber, 10) || 1);
        if (resolvedBackupTrack > ebGetTrackCount(sequence.videoTracks)) {
            return ebResult(false, "V" + resolvedBackupTrack + " does not exist in the active sequence.");
        }
        var sequenceBaseName = ebGetSequenceExportBaseName(sequence);
        if (videoPath) {
            ebRemoveManagedClipsFromTrack(sequence.videoTracks[resolvedBackupTrack - 1], sequenceBaseName, "backup", 0);
            ebRemoveManagedClipsFromAllAudioTracks(sequence, sequenceBaseName, "backup", 0);
        }

        if (videoPath && ebTrackHasClips(sequence.videoTracks[resolvedBackupTrack - 1])) {
            return ebResult(false, "V" + resolvedBackupTrack + " is not empty.");
        }

        var firstImportAudioTrack = ebGetHighestSourceAudioTrackNumber(sequence, sequenceBaseName) + 1;
        var recordedBackupAudioTrack = rebackupLayout && rebackupLayout.backupAudio
            ? parseInt(rebackupLayout.backupAudio.targetTrackNumber, 10) || 0
            : 0;
        var backupVideoAudioTrackNumber = videoPath
            ? (recordedBackupAudioTrack > 0 ? recordedBackupAudioTrack : firstImportAudioTrack)
            : 0;
        var reservedTrackNumbers = {};
        var assignedTrackNumbers = {};
        var audioTargetTrackNumbers = [];
        var audioPlacements = [];
        var nextFallbackAudioTrack = firstImportAudioTrack;
        var firstOtherAudioTrack = 0;
        var finalRequiredAudioTrack = backupVideoAudioTrackNumber;
        var notes = [];
        var importBin = ebGetImportBin(sequence);
        var organizerNotes;
        var i;

        if (backupVideoAudioTrackNumber > 0) {
            reservedTrackNumbers[backupVideoAudioTrackNumber] = true;
            assignedTrackNumbers[backupVideoAudioTrackNumber] = true;
        }

        if (rebackupLayout && rebackupLayout.audioOutputs) {
            for (i = 0; i < rebackupLayout.audioOutputs.length; i++) {
                var layoutTargetTrack = parseInt(rebackupLayout.audioOutputs[i].targetTrackNumber, 10) || 0;
                if (layoutTargetTrack > 0) {
                    reservedTrackNumbers[layoutTargetTrack] = true;
                }
            }
        }

        if (audioEntries && audioEntries.length) {
            for (i = 0; i < audioEntries.length; i++) {
                var recordedAudioPlacement = ebFindRebackupAudioPlacement(rebackupLayout, audioEntries[i]);
                var recordedAudioTrack = recordedAudioPlacement
                    ? parseInt(recordedAudioPlacement.targetTrackNumber, 10) || 0
                    : 0;

                if (recordedAudioTrack > 0 && !assignedTrackNumbers[recordedAudioTrack]) {
                    audioTargetTrackNumbers[i] = recordedAudioTrack;
                    audioPlacements[i] = recordedAudioPlacement;
                } else {
                    nextFallbackAudioTrack = ebGetNextUnreservedTrackNumber(nextFallbackAudioTrack, reservedTrackNumbers);
                    audioTargetTrackNumbers[i] = nextFallbackAudioTrack;
                    audioPlacements[i] = null;
                    nextFallbackAudioTrack += 1;
                }

                reservedTrackNumbers[audioTargetTrackNumbers[i]] = true;
                assignedTrackNumbers[audioTargetTrackNumbers[i]] = true;
                if (firstOtherAudioTrack < 1 || audioTargetTrackNumbers[i] < firstOtherAudioTrack) {
                    firstOtherAudioTrack = audioTargetTrackNumbers[i];
                }
            }

            for (i = 0; i < audioTargetTrackNumbers.length; i++) {
                finalRequiredAudioTrack = Math.max(finalRequiredAudioTrack, audioTargetTrackNumbers[i]);
            }
        }

        if (finalRequiredAudioTrack > 0) {
            ebEnsureAudioTrackCount(sequence, finalRequiredAudioTrack);
        }

        if (videoPath) {
            var videoItem = ebImportProjectItem(videoPath, importBin);
            if (!videoItem) {
                return ebResult(false, "Could not import backup video: " + videoPath);
            }

            var videoWhen = ebCreateTimeFromPlacement(rebackupLayout && rebackupLayout.video ? rebackupLayout.video : null);
            if (sequence.overwriteClip) {
                sequence.overwriteClip(videoItem, videoWhen.seconds, resolvedBackupTrack - 1, backupVideoAudioTrackNumber - 1);
            } else {
                sequence.videoTracks[resolvedBackupTrack - 1].overwriteClip(videoItem, videoWhen);
            }

            ebSetTrackName(sequence.audioTracks[backupVideoAudioTrackNumber - 1], sequenceBaseName + "_BACKUP");
            if (sequence.audioTracks[backupVideoAudioTrackNumber - 1] && sequence.audioTracks[backupVideoAudioTrackNumber - 1].setMute) {
                sequence.audioTracks[backupVideoAudioTrackNumber - 1].setMute(0);
            }

            notes.push(
                "Aligned backup MP4 to V" + resolvedBackupTrack +
                " and its audio to A" + backupVideoAudioTrackNumber +
                " at " + videoWhen.seconds + "s."
            );
        }

        for (i = 0; i < audioEntries.length; i++) {
            var targetTrackNumber = audioTargetTrackNumbers[i];
            var audioEntryTrackNumbers = ebGetAudioEntryTrackNumbers(audioEntries[i]);
            var firstSourceTrackNumber = audioEntryTrackNumbers.length
                ? audioEntryTrackNumbers[0]
                : (parseInt(audioEntries[i].trackNumber, 10) || 0);
            ebRemoveManagedClipsFromAllAudioTracks(sequence, sequenceBaseName, "audio", firstSourceTrackNumber);
            var audioItem = ebImportProjectItem(audioEntries[i].path, importBin);
            if (!audioItem) {
                return ebResult(false, "Could not import audio file: " + audioEntries[i].path);
            }

            var audioWhen = ebCreateTimeFromPlacement(audioPlacements[i]);
            sequence.audioTracks[targetTrackNumber - 1].overwriteClip(audioItem, audioWhen);
            ebSetTrackName(
                sequence.audioTracks[targetTrackNumber - 1],
                sequenceBaseName + "_Track" + ebBuildTrackGroupLabel(audioEntryTrackNumbers)
            );
            if (sequence.audioTracks[targetTrackNumber - 1] && sequence.audioTracks[targetTrackNumber - 1].setMute) {
                sequence.audioTracks[targetTrackNumber - 1].setMute(1);
            }
            notes.push(
                "Aligned " + (audioEntries[i].name || ("Track" + ebBuildTrackGroupLabel(audioEntryTrackNumbers))) +
                " to A" + targetTrackNumber +
                " at " + audioWhen.seconds + "s."
            );
        }

        if (videoPath && backupVideoAudioTrackNumber > 0) {
            ebSetOnlyTrackAudible(sequence, backupVideoAudioTrackNumber - 1);
            notes.push("Muted every other audio track; only the backup MP4 audio on A" + backupVideoAudioTrackNumber + " remains audible.");
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
            ebRemoveUnusedMedia();
            if (app.project && app.project.save) {
                app.project.save();
            }
        } catch (e) {}

        return ebResult(true, (rebackupLayout ? "Re-backup alignment restored.\n" : "Alignment completed.\n") + notes.join("\n"), {
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
