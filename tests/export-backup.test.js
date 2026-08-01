const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function makeCollection(items, countProperty) {
    const collection = {};

    function sync() {
        Object.keys(collection).forEach((key) => {
            if (/^\d+$/.test(key)) {
                delete collection[key];
            }
        });
        items.forEach((item, index) => {
            collection[index] = item;
        });
        collection[countProperty] = items.length;
    }

    collection.removeItem = (item) => {
        const index = items.indexOf(item);
        if (index >= 0) {
            items.splice(index, 1);
            sync();
        }
    };
    sync();
    return collection;
}

function makeTrack(initialMute, clips, name) {
    let muted = initialMute ? 1 : 0;
    return {
        name: name || '',
        clips: makeCollection(clips || [], 'numItems'),
        isMuted() {
            return muted;
        },
        setMute(value) {
            muted = value ? 1 : 0;
        },
        get muteValue() {
            return muted;
        }
    };
}

function loadHostLogic(appOverrides, contextOverrides) {
    const sourcePath = path.join(__dirname, '..', 'jsx', 'export.jsx');
    const source = fs.readFileSync(sourcePath, 'utf8');
    const app = Object.assign({
        enableQE() {},
        findMenuCommandId() {
            return 0;
        },
        project: {
            rootItem: {
                type: 2,
                children: makeCollection([], 'numItems')
            },
            sequences: makeCollection([], 'numSequences'),
            save() {}
        }
    }, appOverrides || {});
    const contextValues = {
        app,
        console,
        File: function File(filePath) {
            const normalized = String(filePath || '').replace(/\//g, '\\');
            this.fsName = normalized;
            this.name = normalized.split('\\').pop();
            this.exists = true;
        },
        Folder: function Folder(folderPath) {
            this.fsName = String(folderPath || '');
            this.exists = true;
        },
        ProjectItemType: { BIN: 2 },
        Time: function Time() {
            this.seconds = 0;
            this.ticks = '';
        },
        $: {
            sleep() {}
        }
    };
    const context = vm.createContext(Object.assign(contextValues, contextOverrides || {}));

    vm.runInContext(source, context, { filename: sourcePath });
    return { context, source };
}

test('Queue Backup Exports has a visual-only toggle and always starts shown', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');

    assert.match(html, /id="toggleQueueBackupSectionButton"[^>]*>Hide</);
    assert.match(html, /id="queueBackupSectionContent" class="stack"/);
    assert.doesNotMatch(html, /id="queueBackupSectionContent"[^>]*is-hidden/);
    assert.match(
        html,
        /id="audioFormatWav"[\s\S]*?<\/div>\s*<\/div>\s*<div class="stack">\s*<div class="section-label">Queue Preview<\/div>/
    );
    assert.match(mainSource, /function toggleQueueBackupSection\(\)/);
    assert.match(
        mainSource,
        /document\.addEventListener\("DOMContentLoaded",[\s\S]*setQueueBackupSectionVisibility\(true\)/
    );
    assert.doesNotMatch(mainSource, /queueBackupSectionVisibleStorage/i);
});

test('Queue Preview omits app-managed backup audio layers', () => {
    const sourceTrack = makeTrack(0, [{ projectItem: { name: 'Dialogue.wav' } }]);
    const backupTrack = makeTrack(0, [{ projectItem: { name: 'OtherSequence_BACKUP_REBKP_TEMP.mp4' } }]);
    const managedAudioTrack = makeTrack(1, [{ projectItem: { name: 'AD_SM QUOTE No Pain Food_Track1.mp3' } }]);
    const sequence = {
        name: 'Scene',
        audioTracks: makeCollection(
            [sourceTrack, backupTrack, managedAudioTrack],
            'numTracks'
        ),
        videoTracks: makeCollection([], 'numTracks')
    };
    const { context } = loadHostLogic({
        project: {
            activeSequence: sequence,
            rootItem: {
                type: 2,
                children: makeCollection([], 'numItems')
            },
            sequences: makeCollection([sequence], 'numSequences'),
            save() {}
        }
    });

    const result = JSON.parse(context.exportBackup.getExportSelectionInfo());
    assert.equal(result.ok, true);
    assert.deepEqual(
        Array.from(result.items, (item) => item.label),
        ['Backup MP4', 'Track 1']
    );
});

test('re-backup ignores other sequence backup names and finds the active sequence backup', () => {
    const sequence = {
        name: 'Scene',
        videoTracks: makeCollection([
            makeTrack(0, [{ projectItem: { name: 'OtherSequence_BACKUP_REBKP_TEMP.mp4' } }]),
            makeTrack(0, [{ projectItem: { name: 'Rendered Mix.mp4' } }], 'Scene_BACKUP')
        ], 'numTracks'),
        audioTracks: makeCollection([], 'numTracks')
    };
    const { context } = loadHostLogic();
    context.sequenceUnderTest = sequence;

    const result = JSON.parse(vm.runInContext(`JSON.stringify({
        backupTrack: ebFindManagedBackupVideoTrackNumber(sequenceUnderTest, 'Scene'),
        otherSequenceInfo: ebGetTrackSequenceManagedInfo(sequenceUnderTest.videoTracks[0], 'Scene'),
        trackNameInfo: ebGetTrackSequenceManagedInfo(sequenceUnderTest.videoTracks[1], 'Scene')
    })`, context));

    assert.equal(result.backupTrack, 2);
    assert.equal(result.otherSequenceInfo.hasBackup, false);
    assert.equal(result.trackNameInfo.hasBackup, true);
});

test('re-backup output paths follow existing backup media paths', () => {
    const sequence = {
        name: 'Scene',
        audioTracks: makeCollection([
            makeTrack(0, [{ projectItem: { name: 'Dialogue.wav' } }]),
            makeTrack(0, [{ projectItem: { name: 'Scene_BACKUP.mp4', getMediaPath() { return 'E:\\Existing\\Scene_BACKUP.mp4'; } } }], 'Scene_BACKUP'),
            makeTrack(0, [{ projectItem: { name: 'Scene_Track1_REBKP_TEMP.mp3', getMediaPath() { return 'E:\\Existing\\Scene_Track1_REBKP_TEMP.mp3'; } } }], 'Scene_Track1')
        ], 'numTracks'),
        videoTracks: makeCollection([
            makeTrack(0, [{ projectItem: { name: 'Scene_BACKUP.mp4', getMediaPath() { return 'E:\\Existing\\Scene_BACKUP.mp4'; } } }], 'Scene_BACKUP')
        ], 'numTracks')
    };
    const { context } = loadHostLogic();
    context.sequenceUnderTest = sequence;

    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const layout = ebCaptureRebackupLayout(sequenceUnderTest, 'Scene');
        return ebBuildRequestedOutputFiles(
            sequenceUnderTest,
            'D:\\Chosen',
            'video.epr',
            'audio.epr',
            'mp3',
            JSON.stringify({ includeVideo: true, audioTracks: [1], audioGroups: [] }),
            true,
            layout
        );
    })())`, context));

    assert.equal(result[0].finalPath, 'E:\\Existing\\Scene_BACKUP.mp4');
    assert.equal(result[0].path, 'E:\\Existing\\Scene_BACKUP_REBKP_TEMP.mp4');
    assert.equal(result[1].finalPath, 'E:\\Existing\\Scene_Track1.mp3');
    assert.equal(result[1].path, 'E:\\Existing\\Scene_Track1_REBKP_TEMP.mp3');
});

test('Re-backup preserves and relinks old media before freeing the final filename', () => {
    const sourcePath = path.win32.join('D:', 'Backups', 'Scene_BACKUP.mp4');
    const files = new Set([sourcePath]);
    let mediaPath = sourcePath;
    let saveCount = 0;
    let blockedCopyTarget = '';

    function normalizeFilePath(filePath) {
        return path.win32.normalize(String(filePath || ''));
    }

    class MockFile {
        constructor(filePath) {
            this.fsName = normalizeFilePath(filePath);
            this.name = path.win32.basename(this.fsName);
        }

        get exists() {
            return files.has(this.fsName);
        }

        copy(targetPath) {
            if (!files.has(this.fsName)) {
                return false;
            }
            const normalizedTargetPath = normalizeFilePath(targetPath);
            if (blockedCopyTarget && normalizedTargetPath === blockedCopyTarget) {
                return false;
            }
            files.add(normalizedTargetPath);
            return true;
        }

        remove() {
            return files.delete(this.fsName);
        }

        rename(newName) {
            if (!files.has(this.fsName)) {
                return false;
            }
            const renamedPath = path.win32.join(path.win32.dirname(this.fsName), String(newName));
            files.delete(this.fsName);
            files.add(renamedPath);
            this.fsName = renamedPath;
            this.name = path.win32.basename(renamedPath);
            return true;
        }
    }

    const mediaItem = {
        type: 1,
        name: path.win32.basename(sourcePath),
        getMediaPath() {
            return mediaPath;
        },
        canChangeMediaPath() {
            return true;
        },
        changeMediaPath(newPath) {
            mediaPath = normalizeFilePath(newPath);
            return 0;
        }
    };
    const rootItem = {
        type: 2,
        children: makeCollection([mediaItem], 'numItems')
    };
    const { context } = loadHostLogic({
        project: {
            rootItem,
            sequences: makeCollection([], 'numSequences'),
            save() {
                saveCount += 1;
            }
        }
    }, {
        File: MockFile
    });
    const requestedFiles = [{
        kind: 'video',
        path: path.win32.join('D:', 'Backups', 'Scene_BACKUP_REBKP_TEMP.mp4'),
        finalPath: sourcePath,
        sourceMediaPath: sourcePath,
        trackNumber: 0,
        trackNumbers: []
    }];
    context.requestedFilesUnderTest = requestedFiles;

    const result = JSON.parse(vm.runInContext(
        'JSON.stringify(ebPrepareRebackupMediaForDirectExport(requestedFilesUnderTest))',
        context
    ));
    const releasePaths = JSON.parse(vm.runInContext(
        'JSON.stringify(ebGetRebackupReleasePaths(requestedFilesUnderTest))',
        context
    ));
    const preservedPath = requestedFiles[0].preservedPaths[0];
    context.preservedClipUnderTest = {
        projectItem: {
            getMediaPath() {
                return preservedPath;
            }
        }
    };
    const recoveredFinalPath = vm.runInContext(
        'ebGetManagedClipFinalMediaPath(preservedClipUnderTest)',
        context
    );

    assert.equal(result.preservedCount, 1);
    assert.equal(requestedFiles[0].path, sourcePath);
    assert.equal(files.has(sourcePath), false);
    assert.equal(files.has(preservedPath), true);
    assert.equal(mediaPath, preservedPath);
    assert.equal(recoveredFinalPath, sourcePath);
    assert.equal(mediaItem.name, path.win32.basename(sourcePath));
    assert.equal(saveCount, 1);
    assert.ok(releasePaths.includes(sourcePath));
    assert.ok(releasePaths.includes(preservedPath));

    blockedCopyTarget = sourcePath;
    context.failedRollbackUnderTest = [{
        sourcePath,
        preservedPath,
        extraPaths: []
    }];
    vm.runInContext('ebRestoreRebackupPreservations(failedRollbackUnderTest)', context);
    assert.equal(files.has(sourcePath), false);
    assert.equal(files.has(preservedPath), true);
    assert.equal(mediaPath, preservedPath);
});
test('preserved filenames remain recognizable after an interrupted Re-backup', () => {
    const { context } = loadHostLogic();
    context.preservedVideoName = 'Scene_BACKUP_REBKP_OLD_1785000000000_1.mp4';
    context.preservedAudioName = 'Scene_Track1-2_REBKP_OLD_1785000000000_2.wav';

    assert.equal(vm.runInContext(
        'ebIsSequenceManagedBackupTrack(preservedVideoName, "Scene")',
        context
    ), true);
    assert.deepEqual(JSON.parse(vm.runInContext(
        'JSON.stringify(ebGetSequenceManagedAudioTrackNumbers(preservedAudioName, "Scene"))',
        context
    )), [1, 2]);
});
test('re-backup audio format changes render selected format and release old audio path', () => {
    const sequence = {
        name: 'Scene',
        audioTracks: makeCollection([
            makeTrack(0, [{ projectItem: { name: 'Dialogue.wav' } }]),
            makeTrack(0, [{ projectItem: { name: 'Scene_Track1.mp3', getMediaPath() { return 'E:\\Existing\\Scene_Track1.mp3'; } } }], 'Scene_Track1')
        ], 'numTracks'),
        videoTracks: makeCollection([
            makeTrack(0, [{ projectItem: { name: 'Scene_BACKUP.mp4', getMediaPath() { return 'E:\\Existing\\Scene_BACKUP.mp4'; } } }], 'Scene_BACKUP')
        ], 'numTracks')
    };
    const { context } = loadHostLogic();
    context.sequenceUnderTest = sequence;

    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const layout = ebCaptureRebackupLayout(sequenceUnderTest, 'Scene');
        const files = ebBuildRequestedOutputFiles(
            sequenceUnderTest,
            'D:\\Chosen',
            'video.epr',
            'audio.epr',
            'wav',
            JSON.stringify({ includeVideo: true, audioTracks: [1], audioGroups: [] }),
            true,
            layout
        );
        return {
            files,
            releasePaths: ebGetRebackupReleasePaths(files)
        };
    })())`, context));

    assert.equal(result.files[1].path, 'E:\\Existing\\Scene_Track1_REBKP_TEMP.wav');
    assert.equal(result.files[1].finalPath, 'E:\\Existing\\Scene_Track1.wav');
    assert.equal(result.files[1].oldFinalPath, 'E:\\Existing\\Scene_Track1.mp3');
    assert.ok(result.releasePaths.includes('E:\\Existing\\Scene_Track1.mp3'));
});

test('re-backup respects unchecked backup video and checked audio items', () => {
    const sequence = {
        name: 'Scene',
        audioTracks: makeCollection([
            makeTrack(0, [{ projectItem: { name: 'Dialogue A1.wav' } }]),
            makeTrack(0, [{ projectItem: { name: 'Dialogue A2.wav' } }]),
            makeTrack(0, [{ projectItem: { name: 'Scene_Track1.mp3', getMediaPath() { return 'E:\\Existing\\Scene_Track1.mp3'; } } }], 'Scene_Track1'),
            makeTrack(0, [{ projectItem: { name: 'Scene_Track2.mp3', getMediaPath() { return 'E:\\Existing\\Scene_Track2.mp3'; } } }], 'Scene_Track2')
        ], 'numTracks'),
        videoTracks: makeCollection([
            makeTrack(0, [{ projectItem: { name: 'Scene_BACKUP.mp4', getMediaPath() { return 'E:\\Existing\\Scene_BACKUP.mp4'; } } }], 'Scene_BACKUP')
        ], 'numTracks')
    };
    const { context } = loadHostLogic();
    context.sequenceUnderTest = sequence;

    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const layout = ebCaptureRebackupLayout(sequenceUnderTest, 'Scene');
        return ebBuildRequestedOutputFiles(
            sequenceUnderTest,
            'D:\\Chosen',
            'video.epr',
            'audio.epr',
            'mp3',
            JSON.stringify({ includeVideo: false, audioTracks: [2], audioGroups: [] }),
            true,
            layout
        );
    })())`, context));

    assert.equal(result.length, 1);
    assert.equal(result[0].kind, 'audio');
    assert.deepEqual(result[0].trackNumbers, [2]);
    assert.equal(result[0].finalPath, 'E:\\Existing\\Scene_Track2.mp3');
});

test('re-backup can export a new MP4 when the old backup video clip is missing', () => {
    const sequence = {
        name: 'Scene',
        getInPoint() {
            return '0';
        },
        getOutPoint() {
            return '10';
        },
        videoTracks: makeCollection([makeTrack(0)], 'numTracks'),
        audioTracks: makeCollection([], 'numTracks')
    };
    const { context } = loadHostLogic({
        project: {
            activeSequence: sequence,
            rootItem: { type: 2, children: makeCollection([], 'numItems') },
            sequences: makeCollection([sequence], 'numSequences'),
            save() {}
        }
    });

    const result = JSON.parse(context.exportBackup.validateBackupExportSettings(
        1,
        'D:\\Backups',
        'video.epr',
        'mp3.epr',
        'wav.epr',
        'mp3',
        JSON.stringify({ includeVideo: true, audioTracks: [], audioGroups: [] }),
        true,
        false
    ));

    assert.equal(result.ok, true);
    assert.equal(result.backupVideoTrackNumber, 1);
});

test('missing sequence In and Out can be set to the full sequence range', () => {
    let inPoint = 0;
    let outPoint = 0;
    let receivedOutTicks = '';
    const sequence = {
        name: 'Scene',
        end: '508032000000',
        getInPoint() {
            return String(inPoint);
        },
        getOutPoint() {
            return String(outPoint);
        },
        setInPoint(value) {
            inPoint = parseFloat(value) || 0;
        },
        setOutPoint(value) {
            receivedOutTicks = String(value.ticks || '');
            outPoint = 2;
        },
        videoTracks: makeCollection([makeTrack(0)], 'numTracks'),
        audioTracks: makeCollection([], 'numTracks')
    };
    const { context } = loadHostLogic({
        project: {
            activeSequence: sequence,
            rootItem: { type: 2, children: makeCollection([], 'numItems') },
            sequences: makeCollection([sequence], 'numSequences'),
            save() {}
        }
    });

    const missingResult = JSON.parse(context.exportBackup.validateBackupExportSettings(
        1,
        'D:\\Backups',
        'video.epr',
        'mp3.epr',
        'wav.epr',
        'mp3',
        JSON.stringify({ includeVideo: true, audioTracks: [], audioGroups: [] }),
        false,
        false
    ));
    const setResult = JSON.parse(context.exportBackup.setActiveSequenceInOutToFullRange());

    assert.equal(missingResult.ok, false);
    assert.equal(missingResult.needsInOut, true);
    assert.equal(setResult.ok, true);
    assert.equal(inPoint, 0);
    assert.equal(outPoint, 2);
    assert.equal(receivedOutTicks, sequence.end);
});

test('empty renamed managed audio tracks do not count as existing backup media', () => {
    const sequence = {
        name: 'Scene',
        audioTracks: makeCollection([
            makeTrack(0, [{ projectItem: { name: 'Dialogue.wav' } }]),
            makeTrack(0, [], 'Scene_Track1')
        ], 'numTracks'),
        videoTracks: makeCollection([], 'numTracks')
    };
    const { context } = loadHostLogic({
        project: {
            rootItem: { type: 2, children: makeCollection([], 'numItems') },
            sequences: makeCollection([sequence], 'numSequences'),
            save() {}
        }
    });
    context.sequenceUnderTest = sequence;

    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const requested = [{
            kind: 'audio',
            trackNumber: 1,
            trackNumbers: [1],
            path: 'D:\\\\Backups\\\\Scene_Track1.mp3',
            finalPath: 'D:\\\\Backups\\\\Scene_Track1.mp3'
        }];
        return {
            emptyConflicts: ebFindExistingProjectConflicts(sequenceUnderTest, requested, 'Scene'),
            managedSelection: ebGetSequenceManagedSelection(sequenceUnderTest, 'Scene')
        };
    })())`, context));

    assert.deepEqual(result.emptyConflicts, []);
    assert.equal(result.managedSelection.trackNumbers['2'], true);
});

test('video visibility is restored exactly after export-only hiding', () => {
    const { context, source } = loadHostLogic();
    const tracks = [
        makeTrack(1),
        makeTrack(0),
        makeTrack(1)
    ];
    const sequence = {
        videoTracks: makeCollection(tracks, 'numTracks')
    };
    context.sequenceUnderTest = sequence;

    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const states = ebCaptureVideoMuteStates(sequenceUnderTest);
        ebHideVideoTracksAbove(sequenceUnderTest, 1);
        const duringExport = [
            sequenceUnderTest.videoTracks[0].muteValue,
            sequenceUnderTest.videoTracks[1].muteValue,
            sequenceUnderTest.videoTracks[2].muteValue
        ];
        ebRestoreVideoMuteStates(sequenceUnderTest, states);
        return {
            duringExport,
            restored: [
                sequenceUnderTest.videoTracks[0].muteValue,
                sequenceUnderTest.videoTracks[1].muteValue,
                sequenceUnderTest.videoTracks[2].muteValue
            ]
        };
    })())`, context));

    assert.deepEqual(result.duringExport, [1, 1, 1]);
    assert.deepEqual(result.restored, [1, 0, 1]);
    assert.match(source, /finally\s*\{[\s\S]*ebRestoreVideoMuteStates\(sequence, originalVideoMuteStates\)/);
});

test('runBackupQueue restores video visibility when direct export throws', () => {
    const videoTracks = [makeTrack(0), makeTrack(0), makeTrack(1)];
    const sequence = {
        videoTracks: makeCollection(videoTracks, 'numTracks'),
        audioTracks: makeCollection([], 'numTracks')
    };
    const { context } = loadHostLogic({
        project: {
            activeSequence: sequence,
            rootItem: { type: 2, children: makeCollection([], 'numItems') },
            sequences: makeCollection([sequence], 'numSequences'),
            name: 'Test.prproj',
            path: 'D:\\Project\\Test.prproj',
            save() {}
        }
    });

    vm.runInContext(`
        ebEnsureFolder = function () { return true; };
        ebSequenceHasInOut = function () { return true; };
        ebGetSequenceExportBaseName = function () { return "Show"; };
        ebValidateBackupTrack = function () { return 1; };
        ebCheckPreset = function () {};
        ebBuildRequestedOutputFiles = function () {
            return [{
                kind: "video",
                path: "D:\\\\Backups\\\\Show_BACKUP.mp4",
                finalPath: "D:\\\\Backups\\\\Show_BACKUP.mp4"
            }];
        };
        ebFindExistingOutputConflicts = function () { return []; };
        ebFindExistingProjectConflicts = function () { return []; };
        ebRemoveUnusedMedia = function () { return false; };
        ebClearAllAudioSoloStates = function () { return 0; };
        ebApplyManagedTrackMutePolicy = function () {};
        ebGetSelectedExportItems = function () {
            return { includeVideo: true, audioTracks: [], audioGroups: [] };
        };
        ebExportSequenceDirect = function () {
            throw new Error("simulated export failure");
        };
    `, context);

    const result = JSON.parse(vm.runInContext(
        'exportBackup.runBackupQueue("D:\\\\Backups","video.epr","mp3.epr","wav.epr","wav",1,false,"{}","premiere",false)',
        context
    ));

    assert.equal(result.ok, false);
    assert.match(result.message, /simulated export failure/);
    assert.deepEqual(videoTracks.map((track) => track.muteValue), [0, 0, 1]);
});

test('backup MP4 audio can be made the only audible audio track', () => {
    const { context, source } = loadHostLogic();
    const tracks = [makeTrack(0), makeTrack(0), makeTrack(1), makeTrack(0)];
    context.sequenceUnderTest = {
        audioTracks: makeCollection(tracks, 'numTracks')
    };

    const muteValues = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        ebSetOnlyTrackAudible(sequenceUnderTest, 2);
        return [
            sequenceUnderTest.audioTracks[0].muteValue,
            sequenceUnderTest.audioTracks[1].muteValue,
            sequenceUnderTest.audioTracks[2].muteValue,
            sequenceUnderTest.audioTracks[3].muteValue
        ];
    })())`, context));

    assert.deepEqual(muteValues, [1, 1, 0, 1]);
    assert.match(source, /ebSetOnlyTrackAudible\(sequence, backupVideoAudioTrackNumber - 1\)/);
});

test('old backup clips are removed from every project sequence, including offline clips', () => {
    const targetPath = 'D:\\Backups\\Show_BACKUP.mp4';
    const onlineClip = {
        projectItem: {
            name: 'Show_BACKUP.mp4',
            getMediaPath() {
                return targetPath;
            }
        }
    };
    const offlineClip = {
        projectItem: {
            name: 'Show_BACKUP.mp4',
            getMediaPath() {
                return '';
            }
        }
    };
    const unrelatedClip = {
        projectItem: {
            name: 'Other.mp4',
            getMediaPath() {
                return 'D:\\Backups\\Other.mp4';
            }
        }
    };
    const firstClips = [onlineClip, unrelatedClip];
    const secondClips = [offlineClip];
    const firstTrack = makeTrack(0, firstClips);
    const secondTrack = makeTrack(0, secondClips);

    onlineClip.remove = () => firstTrack.clips.removeItem(onlineClip);
    unrelatedClip.remove = () => firstTrack.clips.removeItem(unrelatedClip);
    offlineClip.remove = () => secondTrack.clips.removeItem(offlineClip);

    const sequences = makeCollection([
        {
            videoTracks: makeCollection([firstTrack], 'numTracks'),
            audioTracks: makeCollection([], 'numTracks')
        },
        {
            videoTracks: makeCollection([], 'numTracks'),
            audioTracks: makeCollection([secondTrack], 'numTracks')
        }
    ], 'numSequences');
    const { context } = loadHostLogic({
        project: {
            rootItem: { type: 2, children: makeCollection([], 'numItems') },
            sequences,
            save() {}
        }
    });
    context.targetPaths = [targetPath];

    const removed = vm.runInContext('ebRemoveProjectClipsByMediaPaths(targetPaths)', context);

    assert.equal(removed, 2);
    assert.equal(firstTrack.clips.numItems, 1);
    assert.equal(firstTrack.clips[0], unrelatedClip);
    assert.equal(secondTrack.clips.numItems, 0);
});

test('media-path cleanup can target only video or only audio tracks', () => {
    const targetPath = 'D:\\Backups\\Show_BACKUP.mp4';
    const videoClip = {
        projectItem: {
            name: 'Show_BACKUP.mp4',
            getMediaPath() {
                return targetPath;
            }
        }
    };
    const audioClip = {
        projectItem: {
            name: 'Show_BACKUP.mp4',
            getMediaPath() {
                return targetPath;
            }
        }
    };
    const videoTrack = makeTrack(0, [videoClip]);
    const audioTrack = makeTrack(0, [audioClip]);

    videoClip.remove = () => videoTrack.clips.removeItem(videoClip);
    audioClip.remove = () => audioTrack.clips.removeItem(audioClip);

    const sequence = {
        videoTracks: makeCollection([videoTrack], 'numTracks'),
        audioTracks: makeCollection([audioTrack], 'numTracks')
    };
    const { context } = loadHostLogic({
        project: {
            rootItem: { type: 2, children: makeCollection([], 'numItems') },
            sequences: makeCollection([sequence], 'numSequences'),
            save() {}
        }
    });
    context.targetPaths = [targetPath];

    const removedVideoOnly = vm.runInContext('ebRemoveProjectClipsByMediaPaths(targetPaths, "video")', context);
    assert.equal(removedVideoOnly, 1);
    assert.equal(videoTrack.clips.numItems, 0);
    assert.equal(audioTrack.clips.numItems, 1);

    const removedAudioOnly = vm.runInContext('ebRemoveProjectClipsByMediaPaths(targetPaths, "audio")', context);
    assert.equal(removedAudioOnly, 1);
    assert.equal(audioTrack.clips.numItems, 0);
});
test('offline project items must actually leave the project tree before replacement succeeds', () => {
    const targetPath = 'D:\\Backups\\Show_BACKUP.mp4';
    const rootItems = [];
    const root = {
        type: 2,
        children: makeCollection(rootItems, 'numItems')
    };

    function createItem(nodeId, canDeleteAfterOffline) {
        let offline = false;
        const item = {
            nodeId,
            name: 'Show_BACKUP.mp4',
            type: 1,
            getMediaPath() {
                return offline ? '' : targetPath;
            },
            isOffline() {
                return offline;
            },
            setOffline() {
                offline = true;
                return true;
            },
            deleteBin() {
                return false;
            },
            remove() {
                if (offline && canDeleteAfterOffline) {
                    root.children.removeItem(item);
                    return true;
                }
                return false;
            }
        };
        return item;
    }

    const removableItem = createItem('removable', true);
    rootItems.push(removableItem);
    root.children = makeCollection(rootItems, 'numItems');
    const appProject = {
        rootItem: root,
        sequences: makeCollection([], 'numSequences'),
        save() {}
    };
    const { context } = loadHostLogic({ project: appProject });
    context.targetPath = targetPath;

    const removedResult = JSON.parse(vm.runInContext(
        'JSON.stringify(ebReleaseProjectItemsByMediaPath(targetPath))',
        context
    ));
    assert.equal(removedResult.found, 1);
    assert.equal(removedResult.offlined, 1);
    assert.equal(removedResult.removed, 1);
    assert.equal(removedResult.remaining, 0);

    const stubbornItem = createItem('stubborn', false);
    rootItems.push(stubbornItem);
    root.children = makeCollection(rootItems, 'numItems');
    const stubbornResult = JSON.parse(vm.runInContext(
        'JSON.stringify(ebReleaseProjectItemsByMediaPath(targetPath))',
        context
    ));
    assert.equal(stubbornResult.found, 1);
    assert.equal(stubbornResult.offlined, 1);
    assert.equal(stubbornResult.removed, 0);
    assert.equal(stubbornResult.remaining, 1);
    assert.equal(stubbornResult.remainingOnline, 0);
});

test('ordinary footage is deleted through a temporary bin using supported ProjectItem APIs', () => {
    const targetPath = 'D:\\Backups\\Show_BACKUP.mp4';
    const rootItems = [];
    const root = {
        type: 2,
        children: null
    };
    const item = {
        nodeId: 'footage-item',
        name: 'Show_BACKUP.mp4',
        type: 1,
        getMediaPath() {
            return targetPath;
        },
        isOffline() {
            return false;
        }
    };

    function syncRoot() {
        root.children = makeCollection(rootItems, 'numItems');
    }

    item.moveBin = (targetBin) => {
        const index = rootItems.indexOf(item);
        if (index >= 0) {
            rootItems.splice(index, 1);
        }
        targetBin._items.push(item);
        targetBin.children = makeCollection(targetBin._items, 'numItems');
        syncRoot();
        return 0;
    };

    root.createBin = (name) => {
        const bin = {
            nodeId: `bin-${name}`,
            name,
            type: 2,
            _items: [],
            children: makeCollection([], 'numItems'),
            deleteBin() {
                const index = rootItems.indexOf(bin);
                if (index >= 0) {
                    rootItems.splice(index, 1);
                }
                syncRoot();
                return 0;
            }
        };
        rootItems.push(bin);
        syncRoot();
        return bin;
    };

    rootItems.push(item);
    syncRoot();
    const { context } = loadHostLogic({
        project: {
            rootItem: root,
            sequences: makeCollection([], 'numSequences'),
            save() {}
        }
    });
    context.targetPath = targetPath;

    const result = JSON.parse(vm.runInContext(
        'JSON.stringify(ebReleaseProjectItemsByMediaPath(targetPath))',
        context
    ));

    assert.equal(result.found, 1);
    assert.equal(result.movedToCleanupBin, 1);
    assert.equal(result.cleanupBinDeleted, true);
    assert.equal(result.removed, 1);
    assert.equal(result.remaining, 0);
    assert.equal(root.children.numItems, 0);
});

test('project-item cleanup finds media that was already offline before cleanup began', () => {
    const targetPath = 'D:\\Backups\\Show_BACKUP.mp4';
    const rootItems = [];
    const root = {
        type: 2,
        children: makeCollection(rootItems, 'numItems')
    };
    const offlineItem = {
        nodeId: 'already-offline',
        name: 'Show_BACKUP.mp4',
        type: 1,
        getMediaPath() {
            return '';
        },
        isOffline() {
            return true;
        },
        setOffline() {
            return true;
        },
        deleteBin() {
            return false;
        },
        remove() {
            root.children.removeItem(offlineItem);
            return true;
        }
    };
    rootItems.push(offlineItem);
    root.children = makeCollection(rootItems, 'numItems');

    const { context } = loadHostLogic({
        project: {
            rootItem: root,
            sequences: makeCollection([], 'numSequences'),
            save() {}
        }
    });
    context.targetPath = targetPath;

    const result = JSON.parse(vm.runInContext(
        'JSON.stringify(ebReleaseProjectItemsByMediaPath(targetPath))',
        context
    ));

    assert.equal(result.found, 1);
    assert.equal(result.removed, 1);
    assert.equal(result.remaining, 0);
});

test('re-backup cleanup releases both TEMP imports and old final-path media before rename', () => {
    const { context } = loadHostLogic();
    context.expectedFiles = [
        {
            path: 'D:\\Backups\\Show_BACKUP_REBKP_TEMP.mp4',
            finalPath: 'D:\\Backups\\Show_BACKUP.mp4'
        },
        {
            path: 'D:\\Backups\\Show_Track1_REBKP_TEMP.wav',
            finalPath: 'D:\\Backups\\Show_Track1.wav'
        },
        {
            path: 'D:\\Backups\\Show_BACKUP_REBKP_TEMP.mp4',
            finalPath: 'D:\\Backups\\Show_BACKUP.mp4'
        }
    ];

    const paths = JSON.parse(vm.runInContext(
        'JSON.stringify(ebGetRebackupReleasePaths(expectedFiles))',
        context
    ));

    assert.deepEqual(paths, [
        'D:\\Backups\\Show_BACKUP_REBKP_TEMP.mp4',
        'D:\\Backups\\Show_BACKUP.mp4',
        'D:\\Backups\\Show_Track1_REBKP_TEMP.wav',
        'D:\\Backups\\Show_Track1.wav'
    ]);
});

test('selected re-backup video stays visible in the sequence until MP4 export finishes', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');
    const runStart = hostSource.indexOf('exportBackup.runBackupQueue = function');
    const runEnd = hostSource.indexOf('function ebGetRebackupReleasePaths', runStart);
    const runSource = hostSource.slice(runStart, runEnd);
    const exportPosition = runSource.indexOf('ebExportSequenceDirect(sequence, videoPath, videoPresetPath, workAreaType)');
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const panelRunStart = mainSource.indexOf('async function runExport');
    const hostResultPosition = mainSource.indexOf('const result = await callHost(script)', panelRunStart);
    const cleanupPosition = mainSource.indexOf('await prepareRebackupReplacement(manifest)', hostResultPosition);

    assert.doesNotMatch(hostSource, /function ebReleaseRebackupVideoBeforeExport/);
    assert.doesNotMatch(runSource, /ebRemoveProjectClipsByMediaPaths|ebReleaseProjectItemsByMediaPath/);
    assert.doesNotMatch(runSource, /existingBackupVideoTrack\.setMute\(1\)/);
    assert.ok(exportPosition >= 0);
    assert.ok(cleanupPosition > hostResultPosition);
});
test('Media Encoder queue failures keep Re-backup recovery metadata', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');
    const runStart = hostSource.indexOf('exportBackup.runBackupQueue = function');
    const runEnd = hostSource.indexOf('function ebGetRebackupReleasePaths', runStart);
    const runSource = hostSource.slice(runStart, runEnd);

    assert.match(runSource, /throw new Error\("Could not queue the MP4 export in Adobe Media Encoder\."\)/);
    assert.match(runSource, /if \(shouldRebackup\) \{\s*throw new Error\("Could not queue " \+ audioLabel/);
    assert.match(runSource, /if \(rebackupPreservationPrepared && requestedFiles && requestedFiles\.length\)/);
    assert.match(runSource, /recoveryQueuedFiles\.push\(ebBuildQueuedFileFromRequested\(requestedFiles\[recoveryIndex\]\)\)/);
});
test('preserved old files are cleaned only after completed final-name exports', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');
    const runStart = hostSource.indexOf('exportBackup.runBackupQueue = function');
    const runEnd = hostSource.indexOf('function ebGetRebackupReleasePaths', runStart);
    const runSource = hostSource.slice(runStart, runEnd);
    const preparePosition = runSource.indexOf('ebPrepareRebackupMediaForDirectExport(requestedFiles)');
    const directPathPosition = hostSource.indexOf('requested.path = ebToFsPath(requested.finalPath)');
    const exportPosition = runSource.indexOf('ebExportSequenceDirect(sequence, videoPath, videoPresetPath, workAreaType)');

    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const directCleanupStart = mainSource.indexOf('if (manifest.rebackup && parsed.exportMode === EXPORT_MODE_PREMIERE)');
    const readyPosition = mainSource.indexOf('assertExpectedExportsAreReady(manifest)', directCleanupStart);
    const premiereCleanupPosition = mainSource.indexOf('await prepareRebackupReplacement(manifest)', directCleanupStart);
    const finalizeStart = mainSource.indexOf('async function finalizeRebackupFiles');
    const finalizeEnd = mainSource.indexOf('function replacePathExtension', finalizeStart);
    const finalizeSource = mainSource.slice(finalizeStart, finalizeEnd);

    assert.ok(preparePosition >= 0);
    assert.ok(directPathPosition >= 0);
    assert.ok(exportPosition > preparePosition);
    assert.ok(readyPosition > directCleanupStart);
    assert.ok(premiereCleanupPosition > readyPosition);
    assert.match(finalizeSource, /getManifestPreservedPaths/);
    assert.ok(finalizeSource.includes('removeFileIfExists(preservedPaths[i], "preserved old backup file")'));
    assert.ok(mainSource.includes("lowerName.includes(REBACKUP_OLD_MARKER.toLowerCase())"));
    assert.match(mainSource, /New Re-backup export files are not complete yet/);
});
test('re-backup cleanup removes only selected re-exported timeline clips', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');
    const cleanupStart = hostSource.indexOf('exportBackup.prepareRebackupReplacement = function');
    const cleanupEnd = hostSource.indexOf('function ebGetAudioEntryTrackNumbers', cleanupStart);
    const cleanupSource = hostSource.slice(cleanupStart, cleanupEnd);

    assert.match(cleanupSource, /ebRemoveProjectClipsByMediaPaths\(\s*ebGetRebackupReleasePathsByKind\(expectedFiles, "video"\),\s*"video"\s*\)/);
    assert.match(cleanupSource, /ebRemoveProjectClipsByMediaPaths\(\s*ebGetRebackupReleasePathsByKind\(expectedFiles, "audio"\),\s*"audio"\s*\)/);
    assert.doesNotMatch(cleanupSource, /ebRemoveManagedClipsFromTrack/);
    assert.doesNotMatch(cleanupSource, /ebRemoveAllManagedClipsFromAllAudioTracks/);
    assert.doesNotMatch(cleanupSource, /ebFindManagedBackupVideoTrackNumber/);
});

test('automatic alignment uses only files selected for the current export', () => {
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const scannerStart = mainSource.indexOf('function scanExportFolderForSequence');
    const scannerEnd = mainSource.indexOf('function writeExportManifest', scannerStart);
    const scannerSource = mainSource.slice(scannerStart, scannerEnd);
    const recoveryStart = mainSource.indexOf('function collectRebackupRecoveryEntries');
    const recoveryEnd = mainSource.indexOf('async function ensureRebackupTempFilesAreStable', recoveryStart);
    const recoverySource = mainSource.slice(recoveryStart, recoveryEnd);
    const automaticAlignmentCount = (mainSource.match(/manifestOnly: true/g) || []).length;

    assert.match(scannerSource, /const manifestOnly = settings\.manifestOnly === true && !!manifest/);
    assert.match(scannerSource, /if \(!manifestOnly\) \{\s*files\.forEach/);
    assert.match(recoverySource, /const manifestOnly = settings\.manifestOnly === true && !!manifest/);
    assert.match(recoverySource, /if \(!manifestOnly\) \{\s*fs\.readdirSync/);
    assert.equal(automaticAlignmentCount, 2);
    assert.match(
        mainSource,
        /async function alignExistingFolder\(\)[\s\S]*?runAlignmentFlow\(exportFolder \|\| alignFolder, \{\s*skipVideo: false,\s*autoTriggered: false/
    );
});

test('filesystem rename is guarded by successful Premiere cleanup', () => {
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');

    assert.match(
        mainSource,
        /if \(!manifest \|\| manifest\.rebackupReplacementPrepared !== true\)\s*\{\s*throw new Error\("Re-backup cleanup was not verified/
    );
    assert.match(
        mainSource,
        /manifest\.rebackupReplacementPrepared = true;[\s\S]*return parsed;/
    );
    assert.match(
        mainSource,
        /await prepareRebackupReplacement\([^;]+;\s*await finalizeRebackupFiles\(/
    );
});

test('Premiere cleanup saves and settles before local re-backup replacement', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');
    const cleanupStart = hostSource.indexOf('exportBackup.prepareRebackupReplacement = function');
    const settlePosition = hostSource.indexOf('var projectSavedForRelease = ebSettleMediaReleaseAfterCleanup()', cleanupStart);
    const verifyPosition = hostSource.indexOf('if (ebGetOnlineProjectItemCountByMediaPath', cleanupStart);

    assert.match(hostSource, /function ebSaveProjectForMediaRelease\(\)[\s\S]*app\.project\.save\(\)/);
    assert.match(hostSource, /function ebSettleMediaReleaseAfterCleanup\(\)[\s\S]*EB_MEDIA_RELEASE_SAVE_WAIT_MS[\s\S]*ebRemoveUnusedMedia\(\)[\s\S]*EB_MEDIA_RELEASE_WAIT_MS/);
    assert.ok(settlePosition > cleanupStart);
    assert.ok(verifyPosition > settlePosition);
    assert.match(hostSource, /projectSavedForRelease: projectSavedForRelease/);
});

test('imported backup video is orange while backup audio remains brown', () => {
    const appliedLabels = [];
    const { context } = loadHostLogic();
    context.mediaItemUnderTest = {
        setColorLabel(labelIndex) {
            appliedLabels.push(labelIndex);
            return 0;
        }
    };

    const results = JSON.parse(vm.runInContext(
        'JSON.stringify([' +
            'ebSetProjectItemColorLabel(mediaItemUnderTest, EB_BACKUP_VIDEO_ORANGE_LABEL_INDEX),' +
            'ebSetProjectItemColorLabel(mediaItemUnderTest, EB_BACKUP_MEDIA_BROWN_LABEL_INDEX)' +
        '])',
        context
    ));
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');
    const alignStart = hostSource.indexOf('exportBackup.alignMappedFiles = function');
    const alignEnd = hostSource.indexOf('return ebResult(true', alignStart);
    const alignSource = hostSource.slice(alignStart, alignEnd);

    assert.deepEqual(results, [true, true]);
    assert.deepEqual(appliedLabels, [7, 14]);
    assert.match(alignSource, /ebSetProjectItemColorLabel\(videoItem, EB_BACKUP_VIDEO_ORANGE_LABEL_INDEX\)/);
    assert.match(alignSource, /ebSetProjectItemColorLabel\(audioItem, EB_BACKUP_MEDIA_BROWN_LABEL_INDEX\)/);
});

test('Premiere project is saved only once after import and alignment', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');
    const alignStart = hostSource.indexOf('exportBackup.alignMappedFiles = function');
    const savePosition = hostSource.indexOf('app.project.save()', alignStart);
    const importPosition = hostSource.indexOf('var videoItem = ebImportProjectItem', alignStart);
    const audioPolicyPosition = hostSource.indexOf(
        'ebSetOnlyTrackAudible(sequence, backupVideoAudioTrackNumber - 1)',
        alignStart
    );

    assert.ok(savePosition > importPosition);
    assert.ok(savePosition > audioPolicyPosition);
});

test('completion and import recovery messages use compact status text and visible dialogs', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const alignmentStart = mainSource.indexOf('async function runAlignmentFlow');
    const alignmentEnd = mainSource.indexOf('function scheduleExportMonitorTick', alignmentStart);
    const alignmentSource = mainSource.slice(alignmentStart, alignmentEnd);

    assert.match(html, /\.status-box\.is-success[\s\S]*font-size: 14px/);
    assert.match(html, /\.status-box\.is-error[\s\S]*font-size: 17px/);
    assert.match(mainSource, /Backup done without error\./);
    assert.match(mainSource, /Re-backup done without error\./);
    assert.match(mainSource, /return "ALIGNMENT DONE"/);
    assert.match(mainSource, /Please use Align Existing to import the file\./);
    assert.match(mainSource, /setStatus\(lines\.join\("\\n"\), "error"\)/);
    assert.match(mainSource, /setStatus\(lines\.join\("\\n"\), "success"\)/);
    assert.match(mainSource, /function showResultPrompt\([\s\S]*showBlockingMessage\(lines\.join\("\\n\\n"\)\)/);
    assert.match(mainSource, /showResultPrompt\("IMPORT NOT COMPLETED", ALIGNMENT_RECOVERY_TEXT\)/);
    assert.match(alignmentSource, /showResultPrompt\(\s*successTitle,/);
    assert.doesNotMatch(mainSource, /function showDesktopResultWindow\(/);
});

test('normal backup reports existing backup media before empty-track errors', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');
    const existingMessageIndex = hostSource.indexOf('Backup files are already there. Use Re-backup.');
    const emptyTrackIndex = hostSource.indexOf('ebValidateBackupTrack(sequence, backupVideoTrackNumber, true, sequenceBaseName);', existingMessageIndex);

    assert.notEqual(existingMessageIndex, -1);
    assert.notEqual(emptyTrackIndex, -1);
    assert.ok(existingMessageIndex < emptyTrackIndex);
});

test('automatic empty backup track option is visible and unchecked by default', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');

    assert.match(html, /id="autoEmptyBackupTrackCheckbox"/);
    assert.doesNotMatch(html, /id="autoEmptyBackupTrackCheckbox" checked/);
    assert.match(html, />TO EMPTY TRACK</);
    assert.match(html, /BACKUP TO[\s\S]*id="decrementBackupTrackButton"[\s\S]*id="exportVideoTrackInput"[\s\S]*id="incrementBackupTrackButton"[\s\S]*class="track-choice-separator">or<\/span>[\s\S]*TO EMPTY TRACK/);
    assert.match(html, /class="action-line folder-action-line"[\s\S]*id="chooseFolderButton"[\s\S]*id="exportPath"[\s\S]*id="togglePresetSectionButton"[\s\S]*Change Export Presets[\s\S]*id="presetSection"/);
    assert.equal((html.match(/id="togglePresetSectionButton"/g) || []).length, 1);
    assert.match(html, /id="refreshExportSelectionButton"[\s\S]*id="updateButton"[\s\S]*id="toggleQueueBackupSectionButton"/);
    assert.match(mainSource, /function resetAutoEmptyBackupTrackOption\(\)/);
    assert.match(mainSource, /function bindAutoEmptyBackupTrackOption\(\)/);
    assert.match(mainSource, /function bindBackupTrackStepper\(\)/);
    assert.match(mainSource, /backupTrackInput\.disabled = disabled/);
    assert.match(mainSource, /button\.disabled = disabled/);
    assert.match(mainSource, /checkbox\.checked = false/);
    assert.doesNotMatch(mainSource, /autoEmptyBackupTrack.*localStorage|localStorage.*autoEmptyBackupTrack/);
    assert.match(mainSource, /validateBackupExportSettings\(backupVideoTrackNumber, selectedQueueItems, isRebackup, autoEmptyTrack\)/);
});



test('missing In and Out prompt can auto-set the full range or cancel', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');

    assert.match(html, /id="inOutPrompt"/);
    assert.match(html, /id="autoSetInOutCheckbox" checked/);
    assert.match(html, /id="inOutPromptCancelButton"[^>]*>Cancel</);
    assert.match(html, /id="inOutPromptOkButton"[^>]*>OK</);
    assert.match(mainSource, /checkbox\.checked = true/);
    assert.match(mainSource, /okButton\.disabled = !checkbox\.checked/);
    assert.match(mainSource, /if \(!validation\.ok && validation\.needsInOut === true\)/);
    assert.match(mainSource, /const shouldAutoSetInOut = await showInOutPrompt\(\)/);
    assert.match(mainSource, /const inOutResult = await setActiveSequenceInOutToFullRange\(\)/);
    assert.match(mainSource, /validation = await validateBackupExportSettings\(/);
    assert.match(mainSource, /Export cancelled\. Set sequence In and Out manually/);
    assert.match(hostSource, /exportBackup\.setActiveSequenceInOutToFullRange = function/);
    assert.match(hostSource, /sequence\.setInPoint\(0\)/);
    assert.match(hostSource, /sequence\.setOutPoint\(endTime\)/);
});

test('Align Existing prefers selected audio format and cleans opposite-format recovery files', () => {
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');

    assert.match(mainSource, /const preferredAudioFormat = getSelectedAudioFormat\(\)/);
    assert.match(mainSource, /normalizeAudioEntries\(audio, sanitizedBase, preferredAudioFormat\)/);
    assert.match(mainSource, /getAudioEntryFormat\(normalizedEntry\.path\) === preferredFormat/);
    assert.match(mainSource, /oldFinalPath: getOppositeAudioFormatPath\(finalPath\)/);
    assert.match(mainSource, /await removeFileIfExists\(entry\.oldFinalPath, "old-format backup file"\)/);
});

test('re-backup replacement retries direct delete before local rename', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');

    assert.doesNotMatch(html, /id="progressDialog"/);
    assert.match(mainSource, /async function runLocalFileStepWithRetry/);
    assert.match(mainSource, /for \(let attempt = 1; attempt <= 2; attempt \+= 1\)/);
    assert.match(mainSource, /const retryDelays = \[10000\]/);
    assert.match(mainSource, /Retrying \$\{title\} in \$\{waitSeconds\} seconds/);
    assert.match(mainSource, /Retry \$\{attempt - 1\}\/1/);
    assert.match(mainSource, /File is still busy\. Retrying delete in \$\{waitSeconds\} seconds/);
    assert.match(mainSource, /Deleting \$\{label\}/);
    assert.match(mainSource, /fs\.rmSync\(filePath, \{ force: true, maxRetries: 1, retryDelay: 100 \}\)/);
    assert.match(mainSource, /deleteLocalFileWithWindowsShell\(filePath\)/);
    assert.match(mainSource, /deleteLocalFileWithCmd\(filePath\)/);
    assert.match(mainSource, /function deleteLocalFileWithExplorer\(filePath\)/);
    assert.match(mainSource, /await tryExplorerDelete\(filePath, label\)/);
    assert.match(mainSource, /Old file deleted with Windows Explorer/);
    assert.match(mainSource, /function buildDeletePendingPath\(filePath\)/);
    assert.match(mainSource, /_DELETE_PENDING_/);
    assert.match(mainSource, /fs\.renameSync\(filePath, pendingPath\)/);
    assert.match(mainSource, /await moveFileAsideIfNeeded\(filePath, label\)/);
    assert.match(mainSource, /Old file moved aside/);
    assert.doesNotMatch(mainSource, /confirm\(/);
    assert.doesNotMatch(mainSource, /_cleanup/);
    assert.doesNotMatch(mainSource, /showLockedFileRecoveryMessage/);
    assert.doesNotMatch(mainSource, /deleteLocalFileWithElevatedShell/);
    assert.match(mainSource, /File still busy\\n\\nPress Align Existing button to delete and import re-export\./);
});

test('Align Existing removes sequence-managed backup clips before target-track emptiness check', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');

    assert.match(hostSource, /shouldRemove = ebIsManagedBackupTrack\(clipName, baseName\) \|\| ebIsSequenceManagedBackupTrack\(clipName, baseName\)/);
    assert.match(
        hostSource,
        /ebRemoveManagedClipsFromTrack\(sequence\.videoTracks\[resolvedBackupTrack - 1\], sequenceBaseName, "backup", 0\);[\s\S]*if \(videoPath && ebTrackHasClips\(sequence\.videoTracks\[resolvedBackupTrack - 1\]\)\)/
    );
});

test('automatic backup track selection picks the lowest empty video track', () => {
    const { context } = loadHostLogic();
    const sequence = {
        videoTracks: makeCollection([
            makeTrack(0, [{ projectItem: { name: 'V1.mov' } }]),
            makeTrack(0, [{ projectItem: { name: 'V2.mov' } }]),
            makeTrack(0, [{ projectItem: { name: 'V3.mov' } }]),
            makeTrack(0, [{ projectItem: { name: 'V4.mov' } }]),
            makeTrack(0, [{ projectItem: { name: 'V5.mov' } }]),
            makeTrack(0),
            makeTrack(0, [{ projectItem: { name: 'V7.mov' } }]),
            makeTrack(0)
        ], 'numTracks')
    };
    context.sequenceUnderTest = sequence;

    const result = vm.runInContext('ebResolveBackupVideoTrackNumber(sequenceUnderTest, 5, true, false)', context);

    assert.equal(result, 6);
});

test('automatic backup track selection creates a new top video track when none are empty', () => {
    const { context } = loadHostLogic();
    const tracks = [
        makeTrack(0, [{ projectItem: { name: 'V1.mov' } }]),
        makeTrack(0, [{ projectItem: { name: 'V2.mov' } }])
    ];
    const sequence = {
        videoTracks: makeCollection(tracks, 'numTracks')
    };
    context.sequenceUnderTest = sequence;
    context.qe = {
        project: {
            getActiveSequence() {
                return {
                    addTracks(videoTracksToAdd) {
                        for (let i = 0; i < videoTracksToAdd; i += 1) {
                            tracks.push(makeTrack(0));
                            sequence.videoTracks[sequence.videoTracks.numTracks] = tracks[tracks.length - 1];
                            sequence.videoTracks.numTracks = tracks.length;
                        }
                    }
                };
            }
        }
    };

    const result = vm.runInContext('ebResolveBackupVideoTrackNumber(sequenceUnderTest, 5, true, true)', context);

    assert.equal(result, 3);
    assert.equal(sequence.videoTracks.numTracks, 3);
    assert.equal(sequence.videoTracks[2].clips.numItems, 0);
});
