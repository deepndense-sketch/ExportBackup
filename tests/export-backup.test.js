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

function makeTrack(initialMute, clips) {
    let muted = initialMute ? 1 : 0;
    return {
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

function loadHostLogic(appOverrides) {
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
    const context = vm.createContext({
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
    });

    vm.runInContext(source, context, { filename: sourcePath });
    return { context, source };
}

test('Queue Backup Exports has a visual-only toggle and always starts shown', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');

    assert.match(html, /id="toggleQueueBackupSectionButton"[^>]*>Hide</);
    assert.match(html, /id="queueBackupSectionContent" class="stack"/);
    assert.doesNotMatch(html, /id="queueBackupSectionContent"[^>]*is-hidden/);
    assert.match(mainSource, /function toggleQueueBackupSection\(\)/);
    assert.match(
        mainSource,
        /document\.addEventListener\("DOMContentLoaded",[\s\S]*setQueueBackupSectionVisibility\(true\)/
    );
    assert.doesNotMatch(mainSource, /queueBackupSectionVisibleStorage/i);
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

test('Premiere project is saved only once after import and alignment', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');
    const saveCalls = hostSource.match(/app\.project\.save\(\)/g) || [];
    const alignStart = hostSource.indexOf('exportBackup.alignMappedFiles = function');
    const savePosition = hostSource.indexOf('app.project.save()', alignStart);
    const importPosition = hostSource.indexOf('var videoItem = ebImportProjectItem', alignStart);
    const audioPolicyPosition = hostSource.indexOf(
        'ebSetOnlyTrackAudible(sequence, backupVideoAudioTrackNumber - 1)',
        alignStart
    );

    assert.equal(saveCalls.length, 1);
    assert.ok(savePosition > importPosition);
    assert.ok(savePosition > audioPolicyPosition);
});
