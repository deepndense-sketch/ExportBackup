const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function createTestElement() {
    const attributes = new Map();
    const classes = new Set();
    return {
        children: [],
        classList: {
            add(name) {
                classes.add(name);
            },
            contains(name) {
                return classes.has(name);
            },
            remove(name) {
                classes.delete(name);
            },
            toggle(name, force) {
                if (force === true) {
                    classes.add(name);
                } else if (force === false) {
                    classes.delete(name);
                } else if (classes.has(name)) {
                    classes.delete(name);
                } else {
                    classes.add(name);
                }
            }
        },
        disabled: false,
        innerHTML: '',
        onclick: null,
        parentNode: null,
        style: {},
        textContent: '',
        appendChild(child) {
            child.parentNode = this;
            this.children.push(child);
        },
        getAttribute(name) {
            return attributes.has(name) ? attributes.get(name) : null;
        },
        setAttribute(name, value) {
            attributes.set(name, String(value));
        }
    };
}

function createTestDocument() {
    const elementIds = [
        'chooseButton',
        'path',
        'categoryDestination',
        'projectRootDestination',
        'manualDestination',
        'collectButton',
        'compareButton',
        'completionMessage',
        'completionPrompt',
        'completionTitle',
        'currentFile',
        'errorList',
        'missingList',
        'linkExistingBackupButton',
        'progressFill',
        'progressText',
        'refreshProjectButton',
        'sequenceFilterHint',
        'sequenceFilters',
        'selectionSummary',
        'summaryText',
        'trackConflictContinueButton',
        'trackConflictList',
        'trackConflictPrompt',
        'trackConflictStatus',
        'updateButton'
    ];
    const elements = new Map(elementIds.map((id) => [id, createTestElement()]));

    return {
        addEventListener() {},
        createElement() {
            return createTestElement();
        },
        getElementById(id) {
            return elements.get(id) || null;
        },
        querySelector() {
            return null;
        }
    };
}

function createMemoryStorage(initialValues) {
    const values = new Map(Object.entries(initialValues || {}));
    return {
        get length() {
            return values.size;
        },
        getItem(key) {
            return values.has(key) ? values.get(key) : null;
        },
        key(index) {
            return Array.from(values.keys())[index] || null;
        },
        removeItem(key) {
            values.delete(key);
        },
        setItem(key, value) {
            values.set(key, String(value));
        }
    };
}

function loadCollectorLogic(testDocument, options) {
    const settings = options || {};
    const scriptPath = path.join(__dirname, '..', 'js', 'main.js');
    const source = fs.readFileSync(scriptPath, 'utf8');
    const context = vm.createContext({
        Buffer,
        Map,
        Set,
        URL,
        clearInterval,
        clearTimeout,
        console,
        process,
        require,
        setInterval,
        setTimeout,
        CSInterface: function CSInterface() {},
        SystemPath: { EXTENSION: 'extension' },
        document: testDocument || {
            addEventListener() {},
            getElementById() {
                return null;
            },
            querySelector() {
                return null;
            }
        },
        localStorage: settings.localStorage || createMemoryStorage(),
        window: {
            parent: settings.parent,
            addEventListener() {},
            confirm: settings.confirm || (() => true),
            prompt: settings.prompt || (() => null)
        },
        alert: settings.alert || (() => {})
    });

    vm.runInContext(source, context, { filename: scriptPath });
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'js', 'backup-routing.js'), 'utf8'), context);
    return context;
}

test('automatic backup preset uses independent lowest video/audio backup boundaries', () => {
    const context = loadCollectorLogic();
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const video = [1,2,3,4,5,6].map(trackNumber => ({trackNumber, isBackup: trackNumber === 4}));
        const audio = [1,2,3,4,5,6,7,8].map(trackNumber => ({trackNumber, isBackup: trackNumber === 3 || trackNumber === 6}));
        const filter = createSequenceFilter('id', 'Show', video, audio, false);
        applyTrackPresetToFilter(filter, AUTO_BACKUP_PRESET);
        const initial = {video: filter.ignoredVideoTracks, audio: filter.ignoredAudioTracks, preset: filter.selectedPresetId};
        mergeSequenceTrackUsage(filter, createSequenceFilter('id', 'Show', video.map(t => ({...t, isBackup: t.trackNumber === 5})), [], false));
        return {initial, refreshed: filter.ignoredVideoTracks, noAudio: filter.ignoredAudioTracks,
            undeletable: deleteTrackPreset(AUTO_BACKUP_PRESET.id)};
    })())`, context));
    assert.deepEqual(result.initial.video, [4,5,6]);
    assert.deepEqual(result.initial.audio, [3,4,5,6,7,8]);
    assert.equal(result.initial.preset, 'builtin-backup-boundary');
    assert.deepEqual(result.refreshed, [5,6]);
    assert.deepEqual(result.noAudio, []);
    assert.equal(result.undeletable, false);
});

test('new sequences without backups keep all tracks; custom tracks survive refresh', () => {
    const context = loadCollectorLogic();
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const filter = createSequenceFilter('id', 'Show', [{trackNumber:1}], [], false);
        const before = filter.ignoredVideoTracks.slice();
        filter.selectedPresetId = '';
        filter.ignoredVideoTracks = [1];
        mergeSequenceTrackUsage(filter, createSequenceFilter('id', 'Show', [{trackNumber:1}, {trackNumber:2,isBackup:true}], [], false));
        return {before, after: filter.ignoredVideoTracks};
    })())`, context));
    assert.deepEqual(result, {before: [], after: [1]});
});

test('show folder routing preserves full names, including MAIN and INTRO', () => {
    const context = loadCollectorLogic();
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        configuredBackupCategories = ['WOW'];
        return [resolveCollectorCategoryPath('D:/WOW 3250 MAIN.prproj'),
            resolveCollectorCategoryPath('D:/3250 WOW INTRO.prproj'),
            resolveCollectorCategoryPath('D:/3250 WOW Title.prproj')];
    })())`, context));
    assert.notEqual(result[0], result[1]);
    assert.equal(path.basename(result[0]), 'WOW 3250 MAIN');
    assert.equal(path.basename(result[1]), 'WOW 3250 INTRO');
    assert.equal(path.basename(result[2]), 'WOW 3250 Title');
    assert.throws(() => context.resolveCollectorCategoryPath(''), /Save the Premiere project/);
    assert.throws(() => context.resolveCollectorCategoryPath('Unknown.prproj'), /built-in show categories/);
});

test('closing Collector clears sequence selections but keeps destination settings', () => {
    const storage = createMemoryStorage();
    storage.setItem('projectcollector.sequenceFilters:C:/Show.prproj','[{"sequenceID":"old"}]');
    storage.setItem('projectcollector.destinationMode','manual');
    const context = loadCollectorLogic(createTestDocument(), {localStorage:storage});
    const result = vm.runInContext(`(() => {
        clearCollectorSequenceSession();
        latestPlan = {projectPath:'C:/Show.prproj',activeSequenceID:'seq',activeSequenceName:'Show',videoTrackUsage:[],audioTrackUsage:[]};
        renderSequenceFilters = () => {};
        loadSequenceFilters();
        return selectedSequenceFilters.length;
    })()`,context);
    assert.equal(result,0);
    assert.equal(storage.getItem('projectcollector.sequenceFilters:C:/Show.prproj'),null);
    assert.equal(storage.getItem('projectcollector.destinationMode'),'manual');
});

test('adding a sequence reloads the project and automatically refreshes track locks', async () => {
    const context = loadCollectorLogic(createTestDocument());
    const result = JSON.parse(await vm.runInContext(`(async () => {
        const calls = [];
        ensureHostScriptLoaded = async () => true;
        loadProjectPlan = async () => {calls.push('project'); return true;};
        readCurrentActiveSequenceFilter = async () => {calls.push('sequence'); return createSequenceFilter('seq','Show',[],[],false);};
        renderSequenceFilters = () => {};
        refreshAllSelectedSequenceTracks = async () => {calls.push('tracks');};
        await addCurrentActiveSequence();
        return JSON.stringify({calls,count:selectedSequenceFilters.length});
    })()`,context));
    assert.deepEqual(result.calls,['project','sequence','tracks']);
    assert.equal(result.count,1);
});

test('project-root destination uses the saved project directory without category validation or extra nesting', () => {
    const context = loadCollectorLogic();
    const result = vm.runInContext(`(() => {
        projectRootDestination = true;
        destination = 'D:/Unrelated';
        return resolveCollectorDestination({projectPath:'C:/Projects/My Edit/Edit.prproj', projectName:'Edit'});
    })()`, context);
    assert.equal(result, path.dirname('C:/Projects/My Edit/Edit.prproj'));
    assert.throws(() => vm.runInContext('resolveCollectorDestination({projectName:"Unsaved"})', context), /Save the Premiere project/);
    const manual = vm.runInContext(`(() => {
        projectRootDestination = false;
        categoryDestination = false;
        return resolveCollectorDestination({projectName:'Edit'});
    })()`, context);
    assert.equal(manual, path.join('D:/Unrelated', 'Edit'));
});

test('changing destination persists project-root and manual modes', async () => {
    const doc = createTestDocument();
    const storage = createMemoryStorage();
    const context = loadCollectorLogic(doc, {localStorage:storage});
    doc.getElementById('categoryDestination').checked = false;
    doc.getElementById('projectRootDestination').checked = true;
    vm.runInContext(`latestPlan = {projectPath:'C:/Edits/Example.prproj'}; loadProjectPlan = async () => true;`, context);
    await context.updateCollectorDestination();
    assert.equal(storage.getItem('projectcollector.destinationMode'), 'projectRoot');
    assert.equal(doc.getElementById('path').textContent, path.dirname('C:/Edits/Example.prproj'));
    assert.equal(doc.getElementById('chooseButton').disabled, true);
    doc.getElementById('projectRootDestination').checked = false;
    await context.updateCollectorDestination();
    assert.equal(storage.getItem('projectcollector.destinationMode'), 'manual');
    assert.equal(doc.getElementById('chooseButton').disabled, false);
});

test('track locks round-trip from Premiere without changing the timeline', () => {
    const context = vm.createContext({});
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'jsx', 'collector.jsx'), 'utf8'), context);
    const tracks = {numTracks: 5,
        0: {isLocked: () => true}, 1: {isLocked: () => 0},
        2: {isLocked: true}, 3: {}, 4: {isLocked: () => {throw new Error('unavailable');}}
    };
    const entries = JSON.parse(context.pcTrackUsageJson(context.pcTrackCollectionUsage(tracks, 'V', null)));
    assert.deepEqual(entries.map(entry => entry.isLocked), [true, false, true, null, null]);
});

test('track preview expands nested media and matches recursive copy scope', () => {
    const context = loadHostRelinkLogic();
    const collection = (items, countKey) => Object.assign({[countKey]:items.length}, items);
    const media = (name) => ({isSequence:() => false,getMediaPath:() => 'D:/Media/' + name});
    const nestedItem = (id) => ({nodeId:id,isSequence:() => true});
    const track = (items) => ({isLocked:() => false,clips:collection(items.map(projectItem => ({projectItem})), 'numItems')});
    const innerItem = nestedItem('inner-item');
    const outerItem = nestedItem('outer-item');
    const inner = {sequenceID:'inner',name:'Inner',projectItem:innerItem,
        videoTracks:collection([track([media('one.png'),media('two.png'),media('one.png'),outerItem])],'numTracks'),
        audioTracks:collection([track([media('music.wav')])],'numTracks')};
    const outer = {sequenceID:'outer',name:'Outer',projectItem:outerItem,
        videoTracks:collection([track([innerItem,media('three.png')])],'numTracks'),
        audioTracks:collection([],'numTracks')};
    const parent = {sequenceID:'parent',name:'Parent',
        videoTracks:collection([track([outerItem])],'numTracks'),audioTracks:collection([],'numTracks')};
    context.app = {project:{sequences:collection([parent,outer,inner],'numSequences')}};
    const usage = JSON.parse(context.pcTrackUsageJson(context.pcTrackCollectionUsage(parent.videoTracks,'V',null)));
    const copiedPaths = [];
    context.pcCollectSequenceMedia(parent,{},context.pcBuildSequenceMap(),{},{},copiedPaths,{},[]);
    assert.equal(usage[0].clipCount, 1);
    assert.deepEqual(usage[0].mediaPaths.slice().sort(), ['D:/Media/music.wav','D:/Media/one.png','D:/Media/three.png','D:/Media/two.png']);
    assert.deepEqual(copiedPaths.slice().sort(), usage[0].mediaPaths.slice().sort());
    const ui = loadCollectorLogic(createTestDocument());
    ui.previewUsage = usage;
    const countText = vm.runInContext(`(() => {
        const filter = createSequenceFilter('parent','Parent',previewUsage,[],false);
        return renderTrackButtonGroup(filter,'video').children[0].children[1].textContent;
    })()`, ui);
    assert.equal(countText, '4 files');
});

test('default track-lock preset follows current locks independently for video and audio', () => {
    const context = loadCollectorLogic();
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const filter = createSequenceFilter('seq', 'Show',
            [{trackNumber:1,isLocked:true},{trackNumber:2,isLocked:false,isBackup:true}],
            [{trackNumber:1,isLocked:false},{trackNumber:2,isLocked:true}], false);
        validateTrackLockStates([filter]);
        const before = [filter.ignoredVideoTracks.slice(), filter.ignoredAudioTracks.slice()];
        mergeSequenceTrackUsage(filter, createSequenceFilter('seq', 'Show',
            [{trackNumber:1,isLocked:false},{trackNumber:2,isLocked:true}],
            [{trackNumber:1,isLocked:false},{trackNumber:2,isLocked:false}], false));
        validateTrackLockStates([filter]);
        return {before,after:[filter.ignoredVideoTracks,filter.ignoredAudioTracks],preset:filter.selectedPresetId};
    })())`, context));
    assert.deepEqual(result.before, [[1], [2]]);
    assert.deepEqual(result.after, [[2], []]);
    assert.equal(result.preset, 'builtin-track-locks');
    assert.throws(() => vm.runInContext(`(() => {
        const filter = createSequenceFilter('seq','Show',[{trackNumber:1}],[],false);
        applyTrackPresetToFilter(filter,TRACK_LOCK_PRESET);
        validateTrackLockStates([filter]);
    })()`, context), /could not read track locks/);
});

test('copy-time refresh replaces obsolete manual choices with live Premiere locks', async () => {
    const context = loadCollectorLogic(createTestDocument());
    const result = JSON.parse(await vm.runInContext(`(async () => {
        selectedSequenceFilters = [createSequenceFilter('seq', 'Show', [], [], false)];
        selectedSequenceFilters[0].selectedPresetId = '';
        selectedSequenceFilters[0].ignoredVideoTracks = [1];
        callHost = async () => JSON.stringify({sequences:[{
            sequenceID:'seq',sequenceName:'Show',
            videoTrackUsage:[{trackNumber:1,isLocked:false},{trackNumber:2,isLocked:true}],
            audioTrackUsage:[{trackNumber:1,isLocked:true}]
        }]});
        renderSequenceFilters = () => {};
        await refreshAllSelectedSequenceTracks(false);
        validateTrackLockStates(selectedSequenceFilters);
        return JSON.stringify(selectedSequenceFilters[0]);
    })()`, context));
    assert.equal(result.selectedPresetId, 'builtin-track-locks');
    assert.deepEqual(result.ignoredVideoTracks, [2]);
    assert.deepEqual(result.ignoredAudioTracks, [1]);
});

test('lock mode track indicators cannot change the ignore selection', () => {
    const context = loadCollectorLogic(createTestDocument());
    const result = vm.runInContext(`(() => {
        const filter = createSequenceFilter('seq','Show',[{trackNumber:1,isLocked:false,clipCount:1}],[],false);
        applyTrackPresetToFilter(filter,TRACK_LOCK_PRESET);
        const button = renderTrackButtonGroup(filter,'video').children[0].children[0];
        return {click:button.onclick,doubleClick:button.ondblclick,title:button.title};
    })()`, context);
    assert.equal(result.click, null);
    assert.equal(result.doubleClick, null);
    assert.match(result.title, /unlocked: included/);
});

test('automatic lock polling updates changed locks only and discards responses after copying starts', async () => {
    const context = loadCollectorLogic(createTestDocument());
    const result = JSON.parse(await vm.runInContext(`(async () => {
        latestPlan = {projectPath:'D:/Show.prproj'};
        selectedSequenceFilters = [createSequenceFilter('seq','Show',[{trackNumber:1,isLocked:false,clipCount:1}],[],false)];
        let renders = 0, calls = 0, locked = true;
        renderSequenceFilters = () => {renders++;};
        const snapshot = () => JSON.stringify({projectPath:'D:/Show.prproj',sequences:[{sequenceID:'seq',sequenceName:'Show',videoLocks:[locked],audioLocks:[]}]});
        callHost = async () => {calls++; return snapshot();};
        await pollCollectorTrackLocks();
        const ignored = selectedSequenceFilters[0].ignoredVideoTracks.slice();
        await pollCollectorTrackLocks();
        const unchangedRenders = renders;
        locked = false;
        await pollCollectorTrackLocks();
        const included = selectedSequenceFilters[0].ignoredVideoTracks.slice();
        window.frameElement = {hidden:true};
        await pollCollectorTrackLocks();
        window.frameElement.hidden = false;
        isCopying = true;
        await pollCollectorTrackLocks();
        isCopying = false;
        const callsBeforeDelayed = calls;
        let finish;
        locked = true;
        callHost = () => new Promise(resolve => {finish = resolve;});
        const pending = pollCollectorTrackLocks();
        collectionGeneration++;
        finish(snapshot());
        await pending;
        return JSON.stringify({ignored,included,unchangedRenders,renders,callsBeforeDelayed,finalIgnored:selectedSequenceFilters[0].ignoredVideoTracks});
    })()`,context));
    assert.deepEqual(result.ignored,[1]);
    assert.deepEqual(result.included,[]);
    assert.equal(result.unchangedRenders,1);
    assert.equal(result.renders,2);
    assert.equal(result.callsBeforeDelayed,3);
    assert.deepEqual(result.finalIgnored,[]);
});

test('sequence tabs switch the visible panel without changing collection selections', () => {
    const context = loadCollectorLogic(createTestDocument());
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        selectedSequenceFilters = [createSequenceFilter('a','Full sequence name A',[],[],false), createSequenceFilter('b','Full sequence name B',[],[],false)];
        updateSelectionSummary = () => {};
        const before = JSON.stringify(selectedSequenceFilters);
        renderSequenceFilters();
        const container = document.getElementById('sequenceFilters');
        const first = container.children.slice(1).map(card => card.hidden);
        const labels = container.children[0].children.map(tab => tab.children[0].textContent);
        const selectSecond = container.children[0].children[1].children[0].onclick;
        // This lightweight DOM mock does not clear children on innerHTML assignment.
        container.children.length = 0;
        selectSecond();
        const second = container.children.slice(1).map(card => card.hidden);
        return {first,second,labels,unchanged:before === JSON.stringify(selectedSequenceFilters)};
    })())`, context));
    assert.deepEqual(result.first, [false,true]);
    assert.deepEqual(result.second, [true,false]);
    assert.deepEqual(result.labels, ['Full sequence name A','Full sequence name B']);
    assert.equal(result.unchanged, true);
});

test('track display hides ignored tracks without changing copy filters or usage', () => {
    const context = loadCollectorLogic(createTestDocument());
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const filter = createSequenceFilter('seq','Show',[
            {trackNumber:1,isLocked:false,clipCount:2},
            {trackNumber:2,isLocked:true,clipCount:3}
        ],[{trackNumber:1,isLocked:true,clipCount:1}],false);
        const before = JSON.stringify(filter);
        const video = renderTrackButtonGroup(filter,'video');
        const audio = renderTrackButtonGroup(filter,'audio');
        return {unchanged:before === JSON.stringify(filter),
            labels:video.children.map(item => item.children[0].children[0].textContent),
            audio:audio.children[0].textContent,ignored:filter.ignoredVideoTracks};
    })())`, context));
    assert.equal(result.unchanged,true);
    assert.deepEqual(result.labels,['V1']);
    assert.deepEqual(result.ignored,[2]);
    assert.equal(result.audio,'No audio tracks selected for copying.');
});

test('host detects full backup MP4/audio outputs but preserves short source excerpts', () => {
    const context = vm.createContext({
        File: function(value) { this.name = String(value).split(/[\\/]/).pop(); this.fsName = value; }
    });
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'jsx', 'collector.jsx'), 'utf8'), context);
    const clip = (name, duration) => ({start:{seconds:0}, end:{seconds:duration}, projectItem:{name, getMediaPath:() => 'D:/' + name}});
    const track = (name, duration) => ({clips:{0:clip(name,duration), numItems:1}});
    const sequence = {name:'Show', end:String(60 * 254016000000)};
    const tracks = {0:track('Other_BACKUP.mp4', 5), 1:track('Show_Track2-3.wav',60), 2:track('Show_BACKUP_REBKP_TEMP.mp4',60), numTracks:3};
    const audio = JSON.parse(context.pcTrackUsageJson(context.pcTrackCollectionUsage(tracks, 'A', sequence)));
    const video = JSON.parse(context.pcTrackUsageJson(context.pcTrackCollectionUsage(tracks, 'V', sequence)));
    assert.deepEqual(audio.map(t => t.isBackup), [false,true,true]);
    assert.deepEqual(video.map(t => t.isBackup), [false,false,true]);
});

test('panel uses the compact destination labels and places sequence options with backup options', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

    assert.match(html, />Skip files existing here</);
    assert.match(html, />Choose Path</);
    assert.doesNotMatch(html, /1\. Skip Existing Media/);
    assert.doesNotMatch(html, /Files already in this folder are skipped/);
    assert.doesNotMatch(html, /2\. Backup Destination/);
    assert.match(html, /Lock a track in Premiere to ignore its files\. Click Refresh after making changes to the sequence\./);
    assert.match(html, /Track locks are read from Premiere before copying/);
    assert.ok(html.indexOf('id="sequenceOnlyMode"') < html.indexOf('id="copyProjectFile"'));
    assert.ok(html.indexOf('id="createReducedProject"') < html.indexOf('id="copyProjectFile"'));
    assert.ok(html.indexOf('id="linkProjectAfterCollection"') < html.indexOf('id="collectButton"'));
    assert.match(html, /id="linkExistingBackupButton"[^>]*>LINK AN EXISTING BACKUP</);
    assert.match(html, /No files are copied again/);
});

test('completion dialog clearly switches between green success and red error states', () => {
    const testDocument = createTestDocument();
    const context = loadCollectorLogic(testDocument);
    const prompt = testDocument.getElementById('completionPrompt');
    const title = testDocument.getElementById('completionTitle');
    const message = testDocument.getElementById('completionMessage');

    vm.runInContext(`showCompletionPrompt(true, 'Backup and linking finished without errors.')`, context);
    assert.equal(prompt.classList.contains('is-visible'), true);
    assert.equal(prompt.classList.contains('is-success'), true);
    assert.equal(prompt.classList.contains('is-error'), false);
    assert.equal(title.textContent, 'Done without error');
    assert.equal(message.textContent, 'Backup and linking finished without errors.');

    vm.runInContext(`showCompletionPrompt(false, 'Backup and linking finished with 2 errors.')`, context);
    assert.equal(prompt.classList.contains('is-success'), false);
    assert.equal(prompt.classList.contains('is-error'), true);
    assert.equal(title.textContent, 'Finished with errors');
    assert.equal(message.textContent, 'Backup and linking finished with 2 errors.');

    vm.runInContext('hideCompletionPrompt()', context);
    assert.equal(prompt.classList.contains('is-visible'), false);
    assert.equal(prompt.classList.contains('is-success'), false);
    assert.equal(prompt.classList.contains('is-error'), false);
});

test('destination folders are created automatically without newer recursive-mkdir options', () => {
    const context = loadCollectorLogic();
    const rootPath = 'Z:\\';
    const targetPath = 'Z:\\2017_SMTV2 PROJECT FOLDER\\@ SHOWS\\PE\\PE 3219 test\\PE 3219';
    const createdDirectories = new Set([rootPath.toLowerCase()]);
    const mkdirCalls = [];
    const directoryFs = {
        existsSync(candidatePath) {
            return createdDirectories.has(String(candidatePath).toLowerCase());
        },
        statSync(candidatePath) {
            if (!this.existsSync(candidatePath)) {
                throw new Error(`Missing test directory: ${candidatePath}`);
            }
            return {
                isDirectory() {
                    return true;
                }
            };
        },
        mkdirSync(candidatePath) {
            assert.equal(arguments.length, 1);
            const parentPath = path.win32.dirname(candidatePath);
            assert.equal(createdDirectories.has(parentPath.toLowerCase()), true);
            createdDirectories.add(String(candidatePath).toLowerCase());
            mkdirCalls.push(candidatePath);
        }
    };
    context.testDirectoryFs = directoryFs;

    vm.runInContext(
        `ensureDirectorySync(${JSON.stringify(targetPath)}, testDirectoryFs)`,
        context
    );

    assert.equal(createdDirectories.has(targetPath.toLowerCase()), true);
    assert.equal(mkdirCalls.length, 5);

    vm.runInContext(
        `ensureDirectorySync(${JSON.stringify(targetPath)}, testDirectoryFs)`,
        context
    );
    assert.equal(mkdirCalls.length, 5);
});

test('destination creation falls back to the native CEP folder API on mapped drives', () => {
    const context = loadCollectorLogic();
    const rootPath = 'Z:\\';
    const targetPath = 'Z:\\Shows\\PE 3219 test\\PE 3219';
    const createdDirectories = new Set([rootPath.toLowerCase()]);
    const cepCalls = [];
    const directoryFs = {
        existsSync(candidatePath) {
            return createdDirectories.has(String(candidatePath).toLowerCase());
        },
        statSync(candidatePath) {
            if (!this.existsSync(candidatePath)) {
                throw new Error(`Missing test directory: ${candidatePath}`);
            }
            return {
                isDirectory() {
                    return true;
                }
            };
        },
        mkdirSync(candidatePath) {
            const error = new Error(`operation not permitted, mkdir '${candidatePath}'`);
            error.code = 'EPERM';
            throw error;
        }
    };
    const cepFileSystem = {
        makedir(candidatePath) {
            const parentPath = path.win32.dirname(candidatePath);
            if (!createdDirectories.has(parentPath.toLowerCase())) {
                return { err: 3 };
            }
            createdDirectories.add(String(candidatePath).toLowerCase());
            cepCalls.push(candidatePath);
            return { err: 0 };
        }
    };
    context.testDirectoryFs = directoryFs;
    context.testCepFileSystem = cepFileSystem;

    vm.runInContext(
        `ensureDirectorySync(${JSON.stringify(targetPath)}, testDirectoryFs, testCepFileSystem)`,
        context
    );

    assert.equal(createdDirectories.has(targetPath.toLowerCase()), true);
    assert.equal(cepCalls.length, 3);
});

test('destination creation still surfaces genuine network permission failures', () => {
    const context = loadCollectorLogic();
    const rootPath = 'Z:\\';
    const blockedPath = 'Z:\\Restricted';
    const directoryFs = {
        existsSync(candidatePath) {
            return String(candidatePath).toLowerCase() === rootPath.toLowerCase();
        },
        statSync(candidatePath) {
            if (!this.existsSync(candidatePath)) {
                throw new Error(`Missing test directory: ${candidatePath}`);
            }
            return {
                isDirectory() {
                    return true;
                }
            };
        },
        mkdirSync(candidatePath) {
            const error = new Error(`operation not permitted, mkdir '${candidatePath}'`);
            error.code = 'EPERM';
            throw error;
        }
    };
    context.testDirectoryFs = directoryFs;

    assert.throws(
        () => vm.runInContext(
            `ensureDirectorySync(${JSON.stringify(blockedPath)}, testDirectoryFs)`,
            context
        ),
        (error) => error && error.code === 'EPERM'
    );
});

test('unresolved included-and-ignored conflict never copies', () => {
    const context = loadCollectorLogic();
    const reason = vm.runInContext(`(() => {
        const task = { name: 'shared.mov', source: 'C:/Media/shared.mov', binPath: '' };
        const mediaKey = normalizeMediaKey(task.source);
        const trackRuleContext = buildTrackRuleContext(
            [task],
            new Set([mediaKey]),
            new Set([mediaKey])
        );
        const copyRuleContext = createCopyRuleContext({
            treeSelectedTaskSet: new Set([task]),
            sequenceScopedMediaSet: new Set([mediaKey]),
            trackRuleContext
        });
        return getCopySkipReason(task, copyRuleContext);
    })()`, context);

    assert.equal(reason, 'skipped because its track conflict was not resolved');
});

test('ignored Premiere project folder still wins', () => {
    const context = loadCollectorLogic();
    const reason = vm.runInContext(`(() => {
        const task = { name: 'shared.mov', source: 'C:/Media/shared.mov', binPath: 'Do Not Back Up' };
        const copyRuleContext = createCopyRuleContext({
            treeSelectedTaskSet: new Set([task]),
            trackRuleContext: buildTrackRuleContext([task], new Set(), new Set()),
            ignoredProjectFolders: ['Do Not Back Up']
        });
        return getCopySkipReason(task, copyRuleContext);
    })()`, context);

    assert.equal(reason, 'skipped because Premiere folder "Do Not Back Up" is ignored');
});

test('ignored-track matching does not widen to a different file with the same name and size', (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-exact-path-'));
    const ignoredPath = path.join(tempRoot, 'ignored', 'shared.mov');
    const includedPath = path.join(tempRoot, 'included', 'shared.mov');
    fs.mkdirSync(path.dirname(ignoredPath), { recursive: true });
    fs.mkdirSync(path.dirname(includedPath), { recursive: true });
    fs.writeFileSync(ignoredPath, 'same-size-content');
    fs.writeFileSync(includedPath, 'same-size-content');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic();
    const result = vm.runInContext(`(() => {
        const ignoredTask = { name: 'shared.mov', source: ${JSON.stringify(ignoredPath)}, binPath: '' };
        const includedTask = { name: 'shared.mov', source: ${JSON.stringify(includedPath)}, binPath: '' };
        const ignoredMediaSet = new Set([normalizeMediaKey(ignoredTask.source)]);
        const includedMediaSet = new Set([
            normalizeMediaKey(ignoredTask.source),
            normalizeMediaKey(includedTask.source)
        ]);
        const trackRuleContext = buildTrackRuleContext(
            [ignoredTask, includedTask],
            includedMediaSet,
            ignoredMediaSet
        );
        return {
            unrelatedFileIgnored: trackRuleContext.ignoredOnlyMediaSet.has(
                normalizeMediaKey(includedTask.source)
            ),
            conflicts: trackRuleContext.conflicts.map((conflict) => conflict.source)
        };
    })()`, context);

    assert.equal(result.unrelatedFileIgnored, false);
    assert.deepEqual(Array.from(result.conflicts), [ignoredPath]);
});

test('root copy-rule precedence is deterministic across every rule source', () => {
    const context = loadCollectorLogic();
    const result = vm.runInContext(`(() => {
        function evaluate(options) {
            const task = {
                name: options.name || 'media.mov',
                source: options.source,
                binPath: options.binPath || ''
            };
            const mediaKey = normalizeMediaKey(task.source);
            const includedMediaSet = options.includedTrack ? new Set([mediaKey]) : new Set();
            const ignoredMediaSet = options.ignoredTrack ? new Set([mediaKey]) : new Set();
            const trackRuleContext = buildTrackRuleContext(
                [task],
                includedMediaSet,
                ignoredMediaSet
            );
            let conflictDecisionInfo = null;
            if (options.decision) {
                conflictDecisionInfo = buildTrackConflictDecisionInfo(trackRuleContext, [{
                    mediaKey,
                    decision: options.decision
                }]);
            }
            const copyRuleContext = createCopyRuleContext({
                treeSelectedTaskSet: options.sourceSelected ? new Set([task]) : new Set(),
                sequenceScopedMediaSet: options.sequenceScoped
                    ? (options.inSequence ? new Set([mediaKey]) : new Set())
                    : null,
                trackRuleContext,
                conflictDecisionInfo,
                includedProjectFolders: options.includedFolder ? [task.binPath] : [],
                ignoredProjectFolders: options.ignoredFolder ? [task.binPath] : []
            });
            return getCopySkipReason(task, copyRuleContext);
        }

        return {
            ignoredFolderBeforeCopyDecision: evaluate({
                source: 'C:/Media/folder-ignore.mov',
                binPath: 'Ignore',
                sourceSelected: true,
                includedTrack: true,
                ignoredTrack: true,
                decision: 'copy',
                ignoredFolder: true
            }),
            sequenceScopeBeforeCopyDecision: evaluate({
                source: 'C:/Media/out-of-sequence.mov',
                sourceSelected: true,
                sequenceScoped: true,
                inSequence: false,
                includedTrack: true,
                ignoredTrack: true,
                decision: 'copy'
            }),
            conflictCopyBeforeSourceList: evaluate({
                source: 'C:/Media/conflict-copy.mov',
                sourceSelected: false,
                includedTrack: true,
                ignoredTrack: true,
                decision: 'copy'
            }),
            conflictSkip: evaluate({
                source: 'C:/Media/conflict-skip.mov',
                sourceSelected: true,
                includedTrack: true,
                ignoredTrack: true,
                decision: 'skip'
            }),
            unresolvedConflict: evaluate({
                source: 'C:/Media/conflict-unresolved.mov',
                sourceSelected: true,
                includedTrack: true,
                ignoredTrack: true
            }),
            ignoredTrackBeforeIncludedFolder: evaluate({
                source: 'C:/Media/ignored-track.mov',
                binPath: 'Force',
                sourceSelected: true,
                ignoredTrack: true,
                includedFolder: true
            }),
            includedFolderBeforeSourceList: evaluate({
                source: 'C:/Media/included-folder.mov',
                binPath: 'Force',
                sourceSelected: false,
                includedFolder: true
            }),
            sourceListSkip: evaluate({
                source: 'C:/Media/source-list-skip.mov',
                sourceSelected: false
            }),
            normalCopy: evaluate({
                source: 'C:/Media/normal-copy.mov',
                sourceSelected: true
            })
        };
    })()`, context);

    assert.equal(result.ignoredFolderBeforeCopyDecision, 'skipped because Premiere folder "Ignore" is ignored');
    assert.equal(result.sequenceScopeBeforeCopyDecision, 'skipped because it is not used by the chosen sequences');
    assert.equal(result.conflictCopyBeforeSourceList, '');
    assert.equal(result.conflictSkip, 'skipped by your track conflict decision');
    assert.equal(result.unresolvedConflict, 'skipped because its track conflict was not resolved');
    assert.equal(result.ignoredTrackBeforeIncludedFolder, 'skipped by ignored track selection');
    assert.equal(result.includedFolderBeforeSourceList, '');
    assert.equal(result.sourceListSkip, 'skipped by Source File List selection');
    assert.equal(result.normalCopy, '');
});

test('validated conflict decisions reject missing, duplicate, and out-of-scope media', () => {
    const context = loadCollectorLogic();
    const result = vm.runInContext(`(() => {
        const conflictTask = { name: 'conflict.mov', source: 'C:/Media/conflict.mov', binPath: '' };
        const otherTask = { name: 'other.mov', source: 'C:/Media/other.mov', binPath: '' };
        const conflictKey = normalizeMediaKey(conflictTask.source);
        const otherKey = normalizeMediaKey(otherTask.source);
        const trackRuleContext = buildTrackRuleContext(
            [conflictTask, otherTask],
            new Set([conflictKey, otherKey]),
            new Set([conflictKey])
        );
        return {
            valid: buildTrackConflictDecisionInfo(trackRuleContext, [{
                mediaKey: conflictKey,
                decision: 'copy'
            }]).valid,
            missing: buildTrackConflictDecisionInfo(trackRuleContext, []).valid,
            duplicate: buildTrackConflictDecisionInfo(trackRuleContext, [{
                mediaKey: conflictKey,
                decision: 'copy'
            }, {
                mediaKey: conflictKey,
                decision: 'skip'
            }]).valid,
            outside: buildTrackConflictDecisionInfo(trackRuleContext, [{
                mediaKey: otherKey,
                decision: 'skip'
            }]).valid
        };
    })()`, context);

    assert.equal(result.valid, true);
    assert.equal(result.missing, false);
    assert.equal(result.duplicate, false);
    assert.equal(result.outside, false);
});

test('collection automatically skips shared ignored media and copies allowed media', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-ignored-'));
    const sourceDir = path.join(tempRoot, 'source');
    const destinationDir = path.join(tempRoot, 'backup');
    const ignoredSourcePath = path.join(sourceDir, 'shared.mov');
    const allowedSourcePath = path.join(sourceDir, 'allowed.mov');
    const collectedRoot = path.join(destinationDir, 'Ignored_Track_Project', 'CollectedMedias');
    fs.mkdirSync(sourceDir, { recursive: true });
    fs.writeFileSync(ignoredSourcePath, 'ignored-shared-media');
    fs.writeFileSync(allowedSourcePath, 'allowed-media');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic(createTestDocument());
    vm.runInContext(`
        destination = ${JSON.stringify(destinationDir)};
        compareLocation = null;
        copyProjectFile = false;
        linkProjectAfterCollection = false;
        sequenceOnlyMode = false;
        createReducedProject = false;
        latestPlan = {
            projectName: 'Ignored_Track_Project',
            missingMedia: [],
            tasks: [{
                name: 'shared.mov',
                source: ${JSON.stringify(ignoredSourcePath)},
                destination: 'Media/shared.mov',
                binPath: '',
                relativePath: 'Media/shared.mov'
            }, {
                name: 'allowed.mov',
                source: ${JSON.stringify(allowedSourcePath)},
                destination: 'Media/allowed.mov',
                binPath: '',
                relativePath: 'Media/allowed.mov'
            }]
        };
        sourceTree = [];
        buildCopyReadyContext = async function buildCopyReadyContextForTest() {
            const treeSelectedTaskSet = new Set(latestPlan.tasks);
            const ignoredMediaSet = new Set([normalizeMediaKey(latestPlan.tasks[0].source)]);
            const trackRuleContext = buildTrackRuleContext(
                latestPlan.tasks,
                new Set(),
                ignoredMediaSet
            );
            return {
                ok: true,
                copyRuleContext: createCopyRuleContext({
                    treeSelectedTaskSet,
                    trackRuleContext
                }),
                copyWarnings: [],
                sequenceScopeInfo: null,
                trackConflicts: trackRuleContext.conflicts
            };
        };
    `, context);

    await vm.runInContext('collect()', context);

    assert.equal(fs.existsSync(path.join(collectedRoot, 'shared.mov')), false);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'allowed.mov')), true);
    assert.equal(fs.readFileSync(path.join(collectedRoot, 'allowed.mov'), 'utf8'), 'allowed-media');
    assert.equal(vm.runInContext('isCopying', context), false);
});

test('individual row clicks apply mixed Copy and Do not copy choices to the real copy loop', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-individual-'));
    const sourceDir = path.join(tempRoot, 'source');
    const destinationDir = path.join(tempRoot, 'backup');
    const copySourcePath = path.join(sourceDir, 'copy-this.mov');
    const skipSourcePath = path.join(sourceDir, 'skip-this.mov');
    const allowedSourcePath = path.join(sourceDir, 'allowed.mov');
    const collectedRoot = path.join(destinationDir, 'Individual_Choice_Project', 'CollectedMedias');
    fs.mkdirSync(sourceDir, { recursive: true });
    fs.writeFileSync(copySourcePath, 'copy-choice-media');
    fs.writeFileSync(skipSourcePath, 'skip-choice-media');
    fs.writeFileSync(allowedSourcePath, 'allowed-media');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic(createTestDocument());
    vm.runInContext(`
        destination = ${JSON.stringify(destinationDir)};
        compareLocation = null;
        copyProjectFile = false;
        linkProjectAfterCollection = false;
        sequenceOnlyMode = false;
        createReducedProject = false;
        latestPlan = {
            projectName: 'Individual_Choice_Project',
            missingMedia: [],
            tasks: [{
                name: 'copy-this.mov',
                source: ${JSON.stringify(copySourcePath)},
                destination: 'Media/copy-this.mov',
                binPath: '',
                relativePath: 'Media/copy-this.mov'
            }, {
                name: 'skip-this.mov',
                source: ${JSON.stringify(skipSourcePath)},
                destination: 'Media/skip-this.mov',
                binPath: '',
                relativePath: 'Media/skip-this.mov'
            }, {
                name: 'allowed.mov',
                source: ${JSON.stringify(allowedSourcePath)},
                destination: 'Media/allowed.mov',
                binPath: '',
                relativePath: 'Media/allowed.mov'
            }]
        };
        sourceTree = [];
        buildCopyReadyContext = async function buildCopyReadyContextForIndividualChoiceTest() {
            const treeSelectedTaskSet = new Set(latestPlan.tasks);
            const ignoredMediaSet = new Set([
                normalizeMediaKey(latestPlan.tasks[0].source),
                normalizeMediaKey(latestPlan.tasks[1].source)
            ]);
            const includedMediaSet = new Set([
                normalizeMediaKey(latestPlan.tasks[0].source),
                normalizeMediaKey(latestPlan.tasks[1].source)
            ]);
            const trackRuleContext = buildTrackRuleContext(
                latestPlan.tasks,
                includedMediaSet,
                ignoredMediaSet
            );
            return {
                ok: true,
                copyRuleContext: createCopyRuleContext({
                    treeSelectedTaskSet,
                    trackRuleContext
                }),
                copyWarnings: [],
                sequenceScopeInfo: null,
                trackConflicts: trackRuleContext.conflicts
            };
        };
    `, context);

    const collectionPromise = vm.runInContext('collect()', context);
    await new Promise((resolve) => setImmediate(resolve));

    const stateBeforeContinue = vm.runInContext(`(() => {
        const list = document.getElementById('trackConflictList');
        list.onclick({ target: trackConflictPromptState.rows[0].copyButton });
        list.onclick({ target: trackConflictPromptState.rows[1].skipButton });
        return {
            decisions: trackConflictPromptState.decisions.slice(),
            continueDisabled: document.getElementById('trackConflictContinueButton').disabled
        };
    })()`, context);

    assert.deepEqual(Array.from(stateBeforeContinue.decisions), ['copy', 'skip']);
    assert.equal(stateBeforeContinue.continueDisabled, false);

    vm.runInContext('finishTrackConflictPrompt(true)', context);
    await collectionPromise;

    assert.equal(fs.existsSync(path.join(collectedRoot, 'copy-this.mov')), true);
    assert.equal(fs.readFileSync(path.join(collectedRoot, 'copy-this.mov'), 'utf8'), 'copy-choice-media');
    assert.equal(fs.existsSync(path.join(collectedRoot, 'skip-this.mov')), false);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'allowed.mov')), true);
    assert.equal(vm.runInContext('isCopying', context), false);
});

test('Do not copy all skips only listed conflicts and still copies unrelated included media', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-skip-all-'));
    const sourceDir = path.join(tempRoot, 'source');
    const destinationDir = path.join(tempRoot, 'backup');
    const fileNames = ['conflict-one.mov', 'conflict-two.mov', 'included-one.mov', 'included-two.mov'];
    const sourcePaths = fileNames.map((fileName) => path.join(sourceDir, fileName));
    const collectedRoot = path.join(destinationDir, 'Skip_All_Scope_Project', 'CollectedMedias');
    fs.mkdirSync(sourceDir, { recursive: true });
    sourcePaths.forEach((filePath, index) => {
        fs.writeFileSync(filePath, `media-${index + 1}`);
    });
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic(createTestDocument());
    vm.runInContext(`
        destination = ${JSON.stringify(destinationDir)};
        compareLocation = null;
        copyProjectFile = false;
        linkProjectAfterCollection = false;
        sequenceOnlyMode = false;
        createReducedProject = false;
        latestPlan = {
            projectName: 'Skip_All_Scope_Project',
            missingMedia: [],
            tasks: ${JSON.stringify(sourcePaths.map((source, index) => ({
                name: fileNames[index],
                source,
                destination: `Media/${fileNames[index]}`,
                binPath: '',
                relativePath: `Media/${fileNames[index]}`
            })))}
        };
        sourceTree = [];
        buildCopyReadyContext = async function buildCopyReadyContextForSkipAllTest() {
            const treeSelectedTaskSet = new Set(latestPlan.tasks);
            const ignoredMediaSet = new Set([
                normalizeMediaKey(latestPlan.tasks[0].source),
                normalizeMediaKey(latestPlan.tasks[1].source)
            ]);
            const includedMediaSet = new Set(latestPlan.tasks.map((task) => normalizeMediaKey(task.source)));
            const trackRuleContext = buildTrackRuleContext(
                latestPlan.tasks,
                includedMediaSet,
                ignoredMediaSet
            );
            return {
                ok: true,
                copyRuleContext: createCopyRuleContext({
                    treeSelectedTaskSet,
                    trackRuleContext
                }),
                copyWarnings: [],
                sequenceScopeInfo: null,
                trackConflicts: trackRuleContext.conflicts
            };
        };
    `, context);

    const collectionPromise = vm.runInContext('collect()', context);
    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(vm.runInContext('trackConflictPromptState.conflicts.length', context), 2);

    vm.runInContext(`
        setAllTrackConflictChoices('skip');
        finishTrackConflictPrompt(true);
    `, context);
    await collectionPromise;

    assert.equal(fs.existsSync(path.join(collectedRoot, 'conflict-one.mov')), false);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'conflict-two.mov')), false);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'included-one.mov')), true);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'included-two.mov')), true);
    assert.equal(vm.runInContext('isCopying', context), false);
});

test('a refreshed plan keeps unrelated selected files when every conflict is skipped', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'collector-refreshed-conflicts-'));
    t.after(() => fs.rmSync(tempRoot, {recursive:true, force:true}));
    const files = ['shared.mov', 'included.mov', 'unchecked.mov'];
    files.forEach(name => fs.writeFileSync(path.join(tempRoot, name), name));
    const context = loadCollectorLogic(createTestDocument());
    vm.runInContext(`
        destination = ${JSON.stringify(path.join(tempRoot, 'backup'))};
        copyProjectFile = false;
        sequenceOnlyMode = false;
        latestPlan = { projectName:'RefreshTest', missingMedia:[], tasks:${JSON.stringify(files.map(name => ({
            name, source:path.join(tempRoot, name), destination:'Media/' + name, binPath:'Media', relativePath:'Media/' + name
        })))} };
        sourceTree = [];
        buildCopyReadyContext = async () => {
            const trackRuleContext = buildTrackRuleContext(latestPlan.tasks,
                new Set(latestPlan.tasks.map(t => normalizeMediaKey(t.source))),
                new Set([normalizeMediaKey(latestPlan.tasks[0].source)]));
            return {ok:true, copyWarnings:[], sequenceScopeInfo:null,
                copyRuleContext:createCopyRuleContext({treeSelectedTaskSet:new Set(latestPlan.tasks.slice(0,2)), trackRuleContext}),
                trackConflicts:trackRuleContext.conflicts};
        };
    `, context);
    const pending = context.collect();
    await new Promise(resolve => setImmediate(resolve));
    // A host refresh returns newly parsed task objects for exactly the same files.
    vm.runInContext(`
        latestPlan = JSON.parse(JSON.stringify(latestPlan));
        setAllTrackConflictChoices('skip');
        finishTrackConflictPrompt(true);
    `, context);
    await pending;
    assert.equal(fs.existsSync(path.join(tempRoot, 'backup', 'RefreshTest', 'Media', 'included.mov')), true);
    assert.equal(fs.existsSync(path.join(tempRoot, 'backup', 'RefreshTest', 'Media', 'shared.mov')), false);
    assert.equal(fs.existsSync(path.join(tempRoot, 'backup', 'RefreshTest', 'Media', 'unchecked.mov')), false);
});

test('refresh started before copying cannot replace the active plan, even after copying finishes', async () => {
    for (const finishBeforeResponse of [false, true]) {
        const context = loadCollectorLogic(createTestDocument());
        let completeRefresh;
        context.ensureHostScriptLoaded = async () => true;
        context.callHost = () => new Promise(resolve => {completeRefresh = resolve;});
        vm.runInContext(`latestPlan = {projectName:'Original', tasks:[]};`, context);
        const pendingRefresh = context.loadProjectPlan();
        await new Promise(resolve => setImmediate(resolve));
        context.setBusyState(true);
        if (finishBeforeResponse) context.setBusyState(false);
        completeRefresh(JSON.stringify({projectName:'Stale refresh',tasks:[]}));
        assert.equal(await pendingRefresh, false);
        assert.equal(vm.runInContext('latestPlan.projectName', context), 'Original');
    }
});

test('stable task selection does not include an unchecked duplicate in a different bin or a different source', () => {
    const context = loadCollectorLogic();
    const result = vm.runInContext(`(() => {
        const selected = {source:'D:/one/clip.mov', destination:'Media/clip.mov', binPath:'Media'};
        const rules = createCopyRuleContext({treeSelectedTaskSet:new Set([selected])});
        return [
            getCopySkipReason(JSON.parse(JSON.stringify(selected)), rules),
            getCopySkipReason({...selected, source:'D:/two/clip.mov'}, rules),
            getCopySkipReason({...selected, binPath:'Unused'}, rules)
        ];
    })()`, context);
    assert.equal(result[0], '');
    assert.match(result[1], /Source File List/);
    assert.match(result[2], /Source File List/);
});

test('Copy all copies listed conflicts without selecting unrelated media', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-copy-all-'));
    const sourceDir = path.join(tempRoot, 'source');
    const destinationDir = path.join(tempRoot, 'backup');
    const fileNames = ['conflict-one.mov', 'conflict-two.mov', 'unselected.mov'];
    const sourcePaths = fileNames.map((fileName) => path.join(sourceDir, fileName));
    const collectedRoot = path.join(destinationDir, 'Copy_All_Scope_Project', 'CollectedMedias');
    fs.mkdirSync(sourceDir, { recursive: true });
    sourcePaths.forEach((filePath, index) => {
        fs.writeFileSync(filePath, `copy-all-media-${index + 1}`);
    });
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic(createTestDocument());
    vm.runInContext(`
        destination = ${JSON.stringify(destinationDir)};
        compareLocation = null;
        copyProjectFile = false;
        linkProjectAfterCollection = false;
        sequenceOnlyMode = false;
        createReducedProject = false;
        latestPlan = {
            projectName: 'Copy_All_Scope_Project',
            missingMedia: [],
            tasks: ${JSON.stringify(sourcePaths.map((source, index) => ({
                name: fileNames[index],
                source,
                destination: `Media/${fileNames[index]}`,
                binPath: '',
                relativePath: `Media/${fileNames[index]}`
            })))}
        };
        sourceTree = [];
        buildCopyReadyContext = async function buildCopyReadyContextForCopyAllTest() {
            const treeSelectedTaskSet = new Set();
            const ignoredMediaSet = new Set([
                normalizeMediaKey(latestPlan.tasks[0].source),
                normalizeMediaKey(latestPlan.tasks[1].source)
            ]);
            const includedMediaSet = new Set(latestPlan.tasks.map((task) => normalizeMediaKey(task.source)));
            const trackRuleContext = buildTrackRuleContext(
                latestPlan.tasks,
                includedMediaSet,
                ignoredMediaSet
            );
            return {
                ok: true,
                copyRuleContext: createCopyRuleContext({
                    treeSelectedTaskSet,
                    trackRuleContext
                }),
                copyWarnings: [],
                sequenceScopeInfo: null,
                trackConflicts: trackRuleContext.conflicts
            };
        };
    `, context);

    const collectionPromise = vm.runInContext('collect()', context);
    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(vm.runInContext('trackConflictPromptState.conflicts.length', context), 2);

    vm.runInContext(`
        setAllTrackConflictChoices('copy');
        finishTrackConflictPrompt(true);
    `, context);
    await collectionPromise;

    assert.equal(fs.existsSync(path.join(collectedRoot, 'conflict-one.mov')), true);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'conflict-two.mov')), true);
    assert.equal(fs.existsSync(path.join(collectedRoot, 'unselected.mov')), false);
    assert.equal(vm.runInContext('isCopying', context), false);
});

test('per-sequence track reset clears only the requested sequence', () => {
    const context = loadCollectorLogic();
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const filters = [
            {
                sequenceID: 'sequence-a',
                sequenceName: 'Sequence A',
                ignoredVideoTracks: [1, 3],
                ignoredAudioTracks: [2]
            },
            {
                sequenceID: 'sequence-b',
                sequenceName: 'Sequence B',
                ignoredVideoTracks: [4],
                ignoredAudioTracks: [1, 5]
            }
        ];
        const resetCount = clearTrackChoices(filters, 'sequence-b');
        return { filters, resetCount };
    })())`, context));

    assert.equal(result.resetCount, 1);
    assert.deepEqual(result.filters[0].ignoredVideoTracks, [1, 3]);
    assert.deepEqual(result.filters[0].ignoredAudioTracks, [2]);
    assert.deepEqual(result.filters[1].ignoredVideoTracks, []);
    assert.deepEqual(result.filters[1].ignoredAudioTracks, []);
});

test('global sequence refresh updates and resets every selected sequence without changing its scope', () => {
    const context = loadCollectorLogic();
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const filters = [
            createSequenceFilter('sequence-a', 'Sequence A', [
                { trackNumber: 1, clipCount: 1, hasItems: true }
            ], [], true),
            createSequenceFilter('sequence-b', 'Sequence B', [], [
                { trackNumber: 1, clipCount: 1, hasItems: true }
            ], false),
            createSequenceFilter('sequence-missing', 'Deleted Sequence', [
                { trackNumber: 9, clipCount: 9, hasItems: true }
            ], [], false)
        ];
        filters[0].ignoredVideoTracks = [1];
        filters[1].ignoredAudioTracks = [1];
        filters[2].ignoredVideoTracks = [9];

        const refreshResult = applySequenceTrackUsagePlan(filters, {
            sequences: [
                {
                    sequenceID: 'sequence-b',
                    sequenceName: 'Sequence B',
                    videoTrackUsage: [],
                    audioTrackUsage: [
                        { trackNumber: 2, clipCount: 7, hasItems: true }
                    ]
                },
                {
                    sequenceID: 'sequence-a',
                    sequenceName: 'Renamed Sequence A',
                    videoTrackUsage: [
                        { trackNumber: 3, clipCount: 5, hasItems: true }
                    ],
                    audioTrackUsage: []
                }
            ]
        }, true);

        return { filters, refreshResult };
    })())`, context));

    assert.equal(result.refreshResult.refreshedCount, 2);
    assert.equal(result.refreshResult.resetCount, 3);
    assert.deepEqual(result.refreshResult.missingSequences, ['Deleted Sequence']);
    assert.equal(result.filters.length, 3);
    assert.equal(result.filters[0].sequenceName, 'Renamed Sequence A');
    assert.equal(result.filters[0].videoTrackUsage[0].trackNumber, 3);
    assert.equal(result.filters[0].videoTrackUsage[0].clipCount, 5);
    assert.equal(result.filters[1].audioTrackUsage[0].trackNumber, 2);
    assert.equal(result.filters[1].audioTrackUsage[0].clipCount, 7);
    assert.equal(result.filters[2].videoTrackUsage[0].trackNumber, 9);
    result.filters.forEach((filter) => {
        assert.deepEqual(filter.ignoredVideoTracks, []);
        assert.deepEqual(filter.ignoredAudioTracks, []);
    });
});

test('Premiere host refresh plan reads every requested sequence, not only the active sequence', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'collector.jsx'), 'utf8');
    const hostContext = vm.createContext({
        JSON,
        app: {
            project: {
                activeSequence: null,
                sequences: null
            }
        }
    });

    vm.runInContext(`
        function createTestTracks(clipCounts, mediaPrefix) {
            const tracks = [];
            clipCounts.forEach((clipCount, trackIndex) => {
                const clips = [];
                for (let clipIndex = 0; clipIndex < clipCount; clipIndex += 1) {
                    const mediaPath = mediaPrefix + '-track-' + (trackIndex + 1) + '-clip-' + (clipIndex + 1) + '.mov';
                    clips.push({
                        projectItem: {
                            getMediaPath: function getMediaPath() {
                                return mediaPath;
                            }
                        }
                    });
                }
                clips.numItems = clips.length;
                tracks.push({ clips });
            });
            tracks.numTracks = tracks.length;
            return tracks;
        }

        const sequenceA = {
            sequenceID: 'sequence-a',
            name: 'Sequence A',
            videoTracks: createTestTracks([2, 0], 'A-video'),
            audioTracks: createTestTracks([1], 'A-audio')
        };
        const sequenceB = {
            sequenceID: 'sequence-b',
            name: 'Sequence B',
            videoTracks: createTestTracks([3], 'B-video'),
            audioTracks: createTestTracks([0, 4], 'B-audio')
        };
        const sequences = [sequenceA, sequenceB];
        sequences.numSequences = sequences.length;
        app.project.sequences = sequences;
        app.project.activeSequence = sequenceA;
    `, hostContext);
    vm.runInContext(hostSource, hostContext, { filename: 'collector.jsx' });

    const raw = vm.runInContext(`getSequenceTrackUsagePlan(${JSON.stringify(JSON.stringify([
        { sequenceID: 'sequence-b', sequenceName: 'Old B Name' },
        { sequenceID: '', sequenceName: 'Sequence A' },
        { sequenceID: 'missing-sequence', sequenceName: 'Missing Sequence' }
    ]))})`, hostContext);
    const result = JSON.parse(raw);

    assert.equal(result.sequences.length, 2);
    assert.equal(result.sequences[0].sequenceID, 'sequence-b');
    assert.equal(result.sequences[0].videoTrackUsage[0].clipCount, 3);
    assert.equal(result.sequences[0].audioTrackUsage[1].clipCount, 4);
    assert.equal(result.sequences[1].sequenceID, 'sequence-a');
    assert.equal(result.sequences[1].videoTrackUsage[0].clipCount, 2);
    assert.deepEqual(result.missingSequences, ['Missing Sequence']);
});

test('Refresh Project preserves presets while refreshing all sequences', async () => {
    const testDocument = createTestDocument();
    const context = loadCollectorLogic(testDocument);

    const result = await vm.runInContext(`(async () => {
        let loadCount = 0;
        let allSequenceRefreshCount = 0;
        let receivedResetMode = null;
        loadProjectPlan = async function loadProjectPlanForRefreshTest() {
            loadCount += 1;
            return true;
        };
        refreshAllSelectedSequenceTracks = async function refreshAllForRefreshTest(resetChoices) {
            allSequenceRefreshCount += 1;
            receivedResetMode = resetChoices;
            return {
                refreshedCount: 2,
                missingSequences: [],
                resetCount: 2
            };
        };

        await refreshProject();
        return {
            loadCount,
            allSequenceRefreshCount,
            receivedResetMode,
            summary: document.getElementById('summaryText').textContent
        };
    })()`, context);

    assert.equal(result.loadCount, 1);
    assert.equal(result.allSequenceRefreshCount, 1);
    assert.equal(result.receivedResetMode, false);
    assert.match(result.summary, /all 2 selected sequences/i);
});

test('track presets persist per user and generate the next available default name', () => {
    const storage = createMemoryStorage();
    const context = loadCollectorLogic(null, { localStorage: storage });
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        loadTrackPresets();
        const filter = createSequenceFilter(
            'sequence-a',
            'Sequence A',
            [
                { trackNumber: 1, clipCount: 1, hasItems: true },
                { trackNumber: 2, clipCount: 1, hasItems: true }
            ],
            [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            true
        );
        filter.ignoredVideoTracks = [2];
        filter.ignoredAudioTracks = [1];

        const firstDefaultName = getNextTrackPresetName(trackPresets);
        const firstPreset = upsertTrackPresetFromFilter(filter, firstDefaultName);
        const secondDefaultName = getNextTrackPresetName(trackPresets);
        saveTrackPresets();

        trackPresets = [];
        trackPresetsLoaded = false;
        loadTrackPresets();
        const reloadedBeforeUpdate = JSON.parse(JSON.stringify(trackPresets));

        filter.ignoredVideoTracks = [1];
        filter.ignoredAudioTracks = [];
        const updatedPreset = upsertTrackPresetFromFilter(filter, 'project copy preset 1');

        return {
            firstDefaultName,
            secondDefaultName,
            firstPresetId: firstPreset.id,
            reloaded: reloadedBeforeUpdate,
            updatedPreset,
            finalCount: trackPresets.length
        };
    })())`, context));

    assert.equal(result.firstDefaultName, 'Project Copy Preset 1');
    assert.equal(result.secondDefaultName, 'Project Copy Preset 2');
    assert.equal(result.reloaded.length, 1);
    assert.equal(result.reloaded[0].id, result.firstPresetId);
    assert.deepEqual(result.reloaded[0].ignoredVideoTracks, [2]);
    assert.deepEqual(result.reloaded[0].ignoredAudioTracks, [1]);
    assert.equal(result.finalCount, 1);
    assert.equal(result.updatedPreset.id, result.firstPresetId);
    assert.deepEqual(result.updatedPreset.ignoredVideoTracks, [1]);
    assert.deepEqual(result.updatedPreset.ignoredAudioTracks, []);
});

test('manual cleanup preserves only track presets, destination, and unrelated application data', () => {
    const storage = createMemoryStorage({
        'projectcollector.trackPresets': '[{"id":"preset-1","name":"Keep Me"}]',
        'projectcollector.destination': 'D:/Backups',
        'projectcollector.sequenceOnlyMode': '1',
        'projectcollector.copyProjectFile': '1',
        'projectcollector.sequenceFilters:c:/project.prproj': '[{"sequenceName":"Old"}]',
        'another.extension.setting': 'keep'
    });
    const context = loadCollectorLogic(null, { localStorage: storage });
    const cleanupResult = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const deleted = deletePluginRelatedData();
        return {
            deleted,
            remainingProjectKeys: Array.from(
                { length: localStorage.length },
                (_, index) => localStorage.key(index)
            ).filter((key) => key && key.indexOf('projectcollector.') === 0).sort()
        };
    })())`, context));

    assert.equal(cleanupResult.deleted, true);
    assert.deepEqual(cleanupResult.remainingProjectKeys, [
        'projectcollector.destination',
        'projectcollector.trackPresets'
    ]);
    assert.equal(storage.getItem('projectcollector.trackPresets'), '[{"id":"preset-1","name":"Keep Me"}]');
    assert.equal(storage.getItem('projectcollector.destination'), 'D:/Backups');
    assert.equal(storage.getItem('another.extension.setting'), 'keep');
});

test('different sequences can apply different presets from the same shared list', () => {
    const context = loadCollectorLogic(createTestDocument());
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        trackPresetsLoaded = true;
        trackPresets = [
            {
                id: 'dialogue-preset',
                name: 'Dialogue',
                ignoredVideoTracks: [2, 9],
                ignoredAudioTracks: [1]
            },
            {
                id: 'music-preset',
                name: 'Music',
                ignoredVideoTracks: [1],
                ignoredAudioTracks: [2, 8]
            }
        ];
        const usage = {
            video: [
                { trackNumber: 1, clipCount: 1, hasItems: true },
                { trackNumber: 2, clipCount: 1, hasItems: true }
            ],
            audio: [
                { trackNumber: 1, clipCount: 1, hasItems: true },
                { trackNumber: 2, clipCount: 1, hasItems: true }
            ]
        };
        selectedSequenceFilters = [
            createSequenceFilter('sequence-a', 'Sequence A', usage.video, usage.audio, true),
            createSequenceFilter('sequence-b', 'Sequence B', usage.video, usage.audio, false)
        ];
        renderSequenceFilters = function renderSequenceFiltersForPresetTest() {};

        applyTrackPresetToSequence('sequence-a', 'dialogue-preset');
        applyTrackPresetToSequence('sequence-b', 'music-preset');

        const firstSection = renderSequencePresetSection(selectedSequenceFilters[0]);
        const secondSection = renderSequencePresetSection(selectedSequenceFilters[1]);
        const firstSelect = firstSection.children.find((control) => String(control.className || '').indexOf('sequence-preset-select') !== -1);
        const secondSelect = secondSection.children.find((control) => String(control.className || '').indexOf('sequence-preset-select') !== -1);
        return {
            filters: selectedSequenceFilters,
            sharedFirstList: firstSelect.children.map((option) => option.textContent),
            sharedSecondList: secondSelect.children.map((option) => option.textContent)
        };
    })())`, context));

    assert.equal(result.filters[0].selectedPresetId, 'dialogue-preset');
    assert.deepEqual(result.filters[0].ignoredVideoTracks, [2]);
    assert.deepEqual(result.filters[0].ignoredAudioTracks, [1]);
    assert.equal(result.filters[1].selectedPresetId, 'music-preset');
    assert.deepEqual(result.filters[1].ignoredVideoTracks, [1]);
    assert.deepEqual(result.filters[1].ignoredAudioTracks, [2]);
    assert.deepEqual(result.sharedFirstList, ['Custom tracks', 'Use Premiere track locks (default)', 'Auto: exclude backup tracks and higher', 'Dialogue', 'Music']);
    assert.deepEqual(result.sharedSecondList, result.sharedFirstList);
});

test('Save preset prompts with an automatic name and immediately selects the saved preset', () => {
    const storage = createMemoryStorage();
    let receivedDefaultName = '';
    const context = loadCollectorLogic(createTestDocument(), {
        localStorage: storage,
        prompt(message, defaultName) {
            receivedDefaultName = defaultName;
            return 'Interview Tracks';
        }
    });

    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        trackPresetsLoaded = true;
        trackPresets = [];
        selectedSequenceFilters = [
            createSequenceFilter(
                'sequence-a',
                'Sequence A',
                [
                    { trackNumber: 1, clipCount: 1, hasItems: true },
                    { trackNumber: 2, clipCount: 1, hasItems: true }
                ],
                [{ trackNumber: 1, clipCount: 1, hasItems: true }],
                true
            )
        ];
        selectedSequenceFilters[0].ignoredVideoTracks = [2];
        selectedSequenceFilters[0].ignoredAudioTracks = [1];
        renderSequenceFilters = function renderSequenceFiltersForSavePresetTest() {};
        updateSelectionSummary = function updateSelectionSummaryForSavePresetTest() {};

        saveTrackPresetForSequence('sequence-a');
        return {
            presets: trackPresets,
            filter: selectedSequenceFilters[0],
            storedPresets: JSON.parse(localStorage.getItem(TRACK_PRESETS_STORAGE_KEY))
        };
    })())`, context));

    assert.equal(receivedDefaultName, 'Project Copy Preset 1');
    assert.equal(result.presets.length, 1);
    assert.equal(result.presets[0].name, 'Interview Tracks');
    assert.deepEqual(result.presets[0].ignoredVideoTracks, [2]);
    assert.deepEqual(result.presets[0].ignoredAudioTracks, [1]);
    assert.equal(result.filter.selectedPresetId, result.presets[0].id);
    assert.deepEqual(result.storedPresets, result.presets);

    const reloadContext = loadCollectorLogic(null, { localStorage: storage });
    const reloaded = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        latestPlan = {
            activeSequenceID: 'sequence-a',
            activeSequenceName: 'Sequence A',
            videoTrackUsage: [
                { trackNumber: 1, clipCount: 1, hasItems: true },
                { trackNumber: 2, clipCount: 1, hasItems: true }
            ],
            audioTrackUsage: [{ trackNumber: 1, clipCount: 1, hasItems: true }]
        };
        renderSequenceFilters = function renderSequenceFiltersForReloadPresetTest() {};
        loadSequenceFilters();
        return {
            presets: trackPresets,
            filter: selectedSequenceFilters[0]
        };
    })())`, reloadContext));

    assert.equal(reloaded.presets.length, 1);
    assert.equal(reloaded.filter.selectedPresetId, 'builtin-track-locks');
    assert.deepEqual(reloaded.filter.ignoredVideoTracks, []);
    assert.deepEqual(reloaded.filter.ignoredAudioTracks, []);
});

test('manual track edits and resets detach only that sequence from its preset', () => {
    const context = loadCollectorLogic();
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        trackPresetsLoaded = true;
        trackPresets = [{
            id: 'shared-preset',
            name: 'Shared',
            ignoredVideoTracks: [2],
            ignoredAudioTracks: []
        }];
        selectedSequenceFilters = [
            createSequenceFilter(
                'sequence-a',
                'Sequence A',
                [
                    { trackNumber: 1, clipCount: 1, hasItems: true },
                    { trackNumber: 2, clipCount: 1, hasItems: true }
                ],
                [],
                true
            ),
            createSequenceFilter(
                'sequence-b',
                'Sequence B',
                [
                    { trackNumber: 1, clipCount: 1, hasItems: true },
                    { trackNumber: 2, clipCount: 1, hasItems: true }
                ],
                [],
                false
            )
        ];
        selectedSequenceFilters.forEach((filter) => applyTrackPresetToFilter(filter, trackPresets[0]));
        renderSequenceFilters = function renderSequenceFiltersForDetachTest() {};

        toggleIgnoredTrack('sequence-a', 'video', 1);
        const afterManualEdit = selectedSequenceFilters.map((filter) => ({
            selectedPresetId: filter.selectedPresetId,
            ignoredVideoTracks: filter.ignoredVideoTracks.slice()
        }));
        resetTrackSelection('sequence-b');

        return {
            afterManualEdit,
            finalFilters: selectedSequenceFilters
        };
    })())`, context));

    assert.equal(result.afterManualEdit[0].selectedPresetId, '');
    assert.deepEqual(result.afterManualEdit[0].ignoredVideoTracks, [1, 2]);
    assert.equal(result.afterManualEdit[1].selectedPresetId, 'shared-preset');
    assert.deepEqual(result.afterManualEdit[1].ignoredVideoTracks, [2]);
    assert.equal(result.finalFilters[1].selectedPresetId, '');
    assert.deepEqual(result.finalFilters[1].ignoredVideoTracks, []);
    assert.deepEqual(result.finalFilters[0].ignoredVideoTracks, [1, 2]);
});

test('every sequence is removable and an intentionally empty list stays saved', () => {
    const storage = createMemoryStorage();
    const firstContext = loadCollectorLogic(null, { localStorage: storage });
    const firstResult = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        latestPlan = {
            activeSequenceID: 'sequence-a',
            activeSequenceName: 'Sequence A',
            videoTrackUsage: [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            audioTrackUsage: []
        };
        renderSequenceFilters = function renderSequenceFiltersForRemovalTest() {};
        loadSequenceFilters();
        const initialCount = selectedSequenceFilters.length;
        removeSequenceFilter('sequence-a');
        loadSequenceFilters();
        return {
            initialCount,
            remainingCount: selectedSequenceFilters.length,
            storedFilters: JSON.parse(localStorage.getItem(getSequenceFiltersStorageKey()))
        };
    })())`, firstContext));

    assert.equal(firstResult.initialCount, 1);
    assert.equal(firstResult.remainingCount, 0);
    assert.deepEqual(firstResult.storedFilters, []);

    const reloadContext = loadCollectorLogic(null, { localStorage: storage });
    const reloadCount = vm.runInContext(`(() => {
        latestPlan = {
            activeSequenceID: 'sequence-a',
            activeSequenceName: 'Sequence A',
            videoTrackUsage: [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            audioTrackUsage: []
        };
        renderSequenceFilters = function renderSequenceFiltersForEmptyReloadTest() {};
        loadSequenceFilters();
        return selectedSequenceFilters.length;
    })()`, reloadContext);
    assert.equal(reloadCount, 0);

    const testDocument = createTestDocument();
    const renderContext = loadCollectorLogic(testDocument);
    const actionLabels = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        selectedSequenceFilters = [
            createSequenceFilter(
                'sequence-a',
                'Sequence A',
                [{ trackNumber: 1, clipCount: 1, hasItems: true }],
                [],
                true
            )
        ];
        updateSelectionSummary = function updateSelectionSummaryForRemoveButtonTest() {};
        renderSequenceFilters();
        const container = document.getElementById('sequenceFilters');
        const tab = container.children[0].children[0];
        const card = container.children[1];
        return {
            header: [tab.children[1].textContent],
            sections: card.children.map((section) => section.className)
        };
    })())`, renderContext));
    assert.equal(actionLabels.header.length, 1);
    assert.equal(actionLabels.header[0], '×');
    assert.deepEqual(actionLabels.sections, ['sequence-groups']);
});

test('sequence selections are saved separately for each Premiere project', () => {
    const storage = createMemoryStorage();
    const context = loadCollectorLogic(null, { localStorage: storage });
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        renderSequenceFilters = function renderSequenceFiltersForProjectScopeTest() {};
        latestPlan = {
            projectName: 'Project A',
            projectPath: 'C:/Projects/Project A.prproj',
            availableSequences: [{
                sequenceID: 'sequence-a',
                sequenceName: 'Sequence A'
            }],
            activeSequenceID: 'sequence-a',
            activeSequenceName: 'Sequence A',
            videoTrackUsage: [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            audioTrackUsage: []
        };
        loadSequenceFilters();
        selectedSequenceFilters[0].ignoredVideoTracks = [1];
        selectedSequenceFilters[0].selectedPresetId = '';
        selectedSequenceFilters.push(createSequenceFilter(
            'deleted-sequence',
            'Deleted Sequence',
            [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            [],
            false
        ));
        saveSequenceFilters();
        loadSequenceFilters();
        const sameProjectFilters = selectedSequenceFilters.map((filter) => ({
            sequenceName: filter.sequenceName,
            ignoredVideoTracks: filter.ignoredVideoTracks.slice()
        }));

        latestPlan = {
            projectName: 'Project B',
            projectPath: 'C:/Projects/Project B.prproj',
            availableSequences: [{
                sequenceID: 'sequence-b',
                sequenceName: 'Sequence B'
            }],
            activeSequenceID: 'sequence-b',
            activeSequenceName: 'Sequence B',
            videoTrackUsage: [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            audioTrackUsage: []
        };
        loadSequenceFilters();
        const projectBFilters = selectedSequenceFilters.map((filter) => filter.sequenceName);

        latestPlan = {
            projectName: 'Project A',
            projectPath: 'C:/Projects/Project A.prproj',
            availableSequences: [{
                sequenceID: 'sequence-a',
                sequenceName: 'Sequence A'
            }],
            activeSequenceID: 'sequence-a',
            activeSequenceName: 'Sequence A',
            videoTrackUsage: [{ trackNumber: 1, clipCount: 1, hasItems: true }],
            audioTrackUsage: []
        };
        loadSequenceFilters();
        const restoredProjectAFilters = selectedSequenceFilters.map((filter) => ({
            sequenceName: filter.sequenceName,
            ignoredVideoTracks: filter.ignoredVideoTracks.slice()
        }));

        return {
            sameProjectFilters,
            projectBFilters,
            restoredProjectAFilters
        };
    })())`, context));

    assert.deepEqual(result.sameProjectFilters.map((filter) => filter.sequenceName), ['Sequence A']);
    assert.deepEqual(result.sameProjectFilters[0].ignoredVideoTracks, []);
    assert.deepEqual(result.projectBFilters, ['Sequence B']);
    assert.deepEqual(result.restoredProjectAFilters.map((filter) => filter.sequenceName), ['Sequence A']);
    assert.deepEqual(result.restoredProjectAFilters[0].ignoredVideoTracks, []);
});

test('Premiere project plan reports every sequence available in the open project', () => {
    const context = loadHostRelinkLogic();
    const sequenceA = {
        name: 'Sequence A',
        sequenceID: 'sequence-a',
        videoTracks: { numTracks: 0 },
        audioTracks: { numTracks: 0 }
    };
    const sequenceB = {
        name: 'Sequence B',
        sequenceID: 'sequence-b',
        videoTracks: { numTracks: 0 },
        audioTracks: { numTracks: 0 }
    };
    const sequences = {
        0: sequenceA,
        1: sequenceB,
        numSequences: 2
    };
    context.app = {
        project: {
            name: 'Scoped Project.prproj',
            path: 'C:/Projects/Scoped Project.prproj',
            rootItem: {
                type: 2,
                children: createHostChildren([])
            },
            sequences,
            activeSequence: sequenceB
        }
    };

    const plan = JSON.parse(context.getProjectCopyPlan('D:/Backup'));

    assert.deepEqual(
        Array.from(plan.availableSequences, (sequence) => ({
            sequenceID: sequence.sequenceID,
            sequenceName: sequence.sequenceName
        })),
        [
            { sequenceID: 'sequence-a', sequenceName: 'Sequence A' },
            { sequenceID: 'sequence-b', sequenceName: 'Sequence B' }
        ]
    );
    assert.equal(plan.activeSequenceID, 'sequence-b');
});

test('saving duplicate track settings selects the existing preset and does not prompt or duplicate', () => {
    let alertMessage = '';
    let promptCount = 0;
    const context = loadCollectorLogic(createTestDocument(), {
        parent: {showReadablePrompt(options) {
            alertMessage = options.message;
        }},
        prompt() {
            promptCount += 1;
            return 'Should Not Save';
        }
    });

    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        trackPresetsLoaded = true;
        trackPresets = [{
            id: 'existing-preset',
            name: 'Existing Preset',
            ignoredVideoTracks: [2],
            ignoredAudioTracks: [1]
        }];
        selectedSequenceFilters = [
            createSequenceFilter(
                'sequence-a',
                'Sequence A',
                [
                    { trackNumber: 1, clipCount: 1, hasItems: true },
                    { trackNumber: 2, clipCount: 1, hasItems: true }
                ],
                [{ trackNumber: 1, clipCount: 1, hasItems: true }],
                false
            )
        ];
        selectedSequenceFilters[0].ignoredVideoTracks = [2];
        selectedSequenceFilters[0].ignoredAudioTracks = [1];
        renderSequenceFilters = function renderSequenceFiltersForDuplicatePresetTest() {};
        updateSelectionSummary = function updateSelectionSummaryForDuplicatePresetTest() {};

        saveTrackPresetForSequence('sequence-a');
        return {
            presets: trackPresets,
            filter: selectedSequenceFilters[0]
        };
    })())`, context));

    assert.equal(promptCount, 0);
    assert.match(alertMessage, /already exists/i);
    assert.match(alertMessage, /Existing Preset/);
    assert.equal(result.presets.length, 1);
    assert.equal(result.filter.selectedPresetId, 'existing-preset');
});

test('deleting a preset removes it from the shared list but preserves current track choices', () => {
    const storage = createMemoryStorage();
    const context = loadCollectorLogic(createTestDocument(), { localStorage: storage });
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        trackPresetsLoaded = true;
        trackPresets = [
            {
                id: 'delete-me',
                name: 'Delete Me',
                ignoredVideoTracks: [2],
                ignoredAudioTracks: [1]
            },
            {
                id: 'keep-me',
                name: 'Keep Me',
                ignoredVideoTracks: [1],
                ignoredAudioTracks: []
            }
        ];
        selectedSequenceFilters = [
            createSequenceFilter(
                'sequence-a',
                'Sequence A',
                [
                    { trackNumber: 1, clipCount: 1, hasItems: true },
                    { trackNumber: 2, clipCount: 1, hasItems: true }
                ],
                [{ trackNumber: 1, clipCount: 1, hasItems: true }],
                false
            ),
            createSequenceFilter(
                'sequence-b',
                'Sequence B',
                [
                    { trackNumber: 1, clipCount: 1, hasItems: true },
                    { trackNumber: 2, clipCount: 1, hasItems: true }
                ],
                [{ trackNumber: 1, clipCount: 1, hasItems: true }],
                false
            )
        ];
        selectedSequenceFilters.forEach((filter) => applyTrackPresetToFilter(filter, trackPresets[0]));
        const sectionBeforeDelete = renderSequencePresetSection(selectedSequenceFilters[0]);
        const deleteButtonBefore = sectionBeforeDelete.children.find(
            (control) => String(control.className || '').indexOf('sequence-preset-delete') !== -1
        );
        const deleted = deleteTrackPreset('delete-me');

        return {
            deleted,
            deleteButtonLabel: deleteButtonBefore.textContent,
            deleteButtonDisabled: deleteButtonBefore.disabled,
            presets: trackPresets,
            filters: selectedSequenceFilters,
            storedPresets: JSON.parse(localStorage.getItem(TRACK_PRESETS_STORAGE_KEY))
        };
    })())`, context));

    assert.equal(result.deleteButtonLabel, 'Delete');
    assert.equal(result.deleteButtonDisabled, false);
    assert.equal(result.deleted, true);
    assert.deepEqual(result.presets.map((preset) => preset.id), ['keep-me']);
    assert.deepEqual(result.storedPresets.map((preset) => preset.id), ['keep-me']);
    result.filters.forEach((filter) => {
        assert.equal(filter.selectedPresetId, '');
        assert.deepEqual(filter.ignoredVideoTracks, [2]);
        assert.deepEqual(filter.ignoredAudioTracks, [1]);
    });
});

test('BACKUP relink plan targets collected, verified skip-location, and original paths correctly', () => {
    const context = loadCollectorLogic();
    const result = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        const allTasks = [
            { source: 'C:/Original/copied.mov' },
            { source: 'C:/Original/ignored.mov' },
            { source: 'C:/Original/compare-skipped.mov' },
            { source: 'C:/Original/copy-failed.mov' }
        ];
        const copiedTasks = [{
            source: 'C:/Original/copied.mov',
            destinationPath: 'D:/Backup/CollectedMedias/copied.mov'
        }];
        const compareMatches = [{
            task: allTasks[2],
            match: {
                path: 'E:/Skip Library/compare-skipped.mov'
            }
        }];
        return buildLinkProjectTasks(allTasks, copiedTasks, compareMatches);
    })())`, context));

    assert.deepEqual(result, [
        {
            source: 'C:/Original/copied.mov',
            destination: 'D:/Backup/CollectedMedias/copied.mov',
            targetKind: 'collected'
        },
        {
            source: 'C:/Original/ignored.mov',
            destination: 'C:/Original/ignored.mov',
            targetKind: 'original'
        },
        {
            source: 'C:/Original/compare-skipped.mov',
            destination: 'E:/Skip Library/compare-skipped.mov',
            targetKind: 'skip-location'
        },
        {
            source: 'C:/Original/copy-failed.mov',
            destination: 'C:/Original/copy-failed.mov',
            targetKind: 'original'
        }
    ]);
});

test('skip-location matching requires identical SHA-256 content, not only name and size', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-skip-hash-'));
    const sourceDir = path.join(tempRoot, 'source');
    const wrongDir = path.join(tempRoot, 'skip-wrong');
    const correctDir = path.join(tempRoot, 'skip-correct');
    fs.mkdirSync(sourceDir, { recursive: true });
    fs.mkdirSync(wrongDir, { recursive: true });
    fs.mkdirSync(correctDir, { recursive: true });
    const sourcePath = path.join(sourceDir, 'clip.mov');
    const wrongPath = path.join(wrongDir, 'clip.mov');
    const correctPath = path.join(correctDir, 'clip.mov');
    fs.writeFileSync(sourcePath, 'AAAA-same-size');
    fs.writeFileSync(wrongPath, 'BBBB-same-size');
    fs.writeFileSync(correctPath, 'AAAA-same-size');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic();
    const wrongOnlyLookup = context.buildCompareLookup([{
        name: 'clip.mov',
        path: wrongPath,
        size: fs.statSync(wrongPath).size
    }]);
    const wrongMatch = await context.findCompareMatchForTask({ source: sourcePath }, wrongOnlyLookup);
    assert.equal(wrongMatch, null);

    const mixedLookup = context.buildCompareLookup([
        {
            name: 'clip.mov',
            path: wrongPath,
            size: fs.statSync(wrongPath).size
        },
        {
            name: 'clip.mov',
            path: correctPath,
            size: fs.statSync(correctPath).size
        }
    ]);
    const verifiedMatch = await context.findCompareMatchForTask({ source: sourcePath }, mixedLookup);

    assert.equal(verifiedMatch.path, correctPath);
    assert.match(verifiedMatch.sha256, /^[a-f0-9]{64}$/);
});

test('copied project keeps the original project filename and appends BACKUP', () => {
    const context = loadCollectorLogic();
    const names = JSON.parse(vm.runInContext(`JSON.stringify((() => {
        sequenceOnlyMode = true;
        selectedSequenceFilters = [{
            sequenceID: 'different-sequence',
            sequenceName: 'This Sequence Name Must Not Be Used'
        }];
        latestPlan = {
            projectName: 'Different Collection Folder'
        };
        return {
            standard: getCollectedProjectFileName('D:/Projects/TEASER FN 20260721.prproj'),
            mixedCaseExtension: getCollectedProjectFileName('D:/Projects/Evening Show.PRPROJ')
        };
    })())`, context));

    assert.equal(names.standard, 'TEASER FN 20260721 BACKUP.prproj');
    assert.equal(names.mixedCaseExtension, 'Evening Show BACKUP.PRPROJ');
});

test('project copy physically creates the requested BACKUP filename even beside the source project', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-project-name-'));
    const sourceProjectPath = path.join(tempRoot, 'temp.prproj');
    const expectedBackupPath = path.join(tempRoot, 'temp BACKUP.prproj');
    fs.writeFileSync(sourceProjectPath, 'premiere-project-data');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic();
    const result = await context.copyProjectFileIntoCollectedRoot(
        tempRoot,
        sourceProjectPath,
        context.getCollectedProjectFileName(sourceProjectPath)
    );

    assert.equal(result.success, true);
    assert.equal(result.destinationPath, expectedBackupPath);
    assert.equal(fs.existsSync(sourceProjectPath), true);
    assert.equal(fs.existsSync(expectedBackupPath), true);
    assert.equal(fs.readFileSync(expectedBackupPath, 'utf8'), 'premiere-project-data');
    assert.deepEqual(
        fs.readdirSync(tempRoot).filter((name) => name.startsWith('.projectcollector-copy-')),
        []
    );
});

test('backup link map keeps copied media portable when the whole backup folder moves', (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-portable-link-map-'));
    const originalBackupRoot = path.join(tempRoot, 'original-location', 'Show_Project');
    const movedBackupRoot = path.join(tempRoot, 'moved-location', 'Show_Project');
    const originalProjectPath = path.join(tempRoot, 'source', 'Show Project.prproj');
    const originalBackupProjectPath = path.join(originalBackupRoot, 'Show Project BACKUP.prproj');
    const originalCollectedPath = path.join(originalBackupRoot, 'CollectedMedias', 'clip.mov');
    const untouchedOriginalMediaPath = path.join(tempRoot, 'source', 'untouched.mov');
    const skipLocationPath = path.join(tempRoot, 'skip-location', 'verified.mov');
    fs.mkdirSync(path.dirname(originalCollectedPath), { recursive: true });
    fs.mkdirSync(path.dirname(originalProjectPath), { recursive: true });
    fs.writeFileSync(originalBackupProjectPath, 'backup-project');
    fs.writeFileSync(originalCollectedPath, 'copied-media');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic();
    const manifest = context.createBackupLinkManifest(
        originalBackupRoot,
        originalBackupProjectPath,
        originalProjectPath,
        [
            {
                source: path.join(tempRoot, 'source', 'clip.mov'),
                destination: originalCollectedPath,
                targetKind: 'collected'
            },
            {
                source: untouchedOriginalMediaPath,
                destination: untouchedOriginalMediaPath,
                targetKind: 'original'
            },
            {
                source: path.join(tempRoot, 'source', 'verified.mov'),
                destination: skipLocationPath,
                targetKind: 'skip-location'
            }
        ]
    );
    context.writeBackupLinkManifest(originalBackupProjectPath, manifest);

    fs.mkdirSync(path.dirname(movedBackupRoot), { recursive: true });
    fs.renameSync(originalBackupRoot, movedBackupRoot);
    const movedBackupProjectPath = path.join(movedBackupRoot, 'Show Project BACKUP.prproj');
    const savedManifest = context.readBackupLinkManifest(movedBackupProjectPath);
    const resolvedTasks = JSON.parse(JSON.stringify(
        context.resolveBackupLinkManifestTasks(savedManifest, movedBackupProjectPath)
    ));

    assert.equal(
        savedManifest.relinkTasks[0].destination,
        'CollectedMedias/clip.mov'
    );
    assert.deepEqual(resolvedTasks, [
        {
            source: path.join(tempRoot, 'source', 'clip.mov'),
            destination: path.join(movedBackupRoot, 'CollectedMedias', 'clip.mov'),
            targetKind: 'collected'
        },
        {
            source: untouchedOriginalMediaPath,
            destination: untouchedOriginalMediaPath,
            targetKind: 'original'
        },
        {
            source: path.join(tempRoot, 'source', 'verified.mov'),
            destination: skipLocationPath,
            targetKind: 'skip-location'
        }
    ]);
});

test('standalone existing-backup action links without copying media again', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-later-link-'));
    const backupRoot = path.join(tempRoot, 'Show_Project');
    const originalProjectPath = path.join(tempRoot, 'source', 'Show Project.prproj');
    const backupProjectPath = path.join(backupRoot, 'Show Project BACKUP.prproj');
    const copiedMediaPath = path.join(backupRoot, 'CollectedMedias', 'clip.mov');
    fs.mkdirSync(path.dirname(copiedMediaPath), { recursive: true });
    fs.mkdirSync(path.dirname(originalProjectPath), { recursive: true });
    fs.writeFileSync(backupProjectPath, 'backup-project');
    fs.writeFileSync(copiedMediaPath, 'copied-media');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const testDocument = createTestDocument();
    const context = loadCollectorLogic(testDocument);
    const manifest = context.createBackupLinkManifest(
        backupRoot,
        backupProjectPath,
        originalProjectPath,
        [{
            source: path.join(tempRoot, 'source', 'clip.mov'),
            destination: copiedMediaPath,
            targetKind: 'collected'
        }]
    );
    context.writeBackupLinkManifest(backupProjectPath, manifest);
    context.ensureHostScriptLoaded = async () => true;

    let receivedRelinkTasks = null;
    context.callHost = async (script) => {
        if (script === 'saveCurrentProjectAndGetPath()') {
            return JSON.stringify({ projectPath: originalProjectPath });
        }
        if (script.startsWith('linkProjectCopyToCollectedMedia(')) {
            const match = script.match(/^linkProjectCopyToCollectedMedia\("(?:\\.|[^"])+"\,\"((?:\\.|[^"])*)"\)$/);
            assert.ok(match, `Unexpected relink host call: ${script}`);
            receivedRelinkTasks = JSON.parse(JSON.parse(`"${match[1]}"`));
            return JSON.stringify({
                success: true,
                linkedCollectedCount: 1,
                linkedSkipLocationCount: 0,
                linkedOriginalCount: 0,
                failed: []
            });
        }
        throw new Error(`Unexpected host call: ${script}`);
    };

    await context.linkExistingBackupProject(backupProjectPath);

    assert.deepEqual(receivedRelinkTasks, [{
        source: path.join(tempRoot, 'source', 'clip.mov'),
        destination: copiedMediaPath,
        targetKind: 'collected'
    }]);
    assert.equal(testDocument.getElementById('completionPrompt').classList.contains('is-success'), true);
    assert.equal(testDocument.getElementById('completionTitle').textContent, 'Done without error');
    assert.match(testDocument.getElementById('summaryText').textContent, /Existing BACKUP linked without errors/);
});

test('collection sends the complete copied-versus-ignored relink plan to Premiere', async (t) => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'project-collector-relink-plan-'));
    const sourceDir = path.join(tempRoot, 'source');
    const destinationDir = path.join(tempRoot, 'backup');
    const skipLocationDir = path.join(tempRoot, 'skip-location', 'extra', 'nested');
    const copiedSourcePath = path.join(sourceDir, 'copied.mov');
    const ignoredSourcePath = path.join(sourceDir, 'ignored.mov');
    const skipSourcePath = path.join(sourceDir, 'skip.mov');
    const skipMatchPath = path.join(skipLocationDir, 'skip.mov');
    const copiedDestinationPath = path.join(destinationDir, 'Relink_Project', 'CollectedMedias', 'copied.mov');
    const projectCopyPath = path.join(destinationDir, 'Relink_Project', 'Relink Project BACKUP.prproj');
    fs.mkdirSync(sourceDir, { recursive: true });
    fs.mkdirSync(skipLocationDir, { recursive: true });
    fs.writeFileSync(copiedSourcePath, 'copied-media');
    fs.writeFileSync(ignoredSourcePath, 'ignored-media');
    fs.writeFileSync(skipSourcePath, 'verified-skip-media');
    fs.writeFileSync(skipMatchPath, 'verified-skip-media');
    t.after(() => fs.rmSync(tempRoot, { recursive: true, force: true }));

    const context = loadCollectorLogic(createTestDocument());
    vm.runInContext(`
        destination = ${JSON.stringify(destinationDir)};
        compareLocation = ${JSON.stringify(path.join(tempRoot, 'skip-location'))};
        copyProjectFile = true;
        linkProjectAfterCollection = true;
        sequenceOnlyMode = false;
        createReducedProject = false;
        latestPlan = {
            projectName: 'Relink_Project',
            projectPath: 'C:/Projects/Relink Project.prproj',
            missingMedia: [],
            tasks: [{
                name: 'copied.mov',
                source: ${JSON.stringify(copiedSourcePath)},
                destination: 'CollectedMedias/copied.mov',
                binPath: 'CollectedMedias',
                relativePath: 'CollectedMedias/copied.mov'
            }, {
                name: 'ignored.mov',
                source: ${JSON.stringify(ignoredSourcePath)},
                destination: 'CollectedMedias/ignored.mov',
                binPath: 'CollectedMedias',
                relativePath: 'CollectedMedias/ignored.mov'
            }, {
                name: 'skip.mov',
                source: ${JSON.stringify(skipSourcePath)},
                destination: 'CollectedMedias/skip.mov',
                binPath: 'CollectedMedias',
                relativePath: 'CollectedMedias/skip.mov'
            }]
        };
        sourceTree = [];
        buildCopyReadyContext = async function buildCopyReadyContextForRelinkPlanTest() {
            const ignoredMediaSet = new Set([normalizeMediaKey(latestPlan.tasks[1].source)]);
            return {
                ok: true,
                copyRuleContext: createCopyRuleContext({
                    treeSelectedTaskSet: new Set(latestPlan.tasks),
                    trackRuleContext: buildTrackRuleContext(latestPlan.tasks, new Set(), ignoredMediaSet)
                }),
                copyWarnings: [],
                sequenceScopeInfo: null,
                trackConflicts: []
            };
        };
        copiedProjectDestinationRoot = '';
        copyProjectFileIntoCollectedRoot = async function copyProjectForRelinkPlanTest(rootPath) {
            copiedProjectDestinationRoot = rootPath;
            return {
                success: true,
                destinationPath: ${JSON.stringify(projectCopyPath)}
            };
        };
    `, context);

    let receivedRelinkTasks = null;
    context.callHost = async (script) => {
        if (script === 'saveCurrentProjectAndGetPath()') {
            return JSON.stringify({ projectPath: 'C:/Projects/Relink Project.prproj' });
        }
        if (script.startsWith('linkProjectCopyToCollectedMedia(')) {
            const match = script.match(/^linkProjectCopyToCollectedMedia\("(?:\\.|[^"])+"\,\"((?:\\.|[^"])*)"\)$/);
            assert.ok(match, `Unexpected relink host call: ${script}`);
            receivedRelinkTasks = JSON.parse(JSON.parse(`"${match[1]}"`));
            return JSON.stringify({
                success: true,
                offlineCount: 3,
                linkedCount: 3,
                linkedCollectedCount: 1,
                linkedSkipLocationCount: 1,
                linkedOriginalCount: 1,
                skippedNoPathCount: 0,
                failed: []
            });
        }
        throw new Error(`Unexpected host call: ${script}`);
    };

    await vm.runInContext('collect()', context);

    assert.equal(fs.existsSync(copiedDestinationPath), true);
    assert.equal(fs.existsSync(path.join(destinationDir, 'Relink_Project', 'CollectedMedias', 'ignored.mov')), false);
    assert.equal(fs.existsSync(path.join(destinationDir, 'Relink_Project', 'CollectedMedias', 'skip.mov')), false);
    assert.equal(
        vm.runInContext('copiedProjectDestinationRoot', context),
        path.join(destinationDir, 'Relink_Project')
    );
    assert.deepEqual(receivedRelinkTasks, [
        {
            source: copiedSourcePath,
            destination: copiedDestinationPath,
            targetKind: 'collected'
        },
        {
            source: ignoredSourcePath,
            destination: ignoredSourcePath,
            targetKind: 'original'
        },
        {
            source: skipSourcePath,
            destination: skipMatchPath,
            targetKind: 'skip-location'
        }
    ]);

    const savedLinkMap = JSON.parse(fs.readFileSync(
        path.join(destinationDir, 'Relink_Project', 'Project Collector Link Map.json'),
        'utf8'
    ));
    assert.equal(savedLinkMap.backupProjectFileName, 'Relink Project BACKUP.prproj');
    assert.deepEqual(savedLinkMap.relinkTasks.map((task) => ({
        destination: task.destination,
        destinationIsRelative: task.destinationIsRelative,
        targetKind: task.targetKind
    })), [
        {
            destination: 'CollectedMedias/copied.mov',
            destinationIsRelative: true,
            targetKind: 'collected'
        },
        {
            destination: ignoredSourcePath,
            destinationIsRelative: false,
            targetKind: 'original'
        },
        {
            destination: skipMatchPath,
            destinationIsRelative: false,
            targetKind: 'skip-location'
        }
    ]);
});

function loadHostRelinkLogic() {
    const scriptPath = path.join(__dirname, '..', 'jsx', 'collector.jsx');
    const source = fs.readFileSync(scriptPath, 'utf8');
    const context = vm.createContext({
        File: function File(filePath) {
            this.fsName = filePath;
            this.exists = true;
        },
        JSON,
        ProjectItemType: { BIN: 2 },
        String
    });
    vm.runInContext(source, context);
    return context;
}

function createHostChildren(items) {
    const children = { numItems: items.length };
    items.forEach((item, index) => {
        children[index] = item;
    });
    return children;
}

function createHostMediaItem(name, mediaPath, events, options) {
    const settings = options || {};
    let currentPath = mediaPath;
    let offline = false;

    return {
        name,
        type: 1,
        getMediaPath() {
            return currentPath;
        },
        isOffline() {
            return offline;
        },
        setOffline() {
            events.push(`offline:${name}`);
            if (settings.offlineFails) {
                return false;
            }
            offline = true;
            currentPath = '';
            return true;
        },
        changeMediaPath(destination) {
            events.push(`link:${name}:${destination}`);
            if (settings.linkFails) {
                return 1;
            }
            currentPath = destination;
            offline = false;
            return 0;
        }
    };
}

test('Premiere host plan ignores pathless project structures but reports genuine offline media', () => {
    const context = loadHostRelinkLogic();
    let sequenceMediaPathRead = false;
    const sequenceItem = {
        name: 'Baby wakes mom up in the sweetest way possible@Daily Mail - nest',
        type: 1,
        isSequence() {
            return true;
        },
        getMediaPath() {
            sequenceMediaPathRead = true;
            return '';
        }
    };
    const generatedItem = {
        name: 'Adjustment Layer',
        type: 1,
        isSequence() {
            return false;
        },
        isOffline() {
            return false;
        },
        getMediaPath() {
            return '';
        }
    };
    const offlineItem = {
        name: 'Offline Interview',
        type: 1,
        isSequence() {
            return false;
        },
        isOffline() {
            return true;
        },
        getMediaPath() {
            return '';
        }
    };
    const unreadableItem = {
        name: 'Unreadable Host Item',
        type: 1,
        isSequence() {
            return false;
        },
        isOffline() {
            return false;
        },
        getMediaPath() {
            throw new Error('Host media-path read failed');
        }
    };
    const networkItem = {
        name: 'Network Clip',
        type: 1,
        isSequence() {
            return false;
        },
        isOffline() {
            return false;
        },
        getMediaPath() {
            return '\\\\media-server\\news\\network-clip.mov';
        }
    };

    context.app = {
        project: {
            name: 'Portable News Project.prproj',
            path: 'C:/Projects/Portable News Project.prproj',
            rootItem: {
                type: 2,
                children: createHostChildren([
                    sequenceItem,
                    generatedItem,
                    offlineItem,
                    unreadableItem,
                    networkItem
                ])
            }
        }
    };

    const plan = JSON.parse(context.getProjectCopyPlan('\\\\backup-server\\projects'));

    assert.equal(sequenceMediaPathRead, false);
    assert.equal(plan.tasks.length, 1);
    assert.equal(plan.tasks[0].source, '\\\\media-server\\news\\network-clip.mov');
    assert.equal(plan.missingMedia.length, 2);
    assert.equal(
        plan.missingMedia[0],
        'Offline Interview | Media is offline and no media path is available'
    );
    assert.match(
        plan.missingMedia[1],
        /^Unreadable Host Item \| Could not read media path: Error: Host media-path read failed$/
    );
});

test('Premiere host relink takes every file offline before linking copied and ignored media', () => {
    const context = loadHostRelinkLogic();
    const events = [];
    const copiedItem = createHostMediaItem('Copied', 'C:/Original/copied.mov', events);
    const ignoredItem = createHostMediaItem('Ignored', 'C:/Original/ignored.mov', events);
    const skipItem = createHostMediaItem('Skip', 'C:/Original/skip.mov', events);
    const generatedItem = {
        name: 'Generated',
        type: 1,
        getMediaPath() {
            return '';
        }
    };
    const copiedProject = {
        path: 'D:/Backup/My Project BACKUP.prproj',
        rootItem: {
            type: 2,
            children: createHostChildren([copiedItem, ignoredItem, skipItem, generatedItem])
        },
        save() {
            events.push('save:backup');
            return 0;
        },
        closeDocument(saveFirst, promptIfDirty) {
            events.push(`close:backup:${saveFirst}:${promptIfDirty}`);
            return 0;
        }
    };
    const originalProject = {
        path: 'C:/Projects/My Project.prproj',
        rootItem: { type: 2, children: createHostChildren([]) },
        save() {
            events.push('save:original');
            return 0;
        },
        closeDocument(saveFirst, promptIfDirty) {
            events.push(`close:original:${saveFirst}:${promptIfDirty}`);
            app.project = null;
            return 0;
        }
    };
    const app = {
        project: originalProject,
        openDocument(projectPath) {
            events.push(`open:${projectPath}`);
            this.project = projectPath === copiedProject.path ? copiedProject : originalProject;
            return true;
        }
    };
    context.app = app;

    const tasks = JSON.stringify([
        {
            source: 'C:/Original/copied.mov',
            destination: 'D:/Backup/CollectedMedias/copied.mov',
            targetKind: 'collected'
        },
        {
            source: 'C:/Original/ignored.mov',
            destination: 'C:/Original/ignored.mov',
            targetKind: 'original'
        },
        {
            source: 'C:/Original/skip.mov',
            destination: 'E:/Skip Library/skip.mov',
            targetKind: 'skip-location'
        }
    ]);
    const result = JSON.parse(context.linkProjectCopyToCollectedMedia(copiedProject.path, tasks));

    assert.equal(result.success, true);
    assert.equal(result.offlineCount, 3);
    assert.equal(result.linkedCount, 3);
    assert.equal(result.linkedCollectedCount, 1);
    assert.equal(result.linkedSkipLocationCount, 1);
    assert.equal(result.linkedOriginalCount, 1);
    assert.equal(result.skippedNoPathCount, 1);
    assert.deepEqual(result.failed, []);
    assert.ok(events.indexOf('save:original') < events.indexOf('close:original:0:0'));
    assert.ok(events.indexOf('close:original:0:0') < events.indexOf(`open:${copiedProject.path}`));
    assert.ok(events.indexOf('offline:Copied') < events.indexOf('link:Copied:D:/Backup/CollectedMedias/copied.mov'));
    assert.ok(events.indexOf('offline:Ignored') < events.indexOf('link:Copied:D:/Backup/CollectedMedias/copied.mov'));
    assert.ok(events.indexOf('link:Copied:D:/Backup/CollectedMedias/copied.mov') < events.indexOf('save:backup'));
    assert.equal(events.includes('close:backup:0:0'), false);
    assert.equal(copiedItem.getMediaPath(), 'D:/Backup/CollectedMedias/copied.mov');
    assert.equal(ignoredItem.getMediaPath(), 'C:/Original/ignored.mov');
    assert.equal(skipItem.getMediaPath(), 'E:/Skip Library/skip.mov');
    assert.equal(result.backupLeftOpen, true);
    assert.equal(app.project, copiedProject);
});

test('Premiere host finds the BACKUP in app.projects when the original remains the active project', () => {
    const context = loadHostRelinkLogic();
    const events = [];
    const copiedItem = createHostMediaItem('Copied', 'C:/Original/copied.mov', events);
    const copiedProject = {
        path: 'D:/Backup/My Project BACKUP.prproj',
        rootItem: {
            type: 2,
            children: createHostChildren([copiedItem])
        },
        save() {
            events.push('save:backup');
            return 0;
        },
        closeDocument() {
            events.push('close:backup');
            return 0;
        }
    };
    const originalProject = {
        path: 'C:/Projects/My Project.prproj',
        rootItem: { type: 2, children: createHostChildren([]) },
        save() {
            events.push('save:original');
            return 0;
        },
        closeDocument() {
            events.push('close:original');
            app.project = otherProject;
            projects[1] = otherProject;
            projects.numProjects = 1;
            return 0;
        }
    };
    const otherProject = {
        path: 'C:/Projects/Other Project.prproj',
        rootItem: { type: 2, children: createHostChildren([]) }
    };
    const projects = {
        1: originalProject,
        numProjects: 1
    };
    const app = {
        project: originalProject,
        projects,
        openDocument(projectPath) {
            events.push(`open:${projectPath}`);
            if (projectPath === copiedProject.path) {
                projects[2] = copiedProject;
                projects.numProjects = 2;
            } else if (projectPath === originalProject.path) {
                this.project = originalProject;
            }
            return true;
        }
    };
    context.app = app;

    const result = JSON.parse(context.linkProjectCopyToCollectedMedia(
        copiedProject.path,
        JSON.stringify([{
            source: 'C:/Original/copied.mov',
            destination: 'D:/Backup/CollectedMedias/copied.mov',
            targetKind: 'collected'
        }])
    ));

    assert.equal(result.success, true);
    assert.equal(result.linkedCollectedCount, 1);
    assert.equal(copiedItem.getMediaPath(), 'D:/Backup/CollectedMedias/copied.mov');
    assert.equal(events.includes('save:backup'), true);
    assert.equal(events.includes('close:backup'), false);
    assert.equal(events.includes('close:original'), true);
    assert.equal(result.backupLeftOpen, true);
    assert.equal(app.project, otherProject);
});

test('Premiere host does not relink or save a partially offlined BACKUP project', () => {
    const context = loadHostRelinkLogic();
    const events = [];
    const firstItem = createHostMediaItem('First', 'C:/Original/first.mov', events);
    const blockedItem = createHostMediaItem('Blocked', 'C:/Original/blocked.mov', events, {
        offlineFails: true
    });
    const copiedProject = {
        path: 'D:/Backup/My Project BACKUP.prproj',
        rootItem: {
            type: 2,
            children: createHostChildren([firstItem, blockedItem])
        },
        save() {
            events.push('save:backup');
            return 0;
        },
        closeDocument(saveFirst, promptIfDirty) {
            events.push(`close:backup:${saveFirst}:${promptIfDirty}`);
            return 0;
        }
    };
    const originalProject = {
        path: 'C:/Projects/My Project.prproj',
        rootItem: { type: 2, children: createHostChildren([]) },
        save() {
            events.push('save:original');
            return 0;
        },
        closeDocument() {
            events.push('close:original');
            app.project = null;
            return 0;
        }
    };
    const app = {
        project: originalProject,
        openDocument(projectPath) {
            this.project = projectPath === copiedProject.path ? copiedProject : originalProject;
            return true;
        }
    };
    context.app = app;

    const tasks = JSON.stringify([
        {
            source: 'C:/Original/first.mov',
            destination: 'D:/Backup/CollectedMedias/first.mov',
            targetKind: 'collected'
        },
        {
            source: 'C:/Original/blocked.mov',
            destination: 'C:/Original/blocked.mov',
            targetKind: 'original'
        }
    ]);
    const result = JSON.parse(context.linkProjectCopyToCollectedMedia(copiedProject.path, tasks));

    assert.equal(result.success, false);
    assert.equal(result.offlineCount, 1);
    assert.equal(result.linkedCount, 0);
    assert.match(result.failed.join('\n'), /could not take the item offline/i);
    assert.equal(events.some((event) => event.startsWith('link:')), false);
    assert.equal(events.includes('save:backup'), false);
    assert.equal(events.includes('close:backup:0:0'), true);
    assert.equal(app.project, originalProject);
});
