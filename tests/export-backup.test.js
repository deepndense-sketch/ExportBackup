const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

test('selection excludes disabled-only audio in range and restores it when enabled', () => {
    const {context}=loadHostLogic();
    const clip=(disabled,start=0,end=60)=>{
        const c=makeTimedClip('Source.wav',end,'D:/Source.wav');
        c.start.seconds=start;c.disabled=disabled;return c;
    };
    const disabled=clip(true);
    const tracks=[makeTrack(0,[disabled,clip(1)]),makeTrack(0,[clip(true),clip(false)]),
        makeTrack(0,[clip(true),clip(false,120,180)]),makeTrack(1,[clip(false)]),makeTrack(0,[clip(undefined)])];
    const sequence=makeTimedSequence(tracks,[],'Show');
    context.app.project.activeSequence=sequence;
    const listed=()=>{
        const result=JSON.parse(context.exportBackup.getExportSelectionInfo());
        assert.equal(result.ok,true,result.message);
        return result.items.filter(item=>item.kind==='audio').map(item=>item.trackNumber);
    };
    assert.deepEqual(listed(),[2,4,5],'mixed enabled clips, muted tracks and unknown state stay eligible');
    disabled.disabled=false;
    assert.deepEqual(listed(),[1,2,4,5]);
    disabled.disabled=true;
    assert.deepEqual(listed(),[2,4,5]);
    assert.equal(context.ebTrackOccupiedInRange(tracks[0],{startSeconds:0,endSeconds:60}),true,
        'disabled clips still occupy their timeline space and cannot be overwritten');
});

test('empty-track preview reports the range-free track or proposed new track without creating it', () => {
    const {context}=loadHostLogic();
    const sequence=makeTimedSequence([],[makeTrack(0,[makeTimedClip('Source.mp4',60)]),makeTrack(0)],'Show');
    context.app.project.activeSequence=sequence;
    context.ebCreateTopVideoTrack=()=>assert.fail('preview must not create tracks');
    assert.deepEqual(JSON.parse(context.exportBackup.getEmptyTrackPreview()).trackNumber,2);
    sequence.videoTracks[1]=makeTrack(0,[makeTimedClip('Other.mp4',60)]);
    const result=JSON.parse(context.exportBackup.getEmptyTrackPreview());
    assert.equal(result.trackNumber,3);assert.equal(result.createTrack,true);
    sequence.getInPoint=()=>60;sequence.getOutPoint=()=>90;
    assert.equal(JSON.parse(context.exportBackup.getEmptyTrackPreview()).trackNumber,1);
});

test('backup imports into lower video/audio tracks that are empty only within its saved range', () => {
    const {context}=loadHostLogic();
    const clip=(name,start,end)=>{const c=makeTimedClip(name,end,'D:/'+name); c.start.seconds=start; return c;};
    const earlierVideo=clip('Show_BACKUP.mp4',0,60);
    const earlierAudio=clip('Show_Track1.wav',0,120);
    const laterAudio=clip('Later.wav',180,240);
    const audio=[makeTrack(0,[clip('Source.wav',120,180)]),makeTrack(0,[clip('Old.mp4',0,120),laterAudio]),makeTrack(0,[earlierAudio])];
    const video=[makeTrack(0,[clip('Source.mp4',120,180)]),makeTrack(0,[earlierVideo]),makeTrack(0)];
    const sequence=makeTimedSequence(audio,video,'Show');
    sequence.getInPoint=()=>120;sequence.getOutPoint=()=>180;
    context.app.project.activeSequence=sequence;
    assert.equal(context.ebResolveBackupVideoTrackNumber(sequence,1,true,false),2);
    assert.doesNotThrow(()=>context.ebValidateBackupTrack(sequence,2,false,'Show'));
    assert.throws(()=>context.ebValidateBackupTrack(sequence,1,false,'Show'),/not empty/);
    const writes=[];
    sequence.overwriteClip=(item,time,v,a)=>writes.push(['video',v,a,time]);
    audio.forEach((track,index)=>track.overwriteClip=(item,time)=>writes.push(['audio',index,time.seconds]));
    [earlierVideo,earlierAudio,laterAudio].forEach(c=>c.remove=()=>assert.fail('must preserve clips outside import range'));
    context.ebGetImportBin=()=>({name:'BACKUP'});
    context.ebImportProjectItem=mediaPath=>({name:mediaPath});
    context.ebRemoveUnusedMedia=()=>false;
    // Moving marks after export must not change import placement or occupancy checks.
    sequence.getInPoint=()=>500;sequence.getOutPoint=()=>540;
    const range={startSeconds:120,endSeconds:180};
    const result=JSON.parse(context.exportBackup.alignMappedFiles('D:/Show_BACKUP.mp4',JSON.stringify([{path:'D:/Show_Track1.wav',trackNumber:1,exportRange:range}]),2,false,'null','{}',JSON.stringify(range)));
    assert.equal(result.ok,true,result.message);
    assert.deepEqual(writes,[['video',1,1,120],['audio',2,120]]);
    assert.equal(audio[1].clips.numItems,2);
    assert.equal(audio[2].clips[0],earlierAudio);
});

test('recorded alignment refuses overlapping source media before removing existing backups', () => {
    const {context}=loadHostLogic();
    const old=makeTimedClip('Show_BACKUP.mp4',180,'D:/Show_BACKUP.mp4');old.start.seconds=120;
    old.remove=()=>assert.fail('preflight must finish before removing backups');
    const source=makeTimedClip('Source.wav',150,'D:/Source.wav');source.start.seconds=130;
    const sequence=makeTimedSequence([makeTrack(0,[source])],[makeTrack(0,[old])],'Show');
    context.app.project.activeSequence=sequence;
    context.ebGetImportBin=()=>({name:'BACKUP'});
    const range={startSeconds:120,endSeconds:180,targetTrackNumber:1};
    const result=JSON.parse(context.exportBackup.alignMappedFiles('D:/Show_BACKUP.mp4','[]',1,false,JSON.stringify({video:range,backupAudio:range,audioOutputs:[]}),'{}',JSON.stringify(range)));
    assert.equal(result.ok,false);
    assert.match(result.message,/A1 has clips inside the backup range/);
    assert.equal(context.ebTrackOccupiedInRange(makeTrack(0,[{projectItem:{name:'Unknown.wav'}}]),range),true);
});

test('selection lists only source clips in the marked section of mixed backup and teaser tracks', () => {
    const oldName = '3292 WOW P2';
    const sources = [1,2,3,4].map(n => makeTrack(0,[makeTimedClip('Source'+n+'.wav',1208,'D:/Source'+n+'.wav')]));
    const teaser = name => { const clip=makeTimedClip(name,1329,'D:/'+name); clip.start.seconds=1281; return clip; };
    const backup = suffix => makeTimedClip(oldName+suffix,1208,'D:/'+oldName+suffix);
    const sequence = makeTimedSequence(sources.concat([
        makeTrack(0,[backup('_BACKUP.mp4'),teaser('Teaser music.mp3')]),
        makeTrack(0,[backup('_Track1.wav'),teaser('Teaser atmosphere.mp3')]),
        makeTrack(0,[backup('_Track2.wav')]),
        makeTrack(0,[backup('_Track3-4.wav')])
    ]),[makeTrack(0,[backup('_BACKUP.mp4')])],oldName+' Copy 01');
    sequence.end=String(1329*254016000000);
    let start=0,end=1208;
    sequence.getInPoint=()=>start; sequence.getOutPoint=()=>end;
    const {context}=loadHostLogic(); context.app.project.activeSequence=sequence;
    const tracks=()=>{const result=JSON.parse(context.exportBackup.getExportSelectionInfo()); assert.equal(result.ok,true,result.message); return result.items.filter(x=>x.kind==='audio').map(x=>x.trackNumber);};
    assert.deepEqual(tracks(),[1,2,3,4],'main selection excludes backup clips on mixed tracks');
    start=1281;end=1329;
    assert.deepEqual(tracks(),[5,6],'teaser music remains selectable even on backup-named tracks');
    start=1208;end=1281;
    assert.deepEqual(tracks(),[],'gap and clips touching the boundaries do not add source tracks');
});

test('marked range is captured for new backups and original ranges survive changed teaser marks', () => {
    const {context} = loadHostLogic();
    const sequence = makeTimedSequence([],[],'Show');
    sequence.sequenceID='seq'; context.app.project.path='D:/Project.prproj';
    const selectedRange={startSeconds:120,endSeconds:180};
    const files=[{kind:'video',path:'D:/Show_BACKUP.mp4'},{kind:'audio',trackNumbers:[1],path:'D:/Show_Track1.wav'}];
    context.ebAssignRequestedExportRanges(sequence,files,null,selectedRange);
    selectedRange.startSeconds=1562; selectedRange.endSeconds=1600;
    assert.equal(files[0].exportRange.startSeconds,120);
    assert.equal(context.ebBuildQueuedFileFromRequested(files[1]).exportRange.endSeconds,180);
    const layout={video:{startSeconds:120,endSeconds:180},audioOutputs:[{sourceTrackNumbers:[1],startSeconds:120,endSeconds:180}]};
    context.ebAssignRequestedExportRanges(sequence,files,layout,selectedRange);
    assert.equal(files[0].exportRange.startSeconds,120);
    assert.equal(files[1].exportRange.endSeconds,180);
    const calls=[];
    sequence.setInPoint=value=>calls.push(['in',value]); sequence.setOutPoint=value=>calls.push(['out',value]);
    context.ebApplyExportRange(sequence,files[0].exportRange);
    assert.deepEqual(calls,[['in',120],['out',180]]);
    context.exportBackup.backupOwners={known:{path:'D:/Show_Track1.wav',sequenceID:'seq',projectPath:'D:/Project.prproj',exportRange:{startSeconds:120,endSeconds:180}}};
    context.ebAssignRequestedExportRanges(sequence,files.slice(1),{audioOutputs:[{mediaPath:'D:/Show_Track1.wav',sourceTrackNumbers:[1],startSeconds:0}]},selectedRange);
    assert.equal(files[1].exportRange.startSeconds,120,'project-only backup uses persisted range');
    const clip=makeTimedClip('Show_Track1.wav',60,'D:/Show_Track1.wav');clip.start.seconds=120;clip.end.seconds=180;
    sequence.end=String(1600*254016000000);sequence.getInPoint=()=>1562;sequence.getOutPoint=()=>1600;
    assert.ok(context.ebGetBackupTrackCandidate(sequence,makeTrack(0,[clip]),'Show'),'recorded backup is still recognized after marking the teaser');
});

test('new backup imports at saved In point and Align Existing keeps its previous position', () => {
    const {context} = loadHostLogic();
    const placements=[];
    const audio=Array.from({length:5},()=>makeTrack(0));
    audio.forEach((track,index)=>{track.overwriteClip=(item,time)=>placements.push(['audio',index,time.seconds]);});
    const sequence=makeTimedSequence(audio,[makeTrack(0),makeTrack(0)],'Show');
    sequence.getInPoint=()=>1562;sequence.getOutPoint=()=>1600;
    sequence.overwriteClip=(item,time,videoTrack,audioTrack)=>placements.push(['video',videoTrack,time,audioTrack]);
    context.app.project.activeSequence=sequence;
    context.ebGetHighestSourceAudioTrackNumber=()=>2;
    context.ebGetImportBin=()=>({name:'BACKUP'});
    context.ebImportProjectItem=mediaPath=>({name:mediaPath});
    context.ebRemoveManagedClipsFromAllAudioTracks=()=>{};
    context.ebRemoveManagedClipsFromTrack=()=>{};
    context.ebRemoveUnusedMedia=()=>false;
    const range={startSeconds:120,endSeconds:180};
    let result=JSON.parse(context.exportBackup.alignMappedFiles('D:/Show_BACKUP.mp4',JSON.stringify([{path:'D:/Show_Track1.wav',trackNumber:1,exportRange:range}]),2,false,'null','{}',JSON.stringify(range)));
    assert.equal(result.ok,true,result.message);
    assert.equal(placements[0][2],120);
    assert.equal(placements[1][2],120);
    placements.length=0;
    const layout={video:{targetTrackNumber:2,startSeconds:120,endSeconds:180},backupAudio:{targetTrackNumber:3,startSeconds:120},audioOutputs:[{sourceTrackNumbers:[1],targetTrackNumber:4,startSeconds:120,endSeconds:180}]};
    result=JSON.parse(context.exportBackup.alignMappedFiles('D:/Show_BACKUP.mp4',JSON.stringify([{path:'D:/Show_Track1.wav',trackNumber:1}]),2,false,JSON.stringify(layout),'{}'));
    assert.equal(result.ok,true,result.message);
    assert.equal(placements[0][2],120);
    assert.equal(placements[1][2],120);
});

test('partial Re-backup exports exactly the checked items and preserves unchecked backup placements', () => {
    const mainSource = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
    const inputs = [1,2,3,4,5].map(number => ({checked:false,disabled:false,getAttribute:()=>String(number)}));
    const video = {checked:true};
    const panel = vm.createContext({document:{getElementById:()=>({querySelector:()=>video,querySelectorAll:()=>inputs})},mergedAudioGroups:[{trackNumbers:[3,4,5]}]});
    vm.runInContext(mainSource.slice(mainSource.indexOf('function getSelectedQueueItems('),mainSource.indexOf('async function refreshExportSelection(')),panel);
    const {context} = loadHostLogic();
    const sequence = makeTimedSequence(inputs.map((_,index)=>makeTrack(0,[makeTimedClip('Source.wav',60,'D:/Source'+index+'.wav')])));
    const layout = {video:{mediaPath:'D:/Show_BACKUP.mp4',targetTrackNumber:6},backupAudio:{targetTrackNumber:6},audioOutputs:[
        {sourceTrackNumbers:[1],sourceTrackNumber:1,targetTrackNumber:7,mediaPath:'D:/Show_Track1.wav'},
        {sourceTrackNumbers:[2],sourceTrackNumber:2,targetTrackNumber:8,mediaPath:'D:/Show_Track2.wav'},
        {sourceTrackNumbers:[3,4,5],sourceTrackNumber:3,targetTrackNumber:9,mediaPath:'D:/Show_Track3-4-5.wav'}
    ]};
    const build = () => {
        const selected = panel.getSelectedQueueItems();
        const outputs = context.ebBuildRequestedOutputFiles(sequence,'D:/','video.epr','audio.epr','wav',JSON.stringify(selected),true,layout);
        assert.equal(context.ebGetSelectedExportItems(JSON.stringify(selected)).preserveUnselectedAudio,selected.preserveUnselectedAudio);
        return {selected,outputs};
    };
    let result = build();
    assert.deepEqual(Array.from(result.outputs,entry=>entry.kind),['video']);
    assert.equal(context.ebGetObsoleteRebackupAudioFiles(layout,result.outputs,true,true).length,0);
    assert.equal(context.ebBuildRebackupAlignmentLayout(layout,result.outputs,true,true).audioOutputs.length,3);
    video.checked = false; inputs[1].checked = true;
    result = build();
    assert.equal(result.outputs.length,1);
    assert.equal(result.outputs[0].kind,'audio');
    assert.equal(result.outputs[0].trackNumber,2);
    assert.equal(context.ebGetObsoleteRebackupAudioFiles(layout,result.outputs,true,true).length,0);
    assert.equal(context.ebBuildRebackupAlignmentLayout(layout,result.outputs,true,true).audioOutputs[1].targetTrackNumber,8);
    inputs[1].checked = false; inputs[2].checked = true;
    result = build();
    assert.deepEqual(Array.from(result.outputs[0].trackNumbers),[3],'one checked member does not export the whole merged group');
    assert.equal(context.ebGetObsoleteRebackupAudioFiles(layout,result.outputs,true,true).length,0,'keep the old group containing unchecked tracks');
    video.checked = true; inputs.forEach(input=>{input.checked=true;});
    result = build();
    assert.equal(result.selected.preserveUnselectedAudio,false);
    assert.equal(result.outputs.length,4,'full Re-backup still exports video, two singles, and the merged group');
    assert.deepEqual(Array.from(result.outputs[3].trackNumbers),[3,4,5]);
});

test('timeline refresh before export retains unchecked items and does not select newly discovered tracks', async () => {
    const source = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
    const nodes = [{kind:'video',number:0,checked:false},{kind:'audio',number:1,checked:false},{kind:'audio',number:2,checked:true}]
        .map(value=>({...value,getAttribute:name=>name==='data-kind'?value.kind:String(value.number)}));
    const latest = {ok:true,sequenceName:'Current',items:[{kind:'video',trackNumber:0},{kind:'audio',trackNumber:1},{kind:'audio',trackNumber:2},{kind:'audio',trackNumber:3}]};
    let rendered;
    const context = vm.createContext({exportSelectionState:{sequenceName:'Current',items:latest.items.slice(0,3)},mergedAudioGroups:[],
        document:{getElementById:()=>({querySelectorAll:()=>nodes})},getExportSelectionInfo:async()=>latest,renderExportSelectionList:value=>{rendered=value;},setStatus(){}});
    vm.runInContext(source.slice(source.indexOf('function getExportSelectionTrackSignature('),source.indexOf('function parseTrackNumbersFromFileName(')),context);
    await context.syncExportSelectionWithActiveTimeline();
    assert.deepEqual(rendered.items.map(item=>item.selected),[false,false,true,false]);
});

test('disk backup renaming relinks duplicate video/audio items and retains unsafe files', () => {
    const directory = fs.mkdtempSync(path.join(require('node:os').tmpdir(), 'backup-rename-'));
    try {
        const DiskFile = function(value) {
            this.fsName = path.resolve(String(value));
            this.name = path.basename(this.fsName);
            Object.defineProperty(this, 'exists', {get:() => fs.existsSync(this.fsName)});
            this.rename = name => { fs.renameSync(this.fsName, path.join(path.dirname(this.fsName), name)); return true; };
        };
        const {context} = loadHostLogic(null, {File:DiskFile});
        const names = ['Old_BACKUP.mp4','Old_Track1.wav','Old_Track2-3.mp3','Foreign_Track4.wav','Short_Track5.wav'];
        const clips = names.map((name,index) => {
            let currentPath = path.join(directory,name);
            fs.writeFileSync(currentPath,'media-' + index);
            const clip = makeTimedClip(name,index === 4 ? 20 : 60,currentPath);
            clip.name = name;
            clip.projectItem.getMediaPath = () => currentPath;
            clip.projectItem.changeMediaPath = next => {currentPath=next;return 0;};
            return clip;
        });
        let duplicatePath = clips[0].projectItem.getMediaPath();
        const duplicate = {name:names[0],getMediaPath:() => duplicatePath,changeMediaPath:next => {duplicatePath=next;return 0;}};
        const linked = {name:names[0],start:{seconds:0},end:{seconds:60},projectItem:duplicate};
        const sequence = makeTimedSequence([makeTrack(0,[linked]),...clips.slice(1).map(clip=>makeTrack(0,[clip]))],[makeTrack(0,[clips[0]])],'New Show');
        sequence.sequenceID = 'current';
        context.app.project.path = 'D:/Project.prproj';
        context.app.project.rootItem.children = makeCollection([...clips.map(clip=>clip.projectItem),duplicate],'numItems');
        const owners = {}, paths = clips.map(clip=>clip.projectItem.getMediaPath());
        paths.forEach((mediaPath,index)=>{owners[mediaPath]={path:mediaPath,projectPath:'D:/Project.prproj',sequenceID:index===3?'foreign':'current'};});
        const report = {files:[],warnings:[]};
        context.ebRenameOwnedBackupClips(sequence,owners,paths,report);
        assert.equal(report.files.length,3);
        assert.equal(report.warnings.length,0);
        for (let i=0;i<3;i++) {
            const expected = path.join(directory,names[i].replace('Old','New Show'));
            assert.equal(fs.readFileSync(expected,'utf8'),'media-'+i);
            assert.equal(fs.existsSync(paths[i]),false);
            assert.equal(clips[i].projectItem.getMediaPath(),expected);
            assert.equal(clips[i].projectItem.name,path.basename(expected));
        }
        assert.equal(duplicatePath,clips[0].projectItem.getMediaPath());
        assert.equal(linked.name,'New Show_BACKUP.mp4');
        assert.equal(fs.existsSync(paths[3]),true);
        assert.equal(fs.existsSync(paths[4]),true);
        // Never overwrite a different existing file.
        fs.writeFileSync(path.join(directory,'Taken.wav'),'keep');
        const collision = {files:[],warnings:[]};
        assert.equal(context.ebRenameBackupMediaFile(paths[3],'Taken.wav',collision),'');
        assert.equal(fs.readFileSync(path.join(directory,'Taken.wav'),'utf8'),'keep');
        // A partial relink failure restores the original filename and changed references.
        duplicate.changeMediaPath = () => 1;
        const rollback = {files:[],warnings:[]};
        const current = clips[0].projectItem.getMediaPath();
        assert.equal(context.ebRenameBackupMediaFile(current,'Again_BACKUP.mp4',rollback),'');
        assert.equal(fs.existsSync(current),true);
        assert.equal(fs.existsSync(path.join(directory,'Again_BACKUP.mp4')),false);
        assert.equal(clips[0].projectItem.getMediaPath(),current);
        assert.equal(rollback.files.length,0);
        assert.match(rollback.warnings[0],/rolled back/);
    } finally { fs.rmSync(directory,{recursive:true,force:true}); }
});

test('readable prompts preserve long names, default to cancel and queue completion messages', async () => {
    const source = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
    const elements = {};
    const document = {activeElement:null,getElementById:id=>elements[id]};
    for (const id of ['readablePrompt','readablePromptTitle','readablePromptMessage','readablePromptOK','readablePromptCancel','readablePromptSuccess','readablePromptDetails']) {
        elements[id] = {classList:{add(){},remove(){}},setAttribute(name,value){this[name]=value;},focus(){document.activeElement=this;},addEventListener(type,fn){this[type]=fn;},removeEventListener(type){delete this[type];}};
    }
    const context = vm.createContext({document});
    vm.runInContext(source.slice(source.indexOf('let readablePromptQueue'),source.indexOf('const ALIGNMENT_RECOVERY_TEXT')),context);
    const message = 'Long file name '.repeat(1000)+'\nAnother file.wav';
    const confirmation = context.showReadablePrompt({message,cancelText:'Keep names',confirmText:'Rename files'});
    const completion = context.showBlockingMessage('Completed with details', 'success');
    await Promise.resolve();
    assert.equal(elements.readablePromptMessage.textContent,message);
    assert.equal(document.activeElement,elements.readablePromptCancel);
    elements.readablePrompt.keydown({key:'Escape',preventDefault(){},stopPropagation(){}});
    assert.equal(await confirmation,false);
    await Promise.resolve();
    assert.equal(elements.readablePromptMessage.textContent,'Completed with details');
    assert.equal(elements.readablePromptSuccess.textContent,'✓');
    assert.equal(elements.readablePromptMessage.hidden,true);
    elements.readablePromptDetails.onclick();
    assert.equal(elements.readablePromptMessage.hidden,false);
    elements.readablePromptOK.onclick();
    assert.equal(await completion,true);
    for (const kind of ['error','warning']) {
        const problem = context.showReadablePrompt({message:'Needs attention',kind});
        await Promise.resolve();
        assert.equal(elements.readablePromptSuccess.className,'success-mark '+kind);
        assert.equal(elements.readablePromptMessage.hidden,false);
        assert.equal(elements.readablePromptDetails.hidden,true);
        elements.readablePromptOK.onclick();
        await problem;
    }
});

test('disk rename updates ownership and manifest paths for subsequent cleanup and rebackup', () => {
    const source = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
    let saved;
    const context = vm.createContext({getBackupClipOwners:()=>({'d:/old.wav':{path:'D:/Old.wav',sequenceID:'seq'}}),getPathComparisonKey:value=>String(value).toLowerCase(),localStorage:{setItem:(key,value)=>{saved=JSON.parse(value);}}});
    vm.runInContext(source.slice(source.indexOf('function applyBackupFileRenames('),source.indexOf('function createExportManifestFromHostResult(')),context);
    const info = {audio:[{path:'D:/Old.wav'}],manifest:{expectedFiles:[{finalPath:'D:/Old.wav'}]}};
    context.applyBackupFileRenames(info,[{oldPath:'D:/Old.wav',newPath:'D:/New.wav'}]);
    assert.equal(info.audio[0].path,'D:/New.wav');
    assert.equal(info.manifest.expectedFiles[0].finalPath,'D:/New.wav');
    assert.equal(saved['d:/new.wav'].sequenceID,'seq');
    assert.equal(saved['d:/old.wav'],undefined);
});

test('alignment renames only recorded full-sequence backup clips and preserves other media names', () => {
    const {context} = loadHostLogic();
    const own = makeTimedClip('Old Show_BACKUP.mp4',60,'D:/Old Show_BACKUP.mp4');
    const foreign = makeTimedClip('Other_BACKUP.mp4',60,'D:/Other_BACKUP.mp4');
    const short = makeTimedClip('Old Show_Track1.wav',20,'D:/Old Show_Track1.wav');
    const unknown = makeTimedClip('Unrecorded_BACKUP.mp4',60,'D:/Unrecorded_BACKUP.mp4');
    for (const clip of [own,foreign,short,unknown]) clip.name = clip.projectItem.name;
    const sequence = makeTimedSequence([makeTrack(0,[short])],[makeTrack(0,[own]),makeTrack(0,[foreign]),makeTrack(0,[unknown])]);
    sequence.name = 'New Show'; sequence.sequenceID = 'owned-sequence';
    context.app.project.path = 'D:/Project.prproj';
    const owners = {};
    for (const clip of [own,short,foreign]) {
        const mediaPath = clip.projectItem.getMediaPath();
        owners[mediaPath] = {path:mediaPath,projectPath:'D:/Project.prproj',sequenceID:clip === foreign ? 'another-sequence' : 'owned-sequence'};
    }
    const paths = [own,foreign,short,unknown].map(clip => clip.projectItem.getMediaPath());
    assert.equal(context.ebRenameOwnedBackupClips(sequence,owners,paths),1);
    assert.equal(own.name,'New Show_BACKUP.mp4');
    assert.equal(own.projectItem.name,'Old Show_BACKUP.mp4','shared project item is not renamed');
    assert.equal(foreign.name,'Other_BACKUP.mp4');
    assert.equal(short.name,'Old Show_Track1.wav');
    assert.equal(unknown.name,'Unrecorded_BACKUP.mp4');
    assert.equal(context.ebRenameOwnedBackupClips(sequence,owners,paths),0);
    sequence.name = 'Changed Again';
    context.app.project.path = 'D:/AnotherProject.prproj';
    assert.equal(context.ebRenameOwnedBackupClips(sequence,owners,paths),0);
    assert.equal(own.name,'New Show_BACKUP.mp4');
});

test('legacy rename review includes full-length MP4, WAV and merged MP3 but excludes other owners and short clips', () => {
    const {context} = loadHostLogic();
    const clips = ['FN3 P5_BACKUP.mp4','FN3 P5_Track1.wav','FN3 P5_Track2-3.mp3','Other_Track1.wav','Short_Track2.wav'].map((name,i) => {
        const clip = makeTimedClip(name,i === 4 ? 10 : 60,'D:/'+name); clip.name=name; return clip;
    });
    const sequence = makeTimedSequence(clips.slice(1).map(clip => makeTrack(0,[clip])),[makeTrack(0,[clips[0]])]);
    sequence.sequenceID = 'current'; sequence.name = 'New Full Sequence Name';
    context.app.project.activeSequence = sequence; context.app.project.path = 'D:/Project.prproj';
    const paths = clips.map(clip => clip.projectItem.getMediaPath());
    const owners = {other:{path:paths[3],sequenceID:'other',projectPath:'D:/Project.prproj'}};
    const result = JSON.parse(context.exportBackup.getBackupRenameCandidates(JSON.stringify(paths),JSON.stringify(owners)));
    assert.equal(result.ok,true);
    assert.deepEqual(result.candidates.map(entry => entry.name),['FN3 P5_BACKUP.mp4','FN3 P5_Track1.wav','FN3 P5_Track2-3.mp3']);
    result.candidates.forEach(entry => {owners[entry.path]={path:entry.path,sequenceID:'current',projectPath:'D:/Project.prproj'};});
    assert.equal(context.ebRenameOwnedBackupClips(sequence,owners,paths),3);
    assert.equal(clips[1].name,'New Full Sequence Name_Track1.wav');
    assert.equal(clips[2].name,'New Full Sequence Name_Track2-3.mp3');
    assert.equal(clips[3].name,'Other_Track1.wav');
});

test('legacy name review only records ownership after user confirmation', async () => {
    const source = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
    let confirmed = false, records = 0;
    const context = vm.createContext({showReadablePrompt:async () => confirmed,parseHostResult:JSON.parse,escapeForEvalScript:value => value,
        getBackupClipOwners:() => ({}),rememberBackupClipOwners:() => records++,
        callHost:async () => JSON.stringify({ok:true,sequenceID:'seq',projectPath:'D:/Project.prproj',sequenceName:'Show',candidates:[{path:'D:/Old_Track1.wav',name:'Old_Track1.wav'}]})});
    vm.runInContext(source.slice(source.indexOf('async function confirmLegacyBackupNames('),source.indexOf('function createExportManifestFromHostResult(')),context);
    await context.confirmLegacyBackupNames(['D:/Old_Track1.wav']);
    assert.equal(records,0);
    confirmed = true;
    await context.confirmLegacyBackupNames(['D:/Old_Track1.wav']);
    assert.equal(records,1);
});

test('export ownership records final paths under stable sequence IDs', () => {
    const source = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
    let stored = '{}';
    const context = vm.createContext({localStorage:{getItem:() => stored,setItem:(key,value) => {stored=value;}},getPathComparisonKey:value => value.toLowerCase()});
    vm.runInContext(source.slice(source.indexOf('function getBackupClipOwners('),source.indexOf('function createExportManifestFromHostResult(')),context);
    context.rememberBackupClipOwners({sequenceID:'stable-id',projectPath:'D:/Project.prproj',queuedFiles:[{path:'D:/Old_BACKUP_REBKP_TEMP.mp4',finalPath:'D:/Old_BACKUP.mp4'}]});
    assert.deepEqual(JSON.parse(stored)['d:/old_backup.mp4'],{path:'D:/Old_BACKUP.mp4',sequenceID:'stable-id',projectPath:'D:/Project.prproj'});
});

test('update prompt offers Later without installing and blocks updates during Collector copying', () => {
    const source = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
    const elements = {};
    const document = {getElementById:id => elements[id], activeElement:null};
    for (const id of ['updatePrompt','updatePromptTitle','updatePromptNotes','updateLaterButton','updateInstallButton']) {
        const classes = new Set(id === 'updatePrompt' ? ['is-hidden'] : []);
        elements[id] = {classList:{contains:value => classes.has(value),add:value => classes.add(value),remove:value => classes.delete(value)},focus() {document.activeElement=this;}};
    }
    let copying = false, installs = 0;
    elements.collectorPanel = {contentWindow:{isCollectorBusy:() => copying}};
    const context = vm.createContext({document,busy:false,remoteVersion:'4.8.0',remoteVersionNotes:'First fix\nSecond fix <script>text only</script>',localVersion:'4.7.6',compareVersions:() => 1,runGithubUpdate:() => installs++});
    vm.runInContext(source.slice(source.indexOf("let promptedUpdateVersion = ''"),source.indexOf('function escapeForEvalScript(')),context);
    context.showUpdatePrompt(true);
    assert.equal(elements.updatePromptTitle.textContent,'Backup Project 4.8.0 is available');
    assert.equal(elements.updatePromptNotes.textContent,context.remoteVersionNotes);
    assert.equal(elements.updatePrompt.classList.contains('is-hidden'),false);
    context.closeUpdatePrompt();
    context.showUpdatePrompt(true);
    assert.equal(elements.updatePrompt.classList.contains('is-hidden'),true);
    assert.equal(installs,0);
    copying = true;
    context.showUpdatePrompt(false);
    context.installPromptedUpdate();
    assert.equal(installs,0);
    copying = false;
    context.remoteVersionNotes='';
    context.showUpdatePrompt(false);
    assert.equal(elements.updatePromptNotes.textContent,'No release notes were provided for this version.');
    context.installPromptedUpdate();
    assert.equal(installs,1);
    assert.equal(elements.updatePrompt.classList.contains('is-hidden'),true);
});

test('updater launch preserves spaces, apostrophes, and literal PowerShell characters', () => {
    const source = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
    const context = vm.createContext({Buffer});
    vm.runInContext(source.slice(source.indexOf('function buildUpdaterLaunchCommand('),source.indexOf('function runGithubUpdate(')),context);
    const values = ["C:/Users/O'Brien/update.ps1",'C:/Temp/package.zip','C:/CEP/Backup Project','C:/Temp/result.json','C:/Temp/$log.txt'];
    const command = context.buildUpdaterLaunchCommand(...values);
    const encoded = command.match(/'-EncodedCommand','([^']+)'/)[1];
    const payload = Buffer.from(encoded,'base64').toString('utf16le');
    assert.equal(payload, "& 'C:/Users/O''Brien/update.ps1' -ZipPath 'C:/Temp/package.zip' -Destination 'C:/CEP/Backup Project' -ResultPath 'C:/Temp/result.json' -LogPath 'C:/Temp/$log.txt'");
    assert.match(command, /-WindowStyle Hidden/);
});

test('Align Existing discovers old exports and merged-track leftovers without touching unrelated media', async () => {
    const folder = fs.mkdtempSync(path.join(require('node:os').tmpdir(),'backup-cleanup-'));
    try {
        const names = ['Show_BACKUP.mp4','Show_Track3-4.wav','Show_BACKUP_REBKP_OLD_123_1.mp4',
            'Show_Track3_REBKP_OLD_123_2.wav','Show_Track3.wav','Show_Track4.wav',
            'Other_Track3.wav','Show_Track5.wav','Show_Track3-4_REBKP_TEMP.wav','source.png'];
        names.forEach(name => fs.writeFileSync(path.join(folder,name),'ready'));
        const source = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
        const key = file => path.resolve(file).toLowerCase();
        const context = vm.createContext({fs,path,getPathComparisonKey:key,parseHostResult:JSON.parse,escapeForEvalScript:value => value,
            deleteLocalFileNow:file => {fs.unlinkSync(file);return {ok:true};},
            callHost:async () => JSON.stringify({ok:true,safePaths:names.filter(name => name !== 'Show_Track4.wav').map(name => path.join(folder,name))})});
        vm.runInContext(source.slice(source.indexOf('function findBackupLeftovers('),source.indexOf('async function prepareAlignExistingCleanup(')),context);
        const info = {videoPath:path.join(folder,'Show_BACKUP.mp4'),audio:[{path:path.join(folder,'Show_Track3-4.wav')}]};
        const candidates = context.findBackupLeftovers(info);
        assert.equal(candidates.length,5);
        const result = await context.cleanupDiscoveredBackupLeftovers(info);
        assert.equal(result.deleted.length,4);
        assert.deepEqual(Array.from(result.retained),[path.join(folder,'Show_Track4.wav')]);
        for (const name of ['Show_BACKUP.mp4','Show_Track3-4.wav','Other_Track3.wav','Show_Track5.wav','source.png','Show_Track4.wav']) assert.ok(fs.existsSync(path.join(folder,name)));
        assert.equal(context.findBackupLeftovers(info).length,1);
    } finally { fs.rmSync(folder,{recursive:true,force:true}); }
});

test('leftover project-item release never removes media still used by another sequence', () => {
    const {context} = loadHostLogic();
    const sequence = makeTimedSequence([makeTrack(0,[makeTimedClip('Old.wav',60,'D:/Show_Track3.wav')])]);
    context.app.project.sequences = makeCollection([sequence],'numSequences');
    const released = [];
    context.ebReleaseProjectItemsByMediaPath = file => {released.push(file);return {remaining:0};};
    const result = JSON.parse(context.exportBackup.releaseUnusedBackupLeftovers(JSON.stringify(['D:/Show_Track3.wav','D:/Show_BACKUP_REBKP_OLD_1.mp4'])));
    assert.equal(result.ok,true);
    assert.deepEqual(result.retainedPaths,['D:/Show_Track3.wav']);
    assert.deepEqual(released,['D:/Show_BACKUP_REBKP_OLD_1.mp4']);
});

test('Project Root previews and creates a BACKUP subfolder beside the project', async () => {
    const source = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
    const created = [];
    const context = vm.createContext({path:path.win32,
        fs:{mkdirSync:(folder,options) => created.push({folder,recursive:options.recursive}),existsSync:() => false},
        BACKUP_DESTINATION_PROJECT_ROOT:'projectRoot',BACKUP_DESTINATION_PARENT_FOLDER:'parentFolder',BACKUP_DESTINATION_MANUAL:'manual',
        getSelectedBackupDestination:() => 'projectRoot',
        getActiveProjectInfo:async () => ({ok:true,projectPath:'D:\\Edits\\Show.prproj'}),
        updateCategoryDestinationNote() {}});
    vm.runInContext(source.slice(source.indexOf('async function resolveProjectBackupFolder('),source.indexOf('async function refreshResolvedBackupDestination(')),context);
    assert.equal((await context.resolveProjectBackupFolder({create:false})).folderPath,'D:\\Edits\\BACKUP');
    assert.equal(created.length,0);
    assert.equal((await context.resolveProjectBackupFolder({create:true})).folderPath,'D:\\Edits\\BACKUP');
    assert.deepEqual(created,[{folder:'D:\\Edits\\BACKUP',recursive:true}]);
});

test('Parent Folder previews outside the project directory and reuses an existing BACKUP folder', async () => {
    const source = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
    const created = [];
    const context = vm.createContext({path:path.win32,
        fs:{mkdirSync:(folder,options) => created.push({folder,recursive:options.recursive}),existsSync:() => true},
        BACKUP_DESTINATION_PROJECT_ROOT:'projectRoot',BACKUP_DESTINATION_PARENT_FOLDER:'parentFolder',BACKUP_DESTINATION_MANUAL:'manual',
        getSelectedBackupDestination:() => 'parentFolder',
        getActiveProjectInfo:async () => ({ok:true,projectPath:'D:\\Show\\Project\\Edit.prproj'}),
        updateCategoryDestinationNote() {}});
    vm.runInContext(source.slice(source.indexOf('async function resolveProjectBackupFolder('),source.indexOf('async function refreshResolvedBackupDestination(')),context);
    assert.equal((await context.resolveProjectBackupFolder({create:false,requireExisting:true})).folderPath,'D:\\Show\\BACKUP');
    assert.equal(created.length,0);
    assert.equal((await context.resolveProjectBackupFolder({create:true})).folderPath,'D:\\Show\\BACKUP');
    assert.deepEqual(created,[{folder:'D:\\Show\\BACKUP',recursive:true}]);
    context.getActiveProjectInfo = async () => ({ok:true,projectPath:'D:\\Edit.prproj'});
    assert.equal((await context.resolveProjectBackupFolder({create:false})).folderPath,'D:\\BACKUP');
});

test('destination dropdown selects the existing routing mode and only offers browsing for manual paths', () => {
    const source = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const elements = {};
    for (const id of ['backupDestinationFtp','backupDestinationProjectRoot','backupDestinationParentFolder','backupDestinationManual','backupDestinationSelect','chooseFolderButton','exportPath']) {
        elements[id] = {checked:id === 'backupDestinationFtp', events:{}, addEventListener(name, fn) {this.events[name]=fn;}, dispatchEvent(event) {return this.events[event.type]();}};
    }
    const context = vm.createContext({document:{getElementById:id => elements[id]},
        Event:class {constructor(type) {this.type=type;}}, localStorage:{setItem() {}},
        BACKUP_DESTINATION_FTP:'ftp', BACKUP_DESTINATION_PROJECT_ROOT:'projectRoot', BACKUP_DESTINATION_PARENT_FOLDER:'parentFolder', BACKUP_DESTINATION_MANUAL:'manual',
        BACKUP_DESTINATION_STORAGE_KEY:'destination', manualExportFolder:'D:/Manual', exportFolder:null,
        refreshResolvedBackupDestination:async () => context.updateDestinationButtonLabel()});
    for (const name of ['getBackupDestinationInputs','getSelectedBackupDestination','updateDestinationButtonLabel','bindBackupDestinationInputs']) {
        vm.runInContext(source.match(new RegExp('^function ' + name + '\\([^\\n]*\\).*?^}', 'ms'))[0],context);
    }
    context.bindBackupDestinationInputs();
    for (const mode of ['manual','projectRoot','parentFolder','ftp']) {
        elements.backupDestinationSelect.value=mode;
        elements.backupDestinationSelect.events.change();
        assert.equal(context.getSelectedBackupDestination(),mode);
        assert.equal(elements.chooseFolderButton.hidden,mode !== 'manual');
    }
});

// Give older layout fixtures explicit whole-sequence timing.
function completeBackupFixture(sequence) {
    if (sequence.end === undefined) sequence.end = String(60 * 254016000000);
    for (const tracks of [sequence.audioTracks, sequence.videoTracks]) {
        if (!tracks) continue;
        for (let i = 0; i < tracks.numTracks; i++) {
            for (let j = 0; j < tracks[i].clips.numItems; j++) {
                const clip = tracks[i].clips[j];
                if (!clip.start) clip.start = { seconds: 0 };
                if (!clip.end) clip.end = { seconds: 60 };
            }
        }
    }
}

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

function makeTimedClip(name, seconds, mediaPath) {
    return {
        start: { seconds: 0 }, end: { seconds },
        projectItem: { name, getMediaPath() { return mediaPath || ''; } }
    };
}

function makeTimedSequence(audioTracks, videoTracks, name) {
    return {
        name: name || 'Current Show',
        end: String(60 * 254016000000),
        getInPoint() { return 0; },
        getOutPoint() { return 60; },
        audioTracks: makeCollection(audioTracks, 'numTracks'),
        videoTracks: makeCollection(videoTracks || [makeTrack(0)], 'numTracks')
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

test('renamed sequences retain existing backup paths and exclude their audio outputs', () => {
    const sequence = makeTimedSequence([
        makeTrack(0,[makeTimedClip('Voice.wav',60,'D:/Voice.wav')]),
        makeTrack(0,[makeTimedClip('Old_BACKUP.mp4',60,'D:/Backups/Old_BACKUP.mp4')]),
        makeTrack(0,[makeTimedClip('Old_Track1.wav',60,'D:/Backups/Old_Track1.wav')])
    ],[makeTrack(0,[makeTimedClip('Old_BACKUP.mp4',60,'D:/Backups/Old_BACKUP.mp4')])], 'Renamed without category');
    const {context} = loadHostLogic();
    context.app.project.activeSequence = sequence;
    const layout = JSON.parse(context.exportBackup.getActiveBackupLayout()).layout;
    assert.match(layout.video.mediaPath, /Old_BACKUP\.mp4$/);
    assert.equal(layout.baseName,'old');
    const preview = JSON.parse(context.exportBackup.getExportSelectionInfo());
    assert.deepEqual(preview.items.filter(item => item.kind === 'audio').map(item => item.trackNumber),[1]);
    const outputs = context.ebBuildRequestedOutputFiles(sequence,'D:/Backups','video.epr','audio.epr','wav',JSON.stringify({includeVideo:true,audioTracks:[1],replaceAudioLayout:true}),true,layout);
    assert.equal(outputs.length,2);
    assert.ok(outputs.every(output => /Old_(BACKUP|Track1)/.test(output.finalPath)));
});

test('project-bin-only backups are discovered by exact sequence identity', () => {
    const {context} = loadHostLogic();
    context.app.project.activeSequence = makeTimedSequence([makeTrack(0,[makeTimedClip('Voice.wav',60)])],[], 'Show');
    context.app.project.rootItem.children = makeCollection([
        {children:makeCollection([],'numItems'),getMediaPath:() => 'D:/Existing/Show_BACKUP.mp4'},
        {children:makeCollection([],'numItems'),getMediaPath:() => 'E:/Audio/Show_Track1.wav'},
        {getMediaPath:() => 'D:/Other/Other_BACKUP.mp4'}
    ],'numItems');
    const result = JSON.parse(context.exportBackup.getActiveBackupLayout());
    assert.equal(result.ok,true);
    assert.match(result.layout.video.mediaPath,/Show_BACKUP.mp4$/);
    assert.equal(result.layout.audioOutputs.length,1);
    assert.match(result.layout.audioOutputs[0].mediaPath,/Show_Track1.wav$/);
});

test('removing backup video still discovers its family from remaining audio and sibling files', () => {
    const backupFolder = {exists:true,fsName:'D:/Existing',getFiles:() => [
        {exists:true,name:'Old_BACKUP.mp4',fsName:'D:/Existing/Old_BACKUP.mp4'},
        {exists:true,name:'Other_BACKUP.mp4',fsName:'D:/Existing/Other_BACKUP.mp4'}
    ]};
    const {context} = loadHostLogic(null, {File:function(filePath) {
        this.fsName = String(filePath).replace(/\//g,'\\');
        this.name = this.fsName.split('\\').pop();
        this.exists = true;
        this.parent = backupFolder;
    }});
    const sequence = makeTimedSequence([
        makeTrack(0,[makeTimedClip('Voice.wav',60,'D:/Source/Voice.wav')]),
        makeTrack(0,[makeTimedClip('Old_Track1.wav',60,'D:/Existing/Old_Track1.wav')])
    ],[makeTrack(0)],'Renamed sequence');
    context.app.project.activeSequence=sequence;
    const result=JSON.parse(context.exportBackup.getActiveBackupLayout());
    assert.equal(result.ok,true);
    assert.equal(result.layout.baseName,'old');
    assert.match(result.layout.video.mediaPath,/Old_BACKUP.mp4$/);
    assert.equal(result.layout.audioOutputs[0].targetTrackNumber,2);
    const preview=JSON.parse(context.exportBackup.getExportSelectionInfo());
    assert.deepEqual(preview.items.filter(item => item.kind === 'audio').map(item => item.trackNumber),[1]);
    const outputs=context.ebBuildRequestedOutputFiles(sequence,'D:/Existing','video.epr','audio.epr','wav',JSON.stringify({includeVideo:true,audioTracks:[1],replaceAudioLayout:true}),true,result.layout);
    assert.equal(outputs.length,2);
    assert.ok(outputs.every(output => /Old_(BACKUP|Track1)/.test(output.finalPath)));
});

test('Re-backup bypasses new destination routing and Backup reports existing files first', async () => {
    const source = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
    const context = vm.createContext({parseHostResult:JSON.parse,path:path.win32,fs:{existsSync:() => true},ensureHostLoaded:async () => true,
        escapeForEvalScript:value=>value,getBackupClipOwners:()=>({}),
        callHost:async () => JSON.stringify({ok:true,layout:{video:{mediaPath:'D:/Existing/Old_BACKUP.mp4'}}}),
        resolveProjectBackupFolder:() => {throw new Error('Destination naming must not run');}});
    vm.runInContext(source.slice(source.indexOf('async function resolveExportActionDestination('),source.indexOf('async function runExport(')),context);
    assert.equal((await context.resolveExportActionDestination(true)).folderPath,'D:/Existing');
    await assert.rejects(context.resolveExportActionDestination(false),/Backup files already exist/);
    context.callHost=async () => JSON.stringify({ok:true,layout:{audioOutputs:[]}});
    await assert.rejects(context.resolveExportActionDestination(true),/No existing backup files/);
});

test('Align Existing bypasses destination naming and passes discovered paths to alignment', async () => {
    const source = fs.readFileSync(path.join(__dirname,'..','js','main.js'),'utf8');
    let aligned;
    const existing = {sequenceName:'No category',layout:{baseName:'old',video:{mediaPath:'D:/Old_BACKUP.mp4',targetTrackNumber:4},audioOutputs:[{mediaPath:'E:/Old_Track1.wav',sourceTrackNumber:1,sourceTrackNumbers:[1]}]}};
    const context=vm.createContext({busy:false,exportFolder:null,alignFolder:null,DEFAULT_BACKUP_VIDEO_TRACK:5,
        stopPendingCleanupRetry:async () => {}, resolveExportActionDestination:async (flag) => {assert.equal(flag,true); return {folderPath:'D:/',existingBackup:existing};},
        resolveProjectBackupFolder:() => {throw new Error('FTP validation must not run');},
        readManifestForSequence:() => null, getPositiveIntValue:() => 5, updateAlignFolder() {},
        document:{getElementById:() => ({})},alert:message => assert.fail(message),setStatus() {},
        runAlignmentFlow:async (folder,options) => {aligned={folder,options};}});
    vm.runInContext(source.slice(source.indexOf('async function alignExistingFolder()'),source.indexOf('document.addEventListener("DOMContentLoaded"')),context);
    await context.alignExistingFolder();
    assert.equal(aligned.folder,'D:/');
    assert.equal(aligned.options.manifestOnly,true);
    assert.deepEqual(Array.from(aligned.options.manifest.expectedFiles,entry => entry.path),['D:/Old_BACKUP.mp4','E:/Old_Track1.wav']);
});

test('URI-encoded Premiere filenames reuse V7 and recognize MP4 audio-only backups', () => {
    const mediaPath = 'D:/Backups/Current Show_BACKUP.mp4';
    const video = makeTimedClip('Current Show_BACKUP.mp4',60,mediaPath);
    const audio = makeTimedClip('Current Show_BACKUP.mp4',60,mediaPath);
    const sequence = makeTimedSequence([makeTrack(0,[audio])],
        Array.from({length:7},(_,i) => makeTrack(0,i === 6 ? [video] : [])));
    const {context} = loadHostLogic(null,{File:function(value) {
        this.fsName = String(value).replace(/\//g,'\\');
        this.name = encodeURIComponent(this.fsName.split('\\').pop());
        this.exists = true;
    }});
    context.app.project.activeSequence = sequence;
    const layout = context.ebResolveExistingBackupLayout(sequence,sequence.name,7);
    assert.equal(layout.baseName,sequence.name);
    assert.equal(layout.video.targetTrackNumber,7);
    assert.equal(layout.backupAudio.targetTrackNumber,1);
    assert.doesNotThrow(() => context.ebValidateBackupTrack(sequence,7,true,layout.baseName));
    assert.equal(sequence.videoTracks[6].clips.numItems,1,'discovery must retain old backup');
    sequence.videoTracks[6] = makeTrack(0);
    const audioOnly = context.ebResolveExistingBackupLayout(sequence,sequence.name,7);
    assert.equal(audioOnly.video.mediaPath,context.ebGetManagedClipFinalMediaPath(audio));
    assert.equal(audioOnly.video.targetTrackNumber,0);
    assert.equal(audioOnly.backupAudio.targetTrackNumber,1);
    sequence.videoTracks[6] = makeTrack(0,[makeTimedClip('Unrelated.mp4',60,'D:/Unrelated.mp4')]);
    assert.throws(() => context.ebValidateBackupTrack(sequence,7,true,sequence.name),/V7 is not empty/);
});

test('mixed source tracks remain selectable, audible, and untouched by backup cleanup', () => {
    const source = makeTimedClip('WOW 1900 The Apology of Socrates_BACKUP.mp4', 7.274,
        'D:/LocalTests/WOW 1900 The Apology of Socrates_BACKUP.mp4');
    const track = makeTrack(0, [source, makeTimedClip('Dialogue.wav', 10)]);
    let removals = 0;
    source.remove = () => { removals++; };
    const sequence = makeTimedSequence([track]);
    const { context } = loadHostLogic();
    context.app.project.activeSequence = sequence;
    context.seq = sequence;
    context.paths = [source.projectItem.getMediaPath()];
    const queue = JSON.parse(context.exportBackup.getExportSelectionInfo());
    assert.deepEqual(queue.items.map(item => item.label), ['Backup MP4', 'Track 1']);
    assert.equal(queue.items[0].selected, true);
    assert.equal(vm.runInContext('ebFindBackupCandidates(seq, seq.name).length', context), 0);
    vm.runInContext('ebApplyManagedTrackMutePolicy(seq, seq.name)', context);
    assert.equal(track.muteValue, 0);
    vm.runInContext('ebRemoveManagedClipsFromAllAudioTracks(seq, seq.name, "backup", 0, paths)', context);
    assert.equal(removals, 0);
});

test('short clips, unknown timing, and stale track labels cannot establish backup ownership', () => {
    const sequence = makeTimedSequence([
        makeTrack(0, [makeTimedClip('Current Show_BACKUP.mp4', 7)]),
        makeTrack(0, [makeTimedClip('Current Show_Track1.wav', 10)]),
        makeTrack(0, [{ projectItem: { name: 'Current Show_BACKUP.mp4' } }]),
        makeTrack(0, [], 'Current Show_BACKUP')
    ]);
    const { context } = loadHostLogic();
    context.app.project.activeSequence = sequence;
    context.seq = sequence;
    assert.equal(vm.runInContext('ebFindBackupCandidates(seq, seq.name).length', context), 0);
    const queue = JSON.parse(context.exportBackup.getExportSelectionInfo());
    assert.deepEqual(queue.items.map(item => item.label), ['Backup MP4', 'Track 1', 'Track 2', 'Track 3']);
});

test('renamed whole-sequence backups are candidates and selected Re-backup preserves their paths', () => {
    const sequence = makeTimedSequence([
        makeTrack(0, [makeTimedClip('Old Show_BACKUP.mp4', 60, 'D:/LocalTests/Old Show_BACKUP.mp4')]),
        makeTrack(0, [makeTimedClip('Old Show_Track1-2.wav', 60, 'D:/LocalTests/Old Show_Track1-2.wav')])
    ], [makeTrack(0, [makeTimedClip('Old Show_BACKUP.mp4', 60, 'D:/LocalTests/Old Show_BACKUP.mp4')]), makeTrack(0)]);
    const { context } = loadHostLogic();
    context.app.project.activeSequence = sequence;
    context.seq = sequence;
    const candidates = JSON.parse(vm.runInContext('JSON.stringify(ebFindBackupCandidates(seq, seq.name))', context));
    assert.equal(candidates.length, 3);
    assert.equal(candidates[0].name, 'Old Show_BACKUP.mp4');
    assert.equal(candidates[0].exactName, false);
    assert.equal(candidates[0].timelineTrack, 'V1');
    const layout = JSON.parse(vm.runInContext('JSON.stringify(ebCaptureRebackupLayout(seq, seq.name, 1))', context));
    assert.match(layout.video.mediaPath, /Old Show_BACKUP\.mp4$/);
    assert.equal(layout.backupAudio.targetTrackNumber, 1);
    assert.deepEqual(layout.audioOutputs[0].sourceTrackNumbers, [1, 2]);
});

test('duration detection handles sequence ticks and an export range shorter than the full sequence', () => {
    const sequence = makeTimedSequence([
        makeTrack(0, [makeTimedClip('Current Show_Track1.wav', 60)]),
        makeTrack(0, [makeTimedClip('Current Show_Track2.wav', 20)])
    ]);
    sequence.getInPoint = () => 10;
    sequence.getOutPoint = () => 30;
    const { context } = loadHostLogic();
    context.seq = sequence;
    assert.equal(vm.runInContext('ebGetSequenceFullDurationSeconds(seq)', context), 60);
    assert.equal(vm.runInContext('ebFindBackupCandidates(seq, seq.name).length', context), 2);
    delete sequence.getInPoint;
    delete sequence.getOutPoint;
    assert.equal(vm.runInContext('ebFindBackupCandidates(seq, seq.name).length', context), 1);
});

test('normal validation warns for project candidates but blocks only real output files', () => {
    const sequence = makeTimedSequence([
        makeTrack(0, [makeTimedClip('Old Show_Track1.wav', 60, 'D:/LocalTests/Old Show_Track1.wav')])
    ]);
    const existingFiles = new Set();
    const { context } = loadHostLogic(undefined, {
        File: function(filePath) {
            this.fsName = String(filePath).replace(/\//g, '\\');
            this.name = this.fsName.split('\\').pop();
            this.exists = /\.epr$/.test(this.fsName) || existingFiles.has(this.fsName);
        }
    });
    context.app.project.activeSequence = sequence;
    const requestedPath = 'D:\\LocalTests\\Current Show_BACKUP.mp4';
    context.app.project.rootItem.children = makeCollection([{
        name: 'Current Show_BACKUP.mp4', getMediaPath() { return requestedPath; }
    }], 'numItems');
    const validate = () => JSON.parse(context.exportBackup.validateBackupExportSettings(
        1, 'D:/LocalTests', 'video.epr', 'mp3.epr', 'wav.epr', 'wav',
        JSON.stringify({ includeVideo: true, audioTracks: [] }), false, false
    ));
    let result = validate();
    assert.equal(result.ok, true);
    assert.equal(result.backupCandidates[0].name, 'Old Show_Track1.wav');
    assert.equal(result.projectReferences[0].path, requestedPath);
    existingFiles.add(requestedPath);
    result = validate();
    assert.equal(result.ok, false);
    assert.equal(result.hasConflicts, true);
    assert.equal(result.conflicts[0].path, requestedPath);
});

test('alignment cleanup removes only a standalone replacement path, preserving other named backups', () => {
    const old = makeTimedClip('Old Show_BACKUP.mp4', 60, 'D:/LocalTests/Old Show_BACKUP.mp4');
    const current = makeTimedClip('Current Show_BACKUP.mp4', 60, 'D:/LocalTests/Current Show_BACKUP.mp4');
    const tracks = [makeTrack(0, [old]), makeTrack(0, [current])];
    old.remove = () => tracks[0].clips.removeItem(old);
    current.remove = () => tracks[1].clips.removeItem(current);
    const sequence = makeTimedSequence(tracks);
    const { context } = loadHostLogic();
    context.app.project.activeSequence = sequence;
    context.seq = sequence;
    context.paths = [current.projectItem.getMediaPath()];
    assert.equal(vm.runInContext('ebRemoveManagedClipsFromAllAudioTracks(seq, seq.name, "backup", 0, paths)', context), 1);
    assert.equal(tracks[0].clips.numItems, 1);
    assert.equal(tracks[1].clips.numItems, 0);
});

test('candidate prompt names the actual media and honors both continue and cancel', async () => {
    const source = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const start = source.indexOf('async function confirmBackupCandidates(');
    const end = source.indexOf('function formatExistingMediaMessage(', start);
    let message = '';
    let answer = false;
    const context = vm.createContext({ showReadablePrompt(options) { message = options.message; return Promise.resolve(answer); } });
    vm.runInContext(source.slice(start, end), context);
    const validation = { backupCandidates: [{ name: 'Old Show_BACKUP.mp4', path: 'D:/LocalTests/Old Show_BACKUP.mp4',
        timelineTrack: 'V2', durationSeconds: 60, exactName: false }] };
    assert.equal(await context.confirmBackupCandidates(validation), false);
    assert.match(message, /Old Show_BACKUP\.mp4/);
    assert.match(message, /Continue with a new Backup/);
    assert.doesNotMatch(message, /Media already exists/);
    answer = true;
    assert.equal(await context.confirmBackupCandidates(validation), true);
    assert.equal(await context.confirmBackupCandidates({}), true);
});

test('Queue Backup Exports has a visual-only toggle and always starts shown', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');

    assert.doesNotMatch(html, /id="toggleQueueBackupSectionButton"/);
    assert.match(html, /id="queueBackupSectionContent" class="stack"/);
    assert.doesNotMatch(html, /id="queueBackupSectionContent"[^>]*is-hidden/);
    assert.match(
        html,
        /<div class="section-label"[^>]*>Tracks to back up<\/div>[\s\S]*<summary>More settings<\/summary>[\s\S]*id="audioFormatWav"/
    );
    assert.match(mainSource, /function toggleQueueBackupSection\(\)/);
    assert.match(
        mainSource,
        /document\.addEventListener\("DOMContentLoaded",[\s\S]*setQueueBackupSectionVisibility\(true\)/
    );
    assert.doesNotMatch(mainSource, /queueBackupSectionVisibleStorage/i);
});

test('Queue Preview keeps unrelated backup-named source tracks selectable', () => {
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
        ['Backup MP4', 'Track 1', 'Track 2', 'Track 3']
    );
});

test('Queue Preview follows shifted source tracks and drops stale backup source mapping', () => {
    const emptyA1 = makeTrack(0, []);
    const shiftedSources = [2, 3, 4, 5].map((trackNumber) =>
        makeTrack(
            0,
            [{ projectItem: { name: `Source A${trackNumber}.wav` } }],
            trackNumber === 5 ? 'Scene_Track4' : `Audio ${trackNumber}`
        )
    );
    const shiftedBackupVideo = makeTrack(0, [{
        projectItem: {
            name: 'Scene_BACKUP.mp4',
            getMediaPath() { return 'E:\\Existing\\Scene_BACKUP.mp4'; }
        }
    }], 'Scene_BACKUP');
    const shiftedOldAudio = makeTrack(0, [{
        projectItem: {
            name: 'Scene_Track1-4.mp3',
            getMediaPath() { return 'E:\\Existing\\Scene_Track1-4.mp3'; }
        }
    }], 'Scene_Track1-4');
    const sequence = {
        name: 'Scene',
        audioTracks: makeCollection([emptyA1].concat(shiftedSources, shiftedBackupVideo, shiftedOldAudio), 'numTracks'),
        videoTracks: makeCollection([], 'numTracks')
    };
    completeBackupFixture(sequence);
    const { context } = loadHostLogic({
        project: {
            activeSequence: sequence,
            rootItem: { type: 2, children: makeCollection([], 'numItems') },
            sequences: makeCollection([sequence], 'numSequences'),
            save() {}
        }
    });

    const result = JSON.parse(context.exportBackup.getExportSelectionInfo());
    assert.equal(result.ok, true);
    assert.deepEqual(
        Array.from(result.items, (item) => item.label),
        ['Backup MP4', 'Track 2', 'Track 3', 'Track 4', 'Track 5']
    );
    assert.equal(result.items[4].trackName, '');
    assert.deepEqual(Array.from(result.audioGroups), []);
});

test('re-backup ignores unrelated names and a track label alone', () => {
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

    assert.equal(result.backupTrack, 0);
    assert.equal(result.otherSequenceInfo.hasBackup, false);
    assert.equal(result.trackNameInfo.hasBackup, false);
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
    completeBackupFixture(sequence);
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
    let removeCount = 0;
    let renameCount = 0;

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
            removeCount += 1;
            return files.delete(this.fsName);
        }

        rename(newName) {
            renameCount += 1;
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
    assert.equal(renameCount, 1);
    assert.equal(removeCount, 0);
    assert.equal(requestedFiles[0].preservedPaths.length, 2);
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
test('Re-backup retries a temporarily locked selected file after 10 and 20 seconds', () => {
    const sourcePath = path.win32.join('D:', 'Backups', 'Scene_Track1.wav');
    const files = new Set([sourcePath]);
    const sleeps = [];
    let renameCount = 0;
    let removeCount = 0;

    class MockFile {
        constructor(filePath) {
            this.fsName = path.win32.normalize(String(filePath || ''));
            this.name = path.win32.basename(this.fsName);
        }

        get exists() {
            return files.has(this.fsName);
        }

        copy(targetPath) {
            if (!files.has(this.fsName)) return false;
            files.add(path.win32.normalize(String(targetPath)));
            return true;
        }

        remove() {
            removeCount += 1;
            return false;
        }

        rename(newName) {
            renameCount += 1;
            if (renameCount < 3 || !files.has(this.fsName)) return false;
            const renamedPath = path.win32.join(path.win32.dirname(this.fsName), String(newName));
            files.delete(this.fsName);
            files.add(renamedPath);
            this.fsName = renamedPath;
            this.name = path.win32.basename(renamedPath);
            return true;
        }
    }

    const { context } = loadHostLogic(undefined, {
        File: MockFile,
        $: { sleep(milliseconds) { sleeps.push(milliseconds); } }
    });
    context.requestedFilesUnderTest = [{
        kind: 'audio',
        path: path.win32.join('D:', 'Backups', 'Scene_Track1_REBKP_TEMP.wav'),
        finalPath: sourcePath,
        sourceMediaPath: sourcePath,
        trackNumber: 1,
        trackNumbers: [1]
    }];

    const result = JSON.parse(vm.runInContext(
        'JSON.stringify(ebPrepareRebackupMediaForDirectExport(requestedFilesUnderTest))',
        context
    ));

    assert.equal(result.preservedCount, 1);
    assert.equal(files.has(sourcePath), false);
    assert.equal(renameCount, 3);
    assert.equal(removeCount, 2);
    assert.deepEqual(sleeps, [4000, 10000, 20000]);
    assert.equal(context.requestedFilesUnderTest[0].preservedPaths.length, 2);
});test('preserved filenames remain recognizable after an interrupted Re-backup', () => {
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
    completeBackupFixture(sequence);
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
    completeBackupFixture(sequence);
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
    assert.equal(result.managedSelection.trackNumbers['2'], undefined);
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

for (const backupTrack of [1, 3, 7]) test(`rebackup keeps V${backupTrack} visible and hides only higher tracks`, () => {
    const video = makeTimedClip('Old Show_BACKUP.mp4',40,'D:/Backups/Old Show_BACKUP.mp4');
    const audio = makeTimedClip('Old Show_BACKUP.mp4',40,'D:/Backups/Old Show_BACKUP.mp4');
    const videoTracks = Array.from({length:9},(_,i) => makeTrack(0,i === backupTrack - 1 ? [video] : []));
    const sequence = makeTimedSequence([makeTrack(0,[makeTimedClip('Dialogue.wav',60)]),makeTrack(0,[audio])],videoTracks);
    video.start.seconds=120; video.end.seconds=160;
    audio.start.seconds=120; audio.end.seconds=160;
    let markedIn=500, markedOut=540;
    sequence.getInPoint=()=>markedIn; sequence.getOutPoint=()=>markedOut;
    sequence.setInPoint=value=>{markedIn=value;}; sequence.setOutPoint=value=>{markedOut=value;};
    const {context} = loadHostLogic();
    context.app.project.activeSequence = sequence;
    context.ebEnsureFolder = () => true;
    context.ebCheckPreset = () => {};
    context.ebRemoveUnusedMedia = () => false;
    context.ebClearAllAudioSoloStates = () => 0;
    context.ebPrepareRebackupMediaForDirectExport = () => ({preservedCount:0});
    let exports = 0;
    context.ebExportSequenceDirect = () => {
        exports++;
        assert.equal(markedIn,120,'rebackup exports the old backup range, not the newly marked teaser');
        assert.equal(markedOut,160);
        assert.deepEqual(videoTracks.map(t => t.muteValue),videoTracks.map((_,i) => i >= backupTrack ? 1 : 0));
        assert.equal(sequence.audioTracks[1].muteValue,1);
        assert.equal(sequence.audioTracks[0].muteValue,0);
        assert.equal(videoTracks[backupTrack - 1].clips.numItems,1,'old clip must remain during rendering');
    };
    const selection = JSON.stringify({includeVideo:true,audioTracks:[]});
    const validation = JSON.parse(context.exportBackup.validateBackupExportSettings(backupTrack,'D:/Backups','video.epr','mp3.epr','wav.epr','wav',selection,true,false));
    assert.equal(validation.ok,true,validation.message);
    const result = JSON.parse(context.exportBackup.runBackupQueue('D:/Backups','video.epr','mp3.epr','wav.epr','wav',backupTrack,false,selection,'premiere',true,false));
    assert.equal(result.ok,true,result.message);
    assert.equal(exports,1);
    assert.equal(markedIn,500); assert.equal(markedOut,540);
    assert.deepEqual(result.queuedFiles[0].exportRange,{startSeconds:120,endSeconds:160});
    assert.equal(result.rebackupLayout.video.targetTrackNumber,backupTrack);
    assert.equal(result.rebackupLayout.backupAudio.targetTrackNumber,2);
    assert.deepEqual(videoTracks.map(t => t.muteValue),[0,0,0,0,0,0,0,0,0]);
    assert.doesNotThrow(() => context.ebValidateBackupTrack(sequence,backupTrack,false,sequence.name));
    markedIn=120;markedOut=160;
    assert.throws(() => context.ebValidateBackupTrack(sequence,backupTrack,false,sequence.name),/not empty/);
    markedIn=500;markedOut=540;
    context.ebExportSequenceDirect = () => {throw new Error('render failed');};
    const failure = JSON.parse(context.exportBackup.runBackupQueue('D:/Backups','video.epr','mp3.epr','wav.epr','wav',backupTrack,false,selection,'premiere',true,false));
    assert.equal(failure.ok,false);
    assert.match(failure.message,/render failed/);
    assert.equal(markedIn,500); assert.equal(markedOut,540);
    assert.equal(videoTracks[backupTrack - 1].clips.numItems,1);
    assert.deepEqual(videoTracks.map(t => t.muteValue),[0,0,0,0,0,0,0,0,0]);
});

test('runBackupQueue restores video visibility when direct export throws', () => {
    const videoTracks = [makeTrack(0), makeTrack(0), makeTrack(1)];
    const sequence = {
        getInPoint:()=>0, getOutPoint:()=>60,
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

test('Re-backup recognizes its V track through a preserved media path', () => {
    const finalPath = 'D:\\Backups\\Scene_BACKUP.mp4';
    const preservedPath = 'D:\\Backups\\Scene_BACKUP_REBKP_OLD_1785000000000_1.mp4';
    const backupClip = {
        projectItem: {
            name: 'Temporarily relinked media',
            getMediaPath() {
                return preservedPath;
            }
        }
    };
    const backupTrack = makeTrack(0, [backupClip]);
    const sequence = {
        videoTracks: makeCollection([
            makeTrack(0),
            makeTrack(0),
            makeTrack(0),
            makeTrack(0),
            makeTrack(0),
            makeTrack(0),
            makeTrack(0),
            backupTrack
        ], 'numTracks'),
        audioTracks: makeCollection([], 'numTracks')
    };
    backupClip.remove = () => backupTrack.clips.removeItem(backupClip);

    const { context } = loadHostLogic({
        project: {
            rootItem: { type: 2, children: makeCollection([], 'numItems') },
            sequences: makeCollection([sequence], 'numSequences'),
            save() {}
        }
    });
    completeBackupFixture(sequence);
    context.sequenceUnderTest = sequence;
    context.targetPaths = [finalPath];

    assert.equal(
        vm.runInContext('ebFindManagedBackupVideoTrackNumber(sequenceUnderTest, "Scene")', context),
        8
    );
    const layout = JSON.parse(vm.runInContext(
        'JSON.stringify(ebCaptureRebackupLayout(sequenceUnderTest, "Scene"))',
        context
    ));
    assert.equal(layout.video.targetTrackNumber, 8);
    assert.equal(layout.video.mediaPath.toLowerCase(), finalPath.toLowerCase());

    const removed = vm.runInContext('ebRemoveProjectClipsByMediaPaths(targetPaths, "video")', context);
    assert.equal(removed, 1);
    assert.equal(backupTrack.clips.numItems, 0);
});
test('Re-backup fallback inspects only the selected occupied video track', () => {
    const oldFinalPath = 'D:\\Backups\\Original Scene_BACKUP.mp4';
    const oldBackupClip = {
        projectItem: {
            name: 'Original Scene_BACKUP.mp4',
            getMediaPath() {
                return oldFinalPath;
            }
        }
    };
    const selectedTrack = makeTrack(0, [oldBackupClip]);
    const sequence = {
        videoTracks: makeCollection([
            makeTrack(0),
            makeTrack(0),
            makeTrack(0),
            makeTrack(0),
            makeTrack(0),
            makeTrack(0),
            makeTrack(0),
            selectedTrack,
            makeTrack(0, [{ projectItem: { name: 'Another Scene_BACKUP.mp4' } }])
        ], 'numTracks'),
        audioTracks: makeCollection([], 'numTracks')
    };
    const { context } = loadHostLogic();
    completeBackupFixture(sequence);
    context.sequenceUnderTest = sequence;

    assert.equal(
        vm.runInContext('ebFindManagedBackupVideoTrackNumber(sequenceUnderTest, "Renamed Scene")', context),
        0
    );
    assert.equal(
        vm.runInContext('ebFindManagedBackupVideoTrackNumber(sequenceUnderTest, "Renamed Scene", 8)', context),
        8
    );
    const layout = JSON.parse(vm.runInContext(
        'JSON.stringify(ebCaptureRebackupLayout(sequenceUnderTest, "Renamed Scene", 8))',
        context
    ));
    assert.equal(layout.video.targetTrackNumber, 8);
    assert.equal(layout.video.mediaPath.toLowerCase(), oldFinalPath.toLowerCase());
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
test('completed Re-backup files are imported before preserved old-file cleanup', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');
    const runStart = hostSource.indexOf('exportBackup.runBackupQueue = function');
    const runEnd = hostSource.indexOf('function ebGetRebackupReleasePaths', runStart);
    const runSource = hostSource.slice(runStart, runEnd);
    const preparePosition = runSource.indexOf('ebPrepareRebackupMediaForDirectExport(requestedFiles)');
    const directPathPosition = hostSource.indexOf('requested.path = ebToFsPath(requested.finalPath)');
    const exportPosition = runSource.indexOf('ebExportSequenceDirect(sequence, videoPath, videoPresetPath, workAreaType)');

    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const flowStart = mainSource.indexOf('async function runAlignmentFlow');
    const flowEnd = mainSource.indexOf('function scheduleExportMonitorTick', flowStart);
    const flowSource = mainSource.slice(flowStart, flowEnd);
    const alignPosition = flowSource.indexOf('const result = await callHost(script)');
    const cleanupPosition = flowSource.indexOf('attemptPostAlignmentCleanup(');
    const finalizeStart = mainSource.indexOf('async function finalizeRebackupFiles');
    const finalizeEnd = mainSource.indexOf('async function cleanupRebackupPreservedFilesAfterAlignment', finalizeStart);
    const finalizeSource = mainSource.slice(finalizeStart, finalizeEnd);

    assert.ok(preparePosition >= 0);
    assert.ok(directPathPosition >= 0);
    assert.ok(exportPosition > preparePosition);
    assert.ok(alignPosition >= 0);
    assert.ok(cleanupPosition > alignPosition);
    assert.doesNotMatch(finalizeSource, /preservedPaths|cleanupLocalFilesBestEffort/);
    assert.match(flowSource, /New files are imported and aligned\. Old cleanup items pending/);
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
    assert.equal(automaticAlignmentCount, 3);
    assert.match(
        mainSource,
        /async function alignExistingFolder\(\)[\s\S]*?runAlignmentFlow\(exportFolder \|\| alignFolder, \{\s*manifest,\s*manifestOnly: true,\s*skipVideo: false,\s*autoTriggered: false/
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
    assert.match(hostSource, /return ebResult\(true, "Some preserved old project items remain for post-import cleanup\."/);
    assert.doesNotMatch(hostSource, /return ebResult\(false, "Premiere Pro still contains old backup project items/);
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
    assert.match(mainSource, /showReadablePrompt\(\{title:'Import not completed',[^\n]+kind:'error'/);
    assert.match(alignmentSource, /showResultPrompt\(\s*successTitle,/);
    assert.doesNotMatch(mainSource, /function showDesktopResultWindow\(/);
});

test('normal Backup blocks disk collisions and requires an empty target', () => {
    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');
    assert.doesNotMatch(hostSource, /Backup files are already there/);
    assert.match(hostSource, /ebValidateBackupTrack\(sequence, backupVideoTrackNumber, false, sequenceBaseName\)/);
    assert.match(hostSource, /backupCandidates: isRebackupMode/);
    assert.doesNotMatch(hostSource, /conflicts = conflicts\.concat\(ebFindExistingProjectConflicts/);
});

test('automatic empty backup track option is visible and unchecked by default', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');

    assert.match(html, /id="autoEmptyBackupTrackCheckbox"/);
    assert.doesNotMatch(html, /id="autoEmptyBackupTrackCheckbox" checked/);
    assert.match(html, />EMPTY TRACK</);
    assert.match(html, /IMPORT TO[\s\S]*id="decrementBackupTrackButton"[\s\S]*id="exportVideoTrackInput"[\s\S]*id="incrementBackupTrackButton"[\s\S]*EMPTY TRACK/);
    assert.match(html, /class="action-line folder-action-line"[\s\S]*id="chooseFolderButton"[\s\S]*id="exportPath"[\s\S]*id="togglePresetSectionButton"[\s\S]*Change Export Presets[\s\S]*id="presetSection"/);
    assert.equal((html.match(/id="togglePresetSectionButton"/g) || []).length, 1);
    assert.match(html, /id="refreshExportSelectionButton"[\s\S]*id="installedVersionText"/);
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

test('local cleanup uses same-process paths and cannot block successful alignment', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const deleteStart = mainSource.indexOf('function normalizeLocalFilePath');
    const deleteEnd = mainSource.indexOf('async function replaceRebackupFile', deleteStart);
    const deleteSource = mainSource.slice(deleteStart, deleteEnd);
    const flowStart = mainSource.indexOf('async function runAlignmentFlow');
    const flowEnd = mainSource.indexOf('function scheduleExportMonitorTick', flowStart);
    const flowSource = mainSource.slice(flowStart, flowEnd);
    const alignPosition = flowSource.indexOf('const result = await callHost(script)');
    const cleanupPosition = flowSource.indexOf('attemptPostAlignmentCleanup(');

    assert.doesNotMatch(html, /id="progressDialog"/);
    assert.match(deleteSource, /function inspectLocalFile/);
    assert.match(deleteSource, /fs\.lstatSync\(normalizedPath\)/);
    assert.match(deleteSource, /fs\.unlinkSync\(before\.path\)/);
    assert.match(deleteSource, /async function cleanupLocalFilesBestEffort/);
    assert.match(deleteSource, /const resolvedRetryDelays = Array\.isArray\(retryDelays\)/);
    assert.match(deleteSource, /function buildDeletePendingPath/);
    assert.match(deleteSource, /_DELETE_PENDING_/);
    assert.match(deleteSource, /fs\.renameSync\(sourcePath, pendingPath\)/);
    assert.doesNotMatch(deleteSource, /powershell\.exe|cmd\.exe|Shell\.Application|InvokeVerb|deleteLocalFileWithExplorer/);
    assert.ok(alignPosition >= 0);
    assert.ok(cleanupPosition > alignPosition);
    assert.match(flowSource, /Backup files were imported and aligned successfully/);
    assert.doesNotMatch(deleteSource, /confirm\(/);
    assert.doesNotMatch(mainSource, /_cleanup/);
    assert.doesNotMatch(mainSource, /deleteLocalFileWithElevatedShell/);
});
test('successful Premiere cleanup without a remaining-path field stops retrying', async () => {
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const retryStart = mainSource.indexOf('async function retryPremierePreservedMediaRelease');
    const retryEnd = mainSource.indexOf('function getUniqueCleanupPaths', retryStart);
    const retrySource = mainSource.slice(retryStart, retryEnd);
    const preservedPath = path.win32.join('D:', 'Backups', 'Scene_REBKP_OLD_1.mp4');
    const context = vm.createContext({
        ensureHostLoaded: async () => true,
        setStatus() {},
        callHost: async () => JSON.stringify({ ok: true, cleanupPending: false }),
        escapeForEvalScript: (value) => value,
        parseHostResult: JSON.parse,
        getPathComparisonKey: (filePath) => String(filePath || '').toLowerCase(),
        getUniqueCleanupPaths: (filePaths) => Array.from(filePaths || [])
    });
    vm.runInContext(retrySource, context);
    context.manifest = {
        expectedFiles: [{ kind: 'video', preservedPaths: [preservedPath] }],
        premiereCleanupPendingPaths: [preservedPath]
    };

    const result = await vm.runInContext('retryPremierePreservedMediaRelease(manifest)', context);

    assert.equal(result.ok, true);
    assert.deepEqual(Array.from(result.remainingProjectPaths), []);
    assert.deepEqual(Array.from(context.manifest.premiereCleanupPendingPaths), []);
});
test('pending cleanup retries full Align Existing every three seconds until verified', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const retryStart = mainSource.indexOf('function getUniqueCleanupPaths');
    const retryEnd = mainSource.indexOf('async function recoverRebackupTempFiles', retryStart);
    const retrySource = mainSource.slice(retryStart, retryEnd);
    const alignmentStart = mainSource.indexOf('async function runAlignmentFlow');
    const alignmentEnd = mainSource.indexOf('function scheduleExportMonitorTick', alignmentStart);
    const alignmentSource = mainSource.slice(alignmentStart, alignmentEnd);

    assert.match(mainSource, /const CLEANUP_RETRY_INTERVAL_MS = 3000/);
    assert.match(retrySource, /state\.timer = setTimeout\([\s\S]*CLEANUP_RETRY_INTERVAL_MS/);
    assert.match(retrySource, /await runAlignmentFlow\(state\.folderPath, \{/);
    assert.match(retrySource, /manifestOnly: !!state\.manifest/);
    assert.match(retrySource, /cleanupRetryContext: retryContext/);
    assert.match(retrySource, /Removing the currently aligned backup clips and importing them again/);
    assert.match(retrySource, /queuePendingCleanupRetry\(state\)/);
    assert.doesNotMatch(retrySource, /alignMappedFiles/);
    assert.match(alignmentSource, /const cleanupRetryContext = settings\.cleanupRetryContext \|\| null/);
    assert.match(alignmentSource, /if \(!cleanupRetryContext\) \{[\s\S]*await startPendingCleanupRetry\(/);
    assert.match(mainSource, /async function alignExistingFolder\(\)[\s\S]*await stopPendingCleanupRetry\(true\);[\s\S]*await runAlignmentFlow/);
    assert.match(html, /id="cleanupRetryPrompt"/);
    assert.match(html, /id="cleanupRetryAlignButton"[^>]*>Align Existing</);
});
test('preserved paths remain retryable until both Premiere and Windows release them', async () => {
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const cleanupStart = mainSource.indexOf('async function cleanupRebackupPreservedFilesAfterAlignment');
    const cleanupEnd = mainSource.indexOf('function replacePathExtension', cleanupStart);
    const cleanupSource = mainSource.slice(cleanupStart, cleanupEnd);
    const preservedPath = path.win32.join('D:', 'Backups', 'Scene_REBKP_OLD_1.mp4');
    const context = vm.createContext({
        cleanupLocalFilesBestEffort: async () => ({ deleted: [preservedPath], pending: [], errors: {} }),
        getManifestPreservedPaths: (manifest) => manifest.expectedFiles[0].preservedPaths.slice(),
        getPathComparisonKey: (filePath) => String(filePath || '').toLowerCase()
    });
    vm.runInContext(cleanupSource, context);
    context.manifest = {
        expectedFiles: [{ preservedPaths: [preservedPath] }]
    };
    context.preservedPath = preservedPath;

    await vm.runInContext(
        'cleanupRebackupPreservedFilesAfterAlignment(manifest, [0], [preservedPath])',
        context
    );
    assert.deepEqual(Array.from(context.manifest.expectedFiles[0].preservedPaths), [preservedPath]);
    assert.equal(context.manifest.rebackupFinalized, false);

    await vm.runInContext(
        'cleanupRebackupPreservedFilesAfterAlignment(manifest, [0], [])',
        context
    );
    assert.deepEqual(Array.from(context.manifest.expectedFiles[0].preservedPaths), []);
    assert.equal(context.manifest.rebackupFinalized, true);
});
test('best-effort old-file cleanup reports EBUSY as pending instead of throwing', async () => {
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const deleteStart = mainSource.indexOf('function normalizeLocalFilePath');
    const deleteEnd = mainSource.indexOf('async function replaceRebackupFile', deleteStart);
    const deleteSource = mainSource.slice(deleteStart, deleteEnd);
    const targetPath = path.win32.join('D:', 'Backups', 'Scene_BACKUP_REBKP_OLD_1.mp4');
    let targetExists = true;
    let failDelete = true;
    const mockFs = {
        lstatSync() {
            if (!targetExists) {
                const error = new Error('File not found');
                error.code = 'ENOENT';
                throw error;
            }
            return { size: 100 };
        },
        chmodSync() {},
        unlinkSync() {
            if (failDelete) {
                const error = new Error('resource busy or locked');
                error.code = 'EBUSY';
                throw error;
            }
            targetExists = false;
        }
    };
    const context = vm.createContext({
        fs: mockFs,
        path,
        setStatus() {},
        delay: async () => {},
        fileExists: () => targetExists,
        getPathComparisonKey: (filePath) => path.resolve(filePath || '').toLowerCase()
    });
    vm.runInContext(deleteSource, context);
    context.cleanupTarget = targetPath;

    const emptyPathResult = await vm.runInContext(
        'cleanupLocalFilesBestEffort([""], "old backup", [0])',
        context
    );
    assert.deepEqual(Array.from(emptyPathResult.pending), []);

    const pendingResult = await vm.runInContext(
        'cleanupLocalFilesBestEffort([cleanupTarget], "old backup")',
        context
    );
    assert.deepEqual(Array.from(pendingResult.deleted), []);
    assert.deepEqual(Array.from(pendingResult.pending), [targetPath]);
    assert.match(pendingResult.errors[path.resolve(targetPath).toLowerCase()], /^EBUSY:/);

    failDelete = false;
    const deletedResult = await vm.runInContext(
        'cleanupLocalFilesBestEffort([cleanupTarget], "old backup")',
        context
    );
    assert.deepEqual(Array.from(deletedResult.deleted), [targetPath]);
    assert.deepEqual(Array.from(deletedResult.pending), []);
});
test('range replacement removes only the matching backup inside the import range', () => {
    const {context}=loadHostLogic();
    const path='D:/Show_BACKUP.mp4';
    const before=makeTimedClip('Show_BACKUP.mp4',60,path);
    const current=makeTimedClip('Show_BACKUP.mp4',180,path); current.start.seconds=120;
    const other=makeTimedClip('Other.wav',180,'D:/Other.wav'); other.start.seconds=120;
    const removed=[];
    [before,current,other].forEach((clip,i)=>clip.remove=()=>removed.push(i));
    const track=makeTrack(0,[before,current,other]);
    context.ebRemoveBackupClipsInRange(track,'Show','backup',0,[path],{startSeconds:120,endSeconds:180});
    assert.deepEqual(removed,[1]);
    assert.equal(context.ebTrackOccupiedInRange(track,{startSeconds:120,endSeconds:180},[path]),true,'unrelated overlapping source must block replacement');
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

test('project-name routing finds a configured category and number anywhere and builds a canonical folder name', () => {
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const start = mainSource.indexOf('function parseBackupProjectName');
    const end = mainSource.indexOf('function findFtpCategoryFolder', start);
    const context = vm.createContext({
        path,
        getConfiguredBackupCategories() {
            return ['BMD INTRO', 'DAILY NEWS SCROLLS', 'BMD', 'WOW', 'PE', 'HL'];
        }
    });
    vm.runInContext(mainSource.slice(start, end), context);

    const valid = vm.runInContext('parseBackupProjectName("WOW 3226 3227 Pierre Gassendi.prproj")', context);
    const multiWord = vm.runInContext('parseBackupProjectName("BMD INTRO 91 Opening Headlines.prproj")', context);
    const missingEpisode = vm.runInContext('parseBackupProjectName("WOW Pierre Gassendi.prproj")', context);
    const numberOnly = vm.runInContext('parseBackupProjectName("PE 3226.prproj")', context);
    const unknownCategory = vm.runInContext('parseBackupProjectName("NEWS 3226.prproj")', context);
    const reversed = vm.runInContext('parseBackupProjectName("3226 WOW Pierre Gassendi.prproj")', context);
    const embedded = vm.runInContext('parseBackupProjectName("Pierre 3226 WOW Gassendi.prproj")', context);

    assert.equal(valid.ok, true);
    assert.equal(valid.category, 'WOW');
    assert.deepEqual(Array.from(valid.episodeNumbers), ['3226', '3227']);
    assert.equal(multiWord.ok, true);
    assert.equal(multiWord.category, 'BMD INTRO');
    assert.equal(missingEpisode.ok, false);
    assert.equal(numberOnly.ok, true);
    assert.equal(unknownCategory.ok, false);
    assert.equal(reversed.ok, true);
    assert.equal(reversed.canonicalFolderName, 'WOW 3226 Pierre Gassendi');
    assert.equal(embedded.ok, true);
    assert.equal(embedded.canonicalFolderName, 'WOW 3226 Pierre Gassendi');
});

test('backup destination UI exposes FTP category routing, project root, and manual path', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');

    assert.match(html, /id="backupDestinationFtp"[^>]*value="ftp"/);
    assert.match(html, /id="backupDestinationProjectRoot"[^>]*value="projectRoot"/);
    assert.match(html, /id="backupDestinationManual"[^>]*value="manual"/);
    assert.match(html, /id="chooseFolderButton"[^>]*hidden>Choose Path<\/button>/);
    assert.match(mainSource, /const FTP_BACKUP_ROOT = "Y:\\\\@ Backup"/);
    assert.match(mainSource, /destinationMode === BACKUP_DESTINATION_MANUAL/);
    assert.match(mainSource, /manualExportFolder = result\.data\[0\]/);
    assert.match(mainSource, /else \{\s*parsedName = parseBackupProjectName/);
    assert.match(mainSource, /normalizeBackupCategoryName\(entry\.name\) === category/);
    assert.match(mainSource, /folderPath = path\.join\(categoryFolder, parsedName\.canonicalFolderName\)/);
    assert.doesNotMatch(html, /id="projectNamePrompt"/);
    assert.match(html, /id="categoryDestinationNote"/);
    assert.doesNotMatch(html, /id="(?:categoryList|categoryNameInput|addCategoryButton|deleteCategoryButton)"/);
    assert.doesNotMatch(mainSource, /"SM URGENT MESSAGES"/);
    assert.doesNotMatch(mainSource, /"TRIBUTE"/);
    assert.doesNotMatch(mainSource, /"QYP"/);
    assert.match(mainSource, /BACKUP_CATEGORIES_STORAGE_KEY/);
    assert.match(mainSource, /function addBackupCategory\(\)/);
    assert.match(mainSource, /function deleteSelectedBackupCategory\(\)/);
    assert.doesNotMatch(html, /id="changeRebackupAudioLayoutCheckbox"/);
    assert.match(mainSource, /replaceAudioLayout: true/);
    assert.match(mainSource, /await syncExportSelectionWithActiveTimeline\(\)/);
    assert.match(mainSource, /function clearAudioMerges\(\)/);
    assert.match(mainSource, /Array\.isArray\(selectionInfo\.audioGroups\)/);
    assert.match(mainSource, /Clear Merges/);
    assert.doesNotMatch(html, /Copy Existing Backups|id="copyExistingBackupsButton"/);
    assert.match(mainSource, /async function copyExistingBackupsToResolvedLocation\(\)/);
    assert.match(mainSource, /exportBackup\.getActiveBackupLayout\(\)/);
    assert.match(mainSource, /findLegacyBackupForActiveSequence\(\\"\\", true\)/);
    assert.doesNotMatch(mainSource, /Choose Folder Containing Existing Backup Files/);
    assert.match(mainSource, /fs\.copyFileSync\(sourcePath, targetPath\)/);
    assert.match(mainSource, /fs\.statSync\(sourcePath\)\.size !== fs\.statSync\(targetPath\)\.size/);
    assert.match(mainSource, /A different backup file with the same name already exists at the destination/);
});

test('MAIN and INTRO projects share a number-only folder under an existing inclusive range container', () => {
    const mainSource = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
    const start = mainSource.indexOf('function getGroupedEpisodeRouting');
    const end = mainSource.indexOf('async function resolveProjectBackupFolder', start);
    const directoryNames = [
        'WOW 3240-3251_Italy_19990522_SM PL',
        'WOW 3252-3260_Other Group',
        'PE 3240-3260_PE Group'
    ];
    const context = vm.createContext({
        path,
        fs: {
            readdirSync() {
                return directoryNames.map((name) => ({ name, isDirectory() { return true; } }));
            }
        }
    });
    vm.runInContext(mainSource.slice(start, end), context);
    context.mainProject = { title: 'MAIN', episodeNumbers: ['3250'] };
    context.introProject = { title: 'INTRO', episodeNumbers: ['3250'] };

    const mainRoute = vm.runInContext('getGroupedEpisodeRouting(mainProject)', context);
    const introRoute = vm.runInContext('getGroupedEpisodeRouting(introProject)', context);
    const container = vm.runInContext(
        'findEpisodeRangeContainer("Y:\\\\@ Backup\\\\@ WOW BACKUP", "WOW", 3250)',
        context
    );

    assert.equal(mainRoute.enabled, true);
    assert.equal(introRoute.enabled, true);
    assert.equal(mainRoute.episodeNumber, introRoute.episodeNumber);
    assert.equal(path.basename(container), 'WOW 3240-3251_Italy_19990522_SM PL');
    assert.match(mainSource, /folderPath = path\.join\(rangeContainer, String\(groupedRouting\.episodeNumber\)\)/);
    assert.doesNotMatch(mainSource, /groupedRouting\.episodeNumber <= 3240/);

    assert.throws(
        () => vm.runInContext('findEpisodeRangeContainer("Y:\\\\@ Backup\\\\@ WOW BACKUP", "WOW", 3300)', context),
        /Please create the group root folder for the WOW project first/
    );
});

test('changed Re-backup merge layout becomes authoritative and identifies old audio as obsolete', () => {
    const { context } = loadHostLogic();
    const sequence = {
        audioTracks: makeCollection([
            makeTrack(0, [{ projectItem: { name: 'Dialogue A1.wav' } }], 'Dialogue 1'),
            makeTrack(0, [{ projectItem: { name: 'Dialogue A2.wav' } }], 'Dialogue 2')
        ], 'numTracks')
    };
    context.sequenceUnderTest = sequence;
    context.layoutUnderTest = {
        audioOutputs: [
            { sourceTrackNumber: 1, sourceTrackNumbers: [1], mediaPath: 'E:\\Backup\\Old_Track1.mp3', currentMediaPath: 'E:\\Backup\\Old_Track1.mp3' },
            { sourceTrackNumber: 2, sourceTrackNumbers: [2], mediaPath: 'E:\\Backup\\Old_Track2.mp3', currentMediaPath: 'E:\\Backup\\Old_Track2.mp3' }
        ]
    };
    context.selectionUnderTest = { audioTracks: [], audioGroups: [[1, 2]], replaceAudioLayout: true };

    const definitions = JSON.parse(vm.runInContext(
        'JSON.stringify(ebBuildSelectedRebackupAudioDefinitions(sequenceUnderTest, "Scene", selectionUnderTest, layoutUnderTest))',
        context
    ));
    assert.deepEqual(definitions.map((entry) => entry.trackNumbers), [[1, 2]]);
    assert.equal(definitions[0].mediaPath, '');

    context.requestedUnderTest = [{ kind: 'audio', finalPath: 'E:\\Backup\\Scene_Track1-2.mp3', sourceMediaPath: '' }];
    const obsolete = JSON.parse(vm.runInContext(
        'JSON.stringify(ebGetObsoleteRebackupAudioFiles(layoutUnderTest, requestedUnderTest, true))',
        context
    ));
    assert.deepEqual(obsolete, ['E:\\Backup\\Old_Track1.mp3', 'E:\\Backup\\Old_Track2.mp3']);
});

test('changed Re-backup source tracks rebuild audio placement without reserving obsolete gaps', () => {
    const { context } = loadHostLogic();
    context.layoutUnderTest = {
        video: { targetTrackNumber: 2 },
        backupAudio: { targetTrackNumber: 6 },
        audioOutputs: [
            { sourceTrackNumber: 1, sourceTrackNumbers: [1], targetTrackNumber: 7 },
            { sourceTrackNumber: 2, sourceTrackNumbers: [2], targetTrackNumber: 8 },
            { sourceTrackNumber: 3, sourceTrackNumbers: [3], targetTrackNumber: 9 },
            { sourceTrackNumber: 4, sourceTrackNumbers: [4], targetTrackNumber: 10 }
        ]
    };
    context.requestedUnderTest = [2, 3, 4, 5].map((trackNumber) => ({
        kind: 'audio',
        trackNumber,
        trackNumbers: [trackNumber]
    }));

    const rebuilt = JSON.parse(vm.runInContext(
        'JSON.stringify(ebBuildRebackupAlignmentLayout(layoutUnderTest, requestedUnderTest, true))',
        context
    ));
    assert.equal(rebuilt.audioLayoutRebuilt, true);
    assert.equal(rebuilt.backupAudio.targetTrackNumber, 6);
    assert.deepEqual(rebuilt.audioOutputs, []);

    const hostSource = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'export.jsx'), 'utf8');
    assert.match(hostSource, /retainedBackupAudioTrackNumber = backupVideoAudioTrackNumber > 0/);
    assert.match(hostSource, /rebackupLayout: rebackupAlignmentLayout/);
});

test('Re-backup compacts an unchanged audio mapping when an old backup layer has a gap', () => {
    const { context } = loadHostLogic();
    context.layoutUnderTest = {
        video: { targetTrackNumber: 2 },
        backupAudio: { targetTrackNumber: 6, startSeconds: 0 },
        audioOutputs: [
            { sourceTrackNumber: 2, sourceTrackNumbers: [2], targetTrackNumber: 7 },
            { sourceTrackNumber: 4, sourceTrackNumbers: [4], targetTrackNumber: 8 },
            { sourceTrackNumber: 5, sourceTrackNumbers: [5], targetTrackNumber: 10 }
        ]
    };
    context.requestedUnderTest = [
        { kind: 'video' },
        { kind: 'audio', trackNumber: 2, trackNumbers: [2] },
        { kind: 'audio', trackNumber: 4, trackNumbers: [4] },
        { kind: 'audio', trackNumber: 5, trackNumbers: [5] }
    ];

    const rebuilt = JSON.parse(vm.runInContext(
        'JSON.stringify(ebBuildRebackupAlignmentLayout(layoutUnderTest, requestedUnderTest, true))',
        context
    ));
    assert.equal(rebuilt.audioLayoutRebuilt, true);
    assert.equal(rebuilt.backupAudio.targetTrackNumber, 6);
    assert.deepEqual(rebuilt.audioOutputs, []);
});

test('Align Existing legacy rename requires backup filename rules and matching sequence duration', () => {
    const { context } = loadHostLogic();
    const oldVideoPath = 'E:\\Backups\\Old Name_BACKUP.mp4';
    const oldAudioPath = 'E:\\Backups\\Old Name_Track1-2.mp3';
    const sequence = {
        name: 'New Name',
        getInPoint() { return 0; },
        getOutPoint() { return 60; },
        videoTracks: makeCollection([
            makeTrack(0, [{
                start: { seconds: 0 }, end: { seconds: 60 },
                projectItem: { name: 'Old Name_BACKUP.mp4', getMediaPath() { return oldVideoPath; } }
            }])
        ], 'numTracks'),
        audioTracks: makeCollection([
            makeTrack(0, [{
                start: { seconds: 0 }, end: { seconds: 60 },
                projectItem: { name: 'Old Name_Track1-2.mp3', getMediaPath() { return oldAudioPath; } }
            }])
        ], 'numTracks')
    };
    context.app.project.activeSequence = sequence;

    const matched = JSON.parse(vm.runInContext(
        'exportBackup.findLegacyBackupForActiveSequence("E:\\\\Backups")',
        context
    ));
    assert.equal(matched.ok, true);
    assert.equal(matched.found, true);
    assert.equal(matched.candidate.oldBase, 'Old Name');
    assert.equal(matched.candidate.files.length, 2);

    sequence.videoTracks[0].clips[0].end.seconds = 20;
    sequence.audioTracks[0].clips[0].end.seconds = 20;
    const rejected = JSON.parse(vm.runInContext(
        'exportBackup.findLegacyBackupForActiveSequence("E:\\\\Backups")',
        context
    ));
    assert.equal(rejected.found, false);
});
