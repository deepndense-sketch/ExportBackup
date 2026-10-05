const csInterface = new CSInterface();
const fs = require("fs");
const os = require("os");
const path = require("path");
const childProcess = require("child_process");
const https = require("https");

const VIDEO_PRESET_STORAGE_KEY = "exportbackup.videoPresetPath";
const MP3_PRESET_STORAGE_KEY = "exportbackup.mp3PresetPath";
const WAV_PRESET_STORAGE_KEY = "exportbackup.wavPresetPath";
const EXPORT_FOLDER_STORAGE_KEY = "exportbackup.exportFolder";
const ALIGN_FOLDER_STORAGE_KEY = "exportbackup.alignFolder";
const PRESET_SECTION_VISIBLE_STORAGE_KEY = "exportbackup.presetSectionVisible";
const BACKUP_VIDEO_TRACK_STORAGE_KEY = "exportbackup.backupVideoTrack";
const ALIGN_VIDEO_TRACK_STORAGE_KEY = "exportbackup.alignVideoTrack";
const ALIGN_SORT_PROJECT_FILES_STORAGE_KEY = "exportbackup.alignSortProjectFiles";
const AUDIO_FORMAT_STORAGE_KEY = "exportbackup.audioFormat";
const REMOVE_SEQUENCE_MARKERS_STORAGE_KEY = "exportbackup.removeSequenceMarkers";
const EXPORT_MODE_STORAGE_KEY = "exportbackup.exportMode";
const COPY_PROJECT_FILE_STORAGE_KEY = "exportbackup.copyProjectFile";
const BACKUP_DESTINATION_STORAGE_KEY = "exportbackup.backupDestination";
const BACKUP_CATEGORIES_STORAGE_KEY = "exportbackup.backupCategories";

const DEFAULT_BACKUP_VIDEO_TRACK = 5;
const EXPORT_MODE_PREMIERE = "premiere";
const EXPORT_MODE_MEDIA_ENCODER = "mediaEncoder";
const EXPORT_MANIFEST_SUFFIX = "_ExportBackupMap.json";
const REBACKUP_TEMP_MARKER = "_REBKP_TEMP";
const REBACKUP_OLD_MARKER = "_REBKP_OLD_";
const EXPORT_MONITOR_INTERVAL_MS = 5000;
const EXPORT_MONITOR_STABLE_PASSES = 2;
const EXPORT_MONITOR_TIMEOUT_MS = 6 * 60 * 60 * 1000;
const REBACKUP_RECOVERY_STABLE_WAIT_MS = 2000;
const CLEANUP_RETRY_INTERVAL_MS = 3000;
const BACKUP_DESTINATION_FTP = "ftp";
const BACKUP_DESTINATION_PROJECT_ROOT = "projectRoot";
const BACKUP_DESTINATION_PARENT_FOLDER = "parentFolder";
const BACKUP_DESTINATION_MANUAL = "manual";
const FTP_BACKUP_ROOT = "Y:\\@ Backup";
const BACKUP_CATEGORIES = [
    "BMD INTRO", "DAILY NEWS SCROLLS", "CTAW", "GPGW", "SHOW",
    "BMD", "BRE", "GAT", "GOL", "MOS", "NWN", "PCC", "SWA", "VEG", "WAU",
    "WOW", "AR", "AP", "AW", "CS", "EB", "GG", "HL", "KW",
    "LS", "NB", "PE", "SS", "UL", "VE", "VR"
];

let exportFolder = null;
let manualExportFolder = null;
let alignFolder = null;
let hostLoaded = false;
let busy = false;
let videoPresetPath = "";
let mp3PresetPath = "";
let wavPresetPath = "";
let localVersion = "unknown";
let localVersionNotes = "";
let remoteVersion = null;
let remoteVersionNotes = "";
let presetSectionVisible = true;
let queueBackupSectionVisible = true;
let exportMonitorState = null;
let exportSelectionState = null;
let mergedAudioGroups = [];
let nextMergedAudioGroupId = 1;
let inOutPromptResolver = null;
let cleanupRetryState = null;
let configuredBackupCategories = BACKUP_CATEGORIES.slice();

function getExtensionRootPath() {
    try {
        return csInterface.getSystemPath(SystemPath.EXTENSION);
    } catch (error) {
        return path.basename(__dirname).toLowerCase() === "js" ? path.dirname(__dirname) : __dirname;
    }
}

function getVersionFilePath() {
    return path.join(getExtensionRootPath(), "version.json");
}

function getUpdateScriptPath() {
    return path.join(getExtensionRootPath(), "update_from_github.ps1");
}

function getBundledPresetPath(fileName) {
    return path.join(getExtensionRootPath(), "presets", fileName);
}

function getBundledPresetFolderPath() {
    return path.join(getExtensionRootPath(), "presets");
}

function getDefaultVideoPresetPath() {
    return getBundledPresetPath("1080 AIR.epr");
}

function getDefaultMp3PresetPath() {
    return getBundledPresetPath("mp3.epr");
}

function getDefaultWavPresetPath() {
    return getBundledPresetPath("wav.epr");
}

function setStatus(message, kind) {
    const statusBox = document.getElementById("statusBox");
    if (!statusBox) {
        return;
    }

    statusBox.textContent = message;
    statusBox.classList.toggle("is-success", kind === "success");
    statusBox.classList.toggle("is-error", kind === "error");
}

function getCompletionStatusTitle(settings, manifest) {
    if (settings && settings.autoTriggered) {
        return manifest && manifest.rebackup
            ? "Re-backup done without error."
            : "Backup done without error.";
    }

    return "ALIGNMENT DONE";
}

function setPresetSectionVisibility(visible) {
    presetSectionVisible = visible;

    const presetSection = document.getElementById("presetSection");
    const toggleButton = document.getElementById("togglePresetSectionButton");
    if (!presetSection || !toggleButton) {
        return;
    }

    presetSection.classList.toggle("is-hidden", !visible);
    toggleButton.textContent = visible ? "Hide Export Presets" : "Change Export Presets";

    try {
        localStorage.setItem(PRESET_SECTION_VISIBLE_STORAGE_KEY, visible ? "true" : "false");
    } catch (error) {}
}

function togglePresetSection() {
    setPresetSectionVisibility(!presetSectionVisible);
}

function setQueueBackupSectionVisibility(visible) {
    queueBackupSectionVisible = visible;

    const content = document.getElementById("queueBackupSectionContent");
    const toggleButton = document.getElementById("toggleQueueBackupSectionButton");
    if (!content || !toggleButton) {
        return;
    }

    content.classList.toggle("is-hidden", !visible);
    toggleButton.textContent = visible ? "Hide" : "Show";
    toggleButton.setAttribute("aria-expanded", visible ? "true" : "false");
}

function toggleQueueBackupSection() {
    setQueueBackupSectionVisibility(!queueBackupSectionVisible);
}

function getAudioFormatInputs() {
    return {
        mp3: document.getElementById("audioFormatMp3"),
        wav: document.getElementById("audioFormatWav")
    };
}

function getBackupVideoTrackInput() {
    return document.getElementById("exportVideoTrackInput");
}

function getAutoEmptyBackupTrackCheckbox() {
    return document.getElementById("autoEmptyBackupTrackCheckbox");
}

function getAlignVideoTrackInput() {
    return document.getElementById("exportVideoTrackInput");
}

function getRemoveSequenceMarkersCheckbox() {
    return document.getElementById("removeSequenceMarkersCheckbox");
}

function getCopyProjectFileCheckbox() {
    return document.getElementById("copyProjectFileCheckbox");
}

function getBackupDestinationInputs() {
    return {
        ftp: document.getElementById("backupDestinationFtp"),
        projectRoot: document.getElementById("backupDestinationProjectRoot"),
        parentFolder: document.getElementById("backupDestinationParentFolder"),
        manual: document.getElementById("backupDestinationManual")
    };
}

function getSelectedBackupDestination() {
    const inputs = getBackupDestinationInputs();
    if (inputs.parentFolder && inputs.parentFolder.checked) {
        return BACKUP_DESTINATION_PARENT_FOLDER;
    }
    if (inputs.projectRoot && inputs.projectRoot.checked) {
        return BACKUP_DESTINATION_PROJECT_ROOT;
    }
    if (inputs.manual && inputs.manual.checked) {
        return BACKUP_DESTINATION_MANUAL;
    }
    return BACKUP_DESTINATION_FTP;
}

function updateDestinationButtonLabel() {
    const selector = document.getElementById("backupDestinationSelect");
    if (selector) selector.value = getSelectedBackupDestination();
    const button = document.getElementById("chooseFolderButton");
    if (!button) return;
    const mode = getSelectedBackupDestination();
    button.hidden = mode !== BACKUP_DESTINATION_MANUAL;
    button.textContent = mode === BACKUP_DESTINATION_MANUAL ? "Choose Path"
        : mode === BACKUP_DESTINATION_PROJECT_ROOT ? "Project Root" : "FTP";
}

function getExportModeInputs() {
    return {
        premiere: document.getElementById("exportModePremiere"),
        mediaEncoder: document.getElementById("exportModeMediaEncoder")
    };
}

function setBusyState(nextBusy) {
    const audioInputs = getAudioFormatInputs();
    const setDisabled = (id, disabled) => {
        const element = document.getElementById(id);
        if (element) {
            element.disabled = disabled;
        }
    };

    busy = nextBusy;
    setDisabled("chooseFolderButton", nextBusy);
    setDisabled("backupDestinationSelect", nextBusy);
    setDisabled("copyExistingBackupsButton", nextBusy);
    setDisabled("chooseVideoPresetButton", nextBusy);
    setDisabled("chooseMp3PresetButton", nextBusy);
    setDisabled("chooseWavPresetButton", nextBusy);
    setDisabled("exportButton", nextBusy);
    setDisabled("rebackupButton", nextBusy);
    setDisabled("chooseAlignFolderButton", nextBusy);
    setDisabled("alignFolderButton", nextBusy);
    setDisabled("cleanupRetryAlignButton", nextBusy);
    setDisabled("alignSkipVideoCheckbox", nextBusy);
    setDisabled("alignSortProjectFilesCheckbox", nextBusy);
    setDisabled("refreshExportSelectionButton", nextBusy);
    setDisabled("autoEmptyBackupTrackCheckbox", nextBusy);
    setDisabled("decrementBackupTrackButton", nextBusy);
    setDisabled("incrementBackupTrackButton", nextBusy);
    setDisabled("mergeAudioButton", nextBusy);
    setDisabled("clearAudioMergesButton", nextBusy);
    setDisabled("categoryNameInput", nextBusy);
    setDisabled("addCategoryButton", nextBusy);
    setDisabled("deleteCategoryButton", nextBusy);
    setDisabled("categoryList", nextBusy);

    const destinationInputs = getBackupDestinationInputs();
    Object.keys(destinationInputs).forEach((key) => {
        if (destinationInputs[key]) {
            destinationInputs[key].disabled = nextBusy;
        }
    });

    const exportModeInputs = getExportModeInputs();
    Object.keys(exportModeInputs).forEach((key) => {
        if (exportModeInputs[key]) {
            exportModeInputs[key].disabled = nextBusy;
        }
    });

    Object.keys(audioInputs).forEach((key) => {
        if (audioInputs[key]) {
            audioInputs[key].disabled = nextBusy;
        }
    });

    if (getBackupVideoTrackInput()) {
        getBackupVideoTrackInput().disabled = nextBusy;
    }

    if (getAlignVideoTrackInput()) {
        getAlignVideoTrackInput().disabled = nextBusy;
    }

    if (getRemoveSequenceMarkersCheckbox()) {
        getRemoveSequenceMarkersCheckbox().disabled = nextBusy;
    }

    if (getCopyProjectFileCheckbox()) {
        getCopyProjectFileCheckbox().disabled = nextBusy;
    }

    if (!nextBusy) {
        setSelectedExportMode(getSelectedExportMode());
        setSelectedAudioFormat(getSelectedAudioFormat());
        const autoEmptyCheckbox = getAutoEmptyBackupTrackCheckbox();
        setAutoEmptyBackupTrackEnabled(autoEmptyCheckbox && autoEmptyCheckbox.checked);
    }
}

function setUpdateButton(label, isUpdateAvailable, hoverText) {
    const versionText = document.getElementById("installedVersionText");
    versionText.textContent = `Version ${localVersion}`;
    versionText.title = localVersionNotes || "";
    const notice = document.getElementById('updateNotice');
    if (notice) {
        notice.hidden = !isUpdateAvailable;
        document.getElementById('updateNoticeText').textContent = `Backup Project ${remoteVersion} is available`;
    }
}

let promptedUpdateVersion = '';
let updatePromptPreviousFocus = null;

function isUpdateWorkBusy() {
    if (busy) return true;
    const collector = document.getElementById('collectorPanel');
    try { return !!(collector && collector.contentWindow && collector.contentWindow.isCollectorBusy && collector.contentWindow.isCollectorBusy()); }
    catch (error) { return true; }
}

function showUpdatePrompt(automatic) {
    if (!remoteVersion || compareVersions(remoteVersion, localVersion) <= 0 || isUpdateWorkBusy()) return;
    if (automatic && promptedUpdateVersion === remoteVersion) return;
    const prompt = document.getElementById('updatePrompt');
    if (!prompt || !prompt.classList.contains('is-hidden')) return;
    promptedUpdateVersion = remoteVersion;
    updatePromptPreviousFocus = document.activeElement;
    document.getElementById('updatePromptTitle').textContent = `Backup Project ${remoteVersion} is available`;
    const notes = document.getElementById('updatePromptNotes');
    notes.textContent = String(remoteVersionNotes || '').trim() || 'No release notes were provided for this version.';
    notes.scrollTop = 0;
    prompt.classList.remove('is-hidden');
    document.getElementById('updateLaterButton').focus();
}

function closeUpdatePrompt() {
    document.getElementById('updatePrompt').classList.add('is-hidden');
    if (updatePromptPreviousFocus && updatePromptPreviousFocus.focus) updatePromptPreviousFocus.focus();
}

function installPromptedUpdate() {
    if (isUpdateWorkBusy()) return;
    closeUpdatePrompt();
    runGithubUpdate();
}

function handleUpdatePromptKey(event) {
    if (event.key === 'Escape') { event.preventDefault(); closeUpdatePrompt(); }
    if (event.key === 'Tab') {
        event.preventDefault();
        const later = document.getElementById('updateLaterButton');
        const install = document.getElementById('updateInstallButton');
        const focusable = [document.getElementById('updatePromptNotes'), later, install].filter(Boolean);
        const index = focusable.indexOf(document.activeElement);
        focusable[(index + (event.shiftKey ? focusable.length - 1 : 1)) % focusable.length].focus();
    }
}

function escapeForEvalScript(value) {
    return String(value)
        .replace(/\\/g, "\\\\")
        .replace(/"/g, '\\"')
        .replace(/\r/g, "\\r")
        .replace(/\n/g, "\\n");
}

function escapeHtml(value) {
    return String(value || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function callHost(script) {
    return new Promise((resolve) => {
        csInterface.evalScript(script, (result) => resolve(result));
    });
}

async function ensureHostLoaded() {
    if (hostLoaded) {
        return true;
    }

    const extensionPath = csInterface.getSystemPath(SystemPath.EXTENSION).replace(/\\/g, "/");
    const hostPath = `${extensionPath}/jsx/export.jsx`;
    const result = await callHost(`$.evalFile("${escapeForEvalScript(hostPath)}")`);

    if (result === "EvalScript error." || result === "false") {
        setStatus(`Could not load host script.\n${result}`);
        return false;
    }

    hostLoaded = true;
    return true;
}

function parseHostResult(raw) {
    try {
        return JSON.parse(raw);
    } catch (error) {
        return null;
    }
}

function isLockedFileRecoveryMessage(message) {
    return String(message || "").indexOf("File still busy") >= 0 &&
        String(message || "").indexOf("Align Existing") >= 0;
}

let readablePromptQueue = Promise.resolve();
function showReadablePrompt(options) {
    const pending = readablePromptQueue.then(() => new Promise(resolve => {
        const dialog = document.getElementById('readablePrompt');
        const body = document.getElementById('readablePromptMessage');
        const yes = document.getElementById('readablePromptOK');
        const no = document.getElementById('readablePromptCancel');
        const successMark = document.getElementById('readablePromptSuccess');
        const detailsButton = document.getElementById('readablePromptDetails');
        const kind = options.kind || (options.success === true ? 'success' : '');
        const compactSuccess = kind === 'success';
        const previous = document.activeElement;
        successMark.hidden = !kind;
        successMark.className = 'success-mark ' + kind;
        successMark.textContent = kind === 'success' ? '✓' : (kind === 'error' ? '×' : '!');
        detailsButton.hidden = !compactSuccess;
        body.hidden = compactSuccess;
        detailsButton.textContent = 'Show details';
        detailsButton.setAttribute('aria-expanded', 'false');
        detailsButton.onclick = () => {
            body.hidden = !body.hidden;
            detailsButton.textContent = body.hidden ? 'Show details' : 'Hide details';
            detailsButton.setAttribute('aria-expanded', String(!body.hidden));
        };
        document.getElementById('readablePromptTitle').textContent = options.title || 'Backup Project';
        body.textContent = String(options.message || '');
        yes.textContent = options.confirmText || 'OK';
        no.hidden = !options.cancelText;
        no.textContent = options.cancelText || 'Cancel';
        const close = accepted => {
            dialog.classList.add('is-hidden');
            dialog.removeEventListener('keydown', keydown);
            yes.onclick = no.onclick = detailsButton.onclick = null;
            if (previous && previous.focus) previous.focus();
            resolve(accepted);
        };
        const keydown = event => {
            if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(!options.cancelText); }
            if (event.key === 'Tab') {
                const focusable = [detailsButton, body, no, yes].filter(element => !element.hidden);
                const index = focusable.indexOf(document.activeElement);
                event.preventDefault();
                focusable[(index + (event.shiftKey ? focusable.length - 1 : 1)) % focusable.length].focus();
            }
        };
        yes.onclick = () => close(true);
        no.onclick = () => close(false);
        dialog.addEventListener('keydown', keydown);
        dialog.classList.remove('is-hidden');
        body.scrollTop = 0;
        (options.cancelText ? no : yes).focus();
    }));
    readablePromptQueue = pending.catch(() => {});
    return pending;
}

function showBlockingMessage(message, kind) {
    return showReadablePrompt({message, kind:kind || 'error', title:kind === 'success' ? 'Completed successfully' : 'Could not complete'});
}

const ALIGNMENT_RECOVERY_TEXT = "Please use Align Existing to import the file.";

function showResultPrompt(title, message, options) {
    if (options && options.success === true) {
        return showReadablePrompt({title: String(title || 'Completed successfully'), message, success:true});
    }
    if (options) {
        return showReadablePrompt({title: 'Completed with warnings', message, kind:'warning'});
    }
    const lines = [String(title || "Done")];
    if (message) {
        lines.push(String(message));
    }
    showBlockingMessage(lines.join("\n\n"));
}

function showCleanupRetryPrompt(title, message, kind) {
    const prompt = document.getElementById("cleanupRetryPrompt");
    const panel = document.getElementById("cleanupRetryPanel");
    const titleElement = document.getElementById("cleanupRetryTitle");
    const messageElement = document.getElementById("cleanupRetryMessage");
    const alignButton = document.getElementById("cleanupRetryAlignButton");
    if (!prompt || !panel || !titleElement || !messageElement || !alignButton) {
        return;
    }

    titleElement.textContent = (kind === 'success' ? '✓ ' : kind === 'error' ? '× ' : '! ') + String(title || "Cleanup in progress");
    messageElement.textContent = String(message || "");
    panel.classList.toggle("is-success", kind === "success");
    panel.classList.toggle("is-error", kind === "error");
    alignButton.classList.toggle("is-hidden", kind === "success");
    prompt.classList.remove("is-hidden");
}

function closeCleanupRetryPrompt() {
    const prompt = document.getElementById("cleanupRetryPrompt");
    if (prompt) {
        prompt.classList.add("is-hidden");
    }
}

function dismissCleanupRetryPrompt() {
    if (cleanupRetryState) {
        cleanupRetryState.dialogDismissed = true;
    }
    closeCleanupRetryPrompt();
}

function bindCleanupRetryPrompt() {
    const prompt = document.getElementById("cleanupRetryPrompt");
    const closeButton = document.getElementById("cleanupRetryCloseButton");
    const alignButton = document.getElementById("cleanupRetryAlignButton");
    if (!prompt || !closeButton || !alignButton) {
        return;
    }

    closeButton.addEventListener("click", dismissCleanupRetryPrompt);
    prompt.addEventListener("keydown", (event) => {
        if (event.key === "Escape") {
            event.preventDefault();
            dismissCleanupRetryPrompt();
        }
    });
    alignButton.addEventListener("click", async () => {
        alignButton.disabled = true;
        try {
            await alignExistingFolder();
        } finally {
            alignButton.disabled = false;
        }
    });
}
function closeInOutPrompt(shouldAutoSet) {
    const prompt = document.getElementById("inOutPrompt");
    const checkbox = document.getElementById("autoSetInOutCheckbox");
    const resolve = inOutPromptResolver;
    const confirmed = shouldAutoSet === true && !!(checkbox && checkbox.checked);

    inOutPromptResolver = null;
    if (prompt) {
        prompt.classList.add("is-hidden");
    }
    if (resolve) {
        resolve(confirmed);
    }
}

function showInOutPrompt() {
    const prompt = document.getElementById("inOutPrompt");
    const checkbox = document.getElementById("autoSetInOutCheckbox");
    const okButton = document.getElementById("inOutPromptOkButton");

    if (!prompt || !checkbox || !okButton) {
        return Promise.resolve(false);
    }

    if (inOutPromptResolver) {
        closeInOutPrompt(false);
    }

    checkbox.checked = true;
    okButton.disabled = false;
    prompt.classList.remove("is-hidden");
    setTimeout(() => okButton.focus(), 0);

    return new Promise((resolve) => {
        inOutPromptResolver = resolve;
    });
}

function bindInOutPrompt() {
    const prompt = document.getElementById("inOutPrompt");
    const checkbox = document.getElementById("autoSetInOutCheckbox");
    const okButton = document.getElementById("inOutPromptOkButton");
    const cancelButton = document.getElementById("inOutPromptCancelButton");
    if (!prompt || !checkbox || !okButton || !cancelButton) {
        return;
    }

    checkbox.addEventListener("change", () => {
        okButton.disabled = !checkbox.checked;
    });
    okButton.addEventListener("click", () => closeInOutPrompt(true));
    cancelButton.addEventListener("click", () => closeInOutPrompt(false));
    prompt.addEventListener("keydown", (event) => {
        if (event.key === "Escape") {
            event.preventDefault();
            closeInOutPrompt(false);
        }
    });
}

function showAlignmentRecoveryError(details) {
    const lines = ["IMPORT NOT COMPLETED", ALIGNMENT_RECOVERY_TEXT];
    if (details) {
        lines.push("", details);
    }
    setStatus(lines.join("\n"), "error");
    showReadablePrompt({title:'Import not completed', message:lines.slice(1).join('\n'), kind:'error'});
}

async function confirmBackupCandidates(validation) {
    const candidates = Array.isArray(validation.backupCandidates) ? validation.backupCandidates : [];
    const references = Array.isArray(validation.projectReferences) ? validation.projectReferences : [];
    if (!candidates.length && !references.length) {
        return true;
    }
    const lines = ["Possible existing backups in Premiere:"];
    candidates.forEach((candidate) => {
        lines.push("", `${candidate.timelineTrack}: ${candidate.name}`);
        if (candidate.path) lines.push(candidate.path);
        lines.push(`Standalone clip, ${Number(candidate.durationSeconds).toFixed(2)} seconds; matches the sequence or export range.`);
        if (!candidate.exactName) lines.push("Different name: this may be a backup from before the sequence was renamed, or source media.");
    });
    references.forEach((reference) => {
        lines.push("", `Premiere references this output path, but no file was found there:\n${reference.path}`);
    });
    lines.push("", "Continue with a new Backup? Existing candidate clips will be kept. Cancel to review them or use Re-backup.");
    return showReadablePrompt({title:'Possible existing backups', message:lines.join('\n'), kind:'warning', confirmText:'Continue Backup', cancelText:'Cancel'});
}

function formatExistingMediaMessage(validation) {
    const conflicts = Array.isArray(validation && validation.conflicts) ? validation.conflicts : [];
    const seen = {};
    const paths = [];

    conflicts.forEach((conflict) => {
        const conflictPath = conflict && conflict.path ? String(conflict.path) : "";
        if (!conflictPath || seen[conflictPath]) {
            return;
        }

        seen[conflictPath] = true;
        paths.push(conflictPath);
    });

    if (!paths.length) {
        return validation && validation.message
            ? `${validation.message}\nUse Re-backup to replace it.`
            : "Media already exists. Use Re-backup to replace it.";
    }

    return `Media already exists. Use Re-backup to replace it.\n\nPath:\n${paths.join("\n")}`;
}

function fileExists(filePath) {
    try {
        return !!filePath && fs.existsSync(filePath);
    } catch (error) {
        return false;
    }
}

function readJsonFile(filePath) {
    try {
        return JSON.parse(fs.readFileSync(filePath, "utf8"));
    } catch (error) {
        return null;
    }
}

function getPositiveIntValue(elementId, fallbackValue) {
    const element = document.getElementById(elementId);
    if (!element) {
        return fallbackValue;
    }

    const parsed = parseInt(element.value, 10);
    if (!parsed || parsed < 1) {
        return fallbackValue;
    }

    return parsed;
}

function sanitizeSequenceName(value) {
    return String(value || "Active_Sequence")
        .replace(/[\\\/:\*\?"<>\|]/g, "_")
        .trim() || "Active_Sequence";
}

function getManifestPath(folderPath, baseName) {
    return path.join(folderPath, `${sanitizeSequenceName(baseName)}${EXPORT_MANIFEST_SUFFIX}`);
}

function getSelectedAudioFormat() {
    const audioInputs = getAudioFormatInputs();
    return audioInputs.wav && audioInputs.wav.checked ? "wav" : "mp3";
}

function normalizeExportMode(mode) {
    return String(mode || "").toLowerCase() === EXPORT_MODE_MEDIA_ENCODER.toLowerCase()
        ? EXPORT_MODE_MEDIA_ENCODER
        : EXPORT_MODE_PREMIERE;
}

function getSelectedExportMode() {
    const inputs = getExportModeInputs();
    return inputs.mediaEncoder && inputs.mediaEncoder.checked ? EXPORT_MODE_MEDIA_ENCODER : EXPORT_MODE_PREMIERE;
}

function setSelectedExportMode(mode) {
    const inputs = getExportModeInputs();
    const resolved = normalizeExportMode(mode);

    if (inputs.premiere) {
        inputs.premiere.checked = resolved === EXPORT_MODE_PREMIERE;
    }

    if (inputs.mediaEncoder) {
        inputs.mediaEncoder.checked = resolved === EXPORT_MODE_MEDIA_ENCODER;
    }
}

function saveSelectedExportMode(mode) {
    const resolved = normalizeExportMode(mode);
    setSelectedExportMode(resolved);

    try {
        localStorage.setItem(EXPORT_MODE_STORAGE_KEY, resolved);
    } catch (error) {}
}

function setSelectedAudioFormat(format) {
    const audioInputs = getAudioFormatInputs();
    const resolved = String(format || "").toLowerCase() === "wav" ? "wav" : "mp3";

    if (audioInputs.mp3) {
        audioInputs.mp3.checked = resolved === "mp3";
    }

    if (audioInputs.wav) {
        audioInputs.wav.checked = resolved === "wav";
    }
}

function saveSelectedAudioFormat(format) {
    setSelectedAudioFormat(format);

    try {
        localStorage.setItem(AUDIO_FORMAT_STORAGE_KEY, getSelectedAudioFormat());
    } catch (error) {}
}

function loadSavedUiState() {
    try {
        const saved = localStorage.getItem(PRESET_SECTION_VISIBLE_STORAGE_KEY);
        if (saved === "true") {
            presetSectionVisible = true;
            return;
        }
    } catch (error) {}

    presetSectionVisible = false;
}

function useAutoEmptyBackupTrack() {
    const checkbox = getAutoEmptyBackupTrackCheckbox();
    return !!(checkbox && checkbox.checked);
}

function saveBackupVideoTrack(trackNumber) {
    try {
        localStorage.setItem(BACKUP_VIDEO_TRACK_STORAGE_KEY, String(trackNumber));
    } catch (error) {}
}

function saveAlignVideoTrack(trackNumber) {
    try {
        localStorage.setItem(ALIGN_VIDEO_TRACK_STORAGE_KEY, String(trackNumber));
    } catch (error) {}
}

function saveRemoveSequenceMarkers(enabled) {
    try {
        localStorage.setItem(REMOVE_SEQUENCE_MARKERS_STORAGE_KEY, enabled ? "true" : "false");
    } catch (error) {}
}

function applyBackupDefaults(defaults, force) {
    const backupTrackInput = getBackupVideoTrackInput();
    if (!backupTrackInput) {
        return;
    }

    const value = Math.max(
        1,
        parseInt((defaults && defaults.videoTrackNumber) || DEFAULT_BACKUP_VIDEO_TRACK, 10) || DEFAULT_BACKUP_VIDEO_TRACK
    );

    if (force || backupTrackInput.dataset.userEdited !== "true") {
        backupTrackInput.value = String(value);
        backupTrackInput.dataset.autoValue = String(value);
        saveBackupVideoTrack(value);
    }
}

function applyAlignDefaults(defaults, force) {
    const value = Math.max(
        1,
        parseInt((defaults && defaults.videoTrackNumber) || DEFAULT_BACKUP_VIDEO_TRACK, 10) || DEFAULT_BACKUP_VIDEO_TRACK
    );
    saveAlignVideoTrack(value);
}

function setAutoEmptyBackupTrackEnabled(enabled) {
    const disabled = !!enabled;
    const backupTrackInput = getBackupVideoTrackInput();
    if (backupTrackInput) {
        backupTrackInput.disabled = disabled;
    }

    ["decrementBackupTrackButton", "incrementBackupTrackButton"].forEach((id) => {
        const button = document.getElementById(id);
        if (button) {
            button.disabled = disabled;
        }
    });
}

function resetAutoEmptyBackupTrackOption() {
    const checkbox = getAutoEmptyBackupTrackCheckbox();
    if (checkbox) {
        checkbox.checked = false;
    }
    setAutoEmptyBackupTrackEnabled(false);
}

let autoEmptyTrackPreviewRequest = 0;
let autoEmptyTrackPreviewPending = false;

async function refreshAutoEmptyTrackPreview() {
    const preview = document.getElementById('autoEmptyTrackPreview');
    if (!preview) return;
    if (!useAutoEmptyBackupTrack()) {
        autoEmptyTrackPreviewRequest += 1;
        preview.hidden = true;
        preview.textContent = '';
        return;
    }
    preview.hidden = false;
    if (busy || autoEmptyTrackPreviewPending) return;
    const request = ++autoEmptyTrackPreviewRequest;
    autoEmptyTrackPreviewPending = true;
    if (!preview.textContent) preview.textContent = 'Checking…';
    try {
        if (!await ensureHostLoaded()) throw new Error('Host unavailable');
        const result = JSON.parse(await callHost('exportBackup.getEmptyTrackPreview()'));
        if (request !== autoEmptyTrackPreviewRequest || !useAutoEmptyBackupTrack()) return;
        preview.textContent = result && result.ok
            ? `→ V${result.trackNumber}${result.createTrack ? ' (new track)' : ''}`
            : ((result && result.message) || 'Refresh to check track');
        preview.title = 'For the marked range. Checked again before exporting. Re-backup keeps its recorded track.';
    } catch (error) {
        if (request === autoEmptyTrackPreviewRequest) preview.textContent = 'Refresh to check track';
    } finally {
        autoEmptyTrackPreviewPending = false;
    }
}

function bindAutoEmptyBackupTrackOption() {
    const checkbox = getAutoEmptyBackupTrackCheckbox();
    if (!checkbox) {
        return;
    }

    checkbox.addEventListener("change", () => {
        setAutoEmptyBackupTrackEnabled(checkbox.checked);
        refreshAutoEmptyTrackPreview();
    });
    setInterval(() => {
        if (checkbox.checked && !document.hidden && !busy) refreshAutoEmptyTrackPreview();
    }, 3000);
}

function bindBackupTrackStepper() {
    const backupTrackInput = getBackupVideoTrackInput();
    const decrementButton = document.getElementById("decrementBackupTrackButton");
    const incrementButton = document.getElementById("incrementBackupTrackButton");
    if (!backupTrackInput) {
        return;
    }

    const stepTrack = (direction) => {
        if (backupTrackInput.disabled) {
            return;
        }

        const currentValue = getPositiveIntValue("exportVideoTrackInput", DEFAULT_BACKUP_VIDEO_TRACK);
        backupTrackInput.value = String(Math.max(1, currentValue + direction));
        backupTrackInput.dispatchEvent(new Event("input", { bubbles: true }));
    };

    if (decrementButton) {
        decrementButton.addEventListener("click", () => stepTrack(-1));
    }

    if (incrementButton) {
        incrementButton.addEventListener("click", () => stepTrack(1));
    }
}

function markBackupInputsDirty() {
    const backupTrackInput = getBackupVideoTrackInput();
    const alignTrackInput = getAlignVideoTrackInput();
    if (!backupTrackInput) {
        return;
    }

    backupTrackInput.addEventListener("input", () => {
        backupTrackInput.dataset.userEdited = "true";
        saveBackupVideoTrack(getPositiveIntValue("exportVideoTrackInput", DEFAULT_BACKUP_VIDEO_TRACK));
    });

    if (alignTrackInput) {
        alignTrackInput.addEventListener("input", () => {
            alignTrackInput.dataset.userEdited = "true";
            saveAlignVideoTrack(getPositiveIntValue("exportVideoTrackInput", DEFAULT_BACKUP_VIDEO_TRACK));
        });
    }
}

function bindAudioFormatInputs() {
    const audioInputs = getAudioFormatInputs();

    Object.keys(audioInputs).forEach((key) => {
        const input = audioInputs[key];
        if (!input) {
            return;
        }

        input.addEventListener("change", () => {
            if (input.checked) {
                saveSelectedAudioFormat(input.value);
            } else {
                setSelectedAudioFormat(key === "mp3" ? "wav" : "mp3");
                saveSelectedAudioFormat(getSelectedAudioFormat());
            }
        });
    });
}

function bindAlignOptions() {
    const sortCheckbox = document.getElementById("alignSortProjectFilesCheckbox");
    if (!sortCheckbox) {
        return;
    }

    try {
        sortCheckbox.checked = localStorage.getItem(ALIGN_SORT_PROJECT_FILES_STORAGE_KEY) === "true";
    } catch (error) {
        sortCheckbox.checked = false;
    }

    sortCheckbox.addEventListener("change", () => {
        try {
            localStorage.setItem(ALIGN_SORT_PROJECT_FILES_STORAGE_KEY, sortCheckbox.checked ? "true" : "false");
        } catch (error) {}
    });
}

function bindExportOptions() {
    const removeMarkersCheckbox = getRemoveSequenceMarkersCheckbox();
    const copyProjectCheckbox = getCopyProjectFileCheckbox();

    if (removeMarkersCheckbox) {
        try {
            const saved = localStorage.getItem(REMOVE_SEQUENCE_MARKERS_STORAGE_KEY);
            removeMarkersCheckbox.checked = saved !== "false";
        } catch (error) {
            removeMarkersCheckbox.checked = true;
        }

        removeMarkersCheckbox.addEventListener("change", () => {
            saveRemoveSequenceMarkers(removeMarkersCheckbox.checked);
        });
    }

    const exportModeInputs = getExportModeInputs();
    Object.keys(exportModeInputs).forEach((key) => {
        const input = exportModeInputs[key];
        if (!input) {
            return;
        }

        input.addEventListener("change", () => {
            if (input.checked) {
                saveSelectedExportMode(input.value);
            } else {
                setSelectedExportMode(key === "premiere" ? EXPORT_MODE_MEDIA_ENCODER : EXPORT_MODE_PREMIERE);
                saveSelectedExportMode(getSelectedExportMode());
            }
        });
    });

    if (copyProjectCheckbox) {
        try {
            copyProjectCheckbox.checked = localStorage.getItem(COPY_PROJECT_FILE_STORAGE_KEY) === "true";
        } catch (error) {
            copyProjectCheckbox.checked = false;
        }

        copyProjectCheckbox.addEventListener("change", () => {
            try {
                localStorage.setItem(COPY_PROJECT_FILE_STORAGE_KEY, copyProjectCheckbox.checked ? "true" : "false");
            } catch (error) {}
        });
    }
}

async function refreshSuggestedBackupTrack(force) {
    const fallback = { videoTrackNumber: DEFAULT_BACKUP_VIDEO_TRACK };

    if (!(await ensureHostLoaded())) {
        applyBackupDefaults(fallback, force);
        applyAlignDefaults(fallback, force);
        return;
    }

    const result = await callHost("exportBackup.getAlignmentDefaults()");
    const parsed = parseHostResult(result);
    if (!parsed || !parsed.ok) {
        applyBackupDefaults(fallback, force);
        applyAlignDefaults(fallback, force);
        return;
    }

    const defaults = {
        videoTrackNumber: parsed.suggestedVideoTrack || DEFAULT_BACKUP_VIDEO_TRACK
    };

    applyBackupDefaults(defaults, force);
    applyAlignDefaults(defaults, force);
}

function getTempUpdaterScriptPath() {
    return path.join(os.tmpdir(), "ExportBackup_update_launch.ps1");
}

function getTempUpdaterZipPath() {
    return path.join(os.tmpdir(), "ExportBackup_update_package.zip");
}

function getTempUpdaterResultPath() {
    return path.join(os.tmpdir(), "ExportBackup_update_result.json");
}

function getTempUpdaterLogPath() {
    return path.join(os.tmpdir(), "ExportBackup_update_log.txt");
}

function getUserCepExtensionPath() {
    return path.join(process.env.APPDATA || "", "Adobe", "CEP", "extensions", "Backup Project");
}

function readVersionInfo(silent) {
    try {
        const raw = fs.readFileSync(getVersionFilePath(), "utf8");
        const parsed = JSON.parse(raw);
        localVersion = parsed.version || "unknown";
        localVersionNotes = parsed.notes || "";
    } catch (error) {
        localVersion = "unknown";
        localVersionNotes = "";
        if (!silent) {
            setStatus(`Could not read version file.\n${error.message}`);
        }
    }

    return localVersion;
}

function compareVersions(a, b) {
    const aParts = String(a || "0").split(".").map((part) => parseInt(part, 10) || 0);
    const bParts = String(b || "0").split(".").map((part) => parseInt(part, 10) || 0);
    const length = Math.max(aParts.length, bParts.length);

    for (let i = 0; i < length; i += 1) {
        const left = aParts[i] || 0;
        const right = bParts[i] || 0;
        if (left > right) {
            return 1;
        }
        if (left < right) {
            return -1;
        }
    }

    return 0;
}

async function checkForUpdates() {
    const remoteUrl = "https://raw.githubusercontent.com/deepndense-sketch/ExportBackup/main/version.json";
    setUpdateButton(`Version ${localVersion}`, false, localVersionNotes);

    try {
        const remote = await new Promise((resolve, reject) => {
            https.get(remoteUrl, (response) => {
                if (response.statusCode !== 200) {
                    reject(new Error(`HTTP ${response.statusCode}`));
                    response.resume();
                    return;
                }

                let raw = "";
                response.setEncoding("utf8");
                response.on("data", (chunk) => {
                    raw += chunk;
                });
                response.on("end", () => {
                    try {
                        resolve(JSON.parse(raw));
                    } catch (error) {
                        reject(error);
                    }
                });
            }).on("error", reject);
        });

        remoteVersion = remote.version || "unknown";
        remoteVersionNotes = remote.notes || "";

        if (compareVersions(remoteVersion, localVersion) > 0) {
            setUpdateButton(`Update to ${remoteVersion}`, true, remoteVersionNotes);
            showUpdatePrompt(true);
        } else {
            setUpdateButton(`Latest Version ${localVersion}`, false, localVersionNotes);
        }
    } catch (error) {
        remoteVersionNotes = "";
        setUpdateButton(`Version ${localVersion}`, false, localVersionNotes);
    }
}

function delay(ms) {
    return new Promise((resolve) => {
        setTimeout(resolve, ms);
    });
}

function downloadFile(url, destinationPath) {
    return new Promise((resolve, reject) => {
        const file = fs.createWriteStream(destinationPath);
        const request = https.get(url, (response) => {
            if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
                file.close(() => {
                    fs.unlink(destinationPath, () => {
                        downloadFile(response.headers.location, destinationPath).then(resolve).catch(reject);
                    });
                });
                return;
            }

            if (response.statusCode !== 200) {
                file.close(() => {
                    fs.unlink(destinationPath, () => {});
                    reject(new Error(`HTTP ${response.statusCode}`));
                });
                response.resume();
                return;
            }

            response.pipe(file);
            file.on("finish", () => {
                file.close(resolve);
            });
        });

        request.on("error", (error) => {
            file.close(() => {
                fs.unlink(destinationPath, () => {});
                reject(error);
            });
        });

        file.on("error", (error) => {
            file.close(() => {
                fs.unlink(destinationPath, () => {});
                reject(error);
            });
        });
    });
}

async function monitorUpdaterCompletion() {
    const maxAttempts = 10;
    const resultPath = getTempUpdaterResultPath();

    for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
        await delay(3000);

        if (fileExists(resultPath)) {
            try {
                const parsed = readJsonFile(resultPath);
                if (parsed && parsed.ok) {
                    readVersionInfo(true);
                    await checkForUpdates();
                    setStatus(`Update complete.\nInstalled version: ${localVersion}\nRestart Premiere Pro if the panel was already open.`);
                    return;
                }

                setStatus(
                    `Updater failed.\n${(parsed && parsed.message) || "Unknown error."}\n` +
                    `Log: ${(parsed && parsed.logPath) || getTempUpdaterLogPath()}`
                );
                return;
            } catch (error) {
                setStatus(`Updater finished, but the result file could not be read.\n${error.message}`);
                return;
            }
        }

        readVersionInfo(true);
        await checkForUpdates();

        if (remoteVersion && compareVersions(remoteVersion, localVersion) <= 0) {
            setStatus(`Update complete.\nInstalled version: ${localVersion}\nRestart Premiere Pro if the panel was already open.`);
            return;
        }
    }

    setStatus(`Updater finished launching, but this panel still sees version ${localVersion}.\nIf the button stays blue, reopen the panel or restart Premiere Pro and check again.`);
}

function buildUpdaterLaunchCommand(scriptPath, zipPath, destination, resultPath, logPath) {
    const quote = value => "'" + String(value).replace(/'/g, "''") + "'";
    const payload = `& ${quote(scriptPath)} -ZipPath ${quote(zipPath)} -Destination ${quote(destination)} -ResultPath ${quote(resultPath)} -LogPath ${quote(logPath)}`;
    const encoded = Buffer.from(payload, 'utf16le').toString('base64');
    return `Start-Process PowerShell -Verb RunAs -WindowStyle Hidden -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-EncodedCommand','${encoded}'`;
}

function runGithubUpdate() {
    if (isUpdateWorkBusy()) {
        return;
    }

    const updateScriptPath = getUpdateScriptPath();
    if (!fileExists(updateScriptPath)) {
        setStatus("Update script was not found.");
        return;
    }

    if (remoteVersion && compareVersions(remoteVersion, localVersion) <= 0) {
        setStatus("This installation is already up to date.");
        checkForUpdates();
        return;
    }

    const tempUpdaterScriptPath = getTempUpdaterScriptPath();
    const tempUpdaterZipPath = getTempUpdaterZipPath();
    const tempUpdaterResultPath = getTempUpdaterResultPath();
    const tempUpdaterLogPath = getTempUpdaterLogPath();
    const remoteZipUrl = "https://github.com/deepndense-sketch/ExportBackup/archive/refs/heads/main.zip";

    setStatus("Downloading update package from GitHub...");

    try {
        fs.copyFileSync(updateScriptPath, tempUpdaterScriptPath);
        if (fileExists(tempUpdaterZipPath)) {
            fs.unlinkSync(tempUpdaterZipPath);
        }
        if (fileExists(tempUpdaterResultPath)) {
            fs.unlinkSync(tempUpdaterResultPath);
        }
        if (fileExists(tempUpdaterLogPath)) {
            fs.unlinkSync(tempUpdaterLogPath);
        }
    } catch (error) {
        setStatus(`Could not prepare updater.\n${error.message}`);
        return;
    }

    downloadFile(remoteZipUrl, tempUpdaterZipPath)
        .then(() => {
            setStatus("Launching GitHub updater. Accept the Windows permission prompt if it appears.");

            const command = buildUpdaterLaunchCommand(tempUpdaterScriptPath, tempUpdaterZipPath,
                getUserCepExtensionPath(), tempUpdaterResultPath, tempUpdaterLogPath);

            childProcess.execFile(
                "powershell.exe",
                ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
                (error) => {
                    if (error) {
                        setStatus(`Could not launch updater.\n${error.message}`);
                        return;
                    }

                    setStatus(`Updater launched for the CEP extensions folder.\nTarget: ${getUserCepExtensionPath()}\nWaiting for the update result...`);
                    monitorUpdaterCompletion();
                }
            );
        })
        .catch((error) => {
            setStatus(`Could not prepare updater.\n${error.message}`);
        });
}

function loadSavedPresets() {
    const defaults = {
        video: getDefaultVideoPresetPath(),
        mp3: getDefaultMp3PresetPath(),
        wav: getDefaultWavPresetPath()
    };

    try {
        const savedVideo = localStorage.getItem(VIDEO_PRESET_STORAGE_KEY);
        videoPresetPath = savedVideo && savedVideo.trim() && fileExists(savedVideo) ? savedVideo : defaults.video;
    } catch (error) {
        videoPresetPath = defaults.video;
    }

    try {
        const savedMp3 = localStorage.getItem(MP3_PRESET_STORAGE_KEY);
        mp3PresetPath = savedMp3 && savedMp3.trim() && fileExists(savedMp3) ? savedMp3 : defaults.mp3;
    } catch (error) {
        mp3PresetPath = defaults.mp3;
    }

    try {
        const savedWav = localStorage.getItem(WAV_PRESET_STORAGE_KEY);
        wavPresetPath = savedWav && savedWav.trim() && fileExists(savedWav) ? savedWav : defaults.wav;
    } catch (error) {
        wavPresetPath = defaults.wav;
    }
}

function loadSavedPaths() {
    try {
        const savedDestination = localStorage.getItem(BACKUP_DESTINATION_STORAGE_KEY);
        const destinationInputs = getBackupDestinationInputs();
        if (savedDestination === BACKUP_DESTINATION_PROJECT_ROOT && destinationInputs.projectRoot) {
            destinationInputs.projectRoot.checked = true;
        } else if (savedDestination === BACKUP_DESTINATION_PARENT_FOLDER && destinationInputs.parentFolder) {
            destinationInputs.parentFolder.checked = true;
        } else if (savedDestination === BACKUP_DESTINATION_MANUAL && destinationInputs.manual) {
            destinationInputs.manual.checked = true;
        } else if (destinationInputs.ftp) {
            destinationInputs.ftp.checked = true;
        }
    } catch (error) {}

    try {
        const savedExportFolder = localStorage.getItem(EXPORT_FOLDER_STORAGE_KEY);
        if (savedExportFolder && savedExportFolder.trim()) {
            manualExportFolder = savedExportFolder;
            if (getSelectedBackupDestination() === BACKUP_DESTINATION_MANUAL) {
                exportFolder = manualExportFolder;
            }
        }
    } catch (error) {}

    try {
        const savedAlignFolder = localStorage.getItem(ALIGN_FOLDER_STORAGE_KEY);
        if (savedAlignFolder && savedAlignFolder.trim()) {
            alignFolder = savedAlignFolder;
        }
    } catch (error) {}
}

function bindBackupDestinationInputs() {
    const inputs = getBackupDestinationInputs();
    const selector = document.getElementById("backupDestinationSelect");
    if (selector) {
        selector.value = getSelectedBackupDestination();
        selector.addEventListener("change", () => {
            const selected = inputs[selector.value];
            Object.keys(inputs).forEach((key) => { if (inputs[key]) inputs[key].checked = inputs[key] === selected; });
            if (selected) selected.dispatchEvent(new Event("change"));
        });
    }
    Object.keys(inputs).forEach((key) => {
        const input = inputs[key];
        if (!input) {
            return;
        }
        input.addEventListener("change", async () => {
            if (!input.checked) {
                return;
            }
            try {
                localStorage.setItem(BACKUP_DESTINATION_STORAGE_KEY, getSelectedBackupDestination());
            } catch (error) {}
            if (getSelectedBackupDestination() === BACKUP_DESTINATION_MANUAL) {
                exportFolder = manualExportFolder;
            } else {
                exportFolder = null;
            }
            document.getElementById("exportPath").textContent = "Resolving from the Premiere project...";
            await refreshResolvedBackupDestination();
        });
    });
}

function loadSavedBackupSettings() {
    const backupTrackInput = getBackupVideoTrackInput();
    const removeMarkersCheckbox = getRemoveSequenceMarkersCheckbox();

    try {
        const savedTrack = parseInt(localStorage.getItem(BACKUP_VIDEO_TRACK_STORAGE_KEY), 10);
        if (savedTrack && savedTrack > 0) {
            applyBackupDefaults({ videoTrackNumber: savedTrack }, true);
            applyAlignDefaults({ videoTrackNumber: savedTrack }, true);
            if (backupTrackInput) {
                backupTrackInput.dataset.userEdited = "true";
            }
        }
    } catch (error) {
        applyBackupDefaults({ videoTrackNumber: DEFAULT_BACKUP_VIDEO_TRACK }, true);
        applyAlignDefaults({ videoTrackNumber: DEFAULT_BACKUP_VIDEO_TRACK }, true);
    }

    try {
        const savedAlignTrack = parseInt(localStorage.getItem(ALIGN_VIDEO_TRACK_STORAGE_KEY), 10);
        if (savedAlignTrack && savedAlignTrack > 0) {
            applyAlignDefaults({ videoTrackNumber: savedAlignTrack }, true);
        }
    } catch (error) {}

    try {
        const savedFormat = localStorage.getItem(AUDIO_FORMAT_STORAGE_KEY);
        setSelectedAudioFormat(savedFormat || "mp3");
    } catch (error) {
        setSelectedAudioFormat("mp3");
    }

    // Always start in Premiere, even if Media Encoder was used last session.
    setSelectedExportMode(EXPORT_MODE_PREMIERE);

    if (removeMarkersCheckbox) {
        try {
            const savedRemoveMarkers = localStorage.getItem(REMOVE_SEQUENCE_MARKERS_STORAGE_KEY);
            removeMarkersCheckbox.checked = savedRemoveMarkers !== "false";
        } catch (error) {
            removeMarkersCheckbox.checked = true;
        }
    }
}

function saveVideoPreset(nextPath) {
    videoPresetPath = nextPath;

    try {
        localStorage.setItem(VIDEO_PRESET_STORAGE_KEY, nextPath);
    } catch (error) {}

    document.getElementById("videoPresetPath").textContent = videoPresetPath;
}

function saveMp3Preset(nextPath) {
    mp3PresetPath = nextPath;

    try {
        localStorage.setItem(MP3_PRESET_STORAGE_KEY, nextPath);
    } catch (error) {}

    updateAudioPresetDisplay();
}

function saveWavPreset(nextPath) {
    wavPresetPath = nextPath;

    try {
        localStorage.setItem(WAV_PRESET_STORAGE_KEY, nextPath);
    } catch (error) {}

    updateAudioPresetDisplay();
}

function updateAudioPresetDisplay() {
    document.getElementById("mp3PresetPath").textContent = mp3PresetPath;
    document.getElementById("wavPresetPath").textContent = wavPresetPath;
}

function getPresetDialogStartFolder(currentPresetPath) {
    try {
        if (currentPresetPath && fileExists(currentPresetPath)) {
            return path.dirname(currentPresetPath);
        }

        if (currentPresetPath) {
            const currentFolder = path.dirname(currentPresetPath);
            if (currentFolder && fs.existsSync(currentFolder)) {
                return currentFolder;
            }
        }
    } catch (error) {}

    return getBundledPresetFolderPath();
}

function choosePresetFile(title, currentPresetPath) {
    const startFolder = getPresetDialogStartFolder(currentPresetPath);
    let previousCwd = "";

    try {
        previousCwd = process.cwd();
        if (startFolder && fs.existsSync(startFolder)) {
            process.chdir(startFolder);
        }
    } catch (error) {}

    try {
        return window.cep.fs.showOpenDialogEx(false, false, title, startFolder, ["epr"]);
    } finally {
        try {
            if (previousCwd) {
                process.chdir(previousCwd);
            }
        } catch (error) {}
    }
}

async function getActiveSequenceName() {
    if (!(await ensureHostLoaded())) {
        return "";
    }

    const result = await callHost("exportBackup.getActiveSequenceName()");
    return String(result || "").trim();
}

async function getActiveProjectInfo() {
    if (!(await ensureHostLoaded())) {
        return { ok: false, message: "Could not load Premiere host script." };
    }

    const result = await callHost("exportBackup.getActiveProjectInfo()");
    return parseHostResult(result) || { ok: false, message: result || "Could not read the Premiere project file." };
}

function normalizeBackupCategoryName(value) {
    return String(value || "")
        .replace(/^\s*@\s*/, "")
        .replace(/\s+BACKUP\s*$/i, "")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
}

function getConfiguredBackupCategories() {
    return configuredBackupCategories.slice().sort((a, b) => {
        if (b.length !== a.length) {
            return b.length - a.length;
        }
        return a.localeCompare(b);
    });
}

function saveConfiguredBackupCategories() {
    try {
        localStorage.setItem(BACKUP_CATEGORIES_STORAGE_KEY, JSON.stringify(configuredBackupCategories));
    } catch (error) {}
}

function renderCategoryManager(selectedCategory) {
    const select = document.getElementById("categoryList");
    if (!select) {
        return;
    }
    const alphabetical = configuredBackupCategories.slice().sort((a, b) => a.localeCompare(b));
    select.innerHTML = alphabetical
        .map((category) => `<option value="${escapeHtml(category)}">${escapeHtml(category)}</option>`)
        .join("");
    if (selectedCategory && alphabetical.indexOf(selectedCategory) >= 0) {
        select.value = selectedCategory;
    }
}

function loadConfiguredBackupCategories() {
    configuredBackupCategories = BACKUP_CATEGORIES.slice();
}

function addBackupCategory() {
    const input = document.getElementById("categoryNameInput");
    const category = normalizeBackupCategoryName(input ? input.value : "");
    if (!category || !/^[A-Z0-9]+(?: [A-Z0-9]+)*$/.test(category)) {
        showBlockingMessage("Enter a category using letters, numbers, and spaces only.");
        return;
    }
    if (configuredBackupCategories.indexOf(category) >= 0) {
        renderCategoryManager(category);
        setStatus(`Category ${category} already exists.`);
        return;
    }
    configuredBackupCategories.push(category);
    saveConfiguredBackupCategories();
    renderCategoryManager(category);
    if (input) {
        input.value = "";
    }
    exportFolder = null;
    setStatus(`Added backup category: ${category}.`);
    refreshResolvedBackupDestination();
}

function deleteSelectedBackupCategory() {
    const select = document.getElementById("categoryList");
    const category = select ? normalizeBackupCategoryName(select.value) : "";
    if (!category) {
        showBlockingMessage("Select a category to delete.");
        return;
    }
    if (configuredBackupCategories.length <= 1) {
        showBlockingMessage("At least one backup category must remain.");
        return;
    }
    configuredBackupCategories = configuredBackupCategories.filter((candidate) => candidate !== category);
    saveConfiguredBackupCategories();
    renderCategoryManager();
    exportFolder = null;
    setStatus(`Deleted backup category: ${category}.`);
    refreshResolvedBackupDestination();
}

function bindCategoryManager() {
    const input = document.getElementById("categoryNameInput");
    if (input) {
        input.addEventListener("keydown", (event) => {
            if (event.key === "Enter") {
                event.preventDefault();
                addBackupCategory();
            }
        });
    }
}

function parseBackupProjectName(projectName) {
    const projectBaseName = path.basename(String(projectName || ""), path.extname(String(projectName || ""))).trim();
    const normalizedName = projectBaseName.replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
    const upperName = normalizedName.toUpperCase();
    const paddedUpperName = ` ${upperName} `;
    const category = getConfiguredBackupCategories().find((candidate) => paddedUpperName.indexOf(` ${candidate} `) >= 0) || "";
    const categoryIndex = category ? paddedUpperName.indexOf(` ${category} `) : -1;
    const withoutCategory = categoryIndex >= 0
        ? `${normalizedName.substring(0, categoryIndex)} ${normalizedName.substring(categoryIndex + category.length)}`.replace(/\s+/g, " ").trim()
        : normalizedName;
    const parts = withoutCategory ? withoutCategory.split(/\s+/) : [];
    const episodeNumbers = parts.filter((part) => /^\d+$/.test(part));
    const title = parts.filter((part) => !/^\d+$/.test(part)).join(" ").trim();
    const canonicalFolderName = [category].concat(episodeNumbers).concat(title ? [title] : []).filter(Boolean).join(" ");

    return {
        ok: !!category && episodeNumbers.length > 0,
        projectBaseName,
        category,
        episodeNumbers,
        title,
        canonicalFolderName,
        message: "The saved Premiere project name must contain one configured category and at least one number. They may appear anywhere or in either order."
    };
}

function findFtpCategoryFolder(category) {
    if (!fs.existsSync(FTP_BACKUP_ROOT)) {
        throw new Error(`FTP backup root is not available:\n${FTP_BACKUP_ROOT}`);
    }

    const matches = fs.readdirSync(FTP_BACKUP_ROOT, { withFileTypes: true })
        .filter((entry) => entry.isDirectory() && normalizeBackupCategoryName(entry.name) === category)
        .map((entry) => path.join(FTP_BACKUP_ROOT, entry.name));
    if (!matches.length) {
        throw new Error(`No category folder matching "${category}" was found in:\n${FTP_BACKUP_ROOT}`);
    }
    if (matches.length > 1) {
        throw new Error(`More than one FTP category folder matches "${category}":\n${matches.join("\n")}`);
    }
    return matches[0];
}

function getGroupedEpisodeRouting(parsedName) {
    const titleTokens = String((parsedName && parsedName.title) || "").toUpperCase().split(/\s+/).filter(Boolean);
    const role = titleTokens.indexOf("MAIN") >= 0
        ? "MAIN"
        : (titleTokens.indexOf("INTRO") >= 0 ? "INTRO" : "");
    const episodeNumber = parseInt(parsedName && parsedName.episodeNumbers && parsedName.episodeNumbers[0], 10) || 0;
    return {
        enabled: !!role,
        role,
        episodeNumber
    };
}

function findEpisodeRangeContainer(categoryFolder, category, episodeNumber) {
    const escapedCategory = String(category || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const rangePattern = new RegExp(`^${escapedCategory}\\s+(\\d+)\\s*-\\s*(\\d+)(?:\\s|$)`, "i");
    const matches = fs.readdirSync(categoryFolder, { withFileTypes: true })
        .filter((entry) => entry.isDirectory())
        .map((entry) => {
            const searchableName = entry.name.replace(/^\s*@\s*/, "").replace(/_/g, " ").replace(/\s+/g, " ").trim();
            const rangeMatch = searchableName.match(rangePattern);
            if (!rangeMatch) {
                return null;
            }
            const startEpisode = parseInt(rangeMatch[1], 10) || 0;
            const endEpisode = parseInt(rangeMatch[2], 10) || 0;
            return episodeNumber >= startEpisode && episodeNumber <= endEpisode
                ? path.join(categoryFolder, entry.name)
                : null;
        })
        .filter(Boolean);

    if (!matches.length) {
        throw new Error(
            `No existing ${category} episode-range container includes episode ${episodeNumber}.\n` +
            `Please create the group root folder for the ${category} project first.\n` +
            `Required format: ${category} startEpisode-endEpisode_any name\n` +
            `Example: ${category} 3240-3251_any name\n` +
            `Location: ${categoryFolder}`
        );
    }
    if (matches.length > 1) {
        throw new Error(
            `More than one ${category} episode-range container includes episode ${episodeNumber}:\n${matches.join("\n")}`
        );
    }
    return matches[0];
}

function updateCategoryDestinationNote(projectInfo, error) {
    const note = document.getElementById("categoryDestinationNote");
    if (!note) return;
    note.hidden = getSelectedBackupDestination() !== BACKUP_DESTINATION_FTP;
    if (note.hidden) { note.textContent = ""; return; }
    const parsed = projectInfo && projectInfo.projectPath
        ? parseBackupProjectName(projectInfo.projectName || path.basename(projectInfo.projectPath)) : null;
    if (error) {
        const message = String(error.message || error);
        if (/No category folder matching/.test(message)) {
            note.textContent = "Category folder doesn’t exist on Y:. Create the matching category folder on the drive before copying.";
        } else if (/configured category/.test(message)) {
            note.textContent = "Category not found. Use a built-in category in the project name and make sure its matching folder exists on Y:.";
        } else if (/FTP backup root is not available/.test(message)) {
            note.textContent = "Cannot access the backup folder on Y:. Make sure the drive and category folder are available before copying.";
        } else {
            note.textContent = message.replace(/\s+/g, " ");
        }
        return;
    }
    note.textContent = parsed && parsed.category
        ? `Will copy to the ${parsed.category} category on the Y: drive.`
        : "Will copy to the project’s show category on the Y: drive.";
}

async function resolveProjectBackupFolder(options) {
    const settings = options || {};
    const destinationMode = getSelectedBackupDestination();
    const projectInfo = await getActiveProjectInfo();
    updateCategoryDestinationNote(projectInfo);
    if (!projectInfo.ok || !projectInfo.projectPath) {
        throw new Error((projectInfo && projectInfo.message) || "Save the Premiere project before backing it up.");
    }

    let parsedName = null;
    let folderPath;
    if (destinationMode === BACKUP_DESTINATION_PROJECT_ROOT || destinationMode === BACKUP_DESTINATION_PARENT_FOLDER) {
        const projectDirectory = path.dirname(projectInfo.projectPath);
        folderPath = path.join(destinationMode === BACKUP_DESTINATION_PARENT_FOLDER ? path.dirname(projectDirectory) : projectDirectory, "BACKUP");
        if (settings.create !== false) {
            fs.mkdirSync(folderPath, { recursive: true });
        } else if (settings.requireExisting === true && !fs.existsSync(folderPath)) {
            throw new Error(`The project backup folder does not exist yet:\n${folderPath}`);
        }
    } else if (destinationMode === BACKUP_DESTINATION_MANUAL) {
        if (!manualExportFolder || !String(manualExportFolder).trim()) {
            throw new Error("Choose a manual backup folder first.");
        }
        folderPath = path.resolve(manualExportFolder);
        if (!fs.existsSync(folderPath) || !fs.statSync(folderPath).isDirectory()) {
            throw new Error(`The selected manual backup folder is not available:\n${folderPath}`);
        }
    } else {
        parsedName = parseBackupProjectName(projectInfo.projectName || path.basename(projectInfo.projectPath));
        if (!parsedName.ok) {
            throw new Error(`${parsedName.message}\n\nCurrent project: ${parsedName.projectBaseName || "Unnamed"}`);
        }
        const categoryFolder = findFtpCategoryFolder(parsedName.category);
        const groupedRouting = getGroupedEpisodeRouting(parsedName);
        if (groupedRouting.enabled) {
            const rangeContainer = findEpisodeRangeContainer(
                categoryFolder,
                parsedName.category,
                groupedRouting.episodeNumber
            );
            folderPath = path.join(rangeContainer, String(groupedRouting.episodeNumber));
        } else {
            folderPath = path.join(categoryFolder, parsedName.canonicalFolderName);
        }
        if (settings.create !== false) {
            fs.mkdirSync(folderPath, { recursive: true });
        } else if (settings.requireExisting === true && !fs.existsSync(folderPath)) {
            throw new Error(`The project backup folder does not exist yet:\n${folderPath}`);
        }
    }

    return { folderPath, projectInfo, parsedName, destinationMode };
}

async function refreshResolvedBackupDestination() {
    updateDestinationButtonLabel();
    updateCategoryDestinationNote();
    try {
        const resolved = await resolveProjectBackupFolder({ create: false });
        exportFolder = resolved.folderPath;
        updateAlignFolder(exportFolder);
        document.getElementById("exportPath").textContent = exportFolder;
        return resolved;
    } catch (error) {
        document.getElementById("exportPath").textContent = error.message;
        updateCategoryDestinationNote(null, error);
        return null;
    }
}

async function setActiveSequenceInOutToFullRange() {
    if (!(await ensureHostLoaded())) {
        return { ok: false, message: "Could not load Premiere host script." };
    }

    const result = await callHost("exportBackup.setActiveSequenceInOutToFullRange()");
    return parseHostResult(result) || { ok: false, message: result || "Could not set sequence In and Out." };
}

async function validateBackupExportSettings(backupVideoTrackNumber, selectedItems, allowExistingFiles, autoEmptyTrack) {
    if (!(await ensureHostLoaded())) {
        return { ok: false, message: "Could not load Premiere host script." };
    }

    const selectedItemsJson = JSON.stringify(selectedItems || getSelectedQueueItems());
    const result = await callHost(
        `exportBackup.validateBackupExportSettings(` +
            `${backupVideoTrackNumber},` +
            `"${escapeForEvalScript(exportFolder || "")}",` +
            `"${escapeForEvalScript(videoPresetPath || "")}",` +
            `"${escapeForEvalScript(mp3PresetPath || "")}",` +
            `"${escapeForEvalScript(wavPresetPath || "")}",` +
            `"${escapeForEvalScript(getSelectedAudioFormat())}",` +
            `"${escapeForEvalScript(selectedItemsJson)}",` +
            `${allowExistingFiles ? "true" : "false"},` +
            `${autoEmptyTrack ? "true" : "false"}` +
        `)`
    );
    return parseHostResult(result) || { ok: false, message: result || "Unknown validation error." };
}

async function getExportSelectionInfo() {
    if (!(await ensureHostLoaded())) {
        return { ok: false, message: "Could not load Premiere host script." };
    }

    const result = await callHost(`exportBackup.getExportSelectionInfo("${escapeForEvalScript(JSON.stringify(getBackupClipOwners()))}")`);
    return parseHostResult(result) || { ok: false, message: result || "Could not read export selection." };
}

function renderExportSelectionList(selectionInfo) {
    const container = document.getElementById("exportSelectionList");
    if (!container) {
        return;
    }

    if (!selectionInfo || !selectionInfo.ok) {
        container.innerHTML = `<div class="small-note">${(selectionInfo && selectionInfo.message) || "Could not read export selection."}</div>`;
        exportSelectionState = null;
        return;
    }

    exportSelectionState = selectionInfo;
    nextMergedAudioGroupId = 1;
    const items = Array.isArray(selectionInfo.items) ? selectionInfo.items : [];
    mergedAudioGroups = (Array.isArray(selectionInfo.audioGroups) ? selectionInfo.audioGroups : [])
        .map((group) => Array.from(new Set((group || []).map((value) => parseInt(value, 10) || 0).filter((value) => value > 0))).sort((a, b) => a - b))
        .filter((group) => group.length > 1)
        .map((trackNumbers) => ({ id: `g${nextMergedAudioGroupId++}`, trackNumbers }));

    if (!items.length) {
        container.innerHTML = `<div class="small-note">${selectionInfo.message || "No used audio tracks were found in the active sequence yet. The backup MP4 will still be queued."}</div>`;
        return;
    }

    container.innerHTML = items.map((item, index) => {
        const checkboxId = `exportSelectionItem_${index}`;
        const mergeCheckboxId = `mergeSelectionItem_${index}`;
        const checked = item.selected !== false ? "checked" : "";
        const disabled = item.locked ? "disabled" : "";
        const kindLabel = item.kind === "video" ? "Backup video" : `Audio track ${item.trackNumber}`;
        const existingGroup = item.kind === "audio"
            ? mergedAudioGroups.find((group) => group.trackNumbers.indexOf(parseInt(item.trackNumber, 10) || 0) >= 0)
            : null;
        const detail = existingGroup
            ? `Merged group: ${existingGroup.trackNumbers.join(", ")}`
            : (item.kind === "audio" && item.trackName ? item.trackName : "");
        let mergeControl = `<span></span>`;
        if (item.kind === "video") {
            mergeControl = `<span class="selection-list-button-group"><button id="mergeAudioButton" class="secondary selection-list-merge-button" type="button" onclick="mergeSelectedAudioTracks()">Merge Selection</button><button id="clearAudioMergesButton" class="secondary selection-list-merge-button" type="button" onclick="clearAudioMerges()">Clear Merges</button></span>`;
        } else if (item.kind === "audio") {
            mergeControl = `<label class="merge-checkbox-wrap" for="${mergeCheckboxId}"><span>Merge</span><input class="merge-checkbox" type="checkbox" id="${mergeCheckboxId}" data-kind="audio" data-track-number="${item.trackNumber || 0}" data-merged-group="${existingGroup ? existingGroup.id : ""}" ${existingGroup ? "checked disabled" : ""}></label>`;
        }

        return (
            `<div class="selection-item${existingGroup ? " is-merged" : ""}">` +
                `<input class="queue-checkbox" type="checkbox" id="${checkboxId}" data-kind="${item.kind}" data-track-number="${item.trackNumber || 0}" ${checked} ${disabled}>` +
                `<label for="${checkboxId}">` +
                    `<strong>${escapeHtml(kindLabel)}</strong>` +
                    (detail ? `<small>${escapeHtml(detail)}</small>` : "") +
                `</label>` +
                mergeControl +
            `</div>`
        );
    }).join("");
    container.onchange = updateMergeActionVisibility;
    updateMergeActionVisibility();
}

function updateMergeActionVisibility() {
    const merge = document.getElementById("mergeAudioButton");
    const clear = document.getElementById("clearAudioMergesButton");
    if (merge) merge.hidden = getUnmergedCheckedAudioInputs().length < 2;
    if (clear) clear.hidden = !mergedAudioGroups.length;
}

function clearAudioMerges() {
    if (!exportSelectionState) {
        return;
    }
    exportSelectionState.audioGroups = [];
    renderExportSelectionList(exportSelectionState);
    setStatus("Audio merge groups cleared. Select tracks and use Merge Selection to create the new Re-backup layout.");
}

function getUnmergedCheckedAudioInputs() {
    const container = document.getElementById("exportSelectionList");
    if (!container) {
        return [];
    }

    return Array.prototype.slice.call(container.querySelectorAll(".merge-checkbox[data-kind='audio']:checked"))
        .filter((input) => !input.getAttribute("data-merged-group"));
}

function mergeSelectedAudioTracks() {
    const inputs = getUnmergedCheckedAudioInputs();
    if (inputs.length < 2) {
        setStatus("Select two or more audio tracks, then click Merge.");
        return;
    }

    const groupId = `g${nextMergedAudioGroupId++}`;
    const trackNumbers = inputs
        .map((input) => parseInt(input.getAttribute("data-track-number"), 10) || 0)
        .filter((trackNumber) => trackNumber > 0)
        .sort((a, b) => a - b);

    mergedAudioGroups.push({ id: groupId, trackNumbers });

    inputs.forEach((input) => {
        input.setAttribute("data-merged-group", groupId);
        input.disabled = true;
        const item = input.closest(".selection-item");
        if (item) {
            item.classList.add("is-merged");
            const small = item.querySelector("small");
            const label = `Merged group: ${trackNumbers.join(", ")}`;
            if (small) {
                small.textContent = label;
            } else {
                item.querySelector("label").insertAdjacentHTML("beforeend", `<small>${escapeHtml(label)}</small>`);
            }
        }
    });

    updateMergeActionVisibility();
    setStatus(`Merged audio tracks: ${trackNumbers.join(", ")}.`);
}

function getSelectedAudioTrackNumbers() {
    const container = document.getElementById("exportSelectionList");
    if (!container) {
        return [];
    }

    const groupedTrackNumbers = {};
    mergedAudioGroups.forEach((group) => {
        group.trackNumbers.forEach((trackNumber) => {
            groupedTrackNumbers[trackNumber] = true;
        });
    });

    return Array.prototype.slice.call(container.querySelectorAll(".queue-checkbox[data-kind='audio']:checked"))
        .map((input) => parseInt(input.getAttribute("data-track-number"), 10) || 0)
        .filter((trackNumber) => trackNumber > 0 && !groupedTrackNumbers[trackNumber]);
}

function getSelectedQueueItems() {
    const container = document.getElementById("exportSelectionList");
    const videoInput = container ? container.querySelector(".queue-checkbox[data-kind='video']") : null;
    const audioInputs = container ? Array.from(container.querySelectorAll(".queue-checkbox[data-kind='audio']")) : [];
    const checkedTracks = audioInputs.filter(input => input.checked && !input.disabled)
        .map(input => parseInt(input.getAttribute('data-track-number'), 10)).filter(number => number > 0);
    const selectedGroups = mergedAudioGroups.map(group => group.trackNumbers.filter(number => checkedTracks.includes(number)))
        .filter(group => group.length > 1);

    return {
        includeVideo: videoInput ? !!videoInput.checked : true,
        audioTracks: checkedTracks.filter(number => !selectedGroups.some(group => group.includes(number))),
        audioGroups: selectedGroups,
        preserveUnselectedAudio: audioInputs.some(input => !input.checked || input.disabled),
        replaceAudioLayout: true
    };
}

async function refreshExportSelection() {
    if (busy) {
        return;
    }

    renderExportSelectionList({ ok: true, items: [], message: "Reading active sequence tracks..." });
    const selectionInfo = await getExportSelectionInfo();
    renderExportSelectionList(selectionInfo);
    await refreshAutoEmptyTrackPreview();
}

function getExportSelectionTrackSignature(selectionInfo) {
    const items = selectionInfo && Array.isArray(selectionInfo.items) ? selectionInfo.items : [];
    return items
        .filter((item) => item && item.kind === "audio")
        .map((item) => parseInt(item.trackNumber, 10) || 0)
        .filter((trackNumber) => trackNumber > 0)
        .sort((a, b) => a - b)
        .join(",");
}

async function syncExportSelectionWithActiveTimeline() {
    const latestSelection = await getExportSelectionInfo();
    if (!latestSelection || !latestSelection.ok) {
        throw new Error((latestSelection && latestSelection.message) || "Could not read the active sequence audio tracks.");
    }

    const previousSignature = getExportSelectionTrackSignature(exportSelectionState);
    const latestSignature = getExportSelectionTrackSignature(latestSelection);
    if (!exportSelectionState || previousSignature !== latestSignature || exportSelectionState.sequenceName !== latestSelection.sequenceName) {
        if (exportSelectionState && exportSelectionState.sequenceName === latestSelection.sequenceName) {
            const container = document.getElementById('exportSelectionList');
            const choices = {};
            if (container) Array.from(container.querySelectorAll('.queue-checkbox')).forEach(input => {
                choices[input.getAttribute('data-kind') + ':' + input.getAttribute('data-track-number')] = !!input.checked;
            });
            latestSelection.items.forEach(item => {
                const key = item.kind + ':' + (item.trackNumber || 0);
                // Newly discovered tracks should not silently expand this export.
                item.selected = choices[key] === true;
            });
            latestSelection.audioGroups = mergedAudioGroups.map(group => group.trackNumbers);
        }
        renderExportSelectionList(latestSelection);
        setStatus(`Audio track layout updated from the active sequence: ${latestSignature || "no source audio tracks"}.`);
        return true;
    }
    return false;
}

function parseTrackNumbersFromFileName(name, baseName) {
    const lowerName = String(name || "").toLowerCase();
    const lowerBase = String(baseName || "").toLowerCase();
    const prefix = `${lowerBase}_track`;

    if (!lowerName.startsWith(prefix)) {
        return [];
    }

    const remainder = name.substring(prefix.length);
    const dotIndex = remainder.lastIndexOf(".");
    if (dotIndex <= 0) {
        return [];
    }

    const seen = {};
    return remainder.substring(0, dotIndex)
        .split("-")
        .map((part) => parseInt(part, 10) || 0)
        .filter((trackNumber) => {
            if (trackNumber < 1 || seen[trackNumber]) {
                return false;
            }

            seen[trackNumber] = true;
            return true;
        })
        .sort((a, b) => a - b);
}

function parseTrackNumberFromFileName(name, baseName) {
    const trackNumbers = parseTrackNumbersFromFileName(name, baseName);
    return trackNumbers.length ? trackNumbers[0] : 0;
}

function normalizeAudioEntries(entries, baseName, preferredAudioFormat) {
    const normalized = [];
    const indexesByTrackKey = {};
    const preferredFormat = String(preferredAudioFormat || "").toLowerCase();

    (entries || []).forEach((entry) => {
        if (!entry || !entry.path || !fileExists(entry.path)) {
            return;
        }

        const fileName = entry.name || path.basename(entry.path);
        const parsedTrackNumbers = parseTrackNumbersFromFileName(fileName, baseName);
        const rawTrackNumbers = Array.isArray(entry.trackNumbers) && entry.trackNumbers.length
            ? entry.trackNumbers
            : (parsedTrackNumbers.length ? parsedTrackNumbers : [entry.trackNumber]);
        const seenTrackNumbers = {};
        const trackNumbers = rawTrackNumbers
            .map((value) => parseInt(value, 10) || 0)
            .filter((trackNumber) => {
                if (trackNumber < 1 || seenTrackNumbers[trackNumber]) {
                    return false;
                }

                seenTrackNumbers[trackNumber] = true;
                return true;
            })
            .sort((a, b) => a - b);
        const trackNumber = trackNumbers.length ? trackNumbers[0] : 0;
        if (trackNumber < 1) {
            return;
        }

        const normalizedEntry = {
            path: entry.path,
            trackNumber,
            trackNumbers,
            exportRange: entry.exportRange || null,
            name: fileName
        };
        const trackKey = trackNumbers.join("-");
        const score = (preferredFormat && getAudioEntryFormat(normalizedEntry.path) === preferredFormat ? 100 : 0) + (entry.prefer === true ? 10 : 0);
        normalizedEntry.matchScore = score;

        if (Object.prototype.hasOwnProperty.call(indexesByTrackKey, trackKey)) {
            const existingIndex = indexesByTrackKey[trackKey];
            if (score > (normalized[existingIndex].matchScore || 0)) {
                normalized[existingIndex] = normalizedEntry;
            }
            return;
        }

        indexesByTrackKey[trackKey] = normalized.length;
        normalized.push(normalizedEntry);
    });

    normalized.sort((a, b) => a.trackNumber - b.trackNumber);
    normalized.forEach((entry) => { delete entry.matchScore; });
    return normalized;
}
function getAudioEntryTrackKey(entry, baseName) {
    if (!entry) {
        return "";
    }

    const fileName = entry.name || path.basename(entry.path || "");
    const parsedTrackNumbers = parseTrackNumbersFromFileName(fileName, baseName);
    const trackNumbers = Array.isArray(entry.trackNumbers) && entry.trackNumbers.length
        ? entry.trackNumbers
        : (parsedTrackNumbers.length ? parsedTrackNumbers : [entry.trackNumber]);

    return trackNumbers
        .map((value) => parseInt(value, 10) || 0)
        .filter((trackNumber) => trackNumber > 0)
        .sort((a, b) => a - b)
        .join("-");
}
function readManifestForSequence(folderPath, sequenceName) {
    if (!folderPath || !sequenceName) {
        return null;
    }

    const manifestPath = getManifestPath(folderPath, sequenceName);
    if (!fileExists(manifestPath)) {
        return null;
    }

    const manifest = readJsonFile(manifestPath);
    if (!manifest || !manifest.baseName) {
        return null;
    }

    manifest.manifestPath = manifestPath;
    return manifest;
}

function scanExportFolderForSequence(folderPath, sequenceName, manifest, options) {
    const settings = options || {};
    const manifestOnly = settings.manifestOnly === true && !!manifest;
    const sanitizedBase = sanitizeSequenceName(sequenceName);
    const entries = fs.readdirSync(folderPath, { withFileTypes: true });
    const files = entries.filter((entry) => entry.isFile()).map((entry) => entry.name);
    const lowerBase = sanitizedBase.toLowerCase();
    const backupPrefix = `${lowerBase}_backup.`;
    const preferredAudioFormat = getSelectedAudioFormat();

    let videoPath = "";
    let audio = [];
    let audioCandidates = [];

    if (manifest) {
        const expectedFiles = Array.isArray(manifest.expectedFiles) ? manifest.expectedFiles : [];
        expectedFiles.forEach((entry) => {
            if (!entry || !entry.path || !fileExists(entry.path)) {
                return;
            }

            if (entry.kind === "video" && !videoPath) {
                videoPath = entry.path;
                return;
            }

            if (entry.kind === "audio") {
                audio.push({
                    path: entry.path,
                    trackNumber: parseInt(entry.trackNumber, 10) || 0,
                    trackNumbers: Array.isArray(entry.trackNumbers) ? entry.trackNumbers : [],
                    exportRange: entry.exportRange || manifest.exportRange || null,
                    name: entry.name || path.basename(entry.path)
                });
            }
        });
    }

    if (!manifestOnly) {
        files.forEach((fileName) => {
            const absolutePath = path.join(folderPath, fileName);
            const lowerName = fileName.toLowerCase();

            if (lowerName.includes(REBACKUP_OLD_MARKER.toLowerCase())) {
                return;
            }

            if (!videoPath && lowerName.startsWith(backupPrefix)) {
                videoPath = absolutePath;
                return;
            }

            const trackNumbers = parseTrackNumbersFromFileName(fileName, sanitizedBase);
            const trackNumber = trackNumbers.length ? trackNumbers[0] : 0;
            if (trackNumber > 0 && !audio.some((entry) => entry.path === absolutePath)) {
                audio.push({
                    path: absolutePath,
                    trackNumber,
                    trackNumbers,
                    name: fileName
                });
            }
        });
    }

    audioCandidates = audio.slice();
    audio = normalizeAudioEntries(audio, sanitizedBase, preferredAudioFormat);
    const selectedAudioByTrackKey = {};
    audio.forEach((entry) => {
        selectedAudioByTrackKey[getAudioEntryTrackKey(entry, sanitizedBase)] = getPathComparisonKey(entry.path);
    });
    const staleAudioPaths = [];
    const seenStaleAudioPaths = {};
    audioCandidates.forEach((entry) => {
        const trackKey = getAudioEntryTrackKey(entry, sanitizedBase);
        const entryPathKey = getPathComparisonKey(entry.path);
        if (trackKey && selectedAudioByTrackKey[trackKey] && selectedAudioByTrackKey[trackKey] !== entryPathKey && !seenStaleAudioPaths[entryPathKey]) {
            seenStaleAudioPaths[entryPathKey] = true;
            staleAudioPaths.push(entry.path);
        }
    });

    return {
        baseName: sanitizedBase,
        videoPath,
        audio,
        staleAudioPaths,
        folderFiles: files,
        manifest: manifest || null
    };
}

function writeExportManifest(manifest) {
    if (!manifest || !manifest.folderPath || !manifest.baseName) {
        return null;
    }

    const manifestPath = getManifestPath(manifest.folderPath, manifest.baseName);
    const toWrite = Object.assign({}, manifest, { manifestPath });
    fs.writeFileSync(manifestPath, `${JSON.stringify(toWrite, null, 2)}\n`, "utf8");
    return manifestPath;
}

async function waitBeforeLocalFileStep(message, ms) {
    setStatus(`${message}\nWaiting ${Math.round(ms / 1000)} seconds before trying.`);
    await delay(ms);
}

async function waitForFileState(description, verifier) {
    let lastError = null;
    for (let i = 0; i < 8; i += 1) {
        try {
            if (verifier()) {
                return true;
            }
        } catch (error) {
            lastError = error;
        }
        await delay(500);
    }

    if (lastError) {
        throw lastError;
    }
    throw new Error(`${description} did not finish in Windows yet.`);
}

async function runLocalFileStepWithRetry(title, details, action, verifier) {
    let lastError = null;
    const retryDelays = [10000];

    while (true) {
        for (let attempt = 1; attempt <= 2; attempt += 1) {
            if (attempt > 1) {
                const waitSeconds = Math.round(retryDelays[attempt - 2] / 1000);
                await waitBeforeLocalFileStep(
                    `Retrying ${title} in ${waitSeconds} seconds.\n` +
                    `${details}\n` +
                    `Retry ${attempt - 1}/1.\n` +
                    `Last error: ${lastError ? lastError.message : "unknown"}`,
                    retryDelays[attempt - 2]
                );
            }

            try {
                setStatus(`${title}\n${details}\nAttempt ${attempt}/2.`);
                await action();
                await waitForFileState(title, verifier);
                return;
            } catch (error) {
                lastError = error;
            }
        }

        const recoveryMessage = "File still busy\n\nPress Align Existing button to delete and import re-export.";
        setStatus(recoveryMessage);
        throw new Error(recoveryMessage);
    }
}


function normalizeLocalFilePath(filePath) {
    const pathText = String(filePath || "").trim();
    return pathText ? path.normalize(pathText) : "";
}

function inspectLocalFile(filePath) {
    const normalizedPath = normalizeLocalFilePath(filePath);
    if (!normalizedPath) {
        return { path: "", exists: false, error: "", code: "" };
    }

    try {
        const stats = fs.lstatSync(normalizedPath);
        return {
            path: normalizedPath,
            exists: true,
            size: stats.size,
            error: "",
            code: ""
        };
    } catch (error) {
        if (error && error.code === "ENOENT") {
            return { path: normalizedPath, exists: false, error: "", code: "" };
        }
        return {
            path: normalizedPath,
            exists: true,
            error: error && error.message ? error.message : String(error || "Unknown file error"),
            code: error && error.code ? error.code : "UNKNOWN"
        };
    }
}

function deleteLocalFileNow(filePath) {
    const before = inspectLocalFile(filePath);
    if (!before.exists) {
        return { ok: true, path: before.path, error: "", code: "" };
    }

    try {
        fs.chmodSync(before.path, 0o666);
    } catch (error) {}

    let deleteError = null;
    try {
        fs.unlinkSync(before.path);
    } catch (error) {
        deleteError = error;
    }

    const after = inspectLocalFile(before.path);
    if (!after.exists) {
        return { ok: true, path: before.path, error: "", code: "" };
    }

    return {
        ok: false,
        path: before.path,
        error: deleteError && deleteError.message
            ? deleteError.message
            : (after.error || "Windows still reports that the file exists."),
        code: deleteError && deleteError.code
            ? deleteError.code
            : (after.code || "STILL_EXISTS")
    };
}

function buildDeletePendingPath(filePath) {
    const parsed = path.parse(normalizeLocalFilePath(filePath));
    const stamp = new Date().toISOString()
        .replace(/[-:TZ.]/g, "")
        .slice(0, 14);
    let candidate = path.join(parsed.dir, `${parsed.name}_DELETE_PENDING_${stamp}${parsed.ext}`);
    let index = 1;

    while (fileExists(candidate)) {
        candidate = path.join(parsed.dir, `${parsed.name}_DELETE_PENDING_${stamp}_${index}${parsed.ext}`);
        index += 1;
    }

    return candidate;
}

function moveLocalFileAside(filePath) {
    const sourcePath = normalizeLocalFilePath(filePath);
    if (!sourcePath || !inspectLocalFile(sourcePath).exists) {
        return { ok: true, pendingPath: "", error: "", code: "" };
    }

    const pendingPath = buildDeletePendingPath(sourcePath);
    try {
        fs.renameSync(sourcePath, pendingPath);
        return {
            ok: !inspectLocalFile(sourcePath).exists && inspectLocalFile(pendingPath).exists,
            pendingPath,
            error: "",
            code: ""
        };
    } catch (error) {
        return {
            ok: !inspectLocalFile(sourcePath).exists,
            pendingPath: !inspectLocalFile(sourcePath).exists ? pendingPath : "",
            error: error && error.message ? error.message : String(error || "Rename failed"),
            code: error && error.code ? error.code : "UNKNOWN"
        };
    }
}

async function deleteLocalFileWithRetry(filePath, label, retryDelays) {
    const delays = Array.isArray(retryDelays) && retryDelays.length
        ? retryDelays
        : [0, 1500, 5000];
    let lastResult = { ok: true, path: normalizeLocalFilePath(filePath), error: "", code: "" };

    for (let attempt = 0; attempt < delays.length; attempt += 1) {
        if (delays[attempt] > 0) {
            setStatus(
                `Waiting ${Math.round(delays[attempt] / 1000)} seconds before retrying ${label}.\n` +
                `${lastResult.path}\n` +
                `Last error: ${lastResult.code || "UNKNOWN"}: ${lastResult.error || "File still exists."}`
            );
            await delay(delays[attempt]);
        }

        setStatus(`Deleting ${label}.\n${lastResult.path}\nAttempt ${attempt + 1}/${delays.length}.`);
        lastResult = deleteLocalFileNow(lastResult.path);
        if (lastResult.ok) {
            return lastResult;
        }
    }

    return lastResult;
}

async function cleanupLocalFilesBestEffort(filePaths, label, retryDelays) {
    const uniquePaths = [];
    const seen = {};
    (filePaths || []).forEach((filePath) => {
        const normalizedPath = normalizeLocalFilePath(filePath);
        const key = getPathComparisonKey(normalizedPath);
        if (normalizedPath && key && !seen[key] && inspectLocalFile(normalizedPath).exists) {
            seen[key] = true;
            uniquePaths.push(normalizedPath);
        }
    });

    let pending = uniquePaths.slice();
    const deleted = [];
    const errors = {};
    const resolvedRetryDelays = Array.isArray(retryDelays) && retryDelays.length
        ? retryDelays
        : [0, 1500, 5000];

    for (let attempt = 0; attempt < resolvedRetryDelays.length && pending.length; attempt += 1) {
        if (resolvedRetryDelays[attempt] > 0) {
            setStatus(
                `New backup files are already imported.\n` +
                `Retrying old-file cleanup in ${Math.round(resolvedRetryDelays[attempt] / 1000)} seconds.\n` +
                `Files remaining: ${pending.length}.`
            );
            await delay(resolvedRetryDelays[attempt]);
        }

        const nextPending = [];
        pending.forEach((filePath) => {
            const result = deleteLocalFileNow(filePath);
            if (result.ok) {
                deleted.push(filePath);
                delete errors[getPathComparisonKey(filePath)];
            } else {
                nextPending.push(filePath);
                errors[getPathComparisonKey(filePath)] = `${result.code || "UNKNOWN"}: ${result.error || "File still exists."}`;
            }
        });
        pending = nextPending;
    }

    return { deleted, pending, errors, label: label || "old backup file" };
}

async function removeFileIfExists(filePath, description) {
    const label = description || "old local file";
    const normalizedPath = normalizeLocalFilePath(filePath);
    if (!normalizedPath || !inspectLocalFile(normalizedPath).exists) {
        return;
    }

    const deleteResult = await deleteLocalFileWithRetry(normalizedPath, label, [0, 1500, 5000]);
    if (deleteResult.ok) {
        return;
    }

    setStatus(`Delete did not work. Moving ${label} aside.\n${normalizedPath}`);
    const moveResult = moveLocalFileAside(normalizedPath);
    if (moveResult.ok) {
        return;
    }

    throw new Error(
        `Could not free ${label}.\n` +
        `Path: ${normalizedPath}\n` +
        `Delete: ${deleteResult.code || "UNKNOWN"}: ${deleteResult.error || "failed"}\n` +
        `Rename: ${moveResult.code || "UNKNOWN"}: ${moveResult.error || "failed"}\n` +
        "The completed replacement export was left untouched."
    );
}
async function replaceRebackupFile(entry, fileIndex, totalFiles) {
    const label = entry && entry.name ? entry.name : path.basename((entry && (entry.finalPath || entry.path)) || "backup file");
    const prefix = `File ${fileIndex}/${totalFiles}: ${label}`;

    await runLocalFileStepWithRetry(
        "Checking new TEMP file",
        `${prefix}\n${entry.path}`,
        async () => {},
        () => fileExists(entry.path)
    );

    if (entry.oldFinalPath && entry.oldFinalPath !== entry.finalPath) {
        await removeFileIfExists(entry.oldFinalPath, "old-format backup file");
    }

    await removeFileIfExists(entry.finalPath, "old final backup file");

    await runLocalFileStepWithRetry(
        "Renaming TEMP file to final backup name",
        `${prefix}\nTEMP: ${entry.path}\nFinal: ${entry.finalPath}`,
        async () => {
            if (!fileExists(entry.path)) {
                throw new Error(`TEMP file is missing: ${entry.path}`);
            }
            fs.renameSync(entry.path, entry.finalPath);
        },
        () => fileExists(entry.finalPath) && !fileExists(entry.path)
    );

    setStatus(
        "Local backup file replaced.\n" +
        `File ${fileIndex}/${totalFiles}: ${path.basename(entry.finalPath)}`
    );
    return { ok: true };
}

function getManifestPreservedPaths(manifest) {
    const expectedFiles = manifest && Array.isArray(manifest.expectedFiles) ? manifest.expectedFiles : [];
    const paths = [];
    const seen = {};

    expectedFiles.forEach((entry) => {
        const entryPaths = entry && Array.isArray(entry.preservedPaths) ? entry.preservedPaths : [];
        entryPaths.forEach((filePath) => {
            const key = getPathComparisonKey(filePath);
            if (key && !seen[key]) {
                seen[key] = true;
                paths.push(filePath);
            }
        });
    });

    return paths;
}

function getIncompleteExpectedExportPaths(manifest) {
    const expectedFiles = manifest && Array.isArray(manifest.expectedFiles) ? manifest.expectedFiles : [];
    return expectedFiles
        .filter((entry) => {
            if (!entry || !entry.path || !fileExists(entry.path)) {
                return true;
            }
            try {
                return fs.statSync(entry.path).size < 1;
            } catch (error) {
                return true;
            }
        })
        .map((entry) => (entry && entry.path) || "Unknown export file");
}

function assertExpectedExportsAreReady(manifest) {
    const incompletePaths = getIncompleteExpectedExportPaths(manifest);
    if (incompletePaths.length) {
        throw new Error([
            "New Re-backup export files are not complete yet.",
            "The preserved old clips were left untouched.",
            incompletePaths.join(String.fromCharCode(10))
        ].join(String.fromCharCode(10)));
    }
}

async function finalizeRebackupFiles(manifest) {
    const expectedFiles = manifest && Array.isArray(manifest.expectedFiles) ? manifest.expectedFiles : [];
    const replacementFiles = expectedFiles.filter((entry) => entry && entry.path && entry.finalPath && entry.path !== entry.finalPath);

    if (!manifest || manifest.rebackupReplacementPrepared !== true) {
        throw new Error("Re-backup cleanup was not verified, so local backup files were not finalized.");
    }

    // Only legacy _REBKP_TEMP recovery needs a blocking filesystem change
    // before import. Current Re-backup exports already use their final paths.
    if (replacementFiles.length) {
        await waitBeforeLocalFileStep(
            "Premiere cleanup is done. Legacy TEMP-file recovery will start next.",
            1500
        );
        setStatus(`Recovering legacy TEMP files.\nFiles to rename: ${replacementFiles.length}.`);

        for (let i = 0; i < replacementFiles.length; i += 1) {
            const entry = replacementFiles[i];
            await replaceRebackupFile(entry, i + 1, replacementFiles.length);
            entry.path = entry.finalPath;
            entry.name = path.basename(entry.finalPath);
        }
    }

    manifest.rebackupFilesReady = true;
    return manifest;
}

async function cleanupRebackupPreservedFilesAfterAlignment(manifest, retryDelays, premierePendingPaths) {
    const expectedFiles = manifest && Array.isArray(manifest.expectedFiles) ? manifest.expectedFiles : [];
    const preservedPaths = getManifestPreservedPaths(manifest);
    const cleanupResult = await cleanupLocalFilesBestEffort(
        preservedPaths,
        "preserved old backup files",
        retryDelays
    );
    const retainedKeys = {};

    cleanupResult.pending.forEach((filePath) => {
        retainedKeys[getPathComparisonKey(filePath)] = true;
    });
    (premierePendingPaths || []).forEach((filePath) => {
        retainedKeys[getPathComparisonKey(filePath)] = true;
    });

    expectedFiles.forEach((entry) => {
        if (!entry || !Array.isArray(entry.preservedPaths)) {
            return;
        }
        entry.preservedPaths = entry.preservedPaths.filter((filePath) => retainedKeys[getPathComparisonKey(filePath)]);
    });

    const premierePendingCount = Array.isArray(premierePendingPaths) ? premierePendingPaths.length : 0;
    manifest.cleanupPendingPaths = cleanupResult.pending.slice();
    manifest.cleanupErrors = cleanupResult.errors;
    manifest.rebackupPreservedCleanupDone = cleanupResult.pending.length === 0 && premierePendingCount === 0;
    manifest.rebackupFinalized = cleanupResult.pending.length === 0 && premierePendingCount === 0;
    return cleanupResult;
}

function replacePathExtension(filePath, extension) {
    const parsedPath = path.parse(filePath || "");
    const resolvedExtension = String(extension || "").startsWith(".") ? String(extension || "") : `.${extension}`;
    if (!parsedPath.dir || !parsedPath.name) {
        return "";
    }
    return path.join(parsedPath.dir, `${parsedPath.name}${resolvedExtension}`);
}

function getOppositeAudioFormatPath(filePath) {
    const extension = path.extname(filePath || "").toLowerCase();
    if (extension === ".mp3") {
        return replacePathExtension(filePath, ".wav");
    }
    if (extension === ".wav") {
        return replacePathExtension(filePath, ".mp3");
    }
    return "";
}
function getRebackupFinalPath(tempPath) {
    if (!tempPath) {
        return "";
    }

    const parsedPath = path.parse(tempPath);
    if (!parsedPath.name.toUpperCase().endsWith(REBACKUP_TEMP_MARKER)) {
        return "";
    }

    const finalName = parsedPath.name.substring(0, parsedPath.name.length - REBACKUP_TEMP_MARKER.length);
    return path.join(parsedPath.dir, `${finalName}${parsedPath.ext}`);
}

function getPathComparisonKey(filePath) {
    try {
        return path.resolve(filePath || "").toLowerCase();
    } catch (error) {
        return String(filePath || "").toLowerCase();
    }
}

function getAudioEntryFormat(filePath) {
    const extension = path.extname(filePath || "").toLowerCase();
    if (extension === ".wav") {
        return "wav";
    }
    if (extension === ".mp3") {
        return "mp3";
    }
    return "";
}
function buildRebackupRecoveryEntry(tempPath, baseName) {
    const fileName = path.basename(tempPath || "");
    const lowerFileName = fileName.toLowerCase();
    const lowerBaseName = String(baseName || "").toLowerCase();
    const finalPath = getRebackupFinalPath(tempPath);

    if (!finalPath || !lowerFileName.startsWith(`${lowerBaseName}_`)) {
        return null;
    }

    if (lowerFileName.startsWith(`${lowerBaseName}_backup${REBACKUP_TEMP_MARKER.toLowerCase()}.`)) {
        return {
            kind: "video",
            path: tempPath,
            finalPath,
            trackNumber: 0,
            trackNumbers: [],
            name: fileName
        };
    }

    const trackNumbers = parseTrackNumbersFromFileName(fileName, baseName);
    if (!trackNumbers.length) {
        return null;
    }

    return {
        kind: "audio",
        path: tempPath,
        finalPath,
        oldFinalPath: getOppositeAudioFormatPath(finalPath),
        trackNumber: trackNumbers[0],
        trackNumbers,
        name: fileName
    };
}

function collectRebackupRecoveryEntries(folderPath, sequenceName, manifest, options) {
    const settings = options || {};
    const manifestOnly = settings.manifestOnly === true && !!manifest;
    const baseName = sanitizeSequenceName((manifest && manifest.baseName) || sequenceName);
    const entries = [];
    const entryIndexesByFinalPath = {};

    const addEntry = (entry, preferTemp) => {
        if (!entry || !entry.path) {
            return;
        }

        const derivedFinalPath = entry.finalPath || getRebackupFinalPath(entry.path) || entry.path;
        const tempExists = fileExists(entry.path) && !!getRebackupFinalPath(entry.path);
        const finalExists = fileExists(derivedFinalPath);
        if (!tempExists && !finalExists) {
            return;
        }

        const normalizedEntry = Object.assign({}, entry, {
            finalPath: derivedFinalPath,
            oldFinalPath: entry.oldFinalPath || "",
            path: tempExists ? entry.path : derivedFinalPath,
            name: path.basename(tempExists ? entry.path : derivedFinalPath)
        });
        const finalPathKey = getPathComparisonKey(derivedFinalPath);

        if (Object.prototype.hasOwnProperty.call(entryIndexesByFinalPath, finalPathKey)) {
            const existingIndex = entryIndexesByFinalPath[finalPathKey];
            if (preferTemp && tempExists) {
                entries[existingIndex] = normalizedEntry;
            }
            return;
        }

        entryIndexesByFinalPath[finalPathKey] = entries.length;
        entries.push(normalizedEntry);
    };

    if (manifest && Array.isArray(manifest.expectedFiles)) {
        manifest.expectedFiles.forEach((entry) => addEntry(entry, false));
    }

    if (!manifestOnly) {
        fs.readdirSync(folderPath, { withFileTypes: true })
            .filter((entry) => entry.isFile() && entry.name.toUpperCase().includes(REBACKUP_TEMP_MARKER))
            .forEach((entry) => {
                const tempPath = path.join(folderPath, entry.name);
                const recoveryEntry = buildRebackupRecoveryEntry(tempPath, baseName);
                if (recoveryEntry) {
                    recoveryEntry.prefer = true;
                }
                addEntry(recoveryEntry, true);
            });
    }

    return {
        baseName,
        entries,
        hasTempFiles: entries.some((entry) => fileExists(entry.path) && !!getRebackupFinalPath(entry.path))
    };
}

async function ensureRebackupTempFilesAreStable(entries) {
    const tempPaths = (entries || [])
        .map((entry) => entry && entry.path)
        .filter((entryPath) => fileExists(entryPath) && !!getRebackupFinalPath(entryPath));
    let previousSizes = {};

    tempPaths.forEach((tempPath) => {
        previousSizes[tempPath] = fs.statSync(tempPath).size;
    });

    for (let pass = 0; pass < 2; pass += 1) {
        await delay(REBACKUP_RECOVERY_STABLE_WAIT_MS);

        const nextSizes = {};
        let allStable = true;
        tempPaths.forEach((tempPath) => {
            if (!fileExists(tempPath)) {
                allStable = false;
                return;
            }

            const nextSize = fs.statSync(tempPath).size;
            nextSizes[tempPath] = nextSize;
            if (nextSize < 1 || nextSize !== previousSizes[tempPath]) {
                allStable = false;
            }
        });

        if (!allStable) {
            throw new Error(
                "A _REBKP_TEMP file is still being written. Wait for the export to finish, then run Align Existing again."
            );
        }

        previousSizes = nextSizes;
    }
}

function findBackupLeftovers(matchInfo) {
    const current = (matchInfo.audio || []).map(entry => entry.path);
    if (matchInfo.videoPath) current.push(matchInfo.videoPath);
    const ready = file => { try { return fs.statSync(file).isFile() && fs.statSync(file).size > 0; } catch (_) { return false; } };
    const descriptions = current.filter(ready).map(file => {
        const match = path.basename(file).match(/^(.+)_(BACKUP|Track(\d+(?:-\d+)*))\.(mp4|mov|mxf|avi|wav|mp3)$/i);
        return match && {file,base:match[1].toLowerCase(),role:match[2].toLowerCase(),tracks:match[3] ? match[3].split('-') : []};
    }).filter(Boolean);
    const currentKeys = new Set(current.map(getPathComparisonKey));
    const folders = Array.from(new Set(descriptions.map(entry => path.dirname(entry.file))));
    const leftovers = [];
    folders.forEach(folder => {
        fs.readdirSync(folder,{withFileTypes:true}).filter(entry => entry.isFile()).forEach(entry => {
            const file = path.join(folder,entry.name);
            if (currentKeys.has(getPathComparisonKey(file))) return;
            const match = entry.name.match(/^(.+)_(BACKUP|Track(\d+(?:-\d+)*))(_REBKP_OLD_\d+(?:_\d+)?|_REBKP_TEMP)?\.(mp4|mov|mxf|avi|wav|mp3)$/i);
            if (!match) return;
            const tracks = match[3] ? match[3].split('-') : [];
            const replacement = descriptions.find(item => item.base === match[1].toLowerCase() && (
                (match[4] && item.role === match[2].toLowerCase()) ||
                (tracks.length && item.tracks.length > tracks.length && tracks.every(track => item.tracks.includes(track)))
            ));
            if (replacement) leftovers.push({path:file,replacement:replacement.file});
        });
    });
    return leftovers;
}

async function cleanupDiscoveredBackupLeftovers(matchInfo) {
    const candidates = findBackupLeftovers(matchInfo);
    if (!candidates.length) return {deleted:[],retained:[]};
    const result = parseHostResult(await callHost(`exportBackup.releaseUnusedBackupLeftovers("${escapeForEvalScript(JSON.stringify(candidates.map(entry => entry.path)))}")`));
    if (!result || !result.ok) throw new Error(result && result.message || 'Could not verify backup leftover usage.');
    const allowed = new Set((result.safePaths || []).map(getPathComparisonKey));
    const deleted = [], retained = [];
    candidates.forEach(entry => {
        let replacementReady = false;
        try { replacementReady = fs.statSync(entry.replacement).size > 0; } catch (_) {}
        if (!allowed.has(getPathComparisonKey(entry.path)) || !replacementReady) { retained.push(entry.path); return; }
        const removed = deleteLocalFileNow(entry.path);
        (removed.ok ? deleted : retained).push(entry.path);
    });
    return {deleted,retained};
}

async function prepareAlignExistingCleanup(matchInfo) {
    const stalePaths = matchInfo && Array.isArray(matchInfo.staleAudioPaths) ? matchInfo.staleAudioPaths.slice() : [];
    const obsoleteAudioFiles = matchInfo && matchInfo.manifest && Array.isArray(matchInfo.manifest.obsoleteAudioFiles)
        ? matchInfo.manifest.obsoleteAudioFiles
        : [];
    obsoleteAudioFiles.forEach((obsoletePath) => {
        if (obsoletePath && stalePaths.indexOf(obsoletePath) < 0) {
            stalePaths.push(obsoletePath);
        }
    });
    const expectedFiles = [];

    if (matchInfo && matchInfo.videoPath) {
        expectedFiles.push({
            kind: "video",
            path: matchInfo.videoPath,
            finalPath: matchInfo.videoPath
        });
    }

    (matchInfo && Array.isArray(matchInfo.audio) ? matchInfo.audio : []).forEach((entry) => {
        expectedFiles.push({
            kind: "audio",
            path: entry.path,
            finalPath: entry.path,
            oldFinalPath: "",
            trackNumber: entry.trackNumber,
            trackNumbers: entry.trackNumbers,
            name: entry.name
        });
    });

    stalePaths.forEach((stalePath) => {
        expectedFiles.push({
            kind: "audio",
            path: stalePath,
            finalPath: stalePath,
            oldFinalPath: stalePath
        });
    });

    if (expectedFiles.length) {
        await prepareRebackupReplacement({ expectedFiles });
    }

    return { stalePaths };
}

async function prepareRebackupReplacement(manifest) {
    if (!manifest || !Array.isArray(manifest.expectedFiles) || !manifest.expectedFiles.length) {
        return;
    }

    if (!(await ensureHostLoaded())) {
        throw new Error("Could not load Premiere host script before replacing old backup files.");
    }

    const expectedFilesJson = JSON.stringify(manifest.expectedFiles);
    setStatus(
        "Preparing re-backup replacement.\n" +
        "Step 1: removing old backup clips and media from Premiere."
    );
    const result = await callHost(`exportBackup.prepareRebackupReplacement("${escapeForEvalScript(expectedFilesJson)}")`);
    const parsed = parseHostResult(result);
    if (!parsed || parsed.ok === false) {
        const remainingProjectPaths = parsed && Array.isArray(parsed.remainingProjectPaths)
            ? parsed.remainingProjectPaths
            : (parsed && Array.isArray(parsed.remainingOnlinePaths) ? parsed.remainingOnlinePaths : []);
        const remainingPaths = remainingProjectPaths.length
            ? `\nStill present in the Premiere project:\n${remainingProjectPaths.join("\n")}`
            : "";
        throw new Error(
            `${(parsed && parsed.message) || "Could not prepare old backup media for replacement."}${remainingPaths}\n` +
            "Easy fix: save the project, close any Source Monitor/reference using that media, then run Align Existing to retry."
        );
    }

    const remainingProjectPaths = Array.isArray(parsed.remainingProjectPaths)
        ? parsed.remainingProjectPaths
        : [];
    setStatus(
        (remainingProjectPaths.length
            ? "Premiere replacement preparation complete; old ProjectItem cleanup will retry after import.\n"
            : "Premiere cleanup complete.\n") +
        `Timeline clips removed: ${parseInt(parsed.removedTimelineClips, 10) || 0}.\n` +
        `Project items removed/offlined: ${(parseInt(parsed.removedProjectItems, 10) || 0) + (parseInt(parsed.offlinedProjectItems, 10) || 0)}.\n` +
        "Step 2: importing completed replacement files."
    );

    manifest.rebackupReplacementPrepared = true;
    manifest.premiereCleanupPendingPaths = remainingProjectPaths;
    manifest.rebackupCleanup = {
        removedTimelineClips: parseInt(parsed.removedTimelineClips, 10) || 0,
        removedProjectItems: parseInt(parsed.removedProjectItems, 10) || 0,
        offlinedProjectItems: parseInt(parsed.offlinedProjectItems, 10) || 0
    };
    return parsed;
}

async function retryPremierePreservedMediaRelease(manifest) {
    const expectedFiles = manifest && Array.isArray(manifest.expectedFiles) ? manifest.expectedFiles : [];
    const cleanupFiles = expectedFiles
        .filter((entry) => entry && Array.isArray(entry.preservedPaths) && entry.preservedPaths.length)
        .map((entry) => ({
            kind: entry.kind,
            preservedPaths: entry.preservedPaths.slice()
        }));
    const recordedPendingPaths = manifest && Array.isArray(manifest.premiereCleanupPendingPaths)
        ? manifest.premiereCleanupPendingPaths
        : [];
    const representedKeys = {};
    cleanupFiles.forEach((entry) => {
        entry.preservedPaths.forEach((filePath) => {
            representedKeys[getPathComparisonKey(filePath)] = true;
        });
    });
    recordedPendingPaths.forEach((filePath) => {
        const key = getPathComparisonKey(filePath);
        if (key && !representedKeys[key]) {
            representedKeys[key] = true;
            const extension = path.extname(filePath || "").toLowerCase();
            cleanupFiles.push({
                kind: extension === ".mp4" ? "video" : "audio",
                preservedPaths: [filePath]
            });
        }
    });

    const attemptedPaths = getUniqueCleanupPaths(
        cleanupFiles.reduce((allPaths, entry) => allPaths.concat(entry.preservedPaths || []), [])
    );
    if (!cleanupFiles.length) {
        return { ok: true, attempted: false, remainingProjectPaths: [] };
    }

    try {
        if (!(await ensureHostLoaded())) {
            return {
                ok: false,
                attempted: false,
                remainingProjectPaths: attemptedPaths,
                message: "Could not load Premiere host cleanup script."
            };
        }

        setStatus("New files are aligned. Releasing preserved old Premiere references before optional file cleanup...");
        const cleanupFilesJson = JSON.stringify(cleanupFiles);
        const result = await callHost(`exportBackup.prepareRebackupReplacement("${escapeForEvalScript(cleanupFilesJson)}")`);
        const parsed = parseHostResult(result);
        const remainingProjectPaths = parsed && parsed.ok !== false
            ? (Array.isArray(parsed.remainingProjectPaths) ? parsed.remainingProjectPaths : [])
            : attemptedPaths;

        manifest.premiereCleanupPendingPaths = remainingProjectPaths.slice();
        return {
            ok: !!parsed && parsed.ok !== false,
            attempted: true,
            remainingProjectPaths,
            message: parsed && parsed.message ? parsed.message : "Premiere cleanup retry did not return a result."
        };
    } catch (error) {
        return {
            ok: false,
            attempted: true,
            remainingProjectPaths: attemptedPaths,
            message: error && error.message ? error.message : String(error || "Premiere cleanup retry failed.")
        };
    }
}

function getUniqueCleanupPaths(filePaths) {
    const uniquePaths = [];
    const seen = {};
    (filePaths || []).forEach((filePath) => {
        const normalizedPath = normalizeLocalFilePath(filePath);
        const key = getPathComparisonKey(normalizedPath);
        if (normalizedPath && key && !seen[key]) {
            seen[key] = true;
            uniquePaths.push(normalizedPath);
        }
    });
    return uniquePaths;
}

async function attemptPostAlignmentCleanup(manifest, stalePaths, retryDelays) {
    const cleanupErrors = {};
    let deletedOldFileCount = 0;
    let premiereCleanupPendingPaths = [];
    let rebackupPendingPaths = [];

    if (manifest && manifest.rebackup === true) {
        const premiereCleanupRetry = await retryPremierePreservedMediaRelease(manifest);
        premiereCleanupPendingPaths = getUniqueCleanupPaths(premiereCleanupRetry.remainingProjectPaths);
        if (!premiereCleanupRetry.ok) {
            cleanupErrors.premiere = premiereCleanupRetry.message;
        }

        const rebackupCleanupResult = await cleanupRebackupPreservedFilesAfterAlignment(
            manifest,
            retryDelays,
            premiereCleanupPendingPaths
        );
        deletedOldFileCount += rebackupCleanupResult.deleted.length;
        rebackupPendingPaths = rebackupCleanupResult.pending.slice();
        Object.assign(cleanupErrors, rebackupCleanupResult.errors);
    }

    const staleCleanupResult = await cleanupLocalFilesBestEffort(
        stalePaths,
        "older audio-format files",
        retryDelays
    );
    deletedOldFileCount += staleCleanupResult.deleted.length;
    Object.assign(cleanupErrors, staleCleanupResult.errors);

    const pendingPaths = getUniqueCleanupPaths(rebackupPendingPaths.concat(staleCleanupResult.pending));
    return {
        deletedOldFileCount,
        pendingPaths,
        stalePendingPaths: staleCleanupResult.pending.slice(),
        premiereCleanupPendingPaths,
        pendingCount: pendingPaths.length + premiereCleanupPendingPaths.length,
        errors: cleanupErrors
    };
}

function recordPendingCleanup(manifest, cleanupSummary) {
    if (!manifest) {
        return;
    }

    manifest.cleanupPendingPaths = cleanupSummary.pendingPaths.slice();
    manifest.cleanupErrors = cleanupSummary.errors;
    manifest.premiereCleanupPendingPaths = cleanupSummary.premiereCleanupPendingPaths.slice();
    manifest.rebackupFinalized = cleanupSummary.pendingCount === 0;
}

async function stopPendingCleanupRetry(closePrompt) {
    const state = cleanupRetryState;
    if (!state) {
        if (closePrompt !== false) {
            closeCleanupRetryPrompt();
        }
        return;
    }

    cleanupRetryState = null;
    if (state.timer) {
        clearTimeout(state.timer);
        state.timer = null;
    }

    const activeAttempt = state.currentAttempt;
    if (activeAttempt) {
        try {
            await activeAttempt;
        } catch (error) {}
    }

    if (closePrompt !== false) {
        closeCleanupRetryPrompt();
    }
}

function queuePendingCleanupRetry(state) {
    if (!state || cleanupRetryState !== state) {
        return;
    }

    if (state.timer) {
        clearTimeout(state.timer);
    }
    state.timer = setTimeout(() => {
        if (cleanupRetryState !== state) {
            return;
        }
        state.currentAttempt = runPendingCleanupRetryAttempt(state);
    }, CLEANUP_RETRY_INTERVAL_MS);
}

async function runPendingCleanupRetryAttempt(state) {
    if (!state || cleanupRetryState !== state || state.running) {
        return;
    }

    state.running = true;
    state.attempt += 1;
    if (!state.dialogDismissed) {
        showCleanupRetryPrompt(
            "ALIGN EXISTING IN PROGRESS",
            `Removing the currently aligned backup clips and importing them again.\n\nAutomatic attempt ${state.attempt}.\nIf this is taking time, you can press Align Existing between attempts.`,
            "pending"
        );
    }
    setStatus(
        `Automatic Align Existing retry ${state.attempt}.\n` +
        "Removing aligned backup clips, importing the files again, and checking old-file cleanup."
    );

    try {
        const retryContext = {};
        const alignmentSucceeded = await runAlignmentFlow(state.folderPath, {
            manifest: state.manifest,
            manifestOnly: !!state.manifest,
            skipVideo: false,
            sortProjectFiles: false,
            autoTriggered: true,
            cleanupRetryContext: retryContext
        });
        if (cleanupRetryState !== state) {
            return;
        }
        if (!alignmentSucceeded) {
            throw new Error(retryContext.error || "Automatic Align Existing retry did not finish.");
        }

        const cleanupSummary = retryContext.cleanupSummary || {
            stalePendingPaths: [],
            pendingCount: 0
        };
        const pendingCount = parseInt(retryContext.pendingCleanupCount, 10) || 0;
        state.stalePaths = Array.isArray(cleanupSummary.stalePendingPaths)
            ? cleanupSummary.stalePendingPaths.slice()
            : [];

        if (pendingCount === 0) {
            cleanupRetryState = null;
            state.timer = null;
            setStatus(
                `${retryContext.successTitle || state.successTitle}\n` +
                `Align Existing retry completed after ${state.attempt} automatic attempt(s).`,
                "success"
            );
            showCleanupRetryPrompt(
                "ALIGNMENT AND CLEANUP DONE",
                "The backup files were imported and aligned again, and all old cleanup items are finished.",
                "success"
            );
            return;
        }

        setStatus(
            `${retryContext.successTitle || state.successTitle}\n` +
            `Files were imported and aligned again. Cleanup items remaining: ${pendingCount}.\n` +
            "Running Align Existing again in 3 seconds.",
            "success"
        );
        if (!state.dialogDismissed) {
            showCleanupRetryPrompt(
                "ALIGN EXISTING IS RETRYING",
                `The backup files were imported and aligned again.\n\n${pendingCount} cleanup item(s) remain. Running Align Existing again in 3 seconds.`,
                "pending"
            );
        }
        queuePendingCleanupRetry(state);
    } catch (error) {
        if (cleanupRetryState !== state) {
            return;
        }
        const message = error && error.message ? error.message : String(error || "Automatic Align Existing retry failed.");
        setStatus(
            `${state.successTitle}\n` +
            `Automatic Align Existing retry ${state.attempt} could not finish.\n` +
            `Retrying again in 3 seconds.\n${message}`,
            "success"
        );
        if (!state.dialogDismissed) {
            showCleanupRetryPrompt(
                "ALIGN EXISTING IS RETRYING",
                `The last automatic Align Existing attempt could not finish. Retrying again in 3 seconds.\n\n${message}`,
                "pending"
            );
        }
        queuePendingCleanupRetry(state);
    } finally {
        state.running = false;
        state.currentAttempt = null;
    }
}
async function startPendingCleanupRetry(manifest, stalePaths, successTitle, pendingCount, folderPath) {
    await stopPendingCleanupRetry(false);
    const state = {
        manifest: manifest || null,
        folderPath: folderPath || (manifest && manifest.folderPath) || exportFolder || alignFolder || "",
        stalePaths: getUniqueCleanupPaths(stalePaths),
        successTitle: successTitle || "Alignment done.",
        pendingCount: parseInt(pendingCount, 10) || 0,
        attempt: 0,
        timer: null,
        currentAttempt: null,
        running: false,
        dialogDismissed: false
    };
    cleanupRetryState = state;

    showCleanupRetryPrompt(
        "ALIGN EXISTING WILL RETRY",
        `The first import and alignment finished, but ${state.pendingCount} cleanup item(s) remain.\n\nIn 3 seconds, Backup Project will remove the aligned backup clips, import the files again, and restore their positions.`,
        "pending"
    );
    queuePendingCleanupRetry(state);
}
async function recoverRebackupTempFiles(folderPath, sequenceName, manifest, options) {
    const recoveryInfo = collectRebackupRecoveryEntries(folderPath, sequenceName, manifest, options);
    if (!recoveryInfo.hasTempFiles) {
        return {
            recovered: false,
            manifest: manifest || null
        };
    }

    await ensureRebackupTempFilesAreStable(recoveryInfo.entries);

    const recoveryManifest = Object.assign({
        version: 5,
        createdAt: new Date().toISOString(),
        folderPath,
        sequenceName,
        baseName: recoveryInfo.baseName,
        backupVideoTrackNumber: getPositiveIntValue("exportVideoTrackInput", DEFAULT_BACKUP_VIDEO_TRACK),
        audioFormat: getSelectedAudioFormat(),
        exportMode: getSelectedExportMode(),
        rebackup: true,
        rebackupLayout: null,
        projectName: "",
        projectPath: ""
    }, manifest || {});

    recoveryManifest.version = Math.max(5, parseInt(recoveryManifest.version, 10) || 0);
    recoveryManifest.folderPath = folderPath;
    recoveryManifest.sequenceName = recoveryManifest.sequenceName || sequenceName;
    recoveryManifest.baseName = recoveryManifest.baseName || recoveryInfo.baseName;
    recoveryManifest.rebackup = true;
    recoveryManifest.rebackupRecovery = true;
    recoveryManifest.expectedFiles = recoveryInfo.entries;
    recoveryManifest.manifestPath = writeExportManifest(recoveryManifest);

    await prepareRebackupReplacement(recoveryManifest);
    await finalizeRebackupFiles(recoveryManifest);
    recoveryManifest.manifestPath = writeExportManifest(recoveryManifest);

    return {
        recovered: true,
        manifest: recoveryManifest
    };
}

function updateAlignFolder(folderPath) {
    alignFolder = folderPath;
    const alignPathElement = document.getElementById("alignPath");
    if (alignPathElement) {
        alignPathElement.textContent = folderPath || "No folder selected yet.";
    }

    try {
        if (folderPath) {
            localStorage.setItem(ALIGN_FOLDER_STORAGE_KEY, folderPath);
        }
    } catch (error) {}
}

function getBackupClipOwners() {
    try {
        const owners = JSON.parse(localStorage.getItem('exportbackup.clipOwners.v1') || '{}');
        return owners && typeof owners === 'object' && !Array.isArray(owners) ? owners : {};
    }
    catch (error) { return {}; }
}

function rememberBackupClipOwners(parsed) {
    if (!parsed || !parsed.sequenceID || !parsed.projectPath || !Array.isArray(parsed.queuedFiles)) return;
    const owners = getBackupClipOwners();
    parsed.queuedFiles.forEach(entry => {
        const filePath = entry.finalPath || entry.path;
        if (filePath) owners[getPathComparisonKey(filePath)] = Object.assign({}, owners[getPathComparisonKey(filePath)] || {}, {
            path: filePath, projectPath: parsed.projectPath, sequenceID: String(parsed.sequenceID)
        }, entry.exportRange || parsed.exportRange ? {exportRange:entry.exportRange || parsed.exportRange} : {});
    });
    localStorage.setItem('exportbackup.clipOwners.v1', JSON.stringify(owners));
}

async function confirmLegacyBackupNames(paths) {
    const result = parseHostResult(await callHost(`exportBackup.getBackupRenameCandidates("${escapeForEvalScript(JSON.stringify(paths))}","${escapeForEvalScript(JSON.stringify(getBackupClipOwners()))}")`));
    if (!result || !result.ok) throw new Error((result && result.message) || 'Could not verify backup names.');
    const candidates = result.candidates || [];
    if (!candidates.length) return;
    const message = `Do these full-length backup files belong to this sequence?\n\n${result.sequenceName}\n\n${candidates.map(entry => entry.path + '\n→ ' + entry.targetName).join('\n\n')}\n\nRename files changes their names in Explorer and relinks Premiere. Keep names continues alignment without renaming them.`;
    if (!await showReadablePrompt({title:'Confirm backup files', message, confirmText:'Rename files', cancelText:'Keep names'})) return;
    rememberBackupClipOwners({sequenceID:result.sequenceID,projectPath:result.projectPath,
        queuedFiles:candidates.map(entry => ({finalPath:entry.path}))});
}

function applyBackupFileRenames(matchInfo, renamedFiles) {
    if (!renamedFiles || !renamedFiles.length) return;
    const owners = getBackupClipOwners();
    const replacements = {};
    renamedFiles.forEach(entry => {
        const oldKey = getPathComparisonKey(entry.oldPath);
        replacements[oldKey] = entry.newPath;
        const ownerKey = Object.keys(owners).find(key => getPathComparisonKey(owners[key].path) === oldKey);
        if (ownerKey) {
            owners[getPathComparisonKey(entry.newPath)] = Object.assign({}, owners[ownerKey], {path:entry.newPath});
            if (ownerKey !== getPathComparisonKey(entry.newPath)) delete owners[ownerKey];
        }
    });
    const replacePaths = value => {
        if (typeof value === 'string') return replacements[getPathComparisonKey(value)] || value;
        if (value && typeof value === 'object') Object.keys(value).forEach(key => { value[key] = replacePaths(value[key]); });
        return value;
    };
    replacePaths(matchInfo);
    localStorage.setItem('exportbackup.clipOwners.v1', JSON.stringify(owners));
}

function createExportManifestFromHostResult(parsed) {
    return {
        version: 5,
        createdAt: new Date().toISOString(),
        folderPath: exportFolder,
        sequenceName: parsed.sequenceName || "",
        baseName: parsed.baseName || sanitizeSequenceName(parsed.sequenceName || "Active_Sequence"),
        backupVideoTrackNumber: parseInt(parsed.backupVideoTrackNumber, 10) || DEFAULT_BACKUP_VIDEO_TRACK,
        audioFormat: parsed.audioFormat || getSelectedAudioFormat(),
        exportMode: parsed.exportMode || getSelectedExportMode(),
        rebackup: parsed.rebackup === true,
        rebackupPrepared: parsed.rebackupPrepared === true,
        rebackupLayout: parsed.rebackupLayout || null,
        exportRange: parsed.exportRange || null,
        expectedFiles: Array.isArray(parsed.queuedFiles) ? parsed.queuedFiles : [],
        obsoleteAudioFiles: Array.isArray(parsed.obsoleteAudioFiles) ? parsed.obsoleteAudioFiles : [],
        manifestPath: "",
        projectName: parsed.projectName || "",
        projectPath: parsed.projectPath || ""
    };
}

function clearExportCompletionMonitor() {
    if (!exportMonitorState) {
        return;
    }

    if (exportMonitorState.timer) {
        clearTimeout(exportMonitorState.timer);
    }

    exportMonitorState = null;
}

function getCompletionSummary(state) {
    const expectedFiles = state.manifest.expectedFiles || [];
    let completed = 0;

    expectedFiles.forEach((entry) => {
        if (entry && state.stableCounts[entry.path] >= EXPORT_MONITOR_STABLE_PASSES) {
            completed += 1;
        }
    });

    return `${completed}/${expectedFiles.length}`;
}

function copyProjectToFolder(projectPath, destinationFolder) {
    try {
        if (!projectPath) {
            return { ok: false, message: "Premiere saved the project, but its path could not be read for copying." };
        }

        if (!fileExists(projectPath)) {
            return { ok: false, message: `Premiere saved the project, but the file was not found for copying.\n${projectPath}` };
        }

        const destinationPath = path.join(destinationFolder, path.basename(projectPath));
        if (path.resolve(destinationPath).toLowerCase() !== path.resolve(projectPath).toLowerCase()) {
            fs.copyFileSync(projectPath, destinationPath);
        }

        return { ok: true, destinationPath };
    } catch (error) {
        return { ok: false, message: `Project alignment finished, but the project copy could not be created.\n${error.message}` };
    }
}

async function renameLegacyBackupFilesForSequence(folderPath, sequenceName) {
    const raw = await callHost(
        `exportBackup.findLegacyBackupForActiveSequence("${escapeForEvalScript(folderPath)}")`
    );
    const result = parseHostResult(raw);
    if (!result || result.ok === false) {
        throw new Error((result && result.message) || "Could not inspect existing backup names in Premiere.");
    }
    if (!result.found || !result.candidate || !Array.isArray(result.candidate.files)) {
        return { renamedCount: 0, message: result.message || "" };
    }

    const sanitizedBase = sanitizeSequenceName(sequenceName);
    const legacyLayout = result.candidate.layout || null;
    const folderKey = `${path.resolve(folderPath).toLowerCase()}${path.sep}`;
    const operations = [];
    result.candidate.files.forEach((entry) => {
        if (!entry || !entry.path) {
            return;
        }
        const sourcePath = path.resolve(entry.path);
        if (!fileExists(sourcePath)) {
            return;
        }
        if (`${sourcePath.toLowerCase()}${path.sep}`.indexOf(folderKey) !== 0) {
            throw new Error(`Refusing to rename a backup file outside the selected folder:\n${sourcePath}`);
        }
        const extension = path.extname(sourcePath);
        const trackNumbers = Array.isArray(entry.trackNumbers)
            ? entry.trackNumbers.map((value) => parseInt(value, 10) || 0).filter((value) => value > 0)
            : [];
        const targetName = entry.kind === "video"
            ? `${sanitizedBase}_BACKUP${extension}`
            : `${sanitizedBase}_Track${trackNumbers.join("-")}${extension}`;
        if (entry.kind === "audio" && !trackNumbers.length) {
            return;
        }
        const targetPath = path.join(folderPath, targetName);
        if (getPathComparisonKey(sourcePath) === getPathComparisonKey(targetPath)) {
            return;
        }
        if (fileExists(targetPath)) {
            return;
        }
        operations.push({ sourcePath, targetPath, kind: entry.kind });
    });

    if (!operations.length) {
        return { renamedCount: 0, message: "Matching backup files already use the active sequence name." };
    }

    await prepareRebackupReplacement({
        expectedFiles: operations.map((operation) => ({
            kind: operation.kind,
            path: operation.sourcePath,
            finalPath: operation.sourcePath,
            sourceMediaPath: operation.sourcePath
        }))
    });

    const completed = [];
    try {
        operations.forEach((operation) => {
            fs.renameSync(operation.sourcePath, operation.targetPath);
            completed.push(operation);
        });
    } catch (error) {
        completed.reverse().forEach((operation) => {
            try {
                if (fileExists(operation.targetPath) && !fileExists(operation.sourcePath)) {
                    fs.renameSync(operation.targetPath, operation.sourcePath);
                }
            } catch (rollbackError) {}
        });
        try {
            const legacyVideo = result.candidate.files.find((entry) => entry && entry.kind === "video");
            const legacyAudio = result.candidate.files
                .filter((entry) => entry && entry.kind === "audio")
                .map((entry) => ({
                    path: entry.path,
                    trackNumber: parseInt((entry.trackNumbers || [])[0], 10) || 0,
                    trackNumbers: entry.trackNumbers || [],
                    name: path.basename(entry.path || "")
                }));
            const restoreScript = `exportBackup.alignMappedFiles(` +
                `"${escapeForEvalScript(legacyVideo ? legacyVideo.path : "")}",` +
                `"${escapeForEvalScript(JSON.stringify(legacyAudio))}",` +
                `${getPositiveIntValue("exportVideoTrackInput", DEFAULT_BACKUP_VIDEO_TRACK)},false,` +
                `"${escapeForEvalScript(JSON.stringify(legacyLayout))}")`;
            await callHost(restoreScript);
        } catch (restoreError) {}
        throw new Error(`Could not rename all matching backup files. Completed renames were rolled back.\n${error.message}`);
    }

    return {
        renamedCount: completed.length,
        legacyLayout,
        message: `Renamed ${completed.length} backup file(s) to match sequence "${sanitizedBase}".`
    };
}

async function copyExistingBackupsToResolvedLocation() {
    if (busy) {
        return;
    }
    setBusyState(true);
    setStatus("Reading existing backup files from the active sequence...");
    const copiedTargets = [];
    try {
        if (!(await ensureHostLoaded())) {
            throw new Error("Could not load Premiere host script.");
        }
        const sequenceName = await getActiveSequenceName();
        if (!sequenceName) {
            throw new Error("No active sequence is open in Premiere Pro.");
        }

        const destination = await resolveProjectBackupFolder({ create: true });
        const destinationFolder = destination.folderPath;
        const sourcePaths = [];
        const addLayoutPath = (entry) => {
            if (!entry) {
                return;
            }
            const candidate = entry.currentMediaPath || entry.mediaPath || "";
            if (candidate && fileExists(candidate)) {
                sourcePaths.push(candidate);
            }
        };

        const rawLayout = await callHost("exportBackup.getActiveBackupLayout()");
        const activeBackup = parseHostResult(rawLayout);
        if (!activeBackup || activeBackup.ok === false) {
            throw new Error((activeBackup && activeBackup.message) || "Could not read existing backups from Premiere.");
        }
        const layout = activeBackup.layout || {};
        addLayoutPath(layout.video);
        addLayoutPath(layout.backupAudio);
        (layout.audioOutputs || []).forEach(addLayoutPath);

        if (!sourcePaths.length) {
            const rawLegacy = await callHost(
                "exportBackup.findLegacyBackupForActiveSequence(\"\", true)"
            );
            const legacy = parseHostResult(rawLegacy);
            if (!legacy || legacy.ok === false) {
                throw new Error((legacy && legacy.message) || "Could not inspect existing backup files.");
            }
            if (legacy.ambiguous) {
                throw new Error("Multiple backup sets match the active sequence duration, so nothing was copied automatically.");
            }
            if (legacy.found && legacy.candidate && Array.isArray(legacy.candidate.files)) {
                legacy.candidate.files.forEach((entry) => {
                    if (entry && entry.path) {
                        sourcePaths.push(entry.path);
                    }
                });
            }
        }

        const uniqueSourcePaths = Array.from(new Set(sourcePaths.map((filePath) => path.resolve(filePath))));
        if (!uniqueSourcePaths.length) {
            throw new Error("No matching backup files are currently used by the active sequence.");
        }

        let alreadyAtDestination = 0;
        uniqueSourcePaths.forEach((sourcePath) => {
            const targetPath = path.join(destinationFolder, path.basename(sourcePath));
            if (!fileExists(sourcePath)) {
                throw new Error(`Existing backup file was not found:\n${sourcePath}`);
            }
            if (getPathComparisonKey(sourcePath) === getPathComparisonKey(targetPath)) {
                alreadyAtDestination += 1;
                return;
            }
            if (fileExists(targetPath)) {
                if (fs.statSync(sourcePath).size === fs.statSync(targetPath).size) {
                    alreadyAtDestination += 1;
                    return;
                }
                throw new Error(`A different backup file with the same name already exists at the destination:\n${targetPath}`);
            }
            fs.copyFileSync(sourcePath, targetPath);
            if (fs.statSync(sourcePath).size !== fs.statSync(targetPath).size) {
                throw new Error(`Copied backup verification failed:\n${targetPath}`);
            }
            copiedTargets.push(targetPath);
        });

        exportFolder = destinationFolder;
        updateAlignFolder(destinationFolder);
        document.getElementById("exportPath").textContent = destinationFolder;
        setBusyState(false);
        const resultMessage = copiedTargets.length
            ? `Copied ${copiedTargets.length} existing backup file(s).`
            : `All ${alreadyAtDestination} matching backup file(s) are already at the destination.`;
        setStatus(`${resultMessage}\nDestination: ${destinationFolder}`, "success");
        showBlockingMessage(`${resultMessage}\n\nDestination:\n${destinationFolder}`, 'success');
    } catch (error) {
        copiedTargets.forEach((targetPath) => {
            try {
                if (fileExists(targetPath)) {
                    fs.unlinkSync(targetPath);
                }
            } catch (cleanupError) {}
        });
        setBusyState(false);
        setStatus(error.message, "error");
        showBlockingMessage(error.message);
    }
}

async function runAlignmentFlow(folderPath, options) {
    const settings = options || {};
    const cleanupRetryContext = settings.cleanupRetryContext || null;
    const reportAlignmentFailure = (message) => {
        const resolvedMessage = String(message || "Alignment failed.");
        if (cleanupRetryContext) {
            cleanupRetryContext.error = resolvedMessage;
            setStatus(
                `Automatic Align Existing retry could not finish.\n${resolvedMessage}\n` +
                "Retrying again in 3 seconds.",
                "success"
            );
            return false;
        }
        showAlignmentRecoveryError(resolvedMessage);
        return false;
    };

    if (!folderPath) {
        if (cleanupRetryContext) {
            return reportAlignmentFailure("No export folder is available for automatic Align Existing retry.");
        }
        showBlockingMessage("Choose an export folder first.");
        return false;
    }

    setBusyState(true);
    setStatus(settings.autoTriggered
        ? "Export finished. Importing and aligning backup files..."
        : "Loading Premiere host script...");

    try {
        if (!(await ensureHostLoaded())) {
            return reportAlignmentFailure("Could not load Premiere host script.");
        }

        const activeSequenceName = await getActiveSequenceName();
        if (!activeSequenceName) {
            return reportAlignmentFailure("No active sequence is open in Premiere Pro.");
        }

        let legacyRenameLayout = null;
        if (!settings.manifestOnly && !settings.cleanupRetryContext) {
            const renameResult = await renameLegacyBackupFilesForSequence(folderPath, activeSequenceName);
            legacyRenameLayout = renameResult.legacyLayout || null;
            if (renameResult.renamedCount > 0) {
                setStatus(`${renameResult.message}\nImporting and aligning the renamed files...`);
            }
        }

        let manifest = settings.manifest || readManifestForSequence(folderPath, activeSequenceName);
        if (!manifest && legacyRenameLayout) {
            manifest = {
                rebackup: true,
                rebackupLayout: legacyRenameLayout,
                backupVideoTrackNumber: legacyRenameLayout.video
                    ? parseInt(legacyRenameLayout.video.targetTrackNumber, 10) || DEFAULT_BACKUP_VIDEO_TRACK
                    : DEFAULT_BACKUP_VIDEO_TRACK,
                expectedFiles: [],
                obsoleteAudioFiles: []
            };
        }
        const manifestOnly = settings.manifestOnly === true && !!manifest;

        if (
            manifest &&
            manifest.rebackup === true &&
            manifest.rebackupPrepared === true &&
            manifest.rebackupReplacementPrepared !== true &&
            getManifestPreservedPaths(manifest).length
        ) {
            assertExpectedExportsAreReady(manifest);
            setStatus("New Re-backup files found. Releasing selected old Premiere clips before alignment...");
            await prepareRebackupReplacement(manifest);
            await finalizeRebackupFiles(manifest);
            manifest.manifestPath = writeExportManifest(manifest);
        }

        const recoveryResult = await recoverRebackupTempFiles(folderPath, activeSequenceName, manifest, { manifestOnly });
        if (recoveryResult.recovered) {
            manifest = recoveryResult.manifest;
            setStatus("Recovered _REBKP_TEMP files. Importing and restoring the backup media...");
        }
        const matchInfo = scanExportFolderForSequence(folderPath, activeSequenceName, manifest, { manifestOnly });

        if (matchInfo.manifest && matchInfo.manifest.backupVideoTrackNumber) {
            applyAlignDefaults({ videoTrackNumber: matchInfo.manifest.backupVideoTrackNumber }, true);
        }

        if (!matchInfo.videoPath && matchInfo.audio.length === 0) {
            const message =
                "No files could be matched in the chosen folder.\n" +
                `Sequence base: ${matchInfo.baseName}\n` +
                `Folder files: ${matchInfo.folderFiles.join(" | ")}`;
            return reportAlignmentFailure(message);
        }

        if (settings.autoTriggered === false) {
            await confirmLegacyBackupNames([matchInfo.videoPath].concat(matchInfo.audio.map(entry => entry.path)).filter(Boolean));
        }
        const alignmentCleanup = await prepareAlignExistingCleanup(matchInfo);

        const backupVideoTrackNumber = getPositiveIntValue(
            "exportVideoTrackInput",
            (matchInfo.manifest && parseInt(matchInfo.manifest.backupVideoTrackNumber, 10)) || DEFAULT_BACKUP_VIDEO_TRACK
        );
        saveAlignVideoTrack(backupVideoTrackNumber);
        const skipVideoCheckbox = document.getElementById("alignSkipVideoCheckbox");
        const skipBackupVideo = settings.skipVideo === true || (skipVideoCheckbox && skipVideoCheckbox.checked);
        const sortProjectFiles = settings.sortProjectFiles === true || document.getElementById("alignSortProjectFilesCheckbox").checked;
        const resolvedVideoPath = skipBackupVideo ? "" : (matchInfo.videoPath || "");
        const audioJson = JSON.stringify(matchInfo.audio);
        const rebackupLayoutJson = JSON.stringify(
            matchInfo.manifest && matchInfo.manifest.rebackup === true
                ? (matchInfo.manifest.rebackupLayout || null)
                : null
        );
        const ownersJson = JSON.stringify(getBackupClipOwners());
        const videoEntry = matchInfo.manifest && (matchInfo.manifest.expectedFiles || []).find(entry => entry.kind === 'video');
        const exportRangeJson = JSON.stringify(videoEntry && videoEntry.exportRange || matchInfo.manifest && matchInfo.manifest.exportRange || null);
        const script = `exportBackup.alignMappedFiles("${escapeForEvalScript(resolvedVideoPath)}","${escapeForEvalScript(audioJson)}",${backupVideoTrackNumber},${sortProjectFiles},"${escapeForEvalScript(rebackupLayoutJson)}","${escapeForEvalScript(ownersJson)}","${escapeForEvalScript(exportRangeJson)}")`;
        const result = await callHost(script);
        const parsed = parseHostResult(result);

        if (!parsed || parsed.ok === false) {
            return reportAlignmentFailure((parsed && parsed.message) || "Alignment failed.");
        }

        applyBackupFileRenames(matchInfo, parsed.renamedFiles);
        const cleanupSummary = await attemptPostAlignmentCleanup(
            matchInfo.manifest,
            alignmentCleanup && alignmentCleanup.stalePaths,
            [0]
        );
        let extraCleanup = {deleted:[],retained:[]};
        let extraCleanupError = '';
        try { extraCleanup = await cleanupDiscoveredBackupLeftovers(matchInfo); }
        catch (error) { extraCleanupError = error.message; }
        const deletedOldFileCount = cleanupSummary.deletedOldFileCount + extraCleanup.deleted.length;
        const premiereCleanupPendingCount = cleanupSummary.premiereCleanupPendingPaths.length;
        let pendingCleanupCount = cleanupSummary.pendingCount;
        const shouldCopyProject = !!(getCopyProjectFileCheckbox() && getCopyProjectFileCheckbox().checked);
        const copyResult = shouldCopyProject
            ? copyProjectToFolder(parsed.projectPath, folderPath)
            : { ok: false, message: "" };

        const successTitle = getCompletionStatusTitle(settings, matchInfo.manifest);
        const lines = [successTitle, parsed.message || "Alignment completed."];
        if (extraCleanup.retained.length) lines.push('Leftovers still used by a sequence or held by Premiere/Windows were kept:\n' + extraCleanup.retained.join('\n') + '\nRun Align Existing again after they are released.');
        if (extraCleanupError) lines.push('Additional leftover cleanup could not finish: ' + extraCleanupError + '\nRun Align Existing again to retry.');
        if (parsed.importBinName) {
            lines.push(`Imported backup files were added to project bin: ${parsed.importBinName}`);
        }
        if (copyResult.ok) {
            lines.push(`Project copy saved: ${copyResult.destinationPath}`);
        } else if (copyResult.message) {
            lines.push(copyResult.message);
        }
        if (deletedOldFileCount > 0) {
            lines.push(`Deleted old backup files after import: ${deletedOldFileCount}.`);
        }
        if (premiereCleanupPendingCount > 0) {
            lines.push(`Old Premiere ProjectItems still pending cleanup: ${premiereCleanupPendingCount}.`);
        }

        if (matchInfo.manifest) {
            recordPendingCleanup(matchInfo.manifest, cleanupSummary);

            if (cleanupSummary.pendingCount > 0) {
                try {
                    matchInfo.manifest.manifestPath = writeExportManifest(matchInfo.manifest);
                } catch (manifestError) {
                    pendingCleanupCount += 1;
                    lines.push(`Could not save the cleanup record yet: ${manifestError.message}`);
                }
            } else if (matchInfo.manifest.manifestPath) {
                const mapDeleteResult = deleteLocalFileNow(matchInfo.manifest.manifestPath);
                if (!mapDeleteResult.ok) {
                    pendingCleanupCount += 1;
                    lines.push(`Export map cleanup is pending: ${mapDeleteResult.code}: ${mapDeleteResult.error}`);
                }
            }
        }

        if (pendingCleanupCount > 0) {
            lines.push(`New files are imported and aligned. Old cleanup items pending: ${pendingCleanupCount}.`);
            lines.push("Cleanup will retry automatically every 3 seconds. If it is taking time, use Align Existing.");
        }
        setStatus(lines.join("\n"), "success");

        if (cleanupRetryContext) {
            cleanupRetryContext.cleanupSummary = cleanupSummary;
            cleanupRetryContext.pendingCleanupCount = pendingCleanupCount;
            cleanupRetryContext.successTitle = successTitle;
        }

        if (pendingCleanupCount > 0) {
            if (!cleanupRetryContext) {
                await startPendingCleanupRetry(
                    matchInfo.manifest,
                    cleanupSummary.stalePendingPaths,
                    successTitle,
                    pendingCleanupCount,
                    folderPath
                );
            }
        } else if (!cleanupRetryContext) {
            showResultPrompt(
                successTitle,
                (settings.autoTriggered ? 'Backup files were imported and aligned successfully.' : 'Existing backup files were imported and aligned successfully.') + '\n\n' + lines.slice(1).join('\n\n'),
                {success: !(parsed.renameWarnings && parsed.renameWarnings.length) && !extraCleanupError && !extraCleanup.retained.length && !copyResult.message}
            );
        }
        return true;
    } catch (error) {
        return reportAlignmentFailure(`Alignment failed.\n${error.message}`);
    } finally {
        setBusyState(false);
    }
}

function scheduleExportMonitorTick() {
    if (!exportMonitorState) {
        return;
    }

    exportMonitorState.timer = setTimeout(() => {
        monitorExportCompletion().catch((error) => {
            showAlignmentRecoveryError(`Automatic import stopped.\n${error.message}`);
            clearExportCompletionMonitor();
        });
    }, EXPORT_MONITOR_INTERVAL_MS);
}

async function monitorExportCompletion() {
    const state = exportMonitorState;
    if (!state) {
        return;
    }

    if ((Date.now() - state.startedAt) > EXPORT_MONITOR_TIMEOUT_MS) {
        showAlignmentRecoveryError(
            "Automatic import timed out while waiting for Media Encoder.\n" +
            "The queued exports are still in the chosen folder."
        );
        clearExportCompletionMonitor();
        return;
    }

    const expectedFiles = state.manifest.expectedFiles || [];
    if (!expectedFiles.length) {
        clearExportCompletionMonitor();
        return;
    }

    let allStable = true;

    expectedFiles.forEach((entry) => {
        if (!entry || !entry.path || !fileExists(entry.path)) {
            if (entry && entry.path) {
                state.lastSizes[entry.path] = -1;
                state.stableCounts[entry.path] = 0;
            }
            allStable = false;
            return;
        }

        const size = fs.statSync(entry.path).size;
        if (state.lastSizes[entry.path] === size && size > 0) {
            state.stableCounts[entry.path] = (state.stableCounts[entry.path] || 0) + 1;
        } else {
            state.stableCounts[entry.path] = 0;
        }

        state.lastSizes[entry.path] = size;
        if (state.stableCounts[entry.path] < EXPORT_MONITOR_STABLE_PASSES) {
            allStable = false;
        }
    });

    if (allStable) {
        clearExportCompletionMonitor();
        if (state.manifest.rebackup) {
            try {
                await prepareRebackupReplacement(state.manifest);
                await finalizeRebackupFiles(state.manifest);
                writeExportManifest(state.manifest);
            } catch (error) {
                if (isLockedFileRecoveryMessage(error.message)) {
                    showAlignmentRecoveryError(error.message);
                    return;
                }

                showAlignmentRecoveryError(
                    "Re-backup finished exporting, but replacing old files failed.\n" +
                    error.message
                );
                return;
            }
        }
        await runAlignmentFlow(state.manifest.folderPath, {
            manifest: state.manifest,
            manifestOnly: true,
            skipVideo: false,
            sortProjectFiles: false,
            autoTriggered: true
        });
        return;
    }

    setStatus(
        "Queued jobs were sent to Adobe Media Encoder.\n" +
        `Waiting for finished files: ${getCompletionSummary(state)} complete.\n` +
        `Folder: ${state.manifest.folderPath}`
    );
    scheduleExportMonitorTick();
}

function startExportCompletionMonitor(manifest) {
    clearExportCompletionMonitor();

    exportMonitorState = {
        manifest,
        startedAt: Date.now(),
        lastSizes: {},
        stableCounts: {},
        timer: null
    };

    setStatus(
        "Queued jobs were sent to Adobe Media Encoder.\n" +
        `Waiting for finished files: 0/${(manifest.expectedFiles || []).length} complete.\n` +
        `Folder: ${manifest.folderPath}`
    );
    scheduleExportMonitorTick();
}

async function chooseExportFolder() {
    if (busy) {
        return;
    }

    const result = window.cep.fs.showOpenDialogEx(false, true, "Choose Export Folder");
    if (result.data && result.data.length > 0) {
        manualExportFolder = result.data[0];
        exportFolder = manualExportFolder;
        alignFolder = exportFolder;
        const inputs = getBackupDestinationInputs();
        if (inputs.manual) {
            inputs.manual.checked = true;
        }
        updateDestinationButtonLabel();
        updateCategoryDestinationNote();
        try {
            localStorage.setItem(EXPORT_FOLDER_STORAGE_KEY, manualExportFolder);
            localStorage.setItem(ALIGN_FOLDER_STORAGE_KEY, alignFolder);
            localStorage.setItem(BACKUP_DESTINATION_STORAGE_KEY, BACKUP_DESTINATION_MANUAL);
        } catch (error) {}
        document.getElementById("exportPath").textContent = exportFolder;
        setStatus("Export folder selected. Ready.");
    }
}

async function chooseAlignFolder() {
    if (busy) {
        return;
    }

    const result = window.cep.fs.showOpenDialogEx(false, true, "Choose Existing Export Folder");
    if (result.data && result.data.length > 0) {
        updateAlignFolder(result.data[0]);

        try {
            const activeSequenceName = await getActiveSequenceName();
            const manifest = readManifestForSequence(alignFolder, activeSequenceName);
            if (manifest && manifest.backupVideoTrackNumber) {
                applyAlignDefaults({ videoTrackNumber: manifest.backupVideoTrackNumber }, true);
            }
        } catch (error) {}

        setStatus("Existing export folder selected. Ready.");
    }
}

async function chooseVideoPreset() {
    if (busy) {
        return;
    }

    const result = choosePresetFile("Choose Premiere Video Preset (.epr)", videoPresetPath);
    if (result.data && result.data.length > 0) {
        saveVideoPreset(result.data[0]);
        setStatus("Video preset updated. This choice will be remembered until you change it.");
    }
}

async function chooseMp3Preset() {
    if (busy) {
        return;
    }

    const result = choosePresetFile("Choose Premiere MP3 Preset (.epr)", mp3PresetPath);
    if (result.data && result.data.length > 0) {
        saveMp3Preset(result.data[0]);
        setStatus("MP3 preset updated. This choice will be remembered until you change it.");
    }
}

async function chooseWavPreset() {
    if (busy) {
        return;
    }

    const result = choosePresetFile("Choose Premiere WAV Preset (.epr)", wavPresetPath);
    if (result.data && result.data.length > 0) {
        saveWavPreset(result.data[0]);
        setStatus("WAV preset updated. This choice will be remembered until you change it.");
    }
}

async function resolveExportActionDestination(isRebackup) {
    if (!(await ensureHostLoaded())) throw new Error("Could not load Premiere host script.");
    const existing = parseHostResult(await callHost(`exportBackup.getActiveBackupLayout("${escapeForEvalScript(JSON.stringify(getBackupClipOwners()))}")`));
    if (!existing || !existing.ok) throw new Error(existing && existing.message || "Could not inspect existing backups.");
    const layout = existing.layout || {};
    const entries = [layout.video, layout.backupAudio].concat(layout.audioOutputs || []).filter(Boolean);
    const paths = Array.from(new Set(entries.map((entry) => entry.mediaPath || entry.currentMediaPath).filter(Boolean)));
    if (paths.length) {
        if (!isRebackup) {
            const error = new Error("Backup files already exist in the sequence or project.\n\nUse Re-backup to replace them.\n\nExisting files:\n" + paths.join("\n\n"));
            error.promptTitle = 'Backup files already exist';
            error.promptKind = 'warning';
            throw error;
        }
        const folderPath = path.dirname(paths[0]);
        if (!fs.existsSync(folderPath)) throw new Error("The existing backup folder is unavailable:\n" + folderPath);
        return {folderPath, existingBackup: existing};
    }
    if (isRebackup) throw new Error("No existing backup files were found for the active sequence. Use Backup to create the first backup.");
    return resolveProjectBackupFolder({ create: true });
}

async function runExport(isRebackup) {
    if (busy) {
        return;
    }

    await stopPendingCleanupRetry(true);

    try {
        await syncExportSelectionWithActiveTimeline();
    } catch (error) {
        await showBlockingMessage(error.message);
        setStatus(error.message, "error");
        return;
    }

    let destination;
    try {
        destination = await resolveExportActionDestination(isRebackup);
    } catch (error) {
        await showReadablePrompt({title:error.promptTitle || 'Could not start backup', message:error.message, kind:error.promptKind || 'error'});
        setStatus(error.message, "error");
        return;
    }

    exportFolder = destination.folderPath;
    updateAlignFolder(exportFolder);
    document.getElementById("exportPath").textContent = exportFolder;
    try {
        if (destination.destinationMode === BACKUP_DESTINATION_MANUAL) {
            localStorage.setItem(EXPORT_FOLDER_STORAGE_KEY, manualExportFolder);
        }
        if (destination.destinationMode) localStorage.setItem(BACKUP_DESTINATION_STORAGE_KEY, destination.destinationMode);
    } catch (error) {}

    const selectedAudioFormat = getSelectedAudioFormat();
    const selectedAudioPresetPath = selectedAudioFormat === "wav" ? wavPresetPath : mp3PresetPath;
    const backupVideoTrackNumber = getPositiveIntValue("exportVideoTrackInput", DEFAULT_BACKUP_VIDEO_TRACK);
    const autoEmptyTrack = !isRebackup && useAutoEmptyBackupTrack();
    const selectedQueueItems = getSelectedQueueItems();
    const removeSequenceMarkers = !!(getRemoveSequenceMarkersCheckbox() && getRemoveSequenceMarkersCheckbox().checked);
    const selectedExportMode = getSelectedExportMode();

    if (!fileExists(videoPresetPath)) {
        await showBlockingMessage("The selected video preset file was not found. Choose the video preset again.");
        return;
    }

    if (!fileExists(selectedAudioPresetPath)) {
        await showBlockingMessage(`The selected ${selectedAudioFormat.toUpperCase()} preset file was not found.`);
        return;
    }

    saveBackupVideoTrack(backupVideoTrackNumber);
    saveSelectedAudioFormat(selectedAudioFormat);
    saveRemoveSequenceMarkers(removeSequenceMarkers);
    saveSelectedExportMode(selectedExportMode);

    setBusyState(true);
    setStatus("Loading Premiere host script...");

    if (!(await ensureHostLoaded())) {
        showBlockingMessage("Could not load Premiere host script.");
        setBusyState(false);
        return;
    }

    let validation = await validateBackupExportSettings(backupVideoTrackNumber, selectedQueueItems, isRebackup, autoEmptyTrack);
    if (!validation.ok && validation.needsInOut === true) {
        const shouldAutoSetInOut = await showInOutPrompt();
        if (!shouldAutoSetInOut) {
            setStatus("Export cancelled. Set sequence In and Out manually, then start Backup or Re-backup again.");
            setBusyState(false);
            return;
        }

        setStatus("Setting In and Out to the full sequence range...");
        const inOutResult = await setActiveSequenceInOutToFullRange();
        if (!inOutResult.ok) {
            const message = inOutResult.message || "Could not set sequence In and Out.";
            setStatus(message, "error");
            showBlockingMessage(message);
            setBusyState(false);
            return;
        }

        setStatus("Sequence In and Out set. Starting export...");
        validation = await validateBackupExportSettings(backupVideoTrackNumber, selectedQueueItems, isRebackup, autoEmptyTrack);
    }
    const resolvedBackupVideoTrackNumber = parseInt(validation.backupVideoTrackNumber, 10) || backupVideoTrackNumber;
    if (!validation.ok) {
        const message = validation.hasConflicts && !isRebackup
            ? formatExistingMediaMessage(validation)
            : (validation.message || "Backup export validation failed.");
        showBlockingMessage(message);
        setStatus(message);
        setBusyState(false);
        return;
    }

    if (!isRebackup && !await confirmBackupCandidates(validation)) {
        setStatus("Backup cancelled. Review the listed clips before choosing Backup or Re-backup.");
        setBusyState(false);
        return;
    }

    if (autoEmptyTrack) {
        const backupTrackInput = getBackupVideoTrackInput();
        if (backupTrackInput) {
            backupTrackInput.value = String(resolvedBackupVideoTrackNumber);
            backupTrackInput.dataset.autoValue = String(resolvedBackupVideoTrackNumber);
        }
        saveBackupVideoTrack(resolvedBackupVideoTrackNumber);
    }
setStatus(selectedExportMode === EXPORT_MODE_PREMIERE
        ? (isRebackup ? "Rendering checked re-backup files in Premiere Pro...\nExisting backup clips stay until export finishes." : `Rendering backup files in Premiere Pro...\nBackup track: V${resolvedBackupVideoTrackNumber}`)
        : (isRebackup ? "Queueing checked re-backup jobs...\nExisting backup clips stay until export finishes." : `Queueing backup jobs...\nBackup track: V${resolvedBackupVideoTrackNumber}`));

    const selectedItemsJson = JSON.stringify(selectedQueueItems);
    const script = `exportBackup.runBackupQueue("${escapeForEvalScript(exportFolder)}","${escapeForEvalScript(videoPresetPath)}","${escapeForEvalScript(mp3PresetPath)}","${escapeForEvalScript(wavPresetPath)}","${escapeForEvalScript(selectedAudioFormat)}",${resolvedBackupVideoTrackNumber},${removeSequenceMarkers ? "true" : "false"},"${escapeForEvalScript(selectedItemsJson)}","${escapeForEvalScript(selectedExportMode)}",${isRebackup ? "true" : "false"},${autoEmptyTrack ? "true" : "false"})`;
    const result = await callHost(script);
    const parsed = parseHostResult(result);

    if (!parsed || parsed.ok === false) {
        let message = (parsed && parsed.message) || "Backup export failed.";
        const lineBreak = String.fromCharCode(10);

        if (
            parsed &&
            parsed.rebackupPrepared === true &&
            Array.isArray(parsed.queuedFiles) &&
            parsed.queuedFiles.length
        ) {
            try {
                const recoveryManifest = createExportManifestFromHostResult(parsed);
                recoveryManifest.exportFailed = true;
                recoveryManifest.manifestPath = writeExportManifest(recoveryManifest);
                updateAlignFolder(exportFolder);
                message +=
                    lineBreak + lineBreak +
                    "Existing backup clips remain linked to preserved old files." +
                    lineBreak +
                    "Retry Re-backup, or use Align Existing after all new exports exist.";
            } catch (manifestError) {
                message +=
                    lineBreak + lineBreak +
                    "Could not write the preserved-file recovery map: " +
                    manifestError.message;
            }
        }

        setStatus(message, "error");
        showBlockingMessage(message);
        setBusyState(false);
        return;
    }

    try {
        const manifest = createExportManifestFromHostResult(parsed);
        try { rememberBackupClipOwners(parsed); }
        catch (ownershipError) { setStatus('Could not save backup clip ownership: ' + ownershipError.message); }
        manifest.manifestPath = writeExportManifest(manifest);
        if (manifest.rebackup && parsed.exportMode === EXPORT_MODE_PREMIERE) {
            assertExpectedExportsAreReady(manifest);
            await prepareRebackupReplacement(manifest);
            await finalizeRebackupFiles(manifest);
            manifest.manifestPath = writeExportManifest(manifest);
        }
        updateAlignFolder(exportFolder);
        applyBackupDefaults({ videoTrackNumber: manifest.backupVideoTrackNumber }, false);
        if (parsed.exportMode === EXPORT_MODE_PREMIERE) {
            setBusyState(false);
            await runAlignmentFlow(exportFolder, {
                manifest,
                manifestOnly: true,
                skipVideo: false,
                sortProjectFiles: false,
                autoTriggered: true
            });
        } else {
            setBusyState(false);
            startExportCompletionMonitor(manifest);
        }
    } catch (error) {
        setBusyState(false);
        if (isLockedFileRecoveryMessage(error.message)) {
            showAlignmentRecoveryError(error.message);
            return;
        }

        const lineBreak = String.fromCharCode(10);
        const message = isRebackup
            ? (
                "Re-backup export finished, but old backup cleanup was not completed." +
                lineBreak +
                error.message +
                lineBreak +
                "Use Align Existing to retry cleanup and align the completed exports."
            )
            : "Queue created, but the export map could not be written." + lineBreak + error.message;
        if (isRebackup) {
            showAlignmentRecoveryError(message);
        } else {
            setStatus(message, "error");
        }
    }
}

async function alignExistingFolder() {
    if (busy) {
        return;
    }

    await stopPendingCleanupRetry(true);
    let manifest;
    try {
        const destination = await resolveExportActionDestination(true);
        exportFolder = destination.folderPath;
        const existing = destination.existingBackup;
        manifest = readManifestForSequence(exportFolder, existing.sequenceName);
        if (!manifest) manifest = createExistingBackupAlignmentManifest(existing, exportFolder);
        updateAlignFolder(exportFolder);
        document.getElementById("exportPath").textContent = exportFolder;
    } catch (error) {
        await showBlockingMessage(error.message);
        setStatus(error.message, "error");
        return;
    }
    await runAlignmentFlow(exportFolder || alignFolder, {
        manifest,
        manifestOnly: true,
        skipVideo: false,
        autoTriggered: false
    });
}

function createExistingBackupAlignmentManifest(existing, folderPath) {
    const layout = existing.layout || {};
    const expectedFiles = [];
    const video = layout.video || layout.backupAudio;
    if (video) expectedFiles.push({kind: "video", path: video.mediaPath || video.currentMediaPath});
    (layout.audioOutputs || []).forEach((entry) => expectedFiles.push({
        kind: "audio", path: entry.mediaPath || entry.currentMediaPath,
        trackNumber: entry.sourceTrackNumber,
        trackNumbers: entry.sourceTrackNumbers || [entry.sourceTrackNumber]
    }));
    return {folderPath, sequenceName: existing.sequenceName, baseName: layout.baseName || existing.baseName,
        rebackup: true, rebackupLayout: layout, expectedFiles, obsoleteAudioFiles: [],
        backupVideoTrackNumber: layout.video && layout.video.targetTrackNumber || getPositiveIntValue("exportVideoTrackInput", DEFAULT_BACKUP_VIDEO_TRACK)};
}

document.addEventListener("DOMContentLoaded", () => {
    readVersionInfo();
    loadSavedPresets();
    loadConfiguredBackupCategories();
    loadSavedPaths();
    loadSavedUiState();
    bindAudioFormatInputs();
    bindAlignOptions();
    bindExportOptions();
    bindAutoEmptyBackupTrackOption();
    bindBackupTrackStepper();
    bindInOutPrompt();
    bindCleanupRetryPrompt();
    bindBackupDestinationInputs();
    bindCategoryManager();
    resetAutoEmptyBackupTrackOption();
    refreshResolvedBackupDestination();
    markBackupInputsDirty();
    loadSavedBackupSettings();
    setQueueBackupSectionVisibility(true);
    setPresetSectionVisibility(presetSectionVisible);
    setUpdateButton(`Version ${localVersion}`, false, localVersionNotes);
    checkForUpdates();
    document.getElementById("videoPresetPath").textContent = videoPresetPath;
    updateAudioPresetDisplay();
    setStatus("Ready.");
    refreshSuggestedBackupTrack(false);
    refreshExportSelection();
});
