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

let exportFolder = null;
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
    setDisabled("chooseVideoPresetButton", nextBusy);
    setDisabled("chooseMp3PresetButton", nextBusy);
    setDisabled("chooseWavPresetButton", nextBusy);
    setDisabled("exportButton", nextBusy);
    setDisabled("rebackupButton", nextBusy);
    setDisabled("chooseAlignFolderButton", nextBusy);
    setDisabled("alignFolderButton", nextBusy);
    setDisabled("alignSkipVideoCheckbox", nextBusy);
    setDisabled("alignSortProjectFilesCheckbox", nextBusy);
    setDisabled("refreshExportSelectionButton", nextBusy);
    setDisabled("autoEmptyBackupTrackCheckbox", nextBusy);
    setDisabled("decrementBackupTrackButton", nextBusy);
    setDisabled("incrementBackupTrackButton", nextBusy);
    setDisabled("mergeAudioButton", nextBusy);
    setDisabled("updateButton", nextBusy);

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
    const button = document.getElementById("updateButton");
    button.textContent = label;
    button.disabled = busy || !isUpdateAvailable;
    button.title = hoverText || "";
    if (isUpdateAvailable) {
        button.classList.add("update-ready");
        button.classList.remove("secondary");
    } else {
        button.classList.remove("update-ready");
        button.classList.add("secondary");
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

function showBlockingMessage(message) {
    alert(message);
}

const ALIGNMENT_RECOVERY_TEXT = "Please use Align Existing to import the file.";

function showResultPrompt(title, message) {
    const lines = [String(title || "Done")];
    if (message) {
        lines.push(String(message));
    }
    showBlockingMessage(lines.join("\n\n"));
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
    showResultPrompt("IMPORT NOT COMPLETED", ALIGNMENT_RECOVERY_TEXT);
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
    return String(mode || "").toLowerCase() === EXPORT_MODE_PREMIERE.toLowerCase()
        ? EXPORT_MODE_PREMIERE
        : EXPORT_MODE_MEDIA_ENCODER;
}

function getSelectedExportMode() {
    const inputs = getExportModeInputs();
    return inputs.premiere && inputs.premiere.checked ? EXPORT_MODE_PREMIERE : EXPORT_MODE_MEDIA_ENCODER;
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

function bindAutoEmptyBackupTrackOption() {
    const checkbox = getAutoEmptyBackupTrackCheckbox();
    if (!checkbox) {
        return;
    }

    checkbox.addEventListener("change", () => {
        setAutoEmptyBackupTrackEnabled(checkbox.checked);
    });
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
    return path.join(process.env.APPDATA || "", "Adobe", "CEP", "extensions", "ExportBackup");
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

function runGithubUpdate() {
    if (busy) {
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

            const escapedScriptPath = tempUpdaterScriptPath.replace(/'/g, "''");
            const escapedZipPath = tempUpdaterZipPath.replace(/'/g, "''");
            const userDestination = getUserCepExtensionPath().replace(/'/g, "''");
            const escapedResultPath = tempUpdaterResultPath.replace(/'/g, "''");
            const escapedLogPath = tempUpdaterLogPath.replace(/'/g, "''");
            const command = `Start-Process PowerShell -Verb RunAs -ArgumentList '-NoExit','-NoProfile','-ExecutionPolicy','Bypass','-File','${escapedScriptPath}','-ZipPath','${escapedZipPath}','-Destination','${userDestination}','-ResultPath','${escapedResultPath}','-LogPath','${escapedLogPath}'`;

            childProcess.execFile(
                "powershell.exe",
                ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
                (error) => {
                    if (error) {
                        setStatus(`Could not launch updater.\n${error.message}`);
                        return;
                    }

                    setStatus(`Updater launched for the CEP extensions folder.\nTarget: ${getUserCepExtensionPath()}\nAn admin PowerShell window should show copy progress and stay open if something fails.`);
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
        const savedExportFolder = localStorage.getItem(EXPORT_FOLDER_STORAGE_KEY);
        if (savedExportFolder && savedExportFolder.trim()) {
            exportFolder = savedExportFolder;
            alignFolder = savedExportFolder;
            document.getElementById("exportPath").textContent = exportFolder;
        }
    } catch (error) {}

    try {
        const savedAlignFolder = localStorage.getItem(ALIGN_FOLDER_STORAGE_KEY);
        if (!exportFolder && savedAlignFolder && savedAlignFolder.trim()) {
            alignFolder = savedAlignFolder;
            const alignPathElement = document.getElementById("alignPath");
            if (alignPathElement) {
                alignPathElement.textContent = alignFolder;
            }
        }
    } catch (error) {}
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

    try {
        const savedExportMode = localStorage.getItem(EXPORT_MODE_STORAGE_KEY);
        setSelectedExportMode(savedExportMode || EXPORT_MODE_MEDIA_ENCODER);
    } catch (error) {
        setSelectedExportMode(EXPORT_MODE_MEDIA_ENCODER);
    }

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

    const result = await callHost("exportBackup.getExportSelectionInfo()");
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
    mergedAudioGroups = [];
    nextMergedAudioGroupId = 1;
    const items = Array.isArray(selectionInfo.items) ? selectionInfo.items : [];

    if (!items.length) {
        container.innerHTML = `<div class="small-note">${selectionInfo.message || "No used audio tracks were found in the active sequence yet. The backup MP4 will still be queued."}</div>`;
        return;
    }

    container.innerHTML = items.map((item, index) => {
        const checkboxId = `exportSelectionItem_${index}`;
        const mergeCheckboxId = `mergeSelectionItem_${index}`;
        const checked = item.kind === "video" || item.selected !== false ? "checked" : "";
        const disabled = item.locked ? "disabled" : "";
        const kindLabel = item.kind === "video" ? "Backup video" : `Audio track ${item.trackNumber}`;
        const detail = item.kind === "audio" && item.trackName ? item.trackName : "";
        let mergeControl = `<span></span>`;
        if (item.kind === "video") {
            mergeControl = `<button id="mergeAudioButton" class="secondary selection-list-merge-button" type="button" onclick="mergeSelectedAudioTracks()">Merge Selection</button>`;
        } else if (item.kind === "audio") {
            mergeControl = `<label class="merge-checkbox-wrap" for="${mergeCheckboxId}"><span>Merge</span><input class="merge-checkbox" type="checkbox" id="${mergeCheckboxId}" data-kind="audio" data-track-number="${item.trackNumber || 0}" data-merged-group=""></label>`;
        }

        return (
            `<div class="selection-item">` +
                `<input class="queue-checkbox" type="checkbox" id="${checkboxId}" data-kind="${item.kind}" data-track-number="${item.trackNumber || 0}" ${checked} ${disabled}>` +
                `<label for="${checkboxId}">` +
                    `<strong>${escapeHtml(kindLabel)}</strong>` +
                    (detail ? `<small>${escapeHtml(detail)}</small>` : "") +
                `</label>` +
                mergeControl +
            `</div>`
        );
    }).join("");
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

    return {
        includeVideo: videoInput ? !!videoInput.checked : true,
        audioTracks: getSelectedAudioTrackNumbers(),
        audioGroups: mergedAudioGroups.map((group) => group.trackNumbers)
    };
}

async function refreshExportSelection() {
    if (busy) {
        return;
    }

    renderExportSelectionList({ ok: true, items: [], message: "Reading active sequence tracks..." });
    const selectionInfo = await getExportSelectionInfo();
    renderExportSelectionList(selectionInfo);
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


function deleteLocalFileWithExplorer(filePath) {
    if (!filePath || !fileExists(filePath)) {
        return { ok: true, error: "" };
    }

    try {
        const env = Object.assign({}, process.env, { EXPORTBACKUP_DELETE_TARGET: filePath });
        const script = [
            "$p=$env:EXPORTBACKUP_DELETE_TARGET",
            "if (-not $p -or -not (Test-Path -LiteralPath $p)) { exit 0 }",
            "$folder=Split-Path -LiteralPath $p -Parent",
            "$name=Split-Path -LiteralPath $p -Leaf",
            "$shell=New-Object -ComObject Shell.Application",
            "$namespace=$shell.Namespace($folder)",
            "if (-not $namespace) { exit 2 }",
            "$item=$namespace.ParseName($name)",
            "if (-not $item) { exit 3 }",
            "$item.InvokeVerb('delete')",
            "Start-Sleep -Milliseconds 2000",
            "if (Test-Path -LiteralPath $p) { exit 1 }"
        ].join("; ");
        const result = childProcess.spawnSync(
            "powershell.exe",
            ["-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-Command", script],
            { env, windowsHide: false, timeout: 30000, encoding: "utf8" }
        );

        return {
            ok: !fileExists(filePath),
            error: result.error ? result.error.message : (result.stderr || "")
        };
    } catch (error) {
        return { ok: !fileExists(filePath), error: error.message };
    }
}

async function tryExplorerDelete(filePath, label) {
    if (!filePath || !fileExists(filePath)) {
        return true;
    }

    setStatus(`Trying Windows Explorer delete for ${label || "old local file"}.\n${filePath}`);
    return deleteLocalFileWithExplorer(filePath).ok;
}

function deleteLocalFileWithWindowsShell(filePath) {
    if (!filePath || !fileExists(filePath)) {
        return { ok: true, error: "" };
    }

    try {
        const env = Object.assign({}, process.env, { EXPORTBACKUP_DELETE_TARGET: filePath });
        const script = "$p=$env:EXPORTBACKUP_DELETE_TARGET; if ($p -and (Test-Path -LiteralPath $p)) { Remove-Item -LiteralPath $p -Force -ErrorAction Stop }";
        const result = childProcess.spawnSync(
            "powershell.exe",
            ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
            { env, windowsHide: true, timeout: 2500, encoding: "utf8" }
        );

        return {
            ok: !fileExists(filePath),
            error: result.error ? result.error.message : (result.stderr || "")
        };
    } catch (error) {
        return { ok: !fileExists(filePath), error: error.message };
    }
}

function deleteLocalFileWithCmd(filePath) {
    if (!filePath || !fileExists(filePath)) {
        return { ok: true, error: "" };
    }

    try {
        const env = Object.assign({}, process.env, { EXPORTBACKUP_DELETE_TARGET: filePath });
        const result = childProcess.spawnSync(
            "cmd.exe",
            ["/d", "/s", "/c", "del /f /q \"%EXPORTBACKUP_DELETE_TARGET%\""],
            { env, windowsHide: true, timeout: 2500, encoding: "utf8" }
        );

        return {
            ok: !fileExists(filePath),
            error: result.error ? result.error.message : (result.stderr || "")
        };
    } catch (error) {
        return { ok: !fileExists(filePath), error: error.message };
    }
}


function deleteLocalFileNow(filePath) {
    if (!filePath || !fileExists(filePath)) {
        return true;
    }

    try {
        fs.chmodSync(filePath, 0o666);
    } catch (error) {}

    try {
        if (fs.rmSync) {
            fs.rmSync(filePath, { force: true, maxRetries: 1, retryDelay: 100 });
        } else {
            fs.unlinkSync(filePath);
        }
    } catch (error) {}

    if (!fileExists(filePath)) {
        return true;
    }

    if (deleteLocalFileWithWindowsShell(filePath).ok) {
        return true;
    }

    if (deleteLocalFileWithCmd(filePath).ok) {
        return true;
    }

    return !fileExists(filePath);
}

function buildDeletePendingPath(filePath) {
    const parsed = path.parse(filePath);
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
    if (!filePath || !fileExists(filePath)) {
        return { ok: true, pendingPath: "", error: "" };
    }

    const pendingPath = buildDeletePendingPath(filePath);
    try {
        fs.renameSync(filePath, pendingPath);
        return {
            ok: !fileExists(filePath) && fileExists(pendingPath),
            pendingPath,
            error: ""
        };
    } catch (error) {
        return {
            ok: !fileExists(filePath),
            pendingPath: !fileExists(filePath) ? pendingPath : "",
            error: error.message
        };
    }
}

async function moveFileAsideIfNeeded(filePath, label) {
    if (!filePath || !fileExists(filePath)) {
        return true;
    }

    setStatus(`Delete did not work. Trying to move ${label || "old local file"} aside.\n${filePath}`);
    const result = moveLocalFileAside(filePath);
    if (result.ok) {
        setStatus(
            `Old file moved aside.\n` +
            `Original path is free now.\n` +
            `${result.pendingPath || filePath}`
        );
        return true;
    }

    return false;
}

async function removeFileIfExists(filePath, description) {
    if (!filePath || !fileExists(filePath)) {
        return;
    }

    const label = description || "old local file";
    const retryDelays = [10000];
    for (let attempt = 1; attempt <= 2; attempt += 1) {
        setStatus(`Deleting ${label}.\n${filePath}\nAttempt ${attempt}/2.`);
        if (deleteLocalFileNow(filePath)) {
            setStatus(
                `Old file deleted.\n` +
                `Original path is free now.\n` +
                `${filePath}`
            );
            return;
        }

        if (await tryExplorerDelete(filePath, label)) {
            setStatus(
                `Old file deleted with Windows Explorer.\n` +
                `Original path is free now.\n` +
                `${filePath}`
            );
            return;
        }

        if (attempt < 2) {
            const waitSeconds = Math.round(retryDelays[attempt - 1] / 1000);
            setStatus(
                `File is still busy. Retrying delete in ${waitSeconds} seconds.\n` +
                `${filePath}\n` +
                `Retry ${attempt}/1.`
            );
            await delay(retryDelays[attempt - 1]);
        }
    }

    if (await moveFileAsideIfNeeded(filePath, label)) {
        return;
    }

    throw new Error("File still busy\n\nPress Align Existing button to delete and import re-export.");
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
    const preservedPaths = getManifestPreservedPaths(manifest);
    const lineBreak = String.fromCharCode(10);

    if (!manifest || manifest.rebackupReplacementPrepared !== true) {
        throw new Error("Re-backup cleanup was not verified, so local backup files were not finalized.");
    }

    if (replacementFiles.length || preservedPaths.length) {
        await waitBeforeLocalFileStep(
            "Premiere cleanup is done. Local backup cleanup will start next.",
            2500
        );
    }

    if (replacementFiles.length) {
        setStatus(
            "Replacing local backup files." + lineBreak +
            "Files to rename: " + replacementFiles.length + "."
        );

        for (let i = 0; i < replacementFiles.length; i += 1) {
            const entry = replacementFiles[i];
            await replaceRebackupFile(entry, i + 1, replacementFiles.length);
            entry.path = entry.finalPath;
            entry.name = path.basename(entry.finalPath);
        }
    }

    if (preservedPaths.length) {
        setStatus(
            "New exports are complete." + lineBreak +
            "Deleting preserved old backup files: " + preservedPaths.length + "."
        );

        for (let i = 0; i < preservedPaths.length; i += 1) {
            await removeFileIfExists(preservedPaths[i], "preserved old backup file");
        }

        expectedFiles.forEach((entry) => {
            if (entry) {
                entry.preservedPaths = [];
            }
        });
        manifest.rebackupPreservedCleanupDone = true;
    }

    manifest.rebackupFinalized = true;
    return manifest;
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

async function removeStaleAudioFormatFiles(stalePaths) {
    for (const stalePath of (stalePaths || [])) {
        if (stalePath) {
            await removeFileIfExists(stalePath, "older audio format file");
        }
    }
}

async function prepareAlignExistingCleanup(matchInfo) {
    const stalePaths = matchInfo && Array.isArray(matchInfo.staleAudioPaths) ? matchInfo.staleAudioPaths : [];
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

    await removeStaleAudioFormatFiles(stalePaths);
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

    setStatus(
        "Premiere cleanup complete.\n" +
        `Timeline clips removed: ${parseInt(parsed.removedTimelineClips, 10) || 0}.\n` +
        `Project items removed/offlined: ${(parseInt(parsed.removedProjectItems, 10) || 0) + (parseInt(parsed.offlinedProjectItems, 10) || 0)}.\n` +
        "Step 2: replacing local files."
    );

    manifest.rebackupReplacementPrepared = true;
    manifest.rebackupCleanup = {
        removedTimelineClips: parseInt(parsed.removedTimelineClips, 10) || 0,
        removedProjectItems: parseInt(parsed.removedProjectItems, 10) || 0,
        offlinedProjectItems: parseInt(parsed.offlinedProjectItems, 10) || 0
    };
    return parsed;
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
        expectedFiles: Array.isArray(parsed.queuedFiles) ? parsed.queuedFiles : [],
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

async function runAlignmentFlow(folderPath, options) {
    const settings = options || {};

    if (!folderPath) {
        alert("Choose an export folder first.");
        return false;
    }

    setBusyState(true);
    setStatus(settings.autoTriggered
        ? "Export finished. Importing and aligning backup files..."
        : "Loading Premiere host script...");

    try {
        if (!(await ensureHostLoaded())) {
            showAlignmentRecoveryError("Could not load Premiere host script.");
            return false;
        }

        const activeSequenceName = await getActiveSequenceName();
        if (!activeSequenceName) {
            const message = "No active sequence is open in Premiere Pro.";
            showAlignmentRecoveryError(message);
            return false;
        }

        let manifest = settings.manifest || readManifestForSequence(folderPath, activeSequenceName);
        const manifestOnly = settings.manifestOnly === true && !!manifest;

        if (
            manifest &&
            manifest.rebackup === true &&
            manifest.rebackupPrepared === true &&
            manifest.rebackupReplacementPrepared !== true &&
            getManifestPreservedPaths(manifest).length
        ) {
            assertExpectedExportsAreReady(manifest);
            setStatus("New Re-backup files found. Removing preserved old media before alignment...");
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
            showAlignmentRecoveryError(message);
            return false;
        }

        await prepareAlignExistingCleanup(matchInfo);

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
        const script = `exportBackup.alignMappedFiles("${escapeForEvalScript(resolvedVideoPath)}","${escapeForEvalScript(audioJson)}",${backupVideoTrackNumber},${sortProjectFiles},"${escapeForEvalScript(rebackupLayoutJson)}")`;
        const result = await callHost(script);
        const parsed = parseHostResult(result);

        if (!parsed || parsed.ok === false) {
            const message = (parsed && parsed.message) || "Alignment failed.";
            showAlignmentRecoveryError(message);
            return false;
        }

        const shouldCopyProject = !!(getCopyProjectFileCheckbox() && getCopyProjectFileCheckbox().checked);
        const copyResult = shouldCopyProject
            ? copyProjectToFolder(parsed.projectPath, folderPath)
            : { ok: false, message: "" };
        if (matchInfo.manifest && matchInfo.manifest.manifestPath) {
            await removeFileIfExists(matchInfo.manifest.manifestPath, "export map file");
        }

        const successTitle = getCompletionStatusTitle(settings, matchInfo.manifest);
        const lines = [successTitle, parsed.message || "Alignment completed."];
        if (parsed.importBinName) {
            lines.push(`Imported backup files were added to project bin: ${parsed.importBinName}`);
        }
        if (copyResult.ok) {
            lines.push(`Project copy saved: ${copyResult.destinationPath}`);
        } else if (copyResult.message) {
            lines.push(copyResult.message);
        }

        setStatus(lines.join("\n"), "success");
        showResultPrompt(
            successTitle,
            settings.autoTriggered
                ? "Backup files were imported and aligned successfully."
                : "Existing backup files were imported and aligned successfully."
        );
        return true;
    } catch (error) {
        const message = `Alignment failed.\n${error.message}`;
        showAlignmentRecoveryError(message);
        return false;
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
        exportFolder = result.data[0];
        alignFolder = exportFolder;
        try {
            localStorage.setItem(EXPORT_FOLDER_STORAGE_KEY, exportFolder);
            localStorage.setItem(ALIGN_FOLDER_STORAGE_KEY, alignFolder);
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

async function runExport(isRebackup) {
    if (busy) {
        return;
    }

    if (!exportFolder) {
        alert("Choose an export folder first.");
        return;
    }

    const selectedAudioFormat = getSelectedAudioFormat();
    const selectedAudioPresetPath = selectedAudioFormat === "wav" ? wavPresetPath : mp3PresetPath;
    const backupVideoTrackNumber = getPositiveIntValue("exportVideoTrackInput", DEFAULT_BACKUP_VIDEO_TRACK);
    const autoEmptyTrack = !isRebackup && useAutoEmptyBackupTrack();
    const selectedQueueItems = getSelectedQueueItems();
    const removeSequenceMarkers = !!(getRemoveSequenceMarkersCheckbox() && getRemoveSequenceMarkersCheckbox().checked);
    const selectedExportMode = getSelectedExportMode();

    if (!fileExists(videoPresetPath)) {
        alert("The selected video preset file was not found. Choose the video preset again.");
        return;
    }

    if (!fileExists(selectedAudioPresetPath)) {
        alert(`The selected ${selectedAudioFormat.toUpperCase()} preset file was not found.`);
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

    await runAlignmentFlow(exportFolder || alignFolder, {
        skipVideo: false,
        autoTriggered: false
    });
}

document.addEventListener("DOMContentLoaded", () => {
    readVersionInfo();
    loadSavedPresets();
    loadSavedPaths();
    loadSavedUiState();
    bindAudioFormatInputs();
    bindAlignOptions();
    bindExportOptions();
    bindAutoEmptyBackupTrackOption();
    bindBackupTrackStepper();
    bindInOutPrompt();
    resetAutoEmptyBackupTrackOption();
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
