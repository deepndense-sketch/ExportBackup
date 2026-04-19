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

const DEFAULT_BACKUP_VIDEO_TRACK = 5;
const EXPORT_MANIFEST_SUFFIX = "_ExportBackupMap.json";
const EXPORT_MONITOR_INTERVAL_MS = 5000;
const EXPORT_MONITOR_STABLE_PASSES = 2;
const EXPORT_MONITOR_TIMEOUT_MS = 6 * 60 * 60 * 1000;

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
let exportMonitorState = null;
let exportSelectionState = null;

function getExtensionRootPath() {
    try {
        return csInterface.getSystemPath(SystemPath.EXTENSION);
    } catch (error) {
        return __dirname;
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

function getDefaultVideoPresetPath() {
    return getBundledPresetPath("1080 AIR.epr");
}

function getDefaultMp3PresetPath() {
    return getBundledPresetPath("mp3.epr");
}

function getDefaultWavPresetPath() {
    return getBundledPresetPath("wav.epr");
}

function setStatus(message) {
    document.getElementById("statusBox").textContent = message;
}

function setPresetSectionVisibility(visible) {
    presetSectionVisible = visible;

    const presetSection = document.getElementById("presetSection");
    const toggleButton = document.getElementById("togglePresetSectionButton");
    if (!presetSection || !toggleButton) {
        return;
    }

    presetSection.classList.toggle("is-hidden", !visible);
    toggleButton.textContent = visible ? "Hide Presets" : "Show Presets";

    try {
        localStorage.setItem(PRESET_SECTION_VISIBLE_STORAGE_KEY, visible ? "true" : "false");
    } catch (error) {}
}

function togglePresetSection() {
    setPresetSectionVisibility(!presetSectionVisible);
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

function getAlignVideoTrackInput() {
    return document.getElementById("alignVideoTrackInput");
}

function getRemoveSequenceMarkersCheckbox() {
    return document.getElementById("removeSequenceMarkersCheckbox");
}

function setBusyState(nextBusy) {
    const audioInputs = getAudioFormatInputs();

    busy = nextBusy;
    document.getElementById("chooseFolderButton").disabled = nextBusy;
    document.getElementById("chooseVideoPresetButton").disabled = nextBusy;
    document.getElementById("chooseMp3PresetButton").disabled = nextBusy;
    document.getElementById("chooseWavPresetButton").disabled = nextBusy;
    document.getElementById("exportButton").disabled = nextBusy;
    document.getElementById("chooseAlignFolderButton").disabled = nextBusy;
    document.getElementById("alignFolderButton").disabled = nextBusy;
    document.getElementById("alignSkipVideoCheckbox").disabled = nextBusy;
    document.getElementById("alignSortProjectFilesCheckbox").disabled = nextBusy;
    document.getElementById("refreshExportSelectionButton").disabled = nextBusy;
    document.getElementById("updateButton").disabled = nextBusy;

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

function showBlockingMessage(message) {
    alert(message);
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
    const alignTrackInput = getAlignVideoTrackInput();
    if (!alignTrackInput) {
        return;
    }

    const value = Math.max(
        1,
        parseInt((defaults && defaults.videoTrackNumber) || DEFAULT_BACKUP_VIDEO_TRACK, 10) || DEFAULT_BACKUP_VIDEO_TRACK
    );

    if (force || alignTrackInput.dataset.userEdited !== "true") {
        alignTrackInput.value = String(value);
        alignTrackInput.dataset.autoValue = String(value);
        saveAlignVideoTrack(value);
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
            saveAlignVideoTrack(getPositiveIntValue("alignVideoTrackInput", DEFAULT_BACKUP_VIDEO_TRACK));
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
    if (!removeMarkersCheckbox) {
        return;
    }

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
            setUpdateButton(`Version ${localVersion}`, false, localVersionNotes);
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
            document.getElementById("exportPath").textContent = exportFolder;
        }
    } catch (error) {}

    try {
        const savedAlignFolder = localStorage.getItem(ALIGN_FOLDER_STORAGE_KEY);
        if (savedAlignFolder && savedAlignFolder.trim()) {
            alignFolder = savedAlignFolder;
            document.getElementById("alignPath").textContent = alignFolder;
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

async function getActiveSequenceName() {
    if (!(await ensureHostLoaded())) {
        return "";
    }

    const result = await callHost("exportBackup.getActiveSequenceName()");
    return String(result || "").trim();
}

async function validateBackupExportSettings(backupVideoTrackNumber) {
    if (!(await ensureHostLoaded())) {
        return { ok: false, message: "Could not load Premiere host script." };
    }

    const result = await callHost(`exportBackup.validateBackupExportSettings(${backupVideoTrackNumber})`);
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
    const items = Array.isArray(selectionInfo.items) ? selectionInfo.items : [];

    if (!items.length) {
        container.innerHTML = `<div class="small-note">${selectionInfo.message || "No used audio tracks were found in the active sequence yet. The backup MP4 will still be queued."}</div>`;
        return;
    }

    container.innerHTML = items.map((item, index) => {
        const checkboxId = `exportSelectionItem_${index}`;
        const checked = item.selected === false ? "" : "checked";
        const disabled = item.locked ? "disabled" : "";
        const kindLabel = item.kind === "video" ? "Backup video" : `Audio track ${item.trackNumber}`;
        const detail = item.kind === "video"
            ? "Untick this if you do not want to export the backup MP4."
            : "Untick this track if you do not want to export it.";

        return (
            `<label class="selection-item" for="${checkboxId}">` +
                `<input type="checkbox" id="${checkboxId}" data-kind="${item.kind}" data-track-number="${item.trackNumber || 0}" ${checked} ${disabled}>` +
                `<span>` +
                    `<strong>${kindLabel}</strong>` +
                    `<small>${detail}</small>` +
                `</span>` +
            `</label>`
        );
    }).join("");
}

function getSelectedAudioTrackNumbers() {
    const container = document.getElementById("exportSelectionList");
    if (!container) {
        return [];
    }

    return Array.prototype.slice.call(container.querySelectorAll("input[type='checkbox'][data-kind='audio']:checked"))
        .map((input) => parseInt(input.getAttribute("data-track-number"), 10) || 0)
        .filter((trackNumber) => trackNumber > 0);
}

function getSelectedQueueItems() {
    const container = document.getElementById("exportSelectionList");
    const videoInput = container ? container.querySelector("input[type='checkbox'][data-kind='video']") : null;

    return {
        includeVideo: videoInput ? !!videoInput.checked : true,
        audioTracks: getSelectedAudioTrackNumbers()
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

function parseTrackNumberFromFileName(name, baseName) {
    const lowerName = String(name || "").toLowerCase();
    const lowerBase = String(baseName || "").toLowerCase();
    const prefix = `${lowerBase}_track`;

    if (!lowerName.startsWith(prefix)) {
        return 0;
    }

    const remainder = name.substring(prefix.length);
    const dotIndex = remainder.lastIndexOf(".");
    if (dotIndex <= 0) {
        return 0;
    }

    return parseInt(remainder.substring(0, dotIndex), 10) || 0;
}

function normalizeAudioEntries(entries, baseName) {
    const normalized = [];

    (entries || []).forEach((entry) => {
        if (!entry || !entry.path || !fileExists(entry.path)) {
            return;
        }

        const fileName = entry.name || path.basename(entry.path);
        const trackNumber = parseInt(entry.trackNumber, 10) || parseTrackNumberFromFileName(fileName, baseName);
        if (trackNumber < 1) {
            return;
        }

        normalized.push({
            path: entry.path,
            trackNumber,
            name: fileName
        });
    });

    normalized.sort((a, b) => a.trackNumber - b.trackNumber);
    return normalized;
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

function scanExportFolderForSequence(folderPath, sequenceName, manifest) {
    const sanitizedBase = sanitizeSequenceName(sequenceName);
    const entries = fs.readdirSync(folderPath, { withFileTypes: true });
    const files = entries.filter((entry) => entry.isFile()).map((entry) => entry.name);
    const lowerBase = sanitizedBase.toLowerCase();
    const backupPrefix = `${lowerBase}_backup.`;

    let videoPath = "";
    let audio = [];

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
                    name: entry.name || path.basename(entry.path)
                });
            }
        });
    }

    files.forEach((fileName) => {
        const absolutePath = path.join(folderPath, fileName);
        const lowerName = fileName.toLowerCase();

        if (!videoPath && lowerName.startsWith(backupPrefix)) {
            videoPath = absolutePath;
            return;
        }

        const trackNumber = parseTrackNumberFromFileName(fileName, sanitizedBase);
        if (trackNumber > 0 && !audio.some((entry) => entry.path === absolutePath)) {
            audio.push({
                path: absolutePath,
                trackNumber,
                name: fileName
            });
        }
    });

    audio = normalizeAudioEntries(audio, sanitizedBase);

    return {
        baseName: sanitizedBase,
        videoPath,
        audio,
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

function removeFileIfExists(filePath) {
    try {
        if (fileExists(filePath)) {
            fs.unlinkSync(filePath);
        }
    } catch (error) {}
}

function updateAlignFolder(folderPath) {
    alignFolder = folderPath;
    document.getElementById("alignPath").textContent = folderPath || "No folder selected yet.";

    try {
        if (folderPath) {
            localStorage.setItem(ALIGN_FOLDER_STORAGE_KEY, folderPath);
        }
    } catch (error) {}
}

function createExportManifestFromHostResult(parsed) {
    return {
        version: 2,
        createdAt: new Date().toISOString(),
        folderPath: exportFolder,
        sequenceName: parsed.sequenceName || "",
        baseName: parsed.baseName || sanitizeSequenceName(parsed.sequenceName || "Active_Sequence"),
        backupVideoTrackNumber: parseInt(parsed.backupVideoTrackNumber, 10) || DEFAULT_BACKUP_VIDEO_TRACK,
        audioFormat: parsed.audioFormat || getSelectedAudioFormat(),
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
        alert("Choose an align folder first.");
        return false;
    }

    setBusyState(true);
    setStatus(settings.autoTriggered
        ? "Media Encoder finished. Importing and aligning backup files..."
        : "Loading Premiere host script...");

    try {
        if (!(await ensureHostLoaded())) {
            showBlockingMessage("Could not load Premiere host script.");
            return false;
        }

        const activeSequenceName = await getActiveSequenceName();
        if (!activeSequenceName) {
            const message = "No active sequence is open in Premiere Pro.";
            setStatus(message);
            showBlockingMessage(message);
            return false;
        }

        const manifest = settings.manifest || readManifestForSequence(folderPath, activeSequenceName);
        const matchInfo = scanExportFolderForSequence(folderPath, activeSequenceName, manifest);

        if (matchInfo.manifest && matchInfo.manifest.backupVideoTrackNumber) {
            applyAlignDefaults({ videoTrackNumber: matchInfo.manifest.backupVideoTrackNumber }, true);
        }

        if (!matchInfo.videoPath && matchInfo.audio.length === 0) {
            const message =
                "No files could be matched in the chosen folder.\n" +
                `Sequence base: ${matchInfo.baseName}\n` +
                `Folder files: ${matchInfo.folderFiles.join(" | ")}`;
            setStatus(message);
            showBlockingMessage(message);
            return false;
        }

        const backupVideoTrackNumber = getPositiveIntValue(
            "alignVideoTrackInput",
            (matchInfo.manifest && parseInt(matchInfo.manifest.backupVideoTrackNumber, 10)) || DEFAULT_BACKUP_VIDEO_TRACK
        );
        saveAlignVideoTrack(backupVideoTrackNumber);
        const skipBackupVideo = settings.skipVideo === true || document.getElementById("alignSkipVideoCheckbox").checked;
        const sortProjectFiles = settings.sortProjectFiles === true || document.getElementById("alignSortProjectFilesCheckbox").checked;
        const resolvedVideoPath = skipBackupVideo ? "" : (matchInfo.videoPath || "");
        const audioJson = JSON.stringify(matchInfo.audio);
        const script = `exportBackup.alignMappedFiles("${escapeForEvalScript(resolvedVideoPath)}","${escapeForEvalScript(audioJson)}",${backupVideoTrackNumber},${sortProjectFiles})`;
        const result = await callHost(script);
        const parsed = parseHostResult(result);

        if (!parsed || parsed.ok === false) {
            const message = (parsed && parsed.message) || "Alignment failed.";
            setStatus(message);
            showBlockingMessage(message);
            return false;
        }

        const copyResult = copyProjectToFolder(parsed.projectPath, folderPath);
        if (matchInfo.manifest && matchInfo.manifest.manifestPath) {
            removeFileIfExists(matchInfo.manifest.manifestPath);
        }

        const lines = [parsed.message || "Alignment completed."];
        if (parsed.importBinName) {
            lines.push(`Imported backup files were added to project bin: ${parsed.importBinName}`);
        }
        if (copyResult.ok) {
            lines.push(`Project copy saved: ${copyResult.destinationPath}`);
        } else if (copyResult.message) {
            lines.push(copyResult.message);
        }

        setStatus(lines.join("\n"));
        showBlockingMessage(settings.autoTriggered ? "Automatic import and alignment finished." : "Done.");
        return true;
    } catch (error) {
        const message = `Alignment failed.\n${error.message}`;
        setStatus(message);
        showBlockingMessage(message);
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
            setStatus(`Automatic import stopped.\n${error.message}`);
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
        setStatus(
            "Automatic import timed out while waiting for Media Encoder.\n" +
            "The queued exports are still in the chosen folder. Use Align Existing Export Folder after the renders finish."
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
        await runAlignmentFlow(state.manifest.folderPath, {
            manifest: state.manifest,
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
        try {
            localStorage.setItem(EXPORT_FOLDER_STORAGE_KEY, exportFolder);
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

    const result = window.cep.fs.showOpenDialogEx(false, false, "Choose Premiere Video Preset (.epr)", null, ["epr"]);
    if (result.data && result.data.length > 0) {
        saveVideoPreset(result.data[0]);
        setStatus("Video preset updated. This choice will be remembered until you change it.");
    }
}

async function chooseMp3Preset() {
    if (busy) {
        return;
    }

    const result = window.cep.fs.showOpenDialogEx(false, false, "Choose Premiere MP3 Preset (.epr)", null, ["epr"]);
    if (result.data && result.data.length > 0) {
        saveMp3Preset(result.data[0]);
        setStatus("MP3 preset updated. This choice will be remembered until you change it.");
    }
}

async function chooseWavPreset() {
    if (busy) {
        return;
    }

    const result = window.cep.fs.showOpenDialogEx(false, false, "Choose Premiere WAV Preset (.epr)", null, ["epr"]);
    if (result.data && result.data.length > 0) {
        saveWavPreset(result.data[0]);
        setStatus("WAV preset updated. This choice will be remembered until you change it.");
    }
}

async function runExport() {
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
    const selectedQueueItems = getSelectedQueueItems();
    const removeSequenceMarkers = !!(getRemoveSequenceMarkersCheckbox() && getRemoveSequenceMarkersCheckbox().checked);

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

    setBusyState(true);
    setStatus("Loading Premiere host script...");

    if (!(await ensureHostLoaded())) {
        showBlockingMessage("Could not load Premiere host script.");
        setBusyState(false);
        return;
    }

    const validation = await validateBackupExportSettings(backupVideoTrackNumber);
    if (!validation.ok) {
        showBlockingMessage(validation.message || "Backup export validation failed.");
        setStatus(validation.message || "Backup export validation failed.");
        setBusyState(false);
        return;
    }

    setStatus("Queueing Media Encoder jobs and writing export map...");

    const selectedItemsJson = JSON.stringify(selectedQueueItems);
    const script = `exportBackup.runBackupQueue("${escapeForEvalScript(exportFolder)}","${escapeForEvalScript(videoPresetPath)}","${escapeForEvalScript(mp3PresetPath)}","${escapeForEvalScript(wavPresetPath)}","${escapeForEvalScript(selectedAudioFormat)}",${backupVideoTrackNumber},${removeSequenceMarkers ? "true" : "false"},"${escapeForEvalScript(selectedItemsJson)}")`;
    const result = await callHost(script);
    const parsed = parseHostResult(result);

    if (!parsed || parsed.ok === false) {
        const message = (parsed && parsed.message) || "Backup export failed.";
        setStatus(message);
        showBlockingMessage(message);
        setBusyState(false);
        return;
    }

    try {
        const manifest = createExportManifestFromHostResult(parsed);
        manifest.manifestPath = writeExportManifest(manifest);
        updateAlignFolder(exportFolder);
        applyBackupDefaults({ videoTrackNumber: manifest.backupVideoTrackNumber }, false);
        setBusyState(false);
        startExportCompletionMonitor(manifest);
    } catch (error) {
        setBusyState(false);
        const message = `Queue created, but the export map could not be written.\n${error.message}`;
        setStatus(message);
        showBlockingMessage(message);
    }
}

async function alignExistingFolder() {
    if (busy) {
        return;
    }

    await runAlignmentFlow(alignFolder, {
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
    markBackupInputsDirty();
    loadSavedBackupSettings();
    setPresetSectionVisibility(presetSectionVisible);
    setUpdateButton(`Version ${localVersion}`, false, localVersionNotes);
    checkForUpdates();
    document.getElementById("videoPresetPath").textContent = videoPresetPath;
    updateAudioPresetDisplay();
    setStatus("Ready.");
    refreshSuggestedBackupTrack(false);
    refreshExportSelection();
});
