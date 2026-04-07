const csInterface = new CSInterface();
const fs = require("fs");
const path = require("path");

const VERSION_FILE_PATH = path.join(__dirname, "..", "version.json");
const DEFAULT_VIDEO_PRESET_PATH = "D:\\Work\\Tools\\ExportBackup\\presets\\1080 AIR.epr";
const MP3_PRESET_PATH = "D:\\Work\\Tools\\ExportBackup\\presets\\mp3.epr";
const WAV_PRESET_PATH = "C:\\Program Files\\Adobe\\Adobe Media Encoder 2026\\MediaIO\\systempresets\\3F3F3F3F_57415645\\Waveform Audio 48kHz 16-bit.epr";
const VIDEO_PRESET_STORAGE_KEY = "exportbackup.videoPresetPath";
const EXPORT_FOLDER_STORAGE_KEY = "exportbackup.exportFolder";
const ALIGN_FOLDER_STORAGE_KEY = "exportbackup.alignFolder";

let exportFolder = null;
let alignFolder = null;
let hostLoaded = false;
let busy = false;
let videoPresetPath = DEFAULT_VIDEO_PRESET_PATH;
let localVersion = "unknown";

function setStatus(message) {
    document.getElementById("statusBox").textContent = message;
}

function setUpdateMessage(message) {
    document.getElementById("updateValue").textContent = message;
}

function setBusyState(nextBusy) {
    busy = nextBusy;
    document.getElementById("chooseFolderButton").disabled = nextBusy;
    document.getElementById("chooseVideoPresetButton").disabled = nextBusy;
    document.getElementById("exportButton").disabled = nextBusy;
    document.getElementById("chooseAlignFolderButton").disabled = nextBusy;
    document.getElementById("alignFolderButton").disabled = nextBusy;
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

function fileExists(filePath) {
    try {
        return !!filePath && fs.existsSync(filePath);
    } catch (error) {
        return false;
    }
}

function getPositiveIntValue(elementId, fallbackValue) {
    const raw = document.getElementById(elementId).value;
    const parsed = parseInt(raw, 10);
    if (!parsed || parsed < 1) {
        return fallbackValue;
    }
    return parsed;
}

function readVersionInfo() {
    try {
        const raw = fs.readFileSync(VERSION_FILE_PATH, "utf8");
        const parsed = JSON.parse(raw);
        localVersion = parsed.version || "unknown";
    } catch (error) {
        localVersion = "unknown";
    }

    document.getElementById("versionValue").textContent = localVersion;
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
    setUpdateMessage("Checking...");

    try {
        const response = await fetch(remoteUrl, { cache: "no-store" });
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const remote = await response.json();
        const remoteVersion = remote.version || "unknown";

        if (compareVersions(remoteVersion, localVersion) > 0) {
            setUpdateMessage(`Update available: ${remoteVersion}`);
        } else {
            setUpdateMessage("Up to date");
        }
    } catch (error) {
        setUpdateMessage("Check failed");
    }
}

function loadSavedVideoPreset() {
    try {
        const saved = localStorage.getItem(VIDEO_PRESET_STORAGE_KEY);
        if (saved && saved.trim() && fileExists(saved)) {
            videoPresetPath = saved;
            return;
        }
    } catch (error) {}

    videoPresetPath = DEFAULT_VIDEO_PRESET_PATH;
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

function saveVideoPreset(nextPath) {
    videoPresetPath = nextPath;

    try {
        localStorage.setItem(VIDEO_PRESET_STORAGE_KEY, nextPath);
    } catch (error) {}

    document.getElementById("videoPresetPath").textContent = videoPresetPath;
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
        setStatus("Export destination selected. Ready.");
    }
}

async function chooseAlignFolder() {
    if (busy) {
        return;
    }

    const result = window.cep.fs.showOpenDialogEx(false, true, "Choose Existing Export Folder");
    if (result.data && result.data.length > 0) {
        alignFolder = result.data[0];
        try {
            localStorage.setItem(ALIGN_FOLDER_STORAGE_KEY, alignFolder);
        } catch (error) {}
        document.getElementById("alignPath").textContent = alignFolder;
        setStatus("Align folder selected. Ready.");
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

async function runExport() {
    if (busy) {
        return;
    }

    if (!exportFolder) {
        alert("Choose an export folder first.");
        return;
    }

    const exportMp3 = document.getElementById("exportMp3Checkbox").checked;
    const exportWav = document.getElementById("exportWavCheckbox").checked;
    const videoTrackNumber = getPositiveIntValue("videoTrackInput", 1);
    const audioStartTrackNumber = getPositiveIntValue("audioStartTrackInput", 1);
    const saveAlignManifest = document.getElementById("saveAlignManifestCheckbox").checked;

    if (!fileExists(videoPresetPath)) {
        alert("The selected video preset file was not found. Choose the video preset again.");
        return;
    }

    if (exportMp3 && !fileExists(MP3_PRESET_PATH)) {
        alert("The MP3 preset file was not found.");
        return;
    }

    if (exportWav && !fileExists(WAV_PRESET_PATH)) {
        alert("The WAV preset file was not found.");
        return;
    }

    setBusyState(true);
    setStatus("Loading Premiere host script...");

    if (!(await ensureHostLoaded())) {
        setBusyState(false);
        return;
    }

    setStatus("Queueing Media Encoder jobs...");

    const script = `exportBackup.runBackupQueue("${escapeForEvalScript(exportFolder)}","${escapeForEvalScript(videoPresetPath)}","${escapeForEvalScript(MP3_PRESET_PATH)}","${escapeForEvalScript(WAV_PRESET_PATH)}",${exportMp3},${exportWav},${videoTrackNumber},${audioStartTrackNumber},${saveAlignManifest})`;
    const result = await callHost(script);
    const parsed = parseHostResult(result);

    if (parsed && parsed.message) {
        setStatus(parsed.message);
    } else {
        setStatus(result || "No response returned from Premiere.");
    }

    setBusyState(false);
}

async function alignExistingFolder() {
    if (busy) {
        return;
    }

    if (!alignFolder) {
        alert("Choose an align folder first.");
        return;
    }

    setBusyState(true);
    setStatus("Loading Premiere host script...");

    if (!(await ensureHostLoaded())) {
        setBusyState(false);
        return;
    }

    const videoTrackNumber = getPositiveIntValue("alignVideoTrackInput", 1);
    const audioStartTrackNumber = getPositiveIntValue("alignAudioStartTrackInput", 1);

    setStatus("Scanning folder and aligning files...");

    const script = `exportBackup.alignExistingFolder("${escapeForEvalScript(alignFolder)}",${videoTrackNumber},${audioStartTrackNumber})`;
    const result = await callHost(script);
    const parsed = parseHostResult(result);

    if (parsed && parsed.message) {
        setStatus(parsed.message);
    } else {
        setStatus(result || "No response returned from Premiere.");
    }

    setBusyState(false);
}

document.addEventListener("DOMContentLoaded", () => {
    readVersionInfo();
    loadSavedVideoPreset();
    loadSavedPaths();
    checkForUpdates();
    document.getElementById("videoPresetPath").textContent = videoPresetPath;
    document.getElementById("audioPresetPath").textContent = `MP3: ${MP3_PRESET_PATH}\nWAV: ${WAV_PRESET_PATH}`;
    setStatus("Waiting to start.");
});
