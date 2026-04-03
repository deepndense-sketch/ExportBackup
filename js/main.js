const csInterface = new CSInterface();
const fs = require("fs");
const path = require("path");

const DEFAULT_VIDEO_PRESET_PATH = "D:\\Work\\Tools\\ExportBackup\\presets\\1080 AIR.epr";
const MP3_PRESET_PATH = "D:\\Work\\Tools\\ExportBackup\\presets\\mp3.epr";
const WAV_PRESET_PATH = "C:\\Program Files\\Adobe\\Adobe Media Encoder 2026\\MediaIO\\systempresets\\3F3F3F3F_57415645\\Waveform Audio 48kHz 16-bit.epr";
const VIDEO_PRESET_STORAGE_KEY = "exportbackup.videoPresetPath";

let exportFolder = null;
let hostLoaded = false;
let busy = false;
let videoPresetPath = DEFAULT_VIDEO_PRESET_PATH;

function setStatus(message) {
    document.getElementById("statusBox").textContent = message;
}

function setBusyState(nextBusy) {
    busy = nextBusy;
    document.getElementById("chooseFolderButton").disabled = nextBusy;
    document.getElementById("chooseVideoPresetButton").disabled = nextBusy;
    document.getElementById("exportButton").disabled = nextBusy;
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

async function chooseExportFolder() {
    if (busy) {
        return;
    }

    const result = window.cep.fs.showOpenDialogEx(false, true, "Choose Export Folder");
    if (result.data && result.data.length > 0) {
        exportFolder = result.data[0];
        document.getElementById("exportPath").textContent = exportFolder;
        setStatus("Export destination selected. Ready.");
    }
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

function saveVideoPreset(nextPath) {
    videoPresetPath = nextPath;

    try {
        localStorage.setItem(VIDEO_PRESET_STORAGE_KEY, nextPath);
    } catch (error) {}

    document.getElementById("videoPresetPath").textContent = videoPresetPath;
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

    const script = `exportBackup.runBackupQueue("${escapeForEvalScript(exportFolder)}","${escapeForEvalScript(videoPresetPath)}","${escapeForEvalScript(MP3_PRESET_PATH)}","${escapeForEvalScript(WAV_PRESET_PATH)}",${exportMp3},${exportWav})`;
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
    loadSavedVideoPreset();
    document.getElementById("videoPresetPath").textContent = videoPresetPath;
    document.getElementById("audioPresetPath").textContent = `MP3: ${MP3_PRESET_PATH}\nWAV: ${WAV_PRESET_PATH}`;
    setStatus("Waiting to start.");
});
