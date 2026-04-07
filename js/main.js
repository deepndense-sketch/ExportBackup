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
const DEFAULT_ALIGN_VIDEO_TRACK = 5;
const DEFAULT_ALIGN_VIDEO_AUDIO_TRACK = 1;
const DEFAULT_ALIGN_AUDIO_START_TRACK = 2;

let exportFolder = null;
let alignFolder = null;
let hostLoaded = false;
let busy = false;
let videoPresetPath = "";
let mp3PresetPath = "";
let wavPresetPath = "";
let localVersion = "unknown";
let remoteVersion = null;
let presetSectionVisible = true;

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

function setBusyState(nextBusy) {
    busy = nextBusy;
    document.getElementById("chooseFolderButton").disabled = nextBusy;
    document.getElementById("chooseVideoPresetButton").disabled = nextBusy;
    document.getElementById("chooseMp3PresetButton").disabled = nextBusy;
    document.getElementById("chooseWavPresetButton").disabled = nextBusy;
    document.getElementById("exportButton").disabled = nextBusy;
    document.getElementById("chooseAlignFolderButton").disabled = nextBusy;
    document.getElementById("alignFolderButton").disabled = nextBusy;
    document.getElementById("updateButton").disabled = nextBusy;
}

function setUpdateButton(label, isUpdateAvailable) {
    const button = document.getElementById("updateButton");
    button.textContent = label;
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

function getPositiveIntValue(elementId, fallbackValue) {
    const raw = document.getElementById(elementId).value;
    const parsed = parseInt(raw, 10);
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

function getAlignmentDefaultValues() {
    return {
        videoTrackNumber: DEFAULT_ALIGN_VIDEO_TRACK,
        videoAudioTrackNumber: DEFAULT_ALIGN_VIDEO_AUDIO_TRACK,
        audioStartTrackNumber: DEFAULT_ALIGN_AUDIO_START_TRACK
    };
}

function getAlignmentInputs() {
    return {
        videoTrack: document.getElementById("alignVideoTrackInput"),
        videoAudioTrack: document.getElementById("alignVideoAudioTrackInput"),
        audioStartTrack: document.getElementById("alignAudioStartTrackInput")
    };
}

function getBackupInputs() {
    return {
        videoTrack: document.getElementById("exportVideoTrackInput")
    };
}

function applyBackupDefaults(defaults, force) {
    const values = defaults || getAlignmentDefaultValues();
    const inputs = getBackupInputs();

    if (!inputs.videoTrack) {
        return;
    }

    if (force || inputs.videoTrack.dataset.userEdited !== "true") {
        inputs.videoTrack.value = String(values.videoTrackNumber);
        inputs.videoTrack.dataset.autoValue = String(values.videoTrackNumber);
    }
}

function applyAlignmentDefaults(defaults, force) {
    const values = defaults || getAlignmentDefaultValues();
    const inputs = getAlignmentInputs();

    [
        { element: inputs.videoTrack, value: values.videoTrackNumber },
        { element: inputs.videoAudioTrack, value: values.videoAudioTrackNumber },
        { element: inputs.audioStartTrack, value: values.audioStartTrackNumber }
    ].forEach((entry) => {
        if (!entry.element) {
            return;
        }

        const parsedValue = parseInt(entry.value, 10) || 0;
        if (parsedValue < 1) {
            return;
        }

        if (force || entry.element.dataset.userEdited !== "true") {
            entry.element.value = String(parsedValue);
            entry.element.dataset.autoValue = String(parsedValue);
        }
    });
}

function markBackupInputsDirty() {
    const inputs = getBackupInputs();

    [inputs.videoTrack].forEach((input) => {
        if (!input) {
            return;
        }

        input.addEventListener("input", () => {
            input.dataset.userEdited = "true";
        });
    });
}

function markAlignmentInputsDirty() {
    Object.values(getAlignmentInputs()).forEach((input) => {
        if (!input) {
            return;
        }

        input.addEventListener("input", () => {
            input.dataset.userEdited = "true";
        });
    });
}

async function refreshAlignmentDefaults(force) {
    const fallback = getAlignmentDefaultValues();

    if (!(await ensureHostLoaded())) {
        applyBackupDefaults(fallback, force);
        applyAlignmentDefaults(fallback, force);
        return;
    }

    const result = await callHost("exportBackup.getAlignmentDefaults()");
    const parsed = parseHostResult(result);

    if (!parsed || !parsed.ok) {
        applyBackupDefaults(fallback, force);
        applyAlignmentDefaults(fallback, force);
        return;
    }

    const suggestedDefaults = {
        videoTrackNumber: parsed.suggestedVideoTrack || DEFAULT_ALIGN_VIDEO_TRACK,
        videoAudioTrackNumber: parsed.suggestedVideoAudioTrack || DEFAULT_ALIGN_VIDEO_AUDIO_TRACK,
        audioStartTrackNumber: parsed.suggestedAudioStartTrack || DEFAULT_ALIGN_AUDIO_START_TRACK
    };

    applyBackupDefaults(suggestedDefaults, force);
    applyAlignmentDefaults(suggestedDefaults, force);
}

function getManifestAlignmentDefaults(matchInfo) {
    if (!matchInfo) {
        return null;
    }

    const videoTrackNumber = parseInt(matchInfo.manifestBackupVideoTrackNumber, 10) || 0;
    const videoAudioTrackNumber = parseInt(matchInfo.manifestBackupVideoAudioTrackNumber, 10) || 0;
    const audioStartTrackNumber = parseInt(matchInfo.manifestAudioStartTrackNumber, 10) || 0;

    if (!videoTrackNumber && !videoAudioTrackNumber && !audioStartTrackNumber) {
        return null;
    }

    return {
        videoTrackNumber,
        videoAudioTrackNumber,
        audioStartTrackNumber
    };
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
    } catch (error) {
        localVersion = "unknown";
        if (!silent) {
            setStatus(`Could not read version file.\n${error.message}`);
        }
    }

    document.getElementById("versionValue").textContent = localVersion;
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
    setUpdateButton("Checking...", false);

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

        if (compareVersions(remoteVersion, localVersion) > 0) {
            setUpdateButton(`Update to ${remoteVersion}`, true);
        } else {
            setUpdateButton("Up to date", false);
        }
    } catch (error) {
        setUpdateButton("Check failed", false);
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
                const raw = fs.readFileSync(resultPath, "utf8");
                const parsed = JSON.parse(raw);
                if (parsed.ok) {
                    readVersionInfo(true);
                    await checkForUpdates();
                    setStatus(`Update complete.\nInstalled version: ${localVersion}\nRestart Premiere Pro if the panel was already open.`);
                    return;
                }

                setStatus(
                    `Updater failed.\n${parsed.message || "Unknown error."}\n` +
                    `Log: ${parsed.logPath || getTempUpdaterLogPath()}`
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
        if (savedVideo && savedVideo.trim() && fileExists(savedVideo)) {
            videoPresetPath = savedVideo;
        } else {
            videoPresetPath = defaults.video;
        }
    } catch (error) {
        videoPresetPath = defaults.video;
    }

    try {
        const savedMp3 = localStorage.getItem(MP3_PRESET_STORAGE_KEY);
        if (savedMp3 && savedMp3.trim() && fileExists(savedMp3)) {
            mp3PresetPath = savedMp3;
        } else {
            mp3PresetPath = defaults.mp3;
        }
    } catch (error) {
        mp3PresetPath = defaults.mp3;
    }

    try {
        const savedWav = localStorage.getItem(WAV_PRESET_STORAGE_KEY);
        if (savedWav && savedWav.trim() && fileExists(savedWav)) {
            wavPresetPath = savedWav;
        } else {
            wavPresetPath = defaults.wav;
        }
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

function loadSavedUiState() {
    try {
        const saved = localStorage.getItem(PRESET_SECTION_VISIBLE_STORAGE_KEY);
        if (saved === "false") {
            presetSectionVisible = false;
            return;
        }
    } catch (error) {}

    presetSectionVisible = true;
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

function scanExportFolderForSequence(folderPath, sequenceName) {
    const sanitizedBase = sanitizeSequenceName(sequenceName);
    const entries = fs.readdirSync(folderPath, { withFileTypes: true });
    const files = entries.filter((entry) => entry.isFile()).map((entry) => entry.name);
    const lowerBase = sanitizedBase.toLowerCase();
    const backupPrefix = `${lowerBase}_backup.`;
    const manifestName = `${sanitizedBase}_ALIGN.json`;
    const manifestPath = path.join(folderPath, manifestName);

    let manifest = null;
    if (fs.existsSync(manifestPath)) {
        try {
            manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
        } catch (error) {}
    }

    let videoPath = "";
    if (manifest && manifest.videoFile && fs.existsSync(manifest.videoFile)) {
        videoPath = manifest.videoFile;
    }

    const audio = [];

    files.forEach((fileName) => {
        const absolutePath = path.join(folderPath, fileName);
        const lowerName = fileName.toLowerCase();

        if (!videoPath && lowerName.startsWith(backupPrefix)) {
            videoPath = absolutePath;
            return;
        }

        const trackNumber = parseTrackNumberFromFileName(fileName, sanitizedBase);
        if (trackNumber > 0) {
            audio.push({
                path: absolutePath,
                trackNumber,
                name: fileName
            });
        }
    });

    audio.sort((a, b) => a.trackNumber - b.trackNumber);

    return {
        baseName: sanitizedBase,
        manifestPath: fs.existsSync(manifestPath) ? manifestPath : "",
        manifestVideoFile: manifest && manifest.videoFile ? manifest.videoFile : "",
        manifestBackupVideoTrackNumber: manifest && manifest.backupVideoTrackNumber ? manifest.backupVideoTrackNumber : 0,
        manifestBackupVideoAudioTrackNumber: manifest && manifest.backupVideoAudioTrackNumber ? manifest.backupVideoAudioTrackNumber : 0,
        manifestAudioStartTrackNumber: manifest && manifest.audioStartTrackNumber ? manifest.audioStartTrackNumber : 0,
        videoPath,
        audio,
        folderFiles: files
    };
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
        await refreshAlignmentDefaults(false);

        try {
            const activeSequenceName = await getActiveSequenceName();
            if (activeSequenceName) {
                const matchInfo = scanExportFolderForSequence(alignFolder, activeSequenceName);
                const manifestDefaults = getManifestAlignmentDefaults(matchInfo);
                if (manifestDefaults) {
                    applyAlignmentDefaults(manifestDefaults, false);
                }
            }
        } catch (error) {}

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

    const exportMp3 = document.getElementById("exportMp3Checkbox").checked;
    const exportWav = document.getElementById("exportWavCheckbox").checked;

    if (!fileExists(videoPresetPath)) {
        alert("The selected video preset file was not found. Choose the video preset again.");
        return;
    }

    if (exportMp3 && !fileExists(mp3PresetPath)) {
        alert("The MP3 preset file was not found.");
        return;
    }

    if (exportWav && !fileExists(wavPresetPath)) {
        alert("The WAV preset file was not found.");
        return;
    }

    const backupVideoTrackNumber = getPositiveIntValue("exportVideoTrackInput", DEFAULT_ALIGN_VIDEO_TRACK);

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

    setStatus("Queueing Media Encoder jobs...");

    const script = `exportBackup.runBackupQueue("${escapeForEvalScript(exportFolder)}","${escapeForEvalScript(videoPresetPath)}","${escapeForEvalScript(mp3PresetPath)}","${escapeForEvalScript(wavPresetPath)}",${exportMp3},${exportWav},${backupVideoTrackNumber})`;
    const result = await callHost(script);
    const parsed = parseHostResult(result);

    if (parsed && parsed.ok === false) {
        setStatus(parsed.message || "Backup export failed.");
        showBlockingMessage(parsed.message || "Backup export failed.");
    } else if (parsed && parsed.message) {
        setStatus(parsed.message);
        showBlockingMessage("Done.");
    } else {
        setStatus(result || "No response returned from Premiere.");
        showBlockingMessage("Done.");
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
        showBlockingMessage("Could not load Premiere host script.");
        setBusyState(false);
        return;
    }

    const activeSequenceName = await getActiveSequenceName();
    if (!activeSequenceName) {
        setStatus("No active sequence is open in Premiere Pro.");
        showBlockingMessage("No active sequence is open in Premiere Pro.");
        setBusyState(false);
        return;
    }

    setStatus("Scanning folder and aligning files...");

    let matchInfo = null;
    try {
        matchInfo = scanExportFolderForSequence(alignFolder, activeSequenceName);
    } catch (error) {
        setStatus(`Could not read the chosen folder.\n${error.message}`);
        showBlockingMessage(`Could not read the chosen folder.\n${error.message}`);
        setBusyState(false);
        return;
    }

    if (!matchInfo.videoPath && matchInfo.audio.length === 0) {
        const message =
            "No files could be matched in the chosen folder.\n" +
            `Sequence base: ${matchInfo.baseName}\n` +
            `Manifest file: ${matchInfo.manifestPath || "(none)"}\n` +
            `Manifest video: ${matchInfo.manifestVideoFile || "(none)"}\n` +
            `Folder files: ${matchInfo.folderFiles.join(" | ")}`;
        setStatus(message);
        showBlockingMessage(message);
        setBusyState(false);
        return;
    }

    const manifestDefaults = getManifestAlignmentDefaults(matchInfo);
    if (manifestDefaults) {
        applyAlignmentDefaults(manifestDefaults, false);
    }

    const videoTrackNumber = getPositiveIntValue("alignVideoTrackInput", DEFAULT_ALIGN_VIDEO_TRACK);
    const videoAudioTrackNumber = getPositiveIntValue("alignVideoAudioTrackInput", DEFAULT_ALIGN_VIDEO_AUDIO_TRACK);
    const audioStartTrackNumber = getPositiveIntValue("alignAudioStartTrackInput", DEFAULT_ALIGN_AUDIO_START_TRACK);
    const skipBackupVideo = document.getElementById("alignSkipVideoCheckbox").checked;

    const audioJson = JSON.stringify(matchInfo.audio);
    const resolvedVideoPath = skipBackupVideo ? "" : (matchInfo.videoPath || "");
    const script = `exportBackup.alignMatchedFiles("${escapeForEvalScript(resolvedVideoPath)}","${escapeForEvalScript(audioJson)}",${videoTrackNumber},${videoAudioTrackNumber},${audioStartTrackNumber})`;
    const result = await callHost(script);
    const parsed = parseHostResult(result);

    if (parsed && parsed.ok === false) {
        setStatus(parsed.message || "Alignment failed.");
        showBlockingMessage(parsed.message || "Alignment failed.");
    } else if (parsed && parsed.message) {
        setStatus(parsed.message);
        showBlockingMessage("Done.");
    } else {
        setStatus(result || "No response returned from Premiere.");
        showBlockingMessage("Done.");
    }

    setBusyState(false);
}

document.addEventListener("DOMContentLoaded", () => {
    readVersionInfo();
    loadSavedPresets();
    loadSavedPaths();
    loadSavedUiState();
    markBackupInputsDirty();
    markAlignmentInputsDirty();
    applyBackupDefaults(getAlignmentDefaultValues(), true);
    applyAlignmentDefaults(getAlignmentDefaultValues(), true);
    setPresetSectionVisibility(presetSectionVisible);
    setUpdateButton("Check for Updates", false);
    checkForUpdates();
    document.getElementById("videoPresetPath").textContent = videoPresetPath;
    updateAudioPresetDisplay();
    setStatus("Ready.");
    refreshAlignmentDefaults(false);
});
