// Collector routes to a category and a project folder with the category first.
// Collector has its own destination root; category names need no @ prefix.
const COLLECTOR_SHOW_ROOT = "Z:\\2017\\_SMTV2 PROJECT FOLDER\\@ SHOWS";
const BACKUP_CATEGORIES = [
    "BMD INTRO", "DAILY NEWS SCROLLS", "CTAW", "GPGW", "SHOW",
    "BMD", "BRE", "GAT", "GOL", "MOS", "NWN", "PCC", "SWA", "VEG", "WAU",
    "WOW", "AR", "AP", "AW", "CS", "EB", "GG", "HL", "KW",
    "LS", "NB", "PE", "SS", "UL", "VE", "VR"
];
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

function resolveCollectorCategoryPath(projectPath) {
    if (!projectPath) throw new Error('Save the Premiere project before using show category folders.');
    const projectName = path.basename(projectPath).replace(/\.prproj$/i, '');
    const searchable = ' ' + projectName.replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase() + ' ';
    const category = getConfiguredBackupCategories().find(name => searchable.indexOf(' ' + name + ' ') >= 0);
    if (!category) throw new Error('The project name must contain one of the built-in show categories.');
    const categoryPattern = category.split(' ').map(word => word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('[\\s_-]+');
    const match = new RegExp('(^|[\\s_-])(' + categoryPattern + ')(?=$|[\\s_-])', 'i').exec(projectName);
    const categoryStart = match.index + match[1].length;
    const before = projectName.slice(0, categoryStart).replace(/[\s_-]+$/, '');
    const after = projectName.slice(categoryStart + match[2].length).replace(/^[\s_-]+/, '');
    const folderName = [category, before, after].filter(Boolean).join(' ');
    // Preview does not access the network. The copy action creates this folder.
    return path.join(COLLECTOR_SHOW_ROOT, category, folderName);
}
