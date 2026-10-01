const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');
const root = path.join(__dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8').replace(/\r\n/g, '\n');
const declaration = (source, name) => source.match(new RegExp('^function ' + name + '\\([^\\n]*\\).*?^}', 'ms'))[0];

test('Collector category normalization and backup detection stay identical to ExportBackup rules', () => {
    const exportMain = read('js/main.js');
    const routing = read('collector/js/backup-routing.js');
    for (const name of ['normalizeBackupCategoryName','getConfiguredBackupCategories']) {
        assert.equal(declaration(routing, name), declaration(exportMain, name));
    }
    const collectorHost = read('collector/jsx/collector.jsx');
    const exportHost = read('jsx/export.jsx');
    for (const match of collectorHost.matchAll(/^function (pcEb\w+)\(/gm)) {
        const name = match[1];
        assert.equal(declaration(collectorHost, name).replace(/\bpcEb(\w+)/g, 'eb$1'), declaration(exportHost, name.replace(/^pcEb/, 'eb')));
    }
});

test('switching tabs preserves the Collector frame and does not invoke jobs', () => {
    const elements = {};
    for (const id of ['exportPanel','collectorPanel','exportTab','collectorTab']) {
        elements[id] = {hidden:false, setAttribute(key, value) {this[key]=value;}, getAttribute(key) {return this[key];}};
    }
    const context = vm.createContext({window:{scrollTo() {}}, document:{getElementById:id => elements[id]}});
    vm.runInContext(read('js/tabs.js'), context);
    context.selectPluginTab('collector');
    assert.equal(elements.collectorPanel.src, 'collector/index.html');
    assert.equal(elements.exportPanel.hidden, true);
    elements.collectorPanel.job = 'retained';
    context.selectPluginTab('export');
    context.selectPluginTab('collector');
    assert.equal(elements.collectorPanel.job, 'retained');
    assert.equal(elements.collectorTab['aria-selected'], 'true');
});

test('Collector previews exact project-name folders without accessing the show root', () => {
    const expectedRoot = 'Z:\\2017\\_SMTV2 PROJECT FOLDER\\@ SHOWS';
    const context = vm.createContext({
        path: path.win32,
        configuredBackupCategories: ['WOW'],
        fs: {
            existsSync() { throw new Error('Preview must not check network availability'); },
            readdirSync() { throw new Error('Preview must not scan folders'); }
        }
    });
    vm.runInContext(read('collector/js/backup-routing.js'), context);
    assert.equal(context.resolveCollectorCategoryPath('D:\\WOW 3250 Title.prproj'),
        expectedRoot + '\\WOW\\WOW 3250 Title');
    assert.equal(context.resolveCollectorCategoryPath('D:\\3279 3280 WOW Socrates_Is virtue known or learned.prproj'),
        expectedRoot + '\\WOW\\WOW 3279 3280 Socrates_Is virtue known or learned');
    assert.equal(context.resolveCollectorCategoryPath('D:\\3279 WOW INTRO.prproj'),
        expectedRoot + '\\WOW\\WOW 3279 INTRO');
    assert.equal(context.resolveCollectorCategoryPath('D:\\3279_WOW_Socrates_Is virtue.prproj'),
        expectedRoot + '\\WOW\\WOW 3279 Socrates_Is virtue');
});

test('embedded Collector resizes in both directions and keeps dialogs in the parent viewport', () => {
    let bodyHeight = 1600;
    const frame = {hidden:false, style:{}, getBoundingClientRect: () => ({top:-300, bottom:bodyHeight-300})};
    const dialog = {style:{setProperty(name, value) {this[name]=value;}}};
    const parent = {innerHeight:700, addEventListener() {}};
    const context = vm.createContext({
        require,
        window:{parent, frameElement:frame, addEventListener() {}},
        document:{
            documentElement:{classList:{add() {}}},
            body:{getBoundingClientRect: () => ({height:bodyHeight})},
            addEventListener() {}, querySelectorAll: () => [dialog]
        }
    });
    vm.runInContext(read('collector/js/embedded.js'), context);
    context.window.syncCollectorFrameSize();
    assert.equal(frame.style.height, '1600px');
    assert.equal(dialog.style.top, '300px');
    assert.equal(dialog.style.height, '700px');
    bodyHeight = 500;
    context.window.syncCollectorFrameSize();
    assert.equal(frame.style.height, '500px');
    frame.hidden = true;
    bodyHeight = 0;
    context.window.syncCollectorFrameSize();
    assert.equal(frame.style.height, '500px');
});
