// Keep each panel's DOM, options and jobs independent when switching tabs.
function selectPluginTab(name) {
    if (!['export', 'collector'].includes(name)) return;
    const collector = name === 'collector';
    if (document.body) document.body.setAttribute('data-active-tool', name);
    document.getElementById('exportPanel').hidden = collector;
    const panel = document.getElementById('collectorPanel');
    if (collector && !panel.getAttribute('src')) panel.src = 'collector/index.html';
    panel.hidden = !collector;
    if (collector && panel.contentWindow && panel.contentWindow.syncCollectorFrameSize) {
        panel.contentWindow.syncCollectorFrameSize();
    }
    window.scrollTo(0, 0);
    ['export', 'collector'].forEach((id) => {
        document.getElementById(id + 'Tab').setAttribute('aria-selected', String(id === name));
    });
}