// CEP exposes its native bridge on the extension window. Reuse it in this frame.
if (window.parent !== window) {
    window.__adobe_cep__ = window.parent.__adobe_cep__;
    window.cep = window.parent.cep;
    if (typeof require === 'undefined') window.require = window.parent.require;
    document.documentElement.classList.add('embedded-collector');

    // Size the frame to its contents so the extension owns the only page scrollbar.
    window.syncCollectorFrameSize = function () {
        const frame = window.frameElement;
        if (!frame || frame.hidden || !document.body) return;
        const height = Math.ceil(document.body.getBoundingClientRect().height);
        if (frame.style.height !== height + 'px') frame.style.height = height + 'px';
        syncCollectorDialogs();
    };

    function syncCollectorDialogs() {
        const frame = window.frameElement;
        if (!frame || frame.hidden) return;
        const rect = frame.getBoundingClientRect();
        const top = Math.max(0, -rect.top);
        const height = Math.max(0, Math.min(rect.bottom, window.parent.innerHeight) - Math.max(0, rect.top));
        document.querySelectorAll('.done-prompt').forEach(function (dialog) {
            // Keep dialogs in the visible parent viewport, even on a tall collector page.
            dialog.style.top = top + 'px';
            dialog.style.height = height + 'px';
            dialog.style.bottom = 'auto';
            dialog.style.setProperty('--dialog-viewport-height', height + 'px');
        });
    }

    document.addEventListener('DOMContentLoaded', function () {
        if (typeof ResizeObserver !== 'undefined') {
            new ResizeObserver(window.syncCollectorFrameSize).observe(document.body);
        } else {
            new MutationObserver(window.syncCollectorFrameSize).observe(document.querySelector('.app'), {
                subtree: true, childList: true, attributes: true, characterData: true
            });
        }
        window.syncCollectorFrameSize();
    });
    window.addEventListener('resize', window.syncCollectorFrameSize);
    window.parent.addEventListener('scroll', syncCollectorDialogs);
    window.parent.addEventListener('resize', window.syncCollectorFrameSize);
}
