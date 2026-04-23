(function () {
    'use strict';

    // Central API base for production deployments (e.g., Vercel frontend + Render backend).
    var API_BASE = '';
    if (window.APP_CONFIG && typeof window.APP_CONFIG.API_BASE_URL === 'string') {
        API_BASE = window.APP_CONFIG.API_BASE_URL.trim();
    }
    if (API_BASE && API_BASE.charAt(API_BASE.length - 1) === '/') {
        API_BASE = API_BASE.slice(0, -1);
    }

    function apiUrl(path) {
        if (!path) return API_BASE || '/';
        if (path.charAt(0) !== '/') path = '/' + path;
        return API_BASE ? (API_BASE + path) : path;
    }

    /* ──────────────────────────────────────────────
       State Manager
       ────────────────────────────────────────────── */
    const AppState = {
        sessionId: null,
        files: [],       // [{id, original_name, size, size_display, slide_count}]
        mergedFile: null,     // filename string
        pdfFile: null,     // filename string
        warnings: [],       // string[]
        busy: false,
        currentPage: 'upload', // upload | organize | result
    };

    /* ──────────────────────────────────────────────
       DOM Cache
       ────────────────────────────────────────────── */
    const Dom = {
        toastBox: $('#toastBox'),
        procOverlay: $('#procOverlay'),
        procTitle: $('#procTitle'),
        procSub: $('#procSub'),
        dropZone: $('#dropZone'),
        fileInput: $('#fileInput'),
        fileList: $('#fileList'),
        fileCountNum: $('#fileCountNum'),
        fileCountSlides: $('#fileCountSlides'),
        btnMerge: $('#btnMerge'),
        btnDlPptx: $('#btnDlPptx'),
        btnConvPdf: $('#btnConvertPdf'),
        btnDlPdf: $('#btnDlPdf'),
        pdfDownloadArea: $('#pdfDownloadArea'),
        pdfConvertArea: $('#pdfConvertArea'),
        warningsArea: $('#warningsArea'),
        statSlides: $('#statSlides'),
        statFiles: $('#statFiles'),
        statWarnings: $('#statWarnings'),
        resultSummary: $('#resultSummary'),
        pages: {
            upload: $('#pageUpload'),
            organize: $('#pageOrganize'),
            result: $('#pageResult'),
        },
        tabs: $('.nav-tab'),
        seps: $('.tab-sep'),
    };

    /* ──────────────────────────────────────────────
       Toast Module
       ────────────────────────────────────────────── */
    const Toast = {
        ICONS: {
            success: 'fa-circle-check',
            error: 'fa-circle-xmark',
            warn: 'fa-triangle-exclamation',
            info: 'fa-circle-info',
        },

        show(message, type) {
            type = type || 'info';
            var iconClass = this.ICONS[type] || this.ICONS.info;
            var $t = $(
                '<div class="toast-item t-' + type + '" role="alert">' +
                '<i class="fa-solid ' + iconClass + '"></i>' +
                '<span>' + escapeHtml(message) + '</span>' +
                '<button class="toast-close" aria-label="Dismiss"><i class="fa-solid fa-xmark"></i></button>' +
                '</div>'
            );
            $t.find('.toast-close').on('click', function () { Toast._remove($t); });
            Dom.toastBox.append($t);
            requestAnimationFrame(function () { $t.addClass('show'); });
            setTimeout(function () { Toast._remove($t); }, 5500);
        },

        success(msg) { this.show(msg, 'success'); },
        error(msg) { this.show(msg, 'error'); },
        warn(msg) { this.show(msg, 'warn'); },
        info(msg) { this.show(msg, 'info'); },

        _remove($t) {
            $t.removeClass('show');
            setTimeout(function () { $t.remove(); }, 350);
        },
    };

    /* ──────────────────────────────────────────────
       Processing Overlay Module
       ────────────────────────────────────────────── */
    var Overlay = {
        show: function (title, sub) {
            Dom.procTitle.text(title);
            Dom.procSub.text(sub || 'This may take a moment');
            Dom.procOverlay.addClass('active');
        },
        hide: function () {
            Dom.procOverlay.removeClass('active');
        },
    };

    /* ──────────────────────────────────────────────
       Router / Page Navigation Module
       ────────────────────────────────────────────── */
    var PAGES = ['upload', 'organize', 'result'];

    var Router = {
        pageOrder: { upload: 0, organize: 1, result: 2 },

        go: function (pageName) {
            if (AppState.busy) return;
            var targetIdx = this.pageOrder[pageName];
            if (targetIdx === undefined) return;

            // Deactivate all pages
            PAGES.forEach(function (p) { Dom.pages[p].removeClass('active'); });
            // Activate target
            Dom.pages[pageName].addClass('active');
            AppState.currentPage = pageName;

            // Scroll to top of the page container
            Dom.pages[pageName].scrollTop = 0;

            // Update nav tabs
            Dom.tabs.each(function () {
                var tab = $(this);
                var tabPage = tab.data('page');
                var tabIdx = Router.pageOrder[tabPage];
                tab.removeClass('active done');
                if (tabIdx < targetIdx) tab.addClass('done');
                else if (tabIdx === targetIdx) tab.addClass('active');
            });
            Dom.seps.each(function () {
                var sep = $(this);
                var sepIdx = parseInt(sep.data('sep'), 10);
                sep.toggleClass('done', sepIdx < targetIdx);
            });
        },
    };

    /* ──────────────────────────────────────────────
       Upload Module
       ────────────────────────────────────────────── */
    var Upload = {
        init: function () {
            // Click-to-browse
            Dom.dropZone.on('click keydown', function (e) {
                if (e.type === 'keydown' && e.key !== 'Enter' && e.key !== ' ') return;
                e.preventDefault();
                Dom.fileInput.trigger('click');
            });

            Dom.fileInput.on('change', function () {
                if (this.files.length) Upload.handleFiles(this.files);
                this.value = '';
            });

            // Drag and drop
            Dom.dropZone.on('dragover', function (e) {
                e.preventDefault();
                $(this).addClass('drag-over');
            });
            Dom.dropZone.on('dragleave drop', function (e) {
                e.preventDefault();
                $(this).removeClass('drag-over');
            });
            Dom.dropZone.on('drop', function (e) {
                var files = e.originalEvent.dataTransfer.files;
                if (files.length) Upload.handleFiles(files);
            });
        },

        handleFiles: function (fileList) {
            var valid = [];
            for (var i = 0; i < fileList.length; i++) {
                var f = fileList[i];
                if (!f.name.toLowerCase().endsWith('.pptx')) {
                    Toast.warn('"' + f.name + '" is not a .pptx file — skipped.');
                    continue;
                }
                if (f.size > 100 * 1024 * 1024) {
                    Toast.warn('"' + f.name + '" exceeds 100 MB — skipped.');
                    continue;
                }
                valid.push(f);
            }
            if (!valid.length) return;

            AppState.busy = true;
            Overlay.show('Uploading files', valid.length + ' file' + (valid.length > 1 ? 's' : ''));

            var fd = new FormData();
            valid.forEach(function (f) { fd.append('files', f); });
            if (AppState.sessionId) fd.append('session_id', AppState.sessionId);

            $.ajax({
                url: apiUrl('/upload'),
                type: 'POST',
                data: fd,
                processData: false,
                contentType: false,
                success: function (res) {
                    AppState.sessionId = res.session_id;
                    res.files.forEach(function (f) {
                        if (!AppState.files.find(function (x) { return x.id === f.id; })) {
                            AppState.files.push(f);
                        }
                    });
                    FileList.render();
                    Router.go('organize');
                    Toast.success(res.files.length + ' file' + (res.files.length > 1 ? 's' : '') + ' uploaded successfully.');
                },
                error: function (xhr) {
                    var msg = (xhr.responseJSON && xhr.responseJSON.detail) || 'Upload failed.';
                    Toast.error(msg);
                },
                complete: function () {
                    AppState.busy = false;
                    Overlay.hide();
                },
            });
        },
    };

    /* ──────────────────────────────────────────────
       File List Module (Organize Page)
       ────────────────────────────────────────────── */
    var FileList = {
        dragIdx: null,

        init: function () {
            // Delegate all events to the container
            Dom.fileList
                .on('click', '.btn-up', function () {
                    FileList.move($(this).closest('.file-item').data('idx'), -1);
                })
                .on('click', '.btn-down', function () {
                    FileList.move($(this).closest('.file-item').data('idx'), 1);
                })
                .on('click', '.btn-rm', function () {
                    FileList.remove($(this).closest('.file-item').data('idx'));
                })
                .on('dragstart', '.file-item', function (e) {
                    FileList.dragIdx = $(this).data('idx');
                    $(this).addClass('dragging');
                    e.originalEvent.dataTransfer.effectAllowed = 'move';
                })
                .on('dragend', '.file-item', function () {
                    FileList.dragIdx = null;
                    $(this).removeClass('dragging');
                    Dom.fileList.find('.file-item').removeClass('drag-target');
                })
                .on('dragover', '.file-item', function (e) {
                    e.preventDefault();
                    if (FileList.dragIdx === null) return;
                    var ti = $(this).data('idx');
                    if (ti !== FileList.dragIdx) $(this).addClass('drag-target');
                })
                .on('dragleave', '.file-item', function () {
                    $(this).removeClass('drag-target');
                })
                .on('drop', '.file-item', function (e) {
                    e.preventDefault();
                    $(this).removeClass('drag-target');
                    if (FileList.dragIdx === null) return;
                    var ti = $(this).data('idx');
                    if (ti === FileList.dragIdx) return;
                    var item = AppState.files.splice(FileList.dragIdx, 1)[0];
                    AppState.files.splice(ti, 0, item);
                    FileList.dragIdx = null;
                    FileList.render();
                });
        },

        render: function () {
            var files = AppState.files;
            if (!files.length) {
                Dom.fileList.html(
                    '<div class="empty-state">' +
                    '<i class="fa-solid fa-inbox"></i>' +
                    '<p>No files added yet.</p>' +
                    '<button class="btn-ghost" id="btnBrowseEmpty"><i class="fa-solid fa-plus"></i> Browse Files</button>' +
                    '</div>'
                );
                Dom.btnMerge.prop('disabled', true);
                // Bind the empty-state button
                Dom.fileList.find('#btnBrowseEmpty').on('click', function () {
                    Dom.fileInput.trigger('click');
                });
                this._updateCounts();
                return;
            }

            Dom.btnMerge.prop('disabled', false);
            var html = '';
            files.forEach(function (f, i) {
                var slideStr = (f.slide_count != null)
                    ? f.slide_count + ' slide' + (f.slide_count !== 1 ? 's' : '')
                    : 'unknown slides';
                html +=
                    '<div class="file-item" draggable="true" data-idx="' + i + '">' +
                    '<span class="fi-grip" title="Drag to reorder"><i class="fa-solid fa-grip-vertical"></i></span>' +
                    '<span class="fi-pos">' + (i + 1) + '</span>' +
                    '<span class="fi-icon"><i class="fa-solid fa-file-powerpoint"></i></span>' +
                    '<div class="fi-info">' +
                    '<div class="fi-name" title="' + escapeHtml(f.original_name) + '">' + escapeHtml(f.original_name) + '</div>' +
                    '<div class="fi-meta">' +
                    '<span><i class="fa-solid fa-hard-drive"></i> ' + f.size_display + '</span>' +
                    '<span><i class="fa-solid fa-layer-group"></i> ' + slideStr + '</span>' +
                    '</div>' +
                    '</div>' +
                    '<div class="fi-actions">' +
                    '<button class="fi-btn btn-up" title="Move up"' + (i === 0 ? ' disabled' : '') + '>' +
                    '<i class="fa-solid fa-chevron-up"></i>' +
                    '</button>' +
                    '<button class="fi-btn btn-down" title="Move down"' + (i === files.length - 1 ? ' disabled' : '') + '>' +
                    '<i class="fa-solid fa-chevron-down"></i>' +
                    '</button>' +
                    '<button class="fi-btn fi-del btn-rm" title="Remove file">' +
                    '<i class="fa-solid fa-xmark"></i>' +
                    '</button>' +
                    '</div>' +
                    '</div>';
            });
            Dom.fileList.html(html);
            this._updateCounts();
        },

        _updateCounts: function () {
            var files = AppState.files;
            Dom.fileCountNum.text(files.length);
            var totalSlides = 0;
            files.forEach(function (f) {
                if (f.slide_count != null) totalSlides += f.slide_count;
            });
            Dom.fileCountSlides.text(totalSlides);
        },

        move: function (idx, dir) {
            var ni = idx + dir;
            if (ni < 0 || ni >= AppState.files.length) return;
            var tmp = AppState.files[idx];
            AppState.files[idx] = AppState.files[ni];
            AppState.files[ni] = tmp;
            this.render();
        },

        remove: function (idx) {
            var f = AppState.files[idx];
            AppState.files.splice(idx, 1);
            // Notify server
            if (AppState.sessionId) {
                $.post(apiUrl('/remove-file'), { session_id: AppState.sessionId, file_id: f.id });
            }
            this.render();
            if (!AppState.files.length) {
                AppState.sessionId = null;
                Router.go('upload');
            }
            Toast.info('"' + f.original_name + '" removed.');
        },
    };

    /* ──────────────────────────────────────────────
       Merge Module
       ────────────────────────────────────────────── */
    var Merge = {
        run: function () {
            if (AppState.busy || !AppState.files.length) return;
            AppState.busy = true;
            Overlay.show('Merging presentations', AppState.files.length + ' file' + (AppState.files.length > 1 ? 's' : '') + ' in order');

            var fd = new FormData();
            fd.append('session_id', AppState.sessionId);
            fd.append('file_order', AppState.files.map(function (f) { return f.id; }).join(','));

            $.ajax({
                url: apiUrl('/process'),
                type: 'POST',
                data: fd,
                processData: false,
                contentType: false,
                success: function (res) {
                    AppState.mergedFile = res.output_file;
                    AppState.warnings = res.warnings || [];
                    AppState.pdfFile = null;
                    Result.render(res);
                    Router.go('result');
                    Toast.success('Merge complete — ' + res.total_slides + ' slides.');
                },
                error: function (xhr) {
                    var msg = (xhr.responseJSON && xhr.responseJSON.detail) || 'Merge failed.';
                    Toast.error(msg);
                },
                complete: function () {
                    AppState.busy = false;
                    Overlay.hide();
                },
            });
        },
    };

    /* ──────────────────────────────────────────────
       Result Module
       ────────────────────────────────────────────── */
    var Result = {
        render: function (res) {
            // Stats
            Dom.statSlides.text(res.total_slides);
            Dom.statFiles.text(AppState.files.length);
            Dom.statWarnings.text(AppState.warnings.length);
            Dom.resultSummary.text(
                res.total_slides + ' slide' + (res.total_slides !== 1 ? 's' : '') +
                ' merged from ' + AppState.files.length + ' file' + (AppState.files.length !== 1 ? 's' : '')
            );

            // Warnings
            Dom.warningsArea.empty();
            if (AppState.warnings.length) {
                var html = '<div class="result-section"><div class="warnings-box"><strong><i class="fa-solid fa-triangle-exclamation"></i> Warnings</strong><ul>';
                AppState.warnings.forEach(function (w) {
                    html += '<li>' + escapeHtml(w) + '</li>';
                });
                html += '</ul></div></div>';
                Dom.warningsArea.html(html);
            }

            // Reset PDF area
            Dom.pdfDownloadArea.addClass('hide');
            Dom.pdfConvertArea.show();
            Dom.btnConvPdf
                .prop('disabled', false)
                .html('<i class="fa-solid fa-wand-magic-sparkles"></i> Convert to PDF');
        },
    };

    /* ──────────────────────────────────────────────
       PDF Conversion Module
       ────────────────────────────────────────────── */
    var PdfConvert = {
        run: function () {
            if (AppState.busy || !AppState.mergedFile) return;
            AppState.busy = true;
            Dom.btnConvPdf
                .prop('disabled', true)
                .html('<i class="fa-solid fa-spinner fa-spin"></i> Converting...');
            Overlay.show('Converting to PDF', 'This may take a moment for large presentations');

            var fd = new FormData();
            fd.append('session_id', AppState.sessionId);
            fd.append('output_file', AppState.mergedFile);

            $.ajax({
                url: apiUrl('/convert'),
                type: 'POST',
                data: fd,
                processData: false,
                contentType: false,
                success: function (res) {
                    AppState.pdfFile = res.pdf_file;
                    Dom.pdfConvertArea.hide();
                    Dom.pdfDownloadArea.removeClass('hide');
                    Toast.success('PDF generated via ' + res.method + '.');
                },
                error: function (xhr) {
                    var msg = (xhr.responseJSON && xhr.responseJSON.detail) || 'Conversion failed.';
                    Toast.error(msg);
                    Dom.btnConvPdf
                        .prop('disabled', false)
                        .html('<i class="fa-solid fa-wand-magic-sparkles"></i> Convert to PDF');
                },
                complete: function () {
                    AppState.busy = false;
                    Overlay.hide();
                },
            });
        },
    };

    /* ──────────────────────────────────────────────
       Download Module
       ────────────────────────────────────────────── */
    var Download = {
        pptx: function () {
            if (!AppState.sessionId || !AppState.mergedFile) return;
            window.location.href = apiUrl('/download') + '?session_id=' + encodeURIComponent(AppState.sessionId) + '&file_type=pptx';
            Toast.info('Downloading PPTX...');
        },
        pdf: function () {
            if (!AppState.sessionId || !AppState.pdfFile) return;
            window.location.href = apiUrl('/download') + '?session_id=' + encodeURIComponent(AppState.sessionId) + '&file_type=pdf';
            Toast.info('Downloading PDF...');
        },
    };

    /* ──────────────────────────────────────────────
       Utility Functions
       ────────────────────────────────────────────── */
    function escapeHtml(str) {
        var div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    function resetAll() {
        AppState.sessionId = null;
        AppState.files = [];
        AppState.mergedFile = null;
        AppState.pdfFile = null;
        AppState.warnings = [];
        AppState.busy = false;
        Dom.btnConvPdf
            .prop('disabled', false)
            .html('<i class="fa-solid fa-wand-magic-sparkles"></i> Convert to PDF');
        FileList.render();
        Router.go('upload');
        Toast.info('Session reset.');
    }

    /* ──────────────────────────────────────────────
       Event Bindings
       ────────────────────────────────────────────── */
    function bindEvents() {
        // Nav logo — go to upload (only if on result)
        $('#navLogo').on('click', function (e) {
            e.preventDefault();
            if (AppState.currentPage !== 'upload') Router.go('upload');
        });

        // Nav tab clicks — allow going back, not forward
        Dom.tabs.on('click', function () {
            var target = $(this).data('page');
            var currentIdx = Router.pageOrder[AppState.currentPage];
            var targetIdx = Router.pageOrder[target];
            // Can go back or to current
            if (targetIdx <= currentIdx) Router.go(target);
        });

        // Add More button
        $('#btnAddMore').on('click', function () { Dom.fileInput.trigger('click'); });

        // Back button on organize page
        $('#btnBackUpload').on('click', function () { Router.go('upload'); });

        // Merge button
        Dom.btnMerge.on('click', function () { Merge.run(); });

        // Download PPTX
        Dom.btnDlPptx.on('click', function () { Download.pptx(); });

        // Convert to PDF
        Dom.btnConvPdf.on('click', function () { PdfConvert.run(); });

        // Download PDF
        Dom.btnDlPdf.on('click', function () { Download.pdf(); });

        // Start Over
        $('#btnStartOver').on('click', function () { resetAll(); });
    }

    /* ──────────────────────────────────────────────
       Initialization
       ────────────────────────────────────────────── */
    function init() {
        Upload.init();
        FileList.init();
        bindEvents();
        FileList.render();
        Router.go('upload');
    }

    // Run when DOM is ready
    $(init);

})();