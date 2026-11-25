/* ppt time. Read the active WPS PDF and report to this user's local timer. */
(function () {
    "use strict";
    var session = Date.now().toString(36) + "-" + Math.random().toString(36).slice(2);
    var sequence = 0, timer = 0, subscribed = null, eventPage = 0;
    var playing = null, playChangedAt = 0, lastDoc = "", lastOk = 0, lastPage = "";

    function read(obj, key) {
        try {
            if (!obj) return undefined;
            var member = obj[key];
            return typeof member === "function" ? member.call(obj) : member;
        } catch (_) { return undefined; }
    }

    function positive() {
        for (var i = 0; i < arguments.length; i++) {
            var n = Number(arguments[i]);
            if (Number.isSafeInteger(n) && n > 0) return n;
        }
        return null;
    }

    async function property(obj, key) {
        try { return await read(obj, key); }
        catch (_) { return undefined; }
    }

    async function windowPage(pane) {
        var view = read(pane, "ActiveView");
        return positive(await property(pane, "CurrentViewPagelndex"), await property(view, "CurrentPagelndex"),
                        await property(pane, "CurrentViewPageIndex"), await property(view, "CurrentPageIndex"));
    }

    function report(snapshot, order) {
        var bridge = window.pptTimePdfBridge;
        if (!bridge || !/^http:\/\/127\.0\.0\.1:\d+\/report$/.test(bridge.url)) return;
        snapshot.key = bridge.key;
        snapshot.session = session;
        snapshot.seq = order;
        var xhr = new XMLHttpRequest();
        xhr.open("POST", bridge.url, true);
        xhr.timeout = 1600;
        xhr.setRequestHeader("Content-Type", "text/plain;charset=UTF-8");
        xhr.onload = function () { if (xhr.status === 204) lastOk = Date.now(); };
        xhr.send(JSON.stringify(snapshot));
    }

    function subscribe(events) {
        if (!events || events === subscribed) return;
        subscribed = events;
        function listen(name, action) {
            try {
                var earlier = events[name];
                events[name] = function () {
                    try { action.apply(null, arguments); }
                    finally { if (typeof earlier === "function") earlier.apply(this, arguments); }
                };
            } catch (_) { /* Some WPS editions expose only the PDF properties. */ }
        }
        listen("StartPlay", function () { playing = true; playChangedAt = Date.now(); sample(); });
        listen("EndPlay", function () { playing = false; playChangedAt = Date.now(); sample(); });
        listen("CurrentPageChange", function (page) {
            eventPage = positive(page, read(page, "CurrentPage")) || 0;
            sample();
        });
    }

    async function sample() {
        var order = ++sequence;
        try {
            var app = window.Application || (window.wps && window.wps.Application);
            if (!app) return;
            subscribe(read(app, "Sub"));
            var pdfHost = read(app, "PdfApplication");
            var doc = read(app, "ActivePDF") || read(app, "PDFActivePDF") ||
                      read(pdfHost, "ActivePDF") || read(app, "ActiveDocument");
            if (!doc) {
                lastDoc = ""; eventPage = 0; playing = null;
                report({active:false}, order);
                return;
            }
            var name = String((await property(doc, "FullName")) || (await property(doc, "Name")) || "");
            if (lastDoc && name && name !== lastDoc) { eventPage = 0; playing = null; }
            lastDoc = name;
            var page = (await windowPage(read(app, "ActiveWindow"))) || (await windowPage(read(doc, "ActiveWindow"))) || eventPage;
            var count = positive(await property(doc, "PageCount"), await property(doc, "PagesCount"),
                                 await property(doc, "TotalPages"), await property(doc, "GetPageCount"),
                                 await property(read(doc, "Pages"), "Count"), await property(pdfHost, "PagesCount"));
            var mode = await property(doc, "PlayMode");
            if (mode === undefined) mode = await property(pdfHost, "PlayMode");
            var screenMode = await property(doc, "GetScreenModeInPlay");
            if (screenMode === undefined) screenMode = await property(pdfHost, "GetScreenModeInPlay");
            var reportedPlay = mode === true || mode === "active" || mode === "start";
            var orientation = typeof screenMode === "string" ? screenMode : screenMode && screenMode.mode;
            var active = reportedPlay || orientation === "vertical" || orientation === "horizontal";
            if (playing !== null) {
                if (Date.now() - playChangedAt < 1600 || (mode === undefined && screenMode === undefined))
                    active = playing;
                else if (playing === false && !reportedPlay)
                    active = false;
            }
            lastPage = page && count && page <= count ? (page + "/" + count) : "";
            report({active:active, current:page, total:count, document:name}, order);
        } catch (_) { /* A closing document can briefly invalidate the host's objects. */ }
    }

    window.pptTimeLoaded = function () {
        if (!timer) timer = window.setInterval(sample, 850);
        sample();
        return true;
    };
    window.pptTimeStatus = function () {
        sample();
        var ok = Date.now() - lastOk < 3600;
        window.alert(ok ? ("已连接 ppt time" + (lastPage ? " · " + lastPage : "")) :
                          "等待 ppt time。请打开桌面计时器，并在 WPS PDF 加载项菜单中查看连接状态。");
    };
    window.addEventListener("pagehide", function () { report({active:false}, ++sequence); });
    window.addEventListener("DOMContentLoaded", window.pptTimeLoaded);
}());
