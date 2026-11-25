/* Read the actual running slide show in WPS Presentation. */
(function () {
    "use strict";
    if (window.pptTimeWppLoaded) return;

    var session = Date.now().toString(36) + "-" + Math.random().toString(36).slice(2);
    var sequence = 0, timer = 0, configRequested = false, stopped = false;
    var eventWindow = null, endedDocument = null, lastReceipt = 0, lastPage = "";
    var subscribed = {}, nextSubscription = 0;

    function read(target, name) {
        try {
            if (target == null) return null;
            var value = target[name];
            return typeof value === "function" ? value.call(target) : value;
        } catch (_) { return null; }
    }

    function number(value) {
        if (typeof value !== "number" && !(typeof value === "string" && /^\d+$/.test(value))) return null;
        var n = Number(value);
        return isFinite(n) && Math.floor(n) === n && n >= 0 && n <= 2147483647 ? n : null;
    }

    function item(collection, index) {
        try { return collection.Item(index); } catch (_) { return null; }
    }

    function documentName(doc) {
        var name = read(doc, "FullName") || read(doc, "Name");
        return typeof name === "string" ? name : "";
    }

    function viewOf(show) {
        return read(show, "View") ? show : null;
    }

    function windowHandle(target) {
        var value = read(target, "HWND");
        return typeof value === "number" && Math.floor(value) === value &&
               value >= -2147483648 && value <= 4294967295 ? value >>> 0 : 0;
    }

    function snapshot(show, app) {
        var view = read(show, "View");
        if (!view) return null;
        var slide = read(view, "Slide");
        var doc = read(show, "Presentation") || read(slide, "Parent");
        var current = number(read(slide, "SlideIndex")) || null;
        var total = number(read(read(doc, "Slides"), "Count")) || null;
        var owner = windowHandle(item(read(doc, "Windows"), 1)) || windowHandle(app);
        var ended = number(read(view, "State")) === 5;
        // CurrentShowPosition is the order in a custom show, not the slide's physical index.
        var order = number(read(view, "CurrentShowPosition")) || null;
        if (ended || (total && current > total)) current = null;
        return {active:true, ended:ended, current:current, total:total, show_position:ended ? null : order,
                document:documentName(doc), hwnd:windowHandle(show), owner_hwnd:owner};
    }

    function currentState(app) {
        var windows = read(app, "SlideShowWindows");
        var count = number(read(windows, "Count"));
        if (count === 0) { eventWindow = null; endedDocument = null; return null; }
        if (count !== null && count > 1) { eventWindow = null; return null; }

        var docs = read(app, "Presentations");
        var numDocs = number(read(docs, "Count"));
        // Without a global count, only a single open presentation uniquely identifies its show.
        if (count === null && numDocs !== 1) return null;
        var best = null, bestRank = -1;
        function consider(show) {
            var state = snapshot(show, app);
            if (!state || (best && best.document && state.document && best.document !== state.document)) return false;
            var rank = state.ended ? 4 : state.current ? (state.total ? 3 : 2) : state.show_position ? 1 : 0;
            if (rank > bestRank) { best = state; bestRank = rank; }
            return rank >= 3;
        }
        if (consider(item(windows, 1))) return best;
        var active = read(app, "ActivePresentation");
        if (consider(read(active, "SlideShowWindow")) || consider(eventWindow)) return best;
        // The editor in focus may differ from the presentation on the projector.
        for (var i = 1; numDocs !== null && i <= Math.min(numDocs, 128); i++) {
            if (consider(read(item(docs, i), "SlideShowWindow"))) return best;
        }
        return best;
    }

    function report(value) {
        lastPage = value.current ? (value.total ? value.current + "/" + value.total : "第 " + value.current + " 页") : "";
        var bridge = window.pptTimeWppBridge;
        if (!bridge || typeof bridge.key !== "string" || !bridge.key ||
                !/^http:\/\/127\.0\.0\.1:\d+\/report$/.test(bridge.url)) return;
        value.key = bridge.key;
        value.session = session;
        value.seq = ++sequence;
        try {
            var xhr = new XMLHttpRequest();
            xhr.open("POST", bridge.url, true);
            xhr.timeout = 1600;
            xhr.setRequestHeader("Content-Type", "text/plain;charset=UTF-8");
            xhr.onload = function () { if (xhr.status === 204) lastReceipt = Date.now(); };
            xhr.send(JSON.stringify(value));
        } catch (_) { /* The desktop timer may have been stopped. Retry at the next heartbeat. */ }
    }

    function onSlide(show) {
        if (stopped) return;
        endedDocument = null;
        eventWindow = viewOf(show);
        sample();
    }

    function onEnd(doc) {
        if (stopped) return;
        // End can fire while the application's window collection still contains the closing show.
        endedDocument = documentName(doc);
        eventWindow = null;
        report({active:false});
    }

    function listen(app) {
        var now = Date.now();
        if (now < nextSubscription) return;
        nextSubscription = now + 5000;
        var events = read(app, "ApiEvent") || read(window.wps, "ApiEvent");
        if (!events) return;
        var handlers = {SlideShowBegin:onSlide, SlideShowNextSlide:onSlide, SlideShowEnd:onEnd};
        Object.keys(handlers).forEach(function (name) {
            if (subscribed[name]) return;
            try {
                events.AddApiEventListener(name, handlers[name]);
                subscribed[name] = true;
            } catch (_) { /* Older WPS versions can still be sampled on the timer. */ }
        });
    }

    function sample() {
        if (stopped) return;
        var app = read(window, "Application") || read(window.wps, "Application") || read(window.wps, "WppApplication");
        if (!app) return;
        listen(app);
        var state = currentState(app);
        if (state && endedDocument !== null && (!endedDocument || !state.document || state.document === endedDocument))
            state = null;
        report(state || {active:false});
    }

    window.pptTimeWppLoaded = function () {
        if (stopped) return true;
        if (!timer) timer = window.setInterval(sample, 750);
        if (!window.pptTimeWppBridge && !configRequested) {
            // Some WPS releases generate their own index.html and load only main.js.
            configRequested = true;
            var script = document.createElement("script");
            script.src = "connection.js";
            script.onload = sample;
            (document.head || document.documentElement).appendChild(script);
        }
        sample();
        return true;
    };
    window.pptTimeWppStatus = function () {
        sample();
        window.alert(lastReceipt && Date.now() - lastReceipt < 3500 ?
            "已连接 ppt time" + (lastPage ? " · " + lastPage : "，等待放映页码。") :
            "等待 ppt time 连接。请打开桌面计时器，在「WPS 演示页码联动」中查看连接状态。");
    };
    window.addEventListener("pagehide", function () {
        stopped = true;
        window.clearInterval(timer);
        eventWindow = null;
        report({active:false});
    });
    if (document.readyState === "loading") window.addEventListener("DOMContentLoaded", window.pptTimeWppLoaded);
    else window.pptTimeWppLoaded();
}());

