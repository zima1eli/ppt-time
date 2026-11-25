// Run with node --test tests/test_wps_presentation_addin.js.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const test = require("node:test");

const assets = path.join(__dirname, "..", "wps_presentation_addin");
const script = fs.readFileSync(path.join(assets, "main.js"), "utf8");
const config = {url:"http://127.0.0.1:48111/report", key:"x".repeat(40)};
const deckName = "Example.dps";

function makeShow(name = deckName) {
    const doc = {FullName:name, Slides:{Count:12}};
    const view = {State:1, Slide:{SlideIndex:1}, CurrentShowPosition:1};
    const show = {Presentation:doc, View:view, HWND:0};
    doc.SlideShowWindow = show;
    return {doc, view, show};
}

function host({autoIndex = false, emitEvents = true, deferred = false, application = "Application"} = {}) {
    const messages = [], events = {}, dom = {}, subscriptions = {}, loadedScripts = [];
    const docList = [], showList = [];
    const collection = list => ({
        get Count() { return list.length; },
        Item(i) { assert.ok(i >= 1 && i <= list.length); return list[i - 1]; }
    });
    const app = {SlideShowWindows:collection(showList), Presentations:collection(docList)};
    if (emitEvents) app.ApiEvent = {AddApiEventListener(name, fn) {
        subscriptions[name] = 1 + (subscriptions[name] || 0); events[name] = fn;
    }};
    const document = {
        readyState:deferred ? "loading" : "complete",
        createElement() { return {}; },
        head:{appendChild(element) {
            loadedScripts.push(element.src);
            window.pptTimeWppBridge = config;
            element.onload();
        }}
    };
    let tick, timerCount = 0, reply = "", disposed = false;
    const window = {
        setInterval(fn) { tick = fn; return ++timerCount; },
        clearInterval() { disposed = true; },
        addEventListener(event, fn) { dom[event] = fn; },
        alert(message) { reply = message; }
    };
    if (!autoIndex) window.pptTimeWppBridge = config;
    if (application === "Application") window.Application = app;
    else window.wps = {WppApplication: () => app};
    class Request {
        open(method, url, async) {
            assert.deepEqual([method, url, async], ["POST", config.url, true]);
        }
        setRequestHeader(name, value) { assert.equal(name, "Content-Type"); assert.ok(value.startsWith("text/plain")); }
        send(data) {
            const message = JSON.parse(data);
            assert.equal(message.key, config.key);
            assert.ok(message.session && message.seq > (messages.at(-1)?.seq || 0));
            if (messages.length) assert.equal(message.session, messages[0].session);
            messages.push(message);
            this.status = 204; this.onload();
        }
    }
    vm.runInNewContext(script, {window, document, XMLHttpRequest:Request});
    return {window, app, events, dom, messages, subscriptions, loadedScripts, docList, showList,
        pulse:() => tick(), latest:() => messages.at(-1), times:() => timerCount,
        status:() => {window.pptTimeWppStatus(); return reply;}, disposed:() => disposed};
}

test("documented WPP ribbon and both local entry points start exactly one reader", () => {
    const xml = fs.readFileSync(path.join(assets, "ribbon.xml"), "utf8");
    const html = fs.readFileSync(path.join(assets, "index.html"), "utf8");
    assert.match(xml, /onLoad="pptTimeWppLoaded"/);
    assert.match(xml, /onAction="pptTimeWppStatus"/);
    assert.ok(html.indexOf("connection.js") < html.indexOf("main.js"));
    const h = host({autoIndex:true, deferred:true, application:"WppApplication"});
    assert.equal(h.times(), 0);
    h.dom.DOMContentLoaded(); h.window.pptTimeWppLoaded();
    assert.deepEqual(h.loadedScripts, ["connection.js"]);
    assert.equal(h.times(), 1);
    assert.deepEqual(h.subscriptions, {SlideShowBegin:1, SlideShowNextSlide:1, SlideShowEnd:1});
    assert.equal(h.latest().active, false);
    assert.match(h.status(), /已连接/);
});

test("external F5, forward and back jumps, end before collection removal, and reopen", () => {
    const h = host(), a = makeShow();
    h.showList.push(a.show); h.docList.push(a.doc); h.app.ActivePresentation = a.doc;
    a.show.HWND = -2147483648;
    h.events.SlideShowBegin(a.show);
    assert.deepEqual([h.latest().current, h.latest().total, h.latest().hwnd, h.latest().document], [1, 12, 2147483648, deckName]);
    a.view.Slide.SlideIndex = 10; h.events.SlideShowNextSlide(a.show);
    assert.equal(h.latest().current, 10);
    a.view.Slide.SlideIndex = 2; h.pulse();
    assert.equal(h.latest().current, 2);
    assert.match(h.status(), /2\/12/);
    h.events.SlideShowEnd(a.doc); h.pulse(); assert.equal(h.latest().active, false);
    h.showList.length = 0; h.pulse();
    a.view.Slide.SlideIndex = 7; h.showList.push(a.show); h.pulse();
    assert.equal(h.latest().current, 7);
    a.view.State = 5; h.pulse();
    assert.equal(h.latest().ended, true); assert.equal(h.latest().current, null);
    h.dom.pagehide(); const n = h.messages.length;
    assert.equal(h.latest().active, false); assert.equal(h.disposed(), true);
    h.events.SlideShowNextSlide(a.show); h.pulse(); assert.equal(h.messages.length, n);
});

test("fallback uses the document's running view after a throwing or partial global item", () => {
    const h = host({emitEvents:false}), a = makeShow(), editing = {FullName:"Editing.dps"};
    h.app.ActivePresentation = editing; h.docList.push(editing, a.doc); h.showList.push(a.show);
    h.app.SlideShowWindows.Item = () => { throw new Error("E_FAIL"); };
    a.view.Slide.SlideIndex = 5; h.pulse(); assert.equal(h.latest().current, 5);
    h.app.SlideShowWindows.Item = () => ({Presentation:a.doc, View:{State:1, CurrentShowPosition:2}});
    a.view.Slide.SlideIndex = 8; h.pulse();
    assert.deepEqual([h.latest().current, h.latest().total], [8, 12]);
    delete a.view.Slide; a.view.CurrentShowPosition = 2; h.pulse();
    assert.equal(h.latest().current, null); assert.equal(h.latest().show_position, 2);
    h.showList.length = 0; h.pulse(); assert.equal(h.latest().active, false);
});

test("a handle-less projector reports the showing document's owner separately", () => {
    const h = host(), a = makeShow();
    h.app.HWND = 23;
    h.app.ActivePresentation = {Windows:{Item(){return {HWND:99};}}};
    a.doc.Windows = {Item(){return {HWND:-2147483648};}};
    h.showList.push(a.show); h.pulse();
    assert.deepEqual([h.latest().hwnd, h.latest().owner_hwnd], [0, 2147483648]);
    delete a.doc.Windows; h.pulse();
    assert.deepEqual([h.latest().hwnd, h.latest().owner_hwnd], [0, 23]);
    h.app.HWND = true; h.pulse();
    assert.deepEqual([h.latest().hwnd, h.latest().owner_hwnd], [0, 0]);
});

test("unknown collections, editing, and simultaneous shows cannot report a guessed page", () => {
    const h = host(), a = makeShow(), b = makeShow("Second.dps");
    h.app.ActivePresentation = {FullName:deckName}; h.docList.push(h.app.ActivePresentation);
    h.pulse(); assert.equal(h.latest().active, false);
    h.showList.push(a.show, b.show); h.events.SlideShowBegin(a.show); assert.equal(h.latest().active, false);
    h.app.SlideShowWindows = null; h.app.ActivePresentation = a.doc;
    h.docList.push(b.doc); h.pulse(); assert.equal(h.latest().active, false);
    h.docList.splice(0, 2, a.doc); h.pulse(); assert.equal(h.latest().current, 1);
    a.view.Slide.SlideIndex = 13; h.pulse(); assert.equal(h.latest().current, null);
    a.view.Slide.SlideIndex = 4; delete a.doc.Slides; h.pulse();
    assert.equal(h.latest().current, 4); assert.equal(h.latest().total, null);
});

test("event-only show reference is cleared when the host stops projecting", () => {
    const h = host(), a = makeShow();
    h.app.SlideShowWindows = {Count:1, Item() { throw Error("show item unavailable"); }};
    h.events.SlideShowBegin(a.show); assert.equal(h.latest().current, 1);
    a.view.Slide.SlideIndex = 3; h.pulse(); assert.equal(h.latest().current, 3);
    h.app.SlideShowWindows.Count = 0; h.pulse(); assert.equal(h.latest().active, false);
    h.app.SlideShowWindows.Count = 1; h.pulse(); assert.equal(h.latest().active, false);
});

