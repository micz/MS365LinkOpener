// Register the message display script for all newly opened message tabs.
messenger.messageDisplayScripts.register({
    js: [{ file: "js/message-content-script.js" }],
    css: [{ file: "css/message-content-styles.css" }],
});

// Inject script and CSS in all already open message tabs.
messenger.micz = {};
let openTabs = await messenger.tabs.query();
let messageTabs = openTabs.filter(
    tab => ["mail", "messageDisplay"].includes(tab.type)
);
for (let messageTab of messageTabs) {
    browser.tabs.executeScript(messageTab.id, {
        file: "js/message-content-script.js"
    })
    browser.tabs.insertCSS(messageTab.id, {
        file: "css/message-content-styles.css"
    })
}