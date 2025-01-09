/*
 *  Microsoft 365 Link Opener [https://micz.it/thunderbird-addon-microsoft365linkopener/]
 *  Copyright (C) 2024 - 2025  Mic (m@micz.it)

 *  This program is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation, either version 3 of the License, or
 *  (at your option) any later version.

 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU General Public License for more details.

 *  You should have received a copy of the GNU General Public License
 *  along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */


// Register the message display script for all newly opened message tabs.
messenger.messageDisplayScripts.register({
    js: [{ file: "js/mzmo-message-content-script.js" }],
    css: [{ file: "css/mzmo-message-content-styles.css" }],
});

// Inject script and CSS in all already open message tabs.
messenger.micz = {};
let openTabs = await messenger.tabs.query();
let messageTabs = openTabs.filter(
    tab => ["mail", "messageDisplay"].includes(tab.type)
);
for (let messageTab of messageTabs) {
    browser.tabs.executeScript(messageTab.id, {
        file: "js/mzmo-message-content-script.js"
    })
    browser.tabs.insertCSS(messageTab.id, {
        file: "css/mzmo-message-content-styles.css"
    })
}