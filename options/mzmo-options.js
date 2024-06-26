/*
 *  Microsoft 365 Link Opener [https://micz.it/thunderbird-addon-microsoft365linkopener/]
 *  Copyright (C) 2024  Mic (m@micz.it)

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


function saveOptions(e) {
  e.preventDefault();
  browser.storage.sync.set({
    force_msedge: document.querySelector("#force_msedge").checked,
    always_link: document.querySelector("#always_link").checked,
  });
}

function restoreOptions() {
  function setCurrentChoice(result) {
    document.querySelector("#force_msedge").checked = result.force_msedge || false;
    document.querySelector("#always_link").checked = result.always_link || false;
  }

  function onError(error) {
    console.log(`Error: ${error}`);
  }

  let getting = browser.storage.sync.get({force_msedge: false, always_link: false});
  getting.then(setCurrentChoice, onError);
}

document.addEventListener('DOMContentLoaded', () => {
  restoreOptions();
  i18n.updateDocument();
}, { once: true });

document.querySelector("#force_msedge").addEventListener("change", saveOptions);
document.querySelector("#always_link").addEventListener("change", saveOptions);
