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
