function saveOptions(e) {
  e.preventDefault();
  browser.storage.sync.set({
    force_msedge: document.querySelector("#force_msedge").checked,
  });
}

function restoreOptions() {
  function setCurrentChoice(result) {
    document.querySelector("#force_msedge").checked = result.force_msedge || false;
  }

  function onError(error) {
    console.log(`Error: ${error}`);
  }

  let getting = browser.storage.sync.get("force_msedge");
  getting.then(setCurrentChoice, onError);
}

document.addEventListener('DOMContentLoaded', () => {
  restoreOptions();
  i18n.updateDocument();
}, { once: true });

document.querySelector("#force_msedge").addEventListener("change", saveOptions);
