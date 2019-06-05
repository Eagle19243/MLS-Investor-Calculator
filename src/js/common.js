function disableExtension() {
    chrome.browserAction.disable();
}

function enableExtension() {
    chrome.browserAction.disable();
}

// Inject CSS and JS
function injectFiles(tab) {
    chrome.tabs.insertCSS(tab.id, { file: 'vendor/bootstrap/bootstrap.min.css' });
    chrome.tabs.insertCSS(tab.id, { file: 'vendor/bootstrap/bootstrap-extend.min.css' });
    chrome.tabs.insertCSS(tab.id, { file: 'css/style.css' });
    chrome.tabs.executeScript(tab.id, { file: 'vendor/jquery/jquery.min.js' });
    chrome.tabs.executeScript(tab.id, { file: 'vendor/bootstrap/bootstrap.min.js' });
    chrome.tabs.executeScript(tab.id, { file: 'vendor/js-xlsx/xlsx.full.min.js' });
    chrome.tabs.executeScript(tab.id, { file: 'vendor/editabletable/editabletable.js' });
    chrome.tabs.executeScript(tab.id, { file: 'js/popup.js' });
}
