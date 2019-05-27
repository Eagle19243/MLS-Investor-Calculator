chrome.browserAction.onClicked.addListener(_onBrowserActionClicked);

function _onBrowserActionClicked(tab) {
    if (!tab.url) {
        return;
    }
    
    if (URLRegex.test(tab.url)) {
        injectFiles(tab);
    } else {
        // show alert message
    }
}

