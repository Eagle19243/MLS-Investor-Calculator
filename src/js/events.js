chrome.browserAction.onClicked.addListener(_onBrowserActionClicked);

function _onBrowserActionClicked(tab) {
    if (!tab.url) {
        return;
    }
    
    if (SPINURLRegex.test(tab.url) || MATRIXURLRegex.test(tab.url)) {
        injectFiles(tab);
    } else {
        // show alert message
    }
}

