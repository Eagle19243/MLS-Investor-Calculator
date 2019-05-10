chrome.browserAction.onClicked.addListener(_onBrowserActionClicked);
chrome.tabs.onUpdated.addListener(_onTabUpdated);
// chrome.tabs.onActivated.addListener(_onTabActivated);

function _onBrowserActionClicked(tab) {
    if (URLRegex.test(tab.url)) {
        injectFiles(tab);
    }
}

function _onTabUpdated(tabId, changeInfo, tab) {
    determineAvailability(tab);
}

function _onTabActivated(activeInfo) {
    chrome.tabs.query({active: true}, function(tabs) {
        determineAvailability(tabs[0]);
    });
}