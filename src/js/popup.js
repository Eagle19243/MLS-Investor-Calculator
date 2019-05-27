init();

async function init() {
    await preparePopup();
    showPopup();
    populateValues();
    initEventHandler();
}

function initEventHandler() {
    $('.btn-contact').click(contactToAgent);
    $('.btn-export').click(exportToXLS);
}

function preparePopup() {
    return new Promise((resolve) => {
        if ($('.popup-container').length) {
            resolve();
        } else {
            const popupContainer = $('<section class="popup-container"></section>');
            const popupHTMLPath = chrome.extension.getURL('html/popup.html');
            
            $('body').prepend(popupContainer);
    
            popupContainer.load(popupHTMLPath, function(){
                resolve();
            });
        }
    });
}

function showPopup() {
    $('#mls_popup').modal('show');
}

function exportToXLS() {
    $('.table-analytics').table2csv('download', {
        trimContent: false,
        filename: `${getMLSNumber()}.csv`
    });
}

function contactToAgent() {
    const MailSubject = 'Additional information request';
    const MailBody    = 
        `Dear Sandra DeCola, \r\n \r\n \
        I am interested in ${getSubjectPropertyAddress()}, but see the listing is missing \
        the following data, (list of all data missing). \r\n \r\n \
        Please provide if possible. \r\n
        Thank you!`;
    const mailTo      = $('a.SideMenuAgent').attr('href');
    window.open(`${mailTo}?subject=${MailSubject}&body=${encodeURIComponent(MailBody)}`);
}

function populateValues() {
    $('.label-analytics').html(getSubjectPropertyAddress());
    $('#analytics_list_price').html(getListPrice());
    $('#analytics_heating').html(getHeating());
    $('#analytics_gas').html(getGas());
    $('#analytics_electricity').html(getElectricity());
    $('#analytics_water').html(getWater());
    $('#analytics_repairs').html(getRepairs());
    $('#analytics_trash').html(getTrash());
    $('#analytics_sewer').html(getSewer());
    $('#analytics_insurance').html(getInsurance());
    $('#analytics_management').html(getManagement());
    $('#analytics_miscellaneous').html(getMiscellaneous());
    $('#analytics_total').html(getTotalExpenses());
    $('#analytics_unit_1_beds').html(getUnitOneBeds());
    $('#analytics_unit_1_baths').html(getUnitOneBaths());
    $('#analytics_unit_1_current').html(getUnitOneCurrent());
    $('#analytics_unit_2_beds').html(getUnitTwoBeds());
    $('#analytics_unit_2_baths').html(getUnitTwoBaths());
    $('#analytics_unit_2_current').html(getUnitTwoCurrent());
    $('#analytics_unit_3_beds').html(getUnitThreeBeds());
    $('#analytics_unit_3_baths').html(getUnitThreeBaths());
    $('#analytics_unit_3_current').html(getUnitThreeCurrent());
    $('#analytics_unit_total_beds').html(getUnitTotalBeds());
    $('#analytics_unit_total_baths').html(getUnitTotalBaths());
    $('#analytics_unit_total_current').html(getUnitTotalCurrent());
    $('#analytics_number_of_units').html(getNumberOfUnits());
    $('#analytics_price_per_unit').html(getPricePerUnit());
    $('#analytics_gross_rents').html(getGrossRents());
    $('#analytics_vacancy_rate').html(getVacancyRate());
    $('#analytics_net_rents').html(getNetRents());
    $('#analytics_noi').html(getNOI());
    $('#analytics_mortgage_principal').html(getMortgagePrincipal());
    $('#analytics_net_cashflow').html(getNetCashflow());
    $('#analytics_cash').html(getCash());
    $('#analytics_cap').html(getCAP());
    $('#analytics_debt_service').html(getDebtService());
}

function getMLSNumber() {
    const mlsNumberContent = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(3) > table:nth-child(1) > tbody > tr > td > table > tbody > tr:nth-child(1) > td > b').html();
    const mlsNumber = /\d+/i.exec(mlsNumberContent)[0].trim();
    
    return mlsNumber;
}

function getSubjectPropertyAddress() {
    const address1 = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(3) > table:nth-child(2) > tbody > tr:nth-child(1) > td:nth-child(1) > b').html().trim();
    const address2 = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(3) > table:nth-child(2) > tbody > tr:nth-child(2) > td:nth-child(1) > b').html().trim();

    return `${address1}, ${address2}`;
}

function getListPrice() {
    const listPriceContent = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(3) > table:nth-child(2) > tbody > tr:nth-child(1) > td:nth-child(2) > b').html();
    const listPrice = listPriceContent.replace(/(<([^>]+)>)/ig,"").trim();
    
    return listPrice;
}

function getListPriceNumValue() {
    let listPrice = getListPrice();
    listPrice = Number(listPrice.substr(1, listPrice.length - 1).replace(',', ''));

    if (!listPrice) {
        return 0;
    }

    return listPrice
}

function getHeating() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(1) > td:nth-child(1) > b').html().trim();
    let value = 0;
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getGas() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(2) > td:nth-child(1) > b').html().trim();
    let value = 0;
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getElectricity() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(3) > td:nth-child(1) > b').html().trim();
    let value = 0;
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getWater() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(4) > td:nth-child(1) > b').html().trim();
    let value = 0;
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getRepairs() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(1) > td:nth-child(2) > b').html().trim();
    let value = 0;
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getTrash() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(2) > td:nth-child(2) > b').html().trim();
    const value = content.substr(1, content.length - 1);

    return (value.length === 0) ? '0.00' : value;
}

function getSewer() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(3) > td:nth-child(2) > b').html().trim();
    let value = 0;
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getInsurance() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(4) > td:nth-child(2) > b').html().trim();
    let value = 0;
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getManagement() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(1) > td:nth-child(3) > b').html().trim();
    let value = 0;
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getMiscellaneous() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(2) > td:nth-child(3) > b').html().trim();
    let value = 0;
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getTotalExpenses() {
    const total = Number(getHeating()) + Number(getGas()) + Number(getElectricity()) +
        Number(getWater()) + Number(getRepairs()) + Number(getTrash()) + Number(getSewer()) +
        Number(getInsurance()) + Number(getManagement()) + Number(getMiscellaneous());

    return (total === 0) ? '0.00' : total;
}

function getUnitOneBeds() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(9) > tbody > tr:nth-child(2) > td:nth-child(2) > b');
    let value = 0;
    
    if (content.length > 0) {
        value = Number(content.html().trim());
    }
    
    return value;
}

function getUnitOneBaths() {
    return 0;
}

function getUnitOneCurrent() {
    return 0;
}

function getUnitTwoBeds() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(9) > tbody > tr:nth-child(10) > td:nth-child(2) > b');
    let value = 0;
    
    if (content.length > 0) {
        value = Number(content.html().trim());
    }
    
    return value;
}

function getUnitTwoBaths() {
    return 0;
}

function getUnitTwoCurrent() {
    return 0;
}

function getUnitThreeBeds() {
    return 0;
}

function getUnitThreeBaths() {
    return 0;
}

function getUnitThreeCurrent() {
    return 0;
}

function getUnitTotalBeds() {
    return getUnitOneBeds() + getUnitTwoBeds() + getUnitThreeBeds();
}

function getUnitTotalBaths() {
    return getUnitOneBaths() + getUnitTwoBaths() + getUnitThreeBaths();
}

function getUnitTotalCurrent() {
    return getUnitOneCurrent() + getUnitTwoCurrent() + getUnitThreeCurrent();
}

function getNumberOfUnits() {
    const content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(3) > table:nth-child(2) > tbody > tr:nth-child(6) > td:nth-child(1) > b');
    let value = 0;
    
    if (content.length > 0) {
        value = Number(content.html().trim());
    }
    
    return value;
}

function getPricePerUnit() {
    if (getNumberOfUnits() === 0) {
        return '';
    }

    return '$' + formatNumber(parseInt(getListPriceNumValue() / getNumberOfUnits()));
}

function getGrossRents() {
    return (getUnitTotalCurrent() * 12).toFixed(2);
}

function getVacancyRate() {
    const vacancyRatePercent = Number($('#analytics_vacancy_rate_percent').html());
    return (getGrossRents() * vacancyRatePercent).toFixed(2);
}

function getNetRents() {
    return (getGrossRents() - getVacancyRate()).toFixed(2);
}

function getNOI() {function formatNumber(num) {
    return num.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,')
  }
    return 0;
}

function getMortgagePrincipal() {
    return 0;
}

function getNetCashflow() {
    return (getNOI() - getMortgagePrincipal()).toFixed(2);
}

function getCash() {
    return (getNetCashflow() / 1).toFixed(2);
}

function getCAP() {
    if (getListPriceNumValue() === 0) {
        return '';
    }

    return (getNOI() / getListPriceNumValue()).toFixed(2);
}

function getDebtService() {
    if (getMortgagePrincipal() === 0) {
        return '';
    }

    return (getNOI() / getMortgagePrincipal()).toFixed(2);
}

function formatNumber(num) {
    return num.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,')
}

