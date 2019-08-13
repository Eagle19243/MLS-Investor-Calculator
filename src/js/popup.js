init();

async function init() {
    await preparePopup();
    populateValues(false);
    initEventHandler();
}

function initEventHandler() {
    $('.btn-contact').click(contactToAgent);
    $('.btn-export').click(exportToXLS);
    $('#mls_popup').on('hidden.bs.modal', () => {
        $('.popup-container').css('display', 'none');
    });
    $('#analytics_equity_percent').on('change', onEquityPercentChange);
    $('#analytics_mortgage_interest').on('change', onMortgageInterestChange);
    $('#analytics_mortgage_term').on('change', onMortgageTermChange);
    $('#analytics_vacancy_rate_percent').on('change', onVacancyRatePercentChange);
    $('#analytics_repairs_percent').on('change', onRepairsPercentChange);
    $('#analytics_management_percent').on('change', onManagementPercentChange);
    $('#analytics_list_price').on('change', onListPriceChange);
    chrome.runtime.onMessage.addListener(handleMessage);
}

function handleMessage(request, sender, response) {
    if (request.action === 'Show_Popup') {
        showPopup();
    }
}

function onListPriceChange(evt, newValue) {
    if (isNaN(newValue)) {
        return false;
    } else {
        $(evt.target).html(formatNumberCurrency(Number(newValue)));
        populateValues(true);
    }

    return true;
}

function onUnitChange(evt, newValue) {
    if (isNaN(newValue)) {
        return false;
    } else {
        $(evt.target).html(newValue);
        populateValues(true);
    }

    return true;
}

function onExpensesChange(evt, newValue) {
    if (isNaN(newValue)) {
        return false;
    } else {
        $(evt.target).html(formatNumberCurrency(Number(newValue)));
        populateValues(true);
    }

    return true;
}

function onEquityPercentChange(evt, newValue) {
    const percent = getPercentNum(newValue);

    if (!percent && percent !== 0) {
        return false;
    } else {
        $(evt.target).html(`${percent}%`);
        $('#analytics_mortgage_percent').html(`${100 - percent}%`);
        populateValues(true);
    }

    return true;
}

function onMortgageInterestChange(evt, newValue) {
    const percent = getPercentNum(newValue);

    if (!percent && percent !== 0) {
        return false;
    } else {
        $(evt.target).html(`${percent}%`);
        populateValues(true);
    }

    return true;
}

function onMortgageTermChange(evt, newValue) {
    const value = Number(newValue);

    if (!value && value !== 0) {
        return false;
    } else {
        populateValues(true);
    }

    return true;
}

function onVacancyRatePercentChange(evt, newValue) {
    const percent = getPercentNum(newValue);

    if (!percent && percent !== 0) {
        return false;
    } else {
        $(evt.target).html(`${percent}%`);
        populateValues(true);
    }

    return true;
}

function onRepairsPercentChange(evt, newValue) {
    const percent = getPercentNum(newValue);

    if (!percent && percent !== 0) {
        return false;
    } else {
        $(evt.target).html(`${percent}%`);
        populateValues(true);
    }

    return true;
}

function onManagementPercentChange(evt, newValue) {
    const percent = getPercentNum(newValue);

    if (!percent && percent !== 0) {
        return false;
    } else {
        $(evt.target).html(`${percent}%`);
        populateValues(true);
    }

    return true;
}

function getPercentNum(str) {
    let value = str.replace(/%/g, '');
    
    if (isNaN(value) || (Number(value) > 100 || Number(value) < 0)) {
        return null;
    }

    return parseInt(Number(value));
}

function getValuefromCurrency(str) {
    let value = str.substr(1, str.length - 1).replace(/,/g, '');
    
    if (isNaN(value) || Number(value) <= 0) {
        return 0;
    }

    return parseInt(Number(value));
}

/**
 * return URLType against current url
 * 0 : vow.mlspin.com
 * 1 : stwmls.mlsmatrix.com
 */
function getURLType() {
    const SPINURLRegex      = /vow.mlspin.com\/clients\/report.aspx/i;
    const MATRIXURLRegex    = /stwmls.mlsmatrix.com\/Matrix\/Public\/Portal.aspx/i;
    const url = location.href;

    let type = -1;

    if (SPINURLRegex.test(url)) {
        type = 0;
    } else if (MATRIXURLRegex.test(url)) {
        type = 1;
    }

    return type;
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
                addUnitElements();
                resolve();
            });
        }
    });
}

function addUnitElements() {
    const numberOfUnits = getNumberOfUnits();

    if (numberOfUnits > 0) {
        for (let i = numberOfUnits; i > 0; i--) {
            const unitContent = `<tr> \
                                    <td readonly>#${i}</td> \
                                    <td readonly></td> \
                                    <td class="text-right" id="analytics_unit_${i}_beds" readonly></td> \
                                    <td class="text-right" id="analytics_unit_${i}_baths" readonly></td> \
                                    <td class="text-right" id="analytics_unit_${i}_current" readonly></td> \
                                </tr>`;
            $('#analytics_units').after(unitContent);
        }
    } else {
        $('#analytics_units').remove();
    }

}

function showPopup() {
    $('.popup-container').css('display', 'block');
    $('#mls_popup').modal('show');
}

function populateValues(reCalc) {
    $('.label-analytics').html(getSubjectPropertyAddress());
    $('#analytics_list_price').html(formatNumberCurrency(getListPriceNumValue()));
    $('#analytics_equity').html(formatNumberCurrency(getEquity()));
    $('#analytics_mortgage').html(formatNumberCurrency(getMortgage()));

    $('#analytics_number_of_units').html(getNumberOfUnits());
    $('#analytics_price_per_unit').html(formatNumberCurrency(getPricePerUnit()));
    
    $('#analytics_unit_total_beds').html(getUnitTotalBeds());
    $('#analytics_unit_total_baths').html(getUnitTotalBaths());
    $('#analytics_unit_total_current').html(getUnitTotalCurrent());

    $('#analytics_gross_rents_mls').html(formatNumberCurrencyWithDec(getGrossRentsMLS()));
    $('#analytics_gross_rents_addition').html(formatNumberCurrencyWithDec(getGrossRentsAddition()));
    $('#analytics_vacancy_rate').html(formatNumberCurrencyWithDec(getVacancyRate()));
    $('#analytics_net_rents').html(formatNumberCurrency(getNetRents()));
    
    $('#analytics_heating').html(formatNumberCurrencyWithDec(getHeating()));
    $('#analytics_gas').html(formatNumberCurrencyWithDec(getGas()));
    $('#analytics_electricity').html(formatNumberCurrencyWithDec(getElectricity()));
    $('#analytics_water').html(formatNumberCurrencyWithDec(getWater()));
    $('#analytics_repairs').html(formatNumberCurrencyWithDec(getRepairs()));
    $('#analytics_trash').html(formatNumberCurrencyWithDec(getTrash()));
    $('#analytics_sewer').html(formatNumberCurrencyWithDec(getSewer()));
    $('#analytics_insurance').html(formatNumberCurrencyWithDec(getInsurance()));
    $('#analytics_management').html(formatNumberCurrencyWithDec(getManagement()));
    $('#analytics_miscellaneous').html(formatNumberCurrencyWithDec(getMiscellaneous()));
    $('#analytics_taxes').html(formatNumberCurrencyWithDec(getTaxes()));
    $('#analytics_total').html(formatNumberCurrencyWithDec(getTotalExpenses()));

    $('#analytics_noi').html(formatNumberCurrencyWithDec(getNOI()));
    $('#analytics_mortgage_principal').html(formatNumberCurrencyWithDec(getMortgagePrincipal()));
    $('#analytics_monthly_pi').html(formatNumberCurrencyWithDec(getMonthlyPI()));
    $('#analytics_net_cashflow').html(formatNumberCurrencyWithDec(getNetCashflow()));
    $('#analytics_cash').html(getCash());
    $('#analytics_cap').html(getCAP());
    $('#analytics_debt_service').html(getDebtService());

    if (!reCalc) {
        for (let i = 0; i < getNumberOfUnits(); i++) {
            $(`#analytics_unit_${i + 1}_beds`).html(getBedsForUnit(i));
            $(`#analytics_unit_${i + 1}_baths`).html(getBathsForUnit(i));
            $(`#analytics_unit_${i + 1}_current`).html(getCurrentForUnit(i));
    
            if (getBedsForUnit(i).length === 0 &&
                getBathsForUnit(i).length === 0 &&
                getCurrentForUnit(i).length === 0) {
                
                $(`#analytics_unit_${i + 1}_beds`).addClass('thick-outside-box');
                $(`#analytics_unit_${i + 1}_beds`).removeAttr('readonly');
                $(`#analytics_unit_${i + 1}_baths`).addClass('thick-outside-box');
                $(`#analytics_unit_${i + 1}_baths`).removeAttr('readonly');
                $(`#analytics_unit_${i + 1}_current`).addClass('thick-outside-box');
                $(`#analytics_unit_${i + 1}_current`).removeAttr('readonly');
                $(`#analytics_unit_${i + 1}_beds`).parent().addClass('bg-yellow');

                $(`#analytics_unit_${i + 1}_beds`).on('change', onUnitChange);
                $(`#analytics_unit_${i + 1}_current`).on('change', onUnitChange);
            }
        }

        if (Number(getHeating()) === 0) {
            $('#analytics_heating').addClass('bg-yellow thick-outside-box');
            $('#analytics_heating').removeAttr('readonly');
            $('#analytics_heating').on('change', onExpensesChange);
        }

        if (Number(getGas()) === 0) {
            $('#analytics_gas').addClass('bg-yellow thick-outside-box');
            $('#analytics_gas').removeAttr('readonly');
            $('#analytics_gas').on('change', onExpensesChange);
        }

        if (Number(getElectricity()) === 0) {
            $('#analytics_electricity').addClass('bg-yellow thick-outside-box');
            $('#analytics_electricity').removeAttr('readonly');
            $('#analytics_electricity').on('change', onExpensesChange);
        }

        if (Number(getWater()) === 0) {
            $('#analytics_water').addClass('bg-yellow thick-outside-box');
            $('#analytics_water').removeAttr('readonly');
            $('#analytics_water').on('change', onExpensesChange);
        }

        if (Number(getTrash()) === 0) {
            $('#analytics_trash').addClass('bg-yellow thick-outside-box');
            $('#analytics_trash').removeAttr('readonly');
            $('#analytics_trash').on('change', onExpensesChange);
        }

        if (Number(getSewer()) === 0) {
            $('#analytics_sewer').addClass('bg-yellow thick-outside-box');
            $('#analytics_sewer').removeAttr('readonly');
            $('#analytics_sewer').on('change', onExpensesChange);
        }

        if (Number(getInsurance()) === 0) {
            $('#analytics_insurance').addClass('bg-yellow thick-outside-box');
            $('#analytics_insurance').removeAttr('readonly');
            $('#analytics_insurance').on('change', onExpensesChange);
        }

        if (Number(getMiscellaneous()) === 0) {
            $('#analytics_miscellaneous').addClass('bg-yellow thick-outside-box');
            $('#analytics_miscellaneous').removeAttr('readonly');
            $('#analytics_miscellaneous').on('change', onExpensesChange);
        }

        if (Number(getTaxes()) === 0) {
            $('#analytics_taxes').addClass('bg-yellow thick-outside-box');
            $('#analytics_taxes').removeAttr('readonly');
            $('#analytics_taxes').on('change', onExpensesChange);
        }

        $('#table_analytics').editableTableWidget();

    }
}

function getMLSNumber() {
    let mlsNumber = 0;

    if (getURLType() == 0) {
        const mlsNumberContent = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(3) > table:nth-child(1) > tbody > tr > td > table > tbody > tr:nth-child(1) > td > b').html();
        mlsNumber = /\d+/i.exec(mlsNumberContent)[0].trim();
    } else if (getURLType() == 1) {
        mlsNumber = $('#wrapperTable > div > div > div.row.d-bgcolor--systemLightest.d-marginBottom--8.d-marginTop--6.d-paddingBottom--4 > div:nth-child(1) > div > div:nth-child(3) > span:nth-child(2)').html().trim();
    }
    
    return mlsNumber;
}

function getSubjectPropertyAddress() {
    let address1 = '';
    let address2 = '';

    if (getURLType() == 0) {
        address1 = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(3) > table:nth-child(2) > tbody > tr:nth-child(1) > td:nth-child(1) > b').html().trim();
        address2 = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(3) > table:nth-child(2) > tbody > tr:nth-child(2) > td:nth-child(1) > b').html().trim();
    } else if (getURLType() == 1) {
        address1 = $('#wrapperTable > div > div > div.row.d-bgcolor--systemLightest.d-marginBottom--8.d-marginTop--6.d-paddingBottom--4 > div:nth-child(1) > div > div.d-mega.d-fontSize--mega.d-color--brandDark.col-sm-12 > span').html().trim();
        address2 = $('#wrapperTable > div > div > div.row.d-bgcolor--systemLightest.d-marginBottom--8.d-marginTop--6.d-paddingBottom--4 > div:nth-child(1) > div > div.col-sm-12.d-textSoft > span').html().trim();
    }

    return `${address1}, ${address2}`;
}

function getListPriceNumValue() {
    let listPrice = '';
    let lastValue = $('#analytics_list_price').html();
    
    if (lastValue !== '') {
        return getValuefromCurrency(lastValue);
    }

    if (getURLType() == 0) {
        const listPriceContent = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(3) > table:nth-child(2) > tbody > tr:nth-child(1) > td:nth-child(2) > b').html();
        listPrice = listPriceContent.replace(/(<([^>]+)>)/ig, '').trim();
    } else if (getURLType() == 1) {
        listPrice = $('#wrapperTable > div > div > div.row.d-bgcolor--systemLightest.d-marginBottom--8.d-marginTop--6.d-paddingBottom--4 > div.col-sm-6.d-borderWidthLeft--1.d-borderWidthTop--0.d-borderWidthBottom--0.d-borderWidthRight--0.d-borderStyle--solid.d-bordercolor--systemBase > div:nth-child(1) > div.col-xs-9 > span.d-text.d-fontSize--largest.d-color--brandDark').html().trim();
    }
    
    listPrice = Number(listPrice.substr(1, listPrice.length - 1).replace(/,/g, ''));

    if (!listPrice) {
        return 0;
    }

    return listPrice
}

function getEquity() {
    return getListPriceNumValue() * getPercentNum($('#analytics_equity_percent').html()) / 100;
}

function getMortgage() {
    return getListPriceNumValue() - getEquity();
}

function getHeating() {
    let content = '';
    let lastValue = $('#analytics_heating').html();

    if (lastValue !== '') {
        return getValuefromCurrency(lastValue).toString();
    }

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(1) > td:nth-child(1) > b').html().trim();
    } 

    if (content.length > 0) {
        content = content.substr(1, content.length - 1);
    }

    return content;
}

function getGas() {
    let content = '';
    let lastValue = $('#analytics_gas').html();
    
    if (lastValue !== '') {
        return getValuefromCurrency(lastValue).toString();
    }

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(2) > td:nth-child(1) > b').html().trim();
    }
    
    if (content.length > 0) {
        content = content.substr(1, content.length - 1);
    }

    return content;
}

function getElectricity() {
    let content = '';
    let lastValue = $('#analytics_electricity').html();

    if (lastValue !== '') {
        return getValuefromCurrency(lastValue).toString();
    }

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(3) > td:nth-child(1) > b').html().trim();
    }
    
    if (content.length > 0) {
        content = content.substr(1, content.length - 1);
    }

    return content;
}

function getWater() {
    let content = '';
    let lastValue = $('#analytics_water').html();

    if (lastValue !== '') {
        return getValuefromCurrency(lastValue).toString();
    }

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(4) > td:nth-child(1) > b').html().trim();
    }
    
    if (content.length > 0) {
        content = content.substr(1, content.length - 1);
    }

    return content;
}

function getRepairs() {
    return getNetRents() * getPercentNum($('#analytics_repairs_percent').html()) / 100;
}

function getTrash() {
    let content = '';
    let lastValue = $('#analytics_trash').html();

    if (lastValue !== '') {
        return getValuefromCurrency(lastValue).toString();
    }

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(2) > td:nth-child(2) > b').html().trim();
    }

    if (content.length > 0) {
        content = content.substr(1, content.length - 1);
    }

    return content;
}

function getSewer() {
    let content = '';
    let lastValue = $('#analytics_sewer').html();

    if (lastValue !== '') {
        return getValuefromCurrency(lastValue).toString();
    }

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(3) > td:nth-child(2) > b').html().trim();
    }
    
    if (content.length > 0) {
        content = content.substr(1, content.length - 1);
    }

    return content;
}

function getInsurance() {
    let content = '';
    let lastValue = $('#analytics_insurance').html();

    if (lastValue !== '') {
        return getValuefromCurrency(lastValue).toString();
    }

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(4) > td:nth-child(2) > b').html().trim();
    }
    
    if (content.length > 0) {
        content = content.substr(1, content.length - 1);
    }

    return content;
}

function getManagement() {
    return getNetRents() * getPercentNum($('#analytics_management_percent').html()) / 100;
}

function getMiscellaneous() {
    let content = '';
    let lastValue = $('#analytics_miscellaneous').html();

    if (lastValue !== '') {
        return getValuefromCurrency(lastValue).toString();
    }

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(2) > td:nth-child(3) > b').html().trim();
    }
    
    if (content.length > 0) {
        content = content.substr(1, content.length - 1);
    }

    return content;
}

function getTaxes() {
    let content = '';
    let lastValue = $('#analytics_taxes').html();

    if (lastValue !== '') {
        return getValuefromCurrency(lastValue).toString();
    }

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(10) > tbody > tr > td:nth-child(3) > table:nth-child(4) > tbody > tr:nth-child(3) > td > b:nth-child(1)').html().trim();
    } else if (getURLType() == 1) {
        content = $('#wrapperTable > div > div > div:nth-child(3) > div.col-sm-6.d-bgcolor--systemLightest > div:nth-child(5) > div > div > div > div:nth-child(6) > div > div').find('div:contains("Real Estate Taxes") span:contains("$")').html().trim();
    }
    
    if (content.length > 0) {
        content = content.substr(1, content.length - 1).replace(/,/g, '');
    }

    return content;
}

function getTotalExpenses() {
    const total = Number(getHeating()) + Number(getGas()) + Number(getElectricity()) +
        Number(getWater()) + Number(getRepairs()) + Number(getTrash()) + Number(getSewer()) +
        Number(getInsurance()) + Number(getManagement()) + Number(getMiscellaneous()) + Number(getTaxes());

    return total;
}

function getBedsForUnit(index) {
    let value   = '';
    let content = '';
    let lastValue = $(`#analytics_unit_${index + 1}_beds`).html();

    if (lastValue !== '') {
        return lastValue;
    }

    if (getURLType() == 0) {
        content = $($('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(9)').find('td:contains("Bedrooms") > b')[index]);
    } else if (getURLType() == 1) {
        if (index == 0) {
            content = $('.DisplayRow > tbody > tr > td > table > tbody > tr:nth-child(3) > td:nth-child(4) > span.wrapped-field');
        } else {
            content = $($('.DisplayRow > tbody > tr > td > table > tbody > tr:nth-child(2) > td:nth-child(4) > span.wrapped-field')[index - 1]);
        }
    }
    
    if (content.length > 0) {
        value = content.html().trim();
    }
    
    return value;
}

function getBathsForUnit(index) {
    let value   = '';
    let content = '';
    let lastValue = $(`#analytics_unit_${index + 1}_baths`).html();

    if (lastValue !== '') {
        return lastValue;
    }

    if (getURLType() == 0) {
        content = $($('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(9)').find('td:contains("Baths") > b')[index]);
        if (content.length > 0) {
            value = content.html().trim();
        }
    } else if (getURLType() == 1) {
        let fullBaths = '';
        let halfBaths = '';

        if (index == 0) {
            fullBaths = $('.DisplayRow > tbody > tr > td > table > tbody > tr:nth-child(3) > td:nth-child(6) > span.wrapped-field');
            halfBaths = $('.DisplayRow > tbody > tr > td > table > tbody > tr:nth-child(3) > td:nth-child(8) > span.wrapped-field');
        } else {
            fullBaths = $($('.DisplayRow > tbody > tr > td > table > tbody > tr:nth-child(2) > td:nth-child(6) > span.wrapped-field')[index - 1]);
            halfBaths = $($('.DisplayRow > tbody > tr > td > table > tbody > tr:nth-child(2) > td:nth-child(8) > span.wrapped-field')[index - 1]);
        }

        if (fullBaths.length > 0) {
            value += `${fullBaths.html().trim()}f`;
        }

        if (halfBaths.length > 0) {
            value += ` ${halfBaths.html().trim()}h`;
        }
    }
    
    return value;
}

function getCurrentForUnit(index) {
    let value   = '';
    let content = '';
    let lastValue = $(`#analytics_unit_${index + 1}_current`).html();

    if (lastValue !== '') {
        return lastValue;
    }

    if (getURLType() == 0) {
        content = $($('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(9)').find('td:contains("Rent:") > b')[index]);
    } else if (getURLType() == 1) {
        if (index == 0) {
            content = $('.DisplayRow > tbody > tr > td > table > tbody > tr:nth-child(3) > td:nth-child(11) > span.wrapped-field');
        } else {
            content = $($('.DisplayRow > tbody > tr > td > table > tbody > tr:nth-child(2) > td:nth-child(11) > span.wrapped-field')[index - 1]);
        }
    }
    
    if (content.length > 0) {
        value = content.html().trim();

        if (getURLType() == 1) {
            value = value.substr(1, value.length - 1);
        }
    }
    
    return value;
}

function getUnitTotalBeds() {
    let value = 0;

    for (let i = 0; i < getNumberOfUnits(); i++) {
        value += Number(getBedsForUnit(i));
    }

    return value;
}

function getUnitTotalBaths() {
    let value = 0;

    for (let i = 0; i < getNumberOfUnits(); i++) {
        value += Number(getBathsForUnit(i));
    }

    return value;
}

function getUnitTotalCurrent() {
    let value = 0;

    for (let i = 0; i < getNumberOfUnits(); i++) {
        value += Number(getCurrentForUnit(i).replace(/,/g, ''));
    }

    return value;
}

function getNumberOfUnits() {
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(3) > table:nth-child(2)').find('td:contains("Total Units") > b');
    } else if (getURLType() == 1) {
        content = $('#wrapperTable > div > div > div.row.d-bgcolor--systemLightest.d-marginBottom--8.d-marginTop--6.d-paddingBottom--4 > div.col-sm-6.d-borderWidthLeft--1.d-borderWidthTop--0.d-borderWidthBottom--0.d-borderWidthRight--0.d-borderStyle--solid.d-bordercolor--systemBase > div:nth-child(2) > div.col-xs-12.col-sm-8 > div:nth-child(1) > div > div > div:nth-child(1) > span.d-text.d-paddingRight--5.d-fontWeight--bold');
    }

    if (content.length > 0) {
        value = Number(content.html().trim());
    }
    
    return value;
}

function getPricePerUnit() {
    if (getNumberOfUnits() === 0) {
        return '';
    }

    return getListPriceNumValue() / getNumberOfUnits();
}

function getGrossRentsMLS() {
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(1) > td:nth-child(4) > b').html().trim();
    }
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    } 

    if (value == 0) {
        value = (getUnitTotalCurrent() * 12).toFixed(2);
    }

    return value;
}

function getGrossRentsAddition() {
    return (getUnitTotalCurrent() * 12).toFixed(2);
}

function getVacancyRate() {
    return getGrossRentsAddition() * getPercentNum($('#analytics_vacancy_rate_percent').html()) / 100;
}

function getNetRents() {
    return getGrossRentsAddition() - getVacancyRate();
}

function getNOI() {
    return getNetRents() - getTotalExpenses();
}

function getMortgagePrincipal() {
    return getMonthlyPI() * 12;
}

function getMonthlyPI() {
    const mortgageInterest = getPercentNum($('#analytics_mortgage_interest').html()) / 100;
    const mortgageTerm = Number($('#analytics_mortgage_term').html());
    const tmpValue1 = mortgageInterest / 12 ;
    const tmpValue2 = Math.pow((1 + tmpValue1), mortgageTerm * 12);

    return (tmpValue2 * tmpValue1) / (tmpValue2 - 1) * getMortgage();
}

function getNetCashflow() {
    return getNOI() - getMortgagePrincipal();
}

function getCash() {
    if (getEquity() === 0) {
        return '';
    }

    return (getNetCashflow() / getEquity() * 100).toFixed(2) + '%';
}

function getCAP() {
    if (getListPriceNumValue() === 0) {
        return '';
    }

    return (getNOI() / getListPriceNumValue() * 100).toFixed(2) + '%';
}

function getDebtService() {
    if (getMortgagePrincipal() === 0) {
        return '';
    }

    return (getNOI() / getMortgagePrincipal()).toFixed(3);
}

function formatNumberCurrency(num) {
    return '$' + Math.floor(num).toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
}

function formatNumberCurrencyWithDec(num) {
    return '$' + (Number(num).toFixed(2)).toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
}

function exportToXLS() {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.table_to_sheet(document.getElementById('table_analytics'));
    setFormulas(ws);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, `${getSubjectPropertyAddress()}.xlsx`, {cellStyles: true, sheetStubs: true});
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


