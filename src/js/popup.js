init();

async function init() {
    await preparePopup();
    showPopup();
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
}

function onEquityPercentChange(evt, newValue) {
    const percent = getPercentNum(newValue);

    if (!percent) {
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

    if (!percent) {
        return false;
    } else {
        $(evt.target).html(`${percent}%`);
        populateValues(true);
    }

    return true;
}

function onMortgageTermChange(evt, newValue) {
    const value = Number(newValue);

    if (!value) {
        return false;
    } else {
        populateValues(true);
    }

    return true;
}

function onVacancyRatePercentChange(evt, newValue) {
    const percent = getPercentNum(newValue);

    if (!percent) {
        return false;
    } else {
        $(evt.target).html(`${percent}%`);
        populateValues(true);
    }

    return true;
}

function onRepairsPercentChange(evt, newValue) {
    const percent = getPercentNum(newValue);

    if (!percent) {
        return false;
    } else {
        $(evt.target).html(`${percent}%`);
        populateValues(true);
    }

    return true;
}

function onManagementPercentChange(evt, newValue) {
    const percent = getPercentNum(newValue);

    if (!percent) {
        return false;
    } else {
        $(evt.target).html(`${percent}%`);
        populateValues(true);
    }

    return true;
}

function getPercentNum(str) {
    value = str.replace(/%/g, '');
    
    if (isNaN(value)) {
        return null;
    } else if (Number(value) > 100 || Number(value) < 0) {
        return null;
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
                $('#table_analytics').editableTableWidget();
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

    for (let i = 0; i < getNumberOfUnits(); i++) {
        $(`#analytics_unit_${i + 1}_beds`).html(getBedsForUnit(i));
        $(`#analytics_unit_${i + 1}_baths`).html(getBathsForUnit(i));
        $(`#analytics_unit_${i + 1}_current`).html(getCurrentForUnit(i));

        if (getBedsForUnit(i).length === 0 &&
            getBathsForUnit(i).length === 0 &&
            getCurrentForUnit(i).length === 0) {
            
            $(`#analytics_unit_${i + 1}_beds`).addClass('thick-outside-box');
            $(`#analytics_unit_${i + 1}_baths`).addClass('thick-outside-box');
            $(`#analytics_unit_${i + 1}_current`).addClass('thick-outside-box');
            $(`#analytics_unit_${i + 1}_beds`).parent().addClass('bg-yellow');
        }
    }

    if (!reCalc) {
        if (Number(getHeating()) === 0) {
            $('#analytics_heating').addClass('bg-yellow thick-outside-box');
        }

        if (Number(getGas()) === 0) {
            $('#analytics_gas').addClass('bg-yellow thick-outside-box');
        }

        if (Number(getElectricity()) === 0) {
            $('#analytics_electricity').addClass('bg-yellow thick-outside-box');
        }

        if (Number(getWater()) === 0) {
            $('#analytics_water').addClass('bg-yellow thick-outside-box');
        }

        if (Number(getTrash()) === 0) {
            $('#analytics_trash').addClass('bg-yellow thick-outside-box');
        }

        if (Number(getSewer()) === 0) {
            $('#analytics_sewer').addClass('bg-yellow thick-outside-box');
        }

        if (Number(getInsurance()) === 0) {
            $('#analytics_insurance').addClass('bg-yellow thick-outside-box');
        }

        if (Number(getMiscellaneous()) === 0) {
            $('#analytics_miscellaneous').addClass('bg-yellow thick-outside-box');
        }

        if (Number(getTaxes()) === 0) {
            $('#analytics_taxes').addClass('bg-yellow thick-outside-box');
        }

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
    XLSX.writeFile(wb, `${getMLSNumber()}.xlsx`);
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

function setFormulas(ws) {
    const numberOfUnits         = getNumberOfUnits();
    const cListPrice            = 'B1';
    const cEquityValue          = 'B2';
    const cEquityPercent        = 'C2';
    const cMortgageValue        = 'B3';
    const cMortgagePercent      = 'C3';
    const cMortgageInterestRage = 'B6';
    const cMortgageTerm         = 'B7';
    const cNumberOfUnits        = 'B9';
    const cPricePerUnit         = 'B10';
    const cTotalBeds            = `C${14 + numberOfUnits}`;
    const cTotalBaths           = `D${14 + numberOfUnits}`;
    const cTotalCurrent         = `E${14 + numberOfUnits}`;
    const cGrossRentsMLS        = `E${16 + numberOfUnits}`;
    const cGrossRentsAddition   = `E${17 + numberOfUnits}`;
    const cVacancyRatePercent   = `B${18 + numberOfUnits}`;
    const cVacancyRate          = `E${18 + numberOfUnits}`;
    const cNetRents             = `E${19 + numberOfUnits}`;
    const cHeating              = `E${22 + numberOfUnits}`;
    const cGas                  = `E${23 + numberOfUnits}`;
    const cElectricity          = `E${24 + numberOfUnits}`;
    const cWater                = `E${25 + numberOfUnits}`;
    const cRepairPercent        = `B${26 + numberOfUnits}`;
    const cRepair               = `E${26 + numberOfUnits}`;
    const cTrashRemoval         = `E${27 + numberOfUnits}`;
    const cSewer                = `E${28 + numberOfUnits}`;
    const cInsurance            = `E${29 + numberOfUnits}`;
    const cManagementPercent    = `B${30 + numberOfUnits}`;
    const cManagement           = `E${30 + numberOfUnits}`;
    const cMiscellaneous        = `E${31 + numberOfUnits}`;
    const cTaxes                = `E${32 + numberOfUnits}`;
    const cTotalExpenses        = `E${33 + numberOfUnits}`;
    const cNOI                  = `E${35 + numberOfUnits}`;
    const cMortgagePrincipal    = `E${37 + numberOfUnits}`;
    const cMonthlyPI            = `E${38 + numberOfUnits}`;
    const cNetCashflow          = `E${40 + numberOfUnits}`;
    const cCashOnCashReturn     = `E${41 + numberOfUnits}`;
    const cCAP                  = `E${42 + numberOfUnits}`;
    const cDebtService          = `B${44 + numberOfUnits}`;

    const bgPurple              = {
        patternType: 'solid',
        bgColor: { argb: 'FFCE93D8' }
    };
    const bgYellow              = {
        patternType: 'solid',
        bgColor: { argb: 'FFFFFF00' }
    };
    const borderOutsideBox      = {
        top: {style: "thin", color: {auto: 1}},
		right: {style: "thin", color: {auto: 1}},
		bottom: {style: "thin", color: {auto: 1}},
		left: {style: "thin", color: {auto: 1}}
    };
    const borderTop             = {
        top: {style: "thin", color: {auto: 1}}
    };
    const fontBold              = {
        bold: true
    };
    const formatCurrency        = '$#,##0';
    const formatCurrencyWithDec = '$#,##0.00';
    const formatPercent         = '0%';
    const formatPercentWithDec  = '0.00%';

    // List Price
    ws[cListPrice].z       = formatCurrency;

    // Equity
    ws[cEquityValue].z     = formatCurrency;
    ws[cEquityValue].f     = `${cListPrice}*${cEquityPercent}`;
    ws[cEquityPercent].z   = formatPercent;
    ws[cEquityPercent].s   = { fill: bgPurple };

    // Mortgage
    ws[cMortgageValue].z   = formatCurrency;
    ws[cMortgageValue].f   = `${cListPrice}-${cEquityValue}`;
    ws[cMortgagePercent].f = `1-${cEquityPercent}`;
    ws[cMortgagePercent].z = formatPercent;
    ws[cMortgagePercent].s = { fill: bgPurple };

    // Mortgage Interest Rage
    ws[cMortgageInterestRage].z = formatPercent;
    ws[cMortgageInterestRage].s = { fill: bgPurple };

    // Mortgage Term -years
    ws[cMortgageTerm].s = { fill: bgPurple };

    // Price per unit
    ws[cPricePerUnit].z = formatCurrency;
    ws[cPricePerUnit].f = `${cListPrice}/${cNumberOfUnits}`;

    // Units
    for (let i = 0; i < numberOfUnits; i++) {
        const cBeds     = ws[`C${14 + i}`];
        const cBBaths   = ws[`D${14 + i}`];
        const cCurrent  = ws[`E${14 + i}`];

        if (cBeds.v.length === 0 && cBBaths.v.length === 0 && cCurrent.v.length === 0) {
            ws[`A${14 + i}`].s = { fill: bgYellow };
            ws[`B${14 + i}`].s = { fill: bgYellow };
            ws[`C${14 + i}`].s = { fill: bgYellow, border: borderOutsideBox };
            ws[`D${14 + i}`].s = { fill: bgYellow, border: borderOutsideBox };
            ws[`E${14 + i}`].s = { fill: bgYellow, border: borderOutsideBox };
        }
    }

    // Total Beds
    ws[cTotalBeds].f = `SUM(C14:C${14 + numberOfUnits - 1})`;
    ws[cTotalBeds].s = {
        font    : fontBold,
        border  : borderTop
    };
    ws[`A${14 + numberOfUnits}`].s = {
        border: borderTop
    };
    ws[`B${14 + numberOfUnits}`].s = {
        border: borderTop
    };
    // Total Baths
    ws[cTotalBaths].f = `SUM(D14:D${14 + numberOfUnits - 1})`;
    ws[cTotalBaths].s = {
        font    : fontBold,
        border  : borderTop
    };
    // Total Current
    ws[cTotalCurrent].f = `SUM(E14:E${14 + numberOfUnits - 1})`;
    ws[cTotalCurrent].s = {
        font    : fontBold,
        border  : borderTop
    };

    // Gross Rents from MLS
    ws[cGrossRentsMLS].z = formatCurrencyWithDec;

    // Grosss Rents Addition
    ws[cGrossRentsAddition].z = formatCurrencyWithDec;
    ws[cGrossRentsAddition].f = `${cTotalCurrent}*12`;
    
    // Vacancy rate
    ws[cVacancyRate].z = formatCurrencyWithDec;
    ws[cVacancyRate].f = `${cGrossRentsAddition}*${cVacancyRatePercent}`;
    ws[cVacancyRatePercent].z = formatPercent;
    ws[cVacancyRatePercent].s = { fill: bgPurple };
    
    // Net Rents
    ws[cNetRents].z = formatCurrencyWithDec;
    ws[cNetRents].f = `${cGrossRentsAddition}-${cVacancyRate}`;
    ws[cNetRents].s = {
        font    : fontBold,
        border  : borderTop
    };
    ws[`A${19 + numberOfUnits}`].s = {
        font: fontBold,
        border: borderTop
    };
    ws[`B${19 + numberOfUnits}`].s = {
        border: borderTop
    };
    ws[`C${19 + numberOfUnits}`].s = {
        border: borderTop
    };
    ws[`D${19 + numberOfUnits}`].s = {
        border: borderTop
    };

    // Heating
    ws[cHeating].z = formatCurrencyWithDec;
    if (ws[cHeating].v === '$0.00') {
        ws[cHeating].s = { fill: bgYellow, border: borderOutsideBox };
    }

    // Gas
    ws[cGas].z = formatCurrencyWithDec;
    if (ws[cGas].v === '$0.00') {
        ws[cGas].s = { fill: bgYellow, border: borderOutsideBox };
    }

    // Electricity
    ws[cElectricity].z = formatCurrencyWithDec;
    if (ws[cElectricity].v === '$0.00') {
        ws[cElectricity].s = { fill: bgYellow, border: borderOutsideBox };
    }

    // Water
    ws[cWater].z = formatCurrencyWithDec;
    if (ws[cWater].v === '$0.00') {
        ws[cWater].s = { fill: bgYellow, border: borderOutsideBox };
    }

    // Repairs & Maintenance
    ws[cRepair].z = formatCurrencyWithDec;
    ws[cRepair].f = `${cRepairPercent}*${cNetRents}`;
    ws[cRepairPercent].z = formatPercent;
    ws[cRepairPercent].s = { fill: bgPurple };

    // Trash Removal
    ws[cTrashRemoval].z = formatCurrencyWithDec;
    if (ws[cTrashRemoval].v === '$0.00') {
        ws[cTrashRemoval].s = { fill: bgYellow, border: borderOutsideBox };
    }

    // Sewer
    ws[cSewer].z = formatCurrencyWithDec;
    if (ws[cSewer].v === '$0.00') {
        ws[cSewer].s = { fill: bgYellow, border: borderOutsideBox };
    }

    // Insurance
    ws[cInsurance].z = formatCurrencyWithDec;
    if (ws[cInsurance].v === '$0.00') {
        ws[cInsurance].s = { fill: bgYellow, border: borderOutsideBox };
    }

    // Repairs & Maintenance
    ws[cManagement].z = formatCurrencyWithDec;
    ws[cManagement].f = `${cManagementPercent}*${cNetRents}`;
    ws[cManagementPercent].z = formatPercent;
    ws[cManagementPercent].s = { fill: bgPurple };

    // Miscellaneous
    ws[cMiscellaneous].z = formatCurrencyWithDec;
    if (ws[cMiscellaneous].v === '$0.00') {
        ws[cMiscellaneous].s = { fill: bgYellow, border: borderOutsideBox };
    }

    // Taxes
    ws[cTaxes].z = formatCurrencyWithDec;
    if (ws[cTaxes].v === '$0.00') {
        ws[cTaxes].s = { fill: bgYellow, border: borderOutsideBox };
    }
    
    // Total Operating Expenses
    ws[cTotalExpenses].z = formatCurrencyWithDec;
    ws[cTotalExpenses].f = `SUM(${cHeating}:${cTaxes})`;
    ws[cTotalExpenses].s = {
        font: fontBold,
        border: borderTop
    };
    ws[`A${33 + numberOfUnits}`].s = {
        font: fontBold,
        border: borderTop
    };
    ws[`B${33 + numberOfUnits}`].s = {
        border: borderTop
    };
    ws[`C${33 + numberOfUnits}`].s = {
        border: borderTop
    };
    ws[`D${33 + numberOfUnits}`].s = {
        border: borderTop
    };

    // Net Operating Income(NOI)
    ws[cNOI].z = formatCurrencyWithDec;
    ws[cNOI].f = `${cNetRents}-${cTotalExpenses}`;
    ws[cNOI].s = {
        font: fontBold
    };
    ws[`A${35 + numberOfUnits}`].s = {
        font: fontBold
    };

    // Mortgage Principal & Interest
    ws[cMortgagePrincipal].z = formatCurrencyWithDec;
    ws[cMortgagePrincipal].f = `${cMonthlyPI}*12`;
    ws[cMortgagePrincipal].s = {
        font: fontBold
    };
    ws[`A${37 + numberOfUnits}`].s = {
        font: fontBold
    };

    // Monthly PI Payment
    ws[cMonthlyPI].z = formatCurrencyWithDec;
    ws[cMonthlyPI].f = `((((1+(${cMortgageInterestRage}/12))^(${cMortgageTerm}*12))*(${cMortgageInterestRage}/12))/(((1+(${cMortgageInterestRage}/12))^(${cMortgageTerm}*12))-1))*${cMortgageValue}`;
    ws[cMonthlyPI].s = {
        font: fontBold
    };
    ws[`A${38 + numberOfUnits}`].s = {
        font: fontBold
    };

    // Net Cashflow/ROI
    ws[cNetCashflow].z = formatCurrencyWithDec;
    ws[cNetCashflow].f = `${cNOI}-${cMortgagePrincipal}`;
    ws[cNetCashflow].s = {
        font: fontBold
    };
    ws[`A${40 + numberOfUnits}`].s = {
        font: fontBold
    };

    // Cash on Cash Return
    ws[cCashOnCashReturn].z = formatPercentWithDec;
    ws[cCashOnCashReturn].f = `${cNetCashflow}/${cEquityValue}`;
    ws[cCashOnCashReturn].s = {
        font: fontBold
    };
    ws[`A${41 + numberOfUnits}`].s = {
        font: fontBold
    };

    // CAP
    ws[cCAP].z = formatPercentWithDec;
    ws[cCAP].f = `${cNOI}/${cListPrice}`;
    ws[cCAP].s = {
        font: fontBold
    };
    ws[`A${42 + numberOfUnits}`].s = {
        font: fontBold
    };

    // Debt Service Cover Ratio
    ws[cDebtService].z = '0.000';
    ws[cDebtService].f = `${cNOI}/${cMortgagePrincipal}`;
    ws[cDebtService].s = {
        font: fontBold
    };
    ws[`A${44 + numberOfUnits}`].s = {
        font: fontBold
    };
}


