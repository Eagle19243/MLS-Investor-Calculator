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
    $('#mls_popup').on('hidden.bs.modal', () => {
        $('.popup-container').css('display', 'none');
    });
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
                                    <td>#${i}</td> \
                                    <td></td> \
                                    <td class="text-right" id="analytics_unit_${i}_beds"></td> \
                                    <td class="text-right" id="analytics_unit_${i}_baths"></td> \
                                    <td class="text-right" id="analytics_unit_${i}_current"></td> \
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

function exportToXLS() {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.table_to_sheet(document.getElementById('table_analytics'));
    setFormulas(ws);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, `${getMLSNumber()}.xlsx`);
}

function setFormulas(ws) {
    ws['B10'] = {t: 'n', f: 'B1/B9'}; // Price per unit
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
    $('#analytics_heating').html(formatNumber(getHeating()));
    $('#analytics_gas').html(formatNumber(getGas()));
    $('#analytics_electricity').html(formatNumber(getElectricity()));
    $('#analytics_water').html(formatNumber(getWater()));
    $('#analytics_repairs').html(formatNumber(getRepairs()));
    $('#analytics_trash').html(formatNumber(getTrash()));
    $('#analytics_sewer').html(formatNumber(getSewer()));
    $('#analytics_insurance').html(formatNumber(getInsurance()));
    $('#analytics_management').html(formatNumber(getManagement()));
    $('#analytics_miscellaneous').html(formatNumber(getMiscellaneous()));
    $('#analytics_taxes').html(formatNumber(getTaxes()));
    $('#analytics_total').html(formatNumber(getTotalExpenses()));
    $('#analytics_unit_total_beds').html(getUnitTotalBeds());
    $('#analytics_unit_total_baths').html(getUnitTotalBaths());
    $('#analytics_unit_total_current').html(formatNumber(getUnitTotalCurrent()));
    $('#analytics_number_of_units').html(getNumberOfUnits());
    $('#analytics_price_per_unit').html(getPricePerUnit());
    $('#analytics_gross_rents').html(formatNumber(getGrossRents()));
    $('#analytics_vacancy_rate').html(getVacancyRate());
    $('#analytics_net_rents').html(formatNumber(getNetRents()));
    $('#analytics_noi').html(getNOI());
    $('#analytics_mortgage_principal').html(getMortgagePrincipal());
    $('#analytics_net_cashflow').html(getNetCashflow());
    $('#analytics_cash').html(getCash());
    $('#analytics_cap').html(getCAP());
    $('#analytics_debt_service').html(getDebtService());

    for (let i = 0; i < getNumberOfUnits(); i++) {
        $(`#analytics_unit_${i + 1}_beds`).html(getBedsForUnit(i));
        $(`#analytics_unit_${i + 1}_baths`).html(getBathsForUnit(i));
        $(`#analytics_unit_${i + 1}_current`).html(formatNumber(getCurrentForUnit(i)));
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

function getListPrice() {
    let listPrice = '';

    if (getURLType() == 0) {
        const listPriceContent = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(3) > table:nth-child(2) > tbody > tr:nth-child(1) > td:nth-child(2) > b').html();
        listPrice = listPriceContent.replace(/(<([^>]+)>)/ig,"").trim();
    } else if (getURLType() == 1) {
        listPrice = $('#wrapperTable > div > div > div.row.d-bgcolor--systemLightest.d-marginBottom--8.d-marginTop--6.d-paddingBottom--4 > div.col-sm-6.d-borderWidthLeft--1.d-borderWidthTop--0.d-borderWidthBottom--0.d-borderWidthRight--0.d-borderStyle--solid.d-bordercolor--systemBase > div:nth-child(1) > div.col-xs-9 > span.d-text.d-fontSize--largest.d-color--brandDark').html().trim();
    }
    
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
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(1) > td:nth-child(1) > b').html().trim();
    } 

    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getGas() {

    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(2) > td:nth-child(1) > b').html().trim();
    }
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getElectricity() {
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(3) > td:nth-child(1) > b').html().trim();
    }
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getWater() {
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(4) > td:nth-child(1) > b').html().trim();
    }
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getRepairs() {
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(1) > td:nth-child(2) > b').html().trim();
    }
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getTrash() {
    let value   = '';
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(2) > td:nth-child(2) > b').html().trim();
    }

    if (content.length > 0) {
        value = content.substr(1, content.length - 1);
    }

    return (value.length === 0) ? '0.00' : value;
}

function getSewer() {
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(3) > td:nth-child(2) > b').html().trim();
    }
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getInsurance() {
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(4) > td:nth-child(2) > b').html().trim();
    }
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getManagement() {
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(1) > td:nth-child(3) > b').html().trim();
    }
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getMiscellaneous() {
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(2) > td:nth-child(3) > b').html().trim();
    }
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    }

    return value.toFixed(2);
}

function getTaxes() {
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(10) > tbody > tr > td:nth-child(3) > table:nth-child(4) > tbody > tr:nth-child(3) > td > b:nth-child(1)').html().trim();
    } else if (getURLType() == 1) {
        content = $('#wrapperTable > div > div > div:nth-child(3) > div.col-sm-6.d-bgcolor--systemLightest > div:nth-child(5) > div > div > div > div:nth-child(6) > div > div > div:nth-child(3) > div > div.col-xs-7.inherit.col-md-8.col-lg-9.d-paddingTop--2.d-paddingBottom--2.d-borderWidthLeft--0.d-borderWidthRight--0.d-borderStyle--solid.d-bordercolor--systemBase.d-borderWidthTop--1.d-borderWidthBottom--0 > span').html().trim();
    }
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1).replace(',', ''));
    }

    return value.toFixed(2);
}

function getTotalExpenses() {
    const total = Number(getHeating()) + Number(getGas()) + Number(getElectricity()) +
        Number(getWater()) + Number(getRepairs()) + Number(getTrash()) + Number(getSewer()) +
        Number(getInsurance()) + Number(getManagement()) + Number(getMiscellaneous());

    return (total === 0) ? '0.00' : total;
}

function getBedsForUnit(index) {
    let value   = '';
    let content = '';

    if (getURLType() == 0) {
        content = $(`body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(9) > tbody > tr:nth-child(${2 + index * 3}) > td:nth-child(2) > b`);
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
        content = $(`body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(9) > tbody > tr:nth-child(${2 + index * 3}) > td:nth-child(3) > b`);
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
        content = $(`body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(9) > tbody > tr:nth-child(${2 + index * 3}) > td:nth-child(7) > b`);
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
        value += Number(getCurrentForUnit(i).replace(',', ''));
    }

    return value;
}

function getNumberOfUnits() {
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(1) > tbody > tr > td:nth-child(3) > table:nth-child(2) > tbody > tr:nth-child(6) > td:nth-child(1) > b');
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

    return '$' + formatNumber(parseInt(getListPriceNumValue() / getNumberOfUnits()));
}

function getGrossRents() {
    let value = 0;
    let content = '';

    if (getURLType() == 0) {
        content = $('body > center:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(5) > table:nth-child(4) > tbody > tr > td > table:nth-child(7) > tbody > tr:nth-child(1) > td:nth-child(4) > b').html().trim();
    }
    
    if (content.length > 0) {
        value = Number(content.substr(1, content.length - 1));
    } else {
        value = (getUnitTotalCurrent() * 12).toFixed(2);
    }

    return value;
}

function getVacancyRate() {
    const vacancyRatePercent = Number($('#analytics_vacancy_rate_percent').html());
    return (getGrossRents() * vacancyRatePercent).toFixed(2);
}

function getNetRents() {
    return (getGrossRents() - getVacancyRate()).toFixed(2);
}

function getNOI() {
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
    return num.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
}


