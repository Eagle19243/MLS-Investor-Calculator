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
    ws[cEquityPercent].s   = { fgColor: { rgb: 'CE93D8' } };

    // Mortgage
    ws[cMortgageValue].z   = formatCurrency;
    ws[cMortgageValue].f   = `${cListPrice}-${cEquityValue}`;
    ws[cMortgagePercent].f = `1-${cEquityPercent}`;
    ws[cMortgagePercent].z = formatPercent;
    ws[cMortgagePercent].s = { fgColor: { rgb: 'CE93D8' } };

    // Mortgage Interest Rage
    ws[cMortgageInterestRage].z = formatPercent;
    ws[cMortgageInterestRage].s = { fgColor: { rgb: 'CE93D8' } };

    // Mortgage Term -years
    ws[cMortgageTerm].s = { fgColor: { rgb: 'CE93D8' } };

    // Price per unit
    ws[cPricePerUnit].z = formatCurrency;
    ws[cPricePerUnit].f = `${cListPrice}/${cNumberOfUnits}`;

    // Units
    for (let i = 0; i < numberOfUnits; i++) {
        const cBeds     = ws[`C${14 + i}`];
        const cBBaths   = ws[`D${14 + i}`];
        const cCurrent  = ws[`E${14 + i}`];

        if (cBeds.v.length === 0 && cBBaths.v.length === 0 && cCurrent.v.length === 0) {
            ws[`A${14 + i}`].s = { fgColor: { rgb: 'FFFF00' } };
            ws[`B${14 + i}`].s = { fgColor: { rgb: 'FFFF00' } };
            ws[`C${14 + i}`].s = {
                fgColor: { rgb: 'FFFF00' },
                top: {style: "medium", color: {rgb: 0x000000}},
                right: {style: "medium", color: {rgb: 0x000000}},
                bottom: {style: "medium", color: {rgb: 0x000000}},
                left: {style: "medium", color: {rgb: 0x000000}}
            };
            ws[`D${14 + i}`].s = {
                fgColor: { rgb: 'FFFF00' },
                top: {style: "medium", color: {rgb: 0x000000}},
                right: {style: "medium", color: {rgb: 0x000000}},
                bottom: {style: "medium", color: {rgb: 0x000000}},
                left: {style: "medium", color: {rgb: 0x000000}}
            };
            ws[`E${14 + i}`].s = {
                fgColor: { rgb: 'FFFF00' },
                top: {style: "medium", color: {rgb: 0x000000}},
                right: {style: "medium", color: {rgb: 0x000000}},
                bottom: {style: "medium", color: {rgb: 0x000000}},
                left: {style: "medium", color: {rgb: 0x000000}}
            };
        }
    }

    // Total Beds
    ws[cTotalBeds].f = `SUM(C14:C${14 + numberOfUnits - 1})`;
    ws[cTotalBeds].s = { bold: true, top: {style: "medium", color: {rgb: 0x000000}} };
    ws[`A${14 + numberOfUnits}`].s = { top: {style: "medium", color: {rgb: 0x000000}} };
    ws[`B${14 + numberOfUnits}`].s = { top: {style: "medium", color: {rgb: 0x000000}} };
    // Total Baths
    ws[cTotalBaths].f = `SUM(D14:D${14 + numberOfUnits - 1})`;
    ws[cTotalBaths].s = { bold: true, top: {style: "medium", color: {rgb: 0x000000}} };
    // Total Current
    ws[cTotalCurrent].f = `SUM(E14:E${14 + numberOfUnits - 1})`;
    ws[cTotalCurrent].s = { bold: true, top: {style: "medium", color: {rgb: 0x000000}} };

    // Gross Rents from MLS
    ws[cGrossRentsMLS].z = formatCurrencyWithDec;

    // Grosss Rents Addition
    ws[cGrossRentsAddition].z = formatCurrencyWithDec;
    ws[cGrossRentsAddition].f = `${cTotalCurrent}*12`;
    
    // Vacancy rate
    ws[cVacancyRate].z = formatCurrencyWithDec;
    ws[cVacancyRate].f = `${cGrossRentsAddition}*${cVacancyRatePercent}`;
    ws[cVacancyRatePercent].z = formatPercent;
    ws[cVacancyRatePercent].s = { fgColor: { rgb: 'CE93D8' } };
    
    // Net Rents
    ws[cNetRents].z = formatCurrencyWithDec;
    ws[cNetRents].f = `${cGrossRentsAddition}-${cVacancyRate}`;
    ws[cNetRents].s = { bold: true, top: {style: "medium", color: {rgb: 0x000000}} };
    ws[`A${19 + numberOfUnits}`].s = { bold: true, top: {style: "medium", color: {rgb: 0x000000}} };
    ws[`B${19 + numberOfUnits}`].s = { top: {style: "medium", color: {rgb: 0x000000}} };
    ws[`C${19 + numberOfUnits}`].s = { top: {style: "medium", color: {rgb: 0x000000}} };
    ws[`D${19 + numberOfUnits}`].s = { top: {style: "medium", color: {rgb: 0x000000}} };

    // Heating
    ws[cHeating].z = formatCurrencyWithDec;
    if (ws[cHeating].v === 0) {
        ws[cHeating].s = {
            fgColor: { rgb: 'FFFF00' },
            top: {style: "medium", color: {rgb: 0x000000}},
            right: {style: "medium", color: {rgb: 0x000000}},
            bottom: {style: "medium", color: {rgb: 0x000000}},
            left: {style: "medium", color: {rgb: 0x000000}}
        };
    }

    // Gas
    ws[cGas].z = formatCurrencyWithDec;
    if (ws[cGas].v === 0) {
        ws[cGas].s = {
            fgColor: { rgb: 'FFFF00' },
            top: {style: "medium", color: {rgb: 0x000000}},
            right: {style: "medium", color: {rgb: 0x000000}},
            bottom: {style: "medium", color: {rgb: 0x000000}},
            left: {style: "medium", color: {rgb: 0x000000}}
        };
    }

    // Electricity
    ws[cElectricity].z = formatCurrencyWithDec;
    if (ws[cElectricity].v === 0) {
        ws[cElectricity].s = {
            fgColor: { rgb: 'FFFF00' },
            top: {style: "medium", color: {rgb: 0x000000}},
            right: {style: "medium", color: {rgb: 0x000000}},
            bottom: {style: "medium", color: {rgb: 0x000000}},
            left: {style: "medium", color: {rgb: 0x000000}}
        };
    }

    // Water
    ws[cWater].z = formatCurrencyWithDec;
    if (ws[cWater].v === 0) {
        ws[cWater].s = {
            fgColor: { rgb: 'FFFF00' },
            top: {style: "medium", color: {rgb: 0x000000}},
            right: {style: "medium", color: {rgb: 0x000000}},
            bottom: {style: "medium", color: {rgb: 0x000000}},
            left: {style: "medium", color: {rgb: 0x000000}}
        };
    }

    // Repairs & Maintenance
    ws[cRepair].z = formatCurrencyWithDec;
    ws[cRepair].f = `${cRepairPercent}*${cNetRents}`;
    ws[cRepairPercent].z = formatPercent;
    ws[cRepairPercent].s = { fgColor: { rgb: 'CE93D8' } };

    // Trash Removal
    ws[cTrashRemoval].z = formatCurrencyWithDec;
    if (ws[cTrashRemoval].v === 0) {
        ws[cTrashRemoval].s = {
            fgColor: { rgb: 'FFFF00' },
            top: {style: "medium", color: {rgb: 0x000000}},
            right: {style: "medium", color: {rgb: 0x000000}},
            bottom: {style: "medium", color: {rgb: 0x000000}},
            left: {style: "medium", color: {rgb: 0x000000}}
        };
    }

    // Sewer
    ws[cSewer].z = formatCurrencyWithDec;
    if (ws[cSewer].v === 0) {
        ws[cSewer].s = {
            fgColor: { rgb: 'FFFF00' },
            top: {style: "medium", color: {rgb: 0x000000}},
            right: {style: "medium", color: {rgb: 0x000000}},
            bottom: {style: "medium", color: {rgb: 0x000000}},
            left: {style: "medium", color: {rgb: 0x000000}}
        };
    }

    // Insurance
    ws[cInsurance].z = formatCurrencyWithDec;
    if (ws[cInsurance].v === 0) {
        ws[cInsurance].s = {
            fgColor: { rgb: 'FFFF00' },
            top: {style: "medium", color: {rgb: 0x000000}},
            right: {style: "medium", color: {rgb: 0x000000}},
            bottom: {style: "medium", color: {rgb: 0x000000}},
            left: {style: "medium", color: {rgb: 0x000000}}
        };
    }

    // Repairs & Maintenance
    ws[cManagement].z = formatCurrencyWithDec;
    ws[cManagement].f = `${cManagementPercent}*${cNetRents}`;
    ws[cManagementPercent].z = formatPercent;
    ws[cManagementPercent].s = { fgColor: { rgb: 'CE93D8' } };

    // Miscellaneous
    ws[cMiscellaneous].z = formatCurrencyWithDec;
    if (ws[cMiscellaneous].v === 0) {
        ws[cMiscellaneous].s = {
            fgColor: { rgb: 'FFFF00' },
            top: {style: "medium", color: {rgb: 0x000000}},
            right: {style: "medium", color: {rgb: 0x000000}},
            bottom: {style: "medium", color: {rgb: 0x000000}},
            left: {style: "medium", color: {rgb: 0x000000}}
        };
    }

    // Taxes
    ws[cTaxes].z = formatCurrencyWithDec;
    if (ws[cTaxes].v === 0) {
        ws[cTaxes].s = {
            fgColor: { rgb: 'FFFF00' },
            top: {style: "medium", color: {rgb: 0x000000}},
            right: {style: "medium", color: {rgb: 0x000000}},
            bottom: {style: "medium", color: {rgb: 0x000000}},
            left: {style: "medium", color: {rgb: 0x000000}}
        };
    }
    
    // Total Operating Expenses
    ws[cTotalExpenses].z = formatCurrencyWithDec;
    ws[cTotalExpenses].f = `SUM(${cHeating}:${cTaxes})`;
    ws[cTotalExpenses].s = { bold: true, top: {style: "medium", color: {rgb: 0x000000}} };
    ws[`A${33 + numberOfUnits}`].s = { bold: true, top: {style: "medium", color: {rgb: 0x000000}} };
    ws[`B${33 + numberOfUnits}`].s = { top: {style: "medium", color: {rgb: 0x000000}} };
    ws[`C${33 + numberOfUnits}`].s = { top: {style: "medium", color: {rgb: 0x000000}} };
    ws[`D${33 + numberOfUnits}`].s = { top: {style: "medium", color: {rgb: 0x000000}} };

    // Net Operating Income(NOI)
    ws[cNOI].z = formatCurrencyWithDec;
    ws[cNOI].f = `${cNetRents}-${cTotalExpenses}`;
    ws[cNOI].s = { bold: true };
    ws[`A${35 + numberOfUnits}`].s = { bold: true };

    // Mortgage Principal & Interest
    ws[cMortgagePrincipal].z = formatCurrencyWithDec;
    ws[cMortgagePrincipal].f = `${cMonthlyPI}*12`;
    ws[cMortgagePrincipal].s = { bold: true };
    ws[`A${37 + numberOfUnits}`].s = { bold: true };

    // Monthly PI Payment
    ws[cMonthlyPI].z = formatCurrencyWithDec;
    ws[cMonthlyPI].f = `((((1+(${cMortgageInterestRage}/12))^(${cMortgageTerm}*12))*(${cMortgageInterestRage}/12))/(((1+(${cMortgageInterestRage}/12))^(${cMortgageTerm}*12))-1))*${cMortgageValue}`;
    ws[cMonthlyPI].s = { bold: true };
    ws[`A${38 + numberOfUnits}`].s = { bold: true };

    // Net Cashflow/ROI
    ws[cNetCashflow].z = formatCurrencyWithDec;
    ws[cNetCashflow].f = `${cNOI}-${cMortgagePrincipal}`;
    ws[cNetCashflow].s = { bold: true };
    ws[`A${40 + numberOfUnits}`].s = { bold: true };

    // Cash on Cash Return
    ws[cCashOnCashReturn].z = formatPercentWithDec;
    ws[cCashOnCashReturn].f = `${cNetCashflow}/${cEquityValue}`;
    ws[cCashOnCashReturn].s = { bold: true };
    ws[`A${41 + numberOfUnits}`].s = { bold: true };

    // CAP
    ws[cCAP].z = formatPercentWithDec;
    ws[cCAP].f = `${cNOI}/${cListPrice}`;
    ws[cCAP].s = { bold: true };
    ws[`A${42 + numberOfUnits}`].s = { bold: true };

    // Debt Service Cover Ratio
    ws[cDebtService].z = '0.000';
    ws[cDebtService].f = `${cNOI}/${cMortgagePrincipal}`;
    ws[cDebtService].s = { bold: true };
    ws[`A${44 + numberOfUnits}`].s = { bold: true };
}