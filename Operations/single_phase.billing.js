const dbConf = require('../Configuration/dbConfig');
var Excel = require("exceljs");// load exceljs module
const fs = require("fs");
const logger = require("../Configuration/logger");
const Billing_Model = require('../ViewModels/Billing_Model');
const prepSlab = require("../ViewModels/Prep_Slab");
const billGen = require("../ViewModels/Bill_Generate");
const sourcePath = "D:\\Env_Temp_std\\";
const _common = require("./common.service");
module.exports = {
    async GetBillingDetails(from, to) {
        var _query = "select * from envisage_dev.mtrinstallconfig mic join envisage_dev.mstmeter m on mic.micmtrsrno = m.mtrsrno join envisage_dev.mstmtrmodel md on md.mdlid = m.mtrmodelid join envisage_dev.billinglog bl on blmicmtrmdl.blmetersrno = mic.micmtrsrno where blmicmtrmdl.bltimestamp >= " + from + " and blmicmtrmdl.bltimestamp <=" + to;
        let result;
        //logger.info("Query 1 : " + _query);
        dbConf.pool.open(dbConf.connStr, (err, conn) => {
            if (err) {
                console.log(err);
            }
            result = dbConf.runSQL(conn, _query);
            conn.close();

            if (result != undefined) {
                var fromDate = _common.epochToJsDate(from);
                var toDate = _common.epochToJsDate(to);
                var month = fromDate.getMonth() + 1;
                var year = fromDate.getFullYear();
                BillingProcess(result, month, year, from, to);
            }
        });
    }
}

async function BillingProcess(blmicmtrmdl, month, year, from, to) {
    try {
        for (var i = 0; i < blmicmtrmdl.length; i++) {
            var _query = "select * from envisage_dev.cmrattrib cmra join envisage_dev.ccattouslottrfdetail ccd on ccd.tstdconclassid = cmra.cmracurclassid and ccd.tstdconcatid=cmra.cmracurcategoryid";
            _query += " join envisage_dev.prepmsttariff pm on ccd.tstdtariffid = pm.preptrfid where cmra.cmraconsumerid ='" + blmicmtrmdl[i].micconsumerid + "'";
            let cmrcctoutrfmst;
            dbConf.pool.open(dbConf.connStr, (err, conn) => {
                if (err) {
                    console.log(err);
                }
                cmrcctoutrfmst = dbConf.runSQL(conn, _query);
                conn.close();

                if (cmrcctoutrfmst != undefined) {
                    BillingProcess(blmicmtrmdl[i], cmrcctoutrfmst, month, year);
                }
            });
        }
    }
    catch (err) {
        logger.info("Error : " + err);
        console.log(err);
    }
}

async function BillingProcess(blmicmtrmdl, cmrcctoutrfmst, month, year, from, to) {
    try {
        var _folderPath = sourcePath + month + "_" + year + "//";
        fs.access(_folderPath, (error) => {
            if (error) {
                fs.mkdir(_folderPath, (error) => {
                    if (error) {
                        console.log(error);
                    } else {
                        var _trfFolder = _folderPath + cmrcctoutrfmst[0].preptrfid + "//";
                        fs.mkdir(_trfFolder, (err) => {
                            if (err) console.log(err);
                            else {
                                console.log("New Directory created successfully !!" + path);
                                ProcessSPBill(blmicmtrmdl, cmrcctoutrfmst, month, year);
                            }
                        });
                    }
                });
            } else {
                console.log("Given Directory already exists !!");
                ProcessSPBill(blmicmtrmdl, cmrcctoutrfmst, month, year);
            }
        });
    }
    catch (err) {
        logger.info("Error : " + err);
        console.log(err);
    }
}

async function ProcessSPBill(blmicmtrmdl, cmrcctoutrfmst, month, year, from, to) {
    try {
        var d = new Date();
        var _query = "Select * from mtrinstant_sp_log m where m.misplsrno = '" + blmicmtrmdl.micmtrsrno + "' and (m.mispldate >= " + from + " and m.mispldate <= " + to + ") order by 1 desc limit 1;";
        let _mtrinstsp;
        let _prepaydt;
        dbConf.pool.open(dbConf.connStr, (err, conn) => {
            if (err) {
                console.log(err);
            }
            _mtrinstsp = dbConf.runSQL(conn, _query);
            _query = "select * from prepaydtlinst_sp p where p.pdispmtrsrno = '" + blmicmtrmdl.micmtrsrno + "' and (p.pdispdate >= " + from + " and p.pdispdate <= " + to + ") order by 1 desc limit 1;"
            //logger.info("Query 16 : " + _query);
            _prepaydt = dbConf.runSQL(conn, _query);
            conn.close();

            if (_mtrinstsp != undefined) {
                var _spBill = Billing_Model.SPBill_Model;
                _spBill.RTC_absolute = d.toDateString();
                _spBill.Active_Power = _mtrinstsp[0].misplmdkw == undefined ? 0 : _mtrinstsp[0].misplmdkw;
                _spBill.Apparent_Power = _mtrinstsp[0].misplmdkva == undefined ? 0 : _mtrinstsp[0].misplmdkva;
                _spBill.Instantaneous_Frequency = _mtrinstsp[0].misplfreq == undefined ? 0 : _mtrinstsp[0].misplfreq;
                _spBill.Signed_Reactive_power = _mtrinstsp[0].misplkvar == undefined ? 0 : _mtrinstsp[0].misplkvar;
                _spBill.Signed_Three_phase_Power_factor = _mtrinstsp[0].misplpf == undefined ? 0 : _mtrinstsp[0].misplpf;
                _spBill.Cumulative_Energy_KVAh = blmicmtrmdl.blcumkvah == undefined ? 0 : blmicmtrmdl.blcumkvah;
                _spBill.Cumulative_Energy_KVARh_Lag = blmicmtrmdl.blcumkvarh_lag == undefined ? 0 : blmicmtrmdl.blcumkvarh_lag;
                _spBill.Cumulative_Energy_KVARh_Lead = blmicmtrmdl.blcumkvarh_lead == undefined ? 0 : blmicmtrmdl.blcumkvarh_lead;
                _spBill.Maximum_demand_KW = blmicmtrmdl.blmdkw == undefined ? 0 : blmicmtrmdl.blmdkw;
                _spBill.Maximum_demand_KVA = blmicmtrmdl.blmdkva == undefined ? 0 : blmicmtrmdl.blmdkva;
                _spBill.Cumulative_tamper_count = blmicmtrmdl.blcumtampercount == undefined ? 0 : blmicmtrmdl.blcumtampercount;
                _spBill.Billing_Date = blmicmtrmdl.bltimestamp == undefined ? 0 : blmicmtrmdl.bltimestamp;
                _spBill.Signed_Average_PowerFactor = blmicmtrmdl.blavgpf == undefined ? 0.00 : blmicmtrmdl.blavgpf;
                _spBill.Last_Month_Cumulative_EnergykWh = blmicmtrmdl.blcumkwh == undefined ? 0.00 : blmicmtrmdl.blcumkwh;
                _spBill.LCumulative_Energy_for_TOU1 = blmicmtrmdl.blcumkwhtou1 == undefined ? 0.00 : blmicmtrmdl.blcumkwhtou1;
                _spBill.LCumulative_Energy_for_TOU2 = blmicmtrmdl.blcumkwhtou2 == undefined ? 0.00 : blmicmtrmdl.blcumkwhtou2;
                _spBill.LCumulative_Energy_for_TOU3 = blmicmtrmdl.blcumkwhtou3 == undefined ? 0.00 : blmicmtrmdl.blcumkwhtou3;
                _spBill.LCumulative_Energy_for_TOU4 = blmicmtrmdl.blcumkwhtou4 == undefined ? 0.00 : blmicmtrmdl.blcumkwhtou4;
                _spBill.LCumulative_Energy_for_TOU5 = blmicmtrmdl.blcumkwhtou5 == undefined ? 0.00 : blmicmtrmdl.blcumkwhtou5;
                _spBill.LCumulative_Energy_for_TOU6 = blmicmtrmdl.blcumkwhtou6 == undefined ? 0.00 : blmicmtrmdl.blcumkwhtou6;
                _spBill.LCumulative_Energy_for_TOU7 = blmicmtrmdl.blcumkwhtou7 == undefined ? 0.00 : blmicmtrmdl.blcumkwhtou7;
                _spBill.LCumulative_Energy_for_TOU8 = blmicmtrmdl.blcumkwhtou8 == undefined ? 0.00 : blmicmtrmdl.blcumkwhtou8;
                _spBill.LCumulative_Energy_kVARh_Lag = blmicmtrmdl.blcumkvarh_lag == undefined ? 0.00 : blmicmtrmdl.blcumkvarh_lag;
                _spBill.LCumulative_Energy_kVARh_Lead = blmicmtrmdl.blcumkvarh_lead == undefined ? 0.00 : blmicmtrmdl.blcumkvarh_lead;
                _spBill.LCumulative_Energy_kVAh = blmicmtrmdl.blcumkvah == undefined ? 0.00 : blmicmtrmdl.blcumkvah;
                _spBill.LCumulative_Apparent_Energy_for_TOU1 = blmicmtrmdl.blcumkvahtou1 == undefined ? 0.00 : blmicmtrmdl.blcumkvahtou1;
                _spBill.LCumulative_Apparent_Energy_for_TOU2 = blmicmtrmdl.blcumkvahtou2 == undefined ? 0.00 : blmicmtrmdl.blcumkvahtou2;
                _spBill.LCumulative_Apparent_Energy_for_TOU3 = blmicmtrmdl.blcumkvahtou3 == undefined ? 0.00 : blmicmtrmdl.blcumkvahtou3;
                _spBill.LCumulative_Apparent_Energy_for_TOU4 = blmicmtrmdl.blcumkvahtou4 == undefined ? 0.00 : blmicmtrmdl.blcumkvahtou4;
                _spBill.LCumulative_Apparent_Energy_for_TOU5 = blmicmtrmdl.blcumkvahtou5 == undefined ? 0.00 : blmicmtrmdl.blcumkvahtou5;
                _spBill.LCumulative_Apparent_Energy_for_TOU6 = blmicmtrmdl.blcumkvahtou6 == undefined ? 0.00 : blmicmtrmdl.blcumkvahtou6;
                _spBill.LCumulative_Apparent_Energy_for_TOU7 = blmicmtrmdl.blcumkvahtou7 == undefined ? 0.00 : blmicmtrmdl.blcumkvahtou7;
                _spBill.LCumulative_Apparent_Energy_for_TOU8 = blmicmtrmdl.blcumkvahtou8 == undefined ? 0.00 : blmicmtrmdl.blcumkvahtou8;
                _spBill.LMDMaximum_Demand_kW = blmicmtrmdl.blmdkw == undefined ? 0.00 : blmicmtrmdl.blmdkw;
                _spBill.LMD_kW_for_TOU1 = blmicmtrmdl.blmdkwtou1 == undefined ? 0.00 : blmicmtrmdl.blmdkwtou1;
                _spBill.LMD_kW_for_TOU2 = blmicmtrmdl.blmdkwtou2 == undefined ? 0.00 : blmicmtrmdl.blmdkwtou2;
                _spBill.LMD_kW_for_TOU3 = blmicmtrmdl.blmdkwtou3 == undefined ? 0.00 : blmicmtrmdl.blmdkwtou3;
                _spBill.LMD_kW_for_TOU4 = blmicmtrmdl.blmdkwtou4 == undefined ? 0.00 : blmicmtrmdl.blmdkwtou4;
                _spBill.LMD_kW_for_TOU5 = blmicmtrmdl.blmdkwtou5 == undefined ? 0.00 : blmicmtrmdl.blmdkwtou5;
                _spBill.LMD_kW_for_TOU6 = blmicmtrmdl.blmdkwtou6 == undefined ? 0.00 : blmicmtrmdl.blmdkwtou6;
                _spBill.LMD_kW_for_TOU7 = blmicmtrmdl.blmdkwtou7 == undefined ? 0.00 : blmicmtrmdl.blmdkwtou7;
                _spBill.LMD_kW_for_TOU8 = blmicmtrmdl.blmdkwtou8 == undefined ? 0.00 : blmicmtrmdl.blmdkwtou8;
                _spBill.LMDMaximum_Demand_kVA = blmicmtrmdl.blmdkva == undefined ? 0.00 : blmicmtrmdl.blmdkva;
                _spBill.LMD_kVA_for_TOU1 = blmicmtrmdl.blmdkvatou1 == undefined ? 0.00 : blmicmtrmdl.blmdkvatou1;
                _spBill.LMD_kVA_for_TOU2 = blmicmtrmdl.blmdkvatou2 == undefined ? 0.00 : blmicmtrmdl.blmdkvatou2;
                _spBill.LMD_kVA_for_TOU3 = blmicmtrmdl.blmdkvatou3 == undefined ? 0.00 : blmicmtrmdl.blmdkvatou3;
                _spBill.LMD_kVA_for_TOU4 = blmicmtrmdl.blmdkvatou4 == undefined ? 0.00 : blmicmtrmdl.blmdkvatou4;
                _spBill.LMD_kVA_for_TOU5 = blmicmtrmdl.blmdkvatou5 == undefined ? 0.00 : blmicmtrmdl.blmdkvatou5;
                _spBill.LMD_kVA_for_TOU6 = blmicmtrmdl.blmdkvatou6 == undefined ? 0.00 : blmicmtrmdl.blmdkvatou6;
                _spBill.LMD_kVA_for_TOU7 = blmicmtrmdl.blmdkvatou7 == undefined ? 0.00 : blmicmtrmdl.blmdkvatou7;
                _spBill.LMD_kVA_for_TOU8 = blmicmtrmdl.blmdkvatou8 == undefined ? 0.00 : blmicmtrmdl.blmdkvatou8;
                if (_prepaydt != undefined) {
                    _spBill.Cumulative_Emergency_Credit = _prepaydt[0].pdispcumemercred == null ? 0.00 : _prepaydt[0].pdispcumemercred;
                    _spBill.Cum_Monthly_Fixed_Charge_deduction_Amount = _prepaydt[0].pdispcummthfxdchrgdedamt == null ? 0.00 : _prepaydt[0].pdispcummthfxdchrgdedamt;;
                    _spBill.Cum_Demand_Deduction_Amount = _prepaydt[0].pdispcumdmddedamt == null ? 0.00 : _prepaydt[0].pdispcumdmddedamt;
                    _spBill.Cum_Adjustable_Amount = _prepaydt[0].pdispcumadjamt == null ? 0.00 : _prepaydt[0].pdispcumadjamt;
                    _spBill.Cum_Daily_Fixed_Charge_deduction_Amount = _prepaydt[0].pdispcdfcdamt == null ? 0.00 : _prepaydt[0].pdispcdfcdamt;
                    _spBill.CumEnergy_Charge_Deduction_Amount = _prepaydt[0].pdispcumenrgchrgdedamt == null ? 0.00 : _prepaydt[0].pdispcumenrgchrgdedamt;
                }
                console.log(_spBill);
            }
        });
    }
    catch (err) {
        console.log(err);
    }
}

async function CreateFileFolder(path) {
    fs.access(path, (error) => {
        if (error) {
            fs.mkdir(path, (error) => {
                if (error) {
                    console.log(error);
                } else {
                    console.log("New Directory created successfully !!" + path);
                }
            });
        } else {
            //console.log("Given Directory already exists !!");
        }
    });
}