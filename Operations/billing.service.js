const dbConf = require('../Configuration/dbConfig');
var Excel = require("exceljs");// load exceljs module
const fs = require("fs");
const logger = require("../Configuration/logger");
const Billing_Model = require('../ViewModels/Billing_Model');
const spBill = require("../ViewModels/Billing_Model");
const prepSlab = require("../ViewModels/Prep_Slab");
const billGen = require("../ViewModels/Bill_Generate");
const sourcePath = "D:\\Env_Temp_std\\";
const _common = require("./common.service");
const _singlePhaseBill = require("./single_phase.billing");
//GetBillingData(1680307200,1682812800);
var fromTime = 0;
var toTime = 0;
ValidateDate();

async function ValidateDate() {
    console.log("Service Started...");
    var _currentDate = new Date();
    var day = _currentDate.getDate();
    if (day == 27) {
        var firstDay = new Date(_currentDate.getFullYear(), _currentDate.getMonth(), 1);
        var lastDay = new Date(_currentDate.getFullYear(), _currentDate.getMonth() + 1, 0);

        var d = firstDay.valueOf();
        var epoch = d / 1000;
        var fromepoch = epoch;
        d = lastDay.valueOf();
        epoch = d / 1000;
        var toepoch = epoch;
        var billmonth = firstDay.getMonth();

        fromTime = 1680307200;
        toTime = 1682812800;

        //GetBillingData(fromTime, toTime, 03, firstDay, lastDay);
        _singlePhaseBill.GetBillingDetails(fromTime, toTime);
    }
}

async function GetBillingData(from, to, billmonth, firstDay, lastDay) {
    var _query = "select * from envisage_dev.billinglog where bltimestamp >=" + from + " and bltimestamp<=" + to + " and sent is null order by blrecid;";
    let result;
    //logger.info("Query 1 : " + _query);
    dbConf.pool.open(dbConf.connStr, (err, conn) => {
        if (err) {
            console.log(err);
        }
        result = dbConf.runSQL(conn, _query);
        conn.close();

        if (result != undefined) {
            BillGenerationLog(result, billmonth, from, to, firstDay, lastDay);
        }
    });
}


async function BillGenerationLog(result, billmonth, from, to, firstDay, lastDay) {
    try {
        if (result.length > 0) {
            for (var i = 0; i < result.length; i++) {
                var billGens = billGen;
                billGens.billMonth = billmonth;
                billGens.billstdate = firstDay.toString();
                billGens.billenddate = lastDay.toString();
                billGens.billYear = firstDay.getFullYear();
                getMeterConfiguration(result[i], billmonth, from, to, billGens);
            }
        }
    }
    catch (err) {
        logger.info("BillGenerationLog Error : " + err);
    }
}

async function getMeterConfiguration(data, billmonth, from, to, billGens) {
    var _query = "select * from envisage_dev.mtrinstallconfig where micmtrsrno = '" + data.blmetersrno + "'";
    let result;
    //logger.info("Query 2 : " + _query);
    dbConf.pool.open(dbConf.connStr, (err, conn) => {
        if (err) {
            console.log(err);
        }
        result = dbConf.runSQL(conn, _query);
        conn.close();

        if (result != undefined) {
            GetConsumerTarrif(result, data, billmonth, from, to, billGens);
        }
    });
}

async function GetConsumerTarrif(mic, bl, billmonth, from, to, billGens) {
    try {
        billGens.txnconsumerid = mic[0].micconsumerid;
        billGens.txnmtrsrno = mic[0].micmtrsrno;
        //console.log("Consumer Number : " + billGens.Bill_Gen.txnmtrsrno);
        var connum = mic[0].micconsumerid.substring(6, 4);
        billGens.billnumber = billGens.billMonth + billGens.billYear + connum;
        var _query = "select * from envisage_dev.cmrattrib m where m.cmraconsumerid = '" + mic[0].micconsumerid + "'";
        let result;
        let trfData;
        //logger.info("Query 3 : " + _query);
        dbConf.pool.open(dbConf.connStr, (err, conn) => {
            if (err) {
                console.log(err);
            }
            result = dbConf.runSQL(conn, _query);
            let _inst;
            if (result != undefined) {
                _query = "select * from envisage_dev.ccattouslottrfdetail c join prepmsttariff p on c.tstdtariffid = p.preptrfid where c.tstdconclassid= '" + result[0].cmracurclassid + "' and c.tstdconcatid ='" + result[0].cmracurcategoryid + "'";
                //logger.info("Query 4 : " + _query);
                trfData = dbConf.runSQL(conn, _query);
                _query = "select * from envisage_dev.mstconsumermeterrelation cm join envisage_dev.mtrinstallconfig mstin on cm.cmrconsumerid = mstin.micconsumerid";
                _query += " join envisage_dev.extrafields_kinesco e on cm.cmrconsumermasterid = e.consumermasterid where mstin.micconsumerid = '" + billGens.Bill_Gen.txnconsumerid + "'";
                //logger.info("Query 5 : " + _query);
                let exf = dbConf.runSQL(conn, _query);
                _query = "select * from envisage_dev.mstconsumermeterrelation cm join envisage_dev.mtrinstallconfig mstin on cm.cmrconsumerid = mstin.micconsumerid";
                _query += " where mstin.micconsumerid ='" + billGens.txnconsumerid + "'";
                //logger.info("Query 6 : " + _query);
                let _mstrid = dbConf.runSQL(conn, _query);
                if (_mstrid != undefined && _mstrid.length > 0) {
                    _query = "select * from envisage_dev.mstconsumermeterrelation m where m.cmrconsumermasterid = '" + _mstrid[0].cmrconsumermasterid + "' and m.cmrrelationenddate is null";
                    //logger.info("Query 7 : " + _query);
                    _inst = dbConf.runSQL(conn, _query);
                    //logger.info("Query 7 : Result Fetch");
                    if (_inst != undefined && _inst.length > 0)
                        billGens.fixchargunit = _inst.length;
                    else
                        billGens.fixchargunit = "1";
                    //logger.info("Query 7 : Result Fetch 1");
                }
                else
                    billGens.fixchargunit = "1";
                if (exf != undefined && exf.length > 0) {
                    billGens.gstnumber = exf[0].gstnumber;
                }
                else
                    billGens.gstnumber = "";
                conn.close();
                if (trfData != undefined) {
                    var conClass = result[0].cmracurclassid;
                    var conCat = result[0].cmracurcategoryid;
                    billGens.tarrifcode = trfData[0].preptrfid;
                    StartBilling(mic, bl, trfData, billmonth, from, to, conClass, conCat, billGens);
                }
            }
        });
    }
    catch (err) {
        logger.info("GetConsumerTarrif Error : " + err + "\nQuery : " + _query);
    }
}

async function StartBilling(mic, bl, trfData, billmonth, from, to, conClass, conCat, billGens) {
    try {
        var _query = "select * from envisage_dev.preptarrifdiff where ptdftarrid ='" + trfData[0].preptrfid + "'";
        let result;
        //logger.info("Query 8 : " + _query);
        dbConf.pool.open(dbConf.connStr, (err, conn) => {
            if (err) {
                console.log(err);
            }
            result = dbConf.runSQL(conn, _query);

            if (result != undefined) {
                let phase;
                _query = "select * from envisage_dev.mstmeter m join envisage_dev.mtrinstallconfig mic on m.mtrsrno= mic.micmtrsrno join envisage_dev.mstmtrmodel md on m.mtrmodelid = md.mdlid where m.mtrsrno = '" + mic[0].micmtrsrno + "'";
                //logger.info("Query 9 : " + _query);
                phase = dbConf.runSQL(conn, _query);
                _query = "select ab.* from envisage_dev.prepmsttrffenrgyslabs ab join envisage_dev.prepmsttariff pr on ab.preptdltrfid = pr.preptrfid where ab.preptdltrfid = '" + trfData[0].preptrfid + "' and pr.preptrftodate is null order by ab.preptdirecid desc, ab.preptdltrfslotno asc";
                //logger.info("Query 10 : " + _query);
                var _prepSlab = dbConf.runSQL(conn, _query);
                //logger.info("Query 10 : Data Fetch");
                if (_prepSlab != undefined) {
                    ServedBillingFile(mic, bl, trfData, billmonth, from, to, conClass, conCat, result[0].ptdftarrtempdis, phase, _prepSlab, billGens);
                }
                conn.close();
            }
        });
    }
    catch (err) {
        logger.info("StartBilling Error : " + err);
    }
}

async function ServedBillingFile(mic, bl, trfData, billmonth, from, to, conClass, conCat, fileName, phase, _prepSlab, billGens) {
    try {
        logger.info("ServedBillingFile : 001");
        var d = new Date();
        let _spData = Billing_Model.SPBill_Model;
        let _tpData = Billing_Model.TPBill_Model;
        let _mtrinstsp;
        let _prepaydt;
        let _trfdmdslabs;
        let _conld;
        let _meters;
        let _toudtls;
        if (phase[0].mdlphase == "S") {
            var _query = "Select * from mtrinstant_sp_log m where m.misplsrno = '" + mic[0].micmtrsrno + "' and (m.mispldate >= " + from + " and m.mispldate <= " + to + ") order by 1 desc limit 1;";
            //logger.info("Query 11 : " + _query);
            dbConf.pool.open(dbConf.connStr, (err, conn) => {
                if (err) {
                    console.log(err);
                }
                _mtrinstsp = dbConf.runSQL(conn, _query);
                _query = "select * from envisage_dev.prepmsttrffdmdslab bc where bc.preptdltrfid = '" + trfData[0].preptrfid + "'";
                //logger.info("Query 12 : " + _query);
                _trfdmdslabs = dbConf.runSQL(conn, _query);
                _query = "select * from envisage_dev.mstcmrenergydetail c join mtrinstallconfig inst on c.cmreconsumerid = inst.micconsumerid where c.cmreconsumerid = '" + mic[0].micconsumerid + "'";
                //logger.info("Query 13 : " + _query);
                _conld = dbConf.runSQL(conn, _query);
                _query = "select m.mtrsrno,mic.micdcusrno,mm.mdlid,prptr.prepenrgtrftype,prptr.prepdmdtrftype,prptr.prepuntexsdmdchrg,prptr.prepmthfixedchrg,";
                _query += "prptr.prepdlyfixedchrg,prptr.prepemrgncycrlmt,prptr.preprsrvecrlmtrcl,prptr.preptrfflatenerrate,prptr.preptrfflatdmdrate,prptr.preptrfuserid,";
                _query += "prptr.preptrfid,ccat.cattrtouid,mm.mdlphase,mic.micdtrid,mic.micservicestartdate,prptr.preptaxrchrg,prptr.prepuntdmdchrg,";
                _query += "prptr.prepfixcharge,prptr.prepgstper,prptr.prepfuelsuchrg,prptr.prepmtrhire,prptr.preplowvoltsuchrg,prptr.preptrfschemedescription,";
                _query += " cmra.cmrafromdate, cmra.cmrasznsz from envisage_dev.mstmtrmodel mm join envisage_dev.mstmeter m on mm.mdlid = m.mtrmodelid ";
                _query += "envisage_dev.mtrinstallconfig mic on m.mtrsrno = mic.micmtrsrno join envisage_dev.cmrattrib cmra on mic.micconsumerid = cmra.cmraconsumerid ";
                _query += "envisage_dev.ccattarifftourel ccat on cmra.cmracurcategoryid = ccat.cattrconsumercatid and cmra.cmracurclassid = ccat.cattrconsumerclassid ";
                _query += "envisage_dev.ccattouslottrfdetail ccatrf on cmra.cmracurcategoryid = ccatrf.tstdconcatid and cmra.cmracurclassid = ccatrf.tstdconclassid ";
                _query += "envisage_dev.prepmsttariff prptr on ccatrf.tstdtariffid = prptr.preptrfid where m.mtrsrno = '" + mic[0].micmtrsrno + "' and ccat.cattrtodate is null ";
                _query += "mic.micconsumerid ='" + mic[0].micconsumerid + "' and cmra.cmratodate is null and prptr.preptrftodate is null";
                //logger.info("Query 14 : " + _query);
                var znid = "";
                _meters = dbConf.runSQL(conn, _query);
                if (_meters != undefined && _meters.length > 0) {
                    znid = _meters[0].cattrtouid;
                    _query = "select cd.pretsslotno,cd.pretsslotstarthr,cd.pretsslotstartmin,cd.pretsslotendhr,cd.pretsslotendmin,cd.pretsincconamt,cd.pretsincdmdamt,cd.pretsload,cd.pretsminchrgdmdlim,cd.tsmaxdemandlimit from envisage_dev.prepmsttouslot cd where cd.pretstouid = '" + znid + "'";
                    //logger.info("Query 15 : " + _query);
                    _toudtls = dbConf.runSQL(conn, _query);
                    if (_meters[0].cmrasznsz != undefined && _meters[0].cmrasznsz != null)
                        billGens.seznsez = _meters[0].cmrasznsz;
                    else
                        billGens.seznsez = "";
                }
                else
                    billGens.seznsez = "";

                if (_mtrinstsp != undefined) {
                    _query = "select * from prepaydtlinst_sp p where p.pdispmtrsrno = '" + mic[0].micmtrsrno + "' and (p.pdispdate >= " + from + " and p.pdispdate <= " + to + ") order by 1 desc limit 1;"
                    //logger.info("Query 16 : " + _query);
                    _prepaydt = dbConf.runSQL(conn, _query);
                    conn.close();
                    _spData.RTC_absolute = d.toString();
                    _spData.Active_Power = _mtrinstsp[0].misplmdkw == undefined ? 0 : _mtrinstsp[0].misplmdkw;
                    _spData.Apparent_Power = _mtrinstsp[0].misplmdkva == undefined ? 0 : _mtrinstsp[0].misplmdkva;
                    _spData.Instantaneous_Frequency = _mtrinstsp[0].misplfreq == undefined ? 0 : _mtrinstsp[0].misplfreq;
                    _spData.Signed_Reactive_power = _mtrinstsp[0].misplkvar == undefined ? 0 : _mtrinstsp[0].misplkvar;
                    _spData.Signed_Three_phase_Power_factor = _mtrinstsp[0].misplpf == undefined ? 0 : _mtrinstsp[0].misplpf;
                    _spData.Cumulative_Energy_KWh = bl.blcumkwh == undefined ? 0 : bl.blcumkwh;
                    _spData.Cumulative_Energy_KVAh = bl.blcumkvah == undefined ? 0 : bl.blcumkvah;
                    _spData.Cumulative_Energy_KVARh_Lag = bl.blcumkvarh_lag == undefined ? 0 : bl.blcumkvarh_lag;
                    _spData.Cumulative_Energy_KVARh_Lead = bl.blcumkvarh_lead == undefined ? 0 : bl.blcumkvarh_lead;
                    _spData.Maximum_demand_KW = bl.blmdkw == undefined ? 0 : bl.blmdkw;
                    _spData.Maximum_demand_KVA = bl.blmdkva == undefined ? 0 : bl.blmdkva;
                    _spData.Cumulative_tamper_count = bl.blcumtampercount == undefined ? 0 : bl.blcumtampercount;
                    _spData.Billing_Date = bl.bltimestamp == undefined ? 0 : bl.bltimestamp;
                    _spData.Signed_Average_PowerFactor = bl.blavgpf == undefined ? 0.00 : bl.blavgpf;
                    _spData.Last_Month_Cumulative_EnergykWh = bl.blcumkwh == undefined ? 0.00 : bl.blcumkwh;
                    _spData.LCumulative_Energy_for_TOU1 = bl.blcumkwhtou1 == undefined ? 0.00 : bl.blcumkwhtou1;
                    _spData.LCumulative_Energy_for_TOU2 = bl.blcumkwhtou2 == undefined ? 0.00 : bl.blcumkwhtou2;
                    _spData.LCumulative_Energy_for_TOU3 = bl.blcumkwhtou3 == undefined ? 0.00 : bl.blcumkwhtou3;
                    _spData.LCumulative_Energy_for_TOU4 = bl.blcumkwhtou4 == undefined ? 0.00 : bl.blcumkwhtou4;
                    _spData.LCumulative_Energy_for_TOU5 = bl.blcumkwhtou5 == undefined ? 0.00 : bl.blcumkwhtou5;
                    _spData.LCumulative_Energy_for_TOU6 = bl.blcumkwhtou6 == undefined ? 0.00 : bl.blcumkwhtou6;
                    _spData.LCumulative_Energy_for_TOU7 = bl.blcumkwhtou7 == undefined ? 0.00 : bl.blcumkwhtou7;
                    _spData.LCumulative_Energy_for_TOU8 = bl.blcumkwhtou8 == undefined ? 0.00 : bl.blcumkwhtou8;

                    _spData.LCumulative_Energy_kVARh_Lag = bl.blcumkvarh_lag == undefined ? 0.00 : bl.blcumkvarh_lag;
                    _spData.LCumulative_Energy_kVARh_Lead = bl.blcumkvarh_lead == undefined ? 0.00 : bl.blcumkvarh_lead;
                    _spData.LCumulative_Energy_kVAh = bl.blcumkvah == undefined ? 0.00 : bl.blcumkvah;
                    _spData.LCumulative_Apparent_Energy_for_TOU1 = bl.blcumkvahtou1 == undefined ? 0.00 : bl.blcumkvahtou1;
                    _spData.LCumulative_Apparent_Energy_for_TOU2 = bl.blcumkvahtou2 == undefined ? 0.00 : bl.blcumkvahtou2;
                    _spData.LCumulative_Apparent_Energy_for_TOU3 = bl.blcumkvahtou3 == undefined ? 0.00 : bl.blcumkvahtou3;
                    _spData.LCumulative_Apparent_Energy_for_TOU4 = bl.blcumkvahtou4 == undefined ? 0.00 : bl.blcumkvahtou4;
                    _spData.LCumulative_Apparent_Energy_for_TOU5 = bl.blcumkvahtou5 == undefined ? 0.00 : bl.blcumkvahtou5;
                    _spData.LCumulative_Apparent_Energy_for_TOU6 = bl.blcumkvahtou6 == undefined ? 0.00 : bl.blcumkvahtou6;
                    _spData.LCumulative_Apparent_Energy_for_TOU7 = bl.blcumkvahtou7 == undefined ? 0.00 : bl.blcumkvahtou7;
                    _spData.LCumulative_Apparent_Energy_for_TOU8 = bl.blcumkvahtou8 == undefined ? 0.00 : bl.blcumkvahtou8;
                    _spData.LMDMaximum_Demand_kW = bl.blmdkw == undefined ? 0.00 : bl.blmdkw;
                    _spData.LMD_kW_for_TOU1 = bl.blmdkwtou1 == undefined ? 0.00 : bl.blmdkwtou1;
                    _spData.LMD_kW_for_TOU2 = bl.blmdkwtou2 == undefined ? 0.00 : bl.blmdkwtou2;
                    _spData.LMD_kW_for_TOU3 = bl.blmdkwtou3 == undefined ? 0.00 : bl.blmdkwtou3;
                    _spData.LMD_kW_for_TOU4 = bl.blmdkwtou4 == undefined ? 0.00 : bl.blmdkwtou4;
                    _spData.LMD_kW_for_TOU5 = bl.blmdkwtou5 == undefined ? 0.00 : bl.blmdkwtou5;
                    _spData.LMD_kW_for_TOU6 = bl.blmdkwtou6 == undefined ? 0.00 : bl.blmdkwtou6;
                    _spData.LMD_kW_for_TOU7 = bl.blmdkwtou7 == undefined ? 0.00 : bl.blmdkwtou7;
                    _spData.LMD_kW_for_TOU8 = bl.blmdkwtou8 == undefined ? 0.00 : bl.blmdkwtou8;
                    _spData.LMDMaximum_Demand_kVA = bl.blmdkva == undefined ? 0.00 : bl.blmdkva;
                    _spData.LMD_kVA_for_TOU1 = bl.blmdkvatou1 == undefined ? 0.00 : bl.blmdkvatou1;
                    _spData.LMD_kVA_for_TOU2 = bl.blmdkvatou2 == undefined ? 0.00 : bl.blmdkvatou2;
                    _spData.LMD_kVA_for_TOU3 = bl.blmdkvatou3 == undefined ? 0.00 : bl.blmdkvatou3;
                    _spData.LMD_kVA_for_TOU4 = bl.blmdkvatou4 == undefined ? 0.00 : bl.blmdkvatou4;
                    _spData.LMD_kVA_for_TOU5 = bl.blmdkvatou5 == undefined ? 0.00 : bl.blmdkvatou5;
                    _spData.LMD_kVA_for_TOU6 = bl.blmdkvatou6 == undefined ? 0.00 : bl.blmdkvatou6;
                    _spData.LMD_kVA_for_TOU7 = bl.blmdkvatou7 == undefined ? 0.00 : bl.blmdkvatou7;
                    _spData.LMD_kVA_for_TOU8 = bl.blmdkvatou8 == undefined ? 0.00 : bl.blmdkvatou8;
                    if (_prepaydt != undefined) {
                        _spData.Cumulative_Emergency_Credit = _prepaydt[0].pdispcumemercred == null ? 0.00 : _prepaydt[0].pdispcumemercred;
                        _spData.Cum_Monthly_Fixed_Charge_deduction_Amount = _prepaydt[0].pdispcummthfxdchrgdedamt == null ? 0.00 : _prepaydt[0].pdispcummthfxdchrgdedamt;;
                        _spData.Cum_Demand_Deduction_Amount = _prepaydt[0].pdispcumdmddedamt == null ? 0.00 : _prepaydt[0].pdispcumdmddedamt;
                        _spData.Cum_Adjustable_Amount = _prepaydt[0].pdispcumadjamt == null ? 0.00 : _prepaydt[0].pdispcumadjamt;
                        _spData.Cum_Daily_Fixed_Charge_deduction_Amount = _prepaydt[0].pdispcdfcdamt == null ? 0.00 : _prepaydt[0].pdispcdfcdamt;
                        _spData.CumEnergy_Charge_Deduction_Amount = _prepaydt[0].pdispcumenrgchrgdedamt == null ? 0.00 : _prepaydt[0].pdispcumenrgchrgdedamt;
                    }
                    //console.log(_spData);
                }
            });
        }
        else if (phase[0].mdlphase == "T") {
            _tpData.RTC_absolute = d.toString();
            _tpData.Current_IR = 0.00;
            _tpData.Current_IY = 0.00;
            _tpData.Current_IB = 0.00;
            _tpData.Voltage_VRN = 0.00;
            _tpData.Voltage_VYN = 0.00;
            _tpData.Voltage_VBN = 0.00;
            _tpData.Active_Power = 0.00;
            _tpData.Apparent_Power = 0.00;
            _tpData.Instantaneous_Frequency = 0.00;
            _tpData.Signed_Reactive_power = 0.00;
            _tpData.Signed_Three_phase_Power_factor = 0.00;
            _tpData.Cumulative_Energy_KWh = bl.blcumkwh == null ? 0.00 : bl.blcumkwh;
            _tpData.Cumulative_Energy_KVAh = bl.blcumkvah == null ? 0.00 : bl.blcumkvah;
            _tpData.Cumulative_Energy_KVARh_Lag = bl.blcumkvarh_lag == null ? 0.00 : bl.blcumkvarh_lag;
            _tpData.Cumulative_Energy_KVARh_Lead = bl.blcumkvarh_lead == null ? 0.00 : bl.blcumkvarh_lead;
            _tpData.Maximum_demand_KW = bl.blmdkw == null ? 0.00 : bl.blmdkw;
            _tpData.Maximum_demand_KVA = bl.blmdkva == null ? 0.00 : bl.blmdkva;
            _tpData.Cumulative_tamper_count = bl.blcumtampercount == null ? 0 : bl.blcumtampercount;
            _tpData.Billing_Date = bl.bltimestamp == null ? 00000000 : bl.bltimestamp;
            _tpData.Billing_Index = 1;
            _tpData.Signed_Average_PowerFactor = bl.blavgpf == null ? 0.00 : bl.blavgpf;
            _tpData.Last_Month_Cumulative_EnergykWh = bl.blcumkwh == null ? 0.00 : bl.blcumkwh;
            _tpData.LCumulative_Energy_for_TOU1 = bl.blcumkwhtou1 == null ? 0.00 : bl.blcumkwhtou1;
            _tpData.LCumulative_Energy_for_TOU2 = bl.blcumkwhtou2 == null ? 0.00 : bl.blcumkwhtou2;
            _tpData.LCumulative_Energy_for_TOU3 = bl.blcumkwhtou3 == null ? 0.00 : bl.blcumkwhtou3;
            _tpData.LCumulative_Energy_for_TOU4 = bl.blcumkwhtou4 == null ? 0.00 : bl.blcumkwhtou4;
            _tpData.LCumulative_Energy_for_TOU5 = bl.blcumkwhtou5 == null ? 0.00 : bl.blcumkwhtou5;
            _tpData.LCumulative_Energy_for_TOU6 = bl.blcumkwhtou6 == null ? 0.00 : bl.blcumkwhtou6;
            _tpData.LCumulative_Energy_for_TOU7 = bl.blcumkwhtou7 == null ? 0.00 : bl.blcumkwhtou7;
            _tpData.LCumulative_Energy_for_TOU8 = bl.blcumkwhtou8 == null ? 0.00 : bl.blcumkwhtou8;

            _tpData.exLCumulative_Energy_for_TOU1 = bl.bltouexpkwh0 == null ? 0.00 : bl.bltouexpkwh0;
            _tpData.exLCumulative_Energy_for_TOU2 = bl.bltouexpkwh1 == null ? 0.00 : bl.bltouexpkwh1;
            _tpData.exLCumulative_Energy_for_TOU3 = bl.bltouexpkwh2 == null ? 0.00 : bl.bltouexpkwh2;
            _tpData.exLCumulative_Energy_for_TOU4 = bl.bltouexpkwh3 == null ? 0.00 : bl.bltouexpkwh3;
            _tpData.exLCumulative_Energy_for_TOU5 = bl.bltouexpkwh4 == null ? 0.00 : bl.bltouexpkwh4;
            _tpData.exLCumulative_Energy_for_TOU6 = bl.bltouexpkwh5 == null ? 0.00 : bl.bltouexpkwh5;
            _tpData.exLCumulative_Energy_for_TOU7 = bl.bltouexpkwh6 == null ? 0.00 : bl.bltouexpkwh6;
            _tpData.exLCumulative_Energy_for_TOU8 = bl.bltouexpkwh7 == null ? 0.00 : bl.bltouexp;
            _tpData.LCumulative_Energy_kVARh_Lag = bl.blcumkvarh_lag == null ? 0.00 : bl.blcumkvarh_lag;
            _tpData.LCumulative_Energy_kVARh_Lead = bl.blcumkvarh_lead == null ? 0.00 : bl.blcumkvarh_lead;
            _tpData.excum_Energy_kVARh_Lag = bl.blexpkvarhlagq2 == null ? 0.00 : bl.blexpkvarhlagq2;
            _tpData.excum_Energy_kVARh_Lead = bl.blexpkvarhlegq3 == null ? 0.00 : bl.blexpkvarhl;
            _tpData.LCumulative_Energy_kVAh = bl.blcumkvah == null ? 0.00 : bl.blcumkvah;
            _tpData.LCumulative_Apparent_Energy_for_TOU1 = bl.blcumkvahtou1 == null ? 0.00 : bl.blcumkvahtou1;
            _tpData.LCumulative_Apparent_Energy_for_TOU2 = bl.blcumkvahtou2 == null ? 0.00 : bl.blcumkvahtou2;
            _tpData.LCumulative_Apparent_Energy_for_TOU3 = bl.blcumkvahtou3 == null ? 0.00 : bl.blcumkvahtou3;
            _tpData.LCumulative_Apparent_Energy_for_TOU4 = bl.blcumkvahtou4 == null ? 0.00 : bl.blcumkvahtou4;
            _tpData.LCumulative_Apparent_Energy_for_TOU5 = bl.blcumkvahtou5 == null ? 0.00 : bl.blcumkvahtou5;
            _tpData.LCumulative_Apparent_Energy_for_TOU6 = bl.blcumkvahtou6 == null ? 0.00 : bl.blcumkvahtou6;
            _tpData.LCumulative_Apparent_Energy_for_TOU7 = bl.blcumkvahtou7 == null ? 0.00 : bl.blcumkvahtou7;
            _tpData.LCumulative_Apparent_Energy_for_TOU8 = bl.blcumkvahtou8 == null ? 0.00 : bl.blcumkvahtou8;
            _tpData.excum_Energy_kVAh = bl.blexpkvah == null ? 0.00 : bl.blexpkvah;
            _tpData.excum_Apparent_Energy_for_TOU1 = bl.bltouexpkvah0 == null ? 0.00 : bl.bltouexpkvah0;
            _tpData.excum_Apparent_Energy_for_TOU2 = bl.bltouexpkvah1 == null ? 0.00 : bl.bltouexpkvah1;
            _tpData.excum_Apparent_Energy_for_TOU3 = bl.bltouexpkvah2 == null ? 0.00 : bl.bltouexpkvah2;
            _tpData.excum_Apparent_Energy_for_TOU4 = bl.bltouexpkvah3 == null ? 0.00 : bl.bltouexpkvah3;
            _tpData.excum_Apparent_Energy_for_TOU5 = bl.bltouexpkvah4 == null ? 0.00 : bl.bltouexpkvah4;
            _tpData.excum_Apparent_Energy_for_TOU6 = bl.bltouexpkvah5 == null ? 0.00 : bl.bltouexpkvah5;
            _tpData.excum_Apparent_Energy_for_TOU7 = bl.bltouexpkvah6 == null ? 0.00 : bl.bltouexpkvah6;
            _tpData.excum_Apparent_Energy_for_TOU8 = bl.bltouexpkvah7 == null ? 0.00 : bl.bltouexpkvah7;
            _tpData.LMDMaximum_Demand_kW = bl.blmdkw == null ? 0.00 : bl.blmdkw;
            _tpData.LMD_kW_for_TOU1 = bl.blmdkwtou1 == null ? 0.00 : bl.blmdkwtou1;
            _tpData.LMD_kW_for_TOU2 = bl.blmdkwtou2 == null ? 0.00 : bl.blmdkwtou2;
            _tpData.LMD_kW_for_TOU3 = bl.blmdkwtou3 == null ? 0.00 : bl.blmdkwtou3;
            _tpData.LMD_kW_for_TOU4 = bl.blmdkwtou4 == null ? 0.00 : bl.blmdkwtou4;
            _tpData.LMD_kW_for_TOU5 = bl.blmdkwtou5 == null ? 0.00 : bl.blmdkwtou5;
            _tpData.LMD_kW_for_TOU6 = bl.blmdkwtou6 == null ? 0.00 : bl.blmdkwtou6;
            _tpData.LMD_kW_for_TOU7 = bl.blmdkwtou7 == null ? 0.00 : bl.blmdkwtou7;
            _tpData.LMD_kW_for_TOU8 = bl.blmdkwtou8 == null ? 0.00 : bl.blmdkwtou8;
            _tpData.exMDMaximum_Demand_kW = bl.blexpmdkw == null ? 0.00 : bl.blexpmdkw;
            _tpData.exMD_kW_for_TOU1 = bl.bltouexpmdkw0 == null ? 0.00 : bl.bltouexpmdkw0;
            _tpData.exMD_kW_for_TOU2 = bl.bltouexpmdkw1 == null ? 0.00 : bl.bltouexpmdkw1;
            _tpData.exMD_kW_for_TOU3 = bl.bltouexpmdkw2 == null ? 0.00 : bl.bltouexpmdkw2;
            _tpData.exMD_kW_for_TOU4 = bl.bltouexpmdkw3 == null ? 0.00 : bl.bltouexpmdkw3;
            _tpData.exMD_kW_for_TOU5 = bl.bltouexpmdkw4 == null ? 0.00 : bl.bltouexpmdkw4;
            _tpData.exMD_kW_for_TOU6 = bl.bltouexpmdkw5 == null ? 0.00 : bl.bltouexpmdkw5;
            _tpData.exMD_kW_for_TOU7 = bl.bltouexpmdkw6 == null ? 0.00 : bl.bltouexpmdkw6;
            _tpData.exMD_kW_for_TOU8 = bl.bltouexpmdkw7 == null ? 0.00 : bl.bltouexpmdkw7;
            _tpData.LMDMaximum_Demand_kVA = bl.blmdkva == null ? 0.00 : bl.blmdkva;
            _tpData.LMD_kVA_for_TOU1 = bl.blmdkvatou1 == null ? 0.00 : bl.blmdkvatou1;
            _tpData.LMD_kVA_for_TOU2 = bl.blmdkvatou2 == null ? 0.00 : bl.blmdkvatou2;
            _tpData.LMD_kVA_for_TOU3 = bl.blmdkvatou3 == null ? 0.00 : bl.blmdkvatou3;
            _tpData.LMD_kVA_for_TOU4 = bl.blmdkvatou4 == null ? 0.00 : bl.blmdkvatou4;
            _tpData.LMD_kVA_for_TOU5 = bl.blmdkvatou5 == null ? 0.00 : bl.blmdkvatou5;
            _tpData.LMD_kVA_for_TOU6 = bl.blmdkvatou6 == null ? 0.00 : bl.blmdkvatou6;
            _tpData.LMD_kVA_for_TOU7 = bl.blmdkvatou7 == null ? 0.00 : bl.blmdkvatou7;
            _tpData.LMD_kVA_for_TOU8 = bl.blmdkvatou8 == null ? 0.00 : bl.blmdkvatou8;
            _tpData.exMDMaximum_Demand_kVA = bl.blexpmdkva == null ? 0.00 : bl.blexpmdkva;
            _tpData.exMD_kVA_for_TOU1 = bl.bltouexpmdkva0 == null ? 0.00 : bl.bltouexpmdkva0;
            _tpData.exMD_kVA_for_TOU2 = bl.bltouexpmdkva1 == null ? 0.00 : bl.bltouexpmdkva1;
            _tpData.exMD_kVA_for_TOU3 = bl.bltouexpmdkva2 == null ? 0.00 : bl.bltouexpmdkva2;
            _tpData.exMD_kVA_for_TOU4 = bl.bltouexpmdkva3 == null ? 0.00 : bl.bltouexpmdkva3;
            _tpData.exMD_kVA_for_TOU5 = bl.bltouexpmdkva4 == null ? 0.00 : bl.bltouexpmdkva4;
            _tpData.exMD_kVA_for_TOU6 = bl.bltouexpmdkva5 == null ? 0.00 : bl.bltouexpmdkva5;
            _tpData.exMD_kVA_for_TOU7 = bl.bltouexpmdkva6 == null ? 0.00 : bl.bltouexpmdkva6;
            _tpData.exMD_kVA_for_TOU8 = bl.bltouexpmdkva7 == null ? 0.00 : bl.bltouexpmdkva7;

            _tpData.Cumulative_Emergency_Credit = 0.00;
            _tpData.Cum_Monthly_Fixed_Charge_deduction_Amount = 0.00;
            _tpData.Cum_Demand_Deduction_Amount = 0.00;
            _tpData.Cum_Adjustable_Amount = 0.00;
            _tpData.Cum_Daily_Fixed_Charge_deduction_Amount = 0.00;
            _tpData.CumEnergy_Charge_Deduction_Amount = 0.00;
            _tpData.Cumulative_recharge_amount = 0.00;
            _tpData.Cumulative_Balance_Deduction_Register = 0.00;

            //console.log(_tpData);
        }
        ProcessBilling(phase, _spData, _tpData, fileName, trfData, mic, billmonth, _prepSlab, _trfdmdslabs, _conld, _toudtls, billGens, bl);
    }
    catch (err) {
        logger.info("ServedBillingFile Error : " + err);
    }
}


async function ProcessBilling(phase, _spData, _tpData, _fileName, _trfData, _mic, billmonth, _prepSlab, _trfdmdslabs, _conld, _toudtls, billGens, billlog, phase) {
    try {
        var d = new Date();
        var year = d.getFullYear();
        var _folderPath = sourcePath + "//" + billmonth + "_" + year + "//";
        CreateFileFolder(_folderPath);
        var _tarrifFolder = _folderPath + _trfData[0].preptrfid + "//";
        CreateFileFolder(_tarrifFolder);
        // if (r == 1) {
        //
        //     console.log(r);
        //     if (r == 1) {
        //         r = _common.CopyFileToAnother(sourcePath + _fileName, _tarrifFolder + _mic[0].micconsumerid + ".xls");
        //         console.log(r);
        //         if (r == 0) {
        //             UpdateExcel(_tarrifFolder + _mic[0].micconsumerid + ".xls", _spData, _tpData, _trfData, _mic, billmonth, _prepSlab, _trfdmdslabs, _conld, phase, _toudtls, billGens, billlog);
        //         }
        //     }
        // }

        fs.copyFileSync(sourcePath + _fileName, _tarrifFolder + _mic[0].micconsumerid + ".xls");
    }
    catch (err) {
        logger.info("ProcessBilling : " + err);
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

async function UpdateExcel(xlFilePath, _spdata, _tpdata, _trfdata, _mic, billmonth, _prepSlab, _trfdmdslabs, _conld, phase, _toudtls, billGens, billlog) {
    try {
        let _billDate;
        if (_spdata != undefined)
            _billDate = _spdata.Billing_Date;
        let _billlog = billlog;
        let _dlbilllog;
        let _prepDtBl;
        let _consumerdets;
        var _query = "select * from envisage_dev.billinglog bl where bl.blmetersrno = '" + billGens.txnmtrsrno + "' and bl.bltimestamp < " + fromTime + " order by bl.bltimestamp desc limit 1";
        //logger.info("Query 17 : " + _query);
        dbConf.pool.open(dbConf.connStr, (err, conn) => {
            if (err) {
                console.log(err);
            }
            _dlbilllog = dbConf.runSQL(conn, _query);
            _query = "select * from envisage_dev.prepaydtlbillinglog prep where prep.ppdtlblmetersrno = '" + billGens.txnmtrsrno + "' and prep.ppdtlbltimestampid =" + fromTime;
            //logger.info("Query 18 : " + _query);
            _prepDtBl = dbConf.runSQL(conn, _query);
            _query = "select * from envisage_dev.mstconsumermeterrelation cn join envisage_dev.mstconsumer mc on cn.cmrconsumermasterid = mc.csmrmasterid";
            _query += " where cn.cmrconsumerid = '" + billGens.txnconsumerid + "'";
            //logger.info("Query 19 : " + _query);
            _consumerdets = dbConf.runSQL(conn, _query);
            conn.close();
        });
        if (_conld != undefined && _conld.length > 0) {
            billGens.connectedload = _conld[_conld.length - 1].cmrecurconnloadkw;
            billGens.Contractdemand = _conld[_conld.length - 1].cmrecurcontdmdkva;
        }
        else {
            billGens.connectedload = "";
            billGens.Contractdemand = "";
        }
        if (_spdata != undefined) {
            var d = new Date();
            console.log("Process starts : " + d.toLocaleDateString());
            ProcessSinglePhase(billGens, xlFilePath, _spdata, _trfdata, _prepSlab, _trfdmdslabs, _toudtls, phase, _billlog);
            d = new Date();
            console.log("Process End : " + d.toLocaleDateString());
        }
        // else if(_tpdata != undefined) {
        //     logger.info("Test T_PHASE 1");
        //     //Read xlsx file and use then fuction to handle promise before executing next step
        //     workbook.xlsx.readFile(xlFilePath).then(function () {
        //         logger.info("Test T_PHASE 2");
        //         var worksheet = workbook.getWorksheet("Bill Parameter Mapping");
        //         logger.info("Test T_PHASE 3");
        //         worksheet.getCell('D5').value = _trfdata[_trfdata.length - 1].preptrfschemedescription;
        //         logger.info("Test T_PHASE 4");
        //         worksheet.getCell('D6').value = _tpdata.Cum_Adjustable_Amount;
        //         logger.info("Test T_PHASE 5");
        //         worksheet.getCell('D7').value = "MD KVA";
        //         worksheet.getCell('D8').value = "KWh ToT";
        //         worksheet.getCell('D9').value = "No";
        //         worksheet.getCell('D10').value = "Yes";
        //         worksheet.getCell('D11').value = "No";
        //         worksheet.getCell('D12').value = _prepSlab.length;
        //         logger.info("Test T_PHASE 6");
        //         if (_prepSlab.length > 0) {
        //             logger.info("Test T_PHASE 7");
        //             worksheet.getCell('D13').value = "0";
        //             var count = _prepSlab.length;
        //             var rowVal = 14;
        //             var i = 0;
        //             logger.info("Test T_PHASE 8");
        //             for (i = 0; i < count; i++) {
        //                 var cellName = "D" + rowVal;
        //                 worksheet.getCell(cellName).value = _prepSlab[i].preptdlenrgunitsfrom;              //Energy Slab 1 : Start Reading
        //                 rowVal++;
        //                 cellName = "D" + rowVal;
        //                 worksheet.getCell(cellName).value = _prepSlab[i].preptdlenrgunitsto;              //Energy Slab 1 : End Reading
        //                 rowVal++;
        //                 cellName = "D" + rowVal;
        //                 worksheet.getCell(cellName).value = _prepSlab[i].preptdlenerchrgamt;              //Energy Slab 1 : Rate
        //                 rowVal++;
        //             }
        //             if (rowVal != 37) {
        //                 logger.info("Test T_PHASE 9");
        //                 var remaining = 37 - rowVal;
        //                 for (var k = 0; k < remaining; k++) {
        //                     var cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = "0";
        //                     rowVal++;
        //                     worksheet.getCell(cellName).value = "0";
        //                     rowVal++;
        //                     worksheet.getCell(cellName).value = "0";
        //                     rowVal++;
        //                 }
        //                 logger.info("Test T_PHASE 10");
        //             }
        //         }
        //         else {
        //             logger.info("Test T_PHASE 11");
        //             worksheet.getCell('D13').value = _trfdata[0].preptrfflatenerrate;
        //             worksheet.getCell('D14').value = "0";
        //             worksheet.getCell('D15').value = "0";
        //             worksheet.getCell('D16').value = "0";
        //             worksheet.getCell('D17').value = "0";
        //             worksheet.getCell('D18').value = "0";
        //             worksheet.getCell('D19').value = "0";
        //             worksheet.getCell('D20').value = "0";
        //             worksheet.getCell('D21').value = "0";
        //             worksheet.getCell('D22').value = "0";
        //             worksheet.getCell('D23').value = "0";
        //             worksheet.getCell('D24').value = "0";
        //             worksheet.getCell('D25').value = "0";
        //             worksheet.getCell('D26').value = "0";
        //             worksheet.getCell('D27').value = "0";
        //             worksheet.getCell('D28').value = "0";
        //             worksheet.getCell('D29').value = "0";
        //             worksheet.getCell('D30').value = "0";
        //             worksheet.getCell('D31').value = "0";
        //             worksheet.getCell('D32').value = "0";
        //             worksheet.getCell('D33').value = "0";
        //             worksheet.getCell('D34').value = "0";
        //             worksheet.getCell('D35').value = "0";
        //             worksheet.getCell('D36').value = "0";
        //             worksheet.getCell('D37').value = "0";
        //             logger.info("Test T_PHASE 12");
        //         }
        //         if (_trfdmdslabs != undefined && _trfdmdslabs.length > 0) {
        //             worksheet.getCell('D38').value = _trfdmdslabs.length;
        //             worksheet.getCell('D39').value = "0";
        //             logger.info("Test T_PHASE 13");
        //             var count = _trfdmdslabs.length;
        //             var rowVal = 40;
        //             var i = 0;
        //             if (count <= 4) {
        //                 logger.info("Test T_PHASE 14");
        //                 for (i = 0; i < count; i++) {
        //                     var cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = _trfdmdslabs[i].preptdldmdunitsfrom;                     //Demand  Slab 1 : Start Reading
        //                     rowVal++;
        //                     cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = _trfdmdslabs[i].preptdldmdunitsto;                     //Demand  Slab 1 : end Reading
        //                     rowVal++;
        //                     cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = _trfdmdslabs[i].preptdldmdchrgamt;                     //Demand  Slab 1 : Rate
        //                     rowVal++;
        //                 }
        //                 i++;
        //                 if (rowVal != 39) {
        //                     logger.info("Test T_PHASE 15");
        //                     var cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = "0";                    //Demand  Slab 1 : Start Reading
        //                     rowVal++;
        //                     cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = "0";                  //Demand  Slab 1 : end Reading
        //                     rowVal++;
        //                     cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = "0";                  //Demand  Slab 1 : Rate
        //                     rowVal++;
        //                     i++;
        //                 }
        //             }
        //         }
        //         else {
        //             logger.info("Test T_PHASE 16");
        //             worksheet.getCell('D38').value = "0";
        //             worksheet.getCell('D39').value = _trfdata[0].preptrfflatdmdrate;
        //             worksheet.getCell('D40').value = "0";                     //Demand  Slab 1 : Start Reading
        //             worksheet.getCell('D41').value = "0";                     //Demand  Slab 1 : end Reading
        //             worksheet.getCell('D42').value = "0";                     //Demand  Slab 1 : Rate
        //             worksheet.getCell('D43').value = "0";                     //Demand  Slab 2 : Start Reading
        //             worksheet.getCell('D44').value = "0";                     //Demand  Slab 2 : end Reading
        //             worksheet.getCell('D45').value = "0";                     //Demand  Slab 2 : Rate
        //             worksheet.getCell('D46').value = "0";                     //Demand  Slab 3 : Start Reading
        //             worksheet.getCell('D47').value = "0";                     //Demand  Slab 3 : end Reading
        //             worksheet.getCell('D48').value = "0";                     //Demand  Slab 3 : Rate
        //             worksheet.getCell('D49').value = "0";                     //Demand  Slab 4 : Start Reading
        //             worksheet.getCell('D50').value = "0";                     //Demand  Slab 4 : end Reading
        //             worksheet.getCell('D51').value = "0";                     //Demand  Slab 4 : Rate
        //             worksheet.getCell('D52').value = "0";
        //             worksheet.getCell('D53').value = "0";
        //             worksheet.getCell('D54').value = "0";
        //             worksheet.getCell('D55').value = "0";
        //             worksheet.getCell('D56').value = "0";
        //             worksheet.getCell('D57').value = "0";
        //             worksheet.getCell('D58').value = "0";
        //             worksheet.getCell('D59').value = "0";
        //             worksheet.getCell('D60').value = "0";
        //             worksheet.getCell('D61').value = "0";
        //             worksheet.getCell('D62').value = "0";
        //             worksheet.getCell('D63').value = "0";
        //         }
        //         if (_conld != undefined && _conld.length > 0)
        //             worksheet.getCell('D64').value = obj[0].Contractdemand;
        //         else
        //             worksheet.getCell('D64').value = "";
        //         logger.info("Test T_PHASE 17");
        //         worksheet.getCell('D65').value = "900";
        //         worksheet.getCell('D66').value = _trfdata[0].prepuntdmdchrg;
        //         worksheet.getCell('D67').value = _trfdata[0].prepuntexsdmdchrg;
        //         worksheet.getCell('D68').value = _trfdata[0].prepmthfixedchrg;
        //         worksheet.getCell('D69').value = _trfdata[0].prepdlyfixedchrg;
        //         worksheet.getCell('D70').value = _trfdata[0].prepemrgncycrlmt;
        //         worksheet.getCell('D71').value = "0";
        //         logger.info("Test T_PHASE 18");
        //         if (_toudtls != undefined && _toudtls.length > 0) {
        //             logger.info("Test T_PHASE 19");
        //             var toucount = _toudtls.length;
        //             worksheet.getCell('D72').value = toucount;
        //             var rowVal = 73;
        //             var i = 0;
        //             logger.info("Test T_PHASE 20");
        //             for (i = 0; i < toucount; i++) {
        //                 var cellName = "D" + rowVal;
        //                 worksheet.getCell(cellName).value = _toudtls[i].pretsslotstarthr + ":" + _toudtls[i].pretsslotstartmin;
        //                 rowVal++;
        //                 cellName = "D" + rowVal;
        //                 worksheet.getCell(cellName).value = _toudtls[i].pretsslotendhr + ":" + _toudtls[i].pretsslotendmin;
        //                 rowVal++;
        //                 cellName = "D" + rowVal;
        //                 worksheet.getCell(cellName).value = _toudtls[i].tsmaxdemandlimit;
        //                 rowVal++;
        //                 cellName = "D" + rowVal;
        //                 worksheet.getCell(cellName).value = _toudtls[i].pretsminchrgdmdlim;
        //                 rowVal++;
        //                 cellName = "D" + rowVal;
        //                 worksheet.getCell(cellName).value = _toudtls[i].pretsincconamt;
        //                 rowVal++;
        //                 cellName = "D" + rowVal;
        //                 worksheet.getCell(cellName).value = _toudtls[i].pretsincdmdamt;
        //                 rowVal++;
        //             }

        //             logger.info("Test T_PHASE 21");
        //             if (toucount != 8) {
        //                 logger.info("Test T_PHASE 22");
        //                 var remain = 8 - toucount;
        //                 for (var k = 0; k < remain; k++) {
        //                     var cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = "";
        //                     rowVal++;
        //                     cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = "";
        //                     rowVal++;
        //                     cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = "";
        //                     rowVal++;
        //                     cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = "";
        //                     rowVal++;
        //                     cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = "";
        //                     rowVal++;
        //                     cellName = "D" + rowVal;
        //                     worksheet.getCell(cellName).value = "";
        //                     rowVal++;
        //                 }
        //                 logger.info("Test T_PHASE 23");
        //             }
        //         }
        //         else {
        //             logger.info("Test T_PHASE 24");
        //             worksheet.getCell("D79").value = "";
        //             worksheet.getCell("D80").value = "";
        //             worksheet.getCell("D81").value = "";
        //             worksheet.getCell("D82").value = "";
        //             worksheet.getCell("D83").value = "";
        //             worksheet.getCell("D84").value = "";
        //             worksheet.getCell("D85").value = "";
        //             worksheet.getCell("D86").value = "";
        //             worksheet.getCell("D87").value = "";
        //             worksheet.getCell("D88").value = "";
        //             worksheet.getCell("D89").value = "";
        //             worksheet.getCell("D90").value = "";
        //             worksheet.getCell("D91").value = "";
        //             worksheet.getCell("D92").value = "";
        //             worksheet.getCell("D93").value = "";
        //             worksheet.getCell("D94").value = "";
        //             worksheet.getCell("D95").value = "";
        //             worksheet.getCell("D96").value = "";
        //             worksheet.getCell("D97").value = "";
        //             worksheet.getCell("D98").value = "";
        //             worksheet.getCell("D99").value = "";
        //             worksheet.getCell("D100").value = "";
        //             worksheet.getCell("D101").value = "";
        //             worksheet.getCell("D102").value = "";
        //             worksheet.getCell("D103").value = "";
        //             worksheet.getCell("D104").value = "";
        //             worksheet.getCell("D105").value = "";
        //             worksheet.getCell("D106").value = "";
        //             worksheet.getCell("D107").value = "";
        //             worksheet.getCell("D108").value = "";
        //             worksheet.getCell("D109").value = "";
        //             worksheet.getCell("D110").value = "";
        //             worksheet.getCell("D111").value = "";
        //             worksheet.getCell("D112").value = "";
        //             worksheet.getCell("D113").value = "";
        //             worksheet.getCell("D114").value = "";
        //             worksheet.getCell("D115").value = "";
        //             worksheet.getCell("D116").value = "";
        //             worksheet.getCell("D117").value = "";
        //             worksheet.getCell("D118").value = "";
        //             worksheet.getCell("D119").value = "";
        //             worksheet.getCell("D120").value = "";
        //         }
        //         worksheet.getCell("B124").value = phase[0].mtrmodelid;
        //         logger.info("Test T_PHASE 25");
        //         worksheet.getCell("D122").value = billGen.Bill_Gen.fixchargunit;
        //         worksheet.getCell("D123").value = billGen.Bill_Gen.seznsez;
        //         worksheet.getCell("H5").value = _tpdata.RTC_absolute;
        //         worksheet.getCell("H6").value = 0;
        //         worksheet.getCell("H7").value = 0;
        //         worksheet.getCell("H8").value = 0;
        //         worksheet.getCell("H9").value = 0;
        //         worksheet.getCell("H10").value = 0;
        //         worksheet.getCell("H11").value = 0;
        //         worksheet.getCell("H12").value = _tpdata.Signed_Three_phase_Power_factor;
        //         worksheet.getCell("H13").value = _tpdata.Instantaneous_Frequency;
        //         worksheet.getCell("H14").value = _tpdata.Apparent_Power;
        //         worksheet.getCell("H15").value = _tpdata.Active_Power;
        //         worksheet.getCell("H16").value = _tpdata.Signed_Reactive_power;
        //         worksheet.getCell("H17").value = "ddMMMyy";
        //         worksheet.getCell("H18").value = _tpdata.Billing_Index;
        //         worksheet.getCell("H19").value = _tpdata.Signed_Average_PowerFactor;
        //         worksheet.getCell("H20").value = _tpdata.Cumulative_Energy_KWh;
        //         worksheet.getCell("H21").value = _tpdata.Cumulative_Energy_KVAh;
        //         worksheet.getCell("H22").value = _tpdata.Cumulative_Energy_KVARh_Lag;
        //         worksheet.getCell("H23").value = _tpdata.Cumulative_Energy_KVARh_Lead;
        //         worksheet.getCell("H24").value = _tpdata.Maximum_demand_KW;
        //         worksheet.getCell("H25").value = _tpdata.Maximum_demand_KVA;
        //         worksheet.getCell("H26").value = _tpdata.Cumulative_tamper_count;
        //         logger.info("Test T_PHASE 26");
        //         worksheet.getCell("H28").value = _tpdata.LCumulative_Energy_for_TOU1;
        //         worksheet.getCell("H29").value = _tpdata.LCumulative_Energy_for_TOU2;
        //         worksheet.getCell("H30").value = _tpdata.LCumulative_Energy_for_TOU3;
        //         worksheet.getCell("H31").value = _tpdata.LCumulative_Energy_for_TOU4;
        //         worksheet.getCell("H32").value = _tpdata.LCumulative_Energy_for_TOU5;
        //         worksheet.getCell("H33").value = _tpdata.LCumulative_Energy_for_TOU6;
        //         worksheet.getCell("H34").value = _tpdata.LCumulative_Energy_for_TOU7;
        //         worksheet.getCell("H35").value = _tpdata.LCumulative_Energy_for_TOU8;
        //         logger.info("Test T_PHASE 27");
        //         worksheet.getCell("H37").value = _spdata.LCumulative_Apparent_Energy_for_TOU1;
        //         worksheet.getCell("H38").value = _spdata.LCumulative_Apparent_Energy_for_TOU2;
        //         worksheet.getCell("H39").value = _spdata.LCumulative_Apparent_Energy_for_TOU3;
        //         worksheet.getCell("H40").value = _spdata.LCumulative_Apparent_Energy_for_TOU4;
        //         worksheet.getCell("H41").value = _spdata.LCumulative_Apparent_Energy_for_TOU5;
        //         worksheet.getCell("H42").value = _spdata.LCumulative_Apparent_Energy_for_TOU6;
        //         worksheet.getCell("H43").value = _spdata.LCumulative_Apparent_Energy_for_TOU7;
        //         worksheet.getCell("H44").value = _spdata.LCumulative_Apparent_Energy_for_TOU8;
        //         logger.info("Test T_PHASE 28");
        //         worksheet.getCell("H46").value = _spdata.LMD_kW_for_TOU1;
        //         worksheet.getCell("H47").value = _spdata.LMD_kW_for_TOU2;
        //         worksheet.getCell("H48").value = _spdata.LMD_kW_for_TOU3;
        //         worksheet.getCell("H49").value = _spdata.LMD_kW_for_TOU4;
        //         worksheet.getCell("H50").value = _spdata.LMD_kW_for_TOU5;
        //         worksheet.getCell("H51").value = _spdata.LMD_kW_for_TOU6;
        //         worksheet.getCell("H52").value = _spdata.LMD_kW_for_TOU7;
        //         worksheet.getCell("H53").value = _spdata.LMD_kW_for_TOU8;
        //         logger.info("Test T_PHASE 29");
        //         worksheet.getCell("H55").value = _spdata.LMD_kVA_for_TOU1;
        //         worksheet.getCell("H56").value = _spdata.LMD_kVA_for_TOU2;
        //         worksheet.getCell("H57").value = _spdata.LMD_kVA_for_TOU3;
        //         worksheet.getCell("H58").value = _spdata.LMD_kVA_for_TOU4;
        //         worksheet.getCell("H59").value = _spdata.LMD_kVA_for_TOU5;
        //         worksheet.getCell("H60").value = _spdata.LMD_kVA_for_TOU6;
        //         worksheet.getCell("H61").value = _spdata.LMD_kVA_for_TOU7;
        //         worksheet.getCell("H62").value = _spdata.LMD_kVA_for_TOU8;
        //         logger.info("Test T_PHASE 30");
        //         console.log(_billlog[0]);
        //         if (_billlog != undefined)
        //             worksheet.getCell("H64").value = _billlog[0].blcumtampercount;
        //         else
        //             worksheet.getCell("H64").value = "0";
        //         logger.info("Test T_PHASE 31");
        //         worksheet.getCell("H65").value = _spdata.Cumulative_recharge_amount;
        //         worksheet.getCell("H66").value = _spdata.Cumulative_Balance_Deduction_Register;
        //         logger.info("Test T_PHASE 32");
        //         if (_prepDtBl != undefined)
        //             worksheet.getCell("H67").value = _prepDtBl.ppdtlblcumnumofrchrgcnt;
        //         else
        //             worksheet.getCell("H67").value = "";
        //         logger.info("Test T_PHASE 33");
        //         worksheet.getCell("H68").value = _spdata.Cumulative_Emergency_Credit;
        //         worksheet.getCell("H69").value = _spdata.Cum_Adjustable_Amount;
        //         worksheet.getCell("H70").value = _spdata.Cum_Monthly_Fixed_Charge_deduction_Amount;
        //         worksheet.getCell("H71").value = _spdata.Cum_Daily_Fixed_Charge_deduction_Amount;
        //         worksheet.getCell("H72").value = _spdata.Cum_Demand_Deduction_Amount;
        //         worksheet.getCell("H73").value = _spdata.CumEnergy_Charge_Deduction_Amount;
        //         logger.info("Test T_PHASE 34");
        //         worksheet.getCell("H75").value = 0;
        //         worksheet.getCell("H76").value = 0;
        //         worksheet.getCell("H77").value = 0;
        //         worksheet.getCell("H78").value = 0;
        //         worksheet.getCell("H79").value = 0;
        //         logger.info("Test T_PHASE 35");
        //         worksheet.getCell("H81").value = 0;
        //         worksheet.getCell("H82").value = 0;
        //         worksheet.getCell("H83").value = 0;
        //         worksheet.getCell("H84").value = 0;
        //         worksheet.getCell("H85").value = 0;
        //         worksheet.getCell("H86").value = 0;
        //         worksheet.getCell("H87").value = 0;
        //         worksheet.getCell("H88").value = 0;
        //         logger.info("Test T_PHASE 36");
        //         worksheet.getCell("H90").value = 0;
        //         worksheet.getCell("H91").value = 0;
        //         worksheet.getCell("H92").value = 0;
        //         worksheet.getCell("H93").value = 0;
        //         worksheet.getCell("H94").value = 0;
        //         worksheet.getCell("H95").value = 0;
        //         worksheet.getCell("H96").value = 0;
        //         worksheet.getCell("H97").value = 0;
        //         logger.info("Test T_PHASE 37");
        //         worksheet.getCell("H99").value = 0;
        //         worksheet.getCell("H100").value = 0;
        //         worksheet.getCell("H101").value = 0;
        //         worksheet.getCell("H102").value = 0;
        //         worksheet.getCell("H103").value = 0;
        //         worksheet.getCell("H104").value = 0;
        //         worksheet.getCell("H105").value = 0;
        //         worksheet.getCell("H106").value = 0;
        //         worksheet.getCell("H108").value = 0;
        //         worksheet.getCell("H109").value = 0;
        //         worksheet.getCell("H110").value = 0;
        //         worksheet.getCell("H111").value = 0;
        //         worksheet.getCell("H112").value = 0;
        //         worksheet.getCell("H113").value = 0;
        //         worksheet.getCell("H114").value = 0;
        //         worksheet.getCell("H115").value = 0;
        //         logger.info("Test T_PHASE 38");
        //         var conid = "";
        //         if (billGen.Bill_Gen.txnconsumerid.length > 4)
        //             conid = billGen.Bill_Gen.txnconsumerid.substring(6, 4);
        //         else
        //             conid = billGen.Bill_Gen.txnconsumerid;
        //         logger.info("Test T_PHASE 39");
        //         worksheet.getCell("L5").value = conid;
        //         worksheet.getCell("L6").value = billGen.Bill_Gen.billMonth + "" + billGen.Bill_Gen.billYear;
        //         logger.info("Test T_PHASE 40");
        //         if (_consumerdets != undefined) {
        //             logger.info("Test T_PHASE 41");
        //             worksheet.getCell("L7").value = _consumerdets[0].csmrfirstname + " " + _consumerdets[0].csmrlastname;
        //             worksheet.getCell("L8").value = _consumerdets[0].csmraddress1;
        //         }
        //         else {
        //             logger.info("Test 42");
        //             worksheet.getCell("L7").value = "";
        //             worksheet.getCell("L8").value = "";
        //         }
        //         logger.info("Test T_PHASE 43");
        //         worksheet.getCell("L9").value = _trfdata[0].preptaxrchrg + " %";
        //         worksheet.getCell("L10").value = billGen.Bill_Gen.gstnumber;
        //         worksheet.getCell("L11").value = billGen.Bill_Gen.connectedload;
        //         worksheet.getCell("L12").value = billGen.Bill_Gen.billnumber;
        //         worksheet.getCell("L13").value = billGen.Bill_Gen.connecteddate;
        //         worksheet.getCell("L14").value = _trfdata[0].prepfuelsuchrg;
        //         worksheet.getCell("L15").value = _trfdata[0].prepfixcharge;
        //         worksheet.getCell("L16").value = _trfdata[0].prepgstper + " %";
        //         worksheet.getCell("L17").value = _trfdata[0].preplowvoltsuchrg;
        //         worksheet.getCell("L19").value = _trfdata[0].prepmtrhire;
        //         logger.info("Test T_PHASE 44");
        //         if (_billlog == undefined || _billlog == null) {
        //             logger.info("Test T_PHASE 45");
        //             worksheet.getCell("L20").value = "0";
        //             worksheet.getCell("L21").value = "0";
        //             worksheet.getCell("L22").value = "0";
        //             worksheet.getCell("L23").value = "0";
        //             worksheet.getCell("L24").value = "0";
        //             worksheet.getCell("L25").value = "0";
        //             worksheet.getCell("L26").value = "0";
        //             worksheet.getCell("L28").value = "0";
        //             worksheet.getCell("L29").value = "0";
        //             worksheet.getCell("L30").value = "0";
        //             worksheet.getCell("L31").value = "0";
        //             worksheet.getCell("L32").value = "0";
        //             worksheet.getCell("L33").value = "0";
        //             worksheet.getCell("L34").value = "0";
        //             worksheet.getCell("L35").value = "0";
        //             worksheet.getCell("L37").value = "0";
        //             worksheet.getCell("L38").value = "0";
        //             worksheet.getCell("L39").value = "0";
        //             worksheet.getCell("L40").value = "0";
        //             worksheet.getCell("L41").value = "0";
        //             worksheet.getCell("L42").value = "0";
        //             worksheet.getCell("L43").value = "0";
        //             worksheet.getCell("L44").value = "0";
        //             worksheet.getCell("L46").value = "0";
        //             worksheet.getCell("L47").value = "0";
        //             worksheet.getCell("L48").value = "0";
        //             worksheet.getCell("L49").value = "0";
        //             worksheet.getCell("L50").value = "0";
        //             worksheet.getCell("L51").value = "0";
        //             worksheet.getCell("L52").value = "0";
        //             worksheet.getCell("L53").value = "0";
        //             worksheet.getCell("L55").value = "0";
        //             worksheet.getCell("L56").value = "0";
        //             worksheet.getCell("L57").value = "0";
        //             worksheet.getCell("L58").value = "0";
        //             worksheet.getCell("L59").value = "0";
        //             worksheet.getCell("L60").value = "0";
        //             worksheet.getCell("L61").value = "0";
        //             worksheet.getCell("L62").value = "0";
        //         }
        //         else {
        //             logger.info("Test T_PHASE 45");
        //             worksheet.getCell("L20").value = _billlog[0].blcumkwh;
        //             worksheet.getCell("L21").value = _billlog[0].blcumkvah;
        //             worksheet.getCell("L22").value = _billlog[0].blcumkvarh_lag;
        //             worksheet.getCell("L23").value = _billlog[0].blcumkvarh_lead;
        //             worksheet.getCell("L24").value = _billlog[0].blmdkw;
        //             worksheet.getCell("L25").value = _billlog[0].blmdkva;
        //             worksheet.getCell("L26").value = _billlog[0].blcumtampercount;
        //             worksheet.getCell("L28").value = _billlog[0].blcumkwhtou1;
        //             worksheet.getCell("L29").value = _billlog[0].blcumkwhtou2;
        //             worksheet.getCell("L30").value = _billlog[0].blcumkwhtou3;
        //             worksheet.getCell("L31").value = _billlog[0].blcumkwhtou4;
        //             worksheet.getCell("L32").value = _billlog[0].blcumkwhtou5;
        //             worksheet.getCell("L33").value = _billlog[0].blcumkwhtou6;
        //             worksheet.getCell("L34").value = _billlog[0].blcumkwhtou7;
        //             worksheet.getCell("L35").value = _billlog[0].blcumkwhtou8;
        //             worksheet.getCell("L37").value = _billlog[0].blcumkvahtou1;
        //             worksheet.getCell("L38").value = _billlog[0].blcumkvahtou2;
        //             worksheet.getCell("L39").value = _billlog[0].blcumkvahtou3;
        //             worksheet.getCell("L40").value = _billlog[0].blcumkvahtou4;
        //             worksheet.getCell("L41").value = _billlog[0].blcumkvahtou5;
        //             worksheet.getCell("L42").value = _billlog[0].blcumkvahtou6;
        //             worksheet.getCell("L43").value = _billlog[0].blcumkvahtou7;
        //             worksheet.getCell("L44").value = _billlog[0].blcumkvahtou8;
        //             worksheet.getCell("L46").value = _billlog[0].blmdkwtou1;
        //             worksheet.getCell("L47").value = _billlog[0].blmdkwtou2;
        //             worksheet.getCell("L48").value = _billlog[0].blmdkwtou3;
        //             worksheet.getCell("L49").value = _billlog[0].blmdkwtou4;
        //             worksheet.getCell("L50").value = _billlog[0].blmdkwtou5;
        //             worksheet.getCell("L51").value = _billlog[0].blmdkwtou6;
        //             worksheet.getCell("L52").value = _billlog[0].blmdkwtou7;
        //             worksheet.getCell("L53").value = _billlog[0].blmdkwtou8;
        //             worksheet.getCell("L55").value = _billlog[0].blmdkvatou1;
        //             worksheet.getCell("L56").value = _billlog[0].blmdkvatou2;
        //             worksheet.getCell("L57").value = _billlog[0].blmdkvatou3;
        //             worksheet.getCell("L58").value = _billlog[0].blmdkvatou4;
        //             worksheet.getCell("L59").value = _billlog[0].blmdkvatou5;
        //             worksheet.getCell("L60").value = _billlog[0].blmdkvatou6;
        //             worksheet.getCell("L61").value = _billlog[0].blmdkvatou7;
        //             worksheet.getCell("L62").value = _billlog[0].blmdkvatou8;

        //             worksheet.getCell("L75").value = _billlog[0].blexpkvah;
        //             worksheet.getCell("L76").value = _billlog[0].blexpkvarhlagq2;
        //             worksheet.getCell("L77").value = _billlog[0].blexpkvarhlegq3;
        //             worksheet.getCell("L78").value = _billlog[0].blexpmdkw;
        //             worksheet.getCell("L79").value = _billlog[0].blexpmdkva;
        //             logger.info("Test T_PHASE 46");
        //             worksheet.getCell("L81").value = _billlog[0].bltouexpkwh0;
        //             worksheet.getCell("L82").value = _billlog[0].bltouexpkwh1;
        //             worksheet.getCell("L83").value = _billlog[0].bltouexpkwh2;
        //             worksheet.getCell("L84").value = _billlog[0].bltouexpkwh3;
        //             worksheet.getCell("L85").value = _billlog[0].bltouexpkwh4;
        //             worksheet.getCell("L86").value = _billlog[0].bltouexpkwh5;
        //             worksheet.getCell("L87").value = _billlog[0].bltouexpkwh6;
        //             worksheet.getCell("L88").value = _billlog[0].bltouexpkwh7;
        //             logger.info("Test T_PHASE 47");
        //             worksheet.getCell("L90").value = _billlog[0].bltouexpkvah0;
        //             worksheet.getCell("L91").value = _billlog[0].bltouexpkvah1;
        //             worksheet.getCell("L92").value = _billlog[0].bltouexpkvah2;
        //             worksheet.getCell("L93").value = _billlog[0].bltouexpkvah3;
        //             worksheet.getCell("L94").value = _billlog[0].bltouexpkvah4;
        //             worksheet.getCell("L95").value = _billlog[0].bltouexpkvah5;
        //             worksheet.getCell("L96").value = _billlog[0].bltouexpkvah6;
        //             worksheet.getCell("L97").value = _billlog[0].bltouexpkvah7;
        //             logger.info("Test T_PHASE 48");
        //             worksheet.getCell("L99").value = _billlog[0].bltouexpmdkw0;
        //             worksheet.getCell("L100").value = _billlog[0].bltouexpmdkw1;
        //             worksheet.getCell("L101").value = _billlog[0].bltouexpmdkw2;
        //             worksheet.getCell("L102").value = _billlog[0].bltouexpmdkw3;
        //             worksheet.getCell("L103").value = _billlog[0].bltouexpmdkw4;
        //             worksheet.getCell("L104").value = _billlog[0].bltouexpmdkw5;
        //             worksheet.getCell("L105").value = _billlog[0].bltouexpmdkw6;
        //             worksheet.getCell("L106").value = _billlog[0].bltouexpmdkw7;
        //             logger.info("Test T_PHASE 49");
        //             worksheet.getCell("L108").value = _billlog[0].bltouexpmdkva0;
        //             worksheet.getCell("L109").value = _billlog[0].bltouexpmdkva1;
        //             worksheet.getCell("L110").value = _billlog[0].bltouexpmdkva2;
        //             worksheet.getCell("L111").value = _billlog[0].bltouexpmdkva3;
        //             worksheet.getCell("L112").value = _billlog[0].bltouexpmdkva4;
        //             worksheet.getCell("L113").value = _billlog[0].bltouexpmdkva5;
        //             worksheet.getCell("L114").value = _billlog[0].bltouexpmdkva6;
        //             worksheet.getCell("L115").value = _billlog[0].bltouexpmdkva7;
        //             logger.info("Test T_PHASE 50");
        //         }
        //         console.log("Excel Process Complete");
        //         workbook.xlsx.writeFile(xlFilePath);
        //         console.log("Single Phase Bill Complete : " + billGen.Bill_Gen.txnconsumerid);
        //     });
        // }
    }
    catch (err) {
        logger.info(err);
        console.log(err);
    }

}
//phase --- contains all meter info from mstmeter,mstmetermodel etc.

function epochToJsDate(ts) {
    return new Date(ts * 1000);
}


async function ProcessSinglePhase(billGens, xlFilePath, _spdata, _trfdata, _prepSlab, _trfdmdslabs, _toudtls, phase, _billlog) {
    logger.info("Test 1");
    var workbook = new Excel.Workbook();
    //Read xlsx file and use then fuction to handle promise before executing next step
    await workbook.xlsx.readFile(xlFilePath).then(function () {
        logger.info(xlFilePath);
        logger.info("Test 2");
        var worksheet = workbook.getWorksheet("Bill Parameter Mapping");
        const v0 = worksheet.getCell('D5').value;
        logger.info("Test 3");
        worksheet.getCell('D5').value = _trfdata[_trfdata.length - 1].preptrfschemedescription;
        logger.info("Test 4");
        worksheet.getCell('D6').value = _spdata.Cum_Adjustable_Amount;
        logger.info("Test 5");
        worksheet.getCell('D7').value = "MD KVA";
        worksheet.getCell('D8').value = "KWh ToT";
        worksheet.getCell('D9').value = "No";
        worksheet.getCell('D10').value = "Yes";
        worksheet.getCell('D11').value = "No";
        worksheet.getCell('D12').value = _prepSlab.length;
        logger.info("Test 6");
        if (_prepSlab.length > 0) {
            logger.info("Test 7");
            worksheet.getCell('D13').value = "0";
            var count = _prepSlab.length;
            var rowVal = 14;
            var i = 0;
            logger.info("Test 8");
            for (i = 0; i < count; i++) {
                var cellName = "D" + rowVal;
                worksheet.getCell(cellName).value = _prepSlab[i].preptdlenrgunitsfrom;              //Energy Slab 1 : Start Reading
                rowVal++;
                cellName = "D" + rowVal;
                worksheet.getCell(cellName).value = _prepSlab[i].preptdlenrgunitsto;              //Energy Slab 1 : End Reading
                rowVal++;
                cellName = "D" + rowVal;
                worksheet.getCell(cellName).value = _prepSlab[i].preptdlenerchrgamt;              //Energy Slab 1 : Rate
                rowVal++;
            }
            if (rowVal != 37) {
                logger.info("Test 9");
                var remaining = 37 - rowVal;
                for (var k = 0; k < remaining; k++) {
                    var cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = "0";
                    rowVal++;
                    worksheet.getCell(cellName).value = "0";
                    rowVal++;
                    worksheet.getCell(cellName).value = "0";
                    rowVal++;
                }
                logger.info("Test 10");
            }
        }
        else {
            logger.info("Test 11");
            worksheet.getCell('D13').value = _trfdata[0].preptrfflatenerrate;
            worksheet.getCell('D14').value = "0";
            worksheet.getCell('D15').value = "0";
            worksheet.getCell('D16').value = "0";
            worksheet.getCell('D17').value = "0";
            worksheet.getCell('D18').value = "0";
            worksheet.getCell('D19').value = "0";
            worksheet.getCell('D20').value = "0";
            worksheet.getCell('D21').value = "0";
            worksheet.getCell('D22').value = "0";
            worksheet.getCell('D23').value = "0";
            worksheet.getCell('D24').value = "0";
            worksheet.getCell('D25').value = "0";
            worksheet.getCell('D26').value = "0";
            worksheet.getCell('D27').value = "0";
            worksheet.getCell('D28').value = "0";
            worksheet.getCell('D29').value = "0";
            worksheet.getCell('D30').value = "0";
            worksheet.getCell('D31').value = "0";
            worksheet.getCell('D32').value = "0";
            worksheet.getCell('D33').value = "0";
            worksheet.getCell('D34').value = "0";
            worksheet.getCell('D35').value = "0";
            worksheet.getCell('D36').value = "0";
            worksheet.getCell('D37').value = "0";
            logger.info("Test 12");
        }
        if (_trfdmdslabs != undefined && _trfdmdslabs.length > 0) {
            worksheet.getCell('D38').value = _trfdmdslabs.length;
            worksheet.getCell('D39').value = "0";
            logger.info("Test 13");
            var count = _trfdmdslabs.length;
            var rowVal = 40;
            var i = 0;
            if (count <= 4) {
                logger.info("Test 14");
                for (i = 0; i < count; i++) {
                    var cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = _trfdmdslabs[i].preptdldmdunitsfrom;                     //Demand  Slab 1 : Start Reading
                    rowVal++;
                    cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = _trfdmdslabs[i].preptdldmdunitsto;                     //Demand  Slab 1 : end Reading
                    rowVal++;
                    cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = _trfdmdslabs[i].preptdldmdchrgamt;                     //Demand  Slab 1 : Rate
                    rowVal++;
                }
                i++;
                if (rowVal != 39) {
                    logger.info("Test 15");
                    var cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = "0";                    //Demand  Slab 1 : Start Reading
                    rowVal++;
                    cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = "0";                  //Demand  Slab 1 : end Reading
                    rowVal++;
                    cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = "0";                  //Demand  Slab 1 : Rate
                    rowVal++;
                    i++;
                }
            }
        }
        else {
            logger.info("Test 16");
            worksheet.getCell('D38').value = "0";
            worksheet.getCell('D39').value = _trfdata[0].preptrfflatdmdrate;
            worksheet.getCell('D40').value = "0";                     //Demand  Slab 1 : Start Reading
            worksheet.getCell('D41').value = "0";                     //Demand  Slab 1 : end Reading
            worksheet.getCell('D42').value = "0";                     //Demand  Slab 1 : Rate
            worksheet.getCell('D43').value = "0";                     //Demand  Slab 2 : Start Reading
            worksheet.getCell('D44').value = "0";                     //Demand  Slab 2 : end Reading
            worksheet.getCell('D45').value = "0";                     //Demand  Slab 2 : Rate
            worksheet.getCell('D46').value = "0";                     //Demand  Slab 3 : Start Reading
            worksheet.getCell('D47').value = "0";                     //Demand  Slab 3 : end Reading
            worksheet.getCell('D48').value = "0";                     //Demand  Slab 3 : Rate
            worksheet.getCell('D49').value = "0";                     //Demand  Slab 4 : Start Reading
            worksheet.getCell('D50').value = "0";                     //Demand  Slab 4 : end Reading
            worksheet.getCell('D51').value = "0";                     //Demand  Slab 4 : Rate
            worksheet.getCell('D52').value = "0";
            worksheet.getCell('D53').value = "0";
            worksheet.getCell('D54').value = "0";
            worksheet.getCell('D55').value = "0";
            worksheet.getCell('D56').value = "0";
            worksheet.getCell('D57').value = "0";
            worksheet.getCell('D58').value = "0";
            worksheet.getCell('D59').value = "0";
            worksheet.getCell('D60').value = "0";
            worksheet.getCell('D61').value = "0";
            worksheet.getCell('D62').value = "0";
            worksheet.getCell('D63').value = "0";
        }
        if (_conld != undefined && _conld.length > 0)
            worksheet.getCell('D64').value = billGens.Contractdemand;
        else
            worksheet.getCell('D64').value = "0";
        logger.info("Test 17");
        worksheet.getCell('D65').value = "900";
        worksheet.getCell('D66').value = _trfdata[0].prepuntdmdchrg;
        worksheet.getCell('D67').value = _trfdata[0].prepuntexsdmdchrg;
        worksheet.getCell('D68').value = _trfdata[0].prepmthfixedchrg;
        worksheet.getCell('D69').value = _trfdata[0].prepdlyfixedchrg;
        worksheet.getCell('D70').value = _trfdata[0].prepemrgncycrlmt;
        worksheet.getCell('D71').value = "0";
        logger.info("Test 18");
        if (_toudtls != undefined && _toudtls.length > 0) {
            logger.info("Test 19");
            var toucount = _toudtls.length;
            worksheet.getCell('D72').value = toucount;
            var rowVal = 73;
            var i = 0;
            logger.info("Test 20");
            for (i = 0; i < toucount; i++) {
                var cellName = "D" + rowVal;
                worksheet.getCell(cellName).value = _toudtls[i].pretsslotstarthr + ":" + _toudtls[i].pretsslotstartmin;
                rowVal++;
                cellName = "D" + rowVal;
                worksheet.getCell(cellName).value = _toudtls[i].pretsslotendhr + ":" + _toudtls[i].pretsslotendmin;
                rowVal++;
                cellName = "D" + rowVal;
                worksheet.getCell(cellName).value = _toudtls[i].tsmaxdemandlimit;
                rowVal++;
                cellName = "D" + rowVal;
                worksheet.getCell(cellName).value = _toudtls[i].pretsminchrgdmdlim;
                rowVal++;
                cellName = "D" + rowVal;
                worksheet.getCell(cellName).value = _toudtls[i].pretsincconamt;
                rowVal++;
                cellName = "D" + rowVal;
                worksheet.getCell(cellName).value = _toudtls[i].pretsincdmdamt;
                rowVal++;
            }

            logger.info("Test 21");
            if (toucount != 8) {
                logger.info("Test 22");
                var remain = 8 - toucount;
                for (var k = 0; k < remain; k++) {
                    var cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = "";
                    rowVal++;
                    cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = "";
                    rowVal++;
                    cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = "";
                    rowVal++;
                    cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = "";
                    rowVal++;
                    cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = "";
                    rowVal++;
                    cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = "";
                    rowVal++;
                }
                logger.info("Test 23");
            }
        }
        else {
            logger.info("Test 24");
            worksheet.getCell("D79").value = "";
            worksheet.getCell("D80").value = "";
            worksheet.getCell("D81").value = "";
            worksheet.getCell("D82").value = "";
            worksheet.getCell("D83").value = "";
            worksheet.getCell("D84").value = "";
            worksheet.getCell("D85").value = "";
            worksheet.getCell("D86").value = "";
            worksheet.getCell("D87").value = "";
            worksheet.getCell("D88").value = "";
            worksheet.getCell("D89").value = "";
            worksheet.getCell("D90").value = "";
            worksheet.getCell("D91").value = "";
            worksheet.getCell("D92").value = "";
            worksheet.getCell("D93").value = "";
            worksheet.getCell("D94").value = "";
            worksheet.getCell("D95").value = "";
            worksheet.getCell("D96").value = "";
            worksheet.getCell("D97").value = "";
            worksheet.getCell("D98").value = "";
            worksheet.getCell("D99").value = "";
            worksheet.getCell("D100").value = "";
            worksheet.getCell("D101").value = "";
            worksheet.getCell("D102").value = "";
            worksheet.getCell("D103").value = "";
            worksheet.getCell("D104").value = "";
            worksheet.getCell("D105").value = "";
            worksheet.getCell("D106").value = "";
            worksheet.getCell("D107").value = "";
            worksheet.getCell("D108").value = "";
            worksheet.getCell("D109").value = "";
            worksheet.getCell("D110").value = "";
            worksheet.getCell("D111").value = "";
            worksheet.getCell("D112").value = "";
            worksheet.getCell("D113").value = "";
            worksheet.getCell("D114").value = "";
            worksheet.getCell("D115").value = "";
            worksheet.getCell("D116").value = "";
            worksheet.getCell("D117").value = "";
            worksheet.getCell("D118").value = "";
            worksheet.getCell("D119").value = "";
            worksheet.getCell("D120").value = "";
            worksheet.getCell("D124").value = phase[0].mtrmodelid;
        }
        logger.info("Test 25");
        worksheet.getCell("D122").value = billGens.fixchargunit;
        worksheet.getCell("D123").value = billGens.seznsez;
        worksheet.getCell("H5").value = _spdata.RTC_absolute;
        worksheet.getCell("H6").value = 0;
        worksheet.getCell("H7").value = 0;
        worksheet.getCell("H8").value = 0;
        worksheet.getCell("H9").value = 0;
        worksheet.getCell("H10").value = 0;
        worksheet.getCell("H11").value = 0;
        worksheet.getCell("H12").value = _spdata.Signed_Three_phase_Power_factor;
        worksheet.getCell("H13").value = _spdata.Instantaneous_Frequency;
        worksheet.getCell("H14").value = _spdata.Apparent_Power;
        worksheet.getCell("H15").value = _spdata.Active_Power;
        worksheet.getCell("H16").value = _spdata.Signed_Reactive_power;
        var _date = epochToJsDate(_spdata.Billing_Date);
        var day = _date.getDate();
        //var month = _date.toLocaleString('default', { month: 'short' });
        var month = _date.getMonth() + 1;
        var year = _date.getFullYear();
        var datestring = month + "/" + day + "/" + year;
        logger.info(datestring);
        worksheet.getCell("H17").value = datestring;
        worksheet.getCell("H18").value = _spdata.Billing_Index;
        worksheet.getCell("H19").value = _spdata.Signed_Average_PowerFactor;
        worksheet.getCell("H20").value = _spdata.Cumulative_Energy_KWh;
        worksheet.getCell("H21").value = _spdata.Cumulative_Energy_KVAh;
        worksheet.getCell("H22").value = _spdata.Cumulative_Energy_KVARh_Lag;
        worksheet.getCell("H23").value = _spdata.Cumulative_Energy_KVARh_Lead;
        worksheet.getCell("H24").value = _spdata.Maximum_demand_KW;
        worksheet.getCell("H25").value = _spdata.Maximum_demand_KVA;
        worksheet.getCell("H26").value = _spdata.Cumulative_tamper_count;
        logger.info("Test 26");
        worksheet.getCell("H28").value = _spdata.LCumulative_Energy_for_TOU1;
        worksheet.getCell("H29").value = _spdata.LCumulative_Energy_for_TOU2;
        worksheet.getCell("H30").value = _spdata.LCumulative_Energy_for_TOU3;
        worksheet.getCell("H31").value = _spdata.LCumulative_Energy_for_TOU4;
        worksheet.getCell("H32").value = _spdata.LCumulative_Energy_for_TOU5;
        worksheet.getCell("H33").value = _spdata.LCumulative_Energy_for_TOU6;
        worksheet.getCell("H34").value = _spdata.LCumulative_Energy_for_TOU7;
        worksheet.getCell("H35").value = _spdata.LCumulative_Energy_for_TOU8;
        logger.info("Test 27");
        worksheet.getCell("H37").value = _spdata.LCumulative_Apparent_Energy_for_TOU1;
        worksheet.getCell("H38").value = _spdata.LCumulative_Apparent_Energy_for_TOU2;
        worksheet.getCell("H39").value = _spdata.LCumulative_Apparent_Energy_for_TOU3;
        worksheet.getCell("H40").value = _spdata.LCumulative_Apparent_Energy_for_TOU4;
        worksheet.getCell("H41").value = _spdata.LCumulative_Apparent_Energy_for_TOU5;
        worksheet.getCell("H42").value = _spdata.LCumulative_Apparent_Energy_for_TOU6;
        worksheet.getCell("H43").value = _spdata.LCumulative_Apparent_Energy_for_TOU7;
        worksheet.getCell("H44").value = _spdata.LCumulative_Apparent_Energy_for_TOU8;
        logger.info("Test 28");
        worksheet.getCell("H46").value = _spdata.LMD_kW_for_TOU1;
        worksheet.getCell("H47").value = _spdata.LMD_kW_for_TOU2;
        worksheet.getCell("H48").value = _spdata.LMD_kW_for_TOU3;
        worksheet.getCell("H49").value = _spdata.LMD_kW_for_TOU4;
        worksheet.getCell("H50").value = _spdata.LMD_kW_for_TOU5;
        worksheet.getCell("H51").value = _spdata.LMD_kW_for_TOU6;
        worksheet.getCell("H52").value = _spdata.LMD_kW_for_TOU7;
        worksheet.getCell("H53").value = _spdata.LMD_kW_for_TOU8;
        logger.info("Test 29");
        worksheet.getCell("H55").value = _spdata.LMD_kVA_for_TOU1;
        worksheet.getCell("H56").value = _spdata.LMD_kVA_for_TOU2;
        worksheet.getCell("H57").value = _spdata.LMD_kVA_for_TOU3;
        worksheet.getCell("H58").value = _spdata.LMD_kVA_for_TOU4;
        worksheet.getCell("H59").value = _spdata.LMD_kVA_for_TOU5;
        worksheet.getCell("H60").value = _spdata.LMD_kVA_for_TOU6;
        worksheet.getCell("H61").value = _spdata.LMD_kVA_for_TOU7;
        worksheet.getCell("H62").value = _spdata.LMD_kVA_for_TOU8;
        logger.info("Test 30");
        if (_billlog != undefined)
            worksheet.getCell("H64").value = _spdata.Cumulative_tamper_count;
        else
            worksheet.getCell("H64").value = "0";
        logger.info("Test 31");
        worksheet.getCell("H65").value = _spdata.Cumulative_recharge_amount;
        worksheet.getCell("H66").value = _spdata.Cumulative_Balance_Deduction_Register;
        logger.info("Test 32");
        if (_prepDtBl != undefined)
            worksheet.getCell("H67").value = _prepDtBl.ppdtlblcumnumofrchrgcnt;
        else
            worksheet.getCell("H67").value = "";
        logger.info("Test 33");
        worksheet.getCell("H68").value = _spdata.Cumulative_Emergency_Credit;
        worksheet.getCell("H69").value = _spdata.Cum_Adjustable_Amount;
        worksheet.getCell("H70").value = _spdata.Cum_Monthly_Fixed_Charge_deduction_Amount;
        worksheet.getCell("H71").value = _spdata.Cum_Daily_Fixed_Charge_deduction_Amount;
        worksheet.getCell("H72").value = _spdata.Cum_Demand_Deduction_Amount;
        worksheet.getCell("H73").value = _spdata.CumEnergy_Charge_Deduction_Amount;
        logger.info("Test 34");
        worksheet.getCell("H75").value = 0;
        worksheet.getCell("H76").value = 0;
        worksheet.getCell("H77").value = 0;
        worksheet.getCell("H78").value = 0;
        worksheet.getCell("H79").value = 0;
        logger.info("Test 35");
        worksheet.getCell("H81").value = 0;
        worksheet.getCell("H82").value = 0;
        worksheet.getCell("H83").value = 0;
        worksheet.getCell("H84").value = 0;
        worksheet.getCell("H85").value = 0;
        worksheet.getCell("H86").value = 0;
        worksheet.getCell("H87").value = 0;
        worksheet.getCell("H88").value = 0;
        logger.info("Test 36");
        worksheet.getCell("H90").value = 0;
        worksheet.getCell("H91").value = 0;
        worksheet.getCell("H92").value = 0;
        worksheet.getCell("H93").value = 0;
        worksheet.getCell("H94").value = 0;
        worksheet.getCell("H95").value = 0;
        worksheet.getCell("H96").value = 0;
        worksheet.getCell("H97").value = 0;
        logger.info("Test 37");
        worksheet.getCell("H99").value = 0;
        worksheet.getCell("H100").value = 0;
        worksheet.getCell("H101").value = 0;
        worksheet.getCell("H102").value = 0;
        worksheet.getCell("H103").value = 0;
        worksheet.getCell("H104").value = 0;
        worksheet.getCell("H105").value = 0;
        worksheet.getCell("H106").value = 0;
        worksheet.getCell("H108").value = 0;
        worksheet.getCell("H109").value = 0;
        worksheet.getCell("H110").value = 0;
        worksheet.getCell("H111").value = 0;
        worksheet.getCell("H112").value = 0;
        worksheet.getCell("H113").value = 0;
        worksheet.getCell("H114").value = 0;
        worksheet.getCell("H115").value = 0;
        logger.info("Test 38");
        var conid = "";
        if (billGens.txnconsumerid.length > 4)
            conid = billGens.txnconsumerid.substring(6, 4);
        else
            conid = billGens.txnconsumerid;
        logger.info("Test 39");
        console.log("Consumer : " + conid);
        worksheet.getCell("L5").value = conid;
        worksheet.getCell("L6").value = billGens.billMonth + "" + billGens.billYear;
        logger.info("Test 40");
        if (_consumerdets != undefined) {
            logger.info("Test 41");
            worksheet.getCell("L7").value = _consumerdets[0].csmrfirstname + " " + _consumerdets[0].csmrlastname;
            worksheet.getCell("L8").value = _consumerdets[0].csmraddress1;
        }
        else {
            logger.info("Test 42");
            worksheet.getCell("L7").value = "";
            worksheet.getCell("L8").value = "";
        }
        logger.info("Test 43");
        worksheet.getCell("L9").value = _trfdata[0].preptaxrchrg + " %";
        worksheet.getCell("L10").value = billGens.gstnumber;
        worksheet.getCell("L11").value = billGens.connectedload;
        worksheet.getCell("L12").value = billGens.billnumber;
        worksheet.getCell("L13").value = billGens.connecteddate;
        worksheet.getCell("L14").value = _trfdata[0].prepfuelsuchrg;
        worksheet.getCell("L15").value = _trfdata[0].prepfixcharge;
        worksheet.getCell("L16").value = _trfdata[0].prepgstper + " %";
        worksheet.getCell("L17").value = _trfdata[0].preplowvoltsuchrg;
        worksheet.getCell("L19").value = _trfdata[0].prepmtrhire;
        logger.info("Test 44");
        if (_billlog == undefined || _billlog == null) {
            logger.info("Test 45");
            worksheet.getCell("L20").value = "0";
            worksheet.getCell("L21").value = "0";
            worksheet.getCell("L22").value = "0";
            worksheet.getCell("L23").value = "0";
            worksheet.getCell("L24").value = "0";
            worksheet.getCell("L25").value = "0";
            worksheet.getCell("L26").value = "0";
            worksheet.getCell("L28").value = "0";
            worksheet.getCell("L29").value = "0";
            worksheet.getCell("L30").value = "0";
            worksheet.getCell("L31").value = "0";
            worksheet.getCell("L32").value = "0";
            worksheet.getCell("L33").value = "0";
            worksheet.getCell("L34").value = "0";
            worksheet.getCell("L35").value = "0";
            worksheet.getCell("L37").value = "0";
            worksheet.getCell("L38").value = "0";
            worksheet.getCell("L39").value = "0";
            worksheet.getCell("L40").value = "0";
            worksheet.getCell("L41").value = "0";
            worksheet.getCell("L42").value = "0";
            worksheet.getCell("L43").value = "0";
            worksheet.getCell("L44").value = "0";
            worksheet.getCell("L46").value = "0";
            worksheet.getCell("L47").value = "0";
            worksheet.getCell("L48").value = "0";
            worksheet.getCell("L49").value = "0";
            worksheet.getCell("L50").value = "0";
            worksheet.getCell("L51").value = "0";
            worksheet.getCell("L52").value = "0";
            worksheet.getCell("L53").value = "0";
            worksheet.getCell("L55").value = "0";
            worksheet.getCell("L56").value = "0";
            worksheet.getCell("L57").value = "0";
            worksheet.getCell("L58").value = "0";
            worksheet.getCell("L59").value = "0";
            worksheet.getCell("L60").value = "0";
            worksheet.getCell("L61").value = "0";
            worksheet.getCell("L62").value = "0";
        }
        else {
            logger.info("Test 45");
            worksheet.getCell("L20").value = _billlog.blcumkwh;
            worksheet.getCell("L21").value = _billlog.blcumkvah;
            worksheet.getCell("L22").value = _billlog.blcumkvarh_lag;
            worksheet.getCell("L23").value = _billlog.blcumkvarh_lead;
            worksheet.getCell("L24").value = _billlog.blmdkw;
            worksheet.getCell("L25").value = _billlog.blmdkva;
            worksheet.getCell("L26").value = _billlog.blcumtampercount;
            worksheet.getCell("L28").value = _billlog.blcumkwhtou1;
            worksheet.getCell("L29").value = _billlog.blcumkwhtou2;
            worksheet.getCell("L30").value = _billlog.blcumkwhtou3;
            worksheet.getCell("L31").value = _billlog.blcumkwhtou4;
            worksheet.getCell("L32").value = _billlog.blcumkwhtou5;
            worksheet.getCell("L33").value = _billlog.blcumkwhtou6;
            worksheet.getCell("L34").value = _billlog.blcumkwhtou7;
            worksheet.getCell("L35").value = _billlog.blcumkwhtou8;
            worksheet.getCell("L37").value = _billlog.blcumkvahtou1;
            worksheet.getCell("L38").value = _billlog.blcumkvahtou2;
            worksheet.getCell("L39").value = _billlog.blcumkvahtou3;
            worksheet.getCell("L40").value = _billlog.blcumkvahtou4;
            worksheet.getCell("L41").value = _billlog.blcumkvahtou5;
            worksheet.getCell("L42").value = _billlog.blcumkvahtou6;
            worksheet.getCell("L43").value = _billlog.blcumkvahtou7;
            worksheet.getCell("L44").value = _billlog.blcumkvahtou8;
            worksheet.getCell("L46").value = _billlog.blmdkwtou1;
            worksheet.getCell("L47").value = _billlog.blmdkwtou2;
            worksheet.getCell("L48").value = _billlog.blmdkwtou3;
            worksheet.getCell("L49").value = _billlog.blmdkwtou4;
            worksheet.getCell("L50").value = _billlog.blmdkwtou5;
            worksheet.getCell("L51").value = _billlog.blmdkwtou6;
            worksheet.getCell("L52").value = _billlog.blmdkwtou7;
            worksheet.getCell("L53").value = _billlog.blmdkwtou8;
            worksheet.getCell("L55").value = _billlog.blmdkvatou1;
            worksheet.getCell("L56").value = _billlog.blmdkvatou2;
            worksheet.getCell("L57").value = _billlog.blmdkvatou3;
            worksheet.getCell("L58").value = _billlog.blmdkvatou4;
            worksheet.getCell("L59").value = _billlog.blmdkvatou5;
            worksheet.getCell("L60").value = _billlog.blmdkvatou6;
            worksheet.getCell("L61").value = _billlog.blmdkvatou7;
            worksheet.getCell("L62").value = _billlog.blmdkvatou8;

            worksheet.getCell("L75").value = _billlog.blexpkvah;
            worksheet.getCell("L76").value = _billlog.blexpkvarhlagq2;
            worksheet.getCell("L77").value = _billlog.blexpkvarhlegq3;
            worksheet.getCell("L78").value = _billlog.blexpmdkw;
            worksheet.getCell("L79").value = _billlog.blexpmdkva;
            logger.info("Test 46");
            worksheet.getCell("L81").value = _billlog.bltouexpkwh0;
            worksheet.getCell("L82").value = _billlog.bltouexpkwh1;
            worksheet.getCell("L83").value = _billlog.bltouexpkwh2;
            worksheet.getCell("L84").value = _billlog.bltouexpkwh3;
            worksheet.getCell("L85").value = _billlog.bltouexpkwh4;
            worksheet.getCell("L86").value = _billlog.bltouexpkwh5;
            worksheet.getCell("L87").value = _billlog.bltouexpkwh6;
            worksheet.getCell("L88").value = _billlog.bltouexpkwh7;
            logger.info("Test 47");
            worksheet.getCell("L90").value = _billlog.bltouexpkvah0;
            worksheet.getCell("L91").value = _billlog.bltouexpkvah1;
            worksheet.getCell("L92").value = _billlog.bltouexpkvah2;
            worksheet.getCell("L93").value = _billlog.bltouexpkvah3;
            worksheet.getCell("L94").value = _billlog.bltouexpkvah4;
            worksheet.getCell("L95").value = _billlog.bltouexpkvah5;
            worksheet.getCell("L96").value = _billlog.bltouexpkvah6;
            worksheet.getCell("L97").value = _billlog.bltouexpkvah7;
            logger.info("Test 48");
            worksheet.getCell("L99").value = _billlog.bltouexpmdkw0;
            worksheet.getCell("L100").value = _billlog.bltouexpmdkw1;
            worksheet.getCell("L101").value = _billlog.bltouexpmdkw2;
            worksheet.getCell("L102").value = _billlog.bltouexpmdkw3;
            worksheet.getCell("L103").value = _billlog.bltouexpmdkw4;
            worksheet.getCell("L104").value = _billlog.bltouexpmdkw5;
            worksheet.getCell("L105").value = _billlog.bltouexpmdkw6;
            worksheet.getCell("L106").value = _billlog.bltouexpmdkw7;
            logger.info("Test 49");
            worksheet.getCell("L108").value = _billlog.bltouexpmdkva0;
            worksheet.getCell("L109").value = _billlog.bltouexpmdkva1;
            worksheet.getCell("L110").value = _billlog.bltouexpmdkva2;
            worksheet.getCell("L111").value = _billlog.bltouexpmdkva3;
            worksheet.getCell("L112").value = _billlog.bltouexpmdkva4;
            worksheet.getCell("L113").value = _billlog.bltouexpmdkva5;
            worksheet.getCell("L114").value = _billlog.bltouexpmdkva6;
            worksheet.getCell("L115").value = _billlog.bltouexpmdkva7;
            logger.info("Test 50");
        }
        console.log("Excel Process Complete");
        workbook.xlsx.writeFile(xlFilePath);
        console.log("Single Phase Bill Complete : " + xlFilePath);
        return 1;
    });
}