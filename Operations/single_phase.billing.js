const dbConf = require('../Configuration/dbConfig');
var Excel = require("exceljs");// load exceljs module
const fs = require("fs");
const logger = require("../Configuration/logger");
const Billing_Model = require('../ViewModels/Billing_Model');
const prepSlab = require("../ViewModels/Prep_Slab");
const billGen = require("../ViewModels/Bill_Generate");
const sourcePath = "D:\\Env_Temp_std\\";
const _common = require("./common.service");
const Bill_Gen = require('../ViewModels/Bill_Generate');
const _moments = require("moment");
module.exports = {
    async GetBillingDetails(from, to) {
        var _query = "select * from envisage_dev.mtrinstallconfig mic join envisage_dev.mstmeter m on mic.micmtrsrno = m.mtrsrno join envisage_dev.mstmtrmodel md on md.mdlid = m.mtrmodelid join envisage_dev.billinglog bl on bl.blmetersrno = mic.micmtrsrno where bl.bltimestamp >= " + from + " and bl.bltimestamp <=" + to + " and md.mdlphase = 'S'";
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
                BillingProcess(result, month, year, from, to, fromDate);
            }
        });
    }
}

async function BillingProcess(blmicmtrmdl, month, year, from, to, fromDate) {
    try {
        //console.log(year);
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
                console.log("CC Count : " + cmrcctoutrfmst.length + " ---- BL Count : " + blmicmtrmdl.length);
                logger.info("Query ccCount : " + _query);
                if (cmrcctoutrfmst != undefined) {
                    ProcessFileCreation(blmicmtrmdl[i], cmrcctoutrfmst, month, year, from, to, fromDate);
                }
            });
        }
    }
    catch (err) {
        logger.info("Error : " + err);
        console.log(err);
    }
}

async function ProcessFileCreation(blmicmtrmdl, cmrcctoutrfmst, month, year, from, to, fromDate) {
    try {
        var _folderPath = sourcePath + month + "_" + year + "//";
        fs.access(_folderPath, (error) => {
            if (error) {
                fs.mkdir(_folderPath, (error) => {
                    if (error) {
                        console.log(error);
                        var _trfFolder = _folderPath + cmrcctoutrfmst[0].preptrfid + "//";
                        fs.mkdir(_trfFolder, (err) => {
                            if (err) {
                                console.log(err);
                                FetchAndCopyBillFormat(blmicmtrmdl, cmrcctoutrfmst, _trfFolder, from, to);
                            }
                            else {
                                console.log("New Directory created successfully !!" + _trfFolder);
                                FetchAndCopyBillFormat(blmicmtrmdl, cmrcctoutrfmst, _trfFolder, from, to);
                            }
                        });
                    } else {
                        var _trfFolder = _folderPath + cmrcctoutrfmst[0].preptrfid + "//";
                        fs.mkdir(_trfFolder, (err) => {
                            if (err) console.log(err);
                            else {
                                console.log("New Directory created successfully !!" + _trfFolder);
                                FetchAndCopyBillFormat(blmicmtrmdl, cmrcctoutrfmst, _trfFolder, from, to);
                            }
                        });
                    }
                });
            } else {
                console.log("Given Directory already exists !!");
                var _trfFolder = _folderPath + cmrcctoutrfmst[0].preptrfid + "//";
                fs.mkdir(_trfFolder, (err) => {
                    if (err) console.log(err);
                    else {
                        console.log("New Directory created successfully !!" + _trfFolder);
                        FetchAndCopyBillFormat(blmicmtrmdl, cmrcctoutrfmst, _trfFolder, from, to);
                    }
                });
            }
        });
    }
    catch (err) {
        logger.info("Error : " + err);
        console.log(err);
    }
}

async function FetchAndCopyBillFormat(blmicmtrmdl, cmrcctoutrfmst, _trfFolder, from, to) {
    try {
        var _query = "select * from envisage_dev.preptarrifdiff s where s.ptdftarrid = '" + cmrcctoutrfmst[0].preptrfid + "'";
        let _formatData;
        dbConf.pool.open(dbConf.connStr, (err, conn) => {
            if (err) {
                console.log(err);
            }
            _formatData = dbConf.runSQL(conn, _query);
            conn.close();
            if (_formatData != undefined) {
                fs.copyFileSync(sourcePath + _formatData[0].ptdftarrtempdis, _trfFolder + blmicmtrmdl.micconsumerid + ".xls");
                ProcessSPBill(blmicmtrmdl, cmrcctoutrfmst, from, to, _trfFolder + blmicmtrmdl.micconsumerid + ".xls");
            }
        });
    }
    catch (err) {
        console.log(err);
        logger.info(err);
    }
}

async function ProcessSPBill(blmicmtrmdl, cmrcctoutrfmst, from, to, xlFilePath) {
    try {
        //#region File Processed
        //console.log(xlFilePath);
        var workbook = new Excel.Workbook();
        workbook.xlsx.readFile(xlFilePath).then(function () {
            var _spBill = Billing_Model.SPBill_Model;
            var _spBillGen = Bill_Gen;
            //console.log(xlFilePath);
            //#region DataFetch
            var _fromDate = _common.epochToJsDate(from);
            var _toDate = _common.epochToJsDate(to);
            var d = new Date();
            var _query = "Select * from mtrinstant_sp_log m where m.misplsrno = '" + blmicmtrmdl.micmtrsrno + "' and (m.mispldate >= " + from + " and m.mispldate <= " + to + ") order by 1 desc limit 1;";
            let _mtrinstsp;
            let _prepaydt;
            let _prepDtBl;
            let _consumerdets;
            let _energyDetail;
            let _inst;
            let _prepSlab;
            let _meters; let _trfData;
            let _toudtls;
            let _oldBill;
            let _trfDetails;
            let _slabData = [];
            logger.info("Query 01 : " + _query);
            dbConf.pool.open(dbConf.connStr, (err, conn) => {
                if (err) {
                    //console.log(err);
                }
                _mtrinstsp = dbConf.runSQL(conn, _query);
                _query = "select * from prepaydtlinst_sp p where p.pdispmtrsrno = '" + blmicmtrmdl.micmtrsrno + "' and (p.pdispdate >= " + from + " and p.pdispdate <= " + to + ") order by 1 desc limit 1;"
                //logger.info("Query 16 : " + _query);
                _prepaydt = dbConf.runSQL(conn, _query);
                _query = "select * from envisage_dev.prepaydtlbillinglog prep where prep.ppdtlblmetersrno = '" + blmicmtrmdl.micmtrsrno + "' and prep.ppdtlbltimestampid =" + from;
                _prepDtBl = dbConf.runSQL(conn, _query);
                _query = "select * from envisage_dev.mstconsumermeterrelation cn join envisage_dev.mstconsumer mc on cn.cmrconsumermasterid = mc.csmrmasterid";
                _query += " where cn.cmrconsumerid = '" + blmicmtrmdl.micconsumerid + "'";
                _consumerdets = dbConf.runSQL(conn, _query);
                _query = "select * from envisage_dev.mstcmrenergydetails c join mtrinstallconfig inst on c.cmreconsumerid = inst.micconsumerid where c.cmreconsumerid = '" + blmicmtrmdl.micconsumerid + "'";
                _energyDetail = dbConf.runSQL(conn, _query);
                _query = "select * from envisage_dev.mstconsumermeterrelation cm join envisage_dev.mtrinstallconfig mstin on cm.cmrconsumerid = mstin.micconsumerid";
                _query += " where mstin.micconsumerid ='" + blmicmtrmdl.micconsumerid + "'";
                let _mstrid = dbConf.runSQL(conn, _query);
                if (_mstrid != undefined) {
                    _query = "select * from envisage_dev.mstconsumermeterrelation m where m.cmrconsumermasterid = '" + _mstrid[0].cmrconsumermasterid + "' and m.cmrrelationenddate is null";
                    _inst = dbConf.runSQL(conn, _query);
                }
                _query = "select * from envisage_dev.mstconsumermeterrelation cm join envisage_dev.mtrinstallconfig mstin on cm.cmrconsumerid = mstin.micconsumerid";
                _query += " join informix.extrafields_kinesco e on cm.cmrconsumermasterid = e.consumermasterid where mstin.micconsumerid = '" + blmicmtrmdl.micconsumerid + "'";
                logger.info("Query : extrafieds : " + _query);
                let exf = dbConf.runSQL(conn, _query);

                _query = "select * from envisage_dev.ccattouslottrfdetail c join prepmsttariff p on c.tstdtariffid = p.preptrfid where c.tstdconclassid= '" + cmrcctoutrfmst[cmrcctoutrfmst.length - 1].cmracurclassid + "' and c.tstdconcatid ='" + cmrcctoutrfmst[cmrcctoutrfmst.length - 1].cmracurcategoryid + "'";
                _trfData = dbConf.runSQL(conn, _query);
                if (_trfData != undefined && _trfData.length > 0) {
                    _spBill.fixedCharge = _trfData[_trfData.length - 1].prepfixcharge;
                }
                else {
                    _spBill.fixedCharge = "0";
                }

                //console.log(_trfData);
                _query = "select ab.* from envisage_dev.prepmsttrffenrgyslabs ab join envisage_dev.prepmsttariff pr on ab.preptdltrfid = pr.preptrfid where ab.preptdltrfid = '" + cmrcctoutrfmst[0].preptrfid + "' and pr.preptrftodate is null order by ab.preptdirecid desc, ab.preptdltrfslotno asc";
                logger.info("Query _prepSlab : " + _query);
                _prepSlab = dbConf.runSQL(conn, _query);
                if (_prepSlab != undefined && _prepSlab.length > 0) {
                    var _prevSlab = "";
                    for (var i = 0; i < _prepSlab.length; i++) {
                        if (!_prevSlab.includes(_prepSlab[i].preptdltrfslotno))
                            _slabData.push(_prepSlab[i]);
                        _prevSlab += _prepSlab[i].preptdltrfslotno + ",";
                    }
                }
                _query = "select * from envisage_dev.prepmsttrffdmdslab bc where bc.preptdltrfid = '" + cmrcctoutrfmst[0].preptrfid + "'";
                _trfdmdslabs = dbConf.runSQL(conn, _query);
                _query = "select m.mtrsrno,mic.micdcusrno,mm.mdlid,prptr.prepenrgtrftype,prptr.prepdmdtrftype,prptr.prepuntexsdmdchrg,prptr.prepmthfixedchrg,";
                _query += "prptr.prepdlyfixedchrg,prptr.prepemrgncycrlmt,prptr.preprsrvecrlmtrcl,prptr.preptrfflatenerrate,prptr.preptrfflatdmdrate,prptr.preptrfuserid,";
                _query += "prptr.preptrfid,ccat.cattrtouid,mm.mdlphase,mic.micdtrid,mic.micservicestartdate,prptr.preptaxrchrg,prptr.prepuntdmdchrg,";
                _query += "prptr.prepfixcharge,prptr.prepgstper,prptr.prepfuelsuchrg,prptr.prepmtrhire,prptr.preplowvoltsuchrg,prptr.preptrfschemedescription,";
                _query += " cmra.cmrafromdate, cmra.cmrasznsz from envisage_dev.mstmtrmodel mm join envisage_dev.mstmeter m on mm.mdlid = m.mtrmodelid ";
                _query += "envisage_dev.mtrinstallconfig mic on m.mtrsrno = mic.micmtrsrno join envisage_dev.cmrattrib cmra on mic.micconsumerid = cmra.cmraconsumerid ";
                _query += "envisage_dev.ccattarifftourel ccat on cmra.cmracurcategoryid = ccat.cattrconsumercatid and cmra.cmracurclassid = ccat.cattrconsumerclassid ";
                _query += "envisage_dev.ccattouslottrfdetail ccatrf on cmra.cmracurcategoryid = ccatrf.tstdconcatid and cmra.cmracurclassid = ccatrf.tstdconclassid ";
                _query += "envisage_dev.prepmsttariff prptr on ccatrf.tstdtariffid = prptr.preptrfid where m.mtrsrno = '" + blmicmtrmdl.micmtrsrno + "' and ccat.cattrtodate is null ";
                _query += "mic.micconsumerid ='" + blmicmtrmdl.micconsumerid + "' and cmra.cmratodate is null and prptr.preptrftodate is null";
                var znid = "";
                _meters = dbConf.runSQL(conn, _query);

                _query = "select * from envisage_dev.prepmsttariff prptr join envisage_dev.ccattouslottrfdetail ccatrf on ";
                _query += " ccatrf.tstdtariffid = prptr.preptrfid join  envisage_dev.cmrattrib cmra on cmra.cmracurclassid = ccatrf.tstdconclassid and cmra.cmracurcategoryid= ccatrf.tstdtouslotid where cmra.cmraconsumerid = '" + blmicmtrmdl.micconsumerid + "' and cmra.cmratodate is null";
                _trfDetails = dbConf.runSQL(conn, _query);
                //console.log(_meters);
                if (_meters != undefined && _meters.length > 0) {
                    znid = _meters[0].cattrtouid;
                }
                _query = "select cd.pretsslotno,cd.pretsslotstarthr,cd.pretsslotstartmin,cd.pretsslotendhr,cd.pretsslotendmin,cd.pretsincconamt,cd.pretsincdmdamt,cd.pretsload,cd.pretsminchrgdmdlim,cd.tsmaxdemandlimit from envisage_dev.prepmsttouslot cd where cd.pretstouid = '" + znid + "'";
                _toudtls = dbConf.runSQL(conn, _query);
                _query = "select * from envisage_dev.billinglog where blmetersrno = '" + blmicmtrmdl.micmtrsrno + "' and bltimestamp < " + blmicmtrmdl.bltimestamp + " order by blrecid desc limit 1";
                _oldBill = dbConf.runSQL(conn, _query);
                conn.close();
                //console.log(_prepaydt);
                if (_mtrinstsp != undefined) {
                    //#region  SP_Bill

                    _spBill.RTC_absolute = d.toDateString();

                    if (_mtrinstsp.length > 0) {
                        _spBill.Active_Power = _mtrinstsp[0].misplmdkw == undefined ? 0 : _mtrinstsp[0].misplmdkw;
                        _spBill.Apparent_Power = _mtrinstsp[0].misplmdkva == undefined ? 0 : _mtrinstsp[0].misplmdkva;
                        _spBill.Instantaneous_Frequency = _mtrinstsp[0].misplfreq == undefined ? 0 : _mtrinstsp[0].misplfreq;
                        _spBill.Signed_Reactive_power = _mtrinstsp[0].misplkvar == undefined ? 0 : _mtrinstsp[0].misplkvar;
                        _spBill.Signed_Three_phase_Power_factor = _mtrinstsp[0].misplpf == undefined ? 0 : _mtrinstsp[0].misplpf;
                    }
                    _spBill.Cumulative_Energy_KWh = blmicmtrmdl.blcumkwh == undefined ? 0 : blmicmtrmdl.blcumkwh;
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
                    if (_prepaydt != undefined && _prepaydt.length > 0) {
                        _spBill.Cumulative_Emergency_Credit = _prepaydt[0].pdispcumemercred == null ? 0.00 : _prepaydt[0].pdispcumemercred;
                        _spBill.Cum_Monthly_Fixed_Charge_deduction_Amount = _prepaydt[0].pdispcummthfxdchrgdedamt == null ? 0.00 : _prepaydt[0].pdispcummthfxdchrgdedamt;;
                        _spBill.Cum_Demand_Deduction_Amount = _prepaydt[0].pdispcumdmddedamt == null ? 0.00 : _prepaydt[0].pdispcumdmddedamt;
                        _spBill.Cum_Adjustable_Amount = _prepaydt[0].pdispcumadjamt == null ? 0.00 : _prepaydt[0].pdispcumadjamt;
                        _spBill.Cum_Daily_Fixed_Charge_deduction_Amount = _prepaydt[0].pdispcdfcdamt == null ? 0.00 : _prepaydt[0].pdispcdfcdamt;
                        _spBill.CumEnergy_Charge_Deduction_Amount = _prepaydt[0].pdispcumenrgchrgdedamt == null ? 0.00 : _prepaydt[0].pdispcumenrgchrgdedamt;
                    }
                    //console.log(_spBill);
                    //#endregion SP_Bill

                    //#region SP_Bill_Gen
                    var connum = blmicmtrmdl.micconsumerid.substring(6, 4);
                    var firstDay = new Date(_fromDate.getFullYear(), _fromDate.getMonth(), 1);
                    var lastDay = new Date(_fromDate.getFullYear(), _fromDate.getMonth() + 1, 0);
                    let _connecteddate;
                    if (blmicmtrmdl.micservicestartdate != undefined && blmicmtrmdl.micservicestartdate != null)
                        _connecteddate = _common.epochToJsDate(blmicmtrmdl.micservicestartdate);
                    if (_energyDetail != undefined && _energyDetail.length > 0) {
                        _spBillGen.connectedload = _energyDetail[_energyDetail.length - 1].cmrecurconnloadkw;
                        _spBillGen.Contractdemand = _energyDetail[_energyDetail.length - 1].cmrecurcontdmdkva;
                    }
                    else {
                        _spBillGen.Contractdemand = "0";
                        _spBillGen.connectedload = "0";
                    }
                    _spBillGen.Sourcefile = "";
                    _spBillGen.billMonth = _fromDate.getMonth() + 1;
                    _spBillGen.billYear = _fromDate.getFullYear();
                    _spBillGen.billenddate = lastDay.toString();
                    _spBillGen.billnumber = _spBillGen.billMonth + _spBillGen.billYear + connum;
                    _spBillGen.billstdate = firstDay.toString();
                    if (_connecteddate != undefined)
                        _spBillGen.connecteddate = _connecteddate.getDate() + "-" + _connecteddate.getMonth() + 1 + "-" + _connecteddate.getFullYear();
                    if (_mstrid == undefined)
                        _spBillGen.fixchargunit = "1";
                    if (_inst != undefined)
                        _spBillGen.fixchargunit = _inst.length;
                    else
                        _spBillGen.fixchargunit = "1";
                    if (exf != undefined && exf.length > 0) {
                        _spBillGen.gstnumber = exf[0].gstnumber;
                    }
                    else
                        _spBillGen.gstnumber = "";
                    _spBillGen.metermodelno = blmicmtrmdl.mdlid;
                    _spBillGen.region = "";
                    _spBillGen.seznsez = cmrcctoutrfmst[0].cmrasznsz;
                    _spBillGen.tarrifcode = cmrcctoutrfmst[0];
                    _spBillGen.txnconsumerid = blmicmtrmdl.micconsumerid;
                    _spBillGen.txnmtrsrno = blmicmtrmdl.micmtrsrno;
                    _spBillGen.utility = "";
                    _spBillGen.zone = "";
                    //console.log(xlFilePath);
                    //#endregion SP_Bill_Gen
                    //ProcessExcelFile(blmicmtrmdl, cmrcctoutrfmst, _spBill, _spBillGen, trfData, xlFilePath, _prepSlab, _trfdmdslabs, _toudtls, _prepDtBl, _consumerdets)


                }
            });
            //#endregion DataFetch
            logger.info(xlFilePath);
            var worksheet = workbook.getWorksheet("Bill Parameter Mapping");
            const v0 = worksheet.getCell('D5').value;
            if (_trfData != undefined && _trfData.length > 0)
                worksheet.getCell('D5').value = _trfData[_trfData.length - 1].preptrfschemedescription;
            else
                worksheet.getCell('D5').value = "";
            worksheet.getCell('D6').value = _spBill.Cum_Adjustable_Amount;
            worksheet.getCell('D7').value = "MD KVA";
            worksheet.getCell('D8').value = "KWh ToT";
            worksheet.getCell('D9').value = "No";
            worksheet.getCell('D10').value = "Yes";
            worksheet.getCell('D11').value = "No";
            if (_slabData != undefined && _slabData.length > 0) {
                worksheet.getCell('D12').value = _slabData.length;
                worksheet.getCell('D13').value = "0";
                var count = _slabData.length;
                var rowVal = 14;
                var i = 0;
                for (i = 0; i < count; i++) {
                    var cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = _slabData[i].preptdlenrgunitsfrom;              //Energy Slab 1 : Start Reading
                    rowVal++;
                    cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = _slabData[i].preptdlenrgunitsto;              //Energy Slab 1 : End Reading
                    rowVal++;
                    cellName = "D" + rowVal;
                    worksheet.getCell(cellName).value = _slabData[i].preptdlenerchrgamt;              //Energy Slab 1 : Rate
                    rowVal++;
                }
                if (rowVal != 37) {
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
                }
            }
            else {
                worksheet.getCell('D12').value = 0;
                if (_trfDetails != undefined && _trfDetails.length > 0) {
                    console.log(_trfDetails);
                    worksheet.getCell('D39').value = _trfDetails[_trfDetails.length - 1].preptrfflatdmdrate;
                    worksheet.getCell('D13').value = _trfDetails[_trfDetails.length - 1].preptrfflatenerrate;
                    worksheet.getCell('D66').value = _trfDetails[_trfDetails.length - 1].prepuntdmdchrg;
                    worksheet.getCell('D67').value = _trfDetails[_trfDetails.length - 1].prepuntexsdmdchrg;
                    worksheet.getCell('D68').value = _trfDetails[_trfDetails.length - 1].prepmthfixedchrg;
                    worksheet.getCell('D69').value = _trfDetails[_trfDetails.length - 1].prepdlyfixedchrg;
                    worksheet.getCell('D70').value = _trfDetails[_trfDetails.length - 1].prepemrgncycrlmt;
                    worksheet.getCell("L9").value = _trfDetails[_trfDetails.length - 1].preptaxrchrg + " %";
                    worksheet.getCell("L14").value = _trfDetails[_trfDetails.length - 1].prepfuelsuchrg;
                    worksheet.getCell("L15").value = _trfDetails[_trfDetails.length - 1].prepfixcharge;
                    worksheet.getCell("L16").value = _trfDetails[_trfDetails.length - 1].prepgstper + " %";
                    worksheet.getCell("L17").value = _trfDetails[_trfDetails.length - 1].preplowvoltsuchrg;
                    worksheet.getCell("L19").value = _trfDetails[_trfDetails.length - 1].prepmtrhire;
                }
                else {
                    worksheet.getCell('D39').value = "0";
                    worksheet.getCell('D13').value = "0";
                    worksheet.getCell('D66').value = "0";
                    worksheet.getCell('D67').value = "0";
                    worksheet.getCell('D68').value = "0";
                    worksheet.getCell('D69').value = "0";
                    worksheet.getCell('D70').value = "0";
                    worksheet.getCell("L9").value = "";
                    worksheet.getCell("L14").value = "0";
                    worksheet.getCell("L15").value = "0";
                    worksheet.getCell("L16").value = "0";
                    worksheet.getCell("L17").value = "0";
                    worksheet.getCell("L19").value = "0";
                }
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
            }
            if (_trfdmdslabs != undefined && _trfdmdslabs.length > 0) {
                worksheet.getCell('D38').value = _trfdmdslabs.length;
                worksheet.getCell('D39').value = "0";
                var count = _trfdmdslabs.length;
                var rowVal = 40;
                var i = 0;
                if (count <= 4) {
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
                worksheet.getCell('D38').value = "0";
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
            if (_spBillGen.Contractdemand != undefined)
                worksheet.getCell('D64').value = _spBillGen.Contractdemand;
            else
                worksheet.getCell('D64').value = "0";
            worksheet.getCell('D65').value = "900";
            worksheet.getCell('D71').value = "0";
            if (_toudtls != undefined && _toudtls.length > 0) {
                var toucount = _toudtls.length;
                worksheet.getCell('D72').value = toucount;
                var rowVal = 73;
                var i = 0;
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
                if (toucount != 8) {
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
                }
            }
            else {
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
                worksheet.getCell("D124").value = blmicmtrmdl.mtrmodelid;
            }
            worksheet.getCell("D122").value = _spBillGen.fixchargunit;
            worksheet.getCell("D123").value = _spBillGen.seznsez;
            worksheet.getCell("H5").value = _spBill.RTC_absolute;
            worksheet.getCell("H6").value = 0;
            worksheet.getCell("H7").value = 0;
            worksheet.getCell("H8").value = 0;
            worksheet.getCell("H9").value = 0;
            worksheet.getCell("H10").value = 0;
            worksheet.getCell("H11").value = 0;
            worksheet.getCell("H12").value = _spBill.Signed_Three_phase_Power_factor;
            worksheet.getCell("H13").value = _spBill.Instantaneous_Frequency;
            worksheet.getCell("H14").value = _spBill.Apparent_Power;
            worksheet.getCell("H15").value = _spBill.Active_Power;
            worksheet.getCell("H16").value = _spBill.Signed_Reactive_power;
            var _date = _common.epochToJsDate(_spBill.Billing_Date);
            var day = _date.getDate();
            //var month = _date.toLocaleString('default', { month: 'short' });
            var month = _date.getMonth() + 1;
            var year = _date.getFullYear();
            var dtForms = _moments().format("DD-MMM-yy");
            // var dateDigit = ("0" + _date.getDate()).slice(-2);
            // var monthDigit = _date.toLocaleString('en-us', { month: 'short'});
            // var yearDigit = _date.getFullYear().toLocaleString().substring(3,5);
            var datestring = dtForms;
            worksheet.getCell("H17").value = datestring;
            worksheet.getCell("H18").value = _spBill.Billing_Index;
            worksheet.getCell("H19").value = _spBill.Signed_Average_PowerFactor;
            worksheet.getCell("H20").value = _spBill.Cumulative_Energy_KWh;
            worksheet.getCell("H21").value = _spBill.Cumulative_Energy_KVAh;
            worksheet.getCell("H22").value = _spBill.Cumulative_Energy_KVARh_Lag;
            worksheet.getCell("H23").value = _spBill.Cumulative_Energy_KVARh_Lead;
            worksheet.getCell("H24").value = _spBill.Maximum_demand_KW;
            worksheet.getCell("H25").value = _spBill.Maximum_demand_KVA;
            worksheet.getCell("H26").value = _spBill.Cumulative_tamper_count;
            worksheet.getCell("H28").value = _spBill.LCumulative_Energy_for_TOU1;
            worksheet.getCell("H29").value = _spBill.LCumulative_Energy_for_TOU2;
            worksheet.getCell("H30").value = _spBill.LCumulative_Energy_for_TOU3;
            worksheet.getCell("H31").value = _spBill.LCumulative_Energy_for_TOU4;
            worksheet.getCell("H32").value = _spBill.LCumulative_Energy_for_TOU5;
            worksheet.getCell("H33").value = _spBill.LCumulative_Energy_for_TOU6;
            worksheet.getCell("H34").value = _spBill.LCumulative_Energy_for_TOU7;
            worksheet.getCell("H35").value = _spBill.LCumulative_Energy_for_TOU8;
            worksheet.getCell("H37").value = _spBill.LCumulative_Apparent_Energy_for_TOU1;
            worksheet.getCell("H38").value = _spBill.LCumulative_Apparent_Energy_for_TOU2;
            worksheet.getCell("H39").value = _spBill.LCumulative_Apparent_Energy_for_TOU3;
            worksheet.getCell("H40").value = _spBill.LCumulative_Apparent_Energy_for_TOU4;
            worksheet.getCell("H41").value = _spBill.LCumulative_Apparent_Energy_for_TOU5;
            worksheet.getCell("H42").value = _spBill.LCumulative_Apparent_Energy_for_TOU6;
            worksheet.getCell("H43").value = _spBill.LCumulative_Apparent_Energy_for_TOU7;
            worksheet.getCell("H44").value = _spBill.LCumulative_Apparent_Energy_for_TOU8;
            worksheet.getCell("H46").value = _spBill.LMD_kW_for_TOU1;
            worksheet.getCell("H47").value = _spBill.LMD_kW_for_TOU2;
            worksheet.getCell("H48").value = _spBill.LMD_kW_for_TOU3;
            worksheet.getCell("H49").value = _spBill.LMD_kW_for_TOU4;
            worksheet.getCell("H50").value = _spBill.LMD_kW_for_TOU5;
            worksheet.getCell("H51").value = _spBill.LMD_kW_for_TOU6;
            worksheet.getCell("H52").value = _spBill.LMD_kW_for_TOU7;
            worksheet.getCell("H53").value = _spBill.LMD_kW_for_TOU8;
            worksheet.getCell("H55").value = _spBill.LMD_kVA_for_TOU1;
            worksheet.getCell("H56").value = _spBill.LMD_kVA_for_TOU2;
            worksheet.getCell("H57").value = _spBill.LMD_kVA_for_TOU3;
            worksheet.getCell("H58").value = _spBill.LMD_kVA_for_TOU4;
            worksheet.getCell("H59").value = _spBill.LMD_kVA_for_TOU5;
            worksheet.getCell("H60").value = _spBill.LMD_kVA_for_TOU6;
            worksheet.getCell("H61").value = _spBill.LMD_kVA_for_TOU7;
            worksheet.getCell("H62").value = _spBill.LMD_kVA_for_TOU8;
            if (blmicmtrmdl != undefined)
                worksheet.getCell("H64").value = _spBill.Cumulative_tamper_count;
            else
                worksheet.getCell("H64").value = "0";
            worksheet.getCell("H65").value = _spBill.Cumulative_recharge_amount;
            worksheet.getCell("H66").value = _spBill.Cumulative_Balance_Deduction_Register;
            if (_prepDtBl != undefined)
                worksheet.getCell("H67").value = _prepDtBl.ppdtlblcumnumofrchrgcnt;
            else
                worksheet.getCell("H67").value = "";
            worksheet.getCell("H68").value = _spBill.Cumulative_Emergency_Credit;
            worksheet.getCell("H69").value = _spBill.Cum_Adjustable_Amount;
            worksheet.getCell("H70").value = _spBill.Cum_Monthly_Fixed_Charge_deduction_Amount;
            worksheet.getCell("H71").value = _spBill.Cum_Daily_Fixed_Charge_deduction_Amount;
            worksheet.getCell("H72").value = _spBill.Cum_Demand_Deduction_Amount;
            worksheet.getCell("H73").value = _spBill.CumEnergy_Charge_Deduction_Amount;
            worksheet.getCell("H75").value = 0;
            worksheet.getCell("H76").value = 0;
            worksheet.getCell("H77").value = 0;
            worksheet.getCell("H78").value = 0;
            worksheet.getCell("H79").value = 0;
            worksheet.getCell("H81").value = 0;
            worksheet.getCell("H82").value = 0;
            worksheet.getCell("H83").value = 0;
            worksheet.getCell("H84").value = 0;
            worksheet.getCell("H85").value = 0;
            worksheet.getCell("H86").value = 0;
            worksheet.getCell("H87").value = 0;
            worksheet.getCell("H88").value = 0;
            worksheet.getCell("H90").value = 0;
            worksheet.getCell("H91").value = 0;
            worksheet.getCell("H92").value = 0;
            worksheet.getCell("H93").value = 0;
            worksheet.getCell("H94").value = 0;
            worksheet.getCell("H95").value = 0;
            worksheet.getCell("H96").value = 0;
            worksheet.getCell("H97").value = 0;
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
            var conid = "";
            if (_spBillGen.txnconsumerid.length > 4)
                conid = _spBillGen.txnconsumerid.substring(6, 4);
            else
                conid = _spBillGen.txnconsumerid;
            //console.log("Consumer : " + conid);
            worksheet.getCell("L5").value = conid;
            worksheet.getCell("L6").value = _spBillGen.billMonth + "" + _spBillGen.billYear;
            if (_consumerdets != undefined) {
                worksheet.getCell("L7").value = _consumerdets[0].csmrfirstname + " " + _consumerdets[0].csmrlastname;
                worksheet.getCell("L8").value = _consumerdets[0].csmraddress1;
            }
            else {
                worksheet.getCell("L7").value = "";
                worksheet.getCell("L8").value = "";
            }
            if (_spBillGen != undefined)
                worksheet.getCell("L10").value = _spBillGen.gstnumber;
            else
                worksheet.getCell("L10").value = "";
            worksheet.getCell("L11").value = _spBillGen.connectedload;
            worksheet.getCell("L12").value = _spBillGen.billnumber;
            worksheet.getCell("L13").value = _spBillGen.connecteddate;
            if (_oldBill == undefined || _oldBill.length == 0) {
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
                worksheet.getCell("L20").value = _oldBill[0].blcumkwh;
                worksheet.getCell("L21").value = _oldBill[0].blcumkvah;
                worksheet.getCell("L22").value = _oldBill[0].blcumkvarh_lag;
                worksheet.getCell("L23").value = _oldBill[0].blcumkvarh_lead;
                worksheet.getCell("L24").value = _oldBill[0].blmdkw;
                worksheet.getCell("L25").value = _oldBill[0].blmdkva;
                worksheet.getCell("L26").value = _oldBill[0].blcumtampercount;
                worksheet.getCell("L28").value = _oldBill[0].blcumkwhtou1;
                worksheet.getCell("L29").value = _oldBill[0].blcumkwhtou2;
                worksheet.getCell("L30").value = _oldBill[0].blcumkwhtou3;
                worksheet.getCell("L31").value = _oldBill[0].blcumkwhtou4;
                worksheet.getCell("L32").value = _oldBill[0].blcumkwhtou5;
                worksheet.getCell("L33").value = _oldBill[0].blcumkwhtou6;
                worksheet.getCell("L34").value = _oldBill[0].blcumkwhtou7;
                worksheet.getCell("L35").value = _oldBill[0].blcumkwhtou8;
                worksheet.getCell("L37").value = _oldBill[0].blcumkvahtou1;
                worksheet.getCell("L38").value = _oldBill[0].blcumkvahtou2;
                worksheet.getCell("L39").value = _oldBill[0].blcumkvahtou3;
                worksheet.getCell("L40").value = _oldBill[0].blcumkvahtou4;
                worksheet.getCell("L41").value = _oldBill[0].blcumkvahtou5;
                worksheet.getCell("L42").value = _oldBill[0].blcumkvahtou6;
                worksheet.getCell("L43").value = _oldBill[0].blcumkvahtou7;
                worksheet.getCell("L44").value = _oldBill[0].blcumkvahtou8;
                worksheet.getCell("L46").value = _oldBill[0].blmdkwtou1;
                worksheet.getCell("L47").value = _oldBill[0].blmdkwtou2;
                worksheet.getCell("L48").value = _oldBill[0].blmdkwtou3;
                worksheet.getCell("L49").value = _oldBill[0].blmdkwtou4;
                worksheet.getCell("L50").value = _oldBill[0].blmdkwtou5;
                worksheet.getCell("L51").value = _oldBill[0].blmdkwtou6;
                worksheet.getCell("L52").value = _oldBill[0].blmdkwtou7;
                worksheet.getCell("L53").value = _oldBill[0].blmdkwtou8;
                worksheet.getCell("L55").value = _oldBill[0].blmdkvatou1;
                worksheet.getCell("L56").value = _oldBill[0].blmdkvatou2;
                worksheet.getCell("L57").value = _oldBill[0].blmdkvatou3;
                worksheet.getCell("L58").value = _oldBill[0].blmdkvatou4;
                worksheet.getCell("L59").value = _oldBill[0].blmdkvatou5;
                worksheet.getCell("L60").value = _oldBill[0].blmdkvatou6;
                worksheet.getCell("L61").value = _oldBill[0].blmdkvatou7;
                worksheet.getCell("L62").value = _oldBill[0].blmdkvatou8;

                worksheet.getCell("L75").value = _oldBill[0].blexpkvah;
                worksheet.getCell("L76").value = _oldBill[0].blexpkvarhlagq2;
                worksheet.getCell("L77").value = _oldBill[0].blexpkvarhlegq3;
                worksheet.getCell("L78").value = _oldBill[0].blexpmdkw;
                worksheet.getCell("L79").value = _oldBill[0].blexpmdkva;

                worksheet.getCell("L81").value = _oldBill[0].bltouexpkwh0;
                worksheet.getCell("L82").value = _oldBill[0].bltouexpkwh1;
                worksheet.getCell("L83").value = _oldBill[0].bltouexpkwh2;
                worksheet.getCell("L84").value = _oldBill[0].bltouexpkwh3;
                worksheet.getCell("L85").value = _oldBill[0].bltouexpkwh4;
                worksheet.getCell("L86").value = _oldBill[0].bltouexpkwh5;
                worksheet.getCell("L87").value = _oldBill[0].bltouexpkwh6;
                worksheet.getCell("L88").value = _oldBill[0].bltouexpkwh7;

                worksheet.getCell("L90").value = _oldBill[0].bltouexpkvah0;
                worksheet.getCell("L91").value = _oldBill[0].bltouexpkvah1;
                worksheet.getCell("L92").value = _oldBill[0].bltouexpkvah2;
                worksheet.getCell("L93").value = _oldBill[0].bltouexpkvah3;
                worksheet.getCell("L94").value = _oldBill[0].bltouexpkvah4;
                worksheet.getCell("L95").value = _oldBill[0].bltouexpkvah5;
                worksheet.getCell("L96").value = _oldBill[0].bltouexpkvah6;
                worksheet.getCell("L97").value = _oldBill[0].bltouexpkvah7;

                worksheet.getCell("L99").value = _oldBill[0].bltouexpmdkw0;
                worksheet.getCell("L100").value = _oldBill[0].bltouexpmdkw1;
                worksheet.getCell("L101").value = _oldBill[0].bltouexpmdkw2;
                worksheet.getCell("L102").value = _oldBill[0].bltouexpmdkw3;
                worksheet.getCell("L103").value = _oldBill[0].bltouexpmdkw4;
                worksheet.getCell("L104").value = _oldBill[0].bltouexpmdkw5;
                worksheet.getCell("L105").value = _oldBill[0].bltouexpmdkw6;
                worksheet.getCell("L106").value = _oldBill[0].bltouexpmdkw7;

                worksheet.getCell("L108").value = _oldBill[0].bltouexpmdkva0;
                worksheet.getCell("L109").value = _oldBill[0].bltouexpmdkva1;
                worksheet.getCell("L110").value = _oldBill[0].bltouexpmdkva2;
                worksheet.getCell("L111").value = _oldBill[0].bltouexpmdkva3;
                worksheet.getCell("L112").value = _oldBill[0].bltouexpmdkva4;
                worksheet.getCell("L113").value = _oldBill[0].bltouexpmdkva5;
                worksheet.getCell("L114").value = _oldBill[0].bltouexpmdkva6;
                worksheet.getCell("L115").value = _oldBill[0].bltouexpmdkva7;
            }
            // console.log("Excel Process Complete");
            workbook.xlsx.writeFile(xlFilePath);
            var worksheetPB = workbook.getWorksheet("Print_Bill");
            var num = worksheetPB.getCell("I37").value.result;
            //console.log(num);
            var a = ['', 'one ', 'two ', 'three ', 'four ', 'five ', 'six ', 'seven ', 'eight ', 'nine ', 'ten ', 'eleven ', 'twelve ', 'thirteen ', 'fourteen ', 'fifteen ', 'sixteen ', 'seventeen ', 'eighteen ', 'nineteen '];
            var b = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];
            if ((num = num.toString()).length > 9) return 'overflow';
            n = ('000000000' + num).substring(-9).match(/^(\d{2})(\d{2})(\d{2})(\d{1})(\d{2})$/);
            if (!n) return; var str = '';
            str += (n[1] != 0) ? (a[Number(n[1])] || b[n[1][0]] + ' ' + a[n[1][1]]) + 'crore ' : '';
            str += (n[2] != 0) ? (a[Number(n[2])] || b[n[2][0]] + ' ' + a[n[2][1]]) + 'lakh ' : '';
            str += (n[3] != 0) ? (a[Number(n[3])] || b[n[3][0]] + ' ' + a[n[3][1]]) + 'thousand ' : '';
            str += (n[4] != 0) ? (a[Number(n[4])] || b[n[4][0]] + ' ' + a[n[4][1]]) + 'hundred ' : '';
            str += (n[5] != 0) ? ((str != '') ? 'and ' : '') + (a[Number(n[5])] || b[n[5][0]] + ' ' + a[n[5][1]]) + 'only ' : '';
            //console.log(str);

            worksheetPB.getCell("F39").value = str;
            workbook.xlsx.writeFile(xlFilePath);
            //console.log("Single Phase Bill Complete : " + xlFilePath);
            return 1;
        });
        //#endregion File Processed

    }
    catch (err) {
        console.log(err);
    }
}

async function numToWords(num) {
    console.log(num);
    var a = ['', 'one ', 'two ', 'three ', 'four ', 'five ', 'six ', 'seven ', 'eight ', 'nine ', 'ten ', 'eleven ', 'twelve ', 'thirteen ', 'fourteen ', 'fifteen ', 'sixteen ', 'seventeen ', 'eighteen ', 'nineteen '];
    var b = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];
    if ((num = num.toString()).length > 9) return 'overflow';
    n = ('000000000' + num).substring(-9).match(/^(\d{2})(\d{2})(\d{2})(\d{1})(\d{2})$/);
    if (!n) return; var str = '';
    str += (n[1] != 0) ? (a[Number(n[1])] || b[n[1][0]] + ' ' + a[n[1][1]]) + 'crore ' : '';
    str += (n[2] != 0) ? (a[Number(n[2])] || b[n[2][0]] + ' ' + a[n[2][1]]) + 'lakh ' : '';
    str += (n[3] != 0) ? (a[Number(n[3])] || b[n[3][0]] + ' ' + a[n[3][1]]) + 'thousand ' : '';
    str += (n[4] != 0) ? (a[Number(n[4])] || b[n[4][0]] + ' ' + a[n[4][1]]) + 'hundred ' : '';
    str += (n[5] != 0) ? ((str != '') ? 'and ' : '') + (a[Number(n[5])] || b[n[5][0]] + ' ' + a[n[5][1]]) + 'only ' : '';
    console.log(str);
    return str;
}
async function ProcessExcelFile(blmicmtrmdl, cmrcctoutrfmst, _spdata, billGens, _trfData, xlFilePath, _prepSlab, _trfdmdslabs, _toudtls, _prepDtBl, _consumerdets) {
    try {

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