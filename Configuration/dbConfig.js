var ibmdb = require('ibm_db');
const logger = require('../Configuration/logger');
const connStr = "DATABASE=envisage_dev;HOSTNAME=172.19.10.133;UID=informix;PWD=Envisage@123;PORT=9089;PROTOCOL=TCPIP";
const Pool = require("ibm_db").Pool
    , pool = new Pool()
    , cn =connStr;



module.exports = {
    pool,
    connStr,
    // async getQueryExecution(qry){
    //     ibmdb.open(connStr, function (err,conn) {
    //     if (err) return console.log(err);

    //     conn.query('select * from envisage_dev.mstmenu', function (err, data) {
    //       if (err) console.log(err);
    //       else console.log(data);
    //       conn.close(function () {
    //         console.log('done');
    //       });
    //     });
    //   });
    // }

    queryExe(qry) {
        try {
            logger.info("Query : " + qry);
            const res = getQueryExecution(qry);
            return res;
        } catch (err) {
            logger.error("Error In Query : " + qry + "---- Error : " + err);
            console.log(err);
            return null
        }
    },


    apiFunctionWrapper(query) {
        return new Promise((resolve, reject) => {
            getQueryExecution(query,(successResponse) => {
                resolve(successResponse);
            }, (errorResponse) => {
                reject(errorResponse);
            });
        });
    },

    singleQuery(sql, callBack){
       // let {pool, connStr} = this.pool;
        let result;
        //this.connNum++;
        pool.open(connStr, (err, conn) => {
            if(err){
              console.log(err);
            }
            result = runSQL(conn, sql);
            if(callBack){
                console.log(result);
                //return result;
                callBack(result);  
            }
            conn.close();
            //this.connNum--;
        });
    },

    processQuery(sql){
        ibmdb.open(connStr, function(err, conn){
            if (err) return console.log(err);
            var query =sql;
            var result = conn.queryResultSync(query);
            console.log("data = ", result.fetchAllSync());
            console.log("metadata = ", result.getColumnMetadataSync());
            result.closeSync(); // Must call to free to avoid application error.
            conn.closeSync();
            return result;
          });
    },
    runSQL(conn, sql) {
        let result;
        this.queryNum++;
        if(typeof sql === "string"){
            try {
                result = conn.querySync(sql);
            }
            catch (error){
                console.log(error);
            }
        }
        else if(sql.query && sql.startCall && sql.successCall && sql.errorCall){
            let {query, startCall, successCall, errorCall} = sql;
            if(typeof query === "string"){
                startCall(this.pool.poolSize, this.queryNum);
                try{
                    result = conn.querySync(query);
                    successCall(this.pool.poolSize, this.queryNum - 1);
                }
                catch(error){
                    errorCall(this.pool.poolSize, this.queryNum - 1);
                }
            }
        }
        this.queryNum--;
        return result;
    }
}

async function getQueryExecution(qry) {
    ibmdb.open(connStr).then(
        async conn => {
            await conn.query(qry).then(data => {
                conn.closeSync();
                return data;
            }, err => {
                console.log(err);
            });
        }, err => {
            console.log(err)
        }
    );
    
}

