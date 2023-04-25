const dbConf = require('../Configuration/dbConfig');
const fs = require("fs");
module.exports = {

    async CheckFileExists(path) {
        const fileExists = async path => !!(await fs.stat(path).catch(e => false));
        console.log(fileExists);
    },

    async CreateFileFolder(path) {
        await fs.access(path, (error) => {
            if (error) {
               fs.mkdir(path, (error) => {
                    if (error) {
                        console.log(error);
                        return 0;
                    } else {
                         console.log("New Directory created successfully !!" + path);
                         return 1;
                    }
                });
            } else {
                return 1;
                //console.log("Given Directory already exists !!");
            }
        });
    },

    CopyFileToAnother(fromPath, toPath){
        try { 
            fs.copyFileSync(fromPath, toPath);
            return 1;
        }
        catch(err){
            console.log(err);
            return 0;
        }
    },
    epochToJsDate(ts) {
        return new Date(ts * 1000);
    }
}
