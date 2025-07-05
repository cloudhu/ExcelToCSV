const xlsx = require("xlsx");
const fs = require("fs");
const ini = require("ini");
const Path = require("path");
//表格文件后缀名列表
const suffixList = [".xlsx", ".xlsm", ".xls"];

function run(excelPath, outputPath) {
    var config = ini.parse(fs.readFileSync(__dirname + "/config.ini", "UTF-8"));
    if (!excelPath) excelPath = config.excelPath;
    if (!outputPath) outputPath = config.outputPath;
    console.log("excelPath:" + excelPath, "outputPath:" + outputPath);
    //
    const files = fs.readdirSync(excelPath);
    files.forEach(function (fileName, index) {
        //fileName是带后缀的整个文件名
        //extname是后缀名
        //name是不带后缀的文件名
        var extname = Path.extname(fileName);
        if (extname) {
            for (let suffix of suffixList) {
                if (extname === suffix) {
                    saveCSV(excelPath, fileName, outputPath);
                    break;
                }
            }
        }
    });
}

function saveCSV(rootPath, fileName, outputPath) {
    if (fileName.indexOf("~$") >= 0) return;
    let excelPath = Path.resolve(rootPath, fileName);
    console.log("excel:", excelPath);
    let workbook = xlsx.readFile(excelPath);
    //获取表名
    let sheetNames = workbook.SheetNames;
    for (let sheetName of sheetNames) {
        //跳过带#的注释表
        if (sheetName.indexOf("#") >= 0) continue;
        //通过表名得到表对象
        let sheet = workbook.Sheets[sheetName];
        //解析范围
        if (!sheet["!ref"]) continue;
        let range = xlsx.utils.decode_range(sheet["!ref"]);
        let colAmount = range.e.c;
        let rowAmount = range.e.r;
        //去掉被多余计算的列
        let overCol = 0;
        for (let j = 0; j <= colAmount; j++) {
            let address = { c: j, r: 0 };
            let ceil = xlsx.utils.encode_cell(address);
            if (j > 0 && !sheet[ceil]) {
                overCol++;
            }
        }
        colAmount -= overCol;
        //
        console.log(sheetName, colAmount, rowAmount);
        let data = "";
        for (let i = 0; i <= rowAmount; i++) {
            let address = { c: 0, r: i };
            let ceil = xlsx.utils.encode_cell(address);
            if (sheet[ceil]) {
                //跳过带#的行
                let rowName = sheet[ceil].w;
                if (rowName.indexOf("#") >= 0) continue;
            } else {
                //跳过KEY为空的行
                if (i > 0) continue;
            }
            for (let j = 0; j <= colAmount; j++) {
                let address = { c: j, r: 0 };
                let ceil = xlsx.utils.encode_cell(address);
                let colName;
                if (sheet[ceil]) {
                    colName = sheet[ceil].w;
                    //跳过标题带#的列
                    if (colName.indexOf("#") >= 0) {
                        if (j === colAmount) {
                            data = data.slice(0, -1);
                            data += "\n";
                        }
                        continue;
                    }
                }
                //
                address = { c: j, r: i };
                ceil = xlsx.utils.encode_cell(address);
                if (sheet[ceil]) {
                    if (i == 0) {
                        //首行标题
                        data += colName.split("|")[0];
                    } else {
                        //内容
                        if (colName) {
                            let sheetValue = sheet[ceil].w;
                            if (colName.indexOf("|Array") >= 0) {
                                //Obj格式转UEArray格式 Value1,Value2=>"(""Value1"",""Value2"")"
                                let str = `"(`;
                                let obj = sheetValue.split(",");
                                console.log(obj);
                                for (let i in obj) {
                                    if (obj[i] != "") {
                                        str += `""${obj[i]}"",`;
                                    }
                                }
                                str = str.slice(0, -1);
                                str += `)"`;
                                data += str;
                                console.log("------------------" + sheetName);
                            } else if (colName.indexOf("|Map") >= 0) {
                                //Obj格式转UEMap格式 Key1:Value1,Key2:Value2=>"((""Key1"",Value1),(""Key2"",Value2))"
                                let str = `"(`;
                                let obj;
                                eval("obj={" + sheet[ceil].w + "}");
                                for (let key in obj) {
                                    str += `(""${key}"",${obj[key]}),`;
                                }
                                str = str.slice(0, -1);
                                str += `)"`;
                                data += str;
                                console.log("------------------" + sheetName);
                            } else if (colName.indexOf("|Struct") >= 0) {
                                //Obj格式转结构体格式 Key1:Value1,Key2:Value2=>"(Key1=Value1,Key2=Value2)"
                                let str = `"(`;
                                str += sheet[ceil].w.replaceAll(":", "=");
                                str += `)"`;
                                data += str;
                            } else if (colName.indexOf("|GameplayTag") >= 0) {
                                //Equipment.Weapon.MainWeapon -> "(TagName=""Category.Equipment.Weapon.MainWeapon"")"
                                data += `"(TagName=""Category.` + sheet[ceil].w + `"")"`;
                            } else {
                                data += sheet[ceil].w;
                            }
                        } else {
                            data += sheet[ceil].w;
                        }
                    }
                }
                data += j === colAmount ? "\n" : ",";
            }
        }
        //MARK 旧方法，无法定制化
        // let data = xlsx.utils.sheet_to_csv(sheet, { blankrows: false });
        //重命名
        let extname = Path.extname(fileName);
        if (sheetName === "Export") {
            sheetName = "";
        }
        let csvName = fileName.replace(extname, sheetName + ".csv");
        let output = Path.resolve(outputPath, csvName);
        //保存文件
        fs.writeFileSync(output, data);
    }
}
const args = process.argv.splice(2);
run(args[0], args[1]);
