import * as xlsx from 'xlsx';
import * as fs from 'fs';
import ExportTool, { outJsonRoot, outInterfaceRoot } from './exportTool';

export default class ExportSheet {
    private _sheetData: ExportSheetData;
    public constructor(sheetData: ExportSheetData) {
        this._sheetData = sheetData;
    }

    public parseSheet() {
        let sheetName = this._sheetData.sheetName;
        let sheet = this._sheetData.workSheet;
        if (!sheet) {
            console.error("导出数据失败:" + this._sheetData.excel + "不存在" + sheetName + "表格");
            return;
        }
        let type = this._sheetData.type;
        if (type == ExportType.TABLE) {
            this.parseSheetAsTable(sheet, sheetName);
        } else if (type == ExportType.CONFIG) {
            this.parseSheetAsConfig(sheet, sheetName);
        } else {
            throw new Error("错误的类型配置!");
        }
    }

    private async parseSheetAsTable(sheet: xlsx.WorkSheet, sheetName: string) {
        let obj = xlsx.utils.sheet_to_json<any>(sheet, { header: 0 });//转出来的json 0元素是excel第二行 类型，1元素是excel第三行 说明
        let types = obj[0];
        let realData = this.getTableRealJsonData(obj.slice(2), types);
        fs.writeFileSync(outJsonRoot + sheetName + ".json", realData);

        let interfaceName = this.getInterfaceName(sheetName);
        let codeFileStr = ExportTool.parser.parseTableTypes(types, interfaceName);
        let fileName = ExportTool.parser.getInterfaceFileName(interfaceName);
        fs.writeFileSync(outInterfaceRoot + fileName, codeFileStr);
    }

    private parseSheetAsConfig(sheet: xlsx.WorkSheet, sheetName: string) {
        let obj = xlsx.utils.sheet_to_json<any>(sheet, { header: "A" });//转出来的json 0元素是excel第二行 类型，1元素是excel第三行 说明

        //导出config数据
        let resultJsonData = this.getConfigRealJsonData(obj);
        fs.writeFileSync(outJsonRoot + sheetName + ".json", resultJsonData);

        //导出类型
        let interfaceName = this.getInterfaceName(sheetName);
        let codeFileStr = ExportTool.parser.parseConfigTypes(obj, interfaceName);
        let fileName = ExportTool.parser.getInterfaceFileName(interfaceName);
        fs.writeFileSync(outInterfaceRoot + fileName, codeFileStr);
    }

    private getTableRealJsonData(originData: any[], types: any) {
        let resultData: any = {};
        let props = Object.keys(types);

        for (let i = 0; i < props.length; i++) {
            const propName = props[i];
            const typeName = types[propName];
            for (let j = 0; j < originData.length; j++) {
                const originSingledata = originData[j];
                let originProperty = originSingledata[propName]
                if (originProperty !== null && originProperty !== undefined) {
                    let data = this.parseSingleJsonData(typeName, originProperty);
                    originData[j][propName] = data;
                }
            }
        }

        for (let i = 0; i < originData.length; i++) {
            const element = originData[i];
            let key = element[this._sheetData.key] as number | string;
            resultData[key] = element;
        }
        return JSON.stringify(resultData);
    }

    private getConfigRealJsonData(originData: any[]) {
        let resultData: any = {}
        for (let i = 0; i < originData.length; i++) {
            const element = originData[i];
            let propName = element['A'];
            let excelType = element['B'];
            let excelData = element['C'];
            let realData = this.parseSingleJsonData(excelType, excelData);
            resultData[propName] = realData;
        }
        return JSON.stringify(resultData);
    }

    private parseSingleJsonData(typeStr: string, property: any): any {
        if (typeof (property) === "number") {
            property = property.toString();
        }

        if (typeStr === "int") {
            return Number.parseInt(property);
        }
        if (typeStr === "float") {
            return Number.parseFloat(property);
        }
        if (typeStr === "string") {
            return property;
        }
        if (typeStr.endsWith('[]')) {
            let datas = [];
            let props = property.split(',');
            let singleType = typeStr.slice(0, -2);
            for (let i = 0; i < props.length; i++) {
                const prop = props[i];
                let data = this.parseSingleJsonData(singleType, prop);
                datas.push(data);
            }
            return datas;
        }
        return "";
    }

    public getInterfaceName(sheetName: string): string {
        return "Conf" + sheetName[0].toUpperCase() + sheetName.slice(1);
    }
}

export class ExportSheetData {
    excel: string;
    sheetName: string;
    type: ExportType;
    workSheet: xlsx.WorkSheet;
    key: string;
}

export enum ExportType {
    TABLE = "table",//表格形式
    CONFIG = "config",//配置形式
}