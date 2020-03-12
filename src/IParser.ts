export default abstract class IParser {
    public parseTableTypes(types: any, interfaceName: string): string {
        let props = Object.keys(types);
        let bodyStr = "";
        let resultStr = this.getInterfaceTempStr();
        for (let index = 0; index < props.length; index++) {
            const prop = props[index];
            const excelType = types[prop];
            let propstr = this.getInterfacePropStr();
            let codeType = this.getInterfaceType(excelType);
            propstr = propstr.replace("{propName}", prop).replace("{typeName}", codeType);
            bodyStr += propstr;
        }
        return resultStr.replace("{interfaceName}", interfaceName).replace("{body}", bodyStr);
    }

    public parseConfigTypes(excel: any[], interfaceName: string) {
        let bodyStr = "";
        let resultStr = this.getInterfaceTempStr();
        for (let i = 0; i < excel.length; i++) {
            const element = excel[i];
            const propName = element["A"];
            const excelType = element["B"];
            let propStr = this.getInterfacePropStr();
            let codeType = this.getInterfaceType(excelType);
            propStr = propStr.replace("{propName}", propName).replace("{typeName}", codeType);
            bodyStr += propStr;
        }
        return resultStr.replace("{interfaceName}", interfaceName).replace("{body}", bodyStr);
    }

    protected abstract getInterfaceTempStr(): string;
    protected abstract getInterfacePropStr(): string;
    protected abstract getInterfaceType(excelType: string): string;
    public abstract getInterfaceFileName(interfaceName: string): string;
}

export class TsParser extends IParser {
    protected getInterfaceTempStr(): string {
        return "export interface {interfaceName} {\n{body}}"
    }

    protected getInterfacePropStr(): string {
        return "\t{propName}: {typeName};\n";
    }

    protected getInterfaceType(excelType: string): string {
        if (excelType == "int" || excelType == "float") {
            return "number";
        }
        if (excelType == "int[]" || excelType == "float[]") {
            return "number[]";
        }
        if (excelType == "string") {
            return "string";
        }
        if (excelType == "string[]") {
            return "string[]";
        }
        console.error("excel 数据类型解析错误");
        throw new Error("excel 数据类型解析错误");
    }

    public getInterfaceFileName(interfaceName: string): string {
        return interfaceName + ".ts";
    }
}

export class CSharpParser extends IParser {
    protected getInterfaceTempStr(): string {
        return "class {interfaceName} {\n{body}}"
    }

    protected getInterfacePropStr(): string {
        return "\tpublic {typeName} {propName};\n";
    }

    protected getInterfaceType(excelType: string): string {
        if (excelType == "int") {
            return "int";
        }
        if(excelType == "float"){
            return "float";
        }
        if (excelType == "int[]") {
            return "int[]";
        }
        if(excelType == "float[]"){
            return "float[]";
        }
        if (excelType == "string") {
            return "string";
        }
        if (excelType == "string[]") {
            return "string[]";
        }
        console.error("excel 数据类型解析错误");
        throw new Error("excel 数据类型解析错误");
    }

    public getInterfaceFileName(interfaceName: string): string {
        return interfaceName + ".ts";
    }
}