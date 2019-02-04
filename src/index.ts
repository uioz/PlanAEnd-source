import { rowEach,inRange } from "./utils/index";
import { utils, WorkSheet, WorkBook , ParsingOptions,WritingOptions } from 'xlsx';
import { ReadStream } from "fs";

/**
 * 解析选项
 */
export const ParseOptions: ParsingOptions = {
    // 将日期类型数据解析为javascript类型
    // 这样做输出时候时间为文本类型不会被excel解析为数值型
    cellDates: true,
    cellHTML:false, // 禁止富文本添新的字段
    cellText:false // 解析文本到额外的字段中
};

/**
 * 导出选项
 */
export const WriteOptions: WritingOptions = {
    // 禁止将日期类型进行转换
    // 这样可以保留日期的原来文本
    cellDates:false,
    // 开启zip压缩来减少文件体积
    compression:true,
    // 导出的文本类型
    // 使用xlsx类型兼容性好(相较于来说)
    bookType:'xlsx'
};

/**
 * 工具函数将可以读流过程抽象为一个Promise
 * @param stream 可读流
 */
export async function StreamReadAsync(stream: ReadStream){

    const buffers = [];

    for await (const chunk of stream as any) {
        buffers.push(chunk);
    }

    return Buffer.concat(buffers);

}

/**
 * 工具函数,使用Promise包装的传统可读流的效果
 * @param stream 可读流
 */
export function StreamReadPro(stream: ReadStream):Promise<Buffer> {
    return new Promise((resolve,reject)=>{

        const buffers = [];

        stream.on('data', data => buffers.push(data));
        stream.on('end', () => {
            resolve(Buffer.concat(buffers));
            stream.close();
        });
        stream.on('error',reject);

    });
}

/**
 * 校验工作表是否符合上传要求
 * @param workSheet 要被校验的工作表
 */
export function checkSourceData(workSheet:WorkSheet):boolean{
    
    const 
        authRange = 'A1:C1';
    
    enum FieldMap {
        'A1' = 'name',
        'B1' = 'number',
        'C1' = 'speciality'
    }

    if(!inRange(workSheet['!ref'],authRange)){
        return false;
    }

    for (const FieldsName of Object.keys(FieldMap)) {
        if (workSheet[FieldsName].v !== FieldMap[FieldsName]){
            return false;
        }
    }
    
    return true;

}

/**
 * 从工作簿对象上获取默认的工作表对象
 * @param workBook 工作簿对象
 */
export function getDefaultSheets(workBook:WorkBook):WorkSheet|false {

    for (const sheetName of ['Sheet1','sheet1']) {
        if(workBook.SheetNames.indexOf(sheetName) != -1){
            return workBook.Sheets[sheetName];
        }
    }

    return false;
}
