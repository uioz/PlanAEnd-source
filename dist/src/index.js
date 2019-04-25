"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const index_1 = require("./utils/index");
/**
 * 解析选项
 */
exports.ParseOptions = {
    // 将日期类型数据解析为javascript类型
    // 这样做输出时候时间为文本类型不会被excel解析为数值型
    cellDates: true,
    cellHTML: false,
    cellText: false // 解析文本到额外的字段中
};
/**
 * 导出选项
 */
exports.WriteOptions = {
    // 禁止将日期类型进行转换
    // 这样可以保留日期的原来文本
    cellDates: false,
    // 开启zip压缩来减少文件体积
    compression: true,
    // 导出的文本类型
    // 使用xlsx类型兼容性好(相较于来说)
    bookType: 'xlsx'
};
/**
 * 工具函数将可以读流过程抽象为一个Promise
 * @param stream 可读流
 */
async function StreamReadAsync(stream) {
    const buffers = [];
    for await (const chunk of stream) {
        buffers.push(chunk);
    }
    return Buffer.concat(buffers);
}
exports.StreamReadAsync = StreamReadAsync;
/**
 * 工具函数,使用Promise包装的传统可读流的效果
 * @param stream 可读流
 */
function StreamReadPro(stream) {
    return new Promise((resolve, reject) => {
        const buffers = [];
        stream.on('data', data => buffers.push(data));
        stream.on('end', () => {
            resolve(Buffer.concat(buffers));
            stream.close();
        });
        stream.on('error', reject);
    });
}
exports.StreamReadPro = StreamReadPro;
/**
 * 校验工作表是否符合上传要求
 * @param workSheet 要被校验的工作表
 */
function checkSourceData(workSheet) {
    const authRange = 'A1:C1';
    let FieldMap;
    (function (FieldMap) {
        FieldMap["A1"] = "name";
        FieldMap["B1"] = "number";
        FieldMap["C1"] = "speciality";
    })(FieldMap || (FieldMap = {}));
    if (!index_1.inRange(workSheet['!ref'], authRange)) {
        return false;
    }
    for (const FieldsName of Object.keys(FieldMap)) {
        if (workSheet[FieldsName].v !== FieldMap[FieldsName]) {
            return false;
        }
    }
    return true;
}
exports.checkSourceData = checkSourceData;
/**
 * 从工作簿对象上获取默认的工作表对象
 * @param workBook 工作簿对象
 */
function getDefaultSheets(workBook) {
    for (const sheetName of ['Sheet1', 'sheet1']) {
        if (workBook.SheetNames.indexOf(sheetName) != -1) {
            return workBook.Sheets[sheetName];
        }
    }
    return false;
}
exports.getDefaultSheets = getDefaultSheets;
/**
 * 从给定的工作表中获取含有level字符串的键
 * 并且将这些键按照名称从小到大排序
 * @param workSheet 工作表对象
 */
function getLevelIndexs(workSheet) {
    const indexs = [], ref = workSheet['!ref'];
    let len = index_1.getColLen(workSheet);
    while (len--) {
        const key = workSheet[String.fromCharCode(65 + len) + '1'].v;
        if (key.search('level') !== -1) {
            indexs.push(key);
        }
    }
    return indexs.sort();
}
exports.getLevelIndexs = getLevelIndexs;
/**
 * 转换给定的JSON格式的工作表对象将多个含有level的键剔除.
 * 按照keys数组中的顺序添加到specialityPath数组中,在添加到JSON对象上.
 * 这个函数不会修改原来数组中的内容,将会返回一个新的JSON格式的工作表对象.
 * @param jsonizeWorkSheet JSON格式的工作表对象
 * @param keys 含有level键名组成的字符串数组
 */
function transformLevelToArray(jsonizeWorkSheet, keys) {
    const transform = (obj, keys) => {
        for (const key of keys) {
            if (obj[key]) {
                obj.specialityPath.push(obj[key]);
                delete obj[key];
            }
        }
        return obj;
    };
    if (keys.length) {
        return jsonizeWorkSheet.map((data) => transform(Object.assign({}, data, {
            'specialityPath': []
        }), keys));
    }
    else {
        return jsonizeWorkSheet.map((data) => Object.assign({}, data));
    }
}
exports.transformLevelToArray = transformLevelToArray;
/**
 * 过滤符合给定专业字段的数据,
 * 当源数据的speciality字段所含所有的数据是arrayizeWorkSheet中的一项的时候,
 * 则通过过滤.
 * @param arrayizeWorkSheet 数组化的工作表
 * @param Specialities 专业字段数组
 */
function correctSpeciality(arrayizeWorkSheet, Specialities) {
    const SpecialitiesSet = new Set(Specialities);
    return arrayizeWorkSheet.filter(item => SpecialitiesSet.has(item.speciality));
}
exports.correctSpeciality = correctSpeciality;
