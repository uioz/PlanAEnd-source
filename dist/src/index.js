"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __asyncValues = (this && this.__asyncValues) || function (o) {
    if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
    var m = o[Symbol.asyncIterator], i;
    return m ? m.call(o) : (o = typeof __values === "function" ? __values(o) : o[Symbol.iterator](), i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i);
    function verb(n) { i[n] = o[n] && function (v) { return new Promise(function (resolve, reject) { v = o[n](v), settle(resolve, reject, v.done, v.value); }); }; }
    function settle(resolve, reject, d, v) { Promise.resolve(v).then(function(v) { resolve({ value: v, done: d }); }, reject); }
};
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
function StreamReadAsync(stream) {
    return __awaiter(this, void 0, void 0, function* () {
        var e_1, _a;
        const buffers = [];
        try {
            for (var _b = __asyncValues(stream), _c; _c = yield _b.next(), !_c.done;) {
                const chunk = _c.value;
                buffers.push(chunk);
            }
        }
        catch (e_1_1) { e_1 = { error: e_1_1 }; }
        finally {
            try {
                if (_c && !_c.done && (_a = _b.return)) yield _a.call(_b);
            }
            finally { if (e_1) throw e_1.error; }
        }
        return Buffer.concat(buffers);
    });
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
