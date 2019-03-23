"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const xlsx_1 = require("xlsx");
/**
 * 用于测试一个范围(target)是否在另外一个范围(source)内
 * @param source 被比较的范围
 * @param target 比较的范围
 */
exports.inRange = (source, target) => {
    const sourceRagne = xlsx_1.utils.decode_range(source), targetRange = xlsx_1.utils.decode_range(target);
    if (targetRange.s.r >= sourceRagne.s.r && targetRange.s.c >= sourceRagne.s.c &&
        targetRange.e.r <= sourceRagne.e.r && targetRange.e.c <= sourceRagne.e.c) {
        return true;
    }
    return false;
};
/**
 * 将单元格切割为一个数组 例如'A1:B1'最后的结果是['A1','B1']
 * @param cellRange 单元格地址对象
 */
exports.sliceRange = (cellRange) => cellRange.split(':');
/**
 * 获取单元格坐标的行数例如A100的行数是99
 * @param cellPosition 单元格坐标
 */
exports.getRowNumber = (cellPosition) => (cellPosition.match(/[^A-Za-z]+/g) || [])[0] === cellPosition ? NaN : xlsx_1.utils.decode_row(cellPosition.replace(/^[A-Za-z]+/g, ''));
/**
 * 获取单元格列号例如AB10这个单位中的列是AB则对应的列号是27(下标从1开始)
 * @param collPosition 单元格坐标
 */
exports.getColNumber = (collPosition) => {
    try {
        return xlsx_1.utils.decode_col(/^[A-Za-z]+/g.exec(collPosition)[0]);
    }
    catch (error) {
        return NaN;
    }
};
/**
 * 基本迭代函数被其他两个迭代函数依赖
 * @param range 迭代的范围 例如A1:B10
 * @param source 被迭代工作表对象
 * @param callback 回调函数
 * @param dir 迭代方向 行优先或者列优先 默认=true(行优先)
 */
const cellEachBase = (range, source, callback, dir = true) => {
    if (!callback) {
        return;
    }
    let { s: { r: RowStartIndex, c: ColStartIndex }, e: { r: RowEndIndex, c: ColEndIndex } } = xlsx_1.utils.decode_range(range);
    if (dir) {
        while (RowStartIndex <= RowEndIndex) {
            let colIndex = ColStartIndex;
            while (colIndex <= ColEndIndex) {
                const cellIndex = xlsx_1.utils.encode_cell({ c: colIndex, r: RowStartIndex });
                callback(source[cellIndex], cellIndex, source);
                colIndex++;
            }
            RowStartIndex++;
        }
    }
    else {
        while (ColStartIndex <= ColEndIndex) {
            let rowIndex = RowStartIndex;
            while (rowIndex <= RowEndIndex) {
                const cellIndex = xlsx_1.utils.encode_cell({ c: ColStartIndex, r: rowIndex });
                callback(source[cellIndex], cellIndex, source);
                rowIndex++;
            }
            ColStartIndex++;
        }
    }
};
/**
 * 按照行优先的策略进行迭代数据,对于如下数据:
 * ```
 * 1 2
 * 3 4
 * ```
 * 完整迭代顺序为:
 * 1. 1
 * 2. 2
 * 3. 3
 * 4. 4
 * @param range 行迭代的范围 例如 A1:B10
 * @param source 被迭代的工作表对象
 * @param callback 回调函数
 */
exports.rowEach = (range, source, callback) => cellEachBase(range, source, callback, true);
/**
 * 按照列迭代优先的策略,有如下数据:
 * ```
 * 1 2
 * 3 4
 * ```
 * 完整迭代顺序为:
 * 1. 1
 * 2. 3
 * 3. 2
 * 4. 4
 * @param range 列迭代的范围 例如: A1:B10
 * @param source 被迭代的工作表对象
 * @param callback 回调函数
 */
exports.colEach = (range, source, callback) => cellEachBase(range, source, callback, false);
/**
 * 从给定的工作表中获取行数
 * @param workSheet 工作表对象
 */
exports.getRowLen = (workSheet) => exports.getRowNumber(exports.sliceRange(workSheet['!ref'])[1]);
/**
 * 从给定的工作表中获取列数
 */
exports.getColLen = (workSheet) => exports.getColNumber(exports.sliceRange(workSheet['!ref'])[1]);
