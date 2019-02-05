import { WorkSheet, CellObject } from "xlsx";
/**
 * 用于描述单元格迭代函数的接口
 */
export interface CellEachInterface {
    (target: string, data: WorkSheet, callback: (result: CellObject, index: string, data: WorkSheet) => void): void;
}
/**
 * 用于测试一个范围(target)是否在另外一个范围(source)内
 * @param source 被比较的范围
 * @param target 比较的范围
 */
export declare const inRange: (source: string, target: string) => boolean;
/**
 * 将单元格切割为一个数组 例如'A1:B1'最后的结果是['A1','B1']
 * @param cellRange 单元格地址对象
 */
export declare const sliceRange: (cellRange: string) => string[];
/**
 * 获取单元格坐标的行数例如A100的行数是99
 * @param cellPosition 单元格坐标
 */
export declare const getRowNumber: (cellPosition: string) => number;
/**
 * 获取单元格列号例如AB10这个单位中的列是AB则对应的列号是27(下标从1开始)
 * @param collPosition 单元格坐标
 */
export declare const getColNumber: (collPosition: string) => number;
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
export declare const rowEach: CellEachInterface;
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
export declare const colEach: CellEachInterface;
