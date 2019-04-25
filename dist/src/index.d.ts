/// <reference types="node" />
import { WorkSheet, WorkBook, ParsingOptions, WritingOptions } from 'xlsx';
import { ReadStream } from "fs";
/**
 * 解析选项
 */
export declare const ParseOptions: ParsingOptions;
/**
 * 导出选项
 */
export declare const WriteOptions: WritingOptions;
/**
 * 工具函数将可以读流过程抽象为一个Promise
 * @param stream 可读流
 */
export declare function StreamReadAsync(stream: ReadStream): Promise<Buffer>;
/**
 * 工具函数,使用Promise包装的传统可读流的效果
 * @param stream 可读流
 */
export declare function StreamReadPro(stream: ReadStream): Promise<Buffer>;
/**
 * 校验工作表是否符合上传要求
 * @param workSheet 要被校验的工作表
 */
export declare function checkSourceData(workSheet: WorkSheet): boolean;
/**
 * 从工作簿对象上获取默认的工作表对象
 * @param workBook 工作簿对象
 */
export declare function getDefaultSheets(workBook: WorkBook): WorkSheet | false;
/**
 * 从给定的工作表中获取含有level字符串的键
 * 并且将这些键按照名称从小到大排序
 * @param workSheet 工作表对象
 */
export declare function getLevelIndexs(workSheet: WorkSheet): any[];
/**
 * 转换给定的JSON格式的工作表对象将多个含有level的键剔除.
 * 按照keys数组中的顺序添加到specialityPath数组中,在添加到JSON对象上.
 * 这个函数不会修改原来数组中的内容,将会返回一个新的JSON格式的工作表对象.
 * @param jsonizeWorkSheet JSON格式的工作表对象
 * @param keys 含有level键名组成的字符串数组
 */
export declare function transformLevelToArray(jsonizeWorkSheet: Array<object>, keys: Array<string>): Array<object>;
/**
 * 过滤符合给定专业字段的数据,
 * 当源数据的speciality字段所含所有的数据是arrayizeWorkSheet中的一项的时候,
 * 则通过过滤.
 * @param arrayizeWorkSheet 数组化的工作表
 * @param Specialities 专业字段数组
 */
export declare function correctSpeciality(arrayizeWorkSheet: Array<any>, Specialities: Array<string>): Array<any>;
