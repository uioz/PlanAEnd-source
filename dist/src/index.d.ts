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
