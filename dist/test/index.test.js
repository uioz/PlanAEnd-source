"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
require("mocha");
const chai_1 = require("chai");
const index_1 = require("../src/index");
const fs_1 = require("fs");
describe('StreamReadAsync测试', () => {
    describe('读取文件测试', () => {
        it('使用异步迭代测试', () => {
            return index_1.StreamReadAsync(fs_1.createReadStream('./test.xlsx', {
                autoClose: true,
            }));
        });
    });
});
describe('StreamReadPro测试', () => {
    describe('读取文件测试', () => {
        it('包装普通可读流为Promise', () => {
            return index_1.StreamReadPro(fs_1.createReadStream('./test.xlsx', {
                autoClose: true
            }));
        });
    });
});
describe('checkSourceData测试', () => {
    describe('功能测试', () => {
        it('一致性测试', () => {
            const WorkSheet = {
                "A1": {
                    "v": "name",
                    "t": "s"
                },
                "B1": {
                    "v": "number",
                    "t": "s"
                },
                "C1": {
                    "v": "speciality",
                    "t": "s"
                },
                "!ref": "A1:C1"
            };
            chai_1.expect(index_1.checkSourceData(WorkSheet)).eq(true);
        });
        it('非一致性测试', () => {
            const WorkSheet = {
                "A1": {
                    "v": "name",
                    "t": "s"
                },
                "B1": {
                    "v": "word",
                    "t": "s"
                },
                "C1": {
                    "v": "hello",
                    "t": "s"
                },
                "!ref": "A1:C1"
            };
            chai_1.expect(index_1.checkSourceData(WorkSheet)).eq(false);
        });
    });
    describe('函数式测试', () => {
        it('不修改传入的参数', () => {
            const WorkSheet = {
                "A1": {
                    "v": "name",
                    "t": "s"
                },
                "B1": {
                    "v": "word",
                    "t": "s"
                },
                "C1": {
                    "v": "hello",
                    "t": "s"
                },
                "!ref": "A1:C1"
            };
            index_1.checkSourceData(WorkSheet);
            chai_1.expect(WorkSheet).eql(WorkSheet);
        });
    });
});
describe('getDefaultSheets测试', () => {
    const baseWorkBook = {
        "Sheets": {
            "Sheet1": {
                "A1": {
                    "v": "name",
                    "t": "s"
                },
                "B1": {
                    "v": "number",
                    "t": "s"
                },
                "C1": {
                    "v": "specialit",
                    "t": "s"
                },
                "!ref": "A1:C1"
            }
        }
    };
    const workBook1 = Object.assign({ "SheetNames": [
            "Sheet1"
        ] }, baseWorkBook);
    const workBook2 = {
        "SheetNames": [
            "sheet1"
        ],
        "Sheets": {
            "sheet1": {
                "A1": {
                    "v": "name",
                    "t": "s"
                },
                "B1": {
                    "v": "number",
                    "t": "s"
                },
                "C1": {
                    "v": "specialit",
                    "t": "s"
                },
                "!ref": "A1:C1"
            }
        }
    };
    const NoneDeafultSheet = {
        SheetNames: ['abc'],
        Sheets: {
            'abc': {}
        }
    };
    describe('基本功能测试', () => {
        it('WorkSheet.Sheets.Sheet1 == Sheet1', () => chai_1.expect(index_1.getDefaultSheets(workBook1)).eql(workBook1.Sheets.Sheet1));
        it('WorkSheet.Sheets.sheet1 == sheet1', () => chai_1.expect(index_1.getDefaultSheets(workBook2)).eql(workBook2.Sheets.sheet1));
        it('没有默认工作表返回false测试', () => chai_1.expect(index_1.getDefaultSheets(NoneDeafultSheet)).eq(false));
    });
    describe('函数式测试', () => {
        it('不修改传入的参数', () => {
            const WorkBook = workBook1;
            index_1.getDefaultSheets(WorkBook);
            chai_1.expect(WorkBook).eql(WorkBook);
        });
    });
});
