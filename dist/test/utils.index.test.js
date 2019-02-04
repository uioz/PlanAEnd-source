"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
require("mocha");
const chai_1 = require("chai");
const index_1 = require("../src/utils/index");
describe('colEach测试', () => {
    const workSheet = {
        A1: { v: 1, t: 'n' },
        B1: { v: 2, t: 'n' },
        C1: { v: 3, t: 'n' },
        A2: { v: 4, t: 'n' },
        B2: { v: 5, t: 'n' },
        C2: { v: 6, t: 'n' },
        '!ref': 'A1:C2'
    };
    describe('迭代测试', () => {
        const countGenerator = (map) => {
            let index = 0;
            return () => map[++index];
        };
        const count1 = countGenerator({
            '1': 1,
            '2': 4,
            '3': 2,
            '4': 5,
            '5': 3,
            '6': 6,
        });
        const count2 = countGenerator({
            '1': 2,
            '2': 5,
            '3': 3,
            '4': 6,
        });
        it('迭代正确性测试1', () => index_1.colEach('A1:C2', workSheet, (content) => chai_1.expect(content.v).eq(count1())));
        it('迭代正确性测试2', () => index_1.colEach('B1:C2', workSheet, (content) => chai_1.expect(content.v).eq(count2())));
    });
    describe('函数式测试', () => {
        it('不修改传入的参数', () => {
            const range = 'A1:B2';
            index_1.colEach('A1:C2', workSheet, () => { });
            chai_1.expect(range).eq(range);
            // deep equal
            chai_1.expect(workSheet).eql(workSheet);
        });
    });
});
describe('rowEach测试', () => {
    const workSheet = {
        A1: { v: 1, t: 'n' },
        B1: { v: 2, t: 'n' },
        A2: { v: 3, t: 'n' },
        B2: { v: 4, t: 'n' },
        '!ref': 'A1:B2'
    };
    describe('迭代测试', () => {
        const countGenerator = (map) => {
            let index = 0;
            return () => map[++index];
        };
        const count1 = countGenerator({
            '1': 1,
            '2': 2,
            '3': 3,
            '4': 4
        });
        const count2 = countGenerator({
            '1': 1,
            '2': 3
        });
        it('迭代正确性测试1', () => index_1.rowEach('A1:B2', workSheet, (content) => chai_1.expect(content.v).eq(count1())));
        it('迭代正确性测试2', () => index_1.rowEach('A1:A2', workSheet, (content) => chai_1.expect(content.v).eq(count2())));
    });
    describe('函数式测试', () => {
        it('不修改传入的参数', () => {
            const range = 'A1:B2';
            index_1.rowEach(range, workSheet, () => { });
            chai_1.expect(range).eq(range);
            // deep equal
            chai_1.expect(workSheet).eql(workSheet);
        });
    });
});
describe('getColNumber测试', () => {
    describe('获取列号测试', () => {
        it('A1 == 0', () => {
            chai_1.expect(index_1.getColNumber('A1')).eq(0);
        });
        it('A10 == 0', () => {
            chai_1.expect(index_1.getColNumber('A10')).eq(0);
        });
        it('AB10 == 27', () => {
            chai_1.expect(index_1.getColNumber('AB10')).eq(27);
        });
    });
    describe('函数式测试', () => {
        it('不修改传入的参数', () => {
            const source = 'AB10';
            index_1.getColNumber(source);
            chai_1.expect(source).eq(source);
        });
    });
});
describe('getRowNumber测试', () => {
    describe('获取行号测试', () => {
        it('A100 == 99', () => {
            chai_1.expect(index_1.getRowNumber('A100')).eq(99);
        });
        it('C100 == 99', () => {
            chai_1.expect(index_1.getRowNumber('C100')).eq(99);
        });
        it('abc == NaN', () => {
            chai_1.expect(index_1.getRowNumber('abc')).to.be.NaN;
        });
        it('450 == NaN', () => {
            chai_1.expect(index_1.getRowNumber('450')).to.be.NaN;
        });
    });
    describe('函数式测试', () => {
        it('不修改传入的参数', () => {
            const source = 'A100';
            index_1.getRowNumber(source);
            chai_1.expect(source).eq(source);
        });
    });
});
describe('inrange测试', () => {
    describe('地址相等测试', () => {
        it(`'A1:B1','A1:B1'`, () => {
            chai_1.expect(index_1.inRange('A1:B1', 'A1:B1')).eq(true);
        });
        it(`"A1:B1", "A2:B2"==false`, () => {
            chai_1.expect(index_1.inRange("A1:B1", "A2:B2")).eq(false);
        });
        it(`"A2:B10", "A4:B7"`, () => {
            chai_1.expect(index_1.inRange("A2:B10", "A4:B7")).eq(true);
        });
        it(`"B2:D10", "C7:B8"`, () => {
            chai_1.expect(index_1.inRange("B2:D10", "C7:B8")).eq(true);
        });
        it(`"C7:B8", "B2:D10"==false`, () => {
            chai_1.expect(index_1.inRange("C7:B8", "B2:D10")).eq(false);
        });
    });
    describe('函数式测试', () => {
        it('不修改传入的参数', () => {
            const a = 'A1:B1', b = 'A1:C1';
            index_1.inRange(a, b);
            chai_1.expect(a).eq(a);
            chai_1.expect(b).eq(b);
        });
    });
});
describe('sliceRange测试', () => {
    describe('字符串切割', () => {
        it(`A1:B1 == ['A1','B1']`, () => {
            chai_1.expect(index_1.sliceRange('A1:B1')).eql(['A1', 'B1']);
        });
        it(`B2:D20 == ['B2','D20']`, () => {
            chai_1.expect(index_1.sliceRange('B2:D20')).eql(['B2', 'D20']);
        });
        it(`B2-D20 == ['B2-D20']`, () => {
            chai_1.expect(index_1.sliceRange('B2-D20')).eql(['B2-D20']);
        });
    });
    describe('函数式测试', () => {
        it('不修改传入的参数', () => {
            const source = 'B2:D20';
            index_1.sliceRange(source);
            chai_1.expect(source).eq(source);
        });
    });
});
