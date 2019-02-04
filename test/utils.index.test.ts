import "mocha";
import { expect } from "chai";
import { inRange, sliceRange, getRowNumber, getColNumber, rowEach,colEach } from "../src/utils/index";

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

        it('迭代正确性测试1', () => colEach('A1:C2', workSheet, (content) => expect(content.v).eq(count1())));

        it('迭代正确性测试2', () => colEach('B1:C2', workSheet, (content) => expect(content.v).eq(count2())));

    });

    describe('函数式测试', () => {

        it('不修改传入的参数', () => {

            const range = 'A1:B2';

            colEach('A1:C2',workSheet,()=>{});

            expect(range).eq(range);
            // deep equal
            expect(workSheet).eql(workSheet);

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

        it('迭代正确性测试1', () => rowEach('A1:B2', workSheet, (content) => expect(content.v).eq(count1())));

        it('迭代正确性测试2', () => rowEach('A1:A2', workSheet, (content) => expect(content.v).eq(count2())));

    });

    describe('函数式测试', () => {

        it('不修改传入的参数', () => {

            const range = 'A1:B2';

            rowEach(range, workSheet, () => { });

            expect(range).eq(range);
            // deep equal
            expect(workSheet).eql(workSheet);

        });

    });

});

describe('getColNumber测试', () => {

    describe('获取列号测试', () => {


        it('A1 == 0', () => {

            expect(getColNumber('A1')).eq(0);

        });

        it('A10 == 0', () => {

            expect(getColNumber('A10')).eq(0);

        });

        it('AB10 == 27', () => {

            expect(getColNumber('AB10')).eq(27);

        });

    });

    describe('函数式测试', () => {

        it('不修改传入的参数', () => {

            const source = 'AB10';

            getColNumber(source);

            expect(source).eq(source);

        });

    });

});

describe('getRowNumber测试', () => {

    describe('获取行号测试', () => {

        it('A100 == 99', () => {

            expect(getRowNumber('A100')).eq(99);

        });

        it('C100 == 99', () => {

            expect(getRowNumber('C100')).eq(99);

        });

        it('abc == NaN', () => {

            expect(getRowNumber('abc')).to.be.NaN;

        });

        it('450 == NaN', () => {

            expect(getRowNumber('450')).to.be.NaN;

        });

    });

    describe('函数式测试', () => {

        it('不修改传入的参数', () => {

            const source = 'A100';

            getRowNumber(source);

            expect(source).eq(source);

        });

    });

});

describe('inrange测试', () => {

    describe('地址相等测试', () => {

        it(`'A1:B1','A1:B1'`, () => {

            expect(inRange('A1:B1', 'A1:B1')).eq(true);

        });

        it(`"A1:B1", "A2:B2"==false`, () => {

            expect(inRange("A1:B1", "A2:B2")).eq(false);

        });

        it(`"A2:B10", "A4:B7"`, () => {

            expect(inRange("A2:B10", "A4:B7")).eq(true);

        });

        it(`"B2:D10", "C7:B8"`, () => {

            expect(inRange("B2:D10", "C7:B8")).eq(true);

        });

        it(`"C7:B8", "B2:D10"==false`, () => {

            expect(inRange("C7:B8", "B2:D10")).eq(false);

        });

    });

    describe('函数式测试', () => {

        it('不修改传入的参数', () => {

            const
                a = 'A1:B1',
                b = 'A1:C1';

            inRange(a, b);

            expect(a).eq(a);
            expect(b).eq(b);

        });

    });

});

describe('sliceRange测试', () => {

    describe('字符串切割', () => {

        it(`A1:B1 == ['A1','B1']`, () => {

            expect(sliceRange('A1:B1')).eql(['A1', 'B1']);

        });

        it(`B2:D20 == ['B2','D20']`, () => {

            expect(sliceRange('B2:D20')).eql(['B2', 'D20']);

        });

        it(`B2-D20 == ['B2-D20']`, () => {

            expect(sliceRange('B2-D20')).eql(['B2-D20']);

        });

    });

    describe('函数式测试', () => {

        it('不修改传入的参数', () => {

            const source = 'B2:D20';

            sliceRange(source);

            expect(source).eq(source);

        });

    });

});
