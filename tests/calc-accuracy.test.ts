import ExcelJS from 'exceljs';
import { beforeAll, describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import type { CalcReport } from '../src/core/types';

/**
 * Formulas with the results Excel gives, stored as each cell's saved value,
 * the way Excel writes them. Opening the workbook compares the engine with
 * every one, exactly as it does for real files, so the report says where the
 * engine and Excel agree. Data!A1:A10 = 1..10; B1:B5 = apple, Banana, cherry,
 * apple, date; E1:F4 = a 1, b 2, c 3, d 4; E6 is empty.
 */
type Expected = number | string | boolean | { error: string };
const CASES: [string, Expected][] = [
  // Arithmetic and rounding
  ['SUM(1,2,3)', 6],
  ['SUM("1",2)', 3],
  ['10/3', 10 / 3],
  ['0.1+0.2=0.3', true],
  ['ROUND(2.5,0)', 3],
  ['ROUND(-2.5,0)', -3],
  ['ROUND(1.005,2)', 1.01],
  ['ROUNDUP(3.14159,2)', 3.15],
  ['ROUNDDOWN(-3.14159,1)', -3.1],
  ['MROUND(10,3)', 9],
  ['CEILING(4.2,1)', 5],
  ['FLOOR(4.8,1)', 4],
  ['INT(-1.5)', -2],
  ['INT(5.9)', 5],
  ['ROUND(1234.5,-2)', 1200],
  ['ROUNDUP(-1.21,1)', -1.3],
  ['ROUNDDOWN(1.29,1)', 1.2],
  ['ROUND(2.675,2)', 2.68],
  ['MOD(5.5,2)', 1.5],
  ['MOD(-7,-3)', -1],
  ['MOD(1,0)', { error: '#DIV/0!' }],
  ['TRUNC(-1.5)', -1],
  ['MOD(-3,2)', 1],
  ['MOD(3,-2)', -1],
  ['POWER(2,10)', 1024],
  ['SQRT(16)', 4],
  ['PRODUCT(2,3,4)', 24],
  ['QUOTIENT(7,2)', 3],
  ['Data!E6+1', 1],
  // Aggregates and conditions
  ['SUMPRODUCT(Data!A1:A3,Data!A2:A4)', 20],
  ['SUMIF(Data!A1:A10,">5")', 40],
  ['SUMIFS(Data!A1:A10,Data!A1:A10,">=3",Data!A1:A10,"<=5")', 12],
  ['COUNTIF(Data!B1:B5,"apple")', 2],
  ['COUNTIF(Data!B1:B5,"a*")', 2],
  ['COUNTIF(Data!B1:B5,"*an*")', 1],
  ['COUNTA(Data!B1:B5)', 5],
  ['COUNTBLANK(Data!E5:E6)', 2],
  ['AVERAGE(Data!A1:A10)', 5.5],
  ['MEDIAN(1,2,3,4)', 2.5],
  ['STDEV.S(Data!A1:A10)', 3.0276503540974917],
  ['LARGE(Data!A1:A10,2)', 9],
  ['SMALL(Data!A1:A10,3)', 3],
  ['RANK.EQ(7,Data!A1:A10)', 4],
  ['RANK.EQ(7,Data!A1:A10,1)', 7],
  ['RANK(3,Data!A1:A10)', 8],
  ['RANK.EQ(11,Data!A1:A10)', { error: '#N/A' }],
  // Text
  ['LEFT("Hello",2)', 'He'],
  ['MID("Hello",2,3)', 'ell'],
  ['RIGHT("Hello",3)', 'llo'],
  ['LEN("Hello")', 5],
  ['UPPER("abc")', 'ABC'],
  ['PROPER("hello world")', 'Hello World'],
  ['TRIM("  a  b  ")', 'a b'],
  ['SUBSTITUTE("a-b-c","-","+")', 'a+b+c'],
  ['SUBSTITUTE("a-b-c","-","+",2)', 'a-b+c'],
  ['FIND("l","Hello")', 3],
  ['SEARCH("L","Hello")', 3],
  ['TEXT(1234.567,"#,##0.00")', '1,234.57'],
  ['TEXT(0.256,"0.0%")', '25.6%'],
  ['TEXT(DATE(2026,3,1),"yyyy-mm-dd")', '2026-03-01'],
  ['TEXT(DATE(2026,3,1),"dddd")', 'Sunday'],
  ['TEXT(-1234.5,"#,##0;(#,##0)")', '(1,235)'],
  ['TEXT(45352,"d mmm yyyy")', '1 Mar 2024'],
  ['TEXT(0.75,"h:mm AM/PM")', '6:00 PM'],
  ['TEXT("abc","@")', 'abc'],
  ['CONCATENATE("a","b",1)', 'ab1'],
  ['"a"&1&TRUE', 'a1TRUE'],
  ['VALUE("12.5")', 12.5],
  ['REPT("ab",3)', 'ababab'],
  ['EXACT("a","A")', false],
  // Logic and errors
  ['IF(1>2,"y","n")', 'n'],
  ['AND(TRUE,1)', true],
  ['OR(FALSE,0)', false],
  ['NOT(0)', true],
  ['IFERROR(1/0,"div")', 'div'],
  ['IFS(FALSE,1,TRUE,2)', 2],
  ['SWITCH(2,1,"a",2,"b")', 'b'],
  ['ISBLANK(Data!E6)', true],
  ['ISNUMBER("1")', false],
  ['1/0', { error: '#DIV/0!' }],
  ['NA()', { error: '#N/A' }],
  ['"a"+1', { error: '#VALUE!' }],
  // Dates (serial numbers in the 1900 system)
  ['DATE(2024,2,30)', 45352],
  ['DATE(2026,1,1)', 46023],
  ['YEAR(46023)', 2026],
  ['MONTH(45352)', 3],
  ['DAY(45352)', 1],
  ['EOMONTH(DATE(2026,1,15),1)', 46081],
  ['EDATE(DATE(2026,1,31),1)', 46081],
  ['WEEKDAY(DATE(2026,3,1))', 1],
  ['WEEKDAY(DATE(2026,3,1),2)', 7],
  ['NETWORKDAYS(DATE(2026,3,2),DATE(2026,3,13))', 10],
  ['WORKDAY(DATE(2026,3,6),1)', 46090],
  ['DATEDIF(DATE(2020,1,15),DATE(2026,3,1),"y")', 6],
  ['DATEDIF(DATE(2026,1,31),DATE(2026,3,1),"m")', 1],
  ['DATEDIF(DATE(2020,1,15),DATE(2026,3,1),"ym")', 1],
  ['DATEDIF(DATE(2025,6,15),DATE(2026,3,1),"yd")', 259],
  ['DATEDIF(DATE(2026,3,1),DATE(2026,1,1),"d")', { error: '#NUM!' }],
  ['DAYS(DATE(2026,3,1),DATE(2026,1,1))', 59],
  ['YEARFRAC(DATE(2026,1,1),DATE(2026,7,1))', 0.5],
  ['TIME(12,30,0)', 0.5208333333333334],
  ['HOUR(0.75)', 18],
  // Lookups
  ['VLOOKUP("c",Data!E1:F4,2,FALSE)', 3],
  ['VLOOKUP("C",Data!E1:F4,2,FALSE)', 3],
  ['INDEX(Data!F1:F4,MATCH("b",Data!E1:E4,0))', 2],
  ['MATCH(5,{1,3,5,7},0)', 3],
  ['MATCH(6,{1,3,5,7},1)', 3],
  ['CHOOSE(2,"a","b","c")', 'b'],
  ['XLOOKUP("c",Data!E1:E4,Data!F1:F4)', 3],
  // Finance (Excel's documented examples)
  ['PMT(0.05/12,360,200000)', -1073.6432460242797],
  ['FV(0.06/12,10,-200,-500,1)', 2581.4033740601],
  ['NPV(0.1,-10000,3000,4200,6800)', 1188.4434123352207],
  ['IRR({-70000,12000,15000,18000,21000})', -0.021244848272],
  // Functions the engine lacks, added for Excel's files; the rest use Excel's documented examples (Examples sheet).
  ['LOOKUP(4.19,Examples!A1:A5,Examples!B1:B5)', 'orange'],
  ['LOOKUP(5.75,Examples!A1:A5,Examples!B1:B5)', 'yellow'],
  ['LOOKUP(7.66,Examples!A1:A5,Examples!B1:B5)', 'blue'],
  ['LOOKUP(0,Examples!A1:A5,Examples!B1:B5)', { error: '#N/A' }],
  ['LOOKUP("C",{"a","b","c","d";1,2,3,4})', 3],
  ['LOOKUP("bump",{"a",1;"b",2;"c",3})', 2],
  ['XMATCH("Grape",{"Apple","Grape","Pear","Banana","Cherry"})', 2],
  ['XMATCH(4.5,{5,4,3,2,1},1)', 1],
  ['CONCAT("The"," ","sun"," ","will"," ","come"," ","up"," ","tomorrow.")', 'The sun will come up tomorrow.'],
  ['NUMBERVALUE("2.500,27",",",".")', 2500.27],
  ['NUMBERVALUE("3.5%")', 0.035],
  ['ARRAYTOTEXT(TEXTSPLIT("Dakota Lennon Sanchez"," "),1)', '{"Dakota","Lennon","Sanchez"}'],
  ['ARRAYTOTEXT(TEXTSPLIT("1,2,3;4,5,6",",",";"),1)', '{"1","2","3";"4","5","6"}'],
  ['TEXTBEFORE(Examples!C1,"Red")', 'Little '],
  ['TEXTBEFORE(Examples!C1,"basket")', { error: '#N/A' }],
  ['TEXTAFTER(Examples!C1,"Red")', " Riding Hood's red hood"],
  ['TEXTAFTER(Examples!C1,"basket",1,,,"Not found")', 'Not found'],
  ['VALUETOTEXT(Examples!D2)', '1234.01234'],
  ['VALUETOTEXT(Examples!D3,1)', '"Hello"'],
  ['VALUETOTEXT(Examples!E1)', '#VALUE!'],
  ['ARRAYTOTEXT(Examples!D1:E3,0)', 'TRUE, #VALUE!, 1234.01234, Seattle, Hello, 1123'],
  ['ARRAYTOTEXT(Examples!D1:E3,1)', '{TRUE,#VALUE!;1234.01234,"Seattle";"Hello",1123}'],
  ['FIXED(1234.567,1)', '1,234.6'],
  ['FIXED(1234.567,-1)', '1,230'],
  ['FIXED(-1234.567,-1,TRUE)', '-1230'],
  ['FIXED(44.332)', '44.33'],
  ['TYPE(Examples!F1)', 2],
  ['TYPE("MR. "&Examples!F1)', 2],
  ['TYPE(2+Examples!F1)', 16],
  ['TYPE({1,2;3,4})', 64],
  ['ERROR.TYPE(1/0)', 2],
  ['ERROR.TYPE(NA())', 7],
  // Array results, written out by ARRAYTOTEXT to fit one cell.
  ['ARRAYTOTEXT(TAKE({1,2,3;4,5,6;7,8,9},2),1)', '{1,2,3;4,5,6}'],
  ['ARRAYTOTEXT(TAKE({1,2,3;4,5,6;7,8,9},,2),1)', '{1,2;4,5;7,8}'],
  ['ARRAYTOTEXT(TAKE({1,2,3;4,5,6;7,8,9},-2),1)', '{4,5,6;7,8,9}'],
  ['ARRAYTOTEXT(TAKE({1,2,3;4,5,6;7,8,9},2,2),1)', '{1,2;4,5}'],
  ['ARRAYTOTEXT(DROP({1,2,3;4,5,6;7,8,9},2),1)', '{7,8,9}'],
  ['ARRAYTOTEXT(DROP({1,2,3;4,5,6;7,8,9},,2),1)', '{3;6;9}'],
  ['ARRAYTOTEXT(DROP({1,2,3;4,5,6;7,8,9},-2),1)', '{1,2,3}'],
  ['DROP({1,2,3;4,5,6;7,8,9},2,2)', 9],
  ['ARRAYTOTEXT(CHOOSEROWS({1,2;3,4;5,6;7,8;9,10;11,12},1,3,5,1),1)', '{1,2;5,6;9,10;1,2}'],
  ['ARRAYTOTEXT(CHOOSEROWS({1,2;3,4;5,6;7,8;9,10;11,12},{1,3},5,1),1)', '{1,2;5,6;9,10;1,2}'],
  ['ARRAYTOTEXT(CHOOSEROWS({1,2;3,4;5,6;7,8;9,10;11,12},-1,-2),1)', '{11,12;9,10}'],
  ['ARRAYTOTEXT(CHOOSECOLS({1,2,3,4,5;6,7,8,9,10},1,3,5,1),1)', '{1,3,5,1;6,8,10,6}'],
  ['ARRAYTOTEXT(CHOOSECOLS({1,2,3,4,5;6,7,8,9,10},3,-2),1)', '{3,4;8,9}'],
  [
    'ARRAYTOTEXT(TOROW({"Ben","Peter","Mary","Sam";"John","Hillary","Jenny","James";"Agnes","Harry","Felicity","Joe"},,TRUE),1)',
    '{"Ben","John","Agnes","Peter","Hillary","Harry","Mary","Jenny","Felicity","Sam","James","Joe"}',
  ],
  ['ARRAYTOTEXT(TOCOL({"Ben","Peter";"John","Hillary"}),1)', '{"Ben";"Peter";"John";"Hillary"}'],
  ['ARRAYTOTEXT(WRAPROWS({"A","B","C","D","E","F","G"},3),1)', '{"A","B","C";"D","E","F";"G",#N/A,#N/A}'],
  ['ARRAYTOTEXT(WRAPROWS({"A","B","C","D","E","F","G"},3,"x"),1)', '{"A","B","C";"D","E","F";"G","x","x"}'],
  ['ARRAYTOTEXT(WRAPCOLS({"A","B","C","D","E","F","G"},3),1)', '{"A","D","G";"B","E",#N/A;"C","F",#N/A}'],
  ['ARRAYTOTEXT(EXPAND({1,2;3,4},3,3),1)', '{1,2,#N/A;3,4,#N/A;#N/A,#N/A,#N/A}'],
  ['ARRAYTOTEXT(EXPAND({1,2;3,4},3,3,"-"),1)', '{1,2,"-";3,4,"-";"-","-","-"}'],
  ['ARRAYTOTEXT(SORTBY({"Tom";"Fred";"Amy";"Sal";"Fritz";"Xio"},{52;65;22;73;19;39}),1)', '{"Fritz";"Amy";"Xio";"Tom";"Fred";"Sal"}'],
  ['ARRAYTOTEXT(VSTACK({1,2},{3}),1)', '{1,2;3,#N/A}'],
  ['ARRAYTOTEXT(HSTACK({1;2},3),1)', '{1,3;2,#N/A}'],
  // Statistics
  ['AVERAGEIFS(Examples!H1:H4,Examples!H1:H4,">70",Examples!H1:H4,"<90")', 80.5],
  ['AVERAGEIFS(Examples!I1:I4,Examples!I1:I4,">95")', { error: '#DIV/0!' }],
  ['AVERAGEIFS(Examples!J1:J4,Examples!J1:J4,"<>Incomplete",Examples!J1:J4,">80")', 87.5],
  ['AVERAGEIFS(Examples!L1:L6,Examples!M1:M6,"Bellevue",Examples!N1:N6,">2",Examples!O1:O6,"Yes")', 397839],
  ['AVERAGEIFS(Examples!L1:L6,Examples!M1:M6,"Issaquah",Examples!N1:N6,"<=3",Examples!O1:O6,"No")', 230000],
  ['MODE(5.6,4,4,3,2,4)', 4],
  ['MODE.SNGL(5.6,4,4,3,2,4)', 4],
  ['ARRAYTOTEXT(MODE.MULT({1,2,3,4,3,2,1,2,3,5,6,1}),1)', '{1;2;3}'],
  ['ARRAYTOTEXT(FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89}),1)', '{1;2;4;2}'],
  ['ROUND(TRIMMEAN({4,5,6,7,2,3,4,5,1,2,3},0.2),9)', 3.777777778],
  ['ROUND(KURT(3,4,5,2,3,4,5,6,4,7),9)', -0.151799637],
  ['PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},2)', 0.333],
  ['PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},4)', 0.555],
  ['PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},8)', 0.666],
  ['PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},5)', 0.583],
  ['PERCENTRANK({13,12,11,8,4,3,2,1,1,1},2)', 0.333],
  ['PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},7)', 0.7],
  ['PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},5.43)', 0.381],
  ['PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},5.43,1)', 0.3],
  ['PERMUT(100,3)', 970200],
  ['PERMUT(3,2)', 6],
  ['PERMUTATIONA(3,2)', 9],
  ['PERMUTATIONA(2,2)', 4],
  ['RANK.AVG(94,{89;88;92;101;94;97;95})', 4],
  ['PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)', 0.1],
  ['PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)', 0.8],
  ['GESTEP(5,4)', 1],
  ['GESTEP(5,5)', 1],
  ['GESTEP(-4,-5)', 1],
  ['GESTEP(-1)', 0],
  ['ROUND(ERF.PRECISE(0.745),6)', 0.707929],
  ['ROUND(ERF.PRECISE(1),7)', 0.8427008],
  ['ROUND(ERFC.PRECISE(1),7)', 0.1572992],
  ['ROUND(BINOM.DIST.RANGE(60,0.75,48),3)', 0.084],
  ['ROUND(BINOM.DIST.RANGE(60,0.75,45,50),3)', 0.524],
  ['MDETERM({1,3,8,5;1,3,6,1;1,1,1,0;7,3,10,2})', 88],
  ['MDETERM({3,6,1;1,1,0;3,10,2})', 1],
  ['MDETERM({3,6;1,1})', -3],
  ['MDETERM({1,3,8,5;1,3,6,1})', { error: '#VALUE!' }],
  ['ARRAYTOTEXT(MINVERSE({4,-1;2,0}),1)', '{0,0.5;-1,2}'],
  ['ARRAYTOTEXT(MINVERSE({1,2,1;3,4,-1;0,2,0}),1)', '{0.25,0.25,-0.75;0,0,0.5;0.75,-0.25,-0.25}'],
  ['ARRAYTOTEXT(MUNIT(3),1)', '{1,0,0;0,1,0;0,0,1}'],
  ['AGGREGATE(4,6,Examples!U1:U11)', 96],
  ['AGGREGATE(14,6,Examples!U1:U11,3)', 72],
  ['AGGREGATE(15,6,Examples!U1:U11)', { error: '#VALUE!' }],
  ['AGGREGATE(12,6,Examples!U1:U11,Examples!V1:V11)', 68],
  // Regression (Microsoft's office buildings for LINEST)
  ['ROUND(FORECAST(30,{6;7;9;15;21},{20;28;31;38;40}),6)', 10.607253],
  ['ROUND(FORECAST.LINEAR(30,{6;7;9;15;21},{20;28;31;38;40}),6)', 10.607253],
  ['ROUND(INTERCEPT({2;3;9;1;8},{6;5;11;7;5}),7)', 0.0483871],
  ['ROUND(INDEX(TREND({133890;135000;135790;137300;138130;139100;139900;141120;141890;143230;144000;145290},{1;2;3;4;5;6;7;8;9;10;11;12}),1),4)', 133953.3333],
  ['ROUND(INDEX(TREND({133890;135000;135790;137300;138130;139100;139900;141120;141890;143230;144000;145290},{1;2;3;4;5;6;7;8;9;10;11;12},{13;14;15;16;17}),5),4)', 150244.2424],
  ['ROUND(TREND(Examples!T1:T11,Examples!P1:S11,{2500,3,2,25}),2)', 158261.1],
  ['ROUND(INDEX(GROWTH({33100;47300;69000;102000;150000;220000},{11;12;13;14;15;16}),1),5)', 32618.20377],
  ['ROUND(INDEX(GROWTH({33100;47300;69000;102000;150000;220000},{11;12;13;14;15;16},{17;18}),1),4)', 320196.7184],
  ['ROUND(INDEX(LOGEST({33100;47300;69000;102000;150000;220000},{11;12;13;14;15;16},TRUE,FALSE),1),9)', 1.463275628],
  ['ROUND(INDEX(LOGEST({33100;47300;69000;102000;150000;220000},{11;12;13;14;15;16},TRUE,FALSE),1,2),7)', 495.3047702],
  ['ARRAYTOTEXT(LINEST({1;9;5;7},{0;4;2;3},,FALSE),1)', '{2,1}'],
  ['ROUND(INDEX(LINEST(Examples!T1:T11,Examples!P1:S11,TRUE,TRUE),1,1),7)', -234.2371645],
  ['ROUND(INDEX(LINEST(Examples!T1:T11,Examples!P1:S11,TRUE,TRUE),1,4),8)', 27.64138737],
  ['ROUND(INDEX(LINEST(Examples!T1:T11,Examples!P1:S11,TRUE,TRUE),1,5),5)', 52317.83051],
  ['ROUND(INDEX(LINEST(Examples!T1:T11,Examples!P1:S11,TRUE,TRUE),2,1),8)', 13.26801148],
  ['ROUND(INDEX(LINEST(Examples!T1:T11,Examples!P1:S11,TRUE,TRUE),2,5),4)', 12237.3616],
  ['ROUND(INDEX(LINEST(Examples!T1:T11,Examples!P1:S11,TRUE,TRUE),3,1),9)', 0.996747993],
  ['ROUND(INDEX(LINEST(Examples!T1:T11,Examples!P1:S11,TRUE,TRUE),3,2),7)', 970.5784629],
  ['ROUND(INDEX(LINEST(Examples!T1:T11,Examples!P1:S11,TRUE,TRUE),4,1),7)', 459.7536742],
  ['INDEX(LINEST(Examples!T1:T11,Examples!P1:S11,TRUE,TRUE),4,2)', 6],
  ['ROUND(INDEX(LINEST(Examples!T1:T11,Examples!P1:S11,TRUE,TRUE),5,1),0)', 1732393319],
  ['ROUND(INDEX(LINEST(Examples!T1:T11,Examples!P1:S11,TRUE,TRUE),5,2),3)', 5652135.316],
  // Bonds and depreciation
  ['COUPDAYBS(DATE(2011,1,25),DATE(2011,11,15),2,1)', 71],
  ['COUPDAYS(DATE(2011,1,25),DATE(2011,11,15),2,1)', 181],
  ['COUPDAYSNC(DATE(2011,1,25),DATE(2011,11,15),2,1)', 110],
  ['COUPNCD(DATE(2011,1,25),DATE(2011,11,15),2,1)', 40678],
  ['COUPNUM(DATE(2007,1,25),DATE(2008,11,15),2,1)', 4],
  ['COUPPCD(DATE(2011,1,25),DATE(2011,11,15),2,1)', 40497],
  ['ROUND(PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0),5)', 94.63436],
  ['ROUND(YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0),4)', 0.065],
  ['ROUND(DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1),7)', 10.9191453],
  ['ROUND(MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1),5)', 5.73567],
  ['ROUND(ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0),6)', 16.666667],
  ['ROUND(ACCRINT(DATE(2008,3,5),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,FALSE),6)', 15.555556],
  ['ROUND(ACCRINT(DATE(2008,4,5),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0,TRUE),7)', 7.2222222],
  ['ROUND(ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,3),8)', 20.54794521],
  ['ROUND(DISC(DATE(2018,7,1),DATE(2048,1,1),97.975,100,1),6)', 0.000686],
  ['ROUND(INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2),5)', 0.05768],
  ['ROUND(RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2),3)', 1014584.654],
  ['ROUND(PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2),5)', 99.79583],
  ['ROUND(YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2),6)', 0.052823],
  ['ROUND(PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0),8)', 99.98449888],
  ['ROUND(YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0),6)', 0.060954],
  ['ROUND(VDB(2400,300,10*365,0,1),2)', 1.32],
  ['ROUND(VDB(2400,300,10*12,0,1),2)', 40],
  ['ROUND(VDB(2400,300,10,0,1),2)', 480],
  ['ROUND(VDB(2400,300,10*12,6,18),2)', 396.31],
  ['ROUND(VDB(2400,300,10*12,6,18,1.5),2)', 311.81],
  ['ROUND(VDB(2400,300,10,0,0.875,1.5),2)', 315],
];

/**
 * Where this editor's engine (HyperFormula) is known to differ from Excel.
 * Opening a real file keeps Excel's saved value for these and marks them.
 */
const KNOWN_DIFFERENCES = new Set<string>([]);

/** The data of Microsoft's examples for the functions above, cell by cell. */
const EXAMPLES: Record<string, ExcelJS.CellValue> = {
  ...column('A', [4.14, 4.19, 5.17, 5.77, 6.39]),
  ...column('B', ['red', 'orange', 'yellow', 'green', 'blue']),
  C1: "Little Red Riding Hood's red hood",
  ...column('D', [true, 1234.01234, 'Hello']),
  ...column('E', [{ error: '#VALUE!' }, 'Seattle', 1123]),
  F1: 'Smith',
  // Quiz grades and houses for AVERAGEIFS.
  ...table('G', [
    ['Emilio', 75, 85, 87],
    ['Julie', 94, 80, 88],
    ['Hans', 86, 93, 'Incomplete'],
    ['Frederique', 'Incomplete', 75, 75],
  ]),
  ...table('K', [
    ['Cozy Rambler', 230000, 'Issaquah', 3, 'No'],
    ['Snug Bungalow', 197000, 'Bellevue', 2, 'Yes'],
    ['Cool Cape Codder', 345678, 'Bellevue', 4, 'Yes'],
    ['Splendid Split Level', 321900, 'Issaquah', 2, 'Yes'],
    ['Exclusive Tudor', 450000, 'Bellevue', 5, 'Yes'],
    ['Classy Colonial', 395000, 'Bellevue', 4, 'No'],
  ]),
  // Office buildings for LINEST: floor space, offices, entrances, age, assessed value.
  ...table('P', [
    [2310, 2, 2, 20, 142000],
    [2333, 2, 2, 12, 144000],
    [2356, 3, 1.5, 33, 151000],
    [2379, 3, 2, 43, 150000],
    [2402, 2, 3, 53, 139000],
    [2425, 4, 2, 23, 169000],
    [2448, 2, 1.5, 99, 126000],
    [2471, 2, 2, 34, 142900],
    [2494, 3, 3, 23, 163000],
    [2517, 4, 4, 55, 169000],
    [2540, 2, 3, 22, 149000],
  ]),
  // AGGREGATE's data, with errors.
  ...column('U', [{ error: '#DIV/0!' }, 72, 30, { error: '#NUM!' }, 31, 96, 32, 81, 33, 53, 34]),
  ...column('V', [82, 65, 95, 63, 53, 71, 55, 83, 100, 91, 89]),
};

/** Rows of values placed from a column rightwards, starting at row 1. */
function table(first: string, rows: ExcelJS.CellValue[][]): Record<string, ExcelJS.CellValue> {
  const code = first.charCodeAt(0);
  return Object.assign({}, ...rows[0].map((_, c) => column(String.fromCharCode(code + c), rows.map((r) => r[c]))));
}

function column(col: string, values: ExcelJS.CellValue[], from = 1): Record<string, ExcelJS.CellValue> {
  return Object.fromEntries(values.map((v, i) => [`${col}${from + i}`, v]));
}

async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  for (let i = 1; i <= 10; i++) data.getCell(`A${i}`).value = i;
  ['apple', 'Banana', 'cherry', 'apple', 'date'].forEach((v, i) => (data.getCell(`B${i + 1}`).value = v));
  ['a', 'b', 'c', 'd'].forEach((v, i) => {
    data.getCell(`E${i + 1}`).value = v;
    data.getCell(`F${i + 1}`).value = i + 1;
  });
  const examples = wb.addWorksheet('Examples');
  for (const [ref, value] of Object.entries(EXAMPLES)) examples.getCell(ref).value = value;
  const calc = wb.addWorksheet('Calc');
  CASES.forEach(([formula, result], i) => {
    calc.getCell(`A${i + 1}`).value = formula;
    calc.getCell(`B${i + 1}`).value = { formula, result } as ExcelJS.CellFormulaValue;
  });
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let report: CalcReport;
beforeAll(async () => {
  const s = WorkbookSession.open(await workbook(), 'accuracy.xlsx');
  report = s.ensureEngine();
});

describe('calculation accuracy', () => {
  it('matches Excel on every formula in the catalogue, apart from the known differences', () => {
    const differ = report.samples.map((x) => `${x.formula} → Excel ${x.excel ?? ''}, here ${x.engine ?? x.reason}`);
    const unexpected = report.samples.filter((x) => !KNOWN_DIFFERENCES.has(x.formula));
    expect(unexpected.map((x) => `${x.formula} → Excel ${x.excel ?? ''}, here ${x.engine ?? x.reason}`)).toEqual([]);
    expect(report.formulaCells).toBe(CASES.length);
    expect(report.liveCells).toBe(CASES.length - differ.length);
  });
});
