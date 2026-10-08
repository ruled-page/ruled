/**
 * Excel's worksheet functions as the formula editor offers them while a
 * formula is typed: every name, whether this editor calculates it, and for
 * the common ones their arguments and what they do. The page has no
 * calculation engine (it runs in the worker), so this is a plain list;
 * tests/function-catalog.test.ts checks it against what the engine registers.
 * Names, arguments and descriptions are Excel's English ones, as in Excel.
 */

/** Excel functions the engine calculates. */
const CALCULATED =
  'ABS ACCRINT ACCRINTM ACOS ACOSH ACOT ACOTH ADDRESS AGGREGATE AND ARABIC ARRAYTOTEXT ASIN ASINH ATAN ATAN2 ' +
  'ATANH AVEDEV AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS BASE BESSELI BESSELJ BESSELK BESSELY BETA.DIST BETA.INV ' +
  'BETADIST BETAINV BIN2DEC BIN2HEX BIN2OCT BINOM.DIST BINOM.DIST.RANGE BINOM.INV BINOMDIST BITAND BITLSHIFT ' +
  'BITOR BITRSHIFT BITXOR CEILING CEILING.MATH CEILING.PRECISE CHAR CHIDIST CHIINV CHISQ.DIST CHISQ.DIST.RT ' +
  'CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHITEST CHOOSE CHOOSECOLS CHOOSEROWS CLEAN CODE COLUMN COLUMNS COMBIN ' +
  'COMBINA COMPLEX CONCAT CONCATENATE CONFIDENCE CONFIDENCE.NORM CONFIDENCE.T CORREL COS COSH COT COTH COUNT ' +
  'COUNTA COUNTBLANK COUNTIF COUNTIFS COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD COVAR COVARIANCE.P ' +
  'COVARIANCE.S CRITBINOM CSC CSCH CUMIPMT CUMPRINC DATE DATEDIF DATEVALUE DAVERAGE DAY DAYS DAYS360 DB DCOUNT ' +
  'DCOUNTA DDB DEC2BIN DEC2HEX DEC2OCT DECIMAL DEGREES DELTA DEVSQ DGET DISC DMAX DMIN DOLLARDE DOLLARFR ' +
  'DPRODUCT DROP DSTDEV DSTDEVP DSUM DURATION DVAR DVARP EDATE EFFECT EOMONTH ERF ERF.PRECISE ERFC ERFC.PRECISE ' +
  'ERROR.TYPE EVEN EXACT EXP EXPAND EXPON.DIST EXPONDIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FACT FACTDOUBLE ' +
  'FALSE FDIST FILTER FIND FINDB FINV FISHER FISHERINV FIXED FLOOR FLOOR.MATH FLOOR.PRECISE FORECAST ' +
  'FORECAST.LINEAR FORMULATEXT FREQUENCY FTEST FV FVSCHEDULE GAMMA GAMMA.DIST GAMMA.INV GAMMADIST GAMMAINV ' +
  'GAMMALN GAMMALN.PRECISE GAUSS GCD GEOMEAN GESTEP GROWTH HARMEAN HEX2BIN HEX2DEC HEX2OCT HLOOKUP HOUR HSTACK ' +
  'HYPERLINK HYPGEOM.DIST HYPGEOMDIST IF IFERROR IFNA IFS IMABS IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMCOSH ' +
  'IMCOT IMCSC IMCSCH IMDIV IMEXP IMLN IMLOG10 IMLOG2 IMPOWER IMPRODUCT IMREAL IMSEC IMSECH IMSIN IMSINH IMSQRT ' +
  'IMSUB IMSUM IMTAN INDEX INT INTERCEPT INTRATE IPMT IRR ISBLANK ISERR ISERROR ISEVEN ISFORMULA ISLOGICAL ISNA ' +
  'ISNONTEXT ISNUMBER ISO.CEILING ISODD ISOWEEKNUM ISPMT ISREF ISTEXT KURT LARGE LCM LEFT LEFTB LEN LENB LINEST ' +
  'LN LOG LOG10 LOGEST LOGINV LOGNORM.DIST LOGNORM.INV LOGNORMDIST LOOKUP LOWER MATCH MAX MAXA MAXIFS MDETERM ' +
  'MDURATION MEDIAN MID MIDB MIN MINA MINIFS MINUTE MINVERSE MIRR MMULT MOD MODE MODE.MULT MODE.SNGL MONTH ' +
  'MROUND MULTINOMIAL MUNIT N NA NEGBINOM.DIST NEGBINOMDIST NETWORKDAYS NETWORKDAYS.INTL NOMINAL NORM.DIST ' +
  'NORM.INV NORM.S.DIST NORM.S.INV NORMDIST NORMINV NORMSDIST NORMSINV NOT NOW NPER NPV NUMBERVALUE OCT2BIN ' +
  'OCT2DEC OCT2HEX ODD OFFSET OR PDURATION PEARSON PERCENTILE PERCENTILE.EXC PERCENTILE.INC PERCENTOF ' +
  'PERCENTRANK PERCENTRANK.EXC PERCENTRANK.INC PERMUT PERMUTATIONA PHI PI PMT POISSON POISSON.DIST POWER PPMT ' +
  'PRICE PRICEDISC PRICEMAT PROB PRODUCT PROPER PV QUARTILE QUARTILE.EXC QUARTILE.INC QUOTIENT RADIANS RAND ' +
  'RANDARRAY RANDBETWEEN RANK RANK.AVG RANK.EQ RATE RECEIVED REPLACE REPLACEB REPT RIGHT RIGHTB ROMAN ROUND ' +
  'ROUNDDOWN ROUNDUP ROW ROWS RRI RSQ SEARCH SEARCHB SEC SECH SECOND SEQUENCE SERIESSUM SHEET SHEETS SIGN SIN ' +
  'SINH SKEW SKEW.P SLN SLOPE SMALL SORT SORTBY SQRT SQRTPI STANDARDIZE STDEV STDEV.P STDEV.S STDEVA STDEVP ' +
  'STDEVPA STEYX SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SUMSQ SUMX2MY2 SUMX2PY2 SUMXMY2 SWITCH SYD T ' +
  'T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TAN TANH TBILLEQ TBILLPRICE TBILLYIELD TDIST TEXT ' +
  'TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TIME TIMEVALUE TINV TOCOL TODAY TOROW TRANSPOSE TREND TRIM TRIMMEAN ' +
  'TRUE TRUNC TTEST TYPE UNICHAR UNICODE UNIQUE UPPER VALUE VALUETOTEXT VAR VAR.P VAR.S VARA VARP VARPA VDB ' +
  'VLOOKUP VSTACK WEEKDAY WEEKNUM WEIBULL WEIBULL.DIST WORKDAY WORKDAY.INTL WRAPCOLS WRAPROWS XIRR XLOOKUP ' +
  'XMATCH XNPV XOR YEAR YEARFRAC YIELD YIELDDISC YIELDMAT Z.TEST ZTEST';

/** Excel functions the engine doesn't calculate: formulas using them keep Excel's saved value, and Excel calculates them on opening. */
const NOT_CALCULATED =
  'AMORDEGRC AMORLINC AREAS ASC BAHTTEXT BYCOL BYROW CELL CONVERT CUBEKPIMEMBER CUBEMEMBER CUBEMEMBERPROPERTY ' +
  'CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE DBCS DETECTLANGUAGE DOLLAR ENCODEURL FIELDVALUE FILTERXML ' +
  'FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT GETPIVOTDATA GROUPBY IMAGE ' +
  'INDIRECT INFO ISOMITTED LAMBDA LET MAKEARRAY MAP ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD PHONETIC PIVOTBY ' +
  'REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST RTD SCAN STOCKHISTORY TRANSLATE TRIMRANGE WEBSERVICE';

/**
 * Arguments and descriptions of the common functions, one per line:
 * NAME(arguments) What it does. Optional arguments are in brackets, and a
 * trailing "..." repeats the last group, as in Excel's own hints.
 */
const HELP = `
ABS(number) Returns the absolute value of a number.
ACOS(number) Returns the arccosine of a number, in radians.
ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text]) Returns a cell reference as text, given row and column numbers.
AGGREGATE(function_num, options, ref1, [ref2], ...) Returns an aggregate (sum, average, count…), optionally ignoring hidden rows and errors.
AND(logical1, [logical2], ...) Returns TRUE if all of its arguments are TRUE.
AREAS(reference) Returns the number of areas in a reference.
ARRAYTOTEXT(array, [format]) Returns an array of values as text.
ASIN(number) Returns the arcsine of a number, in radians.
ATAN(number) Returns the arctangent of a number, in radians.
ATAN2(x_num, y_num) Returns the arctangent of x and y coordinates, in radians.
AVEDEV(number1, [number2], ...) Returns the average of the absolute deviations of numbers from their mean.
AVERAGE(number1, [number2], ...) Returns the average (arithmetic mean) of its arguments.
AVERAGEA(value1, [value2], ...) Returns the average of its arguments, counting text as 0 and TRUE as 1.
AVERAGEIF(range, criteria, [average_range]) Returns the average of the cells that meet a condition.
AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...) Returns the average of the cells that meet several conditions.
BIN2DEC(number) Converts a binary number to decimal.
BINOM.DIST(number_s, trials, probability_s, cumulative) Returns the binomial distribution probability.
BYCOL(array, lambda) Applies a LAMBDA to each column and returns an array of the results.
BYROW(array, lambda) Applies a LAMBDA to each row and returns an array of the results.
CEILING(number, significance) Rounds a number up to the nearest multiple of significance.
CEILING.MATH(number, [significance], [mode]) Rounds a number up to the nearest integer or multiple of significance.
CELL(info_type, [reference]) Returns information about the formatting, location or contents of a cell.
CHAR(number) Returns the character with the given code.
CHISQ.TEST(actual_range, expected_range) Returns the chi-squared test for independence.
CHOOSE(index_num, value1, [value2], ...) Chooses a value from a list by its position.
CHOOSECOLS(array, col_num1, [col_num2], ...) Returns the given columns of an array.
CHOOSEROWS(array, row_num1, [row_num2], ...) Returns the given rows of an array.
CLEAN(text) Removes all nonprintable characters from text.
CODE(text) Returns the code of the first character in text.
COLUMN([reference]) Returns the column number of a reference.
COLUMNS(array) Returns the number of columns in an array or reference.
COMBIN(number, number_chosen) Returns the number of combinations of a number of items.
CONCAT(text1, [text2], ...) Joins text from several ranges or strings.
CONCATENATE(text1, [text2], ...) Joins several text items into one.
CONFIDENCE.NORM(alpha, standard_dev, size) Returns the confidence interval for a population mean, using a normal distribution.
CONFIDENCE.T(alpha, standard_dev, size) Returns the confidence interval for a population mean, using Student's t-distribution.
CONVERT(number, from_unit, to_unit) Converts a number from one measurement system to another.
CORREL(array1, array2) Returns the correlation coefficient between two data sets.
COS(number) Returns the cosine of an angle.
COUNT(value1, [value2], ...) Counts the cells that contain numbers.
COUNTA(value1, [value2], ...) Counts the cells that are not empty.
COUNTBLANK(range) Counts the empty cells in a range.
COUNTIF(range, criteria) Counts the cells in a range that meet a condition.
COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...) Counts the cells that meet several conditions.
COVARIANCE.P(array1, array2) Returns the population covariance of two data sets.
COVARIANCE.S(array1, array2) Returns the sample covariance of two data sets.
CUMIPMT(rate, nper, pv, start_period, end_period, type) Returns the cumulative interest paid on a loan between two periods.
CUMPRINC(rate, nper, pv, start_period, end_period, type) Returns the cumulative principal paid on a loan between two periods.
DATE(year, month, day) Returns the serial number of a date.
DATEDIF(start_date, end_date, unit) Returns the number of days, months or years between two dates.
DATEVALUE(date_text) Converts a date written as text to a serial number.
DAVERAGE(database, field, criteria) Returns the average of database entries that meet conditions.
DAY(serial_number) Returns the day of the month, from 1 to 31.
DAYS(end_date, start_date) Returns the number of days between two dates.
DAYS360(start_date, end_date, [method]) Returns the number of days between two dates on a 360-day year.
DB(cost, salvage, life, period, [month]) Returns depreciation for a period by the fixed-declining balance method.
DCOUNT(database, field, criteria) Counts the database entries with numbers that meet conditions.
DCOUNTA(database, field, criteria) Counts the non-empty database entries that meet conditions.
DDB(cost, salvage, life, period, [factor]) Returns depreciation for a period by the double-declining balance method.
DEC2BIN(number, [places]) Converts a decimal number to binary.
DEC2HEX(number, [places]) Converts a decimal number to hexadecimal.
DEGREES(angle) Converts radians to degrees.
DEVSQ(number1, [number2], ...) Returns the sum of squares of deviations from the mean.
DGET(database, field, criteria) Returns the single database entry that meets conditions.
DMAX(database, field, criteria) Returns the largest database entry that meets conditions.
DMIN(database, field, criteria) Returns the smallest database entry that meets conditions.
DOLLAR(number, [decimals]) Converts a number to text in currency format.
DROP(array, rows, [columns]) Removes rows or columns from the start or end of an array.
DSUM(database, field, criteria) Adds the database entries that meet conditions.
EDATE(start_date, months) Returns the date a number of months before or after a date.
EFFECT(nominal_rate, npery) Returns the effective annual interest rate.
EOMONTH(start_date, months) Returns the last day of the month a number of months before or after a date.
ERROR.TYPE(error_val) Returns a number for an error value.
EVEN(number) Rounds a number up to the nearest even integer.
EXACT(text1, text2) Checks whether two text values are exactly the same.
EXP(number) Returns e raised to the power of a number.
EXPAND(array, rows, [columns], [pad_with]) Expands an array to the given size.
EXPON.DIST(x, lambda, cumulative) Returns the exponential distribution.
FACT(number) Returns the factorial of a number.
FALSE() Returns the logical value FALSE.
FILTER(array, include, [if_empty]) Filters a range or array by a condition.
FIND(find_text, within_text, [start_num]) Finds one text value within another (case-sensitive).
FIXED(number, [decimals], [no_commas]) Formats a number as text with a fixed number of decimals.
FLOOR(number, significance) Rounds a number down to the nearest multiple of significance.
FLOOR.MATH(number, [significance], [mode]) Rounds a number down to the nearest integer or multiple of significance.
FORECAST(x, known_y's, known_x's) Predicts a value along a linear trend.
FORECAST.LINEAR(x, known_y's, known_x's) Predicts a value along a linear trend.
FORMULATEXT(reference) Returns the formula at a reference as text.
FREQUENCY(data_array, bins_array) Returns how often values occur within ranges, as a vertical array.
FV(rate, nper, pmt, [pv], [type]) Returns the future value of an investment.
GEOMEAN(number1, [number2], ...) Returns the geometric mean.
GETPIVOTDATA(data_field, pivot_table, [field1, item1], ...) Returns data stored in a PivotTable report.
GROWTH(known_y's, [known_x's], [new_x's], [const]) Returns values along an exponential trend.
HARMEAN(number1, [number2], ...) Returns the harmonic mean.
HEX2DEC(number) Converts a hexadecimal number to decimal.
HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup]) Looks in the top row of a table and returns the value from a row you choose.
HOUR(serial_number) Returns the hour of a time, from 0 to 23.
HSTACK(array1, [array2], ...) Joins arrays side by side into one array.
HYPERLINK(link_location, [friendly_name]) Creates a link to a document or a place in the workbook.
IF(logical_test, [value_if_true], [value_if_false]) Returns one value if a condition is TRUE and another if it is FALSE.
IFERROR(value, value_if_error) Returns a value of your choice if a formula gives an error, and its result otherwise.
IFNA(value, value_if_na) Returns a value of your choice if a formula gives #N/A, and its result otherwise.
IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...) Returns the value for the first condition that is TRUE.
IMAGE(source, [alt_text], [sizing], [height], [width]) Shows an image from a web address in the cell.
INDEX(array, row_num, [column_num]) Returns the value at a row and column of a range or array.
INDIRECT(ref_text, [a1]) Returns the reference given by text.
INFO(type_text) Returns information about the current operating environment.
INT(number) Rounds a number down to the nearest integer.
INTERCEPT(known_y's, known_x's) Returns where a linear regression line crosses the y-axis.
IPMT(rate, per, nper, pv, [fv], [type]) Returns the interest paid on an investment in a period.
IRR(values, [guess]) Returns the internal rate of return for a series of cash flows.
ISBLANK(value) Returns TRUE if the value is empty.
ISERR(value) Returns TRUE if the value is any error except #N/A.
ISERROR(value) Returns TRUE if the value is any error.
ISEVEN(number) Returns TRUE if the number is even.
ISFORMULA(reference) Returns TRUE if the cell holds a formula.
ISLOGICAL(value) Returns TRUE if the value is TRUE or FALSE.
ISNA(value) Returns TRUE if the value is #N/A.
ISNONTEXT(value) Returns TRUE if the value is not text.
ISNUMBER(value) Returns TRUE if the value is a number.
ISODD(number) Returns TRUE if the number is odd.
ISOWEEKNUM(date) Returns the ISO week number of the year for a date.
ISREF(value) Returns TRUE if the value is a reference.
ISTEXT(value) Returns TRUE if the value is text.
KURT(number1, [number2], ...) Returns the kurtosis of a data set.
LAMBDA([parameter_or_calculation], ...) Creates a function from parameters and a calculation.
LARGE(array, k) Returns the k-th largest value in a data set.
LCM(number1, [number2], ...) Returns the least common multiple.
LEFT(text, [num_chars]) Returns the first characters of text.
LEN(text) Returns the number of characters in text.
LET(name1, name_value1, calculation_or_name2, [name_value2, calculation_or_name3], ...) Names the results of calculations for use within a formula.
LINEST(known_y's, [known_x's], [const], [stats]) Returns the statistics of a linear trend.
LN(number) Returns the natural logarithm of a number.
LOG(number, [base]) Returns the logarithm of a number to a base.
LOG10(number) Returns the base-10 logarithm of a number.
LOOKUP(lookup_value, lookup_vector, [result_vector]) Looks up a value in a row or column and returns the value in the same place of another.
LOWER(text) Converts text to lowercase.
MAKEARRAY(rows, cols, lambda) Returns an array of a given size, calculated by a LAMBDA.
MAP(array1, lambda_or_array2, ...) Applies a LAMBDA to each value of arrays and returns the results.
MATCH(lookup_value, lookup_array, [match_type]) Returns the position of a value in a range.
MAX(number1, [number2], ...) Returns the largest value.
MAXA(value1, [value2], ...) Returns the largest value, counting text as 0 and TRUE as 1.
MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...) Returns the largest value among cells that meet conditions.
MDETERM(array) Returns the matrix determinant of an array.
MEDIAN(number1, [number2], ...) Returns the median of the numbers.
MID(text, start_num, num_chars) Returns characters from the middle of text.
MIN(number1, [number2], ...) Returns the smallest value.
MINA(value1, [value2], ...) Returns the smallest value, counting text as 0 and TRUE as 1.
MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...) Returns the smallest value among cells that meet conditions.
MINUTE(serial_number) Returns the minute of a time, from 0 to 59.
MINVERSE(array) Returns the inverse matrix of an array.
MIRR(values, finance_rate, reinvest_rate) Returns the modified internal rate of return for a series of cash flows.
MMULT(array1, array2) Returns the matrix product of two arrays.
MOD(number, divisor) Returns the remainder after division.
MODE(number1, [number2], ...) Returns the most common value in a data set.
MODE.SNGL(number1, [number2], ...) Returns the most common value in a data set.
MONTH(serial_number) Returns the month of a date, from 1 to 12.
MROUND(number, multiple) Rounds a number to the nearest multiple.
N(value) Converts a value to a number.
NA() Returns the error value #N/A.
NETWORKDAYS(start_date, end_date, [holidays]) Returns the number of working days between two dates.
NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays]) Returns the number of working days between two dates, with weekends of your choice.
NOMINAL(effect_rate, npery) Returns the nominal annual interest rate.
NORM.DIST(x, mean, standard_dev, cumulative) Returns the normal distribution.
NORM.INV(probability, mean, standard_dev) Returns the inverse of the normal cumulative distribution.
NORM.S.DIST(z, cumulative) Returns the standard normal distribution.
NORM.S.INV(probability) Returns the inverse of the standard normal cumulative distribution.
NOT(logical) Reverses a logical value.
NOW() Returns the current date and time.
NPER(rate, pmt, pv, [fv], [type]) Returns the number of periods for an investment.
NPV(rate, value1, [value2], ...) Returns the net present value of an investment from periodic cash flows and a discount rate.
NUMBERVALUE(text, [decimal_separator], [group_separator]) Converts text to a number, whatever the locale.
ODD(number) Rounds a number up to the nearest odd integer.
OFFSET(reference, rows, cols, [height], [width]) Returns a reference a number of rows and columns from a reference.
OR(logical1, [logical2], ...) Returns TRUE if any argument is TRUE.
PEARSON(array1, array2) Returns the Pearson correlation coefficient.
PERCENTILE(array, k) Returns the k-th percentile of values in a range.
PERCENTILE.EXC(array, k) Returns the k-th percentile, where k is between 0 and 1 exclusive.
PERCENTILE.INC(array, k) Returns the k-th percentile, where k is between 0 and 1 inclusive.
PERCENTOF(data_subset, data_all) Returns the share a subset makes of a total.
PERCENTRANK(array, x, [significance]) Returns the rank of a value in a data set as a percentage.
PERCENTRANK.INC(array, x, [significance]) Returns the rank of a value in a data set as a percentage (0 to 1 inclusive).
PERMUT(number, number_chosen) Returns the number of permutations of a number of items.
PI() Returns the value of pi.
PMT(rate, nper, pv, [fv], [type]) Returns the periodic payment for a loan.
POISSON.DIST(x, mean, cumulative) Returns the Poisson distribution.
POWER(number, power) Returns a number raised to a power.
PPMT(rate, per, nper, pv, [fv], [type]) Returns the payment on the principal of an investment in a period.
PRODUCT(number1, [number2], ...) Multiplies its arguments.
PROPER(text) Capitalises the first letter of each word.
PV(rate, nper, pmt, [fv], [type]) Returns the present value of an investment.
QUARTILE(array, quart) Returns the quartile of a data set.
QUARTILE.EXC(array, quart) Returns the quartile of a data set, from percentiles 0 to 1 exclusive.
QUARTILE.INC(array, quart) Returns the quartile of a data set, from percentiles 0 to 1 inclusive.
QUOTIENT(numerator, denominator) Returns the integer part of a division.
RADIANS(angle) Converts degrees to radians.
RAND() Returns a random number between 0 and 1.
RANDARRAY([rows], [columns], [min], [max], [whole_number]) Returns an array of random numbers.
RANDBETWEEN(bottom, top) Returns a random integer between two numbers.
RANK(number, ref, [order]) Returns the rank of a number in a list.
RANK.AVG(number, ref, [order]) Returns the rank of a number in a list; ties get their average rank.
RANK.EQ(number, ref, [order]) Returns the rank of a number in a list; ties get the top rank.
RATE(nper, pmt, pv, [fv], [type], [guess]) Returns the interest rate per period of an annuity.
REDUCE([initial_value], array, lambda) Reduces an array to one value by applying a LAMBDA to each value.
REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity]) Extracts text that matches a regular expression.
REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity]) Replaces text that matches a regular expression.
REGEXTEST(text, pattern, [case_sensitivity]) Checks whether text matches a regular expression.
REPLACE(old_text, start_num, num_chars, new_text) Replaces characters within text.
REPT(text, number_times) Repeats text a number of times.
RIGHT(text, [num_chars]) Returns the last characters of text.
ROUND(number, num_digits) Rounds a number to a number of digits.
ROUNDDOWN(number, num_digits) Rounds a number down, towards zero.
ROUNDUP(number, num_digits) Rounds a number up, away from zero.
ROW([reference]) Returns the row number of a reference.
ROWS(array) Returns the number of rows in an array or reference.
RSQ(known_y's, known_x's) Returns the square of the Pearson correlation coefficient.
SCAN([initial_value], array, lambda) Applies a LAMBDA to each value and returns each intermediate result.
SEARCH(find_text, within_text, [start_num]) Finds one text value within another (not case-sensitive).
SECOND(serial_number) Returns the second of a time, from 0 to 59.
SEQUENCE(rows, [columns], [start], [step]) Returns a list of sequential numbers.
SHEET([value]) Returns the sheet number of a sheet.
SHEETS([reference]) Returns the number of sheets in a reference.
SIGN(number) Returns the sign of a number: 1, 0 or -1.
SIN(number) Returns the sine of an angle.
SKEW(number1, [number2], ...) Returns the skewness of a distribution.
SLN(cost, salvage, life) Returns straight-line depreciation for one period.
SLOPE(known_y's, known_x's) Returns the slope of a linear regression line.
SMALL(array, k) Returns the k-th smallest value in a data set.
SORT(array, [sort_index], [sort_order], [by_col]) Sorts the contents of a range or array.
SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2], ...) Sorts a range or array by the values in another.
SQRT(number) Returns a positive square root.
STANDARDIZE(x, mean, standard_dev) Returns a normalised value.
STDEV(number1, [number2], ...) Estimates standard deviation from a sample.
STDEV.P(number1, [number2], ...) Returns the standard deviation of a whole population.
STDEV.S(number1, [number2], ...) Estimates standard deviation from a sample.
STEYX(known_y's, known_x's) Returns the standard error of the predicted y-values of a regression.
SUBSTITUTE(text, old_text, new_text, [instance_num]) Replaces old text with new text.
SUBTOTAL(function_num, ref1, [ref2], ...) Returns a subtotal in a list, ignoring other subtotals.
SUM(number1, [number2], ...) Adds its arguments.
SUMIF(range, criteria, [sum_range]) Adds the cells that meet a condition.
SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...) Adds the cells that meet several conditions.
SUMPRODUCT(array1, [array2], ...) Returns the sum of the products of corresponding values.
SUMSQ(number1, [number2], ...) Returns the sum of the squares of its arguments.
SWITCH(expression, value1, result1, [default_or_value2, result2], ...) Compares a value against a list and returns the result for the first match.
SYD(cost, salvage, life, per) Returns sum-of-years' digits depreciation for a period.
T(value) Returns the value if it is text, and empty text otherwise.
T.DIST(x, deg_freedom, cumulative) Returns Student's left-tailed t-distribution.
T.INV(probability, deg_freedom) Returns the left-tailed inverse of Student's t-distribution.
T.INV.2T(probability, deg_freedom) Returns the two-tailed inverse of Student's t-distribution.
T.TEST(array1, array2, tails, type) Returns the probability of Student's t-test.
TAKE(array, rows, [columns]) Returns rows or columns from the start or end of an array.
TAN(number) Returns the tangent of an angle.
TEXT(value, format_text) Formats a number and converts it to text.
TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]) Returns the text after a delimiter.
TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]) Returns the text before a delimiter.
TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...) Joins text from several ranges or strings, with a delimiter between.
TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]) Splits text into rows and columns at delimiters.
TIME(hour, minute, second) Returns the serial number of a time.
TIMEVALUE(time_text) Converts a time written as text to a serial number.
TOCOL(array, [ignore], [scan_by_column]) Returns an array as one column.
TODAY() Returns today's date.
TOROW(array, [ignore], [scan_by_column]) Returns an array as one row.
TRANSPOSE(array) Swaps the rows and columns of an array.
TREND(known_y's, [known_x's], [new_x's], [const]) Returns values along a linear trend.
TRIM(text) Removes spaces from text, except single spaces between words.
TRIMMEAN(array, percent) Returns the mean of the interior of a data set.
TRIMRANGE(range, [trim_rows], [trim_cols]) Leaves out empty rows and columns at the edges of a range.
TRUE() Returns the logical value TRUE.
TRUNC(number, [num_digits]) Truncates a number to an integer or a number of digits.
TYPE(value) Returns a number for the type of a value.
UNICHAR(number) Returns the Unicode character with the given code.
UNICODE(text) Returns the Unicode code of the first character of text.
UNIQUE(array, [by_col], [exactly_once]) Returns the unique values of a range or array.
UPPER(text) Converts text to uppercase.
VALUE(text) Converts text to a number.
VALUETOTEXT(value, [format]) Returns a value as text.
VAR(number1, [number2], ...) Estimates variance from a sample.
VAR.P(number1, [number2], ...) Returns the variance of a whole population.
VAR.S(number1, [number2], ...) Estimates variance from a sample.
VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup]) Looks in the first column of a table and returns the value from a column you choose.
VSTACK(array1, [array2], ...) Stacks arrays one above another into one array.
WEEKDAY(serial_number, [return_type]) Returns the day of the week of a date, as a number.
WEEKNUM(serial_number, [return_type]) Returns the week number of the year for a date.
WORKDAY(start_date, days, [holidays]) Returns the date a number of working days before or after a date.
WORKDAY.INTL(start_date, days, [weekend], [holidays]) Returns the date a number of working days away, with weekends of your choice.
WRAPCOLS(vector, wrap_count, [pad_with]) Wraps a row or column of values into columns.
WRAPROWS(vector, wrap_count, [pad_with]) Wraps a row or column of values into rows.
XIRR(values, dates, [guess]) Returns the internal rate of return for cash flows that are not periodic.
XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]) Searches a range for a match and returns the matching item from another range.
XMATCH(lookup_value, lookup_array, [match_mode], [search_mode]) Returns the position of an item in an array.
XNPV(rate, values, dates) Returns the net present value of cash flows that are not periodic.
XOR(logical1, [logical2], ...) Returns TRUE if an odd number of its arguments are TRUE.
YEAR(serial_number) Returns the year of a date.
YEARFRAC(start_date, end_date, [basis]) Returns the fraction of a year between two dates.
`;

export interface FunctionHelp {
  name: string;
  /** Whether this editor calculates it; formulas using others keep Excel's value until Excel recalculates them. */
  calculated: boolean;
  /** Its arguments as Excel lists them, e.g. "lookup_value, table_array, col_index_num, [range_lookup]". */
  args?: string;
  /** What it does, in a sentence. */
  description?: string;
}

let catalog: Map<string, FunctionHelp> | undefined;

function load(): Map<string, FunctionHelp> {
  if (catalog) return catalog;
  const all = [...CALCULATED.split(' ').map((name) => ({ name, calculated: true })), ...NOT_CALCULATED.split(' ').map((name) => ({ name, calculated: false }))];
  all.sort((a, b) => (a.name < b.name ? -1 : 1));
  catalog = new Map<string, FunctionHelp>(all.map((f) => [f.name, f]));
  for (const line of HELP.split('\n')) {
    const m = /^([A-Z0-9.]+)\((.*?)\) (.+)$/.exec(line);
    const f = m && catalog.get(m[1]);
    if (f) Object.assign(f, { args: m[2], description: m[3] });
  }
  return catalog;
}

/** Every function's name, in alphabetical order. */
export function knownFunctionNames(): string[] {
  return [...load().keys()];
}

/** A function by its name (in any case), or undefined if Excel has no such function. */
export function functionHelp(name: string): FunctionHelp | undefined {
  return load().get(name.toUpperCase());
}

/** The functions whose names start with `prefix` (in any case), in alphabetical order. */
export function functionsStartingWith(prefix: string): FunctionHelp[] {
  const p = prefix.toUpperCase();
  return p ? [...load().values()].filter((f) => f.name.startsWith(p)) : [];
}
