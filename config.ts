const CANVAS_BASE='https://bostoncollege.instructure.com/api/v1';
const PAGINATION_PER_PAGE = 10;
const STUDENT_START_ROW = 4;
const REPORT_START = 'D4';
const SUMMARY_START = 'K4';
const STUDENT_NAME_CELL = 'E1';
const OMIT_STRING = 'Omit';
const DROP_STRING = 'drop';
const REPORT_HEADINGS = ['Group', 'Assignment', 'Possible', 'Actual', 'Drop', 'Use'];
const SUMMARY_HEADINGS = ['Group', 'Weight', 'Average', 'Weighted Avg.'];
const ASSIGNMENT_AVERAGE_FORMULA = `=IFERROR(ArrayFormula(SUM(indirect("I"&MATCH(${SUMMARY_START}, $D:$D, FALSE)+1&":I"&MIN(IF((INDIRECT("I"&MATCH(${SUMMARY_START}, $D:$D, FALSE)+1&":I"))="",ROW(INDIRECT("I"&MATCH(${SUMMARY_START}, $D:$D, FALSE)+1&":I"))-1)))))/ArrayFormula(SUMIF(indirect("I"&MATCH(${SUMMARY_START}, $D:$D, FALSE)+1&":I"&MIN(IF((INDIRECT("I"&MATCH(${SUMMARY_START}, $D:$D, FALSE)+1&":I"))="",ROW(INDIRECT("I"&MATCH(${SUMMARY_START}, $D:$D, FALSE)+1&":I"))-1))), "<>${OMIT_STRING}", indirect("F"&MATCH(${SUMMARY_START}, $D:$D, FALSE)+1&":F"&MIN(IF((INDIRECT("F"&MATCH(${SUMMARY_START}, $D:$D, FALSE)+1&":F"))="",ROW(INDIRECT("F"&MATCH(${SUMMARY_START}, $D:$D, FALSE)+1&":F"))-1))))), #N/A)`;
const WEIGHTED_AVERAGE_FORMULA = '=IF(ISNA(M4), #N/A, L4*M4)';
const VALUE_TO_USE_FORMULA = '=IF(H_row_<>"' + DROP_STRING + '", G_row_, "' + OMIT_STRING + '")';

export {
    CANVAS_BASE,
    PAGINATION_PER_PAGE,
    STUDENT_START_ROW,
    REPORT_START,
    SUMMARY_START,
    STUDENT_NAME_CELL,
    REPORT_HEADINGS,
    SUMMARY_HEADINGS,
    ASSIGNMENT_AVERAGE_FORMULA,
    WEIGHTED_AVERAGE_FORMULA,
    DROP_STRING,
    VALUE_TO_USE_FORMULA
}