import {
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
} from './config';

// Borrowed from https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Template_literals
const template = (strings, ...keys) => {
  return (function(...values) {
    let dict = values[values.length - 1] || {};
    let result = [strings[0]];
    keys.forEach(function(key, i) {
      let value = Number.isInteger(key) ? values[key] : dict[key];
      result.push(value, strings[i + 1]);
    });
    return result.join('');
  });
};

const URL_ASSIGNMENT_GROUPS = template`${'CANVAS_BASE'}/courses/${'courseId'}/assignment_groups`;
const URL_ASSIGNMENTS = template`${'CANVAS_BASE'}/courses/${'courseId'}/assignment_groups/${'assignment_group'}/assignments`;
const URL_STUDENTS = template`${'CANVAS_BASE'}/courses/${'courseId'}/users`;
const URL_SUBMISSIONS = template`${'CANVAS_BASE'}/courses/${'courseId'}/assignments/${'assignmentId'}/submissions/${'studentId'}`;

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Canvas')
      .addItem('Fetch Course', 'fetchCourse')
      .addItem('Authorize', 'authorize')
      .addToUi();
}

const message = (mesg) => {
  SpreadsheetApp.getActiveSpreadsheet().toast(mesg);
};

const getLinkNext = (header):string => {
  const linkrels = header.split(',');

  for (const linkrel of linkrels) {
    const [link, rel] = linkrel.split(';');
    if (rel.indexOf('next') != -1) {
      return link.slice(1, -1);  // Lose the '<' and '>' characters
    }
  }

  // If we get here, there's no rel="next"
  return null;
}

const authorize = ():object => {
  const cache = CacheService.getUserCache();
  let token:string = cache.get('AUTH_TOKEN');

  if (token === null) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Paste in your Canvas auth token');
    if (response.getSelectedButton() === ui.Button.OK) {
      token = response.getResponseText();
      cache.put('AUTH_TOKEN', token);
    }
    else {
      return null;
    }
  }

  const authHeader = {Authorization: `Bearer ${token}`};
  cache.put('authHeader', JSON.stringify(authHeader));

  return authHeader;
};

const fetch_assignment_groups = (courseId, auth_header) => {
  const url = URL_ASSIGNMENT_GROUPS({CANVAS_BASE: CANVAS_BASE, courseId: courseId});
  let assignment_groups = {};

  const resp = JSON.parse(UrlFetchApp.fetch(url, {headers: auth_header}).getContentText());
  for (let ag of resp) {
    assignment_groups[ag.id] = ag;
  }

  return assignment_groups;
};

const fetch_assignment_data = (assignment_groups, courseId, auth_header) => {
  // By assignment group, fetch the assignments in the group. Keep track of possible points,
  // and only retain the assignments that count toward the final grade.
  //
  // This function modifies its argument; it does not return a value.

  for (const ag_id of Object.keys(assignment_groups)) {
    let assignment_list = [];
    let assignments_url = URL_ASSIGNMENTS({
      CANVAS_BASE: CANVAS_BASE,
      courseId: courseId,
      assignment_group: ag_id
    });
    assignments_url += `?per_page=${PAGINATION_PER_PAGE}`;

    while (assignments_url !== null) {
      const resp = UrlFetchApp.fetch(assignments_url, {headers: auth_header});
      const headers = resp.getHeaders();
      const result = JSON.parse(resp.getContentText());

      for (let assignment of result) {
        if (! assignment.omit_from_final_grade && assignment.name !== 'Roll Call Attendance') {
          assignment_list.push(
          {
            id: assignment.id,
            name: assignment.name,
            points_possible: assignment.points_possible
          });
        }
      }

      if (headers['Link']) {
        assignments_url = getLinkNext(headers['Link']);
      }
      else {
        assignments_url = null;
      }
    }

    assignment_groups[ag_id].assignments = assignment_list;
  }
};

const fetch_students = (courseId, auth_header) => {
  let students_url = URL_STUDENTS({CANVAS_BASE: CANVAS_BASE, courseId: courseId});
  students_url += `?enrollment_type=student&per_page=${PAGINATION_PER_PAGE}`;

  let students = [];

  while(students_url !== null) {
    const resp = UrlFetchApp.fetch(students_url, {headers: auth_header});
    const headers = resp.getHeaders();
    const content = resp.getContentText();

    const result = JSON.parse(content);
    students = students.concat(result);

    if (headers['Link']) {
      students_url = getLinkNext(headers['Link']);
    }
    else {
      students_url = null;
    }
  }

  return students;
};

const fetch_grades = (studentId, courseId, nyg, auth_header) => {
  const userprops = PropertiesService.getUserProperties();
  let graded_assignments = {};

  message('Fetching grades from Canvas...');
  const assignment_groups = JSON.parse(userprops.getProperty('assignmentGroups'));
  if (assignment_groups === null) {
    SpreadsheetApp.getUi().alert('Error: could not get assignmentGroups from UserProperties.');
    return {};
  }

  for (const [ag_id, ag] of Object.entries(assignment_groups)) {

    graded_assignments[ag_id] = {};
    graded_assignments[ag_id]['name'] = ag.name;
    graded_assignments[ag_id]['weight'] = ag.group_weight;
    graded_assignments[ag_id]['assignments'] = [];

    for (const assignment of ag['assignments']) {
      let url = URL_SUBMISSIONS({
        CANVAS_BASE: CANVAS_BASE,
        courseId: courseId,
        assignmentId: assignment.id,
        studentId: studentId
      });
      let submission = JSON.parse(UrlFetchApp.fetch(url, {headers: auth_header}).getContentText());

      let score = parseFloat(submission['score']) || nyg;

      let never_drop;
      try {
        never_drop = ag['rules']['never_drop'].indexOf(assignment.id) !== -1;
      } catch {
        never_drop = false;
      }

      graded_assignments[ag_id]['assignments'].push(
          {
            'id': assignment.id,
            'name': assignment.name,
            'score': score,
            'points_possible': assignment.points_possible,
            'never_drop': never_drop,
            'dropped': false
          }
      );
    }

    // Apply drop_lowest and drop_highest rules to the assignments in current group.
    // Mark the ones that should be dropped.

    let droppableAssignments = graded_assignments[ag_id]['assignments'].filter(a => a.never_drop === false);
    let dropThese = [];

    // Drop n lowest and n highest assignments if those rules exist

    if (droppableAssignments.length > 0 && ag.hasOwnProperty('rules') && (
        ag['rules'].hasOwnProperty('drop_lowest') ||
        ag['rules'].hasOwnProperty('drop_highest'))) {
      droppableAssignments.sort((a, b) => (a.score/a.points_possible) - (b.score/b.points_possible));
      let n, a;
      if (ag['rules'].hasOwnProperty('drop_lowest')) {
        n = ag.rules.drop_lowest;

        while (n > 0) {
          a = droppableAssignments.shift();
          if (a !== undefined) {
            dropThese.push(a.id);
          }
          n -= 1;
        }
      }

      if (ag['rules'].hasOwnProperty('drop_highest')) {
        n = ag.rules.drop_highest;
        while (n > 0) {
          a = droppableAssignments.pop();
          if (a !== undefined) {
            dropThese.push(a.id);
          }
          n -= 1;
        }
      }

      // Use dropThese to mark the assignments to be dropped
      let i;

      for (const assgn of graded_assignments[ag_id]['assignments']) {
        if ((i = dropThese.indexOf(assgn.id)) !== -1) {
          assgn.dropped = true;
          dropThese.splice(i, 1);
        }
      }
    }
  }

  message('Done.');
  return graded_assignments;
};

const report = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cache = CacheService.getUserCache();

  let courseId = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('CourseID').getValue();
  if (courseId === '') {
    SpreadsheetApp.getUi().alert('Fill in the Course ID');
    return;
  }

  let nyg = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('NYG').getValue();
  if (nyg === '') {
    SpreadsheetApp.getUi().alert('Supply a value for Not-Yet-Graded');
    return;
  }
  else if (nyg === 'Use Zero') {
    nyg = 0;
  }

  const authHeader = JSON.parse(cache.get('authHeader'));
  if (authHeader === null) {
    SpreadsheetApp.getUi().alert('Please reauthorize from the Canvas menu (cache expiration).');
    return {};
  }

  const curCell = ss.getCurrentCell();
  const studentName = curCell.getValue();
  let writeTo = ss.getRange(STUDENT_NAME_CELL);
  writeTo.setValue(studentName);

  const idCell = ss.getRange(curCell.getRow(), curCell.getColumn() + 1);

  const graded_assignments = fetch_grades(idCell.getValue(), courseId, nyg, authHeader);

  // Render the assignments into the sheet

  // First clear any old data. Select cells that should contain headings and
  // expand data region downward, then clear that region.
  // REPORT_START is where the data begins; headings are one row up from that.
  writeTo = ss.getRange(REPORT_START);
  let headingRange = writeTo.offset(-1, 0, 1, REPORT_HEADINGS.length);
  const dataRange = headingRange.getDataRegion(SpreadsheetApp.Dimension.ROWS);
  dataRange.clear({contentsOnly: true});

  // Write headings first
  writeTo = ss.getRange(REPORT_START).offset(-1, 0);
  for (const heading of REPORT_HEADINGS) {
    writeTo.setValue(heading);
    writeTo = writeTo.offset(0, 1);
  }

  // Now the data
  writeTo = writeTo.offset(1, -REPORT_HEADINGS.length);

  for (const [groupKey, groupData] of Object.entries(graded_assignments)) {

    const groupName = groupData.name;
    writeTo.setValue(groupName);

    writeTo = writeTo.offset(1, 1);

    for (const assignment of groupData.assignments) {
      let columnsToBackUp = 0;

      writeTo.setValue(assignment.name);

      writeTo = writeTo.offset(0, 1);
      columnsToBackUp += 1;
      writeTo.setValue(assignment.points_possible);

      writeTo = writeTo.offset(0, 1);
      columnsToBackUp += 1;
      writeTo.setValue(assignment.score);

      writeTo = writeTo.offset(0, 1);
      columnsToBackUp += 1;
      writeTo.setValue(assignment.dropped ? DROP_STRING : '');

      writeTo = writeTo.offset(0, 1);
      columnsToBackUp += 1;
      let v:string = VALUE_TO_USE_FORMULA;
      v = v.replace(/_row_/g, writeTo.getRow().toString());
      writeTo.setFormula(v);

      writeTo = writeTo.offset(1, -columnsToBackUp);
    }

    writeTo = writeTo.offset(0, -1);
  }

  // Render the summary
  // Clear any old data
  const summaryRange = ss.getRange(SUMMARY_START).offset(0, 0, 1, SUMMARY_HEADINGS.length);
  summaryRange.getDataRegion(SpreadsheetApp.Dimension.ROWS).clearContent();

  // Replace headings we just cleared
  writeTo = ss.getRange(SUMMARY_START).offset(-1, 0);
  for (const summary_heading of SUMMARY_HEADINGS) {
    writeTo.setValue(summary_heading);
    writeTo = writeTo.offset(0, 1);
  }

  // Write the assignment group names and their weights
  writeTo = ss.getRange(SUMMARY_START);
  let groupCount = 0;

  const assignment_groups = JSON.parse(PropertiesService.getUserProperties().getProperty('assignmentGroups'));
  if (assignment_groups === null) {
    SpreadsheetApp.getUi().alert('Error: could not get assignmentGroups from UserProperties.');
    return;
  }

  for (const [agId, ag] of Object.entries(assignment_groups)) {
    groupCount += 1;
    writeTo.setValue(ag.name);
    writeTo = writeTo.offset(0, 1);
    writeTo.setValue(ag.group_weight/100.0);
    writeTo = writeTo.offset(1, -1);
  }

  // For the Average and Weight Avg., write formulas into the
  // cells for the first group and copy the formulas down for
  // the other groups.
  const first_avg_cell = ss.getRange(SUMMARY_START).offset(0, 2);
  first_avg_cell.setFormula(ASSIGNMENT_AVERAGE_FORMULA);
  first_avg_cell.offset(0, 1).setFormula(WEIGHTED_AVERAGE_FORMULA);
  first_avg_cell.offset(0, 0, 1, 2).copyTo(first_avg_cell.offset(1, 0, groupCount-1, 2));
};

const show_students = (students, ss) => {
  // Write student names into sheet. After exhausting students array,
  // if there are cells beneath that contain old names, clear them.
  let row = STUDENT_START_ROW;
  for (const student of students) {
    let cell = ss.getRange(row, 1);
    cell.setValue(student.sortable_name);
    cell = ss.getRange(row++, 2);
    cell.setValue(student.id);
  }

  let nextPair = ss.getRange(row++, 1, 1, 2);

  while (nextPair.getCell(1, 1).getValue() != '') {
    nextPair.clearContent();
    nextPair = ss.getRange(row++, 1, 1, 2);
  }
};

const fetchCourse = () => {
  const authHeader = authorize();
  if (!authHeader) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let courseId = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('CourseID').getValue();
  if (courseId === '') {
    SpreadsheetApp.getUi().alert('Fill in the Course ID');
    return;
  }

  const assignment_groups = fetch_assignment_groups(courseId, authHeader);
  fetch_assignment_data(assignment_groups, courseId, authHeader); // Modifies assignment_groups
  PropertiesService.getUserProperties().setProperty('assignmentGroups', JSON.stringify(assignment_groups));

  const students = fetch_students(courseId, authHeader);
  show_students(students, ss);
};
