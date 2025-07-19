import * as XLSX from "xlsx";
/**
 * Add a comment to an XLSX sheet
 * @param {WorkSheet} ws sheet
 * @param {String|Object} where reference to comment location, such as C1  or {r:1, c:1}
 * @param {String} t Text to push
 */
function pushComment(ws, where, t) {
  if (typeof where === "object") {
    where = XLSX.utils.encode_cell(where);
  }
  ws[where].c = ws[where].c || [];
  ws[where].c.hidden = true;
  ws[where].c.push({ a: "SurveyTool", t });
}

/** just findIndex with a throw
 * @param {string[]} COLUMNS column head array
 * @param {string} name column name
 */
function findCol(COLUMNS, name) {
  const i = COLUMNS.indexOf(name);
  if (i == -1) throw Error(`Could not find column ${name}`);
  return i;
}

export { pushComment, findCol };
