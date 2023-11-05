/* eslint-disable @typescript-eslint/no-unused-vars */
// レスポンス
type Row = {
	[key:string]: string
}

type Result = {
	row: Row,
	index: number,
}

function _getColumns(sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] {
	return sheet.getRange("1:1").getValues()[0] as string[];
}

// 取得用の関数
function _get(sheet: GoogleAppsScript.Spreadsheet.Sheet, id: string, key = 'id'): Result {
	// カラム一覧
  const colmuns = _getColumns(sheet);
	// キーがあるカラムの位置
  const columnId = colmuns.indexOf(key) + 1;
	// 最終行
  const lastRow = sheet.getLastRow();
	// 最終行からキーが一致する行を検索
  const values = sheet.getRange(2, columnId, lastRow - 1, 1)
		.getValues()
		.map(a => `${a[0]}`);
	const rowId = values.indexOf(`${id}`) + 2;
	// 見つからなかった場合はエラー
	if (rowId === 1) return {row: {}, index: -1};
	// 見つかった場合はオブジェクトにして返す
	const row = sheet.getRange(rowId, 1, 1, colmuns.length)
		.getValues()[0] as string[];
	const obj = colmuns.reduce((obj, column, i) => {
		if (column === '') return obj;
		obj[column] = row[i];
		return obj;
	}
	, {} as Row);
	return {
		row: obj,
		index: rowId,
	};
}

// 一覧取得用の関数
// limit: 取得件数
// skip: スキップ件数
// all: _から始まるカラムを含めるかどうか
function _list(sheet: GoogleAppsScript.Spreadsheet.Sheet, limit: number, skip: number, all: boolean): Row[] {
	const rows = _getData(sheet, limit, skip);
	const keys = _getColumns(sheet);
	return rows.map(row => {
		const obj = {};
		row.forEach((item, index) => {
			if (!all && keys[index].indexOf('_') === 0) return;
			obj[String(keys[index])] = item;
		});
		return obj;
	});
}

function strToNum(str: string | number | undefined): number {
	if (typeof str === 'undefined') return 0;
	if (!str) return 0;
	if (!isNaN(Number(str))) return str as number;
	if (typeof str === 'string' && str.trim() === '') return 0;
	if (typeof str === 'string' && isNaN(parseInt(str))) return 0;
	return parseInt(str as string);
}

function _getData(sheet: GoogleAppsScript.Spreadsheet.Sheet, limit: number, skip: number): any[][] {
	if (limit === 0 && skip === 0) {
		return sheet.getDataRange().getValues().splice(1);
	}
	const lastColumn = sheet.getLastColumn();
	const lastRow = sheet.getLastRow();
	if (skip > lastRow + 2) return [[]];
	const ary = sheet.getRange(skip + 2, 1, lastRow - skip - 1, lastColumn).getValues();
	if (limit === 0) return ary;
	return ary.splice(0, limit);
}

function _find(sheet: GoogleAppsScript.Spreadsheet.Sheet, term: string): Row[] {
	const textFinder = sheet.createTextFinder(term);
	const keys = _getColumns(sheet);
	const lastColumn = sheet.getLastColumn();
	const ary:Row[] = [];
	while (textFinder.findNext()) {
		const row = textFinder.getCurrentMatch()!.getRow();
		const values = sheet.getRange(row, 1, 1, lastColumn).getValues()[0] as string[];
		const obj = keys.reduce((obj, key, index) => {
			obj[key] = values[index];
			return obj;
		}, {} as Row);
		ary.push(obj);
	}
	return ary;
}

function create(sheet: GoogleAppsScript.Spreadsheet.Sheet, obj: Row, key = 'id'): GoogleAppsScript.Content.TextOutput {
	const keys = _getColumns(sheet);
	if (typeof obj[key] === 'undefined' || !obj[key] || obj[key] === '') {
		obj[key] = Utilities.getUuid();
	}
	// 重複チェック
	const { row } = _get(sheet, obj[key], key);
	if (Object.keys(row).length > 0) {
		const output = ContentService.createTextOutput();
		output.setMimeType(ContentService.MimeType.JSON);
		output.setContent(JSON.stringify({error: 'duplicate'}));
		return output;
	}
	const newRow = sheet.getLastRow() + 1;
	for (const key of keys) {
		if (typeof obj[key] === 'undefined') continue;
		sheet.getRange(newRow, keys.indexOf(key) + 1).setValue(obj[key]);
	}
	const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
	output.setContent(JSON.stringify(obj));
	return output;
}

// GETリクエスト
function get(sheet: GoogleAppsScript.Spreadsheet.Sheet, id: string, key = 'id'): GoogleAppsScript.Content.TextOutput {
	const { row, index } = _get(sheet, id, key);
	const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
	if (index === -1) {
		output.setContent(JSON.stringify({error: 'not found'}));
	} else {
		output.setContent(JSON.stringify(row));
	}
	return output;
}

function list(sheet: GoogleAppsScript.Spreadsheet.Sheet, limit: string | undefined, skip: string | undefined, all = false): GoogleAppsScript.Content.TextOutput {
	const output = ContentService.createTextOutput();
	output.setMimeType(ContentService.MimeType.JSON);
	output.setContent(JSON.stringify(_list(sheet, strToNum(limit), strToNum(skip), all)));
	return output;
}

function update(sheet: GoogleAppsScript.Spreadsheet.Sheet, id: string, obj: Row, key = 'id'): GoogleAppsScript.Content.TextOutput {
	const keys = _getColumns(sheet);
	const rowId = _get(sheet, id, key).index;
	if (rowId === -1) {
		const output = ContentService.createTextOutput();
		output.setMimeType(ContentService.MimeType.JSON);
		output.setContent(JSON.stringify({error: 'not found'}));
		return output;
	}
	for (const k of keys) {
		if (typeof obj[k] === 'undefined') continue;
		sheet.getRange(rowId, keys.indexOf(k) + 1).setValue(obj[k]);
	}
	return get(sheet, id, key);
}

function remove(sheet: GoogleAppsScript.Spreadsheet.Sheet, id: string, key = 'id'): GoogleAppsScript.Content.TextOutput {
	const rowId = _get(sheet, id, key).index;
	const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
	if (rowId === -1) {
		output.setContent(JSON.stringify({error: 'not found'}));
		return output;
	}
	sheet.deleteRow(rowId);
  output.setContent(JSON.stringify({result: 'ok'}));
	return output;
}

function removeAll(sheet: GoogleAppsScript.Spreadsheet.Sheet): GoogleAppsScript.Content.TextOutput {
	const lastRow = sheet.getLastRow();
	sheet.deleteRows(2, lastRow - 1);
	const output = ContentService.createTextOutput();
	output.setMimeType(ContentService.MimeType.JSON);
	output.setContent(JSON.stringify({result: 'ok'}));
	return output;
}

function count(sheet: GoogleAppsScript.Spreadsheet.Sheet): GoogleAppsScript.Content.TextOutput {
	const lastRow = sheet.getLastRow();
	const count = lastRow - 1;
	const output = ContentService.createTextOutput();
	output.setMimeType(ContentService.MimeType.JSON);
	output.setContent(JSON.stringify({count}));
	return output;
}

function find(sheet: GoogleAppsScript.Spreadsheet.Sheet, term: string): GoogleAppsScript.Content.TextOutput {
	const output = ContentService.createTextOutput();
	output.setMimeType(ContentService.MimeType.JSON);
	output.setContent(JSON.stringify(_find(sheet, term)));
	return output;
}
