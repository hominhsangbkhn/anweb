const path = require('path');
const fs = require('fs');
const Excel = require('exceljs');

/**
 * Fill cells C1..C6 on worksheet named "data" in template3.xlsx and save to out folder.
 * Mapping: C1=code, C2=name, C3=year, C4=school, C5=address, C6=score
 * @param {Object} entry - object with fields: code,name,year,school,address,score
 * @param {string} outFilename - output filename (eg. 'filled_341.xlsx')
 * @returns {Promise<string>} full path to saved file
 */
async function fillDataToTemplate(entry, outFilename) {
	// resolve paths
	const rootDir = __dirname; // g:\Frontend\Anweb\backend
	const templatePath = path.join(rootDir, 'template3.xlsx');
	const outDir = path.join(rootDir, 'out');

	// ensure template exists
	if (!fs.existsSync(templatePath)) {
		throw new Error(`Template not found: ${templatePath}`);
	}
	// ensure out dir exists
	if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

	const workbook = new Excel.Workbook();
	await workbook.xlsx.readFile(templatePath);

	const sheet = workbook.getWorksheet('data');
	if (!sheet) throw new Error('Worksheet "data" not found in template');

	// write values to C1..C6 using helper that respects merged cells
	setCellValueKeepingMerge(sheet, 'C1', entry.name ?? '');
	setCellValueKeepingMerge(sheet, 'C2', entry.year ?? '');
	setCellValueKeepingMerge(sheet, 'C3', entry.school ?? '');
	setCellValueKeepingMerge(sheet, 'C4', entry.address ?? '');
	setCellValueKeepingMerge(sheet, 'C5', entry.address2 ?? '');
	setCellValueKeepingMerge(sheet, 'C6', entry.classcode ?? '');

	const outPath = path.join(outDir, outFilename);
	await workbook.xlsx.writeFile(outPath);
	return outPath;
}

/**
 * Read JSON array from dataPath (defaults to ./data2.json) and add classcode:
 * classcode = 18 for indices 0-19, 19 for 20-39, etc.
 * @param {string} [dataPath] - optional path to JSON file
 * @returns {Array<Object>} array of objects with added `classcode` property
 */
function readDataWithClasscodes(dataPath) {
	const file = dataPath ? path.resolve(dataPath) : path.join(__dirname, 'data2.json');
	if (!fs.existsSync(file)) throw new Error(`Data file not found: ${file}`);
	const raw = fs.readFileSync(file, 'utf8');
	const arr = JSON.parse(raw);
	if (!Array.isArray(arr)) throw new Error('Data file does not contain an array');
	return arr.map((item, idx) => {
		const classcode = 18 + Math.floor(idx / 20);
		return { ...item, classcode };
	});
}

/**
 * Create copies of worksheet "form" from template-all.xlsx.
 * Each copy preserves styles, merges, column widths, row heights, views, pageSetup.
 * Writes item data into C6..C10 while keeping the copied cell styles.
 */
async function createCopiesFromTemplateAll(dataArray, outFilename = 'template-all-filled.xlsx', opts = {}) {
	const rootDir = __dirname;
	const templateFile = opts.templateFile || 'template-all.xlsx';
	const formSheetName = opts.formSheetName || 'form';
	const templatePath = path.join(rootDir, templateFile);
	const outDir = path.join(rootDir, 'out');
	if (!fs.existsSync(templatePath)) {
		throw new Error(`Template not found: ${templatePath}`);
	}
	if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

	const workbook = new Excel.Workbook();
	await workbook.xlsx.readFile(templatePath);

	const formSheet = workbook.getWorksheet(formSheetName);
	if (!formSheet) throw new Error(`Worksheet "${formSheetName}" not found in ${templateFile}`);

	// helper: extract merge ranges reliably from a worksheet
	function getMergeRanges(sheet) {
		const ranges = new Set();
		try {
			if (sheet && sheet.model && sheet.model.merges) {
				Object.keys(sheet.model.merges).forEach(r => ranges.add(r));
			}
			if (sheet && sheet._merges && typeof sheet._merges.keys === 'function') {
				for (const k of sheet._merges.keys()) ranges.add(k);
			} else if (sheet && sheet._merges && typeof sheet._merges === 'object') {
				Object.keys(sheet._merges).forEach(k => ranges.add(k));
			}
		} catch (e) {
			// ignore and return what we have
		}
		console.log("Rangers: " , Array.from(sheet.model.merges));
		
		return Array.from(sheet.model.merges);
	}

	// gather merge ranges from source sheet using helper
	let mergeRanges = getMergeRanges(formSheet);

	// For each item create a copy sheet named by index (string)
	for (let i = 0; i < dataArray.length; i++) {
		const item = dataArray[i] || {};
		const sheetName = "STT-" + String(i); // name is index

		// if a sheet with same name exists, remove it first
		const existing = workbook.getWorksheet(sheetName);
		if (existing) workbook.removeWorksheet(existing.id);

		// create new sheet with same views/pageSetup/properties
		const newSheet = workbook.addWorksheet(sheetName, {
			properties: formSheet.properties || {},
			pageSetup: formSheet.pageSetup || {},
			views: formSheet.views || []
		});

		// preserve internal model and merges so merged ranges survive
		try {
			// clone the sheet.model (structure includes merges, rows, cols metadata)
			if (formSheet.model) {
				newSheet.model = JSON.parse(JSON.stringify(formSheet.model));
				// keep the new name (model may contain original name)
				newSheet.name = sheetName;
			}
			// clone internal _merges map if present
			if (formSheet._merges && typeof formSheet._merges.entries === 'function') {
				newSheet._merges = new Map(formSheet._merges);
			}
		} catch (e) {
			// non-fatal: fall back to applying merges later
		}

		// copy columns (width, key, style, hidden, outlineLevel)
		if (Array.isArray(formSheet.columns)) {
			newSheet.columns = formSheet.columns.map(col => {
				const copy = {
					key: col && col.key,
					header: col && col.header,
					width: col && col.width,
					outlineLevel: col && col.outlineLevel,
					hidden: col && col.hidden,
				};
				// copy column style if present
				if (col && col.style && Object.keys(col.style).length > 0) {
					copy.style = JSON.parse(JSON.stringify(col.style));
				}
				return copy;
			});
		}

		// copy rows and cells (values and styles) and row heights
		formSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
			const targetRow = newSheet.getRow(rowNumber);
			// preserve row height and outline/hidden flags
			if (row.height) targetRow.height = row.height;
			if (row.outlineLevel) targetRow.outlineLevel = row.outlineLevel;
			if (row.hidden) targetRow.hidden = row.hidden;

			row.eachCell({ includeEmpty: true }, (cell) => {
				// copy only source master cells (avoid writing non-master cells inside merged ranges)
				const sourceMaster = cell.master || cell;
				if (sourceMaster.address !== cell.address) return;

				// destination cell before merges applied: use plain address
				let targetCell = newSheet.getCell(cell.address);

				// copy value (including objects like {richText}, formula, etc.)
				targetCell.value = cell.value;

				// copy full style (font, fill, border, alignment, numFmt, protection)
				if (cell.style && Object.keys(cell.style).length > 0) {
					try {
						targetCell.style = JSON.parse(JSON.stringify(cell.style));
					} catch (e) {
						const s = cell.style;
						if (s.font) targetCell.font = JSON.parse(JSON.stringify(s.font));
						if (s.fill) targetCell.fill = JSON.parse(JSON.stringify(s.fill));
						if (s.border) targetCell.border = JSON.parse(JSON.stringify(s.border));
						if (s.alignment) targetCell.alignment = JSON.parse(JSON.stringify(s.alignment));
						if (s.numFmt) targetCell.numFmt = s.numFmt;
					}
				}

				// preserve type/formula/result if present
				if (cell.type === Excel.ValueType.Formula && cell.value && cell.value.formula) {
					targetCell.value = { formula: cell.value.formula, result: cell.value.result };
				}
			});
			targetRow.commit && targetRow.commit();
		});

		// Apply merges after copying master cells so merged areas are created correctly
		try {
			for (const range of mergeRanges) {
				if (typeof range === 'string' && range.includes(':')) {
					try { newSheet.mergeCells(range); } catch (e) { /* ignore individual failures */ }
				}
			}
		} catch (e) {
			// ignore
		}

		// copy sheet-level pageMargins, headerFooter if present
		if (formSheet.pageMargins) newSheet.pageMargins = JSON.parse(JSON.stringify(formSheet.pageMargins));
		if (formSheet.headerFooter) newSheet.headerFooter = JSON.parse(JSON.stringify(formSheet.headerFooter));
		if (formSheet.properties && formSheet.properties.tabColor) newSheet.properties = newSheet.properties || {}, newSheet.properties.tabColor = JSON.parse(JSON.stringify(formSheet.properties.tabColor));

		// fill target cells C6..C10 with item data (only change value to keep existing style)
		setCellValueKeepingMerge(newSheet, 'C6', item.name ?? '');
		setCellValueKeepingMerge(newSheet, 'C7', item.year ?? item.name2 ?? '');
		setCellValueKeepingMerge(newSheet, 'C8', item.school ?? '');
		setCellValueKeepingMerge(newSheet, 'C9', item.address ?? '');
		setCellValueKeepingMerge(newSheet, 'C10', item.address2 ?? '');
		setCellValueKeepingMerge(newSheet, 'F6', "Mã lớp: 18");

		// Re-apply merges once more after writing values to ensure merged ranges persist in output
		try {
			console.log(mergeRanges);
			
			for (const range of mergeRanges) {
				if (typeof range === 'string' && range.includes(':')) {
					console.log('Tôi yêu An:', range);
					
					try { newSheet.mergeCells(range); } catch (e) { /* ignore */ }
				}
			}
		} catch (e) {
			// ignore
		}
	}

	const outPath = path.join(outDir, outFilename);
	await workbook.xlsx.writeFile(outPath);
	return outPath;
}

/**
 * Copy worksheet "form" from template-all.xlsx into the same workbook as a new sheet.
 * Preserves columns, row heights, cell styles and merged ranges.
 * @param {string} newSheetName - name for the copied sheet
 * @param {string} [outFilename='template-all-copy.xlsx'] - output file saved to ./out
 * @param {Object} [opts] - { templateFile: 'template-all.xlsx', formSheetName: 'form' }
 * @returns {Promise<string>} - saved file path
 */
async function copyFormSheetToFile(newSheetName, outFilename = 'template-all-copy.xlsx', opts = {}) {
	const rootDir = __dirname;
	const templateFile = opts.templateFile || 'template-all.xlsx';
	const formSheetName = opts.formSheetName || 'form';
	const templatePath = path.join(rootDir, templateFile);
	const outDir = path.join(rootDir, 'out');

	if (!fs.existsSync(templatePath)) throw new Error(`Template not found: ${templatePath}`);
	if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

	const workbook = new Excel.Workbook();
	await workbook.xlsx.readFile(templatePath);

	const formSheet = workbook.getWorksheet(formSheetName);
	if (!formSheet) throw new Error(`Worksheet "${formSheetName}" not found in ${templateFile}`);

	// create new sheet with same sheet-level options
	const sheetOptions = {
		properties: formSheet.properties || {},
		pageSetup: formSheet.pageSetup || {},
		views: formSheet.views || []
	};
	const newSheet = workbook.addWorksheet(newSheetName, sheetOptions);

	// copy columns
	if (Array.isArray(formSheet.columns)) {
		newSheet.columns = formSheet.columns.map(col => {
			const copy = {
				key: col && col.key,
				header: col && col.header,
				width: col && col.width,
				outlineLevel: col && col.outlineLevel,
				hidden: col && col.hidden,
			};
			if (col && col.style && Object.keys(col.style).length > 0) copy.style = JSON.parse(JSON.stringify(col.style));
			return copy;
		});
	}

	// gather merge ranges
	let mergeRanges = [];
	if (formSheet.model && formSheet.model.merges) {
		mergeRanges = Object.keys(formSheet.model.merges || {});
	} else if (formSheet._merges && typeof formSheet._merges.keys === 'function') {
		for (const k of formSheet._merges.keys()) mergeRanges.push(k);
	}

	// copy rows & master cells (do not overwrite merged non-master cells)
	formSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
		const targetRow = newSheet.getRow(rowNumber);
		if (row.height) targetRow.height = row.height;
		if (row.outlineLevel) targetRow.outlineLevel = row.outlineLevel;
		if (row.hidden) targetRow.hidden = row.hidden;

		row.eachCell({ includeEmpty: true }, (cell) => {
			const sourceMaster = cell.master || cell;
			if (sourceMaster.address !== cell.address) return; // skip non-master cells

			let targetCell = newSheet.getCell(cell.address);

			// copy value
			targetCell.value = cell.value;

			// copy style
			if (cell.style && Object.keys(cell.style).length > 0) {
				try {
					targetCell.style = JSON.parse(JSON.stringify(cell.style));
				} catch (e) {
					const s = cell.style;
					if (s.font) targetCell.font = JSON.parse(JSON.stringify(s.font));
					if (s.fill) targetCell.fill = JSON.parse(JSON.stringify(s.fill));
					if (s.border) targetCell.border = JSON.parse(JSON.stringify(s.border));
					if (s.alignment) targetCell.alignment = JSON.parse(JSON.stringify(s.alignment));
					if (s.numFmt) targetCell.numFmt = s.numFmt;
				}
			}

			// preserve formula/result
			if (cell.type === Excel.ValueType.Formula && cell.value && cell.value.formula) {
				targetCell.value = { formula: cell.value.formula, result: cell.value.result };
			}
		});
		targetRow.commit && targetRow.commit();
	});

	// apply merges after copying master cells
	try {
		for (const range of mergeRanges) {
			if (typeof range === 'string' && range.includes(':')) {
				try { newSheet.mergeCells(range); } catch (e) { /* ignore */ }
			}
		}
	} catch (e) {
		// ignore
	}

	// copy sheet-level extras
	if (formSheet.pageMargins) newSheet.pageMargins = JSON.parse(JSON.stringify(formSheet.pageMargins));
	if (formSheet.headerFooter) newSheet.headerFooter = JSON.parse(JSON.stringify(formSheet.headerFooter));
	if (formSheet.properties && formSheet.properties.tabColor) {
		newSheet.properties = newSheet.properties || {};
		newSheet.properties.tabColor = JSON.parse(JSON.stringify(formSheet.properties.tabColor));
	}

	// re-apply merges after any changes (helps ensure persistence)
	try {
		for (const range of mergeRanges) {
			if (typeof range === 'string' && range.includes(':')) {
				try { newSheet.mergeCells(range); } catch (e) { /* ignore */ }
			}
		}
	} catch (e) {
		// ignore
	}

	const outPath = path.join(outDir, outFilename);
	await workbook.xlsx.writeFile(outPath);
	return outPath;
}

// add helper to write into merged cells safely
function setCellValueKeepingMerge(sheet, address, value) {
	// if address is in a merged range, cell.master points to top-left master cell
	const cell = sheet.getCell(address);
	const target = cell.master || cell;
	target.value = value;
}

// export for external use
module.exports = {
	fillDataToTemplate,
	readDataWithClasscodes,
	createCopiesFromTemplateAll,
	copyFormSheetToFile,
};

// replaced CLI: use readDataWithClasscodes() and call createCopiesFromTemplateAll for all records
if (require.main === module) {
	(async () => {
		try {
			const data = readDataWithClasscodes(); // reads data2.json and adds classcode
			
			
			if (!Array.isArray(data) || data.length === 0) {
				console.error('No records found in data2.json');
				process.exit(1);
			}

			// use only the first 20 items
			const _23 = data.slice(100,119);
			const outName = `template-all-filled_23.xlsx`;
			const saved = await createCopiesFromTemplateAll(_23, outName);
			console.log('Saved:', saved);
		} catch (err) {
			console.error(err);
			process.exit(1);
		}
	})();
}