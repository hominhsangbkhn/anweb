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

	// write values to C1..C6
	sheet.getCell('C1').value = entry.name ?? '';
	sheet.getCell('C2').value = entry.year ?? '';
	sheet.getCell('C3').value = entry.school ?? '';
	sheet.getCell('C4').value = entry.address ?? '';
	sheet.getCell('C5').value = entry.address2 ?? '';
	sheet.getCell('C6').value = entry.classcode ?? '';

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

// export for external use
module.exports = {
	fillDataToTemplate,
	readDataWithClasscodes,
};

// replaced CLI: use readDataWithClasscodes() and pass first item to fillDataToTemplate
if (require.main === module) {
	(async () => {
		try {
			const data = readDataWithClasscodes(); // reads data2.json and adds classcode
			if (!Array.isArray(data) || data.length === 0) {
				console.error('No records found in data2.json');
				process.exit(1);
			}
			const first = data[0];
			const outName = `filled_${first.code || 'record'}_${first.name2 || 'record'}.xlsx`;
			const saved = await fillDataToTemplate(first, outName);
			console.log('Saved:', saved);
		} catch (err) {
			console.error(err);
			process.exit(1);
		}
	})();
}