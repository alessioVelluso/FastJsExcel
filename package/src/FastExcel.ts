import { WfuWorksheet, WfuWorksheetDetails } from "../types/generic.types.js";
import { GenericObject } from "word-file-utils";
import ExcelJS, { Border, Borders, FillPattern, Workbook, Worksheet } from 'exceljs';


type ManageWorksheetDetails = (worksheet:Worksheet, detail:WfuWorksheetDetails,startY:number, startX:number) => number
export type WriteWorkbook<T extends GenericObject = GenericObject> = (output:string, worksheets:WfuWorksheet<T>[]) => Promise<void>;
export type CreateWorkbook<T extends GenericObject = GenericObject> = (worksheets:WfuWorksheet<T>[]) => Workbook


const manageWorksheetDetails:ManageWorksheetDetails = (worksheet, detail,startY, startX) => {
	const constFill = { type: 'pattern', pattern: 'solid', fgColor:{ argb: detail.patternColor ?? 'E6E5E5' } } as FillPattern;
	const commonBorder:Partial<Border> = {style:'thin',color:{argb:'A9A6A7'}}
	const constBorders:Partial<Borders> = {top:commonBorder,left:commonBorder,bottom:commonBorder,right:commonBorder}

	const detailTitleCell = worksheet.getRow(startY).getCell(startX)
	detailTitleCell.style.fill = constFill
	detailTitleCell.style.border = constBorders
	detailTitleCell.value = ' ' + detail.title;
	startY++;

	const values = Object.entries(detail.data);
	const xLimit = Math.ceil(values.length / (detail.rows ?? 1));
	let x = 0, y = 0, i = 0;
	for (const value of values) {
		if (i === (xLimit * (y+1))) { y++; x = 0; }
		const relatedCell = worksheet.getRow(startY + y).getCell(startX + x);
		relatedCell.style.fill = constFill,
		relatedCell.style.border = constBorders
		relatedCell.value = ` ${value[0]}:  ${value[1]}`;

		x++; i++;
	}

	return startY + y;
}

export const createWorkbook:CreateWorkbook = <T extends GenericObject = GenericObject>(worksheets:WfuWorksheet<T>[]) => {
	const workbook = new ExcelJS.Workbook();

	for (const ws of worksheets)
	{
		const worksheet = workbook.addWorksheet(ws.name);
		let initialY = 2
		let initialX = 2;

		if (ws.prepend) {
			const newY = manageWorksheetDetails(worksheet, ws.prepend, initialY, initialX)
			initialY = newY + 2;
		}

		const offsetCellAddress:string = worksheet.getRow(initialY).getCell(initialX).address;
		if (ws.data.length === 0)
		{
			const messageCells:string = worksheet.getRow(initialY + 2).getCell(initialX + 2).address;
			worksheet.mergeCells(`${offsetCellAddress}:${messageCells}`);

			worksheet.getCell(offsetCellAddress).value = 'no data found';
			worksheet.getCell(offsetCellAddress).alignment = { vertical: 'middle', horizontal: 'center' };
		}
		else
		{
			const columns = Object.keys(ws.data[0]).map(x => ({ name: x, filterButton: true }))
			const rows = ws.data.map(row => Object.values(row))
			worksheet.addTable({
				name: `Table_${ws.name}`,
				ref: offsetCellAddress,
				headerRow: true,
				style: { theme: 'TableStyleMedium2', showRowStripes: true },
				columns,
				rows
			});

			initialY += rows.length + 3;
		}

		if (ws.append) {
			const newY = manageWorksheetDetails(worksheet, ws.append, initialY, initialX)
			initialY = newY + 1;
		}


		worksheet.columns.forEach((col, i) => {
			let maxLength = 0;

			if (col.header) maxLength = col.header.length;

			// eachCell is undefined if no cells are present in column
			if (col.eachCell) col.eachCell({ includeEmpty: true }, cell => {
				const cellValue = cell.value ? cell.value.toString() : '';
				maxLength = Math.max(maxLength, cellValue.length);
			});

			worksheet.getColumn(i + 1).width = maxLength + 2;
		});

		worksheet.getColumn(1).width = 10;
	}


	return workbook
};

export const writeWorkbook:WriteWorkbook = async <T extends GenericObject = GenericObject>(output:string, worksheets:WfuWorksheet<T>[]):Promise<void> => {
	const wb = createWorkbook(worksheets);

	await wb.xlsx.writeFile(output + '.xlsx')
	console.log(`File saved in ${output}.xlsx`);
}
