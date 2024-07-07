# Fast Js Excel

`v2.0.1`
This is a package i made for myself but can surely be helpful to others, feel free to contribute if you like it.

> [!WARNING]
> This lib is the most-inclusive between my 3 utils libraries.
> If you don't need excel js but other file utils install [word-file-utils](https://github.com/alessioVelluso/WordFileUtils) or take a look at [utils-stuff](https://github.com/alessioVelluso/UtilsStuff) wich is the lighter package.
> **DO NOT INSTALL ALL THREE LIBS CAUSE ONE IS THE "PARENT" OF THE OTHER:**
> 1. `utils-stuff`
> 2. `word-file-utils` (including utils-stuff)
> 3. `fast-js-excel` (including exceljs, word-file-utils (including utils-stuff))
>
>So if you install word-file utils you can use the classes of utils-stuff and so on, choose the one for your pourpose.

## Install:
```bash
npm install fast-js-excel
```

With this package comes the [alessiovelluso/utils-stuff](https://www.npmjs.com/package/utils-stuff), a package of server/client utilities that can be helpful, be sure to check the library for all the different utilities you can use.
This one has basically two different classes `GenericUtils` & `ClientFilters`, you can import them just as explained from the lib documentation but using "fast-excel".
The same with  the [alessiovelluso/word-file-utils](https://github.com/alessioVelluso/WordFileUtils) wich is included too, giving the chance to use translations and csv/json/object basic parsing but with ease.
```ts
import { GenericUtils } from "word-file-utils"
```



At the moment, the package exports:
```ts
type WriteWorkbook<T extends GenericObject = GenericObject> = (output:string, worksheets:WfuWorksheet<T>[]) => Promise<void>;

export type CreateWorkbook<T extends GenericObject = GenericObject> = (worksheets:WfuWorksheet<T>[]) => Workbook
```


## A brief explanation of the methods:
##### 1. Create Workbook
```ts
createWorkbook: <T extends GenericObject = GenericObject>(worksheets:WfuWorksheet<T>[]) => Promise<Workbook>;
```
Returns an ExcelJs.Workbook ready to be passed with an api
```ts
const  wfu = new  Wfu({ separator:  "|" });

const  data = [{col1:"Test1",col2:"Test2"},{col1:"Test3",col2:"Test4"}]
wfu.createWorkbook<{ Key:string, Value:string }>([
	{
		name:  "Worksheet1", data,
		prepend: {
			title:  "Details", rows:1,
			data: { Test:  "A Text", Test2:  543543, "Test_Date":  new  Date() }
		}
	}
]);
```
*Look at test repos for other examples*


##### 2. Write Workbook
```ts
writeWorkbook:<T extends GenericObject = GenericObject>(output:string, worksheets:WfuWorksheet<T>[]) => Promise<void>;
```
Write a workbook locally using the related method `createWorkbook`

> WRITE THE OUTPOUT WITHOUT THE FINAL EXTENSION LIKE:
> ```ts
> writeWorkbook("../Files/test", [...])
> ```



## Types
```ts
import { Column } from  "exceljs"
import { GoogleTranslateLocales } from  "./translate.types"


export  interface  TranslationConfig { translatingCol:string, cultureFrom:GoogleTranslateLocales, cultureTo:GoogleTranslateLocales }

export  interface  TranslateCsvConfig  extends  TranslationConfig { csvFilepath:string, outFilepath:string, separator?:string }

export  interface  TranslationMakerConstructor { separator?:string, errorTranslationValue?:string, translationColumnName?:string }

export  interface  WfuExcelColumn  extends  Partial<Column> { name:string, parse?: 'date' };

export  type  WfuWorksheetDetails = { title:string, rows?:number, data: GenericObject, patternColor?: string }

export  interface  WfuWorksheet<T  extends  GenericObject = GenericObject> {
	name: string,
	data:T[],
	prepend?: WfuWorksheetDetails
	append?: WfuWorksheetDetails,
}
```
