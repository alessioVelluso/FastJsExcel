import { writeWorkbook } from "fast-excel"
import { WordFileUtils } from "word-file-utils"
import gu from "./utils";

const wfu = new WordFileUtils();


const data = wfu.parseCsvToObjectList("../Files/MockData.csv", ",");
writeWorkbook("../Files/Test", [
    {
        name: "Worksheet_1", data,
        prepend: {
			title:  "Details", rows:1,
			data: { Test:  "A Text", Test2:  543543, "Test_Date":  new  Date() }
		}
    }
]);


gu.log("Hello this is another library for logging and utils", "cyan")
