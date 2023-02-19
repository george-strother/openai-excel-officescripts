// OpenAI in Excel Script.  Meant for usage within Power Automate and Excel Online

async function cellRangeToJSON(workbook: ExcelScript.Workbook, sheetName: string, range: string) {
    const sheet = workbook.getWorksheet(sheetName);
    const rangeValues = sheet.getRange(range).getValues();
    const rangeValuesJSON = JSON.stringify(rangeValues);
    return rangeValuesJSON;
}

async function main(workbook: ExcelScript.Workbook) {
    
    // Set the OpenAI API key - You'll need to add this in the Excel file or replace this part with your key
    const apiKey = workbook.getWorksheet("API").getRange("B1").getValue();
    const endpoint: string = "https://api.openai.com/v1/completions";

    // get worksheet info
    const promptSheet = workbook.getWorksheet("Prompt");
    // the ask - if cell D5 of Prompt contains False, we'll just use the value of B2
    //    if it contains True, we'll combine the value of B2 with the cellRangeToJSON function
    //     then we'll combine the cell value of B2
    //     with the CellRangeToJSON function of 
    //     the "Data" worksheet and the range of A1:B1000
    //     otherwise we just use the value of B2
    const splitString = "/* Based on the first 1000 rows of data in the 'Data' sheet rows 1 to 1000 as json: */";
    const mytext = promptSheet.getRange("D5").getValue() ? promptSheet.getRange("B2").getValue() + splitString + cellRangeToJSON(workbook, "Data", "A1:B1000") : promptSheet.getRange("B2").getValue();
    
    // useful if if we get more than one row back
    const result = workbook.getWorksheet("Result");
    result.getRange("A1:D1000").clear();
    promptSheet.getRange("B3").setValue(" ")

    // Set the model engine and prompt
    const model: string = "text-davinci-002";
  const prompt: (string | boolean | number) = mytext;

    // Set the HTTP headers
    const headers: Headers = new Headers();
    headers.append("Content-Type", "application/json");
    headers.append("Authorization", `Bearer ${apiKey}`);

    // Set the HTTP body
  const body: (string | boolean | number) = JSON.stringify({
        model: model,
        prompt: prompt,
        max_tokens: 1024,
        n: 1,
        temperature: 0.5,
    });

    // Send the HTTP request
    const response: Response = await fetch(endpoint, {
        method: "POST",
        headers: headers,
        body: body,
    });

    // Parse the response as JSON
  const json: { choices: { text:(string | boolean | number )}[] } = await response.json();

    // Get the answer - i.e. output
    const text: (string | boolean | number) = json.choices[0].text;

    // Output the generated text
   // console.log(text);
  
   const output = promptSheet.getRange("B4");
   
   output.setValue(text);

  const cell = promptSheet.getRange("B4");

  // Split the cell contents by new line

  const arr = cell.getValue().toString().split("\n");

  const newcell = result.getRange("A1");

  var offset = 0;
  // console.log (arr)

  for (let i = 0; i < arr.length; i++) {
    // Write the value to the next cell
   
    if (arr[i].length > 0) {
      newcell.getOffsetRange(offset, 0).setValue(arr[i]);
    
      offset++;
    }
  }

 // console.log(offset)
  if (offset > 1) {
    promptSheet.getRange("B3").setValue("Check 'Result' sheet to get answers separated by multiple rows")

  }
}


