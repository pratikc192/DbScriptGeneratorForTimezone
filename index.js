//https://dev.virtualearth.net/REST/v1/TimeZone/?query={query}&datetime={datetime_utc}&key={BingMapsKey}
import fetch from 'node-fetch';
import * as XLSX from 'xlsx/xlsx.mjs';
import * as fs from 'fs';
XLSX.set_fs(fs);

const apiUrl = 'https://dev.virtualearth.net/REST/v1/TimeZone/';
const apiKey = 'Auca4p_exxDUrZg0DodpLkAcVn2f-rhLk4gOXVXgf0flh78F6qiWE0kmWVzgDJ0b';

parseExcel("demo.xlsx").forEach(async (element) => {
  //console.log(element.data);
  let cityAndState='', timezone='', sqlQueries='';
  for(let clinic of element.data){
    cityAndState = clinic.city + ', ' + clinic.state;
    try{
      timezone = await getTimezone(cityAndState);
    } catch (e){
      console.error(e);
    }
    console.log("update segway.clinic set timezone='" + timezone + "' where id=" + clinic.id + ";" + "\n");
    sqlQueries = "update segway.clinic set timezone='" + timezone + "' where id=" + clinic.id + ";" + "\n" + sqlQueries;
  }

  writeToFile(sqlQueries);
});

function parseExcel(filename){
  const excelData = XLSX.readFile(filename);
  return Object.keys(excelData.Sheets).map(name => ({
      name,
      data: XLSX.utils.sheet_to_json(excelData.Sheets[name]),
  }));
}

async function getTimezone(cityAndState){
  let url = apiUrl + '?key=' + apiKey + '&query=' + cityAndState;
  console.log(url);
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`HTTP error! status: ${response.status}`);
  }
  const data = await response.json();
  return data.resourceSets[0].resources[0].timeZoneAtLocation[0].timeZone[0].ianaTimeZoneId;
}

async function writeToFile(data){
  //console.log('\n\nQUERIES:\n'+data);
  fs.writeFile('output.txt', data, err => {
    if (err) {
      console.error(err);
    }
  });
}


