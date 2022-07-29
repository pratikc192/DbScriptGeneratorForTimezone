//https://dev.virtualearth.net/REST/v1/TimeZone/?query={query}&datetime={datetime_utc}&key={BingMapsKey}
import fetch from 'node-fetch';
import * as XLSX from 'xlsx/xlsx.mjs';
import * as fs from 'fs';
XLSX.set_fs(fs);

const parseExcel = (filename) => {

  const excelData = XLSX.readFile(filename);

  return Object.keys(excelData.Sheets).map(name => ({
      name,
      data: XLSX.utils.sheet_to_json(excelData.Sheets[name]),
  }));
};

parseExcel("demo.xlsx").forEach(async (element) => {
  //console.log(element.data);
  let cityState='', id='', url='', timezone='', sqlQueries='';
  for(let clinic of element.data){
    id = clinic.id;
    cityState = clinic.city + ', ' + clinic.state;
    url='https://dev.virtualearth.net/REST/v1/TimeZone/?query='+cityState+'&key=Auca4p_exxDUrZg0DodpLkAcVn2f-rhLk4gOXVXgf0flh78F6qiWE0kmWVzgDJ0b';
    console.log(url);
    timezone = await getTimezone(url);
    //console.log(data.resourceSets[0].resources[0].timeZoneAtLocation[0].timeZone[0].ianaTimeZoneId);
    console.log('\n' + cityState +', timezone=' + timezone + ' where id='+ id + '\n');
    sqlQueries = 'update segway.clinic set timezone="' + timezone + '" where id="' + id + '";' + '\n' + sqlQueries;
  }

  writeToFile(sqlQueries);
});

async function getTimezone(url){
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


