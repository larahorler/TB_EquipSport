const xlsx = require('xlsx');
const fs = require('fs');

// Charger le fichier Excel
const workbook = xlsx.readFile('/Users/larahoerler/Desktop/Equip/SEO/stations_list.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convertir la feuille de calcul en JSON
const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

// Définir les descriptions
const descriptionMs = `Discover our Equip Sport multisport station dedicated to [Name of the Sport]. Enjoy our high-quality, self-service equipment, easily accessible via our mobile app. Whether you're a casual or an enthusiast, our station is designed for you to discover an enjoyable and dynamic sports experience. Let’s move with Equip Sport and share moments of fun activity.

✔️ Easy rental via our "Equip Sport" app 
✔️ Station close to your playing field
✔️ Ideal for all levels, from beginner to expert`;

const descriptionSup = `Discover our Equip Sport station dedicated to renting Stand-Up Paddle (SUP) equipment. Enjoy our high-quality equipment (Board, Paddle, lifejacket) available in self-service via our mobile app. Whether you're a beginner or experienced, book and rent your SUP board easily to spend fun moments on the water.

✔️ Easy rental via our "Equip Sport" app 
✔️ Possibility to book in advance 
✔️ Station close to the water 
✔️ Equipment suitable for all levels
✔️ From CHF 12/per hour`;

const descriptionPadel = `Discover our Equip Sport station dedicated to Padel, where you can rent rackets and balls to play without any constraints. Enjoy our high-quality equipment accessible via our mobile app. Whether you're a beginner or an experienced player, book your Padel gear easily and enjoy fun moments with friends.

✔️ Easy rental via our "Equip Sport" app 
✔️ Possibility to book in advance 
✔️ Station located near Padel courts 
✔️ High-quality equipment for all levels
✔️ From CHF 7/per hour`;

// Define the short descriptions
const shortDescriptionMs = `Discover our Equip Sport multisport station for [Name of the Sport]. Enjoy high-quality, self-service equipment via our app. Move with Equip Sport and share moments of fun.`;

const shortDescriptionSup = `Discover our Equip Sport station dedicated to renting Stand-Up Paddle (SUP) equipment. Enjoy our high-quality equipment available in self-service via our mobile app.`;

const shortDescriptionPadel = `Discover our Equip Sport station for Padel. Rent rackets and balls easily. Enjoy high-quality equipment accessible via our mobile app.`;

// Identify the columns 'name', 'descriptionLong', and 'descriptionShort'
const nameColIndex = data[0].indexOf('name');
const descLongColIndex = data[0].indexOf('descriptionLong');
const descShortColIndex = data[0].indexOf('descriptionShort');

// Update the descriptions based on the sport in the 'name' column
for (let i = 1; i < data.length; i++) {
  const name = data[i][nameColIndex];
  if (name.includes('SUP')) {
    data[i][descLongColIndex] = descriptionSup;
    data[i][descShortColIndex] = shortDescriptionSup;
  } else if (name.includes('Padel')) {
    data[i][descLongColIndex] = descriptionPadel;
    data[i][descShortColIndex] = shortDescriptionPadel;
  } else {
    const sportName = name.split('(').pop().replace(')', '').trim();
    data[i][descLongColIndex] = descriptionMs.replace('[Name of the Sport]', sportName);
    data[i][descShortColIndex] = shortDescriptionMs.replace('[Name of the Sport]', sportName);
  }
}

// Convertir le JSON mis à jour en feuille de calcul
const newWorksheet = xlsx.utils.aoa_to_sheet(data);
workbook.Sheets[sheetName] = newWorksheet;

// Sauvegarder le fichier Excel mis à jour
const outputFilePath = '/Users/larahoerler/Desktop/Equip/SEO/stations_list_updated.xlsx';
xlsx.writeFile(workbook, outputFilePath);

console.log('File updated successfully!');