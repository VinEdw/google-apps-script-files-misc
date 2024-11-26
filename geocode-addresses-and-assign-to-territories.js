// Section: Column Names

const addressCols = [
  "ImportedSuccessfully",
  "TerritoryID",
  "TerritoryNumber",
  "Category",
  "ApartmentNumber",
  "Number",
  "Street",
  "Suburb",
  "PostalCode",
  "State",
  "Name",
  "Phone",
  "Status",
  "Latitude",
  "Longitude",
  "Notes",
];

const territoryCols = [
  "TerritoryID",
  "CategoryCode",
  "Category",
  "Number",
  "Suffix",
  "Area",
  "Type",
  "Link1",
  "Link2",
  "CustomNotes1",
  "CustomNotes2",
  "Boundary",
];

// Section: Utilities

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Scripts")
    .addItem("Geocode Selected Addresses", "geocodeAddresses")
    .addItem("Assign Selected Addresses to Territories", "assignTerritories")
    .addToUi();
}

function getActiveRows() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SS.getActiveSheet();
  const range = sheet.getActiveRange();

  let startingRow = range.getRow();
  let height = range.getHeight();

  if (startingRow === 1) {
    if (height === 1) {
      throw new Error("Cannot operate on header row");
    }
    else {
      ++startingRow;
      --height;
    }
  }

  return sheet.getRange(startingRow, 1, height, sheet.getLastColumn());
}

function setColumnOfRange(range, columnIdx, values) {
  return range.offset(0, columnIdx, range.getHeight(), 1)
    .setValues(values.map(x => [x]));
}

// Section: Geocoding

function geocodeAddresses() {
  const rowRange = getActiveRows();
  if (rowRange.getSheet().getName() !== "Addresses") {
    throw new Error("Must be on 'Addresses' sheet");
  }

  const latValues = [];
  const lngValues = [];

  for (const row of rowRange.getValues()) {
    const number = row[addressCols.indexOf("Number")];
    const street = row[addressCols.indexOf("Street")];
    const suburb = row[addressCols.indexOf("Suburb")];
    const postalCode = row[addressCols.indexOf("PostalCode")];
    const state = row[addressCols.indexOf("State")];

    let address = `${number} ${street}, ${suburb}, ${state} ${postalCode}`;

    const geocodeData = Maps.newGeocoder()
      .setLanguage("en")
      .setBounds(24.712577, -126.203344, 50.971278, -58.593496)
      .geocode(address);

    if (geocodeData.status === "OK") {
      const lat = geocodeData.results[0].geometry.location.lat;
      const lng = geocodeData.results[0].geometry.location.lng;
      latValues.push(lat);
      lngValues.push(lng);
    }
    else {
      latValues.push(null);
      lngValues.push(null);
    }
  }

  setColumnOfRange(rowRange, addressCols.indexOf("Latitude"), latValues);
  setColumnOfRange(rowRange, addressCols.indexOf("Longitude"), lngValues);
}

// Section: Territory Assigning

function boundaryStringToArray(boundaryString) {
  return boundaryString.slice(1, -1)
    .split("],[")
    .map(str => str.split(",")
      .map(num => parseFloat(num))
      .reverse()
    );
}

function getTerritories() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SS.getSheetByName("Territories");
  const range = sheet.getDataRange();

  const territories = [];
  for (const row of range.getValues().slice(1)) {
    const territory = {
      id: row[territoryCols.indexOf("TerritoryID")],
      number: row[territoryCols.indexOf("Number")],
      category: row[territoryCols.indexOf("Category")],
      boundary: boundaryStringToArray(row[territoryCols.indexOf("Boundary")]),
    }
    territories.push(territory);
  }

  return territories;
}

// Function checks if a ray starting at a point (lat0, lng0) and extended along the longitude to greater latitudes
// intersects with a line segment connecting (lat1, lng1) and (lat2, lng2)
// Returns true if the ray does intersects
// Returns false if the ray does not intersect
//
// If the ray passes through at least one of the points, then only count the intersection if the other point has a greater longitude
// This prevents double counting the intersection in point_in_polygon()
// Effectively, this nudges the point to a slightly higher longitude for the purpose of the test
//
// Cases where the point is on the segment do not count
function checkForIntersection(lat0, lng0, lat1, lng1, lat2, lng2) {
  // Check if the line segment does not even cross the latitude of the ray
  // This happens when:
  // - Both ends of the line segment have a latitude that is too low
  // - Both ends of the line segment have a longitude that is too high
  // - Both ends of the line segment have a longitude that is too low
  if ((Math.max(lat1, lat2) < lat0) || (Math.min(lng1, lng2) > lng0) || (Math.max(lng1, lng2) < lng0)) {
    return false;
  }
  // If the ray passes through at least one of the points, then only count the intersection if the other point has a greater longitude
  // This prevents double counting the intersection in point_in_polygon()
  if ((lng1 === lng0 && lat1 > lat0) || (lng2 === lng0 && lat2 > lat0)) {
    return Math.max(lng1, lng2) > lng0;
  }
  // After those checks, if both ends of the line segment have latitudes that are greater than lat0, then the ray definitely intersects
  if (Math.min(lat1, lat2) > lat0) {
    return true;
  }
  // Calculate the exact latitude where the line crosses the longitude level of the ray
  // If that latitude value is greater than the latitude of the point, then the line intersects the ray
  // Otherwise, the line does not intersect the ray
  let slope = (lat2 - lat1) / (lng2 - lng1);
  let lat_intersect = slope * (lng0 - lng1) + lat1;
  return lat_intersect > lat0;
}

// Function checks if a (lat, lng) point lies within a polygon defined by a list of points
// Returns true if the point is inside
// Returns false if the point is outside
// Points that lie on the edge of the polygon return an ambiguous result
function pointInPolygon(lat, lng, polygonPoints) {
  let intersection_count = 0;
  for (let i = 0; i < polygonPoints.length; ++i) {
    // Set the initial point
    let pair1 = polygonPoints[i];
    let lat1 = pair1[0];
    let lng1 = pair1[1];
    // Set the next point
    let pair2 = polygonPoints[(i + 1) % polygonPoints.length];
    let lat2 = pair2[0];
    let lng2 = pair2[1];
    // Check for an intersection between the ray extended from the point and the polygon edge formed by the points
    intersection_count += checkForIntersection(lat, lng, lat1, lng1, lat2, lng2);
  }
  // If the intersection_count is even, then the point is outside the polygon
  // If the intersection_count is odd, then the point is inside the polygon
  return intersection_count % 2 !== 0;
}


function assignTerritories() {
  const rowRange = getActiveRows();
  if (rowRange.getSheet().getName() !== "Addresses") {
    throw new Error("Must be on 'Addresses' sheet");
  }

  const territories = getTerritories();
  const territoryIds = [];
  const territoryNumbers = [];
  const territoryCategories = [];

  for (const row of rowRange.getValues()) {
    const lat = row[addressCols.indexOf("Latitude")];
    const lng = row[addressCols.indexOf("Longitude")];

    territoryIds.push(null);
    territoryNumbers.push(null);
    territoryCategories.push(null);

    if (lat === "" || lng === "") {
      continue;
    }

    for (const territory of territories) {
      if (pointInPolygon(lat, lng, territory.boundary)) {
        territoryIds.splice(-1, 1, [territory.id]);
        territoryNumbers.splice(-1, 1, [territory.number]);
        territoryCategories.splice(-1, 1, [territory.category]);
        break;
      }
    }
  }

  setColumnOfRange(rowRange, addressCols.indexOf("TerritoryID"), territoryIds);
  setColumnOfRange(rowRange, addressCols.indexOf("TerritoryNumber"), territoryNumbers);
  setColumnOfRange(rowRange, addressCols.indexOf("Category"), territoryCategories);
}
