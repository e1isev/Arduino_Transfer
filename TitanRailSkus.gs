/**
 * Titan Rail SKU lookup table.
 * Each entry: [name, uuid, sku]
 * Sourced from the live sheet product catalogue.
 */
var TITAN_RAIL_SKUS = [
  ['Titan Rail Stainless Steel Joining Black Buckle', '763d7fba-697a-48c8-9e13-92280a5ac545', 'TRBKLEBK'],
  ['Titan Rail Stainless Steel White Joining Buckle', '82967c5f-f50d-4f45-bd85-06d8c2603724', 'TRBKLEWT'],
  ['Titan Rail Black Internal Corner Bracket',        'b6965291-b7cc-458c-991e-5737016e9858', 'TRBRKBK-C'],
  ['Titan Rail Black End Bracket System',             '8a164c6d-9ceb-43e0-886d-beaa684c2fba', 'TRBRKBK-E'],
  ['Titan Rail Black Front Brackets 25pk',            '1e2e1b94-9acb-41a0-bfaf-92f1dad07c68', 'TRBRKBOXBK-F'],
  ['Titan Rail Black Top Brackets 25pk',              'a27fc206-12e2-472d-a18e-8947e1700ca4', 'TRBRKBOXBK-T'],
  ['Titan Rail White Front Brackets 25pk',            '1c6755a8-1ff6-46f3-b2e1-157aa60320a7', 'TRBRKBOXWT-F'],
  ['Titan Rail White Top Brackets 25pk',              '158612d7-2ad2-4e98-8838-46e2a531810a', 'TRBRKBOXWT-T'],
  ['Titan Rail White Internal Corner Bracket',        'da1ef205-5176-4f4c-8160-27fb008943a4', 'TRBRKWT-C'],
  ['Titan Rail White End Bracket System',             '8415ce91-a8ee-45f6-b842-22b1617514cf', 'TRBRKWT-E'],
  ['Titan Rail - Black 100m',                         '8fbb7f6d-6439-4368-bc30-952d731b1954', 'TRRAILBK100m'],
  ['Titan Rail - Black 200m',                         '6a4c769c-e8d4-4a6c-8736-a3db103f0cc0', 'TRRAILBK200m'],
  ['Titan Rail - White 100m',                         '46c038ee-d181-4063-a977-5937c1af8728', 'TRRAILWT100m'],
  ['Titan Rail - White 200m',                         'a60ef9d4-6445-4be6-bc38-af12485e18a2', 'TRRAILWT200m'],
  ['Titan Rail Spinner Galvanised',                   'a7b7f0a6-7deb-4488-9de0-ee001bf9392c', 'TRSPINNER'],
  ['Titan Rail Black Line Strainer',                  '154d9d16-56cc-43e9-bd36-b3c01d183a18', 'TRSTRAINBK'],
  ['Titan Rail White Line Strainer',                  '9c2098e9-273c-435c-96bc-53c5fe0623d4', 'TRSTRAINWT'],
  ['Titan Wire 200m Black',                           '9958efc6-c58d-4709-a4f7-670dae338763', 'TRWIREBK200m'],
  ['Titan Wire 400m Black',                           'bc321845-5734-44b4-aa6e-1037395c71a1', 'TRWIREBK400m'],
  ['Titan Wire 200m White',                           'cec62459-921d-4052-8281-1c0297059d8e', 'TRWIREWT200m'],
  ['Titan Wire 400m White',                           'e8089410-a3bf-4c21-8f02-d72e6deba279', 'TRWIREWT400m'],
];

/**
 * Finds a Titan Rail product entry by SKU (case-insensitive).
 * Returns {name, uuid, sku} or null if not found.
 *
 * @param {string} sku
 * @returns {{name: string, uuid: string, sku: string}|null}
 */
function getTitanRailBySku(sku) {
  var upper = String(sku).trim().toUpperCase();
  for (var i = 0; i < TITAN_RAIL_SKUS.length; i++) {
    if (TITAN_RAIL_SKUS[i][2].toUpperCase() === upper) {
      return { name: TITAN_RAIL_SKUS[i][0], uuid: TITAN_RAIL_SKUS[i][1], sku: TITAN_RAIL_SKUS[i][2] };
    }
  }
  return null;
}

/**
 * Finds all Titan Rail product entries whose name contains the given search term
 * (case-insensitive).
 * Returns an array of {name, uuid, sku} objects.
 *
 * @param {string} term
 * @returns {Array<{name: string, uuid: string, sku: string}>}
 */
function searchTitanRailByName(term) {
  var lower = String(term).trim().toLowerCase();
  return TITAN_RAIL_SKUS
    .filter(function(row) { return row[0].toLowerCase().indexOf(lower) !== -1; })
    .map(function(row)    { return { name: row[0], uuid: row[1], sku: row[2] }; });
}
