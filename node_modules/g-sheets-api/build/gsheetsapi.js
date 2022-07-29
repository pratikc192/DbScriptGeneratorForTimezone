"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var _crossFetch = _interopRequireDefault(require("cross-fetch"));

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var gsheetsAPI = function gsheetsAPI(_ref) {
  var apiKey = _ref.apiKey,
      sheetId = _ref.sheetId,
      sheetName = _ref.sheetName,
      _ref$sheetNumber = _ref.sheetNumber,
      sheetNumber = _ref$sheetNumber === void 0 ? 1 : _ref$sheetNumber;

  try {
    var sheetNameStr = sheetName && sheetName !== '' ? encodeURIComponent(sheetName) : "Sheet".concat(sheetNumber);
    var sheetsUrl = "https://sheets.googleapis.com/v4/spreadsheets/".concat(sheetId, "/values/").concat(sheetNameStr, "?dateTimeRenderOption=FORMATTED_STRING&majorDimension=ROWS&valueRenderOption=FORMATTED_VALUE&key=").concat(apiKey);
    return (0, _crossFetch.default)(sheetsUrl).then(function (response) {
      if (!response.ok) {
        console.log('there is an error in the gsheets response');
        throw new Error('Error fetching GSheet');
      }

      return response.json();
    }).then(function (data) {
      return data;
    }).catch(function (err) {
      throw new Error('Failed to fetch from GSheets API. Check your Sheet Id and the public availability of your GSheet.');
    });
  } catch (err) {
    throw new Error("General error when fetching GSheet: ".concat(err));
  }
};

var _default = gsheetsAPI;
exports.default = _default;