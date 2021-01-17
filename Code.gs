// This is "Sheet1" by default. Keep it in sync after any renames.
var SHEET_NAME = 'Video Stats';

// This is the named range containing all video IDs.
var VIDEO_ID_RANGE_NAME = 'IDs';

// Update these values after adding/removing columns.
var Column = {
  VIEWS: 'C',
  LIKES: 'D',
  DISLIKES: 'E',
  COMMENTS: 'F',
  DURATION: 'G'
};

// Adds a "YouTube" context menu to manually update stats.
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var entries = [{name: "Update Stats", functionName: "updateStats"}];

  spreadsheet.addMenu("YouTube", entries);
};

function updateStats() {
  var spreadsheet = SpreadsheetApp.getActive();
  var videoIds = getVideoIds();
  var stats = getStats(videoIds.join(','));
  writeStats(stats);
}

// Gets all video IDs from the range and ignores empty values.
function getVideoIds() {
  var spreadsheet = SpreadsheetApp.getActive();
  var range = spreadsheet.getRangeByName(VIDEO_ID_RANGE_NAME);
  var values = range.getValues();
  var videoIds = [];
  for (var i = 0; i < values.length; i++) {
    var value = values[i][0];
    if (!value) {
      return videoIds;
    }
    videoIds.push(value);
  }
  return videoIds;
}

// Queries the YouTube API to get stats for all videos.
function getStats(videoIds) {
  return YouTube.Videos.list('contentDetails,statistics', {'id': videoIds}).items;
}

// Converts the API results to cells in the sheet.
function writeStats(stats) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  var durationPattern = new RegExp(/PT((\d+)M)?(\d+)S/);
  for (var i = 0; i < stats.length; i++) {
    var cell = sheet.setActiveCell(Column.VIEWS + (2+i));
    cell.setValue(stats[i].statistics.viewCount);
    cell = sheet.setActiveCell(Column.LIKES + (2+i));
    cell.setValue(stats[i].statistics.likeCount);
    cell = sheet.setActiveCell(Column.DISLIKES + (2+i));
    cell.setValue(stats[i].statistics.dislikeCount);
    cell = sheet.setActiveCell(Column.COMMENTS + (2+i));
    cell.setValue(stats[i].statistics.commentCount);
    cell = sheet.setActiveCell(Column.DURATION + (2+i));
    var duration = stats[i].contentDetails.duration;
    var result = durationPattern.exec(duration);
    var min = result && result[2] || '00';
    var sec = result && result[3] || '00';
    cell.setValue('00:' + min + ':' + sec);
  }
}