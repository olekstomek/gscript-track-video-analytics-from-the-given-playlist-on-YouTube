// This is "Sheet1" by default. Keep it in sync after any renames.
var SHEET_NAME = 'Video Stats'

// This is the named range containing all video IDs.
var VIDEO_ID_RANGE_NAME = 'IDs'

// Update these values after adding/removing columns.
var Column = {
  TITLE: 'A',
  PUBLISHEDAT: 'B',
  VIEWS: 'C',
  LIKES: 'D',
  DISLIKES: 'E',
  COMMENTS: 'F',
  DURATION: 'G'
}

// Regex for standrat time iso8601 in YouTube API, eg: PT1H36M26S.
var iso8601DurationRegex = /(-)?PT(?:([.,\d]+)H)?(?:([.,\d]+)M)?(?:([.,\d]+)S)?/

// Adds a "YouTube" context menu to manually update stats.
function onOpen() {
  var entries = [{name: "Add URL of playlist", functionName: "updateStats"}, 
                 {name: "Refresh data", functionName: "refreshData"},
                 {name: "Clear data", functionName: "clearData"}]
  
  SpreadsheetApp.getActive().addMenu("YouTube stats", entries)
}

// Runs all the necessary methods to collect the results.
function updateStats() {
  var videosFromPlaylist = retrieveVideosFromPlaylist(null)
  var videoIds = getVideoIds(videosFromPlaylist)
  var stats = getStats(videoIds.join(','))

  writeStats(stats)
}

// Updates the last checked playlist. It is also used by the trigger that updates the data of the playlist videos automatically every specified period of time.
function refreshData() {
  var videosFromPlaylist = retrieveVideosFromPlaylist(SpreadsheetApp.getActive().getActiveSheet().getRange("D1").getRichTextValue().getLinkUrl())
  var videoIds = getVideoIds(videosFromPlaylist)
  var stats = getStats(videoIds.join(','))
  
  writeStats(stats)
}

// Gets all video IDs from the range.
function getVideoIds(videosFromPlaylist) {
  var videoIds = []
  
  for (var j = 0; j < videosFromPlaylist.length; j++) {
    var playlistItem = videosFromPlaylist[j]
    videoIds.push(playlistItem.snippet.resourceId.videoId)
  }
  
  return videoIds
}

// Shows a window with an input to the url for a playlist on YouTube. Extracts the playlist id from the given url. 
// Returns the list of movies in the specified playlist.
function retrieveVideosFromPlaylist(playlistToRefresh) {
  var response
  if (playlistToRefresh !== null) {
    response = playlistToRefresh
  } else {
    response = SpreadsheetApp.getUi().prompt('Enter url to playlist with videos on YouTube').getResponseText()
    showGivenPlaylist(response)
  }
  
  var splitResponse = response.split("list=")
  var playlistId = splitResponse[1]
  var nextPageToken = ''
  var listOfVideosOnPlaylist = []
  
  // This loop retrieves a set of playlist items and checks the nextPageToken in the
  // response to determine whether the list contains additional items. It repeats that process
  // until it has retrieved all of the items in the list.
  while (nextPageToken != null) {
    var playlistResponse = YouTube.PlaylistItems.list('snippet', {
      playlistId: playlistId,
      maxResults: 50,
      pageToken: nextPageToken
    })
    
    for (var j = 0; j < playlistResponse.items.length; j++) {
      var playlistItem = playlistResponse.items[j]
      listOfVideosOnPlaylist.push(playlistItem)
    }
    nextPageToken = playlistResponse.nextPageToken
  }
  
  return listOfVideosOnPlaylist
}

// Queries the YouTube API to get stats for all videos.
function getStats(videoIds) {
  return YouTube.Videos.list('contentDetails,statistics,snippet', {'id': videoIds}).items
}

// Converts the API results to cells in the sheet. Updates the time of the last update of data in the file.
function writeStats(stats) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME)
  
  for (var i = 0; i < stats.length; i++) {
    var cell = sheet.setActiveCell(Column.TITLE + (3+i))
    cell.setValue('=HYPERLINK("https://www.youtube.com/watch?v=' + stats[i].id + '";"' + stats[i].snippet.title + '")')
    cell = sheet.setActiveCell(Column.PUBLISHEDAT + (3+i))
    cell.setValue(setFormatDateAndTime(new Date(stats[i].snippet.publishedAt)))
    cell = sheet.setActiveCell(Column.VIEWS + (3+i))
    cell.setValue(stats[i].statistics.viewCount)
    cell = sheet.setActiveCell(Column.LIKES + (3+i))
    cell.setValue(stats[i].statistics.likeCount)
    cell = sheet.setActiveCell(Column.DISLIKES + (3+i))
    cell.setValue(stats[i].statistics.dislikeCount)
    cell = sheet.setActiveCell(Column.COMMENTS + (3+i))
    cell.setValue(stats[i].statistics.commentCount)
    cell = sheet.setActiveCell(Column.DURATION + (3+i))
    cell.setValue(parseISO8601Duration(stats[i].contentDetails.duration))
  }
  
  lastUpdateInformation()
}

// Parse time from ISO 8601 to HH:mm:ss.
function parseISO8601Duration(iso8601Duration) {
    var matches = iso8601Duration.match(iso8601DurationRegex)

    var value = {
        hours: matches[2] === undefined ? 0 : matches[2],
        minutes: matches[3] === undefined ? 0 : matches[3],
        seconds: matches[4] === undefined ? 0 : matches[4]
    }

    return value.hours + ":" + value.minutes + ":" + value.seconds
}

// Sets the last date when data was updated.
function lastUpdateInformation() {
  SpreadsheetApp.getActive().getSheetByName(SHEET_NAME).getRange("B1").setValue(setFormatDateAndTime(new Date()))
}

// Sets format for date and time
function setFormatDateAndTime(dateAndTime) {
  return Utilities.formatDate(dateAndTime, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd-MMM-yyyy | HH:mm:ss")
}

// Enters the given link to the playlist into the cell.
function showGivenPlaylist(response) {
  SpreadsheetApp.getActive().getSheetByName(SHEET_NAME).getRange("D1").setValue('=HYPERLINK("' + response + '";"playlist")')
}

// Clears downloaded data when requested by the user.
function clearData() {
  var activeSheet = SpreadsheetApp.getActive().getActiveSheet()
  
  activeSheet.getRange("A3:G").clearContent()
  activeSheet.getRange("B1").clearContent()
  activeSheet.getRange("D1").clearContent()
}