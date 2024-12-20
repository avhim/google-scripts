const AD_ACCOUNT_ID = 'AD_ACCOUNT_ID'

// ad, adset, campaign, account
const LEVEL = 'account'

// https://developers.facebook.com/docs/marketing-api/insights/parameters#fields
const FIELDS = 'campaign_name,clicks,cpc,impressions,spend,date_start,date_stop'

// https://developers.facebook.com/docs/marketing-api/insights/parameters#param
const DATE_RANGE = 'last_month'

// user access token linked to a Facebook app
const ACCESS_TOKEN = 'ACCESS_TOKEN'

// https://developers.facebook.com/docs/marketing-api/insights/parameters#param
const FILTERING = ''

const SPREADSHEET_ID = 'SPREADSHEET_ID'
const TAB_NAME = 'fb_cost'

function getFacebookSpendLastMonth() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet_name = spreadsheet.getSheetByName(TAB_NAME);

  const facebook_url = `https://graph.facebook.com/v21.0/act_${AD_ACCOUNT_ID}/insights?level=${LEVEL}&fields=${FIELDS}&date_preset=${DATE_RANGE}&access_token=${ACCESS_TOKEN}`;


  const encodedFacebookUrl = encodeURI(facebook_url)
  const options = {
    method: 'get',
  };
  
  // Fetches & parses the URL 
  const fetchRequest = UrlFetchApp.fetch(encodedFacebookUrl, options);
  const results = JSON.parse(fetchRequest.getContentText());

  // Caches the report run ID
    const rows = results.data.map(item => [
    item.campaign_name || 'N/A',
    item.clicks || '0',
    item.impressions || '0',
    item.cpc || '0',
    item.spend || '0',
    item.date_start || 'N/A',
    item.date_stop || 'N/A',
  ]);

    if (rows.length) {
    sheet_name.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  } else {
    Logger.log('No data found.');
  }
}
