function getExchangeRate(baseCurrency, targetCurrency) {

  var apiKey = "insert_here";
  var exchangeRate;

  var response = UrlFetchApp.fetch("https://openexchangerates.org/api/latest.json?app_id=" + apiKey);
  var data = JSON.parse(response.getContentText());
  exchangeRate = data.rates[targetCurrency] / data.rates[baseCurrency];

  return exchangeRate;
  
}

