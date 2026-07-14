const API_KEY = 'v6GFvkweNo7DK7yD3ylIZ9w52aKBU0eJ7wLXkSR3';
const COUNTRY = 'KR';
const REGION = 'kic';
const BASE_URL = `https://api-${REGION}.lgthinq.com`;
const SHEET_NAME = 'home_environment';
const WEATHER_LOCATIONS = [
  { key: 'busan', name: '부산', latitude: 35.1796, longitude: 129.0756 },
  { key: 'ulsan', name: '울산', latitude: 35.5384, longitude: 129.3114 }
];

function headers_() {
  const props = PropertiesService.getScriptProperties();

  return {
    Authorization: `Bearer ${props.getProperty('LG_THINQ_PAT')}`,
    'x-country': COUNTRY,
    'x-message-id': Utilities.getUuid(),
    'x-client-id': props.getProperty('LG_THINQ_CLIENT_ID') || Utilities.getUuid(),
    'x-api-key': API_KEY,
    'x-service-phase': 'OP',
  };
}

function apiGet_(path) {
  const res = UrlFetchApp.fetch(`${BASE_URL}/${path}`, {
    method: 'get',
    headers: headers_(),
    muteHttpExceptions: true,
  });

  const text = res.getContentText();
  const json = JSON.parse(text);

  if (res.getResponseCode() >= 300) {
    throw new Error(text);
  }

  return json.response || json;
}

function listLgDevices() {
  const devices = apiGet_('devices');
  Logger.log(JSON.stringify(devices, null, 2));
}

function logHomeEnvironment() {
  const props = PropertiesService.getScriptProperties();

  const acDeviceId = props.getProperty('LG_AC_DEVICE_ID');
  const dehumidifierDeviceId = props.getProperty('LG_DEHUMIDIFIER_DEVICE_ID');

  const acState = apiGet_(`devices/${acDeviceId}/state`);
  const dehumidifierState = apiGet_(`devices/${dehumidifierDeviceId}/state`);

  const acOperationMode = findFirstValue_(acState, [
    'airConOperationMode',
    'airconditionerOperationMode',
    'operationMode',
    'power'
  ]);

  const acCurrentTemperature = findFirstValue_(acState, [
    'currentTemperature',
    'currentTemp',
    'roomTemperature',
    'indoorTemperature',
    'airTemperature'
  ]);

  const dehumidifierOperationMode = findValue_(dehumidifierState, 'dehumidifierOperationMode');
  const rawDehumidifierHumidity = findValue_(dehumidifierState, 'currentHumidity');

  const dehumidifierHumidity =
    dehumidifierOperationMode === 'POWER_OFF' || rawDehumidifierHumidity === 0
      ? ''
      : rawDehumidifierHumidity;

  const dehumidifierTargetHumidity = findValue_(dehumidifierState, 'targetHumidity');
  const dehumidifierJobMode = findValue_(dehumidifierState, 'currentJobMode');
  const dehumidifierWindStrength = findValue_(dehumidifierState, 'windStrength');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'timestamp',
      'ac_current_temperature',
      'ac_operation_mode',
      'dehumidifier_current_humidity',
      'dehumidifier_target_humidity',
      'dehumidifier_operation_mode',
      'dehumidifier_job_mode',
      'dehumidifier_wind_strength',
      'ac_raw_json',
      'dehumidifier_raw_json'
    ]);
  }

  sheet.appendRow([
    new Date(),
    acCurrentTemperature,
    acOperationMode,
    dehumidifierHumidity,
    dehumidifierTargetHumidity,
    dehumidifierOperationMode,
    dehumidifierJobMode,
    dehumidifierWindStrength,
    JSON.stringify(acState),
    JSON.stringify(dehumidifierState)
  ]);
}

function findValue_(obj, key) {
  if (!obj || typeof obj !== 'object') return '';

  if (Object.prototype.hasOwnProperty.call(obj, key)) {
    return obj[key];
  }

  for (const k in obj) {
    const found = findValue_(obj[k], key);
    if (found !== '') return found;
  }

  return '';
}

function findFirstValue_(obj, keys) {
  for (const key of keys) {
    const value = findValue_(obj, key);
    if (value !== '') return value;
  }

  return '';
}

function setupFiveMinuteTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'logHomeEnvironment')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('logHomeEnvironment')
    .timeBased()
    .everyMinutes(5)
    .create();
}

const WEBAPP_REFRESH_SECONDS = 300;
const WEBAPP_TIMEZONE = 'Asia/Seoul';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('우리집 온습도 대시보드')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData(options) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('home_environment');
  if (!sheet) {
    throw new Error('home_environment 시트를 찾을 수 없습니다.');
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return {
      generatedAt: formatDateTime_(new Date()),
      refreshSeconds: WEBAPP_REFRESH_SECONDS,
      regionalWeathers: getRegionalWeathers_(),
      latest: null,
      recentHour: [],
      recentDay: [],
      selectedRange: [],
      recentRows: [],
      dailyRunTimes: []
    };
  }

  const headers = values[0];
  const rows = values.slice(1)
    .map(row => normalizeEnvironmentRow_(headers, row))
    .filter(row => row.timestampDate)
    .sort((a, b) => a.timestampDate.getTime() - b.timestampDate.getTime());

  const latest = rows.length ? rows[rows.length - 1] : null;
  const now = new Date();
  const range = normalizeDashboardRange_(options, now);
  const oneHourAgo = new Date(now.getTime() - 60 * 60 * 1000);
  const oneDayAgo = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const selectedRows = rows.filter(row =>
    row.timestampDate >= range.start && row.timestampDate <= range.end
  );

  return {
    generatedAt: formatDateTime_(now),
    refreshSeconds: WEBAPP_REFRESH_SECONDS,
    rangeLabel: `${formatDateTime_(range.start)} - ${formatDateTime_(range.end)}`,
    regionalWeathers: getRegionalWeathers_(),
    latest: latest ? toClientRow_(latest) : null,
    recentHour: rows.filter(row => row.timestampDate >= oneHourAgo).map(toClientRow_),
    recentDay: rows.filter(row => row.timestampDate >= oneDayAgo).map(toClientRow_),
    selectedRange: selectedRows.map(toClientRow_),
    dailyRunTimes: buildDailyRunTimes_(selectedRows),
    recentRows: rows.slice(-12).map(toClientRow_)
  };
}

function getRegionalWeathers_() {
  return WEATHER_LOCATIONS.map(getWeatherForLocation_);
}

function getWeatherForLocation_(location) {
  const url = [
    'https://api.open-meteo.com/v1/forecast',
    `?latitude=${location.latitude}`,
    `&longitude=${location.longitude}`,
    '&current=temperature_2m,relative_humidity_2m',
    `&timezone=${encodeURIComponent(WEBAPP_TIMEZONE)}`,
    '&forecast_days=1'
  ].join('');

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });

    if (response.getResponseCode() >= 300) {
      return {
        key: location.key,
        location: location.name,
        available: false,
        error: `${location.name} 날씨를 불러오지 못했습니다.`
      };
    }

    const data = JSON.parse(response.getContentText());
    const current = data.current || {};

    return {
      key: location.key,
      available: true,
      location: location.name,
      temperature: toNumberOrNull_(current.temperature_2m),
      humidity: toNumberOrNull_(current.relative_humidity_2m),
      observedAt: current.time ? formatDateTime_(new Date(current.time)) : ''
    };
  } catch (error) {
    return {
      key: location.key,
      location: location.name,
      available: false,
      error: `${location.name} 날씨를 불러오지 못했습니다.`
    };
  }
}

function normalizeEnvironmentRow_(headers, row) {
  const record = {};
  headers.forEach((header, index) => {
    record[header] = row[index];
  });

  const timestampDate = toDate_(record.timestamp);
  const acTemperature = toNumberOrNull_(record.ac_current_temperature);
  const rawHumidity = toNumberOrNull_(record.dehumidifier_current_humidity);
  const dehumidifierOperationMode = textOrBlank_(record.dehumidifier_operation_mode);
  const dehumidifierHumidity =
    dehumidifierOperationMode === 'POWER_OFF' || rawHumidity === 0 ? null : rawHumidity;

  return {
    timestampDate,
    timestamp: timestampDate ? formatDateTime_(timestampDate) : '',
    timeLabel: timestampDate ? Utilities.formatDate(timestampDate, WEBAPP_TIMEZONE, 'M/d HH:mm') : '',
    acCurrentTemperature: acTemperature,
    acOperationMode: textOrBlank_(record.ac_operation_mode) || '확인 전',
    dehumidifierCurrentHumidity: dehumidifierHumidity,
    dehumidifierTargetHumidity: toNumberOrNull_(record.dehumidifier_target_humidity),
    dehumidifierOperationMode,
    dehumidifierJobMode: textOrBlank_(record.dehumidifier_job_mode),
    dehumidifierWindStrength: textOrBlank_(record.dehumidifier_wind_strength)
  };
}

function toClientRow_(row) {
  return {
    timestamp: row.timestamp,
    timeLabel: row.timeLabel,
    acCurrentTemperature: row.acCurrentTemperature,
    acOperationMode: row.acOperationMode,
    dehumidifierCurrentHumidity: row.dehumidifierCurrentHumidity,
    dehumidifierTargetHumidity: row.dehumidifierTargetHumidity,
    dehumidifierOperationMode: row.dehumidifierOperationMode,
    dehumidifierJobMode: row.dehumidifierJobMode,
    dehumidifierWindStrength: row.dehumidifierWindStrength
  };
}

function toDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (!value) return null;

  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function toNumberOrNull_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const number = Number(value);
  return isNaN(number) ? null : number;
}

function textOrBlank_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function formatDateTime_(date) {
  return Utilities.formatDate(date, WEBAPP_TIMEZONE, 'M/d HH:mm');
}

function normalizeDashboardRange_(options, now) {
  const end = options && options.end ? new Date(options.end) : now;
  const start = options && options.start ? new Date(options.start) : new Date(end.getTime() - 24 * 60 * 60 * 1000);

  if (isNaN(start.getTime()) || isNaN(end.getTime())) {
    throw new Error('기간 값이 올바르지 않습니다.');
  }

  if (start.getTime() > end.getTime()) {
    throw new Error('시작 시간이 종료 시간보다 늦습니다.');
  }

  return { start, end };
}

function buildDailyRunTimes_(rows) {
  if (!rows.length) return [];

  const byDay = {};
  rows.forEach((row, index) => {
    const next = rows[index + 1];
    const minutes = next
      ? Math.max(0, (next.timestampDate.getTime() - row.timestampDate.getTime()) / 60000)
      : inferLastIntervalMinutes_(rows);
    const day = Utilities.formatDate(row.timestampDate, WEBAPP_TIMEZONE, 'yyyy-MM-dd');

    if (!byDay[day]) {
      byDay[day] = {
        date: day,
        acPowerOnMinutes: 0,
        dehumidifierPowerOnMinutes: 0
      };
    }

    if (row.acOperationMode === 'POWER_ON') {
      byDay[day].acPowerOnMinutes += minutes;
    }

    if (row.dehumidifierOperationMode === 'POWER_ON') {
      byDay[day].dehumidifierPowerOnMinutes += minutes;
    }
  });

  return Object.keys(byDay).sort().map(day => {
    const summary = byDay[day];
    return {
      date: summary.date,
      acPowerOnMinutes: Math.round(summary.acPowerOnMinutes),
      acPowerOnHours: roundHours_(summary.acPowerOnMinutes),
      dehumidifierPowerOnMinutes: Math.round(summary.dehumidifierPowerOnMinutes),
      dehumidifierPowerOnHours: roundHours_(summary.dehumidifierPowerOnMinutes)
    };
  });
}

function inferLastIntervalMinutes_(rows) {
  if (rows.length < 2) return 5;

  const previous = rows[rows.length - 2].timestampDate;
  const last = rows[rows.length - 1].timestampDate;
  const minutes = (last.getTime() - previous.getTime()) / 60000;
  return minutes > 0 && minutes <= 60 ? minutes : 5;
}

function roundHours_(minutes) {
  return Math.round((minutes / 60) * 10) / 10;
}
