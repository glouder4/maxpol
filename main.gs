// ВАЖНО: Вставьте ID вашей Google Таблицы сюда
// Чтобы получить ID: откройте вашу таблицу и посмотрите в URL:
// https://docs.google.com/spreadsheets/d/ИД_ТАБЛИЦЫ/edit
const SHEET_ID = '';

// Константы для работы с таблицей
const CONTACTS_SHEET_NAME = 'Логи запросов';
const DEBUG_LOG_SHEET_NAME = 'Логи отладки';
const CONTACTS_COLUMNS = ['Время', 'ID контакта', 'Готовность ехать к заказчику', 'Готовность работать с менеджером', 'Срочность', 'Деньги'];
const DEAL_SELECT_FIELDS = [
  'ID',
  'TITLE',
  'STAGE_ID',
  'CATEGORY_ID',
  'ASSIGNED_BY_ID',
  'STATUS_ID',
  'OPPORTUNITY',
  'CURRENCY_ID',
  'DATE_CREATE',
  'CLOSEDATE',
  'UF_CRM_1761643310581',
  'UF_CRM_1760955888661',
  'UF_CRM_1760953632886',
  'UF_CRM_1760953604180',
  'UF_CRM_1760953527241',
  'UF_CRM_1760954328306'
];

// ID поля-контейнера с данными формы (содержит JSON строку)
const FORM_DATA_FIELD_ID = 'UF_CRM_1761752668447';
const DEAL_UPDATE_WEBHOOK_URL = 'https://makspol.bitrix24.ru/rest/';

// Поля для извлечения из Bitrix24 (названия в custom JSON)
const BITRIX_FIELDS = {
  GOTOVNOST_EHAT_NA_VSTRECHU: 'Gotovnost_ehat_na_vstrechu_k_zastroyshchiku',
  GOTOVNOST_RABOTAT_S_MENEDZHEROM: 'Gotovnost_rabotat_s_menedzherom',
  SROCHNOST: 'Srochnost',
  DENGI: 'Dengi'
};

// Ассоциации полей с ID в Bitrix24 (для обновления контакта)
const BITRIX_FIELD_IDS = {
  UF_CRM_1764582569038: 'gotovnostEhat', // Готовность ехать на встречу с застройщиком - список
  UF_CRM_1764582544287: 'gotovnostRabotat', // Готовность работать с менеджером - список
  UF_CRM_1761752594935: 'srochnost', // Срочность - список
  UF_CRM_1761752624499: 'dengi' // Деньги - список
};

// Webhook для обновления контакта
const UPDATE_WEBHOOK_URL = 'https://makspol.bitrix24.ru/rest/';

// Маппинг текстовых значений на ID списковых полей
const LIST_VALUE_MAPPING = {
  // Срочность (UF_CRM_1761752594935)
  srochnost: {
    'ГОТОВ КУПИТЬ СЕЙЧАС': 15758,
    'В ЭТОМ МЕСЯЦЕ': 15760,
    'В ТЕЧЕНИЕ 2 Х МЕСЯЦЕВ': 15762,
    'НЕТ СРОЧНОСТИ': 15764,
    'НЕ СРОЧНО': 15764,
    'НЕ УТОЧНИЛ': 16602
  },
  // Деньги (UF_CRM_1761752624499)
  dengi: {
    'СВОИ': 15766,
    'КРЕДИТНЫЕ': 15768,
    'РАССРОЧКА': 15770,
    'НЕ УТОЧНИЛ': 16598
  },
  // Готовность работать с менеджером (UF_CRM_1764582544287)
  gotovnostRabotat: {
    'ДА': 16606,
    'НЕТ': 16608,
    'НЕ УТОЧНИЛ': 16610
  },
  // Готовность ехать на встречу с застройщиком (UF_CRM_1764582569038)
  gotovnostEhat: {
    'ДА': 16612,
    'НЕТ': 16614,
    'НЕ УТОЧНИЛ': 16616
  }
};

const DEAL_LIST_VALUE_MAPPING = {
  srochnost: {
    'ГОТОВ КУПИТЬ СЕЙЧАС': 14826,
    'В ЭТОМ МЕСЯЦЕ': 14828,
    'В ТЕЧЕНИЕ 2 Х МЕСЯЦЕВ': 14830,
    'НЕТ СРОЧНОСТИ': 14832,
    'НЕ СРОЧНО': 14832,
    'НЕ УТОЧНИЛ': 16604
  },
  dengi: {
    'СВОИ': 14822,
    'КРЕДИТНЫЕ': 14824,
    'РАССРОЧКА': 14898,
    'НЕ УТОЧНИЛ': 16600
  },
  // Готовность работать с менеджером (UF_CRM_1764582743319)
  gotovnostRabotat: {
    'ДА': 16618,
    'НЕТ': 16620,
    'НЕ УТОЧНИЛ': 16622
  },
  // Готовность ехать на встречу с застройщиком (UF_CRM_1764582768429)
  gotovnostEhat: {
    'ДА': 16624,
    'НЕТ': 16626,
    'НЕ УТОЧНИЛ': 16628
  }
};

/**
 * Логирует сообщение в таблицу отладки
 */
function logToSheet(message) {
  console.log(message);
  /*
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    let sheet = spreadsheet.getSheetByName(DEBUG_LOG_SHEET_NAME);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(DEBUG_LOG_SHEET_NAME);
      sheet.appendRow(['Время', 'Сообщение']);
      formatSheetHeader(sheet, 1, 1, 2);
    }

    sheet.appendRow([new Date().toISOString(), message]);
    sheet.autoResizeColumns(1, 2);
  } catch (error) {
    // Игнорируем ошибки логирования
  }
  */
}

/**
 * Форматирует заголовок таблицы
 */
function formatSheetHeader(sheet, startRow, startCol, numCols) {
  const headerRange = sheet.getRange(startRow, startCol, 1, numCols);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
}

/**
 * Логирует контрольную точку с измерением времени
 */
function logCheckpoint(label, startedAt = null) {
  const now = new Date();
  if (startedAt instanceof Date) {
    const diffMs = now.getTime() - startedAt.getTime();
    const diffSeconds = Math.round(diffMs / 100) / 10;
    logToSheet(`${label} (+${diffSeconds} сек)`);
  } else {
    logToSheet(label);
  }
  return now;
}

/**
 * Нормализует строковое значение для сопоставления
 */
function normalizeFieldValue(value) {
  if (!value || typeof value !== 'string') {
    return '';
  }
  return value.trim().toUpperCase();
}

/**
 * Определяет, установлен ли чекбокс в Bitrix24
 */
function isBitrixCheckboxChecked(value) {
  if (value === true) {
    return true;
  }
  if (value === false || value === 0 || value === null || value === undefined) {
    return false;
  }

  const normalized = normalizeFieldValue(String(value));

  if (!normalized) {
    return false;
  }

  const positiveValues = ['Y', 'YES', 'TRUE', '1'];
  const negativeValues = ['N', 'NO', 'FALSE', '0',''];

  if (positiveValues.includes(normalized)) {
    return true;
  }

  if (negativeValues.includes(normalized)) {
    return false;
  }

  return normalized.length > 0;
}

/**
 * Извлекает ID контакта из строки вида "CONTACT_139204"
 */
function extractContactIdFromString(value) {
  if (!value || typeof value !== 'string') return '';
  if (value.startsWith('CONTACT_')) {
    return value.replace('CONTACT_', '');
  }
  return '';
}

/**
 * Извлекает ID контакта из POST запроса
 */
function extractContactId(e) {
  let contactId = '';
  
  // Логируем все параметры для отладки
  logToSheet('Все параметры запроса: ' + JSON.stringify(e.parameter));
  if (e.postData && e.postData.contents) {
    logToSheet('Содержимое POST body: ' + e.postData.contents);
  }
  
  // Проверяем все возможные варианты document_id параметров
  const possibleParams = [
    'document_id[1]',
    'document_id[2]',
    'document_id[3]',
    'document_id[0]',
    'id',
    'contact_id',
    'CONTACT_ID'
  ];
  
  // Сначала проверяем параметры запроса
  if (e.parameter) {
    for (let paramName of possibleParams) {
      if (e.parameter[paramName]) {
        const value = e.parameter[paramName];
        logToSheet(`Найден параметр ${paramName}: ${value}`);
        contactId = extractContactIdFromString(value);
        if (contactId) {
          logToSheet(`Извлечен contactId из ${paramName}: ${contactId}`);
          return contactId;
        }
      }
    }
  }
  
  // Если не нашли в параметрах, проверяем POST body
  if (!contactId && e.postData && e.postData.contents) {
    try {
      const params = new URLSearchParams(e.postData.contents);
      for (let paramName of possibleParams) {
        const value = params.get(paramName);
        if (value) {
          logToSheet(`Найден параметр в body ${paramName}: ${value}`);
          contactId = extractContactIdFromString(value);
          if (contactId) {
            logToSheet(`Извлечен contactId из body ${paramName}: ${contactId}`);
            return contactId;
          }
        }
      }
      
      // Также пробуем распарсить как JSON
      try {
        const jsonData = JSON.parse(e.postData.contents);
        logToSheet('POST body как JSON: ' + JSON.stringify(jsonData));
        // Проверяем возможные JSON поля
        if (jsonData.id) {
          contactId = String(jsonData.id);
          logToSheet(`Извлечен contactId из JSON.id: ${contactId}`);
          return contactId;
        }
        if (jsonData.contact_id) {
          contactId = String(jsonData.contact_id);
          logToSheet(`Извлечен contactId из JSON.contact_id: ${contactId}`);
          return contactId;
        }
      } catch (jsonError) {
        logToSheet('POST body не является JSON: ' + jsonError.toString());
      }
    } catch (parseError) {
      logToSheet('Ошибка парсинга POST body: ' + parseError.toString());
    }
  }
  
  logToSheet('Не удалось извлечь contactId из запроса');
  return contactId;
}

function doPost(e) {
  try {
    // Собираем информацию о запросе
    const timestamp = new Date().toISOString();
    
    let checkpoint = logCheckpoint('POST запрос получен');
    
    // Извлекаем contactId из POST данных
    let contactId = extractContactId(e);
    
    checkpoint = logCheckpoint('Извлечен contactId', checkpoint);
    logToSheet('Значение contactId: ' + contactId);
    
    // Получаем данные контакта из Bitrix24
    let contactFields = getContactFields(contactId);
    checkpoint = logCheckpoint('Получены данные контакта из Bitrix24', checkpoint);

    // Получаем сделки, связанные с контактом
    const contactDeals = getContactDeals(contactId);
    checkpoint = logCheckpoint(`Получено сделок контакта: ${contactDeals.length}`, checkpoint);
    
    // Записываем в Google Sheets
    writeToSheet([timestamp, contactId, contactFields.gotovnostEhat, contactFields.gotovnostRabotat, contactFields.srochnost, contactFields.dengi]);
    checkpoint = logCheckpoint('Данные записаны в Google Sheets', checkpoint);
    
    // Обновляем контакт в Bitrix24, если есть определенные поля
    updateContactInBitrix24(contactId, contactFields);
    checkpoint = logCheckpoint('Завершено обновление контакта в Bitrix24', checkpoint);

    // Обновляем связанные сделки на основании данных контакта
    updateDealsWithContactData(contactDeals, contactFields);
    checkpoint = logCheckpoint('Завершено обновление сделок контакта', checkpoint);
    
    // Возвращаем данные запроса в JSON формате
    return ContentService.createTextOutput(
      JSON.stringify({
        status: 'success',
        message: 'Данные успешно получены и записаны в таблицу',
        contactId: contactId,
        deals: contactDeals,
        timestamp: timestamp
      }, null, 2)
    ).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    // Если произошла ошибка, возвращаем её описание
    return ContentService.createTextOutput(
      JSON.stringify({
        status: 'error',
        message: 'Ошибка при обработке запроса',
        error: error.toString()
      }, null, 2)
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Получает данные контакта из Bitrix24 и извлекает нужные поля
 */
function getContactFields(contactId) {
  const fields = {
    gotovnostEhat: '',
    gotovnostRabotat: '',
    srochnost: '',
    dengi: '',
    rawFormData: ''
  };
  let checkpoint = logCheckpoint(`Начало получения данных для контакта ${contactId}`);
  try {
    const url = `https://makspol.bitrix24.ru/rest/70/?id=${contactId}`;
    logToSheet('URL запроса: ' + url);
    
    const response = UrlFetchApp.fetch(url);
    checkpoint = logCheckpoint('Получен ответ от Bitrix24 (fetch)', checkpoint);
    const content = response.getContentText();
    logToSheet('Сырой ответ от Bitrix24: ' + content);
    
    // Парсим JSON и извлекаем только нужное поле
    const data = JSON.parse(content);
    checkpoint = logCheckpoint('Ответ Bitrix24 распарсен', checkpoint);
    logToSheet('Распарсенный ответ: ' + JSON.stringify(data));
    
    // Извлекаем поле-контейнер с данными формы
    const fieldString = data.result && data.result[FORM_DATA_FIELD_ID] ? data.result[FORM_DATA_FIELD_ID] : '';
    fields.rawFormData = fieldString || '';
    logToSheet(`Поле ${FORM_DATA_FIELD_ID}: ` + fieldString);
    
    // Если поле содержит строку с JSON, распарсим её
    if (fieldString && typeof fieldString === 'string') {
      try {
        const parsed = JSON.parse(fieldString);
        checkpoint = logCheckpoint('Распарсен JSON из пользовательского поля', checkpoint);
        
        // Теперь извлекаем поле "custom" (это объект, а не строка)
        if (parsed.custom && typeof parsed.custom === 'object') {
          const customData = parsed.custom;
          
          // Извлекаем нужные поля используя константы
          fields.gotovnostEhat = customData[BITRIX_FIELDS.GOTOVNOST_EHAT_NA_VSTRECHU] || '';
          fields.gotovnostRabotat = customData[BITRIX_FIELDS.GOTOVNOST_RABOTAT_S_MENEDZHEROM] || '';
          fields.srochnost = customData[BITRIX_FIELDS.SROCHNOST] || '';
          fields.dengi = customData[BITRIX_FIELDS.DENGI] || '';
          
          logToSheet(`Извлеченные поля: gotovnostEhat=${fields.gotovnostEhat}, gotovnostRabotat=${fields.gotovnostRabotat}, srochnost=${fields.srochnost}, dengi=${fields.dengi}`);
        } else {
          logToSheet('Поле custom не найдено или не объект. Тип: ' + typeof parsed.custom);
        }
      } catch (e) {
        logToSheet('Ошибка парсинга первого JSON: ' + e.toString());
      }
    } else {
      logToSheet('fieldString пустой или не строка');
    }
    
  } catch (error) {
    logToSheet('Ошибка при получении данных контакта: ' + error.toString());
  }
  
  logCheckpoint(`Завершено получение данных контакта ${contactId}`, checkpoint);
  logToSheet('Возвращаемые поля: ' + JSON.stringify(fields));
  return fields;
}

/**
 * Получает список сделок, связанных с контактом
 */
function getContactDeals(contactId) {
  if (!contactId) {
    logToSheet('getContactDeals: не указан contactId');
    return [];
  }

  try {
    logToSheet(`Запрос сделок для контакта ${contactId}`);
    const url = 'https://makspol.bitrix24.ru/rest/';
    const payload = {
      filter: {
        CONTACT_ID: contactId
      },
      select: DEAL_SELECT_FIELDS
    };

    logToSheet(`Payload crm.deal.list: ${JSON.stringify(payload)}`);

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const content = response.getContentText();
    logToSheet(`Ответ crm.deal.list: ${content}`);

    const data = JSON.parse(content);
    if (!data.result || !Array.isArray(data.result)) {
      logToSheet('crm.deal.list: результат не содержит массива');
      return [];
    }

    return data.result;
  } catch (error) {
    logToSheet(`Ошибка при получении сделок контакта: ${error.toString()}`);
    return [];
  }
}

/**
 * Обновляет сделки данными из контакта
 */
function updateDealsWithContactData(deals, contactFields) {
  if (!Array.isArray(deals) || deals.length === 0) {
    logToSheet('updateDealsWithContactData: сделки для обновления отсутствуют');
    return;
  }

  const normalizedValues = {
    gotovnostRabotat: normalizeFieldValue(contactFields.gotovnostRabotat),
    gotovnostEhat: normalizeFieldValue(contactFields.gotovnostEhat),
    srochnost: normalizeFieldValue(contactFields.srochnost),
    dengi: normalizeFieldValue(contactFields.dengi)
  };

  deals.forEach(deal => {
    const dealId = deal && deal.ID ? String(deal.ID) : '';
    if (!dealId) {
      logToSheet('updateDealsWithContactData: пропущена сделка без ID');
      return;
    }

    const alreadyProcessed = isBitrixCheckboxChecked(deal.UF_CRM_1761643310581);
    if (alreadyProcessed) {
      logToSheet(`Сделка ${dealId} пропущена: флаг обработки уже установлен`);
      return;
    }

    const updateFields = {};

    if (contactFields.rawFormData) {
      updateFields.UF_CRM_1760955888661 = contactFields.rawFormData;
    }

    // Обработка gotovnostRabotat (список)
    if (normalizedValues.gotovnostRabotat && normalizedValues.gotovnostRabotat !== 'FALSE') {
      const mappedGotovnostRabotat = DEAL_LIST_VALUE_MAPPING.gotovnostRabotat[normalizedValues.gotovnostRabotat];
      if (mappedGotovnostRabotat) {
        updateFields.UF_CRM_1764582743319 = mappedGotovnostRabotat;
      }
    }

    // Обработка gotovnostEhat (список)
    if (normalizedValues.gotovnostEhat && normalizedValues.gotovnostEhat !== 'FALSE') {
      const mappedGotovnostEhat = DEAL_LIST_VALUE_MAPPING.gotovnostEhat[normalizedValues.gotovnostEhat];
      if (mappedGotovnostEhat) {
        updateFields.UF_CRM_1764582768429 = mappedGotovnostEhat;
      }
    }

    if (normalizedValues.srochnost && normalizedValues.srochnost !== 'FALSE') {
      const mappedSrochnost = DEAL_LIST_VALUE_MAPPING.srochnost[normalizedValues.srochnost];
      if (mappedSrochnost) {
        updateFields.UF_CRM_1760954328306 = mappedSrochnost;
      }
    }

    if (normalizedValues.dengi && normalizedValues.dengi !== 'FALSE') {
      const mappedDengi = DEAL_LIST_VALUE_MAPPING.dengi[normalizedValues.dengi];
      if (mappedDengi) {
        updateFields.UF_CRM_1760953527241 = mappedDengi;
      }
    }

    updateFields.UF_CRM_1761643310581 = true;

    if (Object.keys(updateFields).length === 0) {
      logToSheet(`Сделка ${dealId}: нет данных для обновления`);
      return;
    }

    try {
      const payload = {
        id: dealId,
        fields: updateFields
      };

      logToSheet(`Обновление сделки ${dealId}: ${JSON.stringify(payload)}`);

      const response = UrlFetchApp.fetch(DEAL_UPDATE_WEBHOOK_URL, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const content = response.getContentText();
      logToSheet(`Результат обновления сделки ${dealId}: ${content}`);
    } catch (error) {
      logToSheet(`Ошибка при обновлении сделки ${dealId}: ${error.toString()}`);
    }
  });
}

/**
 * Обновляет контакт в Bitrix24 через вебхук
 */
function updateContactInBitrix24(contactId, fields) {
  try {
    logToSheet(`Обновление контакта ${contactId} в Bitrix24`);
    let checkpoint = logCheckpoint('Начало подготовки данных для обновления Bitrix24');
    
    // Подготавливаем поля для обновления (только определенные поля)
    const updateFields = {};
    const normalizedValues = {
      gotovnostRabotat: normalizeFieldValue(fields.gotovnostRabotat),
      gotovnostEhat: normalizeFieldValue(fields.gotovnostEhat),
      srochnost: normalizeFieldValue(fields.srochnost),
      dengi: normalizeFieldValue(fields.dengi)
    };
    
    // Списковые поля - конвертируем текст в ID
    // Важно: обновляем только если значение не 'false' и есть в маппинге
    // Используем ID из BITRIX_FIELD_IDS для контактов
    const gotovnostRabotatFieldId = Object.keys(BITRIX_FIELD_IDS).find(key => BITRIX_FIELD_IDS[key] === 'gotovnostRabotat');
    const gotovnostEhatFieldId = Object.keys(BITRIX_FIELD_IDS).find(key => BITRIX_FIELD_IDS[key] === 'gotovnostEhat');
    const srochnostFieldId = Object.keys(BITRIX_FIELD_IDS).find(key => BITRIX_FIELD_IDS[key] === 'srochnost');
    const dengiFieldId = Object.keys(BITRIX_FIELD_IDS).find(key => BITRIX_FIELD_IDS[key] === 'dengi');
    
    // Обработка gotovnostRabotat (список)
    if (normalizedValues.gotovnostRabotat && normalizedValues.gotovnostRabotat !== 'FALSE' && LIST_VALUE_MAPPING.gotovnostRabotat[normalizedValues.gotovnostRabotat]) {
      if (gotovnostRabotatFieldId) {
        updateFields[gotovnostRabotatFieldId] = LIST_VALUE_MAPPING.gotovnostRabotat[normalizedValues.gotovnostRabotat];
      }
    }
    
    // Обработка gotovnostEhat (список)
    if (normalizedValues.gotovnostEhat && normalizedValues.gotovnostEhat !== 'FALSE' && LIST_VALUE_MAPPING.gotovnostEhat[normalizedValues.gotovnostEhat]) {
      if (gotovnostEhatFieldId) {
        updateFields[gotovnostEhatFieldId] = LIST_VALUE_MAPPING.gotovnostEhat[normalizedValues.gotovnostEhat];
      }
    }
    
    if (normalizedValues.srochnost && normalizedValues.srochnost !== 'FALSE' && LIST_VALUE_MAPPING.srochnost[normalizedValues.srochnost]) {
      if (srochnostFieldId) {
        updateFields[srochnostFieldId] = LIST_VALUE_MAPPING.srochnost[normalizedValues.srochnost]; // Срочность - список
      }
    }
    if (normalizedValues.dengi && normalizedValues.dengi !== 'FALSE' && LIST_VALUE_MAPPING.dengi[normalizedValues.dengi]) {
      if (dengiFieldId) {
        updateFields[dengiFieldId] = LIST_VALUE_MAPPING.dengi[normalizedValues.dengi]; // Деньги - список
      }
    }
    checkpoint = logCheckpoint('Сформированы поля для обновления Bitrix24', checkpoint);
    
    // Если нет полей для обновления, выходим
    if (Object.keys(updateFields).length === 0) {
      logToSheet('Нет определенных полей для обновления');
      return null;
    }
    
    logToSheet(`Обновляем поля: ${Object.keys(updateFields).join(', ')}`);
    
    // Формируем данные для POST запроса
    const payload = {
      id: contactId,
      fields: updateFields
    };
    
    const payloadJson = JSON.stringify(payload);
    logToSheet(`URL: ${UPDATE_WEBHOOK_URL}, Payload: ${payloadJson}`);
    
    // Отправляем POST запрос с JSON телом
    const response = UrlFetchApp.fetch(UPDATE_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: payloadJson
    });
    
    const result = response.getContentText();
    logToSheet(`Результат обновления: ${result}`);
    logCheckpoint('Завершено обновление Bitrix24', checkpoint);
    
    return JSON.parse(result);
    
  } catch (error) {
    logToSheet(`Ошибка при обновлении контакта: ${error.toString()}`);
    return null;
  }
}

/**
 * Записывает данные в Google Таблицу
 */
function writeToSheet(rowData) {
  try {
    // Получаем таблицу по ID
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    let sheet = spreadsheet.getSheetByName(CONTACTS_SHEET_NAME);
    
    // Если лист не существует, создаем его
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONTACTS_SHEET_NAME);
      // Добавляем заголовки
      sheet.appendRow(CONTACTS_COLUMNS);
      // Форматируем заголовки
      formatSheetHeader(sheet, 1, 1, CONTACTS_COLUMNS.length);
    }
    
    // Добавляем новую строку с данными
    sheet.appendRow(rowData);
    
    // Автоматически расширяем ширину колонок
    sheet.autoResizeColumns(1, CONTACTS_COLUMNS.length);
    
  } catch (error) {
    logToSheet('Ошибка при записи в таблицу: ' + error.toString());
    throw error;
  }
}

/**
 * Функция для инициализации таблицы (запустите вручную один раз)
 * Создаст новый лист с заголовками
 */
function initializeSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    let sheet = spreadsheet.getSheetByName(CONTACTS_SHEET_NAME);
    
    // Если лист существует, удаляем его
    if (sheet) {
      spreadsheet.deleteSheet(sheet);
    }
    
    // Создаем новый лист
    sheet = spreadsheet.insertSheet(CONTACTS_SHEET_NAME);
    
    // Добавляем заголовки
    sheet.appendRow(CONTACTS_COLUMNS);
    
    // Форматируем заголовки
    formatSheetHeader(sheet, 1, 1, CONTACTS_COLUMNS.length);
    
    // Автоматически расширяем ширину колонок
    sheet.autoResizeColumns(1, CONTACTS_COLUMNS.length);
    
    logToSheet('Таблица успешно инициализирована!');
  } catch (error) {
    logToSheet('Ошибка при инициализации таблицы: ' + error.toString());
  }
}
