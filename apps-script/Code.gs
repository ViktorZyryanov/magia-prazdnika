/**
 * «Магия праздника» — приём заявок с сайта в Google Таблицу.
 *
 * Как установить:
 * 1. Откройте нужную Google Таблицу → Расширения → Apps Script.
 * 2. Удалите весь старый код в файле Code.gs и вставьте этот файл целиком. Сохраните.
 * 3. Развернуть → Управление развёртываниями → ✎ (карандаш) у текущего развёртывания →
 *    Версия: «Новая версия» → Выполнять как: «Я» → У кого есть доступ: «Все» → Развернуть.
 *    Так адрес /exec, который уже указан на сайте, останется прежним.
 * 4. Откройте адрес /exec в браузере. Должен появиться ответ {"ok":true,...}.
 *    Если видите «Script function not found: doGet», значит развёрнута старая версия (повторите шаг 3).
 */

var SHEET_NAME   = 'Заявки';
var LOG_NAME     = 'Лог';
var NOTIFY_EMAIL = '';   // почта для уведомлений о новых заявках (можно оставить пустой)

// Порядок столбцов совпадает с прежней версией скрипта, чтобы старые строки не съехали.
// «Тариф» оставлен для совместимости: блок тарифов с сайта убран, столбец будет пустым.
var HEADERS = ['Дата заявки', 'Имя', 'Телефон', 'Дата праздника', 'Время', 'Формат',
               'Тариф', 'Стоимость', 'Ребёнок', 'Комментарий', 'Источник', 'Согласие', 'leadId', 'Район', 'Населённый пункт', 'Доплата за выезд', 'Количество детей', 'Надбавка за дату и время', 'Взрослых (корпоратив)', 'Детей (корпоратив)', 'Доплата за гостей', 'Скидка за ранний заказ', 'Стоимость до скидки', 'Базовая цена формата', 'Доплата за детей в детском формате', 'Итог числом', 'Правило тарифа', 'Промокод', 'Скидка по промокоду'];
var LEAD_ID_COL = HEADERS.indexOf('leadId') + 1;

function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (headers) {
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.setFrozenRows(1);
    } else if (name === SHEET_NAME) {
      // Добавляем новые столбцы после прежнего leadId; существующие строки не сдвигаем.
      var existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 13)).getValues()[0];
      for (var i = 13; i < headers.length; i++) {
        if (!existing[i]) sh.getRange(1, i + 1).setValue(headers[i]);
      }
    }
  }
  return sh;
}

/* Принимает данные в любом виде: JSON в теле, payload=JSON, обычная форма или GET-параметры */
function parse_(e) {
  var raw = (e && e.postData && e.postData.contents) ? String(e.postData.contents) : '';
  if (raw.charAt(0) === '{') {
    try { return JSON.parse(raw); } catch (x) {}
  }
  var p = (e && e.parameter) || {};
  if (p.payload) {
    try { return JSON.parse(p.payload); } catch (x) {}
  }
  return p;
}

function logError_(err) {
  try { sheet_(LOG_NAME, ['Дата и время', 'Ошибка']).appendRow([new Date(), String(err)]); }
  catch (x) {}
}

function hasLead_(sh, leadId) {
  if (!leadId || sh.getLastRow() < 2) return false;
  var ids = sh.getRange(2, LEAD_ID_COL, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(leadId)) return true;
  }
  return false;
}

/* Не даём формулам выполниться из пользовательского ввода (=, +, -, @ в начале строки) */
function clean_(v) {
  var str = (v === undefined || v === null) ? '' : String(v).slice(0, 1000);
  return /^[=+\-@]/.test(str) ? "'" + str : str;
}

// ===== Availability: normalize Sheets cells and use the same district rule as the site =====
var CLOSE_GROUP_={'Центр':1,'Красный Камень':1,'Вые':1,'Лебяжка':1,'ВМЗ':1};
var FAR_PAIRS_={
 'Вагонка|ВМЗ':1,'Вагонка|Тагилстрой':1,'Вагонка|ГГМ':1,'Вагонка|Старатель':1,'Вагонка|Пригород':1,
 'ГГМ|Тагилстрой':1,'ГГМ|ВМЗ':1,'ГГМ|Пригород':1,'Тагилстрой|ВМЗ':1,'Тагилстрой|Центр':1,'Тагилстрой|Пригород':1
};
function canonicalDistrict_(v){
 var x=String(v||'').trim();
 var a={'Центральный':'Центр','Красный':'Красный Камень','Верхняя Салда':'ВМЗ','В. Салда':'ВМЗ','Тагилстройский':'Тагилстрой','Гагарский':'ГГМ'};
 return a[x]||x;
}
function districtKey_(date,time,district,suburb){var d=canonicalDistrict_(district),s=canonicalDistrict_(suburb);return d==='Пригород'?(s||'Пригород'):d;}
function conflictDistrict_(a,b){
 a=canonicalDistrict_(a);b=canonicalDistrict_(b);if(!a||!b)return false;if(a===b)return false;
 var cityA=!!CLOSE_GROUP_[a]||['Вагонка','Тагилстрой','ГГМ','Старатель'].indexOf(a)>=0;
 var cityB=!!CLOSE_GROUP_[b]||['Вагонка','Тагилстрой','ГГМ','Старатель'].indexOf(b)>=0;
 if(!cityA||!cityB)return true; // suburb/unknown: require extra travel margin
 if(CLOSE_GROUP_[a]&&CLOSE_GROUP_[b])return false;
 return !!FAR_PAIRS_[a+'|'+b]||!!FAR_PAIRS_[b+'|'+a];
}
function busyDate_(v){
 if(v instanceof Date)return Utilities.formatDate(v,'Asia/Yekaterinburg','yyyy-MM-dd');
 var s=String(v||'').trim();if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;
 var m=s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);
 return m?m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2):s;
}
function busyTime_(v){
 if(v instanceof Date)return Utilities.formatDate(v,'Asia/Yekaterinburg','HH:mm');
 if(typeof v==='number'&&isFinite(v)){var minutes=Math.round(v*24*60);return ('0'+Math.floor(minutes/60)).slice(-2)+':'+('0'+(minutes%60)).slice(-2);}
 var s=String(v||'').trim();var m=s.match(/(\d{1,2}):(\d{2})/);return m?('0'+m[1]).slice(-2)+':'+m[2]:s;
}
function busyItems_(){
 var sh=sheet_(SHEET_NAME,HEADERS),last=sh.getLastRow();if(last<2)return [];
 var dc=HEADERS.indexOf('Дата праздника')+1,tc=HEADERS.indexOf('Время')+1,rc=HEADERS.indexOf('Район')+1,sc=HEADERS.indexOf('Населённый пункт')+1;
 var first=Math.min(dc,tc,rc,sc),width=Math.max(dc,tc,rc,sc)-first+1;
 var values=sh.getRange(2,first,last-1,width).getValues(),out=[];
 for(var i=0;i<values.length;i++){
  var date=busyDate_(values[i][dc-first]),time=busyTime_(values[i][tc-first]),district=canonicalDistrict_(values[i][rc-first]),suburb=canonicalDistrict_(values[i][sc-first]);
  if(date&&time&&district)out.push({date:date,time:time,district:district,suburb:suburb});
 }
 return out;
}
function busyBlocks_(date,time,district,suburb){
 var cur=districtKey_(date,time,district,suburb),out=[],hour=parseInt(String(time||'').slice(0,2),10);
 if(!date||!time||!district||!isFinite(hour))return out;
 var items=busyItems_();
 for(var i=0;i<items.length;i++){
  var it=items[i];if(it.date!==busyDate_(date))continue;
  var busyHour=parseInt(String(it.time||'').slice(0,2),10);if(!isFinite(busyHour)||Math.abs(busyHour-hour)>1)continue;
  var booked=districtKey_(it.date,it.time,it.district,it.suburb);
  // Точный занятый час запрещён всем; соседний час — только если районы далеко.
  if(busyHour===hour||(Math.abs(busyHour-hour)===1&&conflictDistrict_(cur,booked)))out.push({date:it.date,time:it.time,district:it.district,suburb:it.suburb,key:booked});
 }
 return out;
}
function timeAllowedByBusy_(date,time,district,suburb){return busyBlocks_(date,time,district,suburb).length===0;}

// Recalculate on the server: a client cannot claim the early-booking discount after the deadline.
function quote_(d) {
  var base = {gift:2500, home:3500, show:4500, corp:5500}[d.formatId];
  if (!base) throw new Error('unknown format');
  var adults=Number(d.corpAdults || 0), children=Number(d.corpChildren || 0), childCount=Number(d.childCount || 1);
  if (d.formatId === 'corp') {
    if (!Number.isInteger(adults) || !Number.isInteger(children) || adults<0 || children<0 || adults>30 || children>30 || adults+children<8 || adults+children>30) throw new Error('invalid guests');
  } else if (!Number.isInteger(childCount) || childCount<1 || childCount>50) throw new Error('invalid child count');
  if (!/^\d{4}-12-(2[0-9]|3[01])$/.test(d.date || '')) throw new Error('invalid date');
  var day=Number(d.date.slice(8)), hour=Number(String(d.time || '').slice(0,2));
  if (!/^\d{2}:00$/.test(d.time || '') || hour<10 || hour>(day===31?23:day===30?22:21)) throw new Error('invalid time');
  var seasonal=day<=29?(hour>=19?600:0):day===30?(hour>=16?1100:500):hour>=23?7500:hour>=22?6500:hour>=20?3700:hour>=16?2500:1200;
  var guests=d.formatId==='corp'?(adults>0 && children>0 ? Math.max(0,adults+children-12)*500 : adults>0 ? Math.max(0,adults-12)*300 : Math.max(0,children-12)*200):Math.max(0,childCount-3)*500;
  var travel=d.district==='Пригород'?1000:0;
  var today=Utilities.formatDate(new Date(),'Asia/Yekaterinburg','yyyy-MM-dd');
  var discount=today<'2026-12-01'?Math.round(base*0.1):0;
  var before=base+guests+seasonal+travel;
  var promoCode=String(d.promoCode || '');
  var promoDiscount=promoCode==='GODNAZAD'?Math.round(before*0.1):0;
  var total=before-discount-promoDiscount;
  d.promoCode=promoCode;d.promoDiscount=promoDiscount;
  d.corpAdults=d.formatId==='corp'?adults:0;d.corpChildren=d.formatId==='corp'?children:0;
  d.childCount=d.formatId==='corp'?'':childCount;
  d.travelFee=travel;d.dateTimeSurcharge=seasonal;d.guestFee=guests;d.formatDiscount=discount;d.beforeDiscount=before;d.basePrice=base;d.childFee=d.formatId==='corp'?0:guests;d.totalNumeric=total;d.pricingRule=d.formatId==='corp'?(adults>0&&children>0?'смешанная группа: общий порог 12, далее +500 ₽/гость':adults>0?'только взрослые: сверх 12 +300 ₽':'только дети: сверх 12 +200 ₽'):'детский формат: сверх 3 детей +500 ₽';
  d.price='от '+String(total).replace(/\B(?=(\d{3})+(?!\d))/g,' ')+' ₽';
  return d;
}

var RATES_HEADERS=['Формат','База, ₽','Входит','Сверх лимита','Сезон 20–29 декабря','Сезон 30 декабря','Сезон 31 декабря','Выезд в пригород','Ранняя скидка','Формула','Дополнительная скидка по промокоду'];
var RATES_ROWS=[
 ['Подарок от Деда Мороза',2500,'до 3 детей','+500 ₽/ребёнок','10–18: +0; 19–21: +600','10–15: +500; 16–22: +1100','10–15: +1200; 16–19: +2500; 20–21: +3700; 22: +6500; 23: +7500',1000,'−250 ₽ до 01.12.2026','База + дети + выезд + сезон − ранняя скидка − скидка по промокоду','GODNAZAD: −10% от базы + гостей/детей + выезда + сезонной надбавки'],
 ['Дед Мороз на дом',3500,'до 3 детей','+500 ₽/ребёнок','10–18: +0; 19–21: +600','10–15: +500; 16–22: +1100','10–15: +1200; 16–19: +2500; 20–21: +3700; 22: +6500; 23: +7500',1000,'−350 ₽ до 01.12.2026','База + дети + выезд + сезон − ранняя скидка − скидка по промокоду','GODNAZAD: −10% от базы + гостей/детей + выезда + сезонной надбавки'],
 ['Интерактивное шоу',4500,'до 3 детей','+500 ₽/ребёнок','10–18: +0; 19–21: +600','10–15: +500; 16–22: +1100','10–15: +1200; 16–19: +2500; 20–21: +3700; 22: +6500; 23: +7500',1000,'−450 ₽ до 01.12.2026','База + дети + выезд + сезон − ранняя скидка − скидка по промокоду','GODNAZAD: −10% от базы + гостей/детей + выезда + сезонной надбавки'],
 ['Корпоратив',5500,'8–30 гостей; база до 12','только взрослые: +300 ₽ после 12; только дети: +200 ₽ после 12; смешанная группа: +500 ₽ после 12 общих гостей','10–18: +0; 19–21: +600','10–15: +500; 16–22: +1100','10–15: +1200; 16–19: +2500; 20–21: +3700; 22: +6500; 23: +7500',1000,'−550 ₽ до 01.12.2026','База + гости + выезд + сезон − ранняя скидка − скидка по промокоду','GODNAZAD: −10% от базы + гостей/детей + выезда + сезонной надбавки']
];
function ensureRates_(){
 var sh=sheet_('Стоимость форматов',RATES_HEADERS);
 for(var r=0;r<RATES_ROWS.length;r++){
   var row=RATES_ROWS[r],rowNumber=r+2;
   if(sh.getLastRow()<rowNumber){sh.appendRow(row);continue;}
   var old=sh.getRange(rowNumber,1,1,row.length).getValues()[0];
   for(var c=0;c<row.length;c++) if(String(old[c])!==String(row[c]))sh.getRange(rowNumber,c+1).setValue(row[c]);
 }
 return sh;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var d = parse_(e);
    if (d.formatId) { try { quote_(d); } catch (err) { return out_({ok:false,error:String(err)},e); } }
    if (!d.name && !d.phone) return out_({ ok: false, error: 'empty' }, e);

    var sh = sheet_(SHEET_NAME, HEADERS);
    ensureRates_();
    if (hasLead_(sh, d.leadId)) return out_({ ok: true, duplicate: true }, e);
    if (d.date && d.time && d.district && !timeAllowedByBusy_(d.date,d.time,d.district,d.suburb)) return out_({ ok:false, reason:'busy', error:'busy' }, e);

    sh.appendRow([
      new Date(), clean_(d.name), clean_(d.phone), clean_(d.date), clean_(d.time),
      clean_(d.format), clean_(d.tariff || ''), clean_(d.price), clean_(d.child), clean_(d.comment),
      clean_(d.source), (d.consent === false || d.consent === 'false') ? 'нет' : 'да', clean_(d.leadId), clean_(d.district), clean_(d.suburb), d.travelFee === 1000 ? 1000 : 0, d.format === 'Корпоратив' ? '' : (Number(d.childCount) || 1), Number(d.dateTimeSurcharge) || 0, Number(d.corpAdults) || 0, Number(d.corpChildren) || 0, Number(d.guestFee) || 0, Number(d.formatDiscount) || 0, Number(d.beforeDiscount) || 0, Number(d.basePrice) || 0, Number(d.childFee) || 0, Number(d.totalNumeric) || 0, clean_(d.pricingRule || ''), clean_(d.promoCode || ''), Number(d.promoDiscount) || 0
    ]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, 'Новая заявка с сайта: ' + (d.name || ''),
        'Имя: ' + (d.name || '') + '\nТелефон: ' + (d.phone || '') +
        '\nДата: ' + (d.date || '') + ' ' + (d.time || '') +
        '\nФормат: ' + (d.format || '') + '\nСтоимость: ' + (d.price || '') +
        '\nРебёнок: ' + (d.child || '') + '\nРайон: ' + (d.district || '') + (d.suburb ? ', ' + d.suburb : '') + '\nДоплата за выезд: ' + (d.travelFee || 0) + ' ₽' + '\nНадбавка за дату и время: ' + (d.dateTimeSurcharge || 0) + ' ₽' + '\nВзрослых: ' + (d.corpAdults || 0) + '; детей на корпоративе: ' + (d.corpChildren || 0) + '\nДоплата за гостей: ' + (d.guestFee || 0) + ' ₽; скидка за ранний заказ: ' + (d.formatDiscount || 0) + ' ₽' + '\nПромокод: ' + (d.promoCode || '—') + '; скидка по промокоду: ' + (d.promoDiscount || 0) + ' ₽' + '\nКоличество детей: ' + (d.childCount || '—') + '\nКомментарий: ' + (d.comment || ''));
    }
    return out_({ ok: true }, e);
  } catch (err) {
    logError_(err);
    return out_({ ok: false, error: String(err) }, e);
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* Проверка связи (?action=ping) и подтверждение записи (?action=check&leadId=...) */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    var sh = sheet_(SHEET_NAME, HEADERS);
    ensureRates_();
    if (p.action === 'check') {
      return out_({ ok: true, found: hasLead_(sh, p.leadId), rows: sh.getLastRow() }, e);
    }
    if (p.action === 'busy') {
      return out_({ ok: true, items: busyItems_() }, e);
    }
    if (p.action === 'checkbusy') {
      var blocks=busyBlocks_(String(p.date||''),String(p.time||''),String(p.district||''),String(p.suburb||''));
      if(blocks.length) return out_({ ok:false, reason:'busy', blocks: blocks.slice(0,5) }, e);
      return out_({ ok:true, allowed:true }, e);
    }
    return out_({ ok: true, sheet: sh.getName(), rows: sh.getLastRow(),
                  note: 'Приём заявок работает.' }, e);
  } catch (err) {
    logError_(err);
    return out_({ ok: false, error: String(err) }, e);
  }
}

/* Ответ в виде JSON или JSONP (если сайт прислал callback) */
function out_(obj, e) {
  var cb = (e && e.parameter && e.parameter.callback) ? String(e.parameter.callback) : '';
  if (cb && !/^[A-Za-z_$][\w$]*$/.test(cb)) cb = '';
  var body = JSON.stringify(obj);
  return cb
    ? ContentService.createTextOutput(cb + '(' + body + ')').setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}
