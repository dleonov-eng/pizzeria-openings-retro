/** @OnlyCurrentDoc */
const RETRO_ID = '1XT382mmH0xkTwaAqO0on7GSnPfrHFnE443TFUKiCGLY';

function onOpen() {
  addRetroMenu_();
}

function addRetroMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('Ретро')
    .addItem('Открыть лендинг', 'showRetroReport')
    .addToUi();
}

function retroBook_() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book || book.getId() !== RETRO_ID) {
    throw new Error('Откройте скрипт из таблицы «Ретро. Открытия».');
  }
  return book;
}

function showRetroReport() {
  retroBook_();
  const page = HtmlService.createHtmlOutput(`<!doctype html>
    <html lang="ru"><head><meta charset="utf-8">
    <style>html,body,iframe {width:100%;height:100%;margin:0;border:0;display:block}</style>
    </head><body>
    <iframe id="report" title="Ретро открытий пиццерий"></iframe>
    <script>
      const report = document.getElementById('report');
      const site = new URL('https://dleonov-eng.github.io/pizzeria-openings-retro/');
      window.addEventListener('message', function (event) {
        if (event.origin !== site.origin || event.source !== report.contentWindow) return;
        if (event.data !== 'retro.readSummary' || event.ports.length !== 1) return;
        const port = event.ports[0];
        function reply(message) {
          port.postMessage(message);
          port.close();
        }
        google.script.run
          .withFailureHandler(function (error) { reply({error: error.message}); })
          .withSuccessHandler(function (result) { reply({result: result}); })
          .readRetroSummary();
      });
      report.src = site.href + '#sheet';
    </script></body></html>`)
    .setWidth(1380)
    .setHeight(820);
  SpreadsheetApp.getUi().showModelessDialog(page, 'Ретро открытий пиццерий');
}

function readRetroSummary() {
  const book = retroBook_();
  const sheet = book.getSheetByName('Сводная');
  if (!sheet) throw new Error('Вкладка «Сводная» не найдена.');

  const count = sheet.getLastRow();
  if (!count) throw new Error('Вкладка «Сводная» пуста.');

  const zone = book.getSpreadsheetTimeZone();
  const values = sheet.getRange(1, 1, count, 40).getValues().map(row =>
    row.map(cell => cell instanceof Date
      ? Utilities.formatDate(cell, zone, 'yyyy-MM-dd')
      : cell)
  );
  return {values, readAt: new Date().toISOString()};
}
