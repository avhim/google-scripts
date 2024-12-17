function getEmailsFromSender() {
  const senderEmail = "some_user@gmail.com";
  const sheetId = 'SHEET-ID'; // part of URL
  const sheet = SpreadsheetApp.openById(sheetId);

  const now = new Date();
  const firstDayOfCurrentMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const lastDayOfPreviousMonth = new Date(firstDayOfCurrentMonth - 1);
  const firstDayOfPreviousMonth = new Date(lastDayOfPreviousMonth.getFullYear(), lastDayOfPreviousMonth.getMonth(), 1);

  const startDate = Utilities.formatDate(firstDayOfPreviousMonth, Session.getScriptTimeZone(), "yyyy/MM/dd");
  const endDate = Utilities.formatDate(lastDayOfPreviousMonth, Session.getScriptTimeZone(), "yyyy/MM/dd");

  if (!sheet.getSheetByName(startDate)) {
    sheet.insertSheet(startDate);
  }

  const headers = ['created', 'date_tour', 'tour', 'cost', 'name', 'phone', 'email', 'details', 'link_to_email'];
  sheet.appendRow(headers);
  
  const query = `subject:(Заявка с сайта) from:${senderEmail} after:${startDate} before:${endDate}`;
  const threads = GmailApp.search(query);

  const dateAndTourRegex = /(\d{2}\.\d{2}\.\d{4} - [^\n]+)/;
  const costRegex = /(\d+)\s*BYN/g;
  const nameRegex = /Имя:\s*([^\n]+)/;
  const phoneRegex = /Телефон:\s*([^\n]+)/;
  const emailRegex = /Email:\s*([^\n]+)/; 
  const detailsRegex = /Детали заявки:\s*(https:\/\/[^\s]+)/;

  function sumBYN(body) {
    let total = 0;
    let match;
    while ((match = costRegex.exec(body)) !== null) {
      total += parseInt(match[1], 10);
    }
    return total;
  }

  // Множество для отслеживания уникальных комбинаций телефона, тура, даты и стоимости
  const uniqueRequests = new Set();

  if (threads.length > 0) {
    for (let thread of threads) {
      const messages = thread.getMessages();
      for (let message of messages) {
        const created = message.getDate();
        const body = message.getPlainBody();

        const dateAndTourMatch = body.match(dateAndTourRegex);
        const dateTour = dateAndTourMatch ? dateAndTourMatch[1].split(' - ')[0] : null;
        const tour = dateAndTourMatch ? dateAndTourMatch[1].split(' - ')[1] : null;
        const cost = sumBYN(body);
        const name = body.match(nameRegex)?.[1] || null;
        const phone = body.match(phoneRegex)?.[1]?.replace(/[^\d]/g, '') || null;
        const email = body.match(emailRegex)?.[1] || null;
        const details = body.match(detailsRegex)?.[1] || null;

        // Генерация уникального идентификатора для комбинации телефона, тура, даты и стоимости
        const uniqueIdentifier = `${phone}_${dateTour}_${tour}_${cost}`;

        // Проверка на дубликат
        if (uniqueRequests.has(uniqueIdentifier)) {
          Logger.log(`Duplicate found for ${uniqueIdentifier}. Skipping.`);
          continue;
        }

        uniqueRequests.add(uniqueIdentifier);

        const row = [created, dateTour, tour, cost, name, phone, email, details, thread.getPermalink()];
        sheet.appendRow(row);
      }
    }
  } else {
    Logger.log(`No emails found from ${senderEmail} in period: ${startDate} - ${endDate}`);
  }
}
