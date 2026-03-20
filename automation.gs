function createCalendarEventsFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getDefaultCalendar();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // --- Column Mapping ---
    const timestamp        = row[0];   // A - Timestamp
    const businessUnit     = row[1];   // B - Business Unit
    const eventName        = row[2];   // C - Event Name
    const eventCategory    = row[3];   // D - Event Category
    const eventType        = row[4];   // E - Event Type
    const startDate        = row[5];   // F - Event Start Date
    const endDate          = row[6];   // G - Event End Date
    const startTimeStr     = row[7];   // H - Event Start Time
    const endTimeStr       = row[8];   // I - Event End Time
    const location         = row[9];   // J - Event Location/Platform
    const objective        = row[10];  // K - Event Objective
    const targetAudience   = row[11];  // L - Target Audience
    const numParticipants  = row[12];  // M - Estimated Number of Participants
    const priorityLevel    = row[13];  // N - Priority Level
    const picName          = row[14];  // O - PIC Full Name
    const picEmail         = row[15];  // P - PIC Email Address
    const picContact       = row[16];  // Q - PIC Contact Number

    // Skip empty rows
    if (!eventName || !startDate) continue;

    // Skip already processed rows
    if (row[23] === '✅ Created') continue; 

    // Build start & end datetime
    const startDateTime = combineDateAndTime(new Date(startDate), startTimeStr);
    const endDateTime   = combineDateAndTime(new Date(endDate), endTimeStr);

    // Build description
    const description = [
    `📋 Business Unit: ${businessUnit}`,
    `🏷️ Category: ${eventCategory}`,
    `⚡ Type: ${eventType}`,
    `🎯 Objective: ${objective}`,
    `👥 Target Audience: ${targetAudience}`,
    `🔢 Est. Participants: ${numParticipants}`,
    `🚨 Priority: ${priorityLevel}`,
    `👤 PIC: ${picName}`,
    `📧 PIC Email: ${picEmail}`,
    `📞 PIC Contact: ${picContact}`,
    `🕐 Submitted: ${timestamp}`
      ].join('\n');

    // Create the calendar event
    calendar.createEvent(eventName, startDateTime, endDateTime, {
      location: location,
      description: description
    });

    // Mark row as created
    sheet.getRange(i + 1, 24).setValue('✅ Created');

    Logger.log(`✅ Created: ${eventName} | ${startDateTime} → ${endDateTime}`);
  }

  SpreadsheetApp.getUi().alert('🎉 All events have been added to your calendar!');
  Logger.log('🎉 All events have been added to your calendar!');
}

function combineDateAndTime(date, timeStr) {
  const result = new Date(date);

  if (timeStr) {
    const timeParts = timeStr.toString().match(/(\d+):(\d+):?(\d*)\s*(AM|PM)?/i);
    if (timeParts) {
      let hours   = parseInt(timeParts[1]);
      const mins  = parseInt(timeParts[2]);
      const ampm  = timeParts[4];

      if (ampm && ampm.toUpperCase() === 'PM' && hours !== 12) hours += 12;
      if (ampm && ampm.toUpperCase() === 'AM' && hours === 12) hours = 0;

      result.setHours(hours, mins, 0, 0);
    }
  }

  return result;
}