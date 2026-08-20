const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '../.env') });
const { updateTeamsMeeting } = require('../services/msGraphService');

const testData = [
  {
    "id": "5",
    "subjectCode": "TOT",
    "subject": "TOT - Online Refresher Training Programme Virtual Classroom",
    "teamsMeetingId": "MSo4OGQ0MzMyNC05Njk1LTQxOTAtOWExYi0zOGE0YjNjYjM5NWYqMCoqMTk6bWVldGluZ19Nalk1TkRJMk5UTXRabU5qWXkwME9XVTRMVGc1TWpZdE5UUmtaVEZoT0RabVltTTVAdGhyZWFkLnYy",
    "creatorEmail": "emigadde@mubs.ac.ug",
    "meetingDate": "2026-08-08",
    "startTime": "10:00",
    "endTime": "12:00",
    "coOrganizers": "[{\"name\":\"Latigo Simon Peter\",\"email\":\"slatigo@mubs.ac.ug\"},{\"name\":\"Nagujja Shakira\",\"email\":\"snagujja@mubs.ac.ug\"},{\"name\":\"Lubega Juma\",\"email\":\"jlubega@mubs.ac.ug\"},{\"name\":\"Migadde Elias\",\"email\":\"emigadde@mubs.ac.ug\"}]"
  },
  {
    "id": "748",
    "subjectCode": "MBA7101",
    "subject": "MBA7101 - GENERAL MANAGEMENT Virtual Classroom",
    "teamsMeetingId": "MSo4N2Y4MGQyOS00NGU2LTQxYjYtYTU4Yy00NDAzN2E1ZGUyMjUqMCoqMTk6bWVldGluZ19ZamRtTlRRd05HRXRaR1ZrTkMwME9ETXlMV0ZsTlRRdE5XTTNNR1EwWXpneE16VmpAdGhyZWFkLnYy",
    "creatorEmail": "mba1@mubs.ac.ug",
    "meetingDate": "2026-08-10",
    "startTime": "15:31",
    "endTime": "23:59",
    "coOrganizers": "[{\"name\":\"Prof. Vincent Bagire (PhD)\",\"email\":\"vbagire@mubs.ac.ug\"},{\"name\":\"Shakirah Nagujja\",\"email\":\"snagujja@mubs.ac.ug\"}]"
  },
  {
    "id": "749",
    "subjectCode": "MBA7101",
    "subject": "MBA7101 - GENERAL MANAGEMENT Virtual Classroom",
    "teamsMeetingId": "MSo4N2Y4MGQyOS00NGU2LTQxYjYtYTU4Yy00NDAzN2E1ZGUyMjUqMCoqMTk6bWVldGluZ19ZekV4WXpBME9USXRORE5pT1MwME5XUTVMV0l3WkRNdFl6ZGlOMk5oTW1RMU1UaG1AdGhyZWFkLnYy",
    "creatorEmail": "mba1@mubs.ac.ug",
    "meetingDate": "2026-08-10",
    "startTime": "15:32",
    "endTime": "23:59",
    "coOrganizers": "[{\"name\":\"Prof. Vincent Bagire (PhD)\",\"email\":\"vbagire@mubs.ac.ug\"},{\"name\":\"Shakirah Nagujja\",\"email\":\"snagujja@mubs.ac.ug\"}]"
  }
];

async function runJsonTest() {
    console.log(`🧪 Starting JSON DATA TEST for ${testData.length} meetings...\n`);

    for (let i = 0; i < testData.length; i++) {
        const item = testData[i];
        console.log(`[${i + 1}/${testData.length}] Patching ID ${item.id}: "${item.subject}"`);

        let coOrganizersList = [];
        if (item.coOrganizers) {
            try {
                coOrganizersList = typeof item.coOrganizers === 'string' 
                    ? JSON.parse(item.coOrganizers) 
                    : item.coOrganizers;
            } catch (e) {
                console.warn(`   ⚠️ Failed to parse coOrganizers JSON: ${e.message}`);
            }
        }

        const emails = coOrganizersList.map(c => c.email).filter(Boolean);
        if (emails.length > 0) {
            console.log(`   📋 Co-Organizers: ${emails.join(', ')}`);
        } else {
            console.log(`   ℹ️ No co-organizers for this meeting`);
        }

        const startDateTime = `${item.meetingDate}T${item.startTime}:00Z`;
        const endDateTime = `${item.meetingDate}T${item.endTime}:00Z`;

        try {
            await updateTeamsMeeting({
                organizerEmail: item.creatorEmail,
                teamsMeetingId: item.teamsMeetingId,
                subject: item.subject,
                startDateTime: startDateTime,
                endDateTime: endDateTime,
                coOrganizers: emails
            });
            console.log(`   ✅ Patch applied successfully.\n`);
        } catch (error) {
            console.error(`   ❌ Failed to patch meeting ID ${item.id}: ${error.message}\n`);
        }
    }

    console.log(`🎉 JSON TEST COMPLETED.`);
}

runJsonTest();