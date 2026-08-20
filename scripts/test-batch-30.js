const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '../.env') });
const { Meeting } = require('../models');
const { updateTeamsMeeting } = require('../services/msGraphService');

function parseCoOrganizers(coOrgsData) {
    if (!coOrgsData) return [];
    if (Array.isArray(coOrgsData)) return coOrgsData;
    try {
        const parsed = JSON.parse(coOrgsData);
        return Array.isArray(parsed) ? parsed : [];
    } catch {
        return [];
    }
}

function buildIsoDateString(dateStr, timeStr, defaultTime) {
    if (!dateStr) return new Date().toISOString();
    const cleanTime = (timeStr || defaultTime).trim();
    const formattedTime = cleanTime.length === 5 ? `${cleanTime}:00` : cleanTime;
    const dateObj = new Date(`${dateStr}T${formattedTime}`);
    return isNaN(dateObj.getTime()) ? new Date().toISOString() : dateObj.toISOString();
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function runTestBatch30() {
    console.log('🧪 Starting TEST RUN for the FIRST 30 Teams meetings...\n');

    try {
        const meetings = await Meeting.findAll();
        // 🔒 Strictly process only the first 30 active meetings
        const teamsMeetings = meetings.filter((m) => m.teamsMeetingId).slice(0, 30);

        console.log(`📊 Loaded ${teamsMeetings.length} meetings for batch testing.\n`);

        let successCount = 0;
        let failureCount = 0;

        for (let i = 0; i < teamsMeetings.length; i++) {
            const meeting = teamsMeetings[i];
            const indexStr = `[${i + 1}/${teamsMeetings.length}]`;

            try {
                const rawCoOrgs = parseCoOrganizers(meeting.coOrganizers);
                const coOrganizerEmails = rawCoOrgs
                    .map((c) => (typeof c === 'string' ? c : c?.email))
                    .filter(Boolean);

                const startDateTime = buildIsoDateString(meeting.meetingDate, meeting.startTime, '08:00');
                const endDateTime = buildIsoDateString(meeting.expiryDate || meeting.meetingDate, meeting.endTime, '17:00');

                console.log(`${indexStr} Patching ID ${meeting.id}: "${meeting.subject}"`);
                if (coOrganizerEmails.length > 0) {
                    console.log(`   📋 DB Co-Organizers: ${coOrganizerEmails.join(', ')}`);
                } else {
                    console.log(`   ℹ️ No co-organizers in DB for this meeting`);
                }

                await updateTeamsMeeting({
                    organizerEmail: meeting.creatorEmail,
                    teamsMeetingId: meeting.teamsMeetingId,
                    subject: meeting.subject,
                    description: meeting.description,
                    startDateTime,
                    endDateTime,
                    coOrganizers: coOrganizerEmails
                });

                console.log(`   ✅ Patch applied successfully.\n`);
                successCount++;
            } catch (err) {
                console.error(`   ❌ Failed to patch meeting ID ${meeting.id}: ${err.message}\n`);
                failureCount++;
            }

            await sleep(500);
        }

        console.log('====================================================');
        console.log(`🎉 Test batch completed!`);
        console.log(`✅ Success: ${successCount}`);
        console.log(`❌ Failed:  ${failureCount}`);
        console.log('====================================================');

        process.exit(0);
    } catch (error) {
        console.error('💥 Fatal error during test batch:', error);
        process.exit(1);
    }
}

runTestBatch30();