require('dotenv').config();
const { Meeting } = require('../models'); // Adjust path to models index if needed
const { updateTeamsMeeting } = require('../services/msGraphService');

// Utility to safely parse coOrganizers JSON from TEXT column
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

// Utility to normalize ISO string from DATEONLY + STRING time
function buildIsoDateString(dateStr, timeStr, defaultTime) {
    if (!dateStr) return new Date().toISOString();
    const cleanTime = (timeStr || defaultTime).trim();
    // Ensures time format is HH:MM:SS
    const formattedTime = cleanTime.length === 5 ? `${cleanTime}:00` : cleanTime;
    const dateObj = new Date(`${dateStr}T${formattedTime}`);
    return isNaN(dateObj.getTime()) ? new Date().toISOString() : dateObj.toISOString();
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function runBulkSecurityPatch() {
    console.log('🚀 Starting bulk security patch for existing Teams meetings...\n');

    try {
        const meetings = await Meeting.findAll();
        const teamsMeetings = meetings.filter((m) => m.teamsMeetingId);

        console.log(`📊 Found ${teamsMeetings.length} active Teams meeting(s) to patch.\n`);

        let successCount = 0;
        let failureCount = 0;

        for (let i = 0; i < teamsMeetings.length; i++) {
            const meeting = teamsMeetings[i];
            const indexStr = `[${i + 1}/${teamsMeetings.length}]`;

            try {
                // Extract co-organizer emails from JSON TEXT field
                const rawCoOrgs = parseCoOrganizers(meeting.coOrganizers);
                const coOrganizerEmails = rawCoOrgs
                    .map((c) => (typeof c === 'string' ? c : c.email))
                    .filter(Boolean);

                const startDateTime = buildIsoDateString(meeting.meetingDate, meeting.startTime, '08:00');
                const endDateTime = buildIsoDateString(meeting.expiryDate || meeting.meetingDate, meeting.endTime, '17:00');

                console.log(`${indexStr} Patching: "${meeting.subject}" (ID: ${meeting.id})`);

                // Send security patch to MS Graph API
                await updateTeamsMeeting({
                    organizerEmail: meeting.creatorEmail,
                    teamsMeetingId: meeting.teamsMeetingId,
                    subject: meeting.subject,
                    description: meeting.description,
                    startDateTime,
                    endDateTime,
                    coOrganizers: coOrganizerEmails
                });

                console.log(`   ✅ Permissions updated (roleIsPresenter + coOrganizers mapped).\n`);
                successCount++;
            } catch (err) {
                console.error(`   ❌ Failed to patch meeting ID ${meeting.id}: ${err.message}\n`);
                failureCount++;
            }

            // Rate-limit buffer for Microsoft Graph API
            await sleep(500);
        }

        console.log('====================================================');
        console.log(`🎉 Bulk update completed!`);
        console.log(`✅ Success: ${successCount}`);
        console.log(`❌ Failed:  ${failureCount}`);
        console.log('====================================================');

        process.exit(0);
    } catch (error) {
        console.error('💥 Fatal error during bulk update:', error);
        process.exit(1);
    }
}

runBulkSecurityPatch();