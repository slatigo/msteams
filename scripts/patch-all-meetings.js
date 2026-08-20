const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '../.env') });
const { Meeting } = require('../models');
const { updateTeamsMeeting } = require('../services/msGraphService');

async function patchAllMeetings() {
    console.log('🚀 Starting FULL batch patch for all Microsoft Teams meetings in the database...\n');

    try {
        const meetings = await Meeting.findAll();
        const validMeetings = meetings.filter(m => m.teamsMeetingId);

        console.log(`📊 Found ${validMeetings.length} meetings to process.\n`);

        let successCount = 0;
        let failCount = 0;

        for (let i = 0; i < validMeetings.length; i++) {
            const meeting = validMeetings[i];
            console.log(`[${i + 1}/${validMeetings.length}] Patching ID ${meeting.id}: "${meeting.subject}"`);

            try {
                let coOrgs = [];
                if (meeting.coOrganizers) {
                    coOrgs = typeof meeting.coOrganizers === 'string'
                        ? JSON.parse(meeting.coOrganizers)
                        : meeting.coOrganizers;
                }

                if (coOrgs.length > 0) {
                    console.log(`   📋 DB Co-Organizers: ${coOrgs.join(', ')}`);
                } else {
                    console.log(`   ℹ️ No co-organizers in DB for this meeting`);
                }

                await updateTeamsMeeting({
                    organizerEmail: meeting.creatorEmail,
                    teamsMeetingId: meeting.teamsMeetingId,
                    subject: meeting.subject,
                    startDateTime: meeting.startTime,
                    endDateTime: meeting.endTime,
                    coOrganizers: coOrgs
                });

                console.log(`   ✅ Patch applied successfully.\n`);
                successCount++;
            } catch (err) {
                console.error(`   ❌ Failed to patch meeting ID ${meeting.id}: ${err.message}\n`);
                failCount++;
            }

            // Small delay to prevent API rate limiting (Throttling / 429)
            await new Promise(resolve => setTimeout(resolve, 200));
        }

        console.log('====================================================');
        console.log(`🎉 Full database patch completed!`);
        console.log(`✅ Success: ${successCount}`);
        console.log(`❌ Failed:  ${failCount}`);
        console.log('====================================================');

        process.exit(0);
    } catch (error) {
        console.error('💥 Fatal error during batch patch:', error);
        process.exit(1);
    }
}

patchAllMeetings();