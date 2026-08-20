const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '../.env') });
const axios = require('axios');
const { Meeting } = require('../models');

const TENANT_ID = process.env.AZURE_TENANT_ID;
const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

async function getAccessToken() {
    const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
    const params = new URLSearchParams({
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        grant_type: 'client_credentials',
        scope: 'https://graph.microsoft.com/.default'
    });

    const response = await axios.post(tokenUrl, params.toString(), {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });

    return response.data.access_token;
}

async function verifyBatch30() {
    console.log('🔍 Verifying live Microsoft Graph API settings for the first 30 meetings...\n');

    try {
        const token = await getAccessToken();
        const headers = { Authorization: `Bearer ${token}` };

        const meetings = await Meeting.findAll();
        const testMeetings = meetings.filter((m) => m.teamsMeetingId).slice(0, 30);

        for (let i = 0; i < testMeetings.length; i++) {
            const m = testMeetings[i];
            console.log(`[${i + 1}/30] Checking ID ${m.id}: "${m.subject}"`);

            try {
                // 1. Resolve Organizer GUID
                const userRes = await axios.get(
                    `https://graph.microsoft.com/v1.0/users/${m.creatorEmail}?$select=id`,
                    { headers }
                );
                const organizerGuid = userRes.data.id;

                // 2. Fetch live meeting object from Graph API
                const endpoint = `https://graph.microsoft.com/v1.0/users/${organizerGuid}/onlineMeetings/${m.teamsMeetingId}`;
                const res = await axios.get(endpoint, { headers });
                const liveMeeting = res.data;

                const coOrgs = liveMeeting.participants?.coOrganizers || [];
                const coOrgUPNs = coOrgs.map(c => c.upn || c.identity?.user?.id);

                console.log(`   • Allowed Presenters: ${liveMeeting.allowedPresenters}`);
                console.log(`   • Lobby Bypass Scope: ${liveMeeting.lobbyBypassSettings?.scope}`);
                console.log(`   • Live Co-Organizers: ${coOrgs.length > 0 ? coOrgUPNs.join(', ') : 'None / Empty'}\n`);

            } catch (err) {
                console.log(`   ❌ Could not verify meeting ID ${m.id}: ${err.response?.data?.error?.message || err.message}\n`);
            }
        }

        process.exit(0);
    } catch (error) {
        console.error('💥 Verification failed:', error);
        process.exit(1);
    }
}

verifyBatch30();