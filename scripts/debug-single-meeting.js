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
    const res = await axios.post(tokenUrl, params.toString(), {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });
    return res.data.access_token;
}

async function debugMeeting5() {
    try {
        const token = await getAccessToken();
        const headers = { Authorization: `Bearer ${token}` };

        const meeting = await Meeting.findByPk(5);
        console.log(`📌 Testing Meeting ID 5: "${meeting.subject}"`);
        console.log(`   Organizer: ${meeting.creatorEmail}`);
        console.log(`   Raw DB Co-Organizers: ${meeting.coOrganizers}\n`);

        // 1. Resolve Organizer GUID
        const orgRes = await axios.get(`https://graph.microsoft.com/v1.0/users/${meeting.creatorEmail}?$select=id`, { headers });
        const organizerGuid = orgRes.data.id;
        console.log(`✅ Organizer GUID: ${organizerGuid}`);

        // 2. Resolve Co-Organizers GUIDs cleanly (without extra 'role' property)
        const rawCoOrgs = JSON.parse(meeting.coOrganizers);
        const coOrgPromises = rawCoOrgs
            .map(c => c.email ? c.email.replace(/\s+/g, '').toLowerCase() : null)
            .filter(e => e && e !== meeting.creatorEmail.toLowerCase())
            .map(async (email) => {
                try {
                    const userRes = await axios.get(`https://graph.microsoft.com/v1.0/users/${email}?$select=id`, { headers });
                    console.log(`✅ Co-Organizer GUID found for ${email}: ${userRes.data.id}`);
                    return {
                        identity: { user: { id: userRes.data.id } },
                        upn: email
                    };
                } catch (err) {
                    console.log(`❌ Failed lookup for ${email}: ${err.message}`);
                    return null;
                }
            });

        const coOrganizersArray = (await Promise.all(coOrgPromises)).filter(Boolean);

        // 3. Send Clean PATCH Payload
        const patchPayload = {
            allowedPresenters: 'roleIsPresenter',
            lobbyBypassSettings: { scope: 'organization' },
            participants: {
                coOrganizers: coOrganizersArray
            }
        };

        console.log('\n🚀 Sending PATCH payload to Graph API...');
        const patchEndpoint = `https://graph.microsoft.com/v1.0/users/${organizerGuid}/onlineMeetings/${meeting.teamsMeetingId}`;
        await axios.patch(patchEndpoint, patchPayload, { headers });
        console.log('✅ PATCH request succeeded!');

        // 4. Fetch live meeting to verify
        console.log('\n🔍 Re-fetching live meeting from Graph API...');
        const getRes = await axios.get(patchEndpoint, { headers });
        console.log('\n================ LIVE PARTICIPANTS RESPONSE ================');
        console.log(JSON.stringify(getRes.data.participants, null, 2));
        console.log('============================================================');

    } catch (err) {
        console.error('💥 Error:', err.response?.data || err.message);
    }
}

debugMeeting5();