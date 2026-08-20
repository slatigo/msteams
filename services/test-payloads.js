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

async function testPayloads() {
    try {
        const token = await getAccessToken();
        const headers = { Authorization: `Bearer ${token}` };

        const meeting = await Meeting.findByPk(5);
        console.log(`📌 Testing payload schema on Meeting ID 5...`);

        const orgRes = await axios.get(`https://graph.microsoft.com/v1.0/users/${meeting.creatorEmail}?$select=id`, { headers });
        const organizerGuid = orgRes.data.id;

        const rawCoOrgs = JSON.parse(meeting.coOrganizers);
        const coOrgPromises = rawCoOrgs
            .map(c => c.email ? c.email.replace(/\s+/g, '').toLowerCase() : null)
            .filter(e => e && e !== meeting.creatorEmail.toLowerCase())
            .map(async (email) => {
                const userRes = await axios.get(`https://graph.microsoft.com/v1.0/users/${email}?$select=id`, { headers });
                return {
                    identity: { user: { id: userRes.data.id } },
                    upn: email
                };
            });

        const coOrganizersArray = (await Promise.all(coOrgPromises)).filter(Boolean);
        const patchEndpoint = `https://graph.microsoft.com/v1.0/users/${organizerGuid}/onlineMeetings/${meeting.teamsMeetingId}`;

        // Map co-organizers into attendees array with explicit role
        const attendeesArray = coOrganizersArray.map(c => ({
            identity: c.identity,
            upn: c.upn,
            role: 'presenter'
        }));

        console.log('\n🚀 Sending PATCH payload with attendees + roles specified...');
        const patchPayload = {
            allowedPresenters: 'roleIsPresenter',
            lobbyBypassSettings: { scope: 'organization' },
            participants: {
                coOrganizers: coOrganizersArray,
                attendees: attendeesArray
            }
        };

        await axios.patch(patchEndpoint, patchPayload, { headers });
        console.log('✅ PATCH SUCCESSFUL!');

        console.log('\n🔍 Fetching live meeting from Graph API...');
        const getRes = await axios.get(patchEndpoint, { headers });
        console.log('\n================ LIVE PARTICIPANTS ================');
        console.log(JSON.stringify(getRes.data.participants, null, 2));
        console.log('==================================================');

    } catch (err) {
        console.error('❌ Error:', err.response?.data || err.message);
    }
}

testPayloads();