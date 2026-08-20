const axios = require('axios');
require('dotenv').config();

const TENANT_ID = process.env.AZURE_TENANT_ID;
const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

// 1. Enter the organizer's email directly
const ORGANIZER_EMAIL = 'bcom1@mubs.ac.ug';

// 2. Full Teams Join URL from Moodle
const TEAMS_JOIN_URL = 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_ZjFiNTliYTgtN2UwNS00MmRkLTg1MmItM2RjNTJhZTVhYTkw%40thread.v2/0?context=%7b%22Tid%22%3a%22e9220e78-c793-4150-b529-c9bbb0e979d3%22%2c%22Oid%22%3a%22a6f367fc-6b0c-4b19-b28c-6b384f5d74f3%22%7d';

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

async function inspectMeetingPermissions() {
  try {
    console.log('🔑 Authenticating with Microsoft Graph API...');
    const token = await getAccessToken();
    const headers = { Authorization: `Bearer ${token}` };

    // Step 1: Automatically resolve email to Azure GUID
    console.log(`🔍 Resolving Azure User GUID for ${ORGANIZER_EMAIL}...`);
    const userRes = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${ORGANIZER_EMAIL}?$select=id`,
      { headers }
    );
    const organizerGuid = userRes.data.id;
    console.log(`✅ Resolved GUID: ${organizerGuid}`);

    // Step 2: Fetch meeting using resolved GUID
    console.log('🔍 Fetching live meeting permissions...');
    const endpoint = `https://graph.microsoft.com/v1.0/users/${organizerGuid}/onlineMeetings?$filter=joinWebUrl eq '${encodeURIComponent(TEAMS_JOIN_URL)}'`;
    
    const res = await axios.get(endpoint, { headers });
    const meetings = res.data.value;

    if (!meetings || meetings.length === 0) {
      console.log('❌ Meeting not found for user:', ORGANIZER_EMAIL);
      return;
    }

    const meeting = meetings[0];

    console.log('\n================ TEAMS SECURITY PERMISSIONS ================');
    console.log(`Meeting ID:          ${meeting.id}`);
    console.log(`Subject:             ${meeting.subject || 'N/A'}`);
    console.log(`Allowed Presenters:  ${meeting.allowedPresenters}`); 
    console.log(`Lobby Bypass Scope:  ${meeting.lobbyBypassSettings?.scope || 'N/A'}`);
    console.log('\n--- PARTICIPANTS CONFIGURATION ---');
    console.log('Organizer:', JSON.stringify(meeting.participants?.organizer, null, 2));
    console.log('Co-Organizers:', JSON.stringify(meeting.participants?.coOrganizers, null, 2));
    console.log('Attendees:', JSON.stringify(meeting.participants?.attendees, null, 2));
    console.log('============================================================\n');

    if (meeting.allowedPresenters === 'organization') {
      console.log('⚠️ SECURITY RISK DETECTED: allowedPresenters is set to "organization".');
      console.log('👉 ALL logged-in students are granted Presenter role (can mute/kick others).');
    } else if (meeting.allowedPresenters === 'roleIsPresenter') {
      console.log('✅ SECURE: allowedPresenters is set to "roleIsPresenter".');
      console.log('👉 Only Organizers and Co-Organizers are Presenters. Students are Attendees.');
    }

  } catch (error) {
    const errorDetails = error.response?.data || error.message;
    console.error('❌ Failed to inspect meeting:', JSON.stringify(errorDetails, null, 2));
  }
}

inspectMeetingPermissions();