// services/msGraphService.js
// Complete working version with all functionality

const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const dns = require('dns');

dns.setDefaultResultOrder('ipv4first');

const msalConfig = {
    auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
        clientSecret: process.env.AZURE_CLIENT_SECRET,
    }
};

const cca = new msal.ConfidentialClientApplication(msalConfig);
const COORGANIZER_PREFER_HEADER = 'include-unknown-enum-members';

async function getAccessToken() {
    const authResult = await cca.acquireTokenByClientCredential({
        scopes: ['https://graph.microsoft.com/.default'],
    });
    return authResult.accessToken;
}

async function getGraphClient() {
    const token = await getAccessToken();
    return Client.init({
        authProvider: (done) => done(null, token)
    });
}

async function getUserGuidByEmail(client, email) {
    try {
        const user = await client.api(`/users/${email}`).select('id,displayName,userPrincipalName').get();
        return {
            id: user.id,
            displayName: user.displayName || email,
            email: user.userPrincipalName || email
        };
    } catch (error) {
        throw new Error(`Azure AD user lookup failed for ${email}: ${error.message}`);
    }
}

async function resolveCoOrganizers(client, coOrganizers, organizerEmail) {
    let resolvedUsers = [];

    if (Array.isArray(coOrganizers)) {
        const filteredEmails = coOrganizers
            .map(c => {
                const rawEmail = typeof c === 'string' ? c : c?.email || c?.upn;
                return rawEmail ? rawEmail.replace(/\s+/g, '').toLowerCase() : null;
            })
            .filter(email => email && email !== organizerEmail.toLowerCase());

        if (filteredEmails.length === 0) {
            return resolvedUsers;
        }

        console.log(`📧 Resolving ${filteredEmails.length} co-organizers:`, filteredEmails);

        const coOrgPromises = filteredEmails.map(async (email) => {
            try {
                const user = await getUserGuidByEmail(client, email);
                console.log(`✅ Resolved ${email} -> ${user.id}`);
                return {
                    id: user.id,
                    email: user.email,
                    displayName: user.displayName
                };
            } catch (err) {
                console.warn(`⚠️ Skipping co-organizer ${email}: ${err.message}`);
                return null;
            }
        });

        resolvedUsers = (await Promise.all(coOrgPromises)).filter(Boolean);
        console.log(`✅ Successfully resolved ${resolvedUsers.length} co-organizers`);
    }

    return resolvedUsers;
}

function buildParticipantsPayload(resolvedUsers) {
    if (!resolvedUsers || resolvedUsers.length === 0) {
        return {
            allowedPresenters: 'organizer',
            participants: {
                attendees: []
            }
        };
    }

    const attendeesArray = resolvedUsers.map(u => ({
        identity: { 
            user: { 
                id: u.id,
                displayName: u.displayName
            } 
        },
        upn: u.email,
        role: 'coorganizer'
    }));

    return {
        allowedPresenters: 'roleIsPresenter',
        participants: {
            attendees: attendeesArray
        }
    };
}

// CREATE MEETING - All settings can be set here
async function createTeamsMeeting({ 
    organizerEmail, 
    subject, 
    startDateTime, 
    endDateTime, 
    coOrganizers = [],
    lobbyBypassScope = 'organization',
    isEntryExitAnnounced = true,
    autoAdmittedUsers = 'everyoneInSameOrganization'
}) {
    const client = await getGraphClient();
    const organizer = await getUserGuidByEmail(client, organizerEmail);
    const resolvedUsers = await resolveCoOrganizers(client, coOrganizers, organizerEmail);
    const { allowedPresenters, participants } = buildParticipantsPayload(resolvedUsers);

    const meetingPayload = {
        subject: subject,
        startDateTime: startDateTime,
        endDateTime: endDateTime,
        allowedPresenters: allowedPresenters,
        lobbyBypassSettings: {
            scope: lobbyBypassScope
        },
        participants: participants,
        isEntryExitAnnounced: isEntryExitAnnounced,
        autoAdmittedUsers: autoAdmittedUsers
    };

    console.log(`📝 Creating meeting with ${resolvedUsers.length} co-organizers`);

    const result = await client
        .api(`/users/${organizer.id}/onlineMeetings`)
        .header('Prefer', COORGANIZER_PREFER_HEADER)
        .post(meetingPayload);

    return {
        teamsMeetingId: result.id,
        joinUrl: result.joinWebUrl
    };
}

// UPDATE CO-ORGANIZERS ONLY - This is the working PATCH approach
async function updateCoOrganizersOnly({ organizerEmail, teamsMeetingId, coOrganizers = [] }) {
    if (!teamsMeetingId) {
        console.warn('⚠️ No teamsMeetingId provided, skipping update');
        return;
    }

    const client = await getGraphClient();
    const organizer = await getUserGuidByEmail(client, organizerEmail);
    const resolvedUsers = await resolveCoOrganizers(client, coOrganizers, organizerEmail);
    const { allowedPresenters, participants } = buildParticipantsPayload(resolvedUsers);

    // ONLY update participants and allowedPresenters
    const patchPayload = {
        participants: participants,
        allowedPresenters: allowedPresenters
    };

    console.log(`📝 Updating co-organizers only for meeting ${teamsMeetingId}`);

    const result = await client
        .api(`/users/${organizer.id}/onlineMeetings/${teamsMeetingId}`)
        .header('Prefer', COORGANIZER_PREFER_HEADER)
        .patch(patchPayload);

    console.log(`✅ Co-organizers updated successfully`);
    return result;
}

// UPDATE MEETING DETAILS (subject, times) - Separate from co-organizers
async function updateMeetingDetails({ organizerEmail, teamsMeetingId, subject, startDateTime, endDateTime }) {
    if (!teamsMeetingId) return;

    const client = await getGraphClient();
    const organizer = await getUserGuidByEmail(client, organizerEmail);

    const patchPayload = {
        subject: subject,
        startDateTime: startDateTime,
        endDateTime: endDateTime
    };

    console.log(`📝 Updating meeting details for ${teamsMeetingId}`);

    const result = await client
        .api(`/users/${organizer.id}/onlineMeetings/${teamsMeetingId}`)
        .patch(patchPayload);

    return result;
}

// UPDATE MEETING - Combined but with proper separation
async function updateTeamsMeeting({ 
    organizerEmail, 
    teamsMeetingId, 
    subject, 
    startDateTime, 
    endDateTime, 
    coOrganizers = []
}) {
    if (!teamsMeetingId) return;

    try {
        // Update co-organizers first (works with v1.0)
        if (coOrganizers && coOrganizers.length > 0) {
            await updateCoOrganizersOnly({ organizerEmail, teamsMeetingId, coOrganizers });
        }
        
        // Then update basic details (works with v1.0)
        if (subject || startDateTime || endDateTime) {
            await updateMeetingDetails({ organizerEmail, teamsMeetingId, subject, startDateTime, endDateTime });
        }
        
        console.log(`✅ Meeting ${teamsMeetingId} updated successfully`);
        
    } catch (error) {
        console.error(`❌ Failed to update meeting: ${error.message}`);
        throw error;
    }
}

// GET MEETING
async function getTeamsMeeting({ organizerEmail, teamsMeetingId }) {
    const client = await getGraphClient();
    const organizer = await getUserGuidByEmail(client, organizerEmail);

    const result = await client
        .api(`/users/${organizer.id}/onlineMeetings/${teamsMeetingId}`)
        .header('Prefer', COORGANIZER_PREFER_HEADER)
        .get();
    
    const attendees = result?.participants?.attendees || [];
    const coOrganizers = attendees.filter(a => a.role === 'coorganizer');
    console.log(`📋 Meeting ${teamsMeetingId} has ${coOrganizers.length} co-organizers`);
    
    return result;
}

// DELETE MEETING
async function deleteTeamsMeeting({ organizerEmail, teamsMeetingId }) {
    if (!teamsMeetingId) return;
    const client = await getGraphClient();
    const organizer = await getUserGuidByEmail(client, organizerEmail);

    await client
        .api(`/users/${organizer.id}/onlineMeetings/${teamsMeetingId}`)
        .delete();
}

// GET MEETING PROPERTIES - Useful for debugging
async function getMeetingProperties({ organizerEmail, teamsMeetingId }) {
    const meeting = await getTeamsMeeting({ organizerEmail, teamsMeetingId });
    
    return {
        subject: meeting.subject,
        startDateTime: meeting.startDateTime,
        endDateTime: meeting.endDateTime,
        allowedPresenters: meeting.allowedPresenters,
        lobbyBypassSettings: meeting.lobbyBypassSettings,
        participants: {
            attendees: meeting.participants?.attendees || [],
            organizers: meeting.participants?.organizers || []
        },
        isEntryExitAnnounced: meeting.isEntryExitAnnounced,
        autoAdmittedUsers: meeting.autoAdmittedUsers
    };
}

module.exports = { 
    createTeamsMeeting, 
    updateTeamsMeeting, 
    getTeamsMeeting, 
    deleteTeamsMeeting,
    updateCoOrganizersOnly,
    updateMeetingDetails,
    getMeetingProperties,
    getGraphClient  // Export for scripts
};