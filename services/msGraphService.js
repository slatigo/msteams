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
        const user = await client.api(`/users/${email}`).select('id').get();
        return user.id;
    } catch (error) {
        throw new Error(`Azure AD user lookup failed for ${email}: ${error.message}`);
    }
}

// Utility to clean emails and filter co-organizers
async function resolveCoOrganizers(client, coOrganizers, organizerEmail) {
    if (!Array.isArray(coOrganizers)) return [];

    const filteredEmails = coOrganizers
        .map(c => {
            const rawEmail = typeof c === 'string' ? c : c?.email;
            // Sanitizes spaces (e.g. "gkituyi @mubs.ac.ug" -> "gkituyi@mubs.ac.ug")
            return rawEmail ? rawEmail.replace(/\s+/g, '').toLowerCase() : null;
        })
        .filter(email => email && email !== organizerEmail.toLowerCase());

    const coOrgPromises = filteredEmails.map(async (email) => {
        try {
            const userGuid = await getUserGuidByEmail(client, email);
            return {
                identity: {
                    user: { id: userGuid }
                },
                upn: email,
                role: 'coOrganizer'
            };
        } catch (err) {
            console.warn(`Skipping co-organizer ${email}: ${err.message}`);
            return null;
        }
    });

    return (await Promise.all(coOrgPromises)).filter(Boolean);
}

async function createTeamsMeeting({ organizerEmail, subject, startDateTime, endDateTime, coOrganizers = [] }) {
    const client = await getGraphClient();
    const organizerGuid = await getUserGuidByEmail(client, organizerEmail);

    const coOrganizersArray = await resolveCoOrganizers(client, coOrganizers, organizerEmail);

    const attendeesArray = [
        {
            identity: { user: { id: organizerGuid } },
            upn: organizerEmail,
            role: 'presenter'
        }
    ];

    const meetingPayload = {
        subject: subject,
        startDateTime: startDateTime,
        endDateTime: endDateTime,
        allowedPresenters: 'roleIsPresenter',
        lobbyBypassSettings: {
            scope: 'organization' // Options: 'organization' (org users bypass) or 'everyone' (all bypass)
        },
        participants: {
            attendees: attendeesArray,
            coOrganizers: coOrganizersArray
        }
    };

    const result = await client
        .api(`/users/${organizerGuid}/onlineMeetings`)
        .post(meetingPayload);

    return {
        teamsMeetingId: result.id,
        joinUrl: result.joinWebUrl
    };
}

async function updateTeamsMeeting({ organizerEmail, teamsMeetingId, subject, startDateTime, endDateTime, coOrganizers = [] }) {
    if (!teamsMeetingId) return;
    const client = await getGraphClient();
    const organizerGuid = await getUserGuidByEmail(client, organizerEmail);

    const coOrganizersArray = await resolveCoOrganizers(client, coOrganizers, organizerEmail);

    const attendeesArray = [
        {
            identity: { user: { id: organizerGuid } },
            upn: organizerEmail,
            role: 'presenter'
        }
    ];

    const patchPayload = {
        subject: subject,
        startDateTime: startDateTime,
        endDateTime: endDateTime,
        allowedPresenters: 'roleIsPresenter',
        // 🔓 Explicitly sets lobby settings on PATCH so students aren't blocked
        lobbyBypassSettings: {
            scope: 'organization' // Change to 'everyone' if guests/external emails need auto-entry
        },
        participants: {
            attendees: attendeesArray,
            coOrganizers: coOrganizersArray
        }
    };

    await client
        .api(`/users/${organizerGuid}/onlineMeetings/${teamsMeetingId}`)
        .patch(patchPayload);
}

async function deleteTeamsMeeting({ organizerEmail, teamsMeetingId }) {
    if (!teamsMeetingId) return;
    const client = await getGraphClient();
    const organizerGuid = await getUserGuidByEmail(client, organizerEmail);

    await client
        .api(`/users/${organizerGuid}/onlineMeetings/${teamsMeetingId}`)
        .delete();
}

module.exports = { createTeamsMeeting, updateTeamsMeeting, deleteTeamsMeeting };