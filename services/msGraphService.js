const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const dns = require('dns');

// Force IPv4 resolution to prevent Node.js v22 network fetch failures with Azure
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

/**
 * Resolves an email/UPN to an Azure AD User Object ID (GUID)
 */
async function getUserGuidByEmail(client, email) {
    try {
        const user = await client.api(`/users/${email}`).select('id').get();
        return user.id;
    } catch (error) {
        throw new Error(`Azure AD user lookup failed for ${email}: ${error.message}`);
    }
}

/**
 * Creates an MS Teams Online Meeting with optional Co-Organizers
 */
async function createTeamsMeeting({ organizerEmail, subject, startDateTime, endDateTime, coOrganizers = [] }) {
    const client = await getGraphClient();

    // 1. Resolve Organizer UPN to Azure AD Object ID (GUID)
    const organizerGuid = await getUserGuidByEmail(client, organizerEmail);

    const meetingPayload = {
        subject: subject,
        startDateTime: startDateTime,
        endDateTime: endDateTime,
        allowedPresenters: 'everyone',
        lobbyBypassSettings: {
            scope: 'organization'
        }
    };

    // 2. Resolve Co-Organizers to Azure AD Object IDs if present
    if (Array.isArray(coOrganizers) && coOrganizers.length > 0) {
        const filteredEmails = coOrganizers.filter(email => email !== organizerEmail);
        
        const attendeePromises = filteredEmails.map(async (email) => {
            try {
                const userGuid = await getUserGuidByEmail(client, email);
                return {
                    identity: {
                        user: { id: userGuid }
                    },
                    role: 'coorganizer'
                };
            } catch (err) {
                console.warn(`Skipping co-organizer ${email}: ${err.message}`);
                return null;
            }
        });

        const attendees = (await Promise.all(attendeePromises)).filter(Boolean);

        if (attendees.length > 0) {
            meetingPayload.participants = { attendees };
        }
    }

    // 3. Create meeting using the organizer GUID
    const result = await client
        .api(`/users/${organizerGuid}/onlineMeetings`)
        .post(meetingPayload);

    return {
        teamsMeetingId: result.id,
        joinUrl: result.joinWebUrl
    };
}

/**
 * Deletes an MS Teams Online Meeting by ID
 */
async function deleteTeamsMeeting({ organizerEmail, teamsMeetingId }) {
    if (!teamsMeetingId) return;
    const client = await getGraphClient();
    
    // Resolve Organizer UPN to Azure AD Object ID (GUID)
    const organizerGuid = await getUserGuidByEmail(client, organizerEmail);

    await client
        .api(`/users/${organizerGuid}/onlineMeetings/${teamsMeetingId}`)
        .delete();
}

module.exports = { createTeamsMeeting, deleteTeamsMeeting };