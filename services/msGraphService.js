const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const dns = require('dns');

// Force IPv4 resolution to prevent Node.js network fetch failures with Azure
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
 * Creates an MS Teams Online Meeting with Co-Organizers and secure permissions
 */
async function createTeamsMeeting({ organizerEmail, subject, startDateTime, endDateTime, coOrganizers = [] }) {
    const client = await getGraphClient();
    const organizerGuid = await getUserGuidByEmail(client, organizerEmail);

    const meetingPayload = {
        subject: subject,
        startDateTime: startDateTime,
        endDateTime: endDateTime,
        allowedPresenters: 'roleIsPresenter', // 🔒 Restricts Presenter role to Organizer + Co-Organizers
        lobbyBypassSettings: {
            scope: 'organization'
        }
    };

    if (Array.isArray(coOrganizers) && coOrganizers.length > 0) {
        const filteredEmails = coOrganizers.filter(email => email.toLowerCase() !== organizerEmail.toLowerCase());
        
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

        const coOrganizersArray = (await Promise.all(coOrgPromises)).filter(Boolean);

        if (coOrganizersArray.length > 0) {
            meetingPayload.participants = {
                coOrganizers: coOrganizersArray
            };
        }
    }

    const result = await client
        .api(`/users/${organizerGuid}/onlineMeetings`)
        .post(meetingPayload);

    return {
        teamsMeetingId: result.id,
        joinUrl: result.joinWebUrl
    };
}

/**
 * Updates an existing MS Teams Online Meeting
 */
async function updateTeamsMeeting({ organizerEmail, teamsMeetingId, subject, startDateTime, endDateTime, coOrganizers = [] }) {
    if (!teamsMeetingId) return;
    const client = await getGraphClient();
    const organizerGuid = await getUserGuidByEmail(client, organizerEmail);

    const patchPayload = {
        subject: subject,
        startDateTime: startDateTime,
        endDateTime: endDateTime,
        allowedPresenters: 'roleIsPresenter'
    };

    if (Array.isArray(coOrganizers)) {
        const filteredEmails = coOrganizers.filter(email => email.toLowerCase() !== organizerEmail.toLowerCase());
        
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

        const coOrganizersArray = (await Promise.all(coOrgPromises)).filter(Boolean);

        patchPayload.participants = {
            coOrganizers: coOrganizersArray
        };
    }

    await client
        .api(`/users/${organizerGuid}/onlineMeetings/${teamsMeetingId}`)
        .patch(patchPayload);
}

/**
 * Deletes an MS Teams Online Meeting by ID
 */
async function deleteTeamsMeeting({ organizerEmail, teamsMeetingId }) {
    if (!teamsMeetingId) return;
    const client = await getGraphClient();
    const organizerGuid = await getUserGuidByEmail(client, organizerEmail);

    await client
        .api(`/users/${organizerGuid}/onlineMeetings/${teamsMeetingId}`)
        .delete();
}

module.exports = { createTeamsMeeting, updateTeamsMeeting, deleteTeamsMeeting };