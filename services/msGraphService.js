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

async function createTeamsMeeting({ organizerEmail, subject, startDateTime, endDateTime, coOrganizers = [] }) {
    const client = await getGraphClient();
    const organizerGuid = await getUserGuidByEmail(client, organizerEmail);

    let coOrganizersArray = [];
    if (Array.isArray(coOrganizers) && coOrganizers.length > 0) {
        const filteredEmails = coOrganizers.filter(email => email && email.toLowerCase() !== organizerEmail.toLowerCase());
        
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

        coOrganizersArray = (await Promise.all(coOrgPromises)).filter(Boolean);
    }

    // 🔒 Include Organizer in attendees list with explicit role: 'presenter'
    const attendeesArray = [
        {
            identity: {
                user: { id: organizerGuid }
            },
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
            scope: 'organization'
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

    let coOrganizersArray = [];
    if (Array.isArray(coOrganizers)) {
        const filteredEmails = coOrganizers.filter(email => email && email.toLowerCase() !== organizerEmail.toLowerCase());
        
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

        coOrganizersArray = (await Promise.all(coOrgPromises)).filter(Boolean);
    }

    // 🔒 Include Organizer in attendees list with explicit role: 'presenter'
    const attendeesArray = [
        {
            identity: {
                user: { id: organizerGuid }
            },
            upn: organizerEmail,
            role: 'presenter'
        }
    ];

    const patchPayload = {
        subject: subject,
        startDateTime: startDateTime,
        endDateTime: endDateTime,
        allowedPresenters: 'roleIsPresenter',
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