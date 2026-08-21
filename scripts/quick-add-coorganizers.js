// scripts/graph-add-coorganizers-v3.js
// Add co-organizers directly via Graph API - with correct meeting ID
const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '../.env') });
const { getGraphClient } = require('../services/msGraphService');

// ============================================================
// ⚡ EDIT THESE VALUES - Using the correct data from your DB
// ============================================================

// 1. The FULL Teams meeting ID from your database
const TEAMS_MEETING_ID = 'MSplZDFhNDQ1NC1mNTZkLTRmMmEtOGM4Ny02MWIxOTE0ZWQwNzMqMCoqMTk6bWVldGluZ19aREptTnpFeU1tWXRZamN4T1MwME5UbGpMV0ZtTlRRdE4yVTBNVFJoT0RFMlpHRmxAdGhyZWFkLnYy';

// 2. The original organizer of the meeting (email)
const ORGANIZER_EMAIL = 'bbc2@mubs.ac.ug';

// 3. The people you want to add as co-organizers (comma separated)
const NEW_COORGANIZERS = 'gbyomire@mubs.ac.ug';

// ============================================================
// DO NOT EDIT BELOW THIS LINE
// ============================================================

// Resolve user email to GUID
async function getUserGuid(client, email) {
    try {
        const user = await client.api(`/users/${email}`).select('id,displayName,userPrincipalName').get();
        return {
            id: user.id,
            displayName: user.displayName || email,
            email: user.userPrincipalName || email
        };
    } catch (error) {
        console.error(`❌ User lookup failed for ${email}: ${error.message}`);
        return null;
    }
}

async function addCoOrganizersDirectly() {
    console.log('🚀 Adding Co-Organizers Directly via Graph API\n');
    console.log('=' .repeat(60));

    console.log(`📋 Meeting ID: ${TEAMS_MEETING_ID.substring(0, 40)}...`);
    console.log(`👤 Organizer: ${ORGANIZER_EMAIL}`);
    
    const emails = NEW_COORGANIZERS.split(',').map(e => e.trim()).filter(Boolean);
    console.log(`📧 Adding: ${emails.join(', ')}`);

    try {
        const client = await getGraphClient();

        // 1. Resolve organizer to GUID
        console.log('\n🔍 Resolving organizer to GUID...');
        const organizer = await getUserGuid(client, ORGANIZER_EMAIL);
        if (!organizer) {
            console.log('❌ Could not resolve organizer email');
            process.exit(1);
        }
        console.log(`✅ Organizer resolved: ${organizer.id}`);

        // 2. Resolve co-organizers to GUIDs
        console.log('\n🔍 Resolving co-organizers to GUIDs...');
        const resolvedUsers = [];
        for (const email of emails) {
            if (email.toLowerCase() === ORGANIZER_EMAIL.toLowerCase()) {
                console.log(`⚠️ Skipping ${email} (is the organizer)`);
                continue;
            }
            const user = await getUserGuid(client, email);
            if (user) {
                resolvedUsers.push(user);
                console.log(`✅ Resolved ${email} -> ${user.id}`);
            }
        }

        if (resolvedUsers.length === 0) {
            console.log('❌ No valid co-organizers to add');
            process.exit(1);
        }

        // 3. Build the attendees array
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

        // 4. Build the PATCH payload
        const patchPayload = {
            participants: {
                attendees: attendeesArray
            },
            allowedPresenters: 'roleIsPresenter'
        };

        console.log('\n📝 Sending PATCH request to Graph API...');
        console.log(`   Endpoint: /users/${organizer.id}/onlineMeetings/${TEAMS_MEETING_ID}`);

        // 5. Send the PATCH request
        const result = await client
            .api(`/users/${organizer.id}/onlineMeetings/${TEAMS_MEETING_ID}`)
            .header('Prefer', 'include-unknown-enum-members')
            .patch(patchPayload);

        console.log('\n✅ SUCCESS! Co-organizers added!');
        console.log(`   Added: ${resolvedUsers.map(u => u.email).join(', ')}`);

        // 6. Verify by fetching the updated meeting
        console.log('\n🔍 Verifying update...');
        const verified = await client
            .api(`/users/${organizer.id}/onlineMeetings/${TEAMS_MEETING_ID}`)
            .header('Prefer', 'include-unknown-enum-members')
            .get();

        const attendees = verified?.participants?.attendees || [];
        const coOrganizers = attendees.filter(a => a.role === 'coorganizer');
        const coOrgEmails = coOrganizers.map(a => a.upn || a.identity?.user?.id);

        console.log(`   Co-organizers in meeting: ${coOrgEmails.join(', ') || 'None'}`);

        // Also check existing co-organizers
        const existingCoOrgs = attendees.filter(a => a.role === 'coorganizer');
        if (existingCoOrgs.length > 0) {
            console.log(`\n📋 All co-organizers now in meeting:`);
            existingCoOrgs.forEach((c, i) => {
                const email = c.upn || c.identity?.user?.id;
                console.log(`   ${i + 1}. ${email}`);
            });
        }

        if (coOrgEmails.length > 0) {
            console.log('\n🎉 All done! Co-organizers are now set in the meeting.');
        } else {
            console.log('\n⚠️ Warning: No co-organizers found after update.');
        }

    } catch (error) {
        console.error('\n❌ Failed:', error.message);
        if (error.statusCode === 400) {
            console.error('📋 Error details:', error.body);
            try {
                const details = JSON.parse(error.body);
                console.error('   Message:', details.message);
            } catch (e) {}
        }
        process.exit(1);
    }

    process.exit(0);
}

addCoOrganizersDirectly();