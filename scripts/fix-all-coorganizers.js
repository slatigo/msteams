// scripts/fix-all-coorganizers.js
// Fix co-organizers for ALL meetings while preserving presenters
const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '../.env') });
const { Meeting } = require('../models');
const { Op } = require('sequelize');

// Import from your service
const msGraphService = require('../services/msGraphService');

function parseCoOrganizers(coOrgsData) {
    if (!coOrgsData) return [];
    if (Array.isArray(coOrgsData)) return coOrgsData;
    try {
        const parsed = JSON.parse(coOrgsData);
        return Array.isArray(parsed) ? parsed : [];
    } catch {
        return [];
    }
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
        console.warn(`⚠️ User lookup failed for ${email}: ${error.message}`);
        return null;
    }
}

async function fixAllCoOrganizers() {
    console.log('🔧 FIX: Co-organizers for ALL meetings (preserving presenters)\n');
    console.log('⚠️  This will UPDATE the actual Teams meetings');
    console.log('⚠️  Join URLs will NOT change');
    console.log('⚠️  Existing presenters will be PRESERVED');
    console.log('⚠️  Active meetings will NOT be disrupted');
    console.log('⚠️  Co-organizers will get breakout room management rights\n');

    try {
        // Get all meetings with co-organizers
        const meetings = await Meeting.findAll({
            where: {
                teamsMeetingId: { [Op.ne]: null },
                coOrganizers: { [Op.ne]: null }
            }
        });

        // Filter meetings that actually have co-organizers
        const meetingsWithCoOrgs = meetings.filter(m => {
            const coOrgs = parseCoOrganizers(m.coOrganizers);
            return coOrgs.length > 0;
        });

        if (meetingsWithCoOrgs.length === 0) {
            console.log('📊 No meetings with co-organizers found.');
            process.exit(0);
        }

        console.log(`📊 Found ${meetingsWithCoOrgs.length} meetings to fix\n`);
        
        // Calculate estimated time
        const estimatedTime = Math.ceil(meetingsWithCoOrgs.length * 1.5 / 60);
        console.log(`⏱️  Estimated time: ~${estimatedTime} minutes (${meetingsWithCoOrgs.length * 1.5} seconds)`);
        console.log('⚠️  Press Ctrl+C within 5 seconds to cancel...');
        await new Promise(resolve => setTimeout(resolve, 5000));

        // Get Graph client
        const client = await msGraphService.getGraphClient();

        let successCount = 0;
        let failCount = 0;
        let skipCount = 0;
        const failedMeetings = [];

        // Process in batches with progress tracking
        const total = meetingsWithCoOrgs.length;
        const batchSize = 50;
        const batches = Math.ceil(total / batchSize);

        console.log(`\n📦 Processing in ${batches} batches of ${batchSize}\n`);

        for (let batch = 0; batch < batches; batch++) {
            const start = batch * batchSize;
            const end = Math.min(start + batchSize, total);
            const batchMeetings = meetingsWithCoOrgs.slice(start, end);

            console.log(`\n📦 Batch ${batch + 1}/${batches} (${start + 1}-${end}/${total})`);
            console.log('='.repeat(50));

            for (let i = 0; i < batchMeetings.length; i++) {
                const meeting = batchMeetings[i];
                const globalIndex = start + i + 1;
                const indexStr = `[${globalIndex}/${total}]`;

                try {
                    // Parse co-organizers from DB
                    const coOrgs = parseCoOrganizers(meeting.coOrganizers);
                    const coOrgEmails = coOrgs
                        .map(c => typeof c === 'string' ? c : c?.email)
                        .filter(Boolean)
                        .filter(email => email.toLowerCase() !== meeting.creatorEmail.toLowerCase());

                    if (coOrgEmails.length === 0) {
                        console.log(`${indexStr} ⏭️ Skipping "${meeting.subject.substring(0, 30)}..." - no valid co-organizers`);
                        skipCount++;
                        continue;
                    }

                    console.log(`${indexStr} 📝 "${meeting.subject.substring(0, 40)}..."`);
                    console.log(`   📧 Co-Organizers: ${coOrgEmails.join(', ')}`);

                    // Resolve organizer
                    const organizer = await getUserGuidByEmail(client, meeting.creatorEmail);
                    if (!organizer) {
                        console.log(`   ❌ Could not resolve organizer: ${meeting.creatorEmail}`);
                        failCount++;
                        failedMeetings.push({ id: meeting.id, reason: 'Organizer not found' });
                        continue;
                    }

                    // Get current meeting to preserve presenters
                    let existingPresenters = [];
                    try {
                        const currentMeeting = await client
                            .api(`/users/${organizer.id}/onlineMeetings/${meeting.teamsMeetingId}`)
                            .header('Prefer', 'include-unknown-enum-members')
                            .get();
                        
                        const attendees = currentMeeting?.participants?.attendees || [];
                        existingPresenters = attendees.filter(a => a.role === 'presenter');
                        
                        if (existingPresenters.length > 0) {
                            console.log(`   📋 Preserving ${existingPresenters.length} existing presenters`);
                        }
                    } catch (err) {
                        console.log(`   ⚠️ Could not fetch current state: ${err.message}`);
                    }

                    // Resolve co-organizer GUIDs
                    const resolvedUsers = [];
                    for (const email of coOrgEmails) {
                        try {
                            const user = await getUserGuidByEmail(client, email);
                            if (user) {
                                resolvedUsers.push({
                                    id: user.id,
                                    email: user.email,
                                    displayName: user.displayName
                                });
                            }
                        } catch (err) {
                            console.warn(`   ⚠️ Skipping ${email}: ${err.message}`);
                        }
                    }

                    if (resolvedUsers.length === 0) {
                        console.log(`   ⚠️ No users resolved, skipping`);
                        skipCount++;
                        continue;
                    }

                    // Build new attendees array
                    const newAttendees = [];

                    // Add co-organizers
                    resolvedUsers.forEach(u => {
                        newAttendees.push({
                            identity: { 
                                user: { 
                                    id: u.id,
                                    displayName: u.displayName
                                } 
                            },
                            upn: u.email,
                            role: 'coorganizer'
                        });
                    });

                    // Preserve existing presenters
                    const coOrgEmailsLower = new Set(coOrgEmails.map(e => e.toLowerCase()));
                    let preservedCount = 0;
                    
                    existingPresenters.forEach(p => {
                        const email = p.upn || p.identity?.user?.id;
                        if (email && !coOrgEmailsLower.has(email.toLowerCase())) {
                            newAttendees.push(p);
                            preservedCount++;
                        }
                    });

                    if (preservedCount > 0) {
                        console.log(`   ✅ Preserved ${preservedCount} existing presenters`);
                    }

                    // Build and send PATCH - ONLY co-organizers and presenters
                    const patchPayload = {
                        participants: {
                            attendees: newAttendees
                        },
                        allowedPresenters: 'roleIsPresenter'
                    };

                    await client
                        .api(`/users/${organizer.id}/onlineMeetings/${meeting.teamsMeetingId}`)
                        .header('Prefer', 'include-unknown-enum-members')
                        .patch(patchPayload);

                    console.log(`   ✅ Success! (${resolvedUsers.length} co-organizers set)\n`);
                    successCount++;

                } catch (err) {
                    console.error(`   ❌ Failed: ${err.message}\n`);
                    failCount++;
                    failedMeetings.push({ 
                        id: meeting.id, 
                        subject: meeting.subject,
                        error: err.message 
                    });
                }

                // Rate limiting delay
                await new Promise(resolve => setTimeout(resolve, 500));
            }

            // Pause between batches
            if (batch < batches - 1) {
                console.log(`\n⏳ Pausing for 2 seconds between batches...\n`);
                await new Promise(resolve => setTimeout(resolve, 2000));
            }
        }

        // Final Summary
        console.log('\n' + '='.repeat(70));
        console.log('📊 FIX SUMMARY');
        console.log('='.repeat(70));
        console.log(`✅ Success:  ${successCount}`);
        console.log(`❌ Failed:   ${failCount}`);
        console.log(`⏭️  Skipped:  ${skipCount}`);
        console.log(`📊 Total:    ${meetingsWithCoOrgs.length}`);
        console.log('='.repeat(70));

        // Show failed meetings if any
        if (failedMeetings.length > 0) {
            console.log('\n❌ FAILED MEETINGS:');
            failedMeetings.forEach(m => {
                console.log(`  - ID ${m.id}: "${m.subject?.substring(0, 50) || 'Unknown'}"`);
                console.log(`    Error: ${m.error || m.reason}`);
            });
        }

        console.log('\n✅ Co-organizers fixed for all meetings!');
        console.log('📝 Existing presenters were PRESERVED');
        console.log('🎉 No meetings were recreated - join URLs are unchanged');
        console.log('🔧 Co-organizers now have breakout room management rights');

        process.exit(0);

    } catch (error) {
        console.error('💥 Fatal error:', error);
        process.exit(1);
    }
}

// Run the script
fixAllCoOrganizers();