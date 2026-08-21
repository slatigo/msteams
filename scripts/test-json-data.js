// scripts/quick-update-meetings.js
// Quick update for specific meetings from JSON export
const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '../.env') });
const { updateCoOrganizersOnly } = require('../services/msGraphService');

// Meeting data from JSON export
const meetingsData = [
    {
        "id": "147",
        "subjectCode": "M2020",
        "subject": "COURSE LEADER'S TRAINING",
        "description": null,
        "meetingDate": "2026-08-10",
        "startTime": "17:00",
        "endTime": "19:00",
        "joinUrl": "https:\/\/teams.microsoft.com\/l\/meetup-join\/19%3ameeting_MmRhOWQ1M2MtMmQ5YS00YmZmLWE4ODctOWEzM2FmMWY1MWM0%40thread.v2\/0?context=%7b%22Tid%22%3a%22e9220e78-c793-4150-b529-c9bbb0e979d3%22%2c%22Oid%22%3a%2288d43324-9695-4190-9a1b-38a4b3cb395f%22%7d",
        "teamsMeetingId": "MSo4OGQ0MzMyNC05Njk1LTQxOTAtOWExYi0zOGE0YjNjYjM5NWYqMCoqMTk6bWVldGluZ19NbVJoT1dRMU0yTXRNbVE1WVMwMFltWm1MV0U0T0RjdE9XRXpNMkZtTVdZMU1XTTBAdGhyZWFkLnYy",
        "creatorEmail": "emigadde@mubs.ac.ug",
        "createdAt": "2026-08-08 13:02:38",
        "updatedAt": "2026-08-08 13:02:38",
        "expiryDate": "2026-08-10",
        "coOrganizers": "[{\"name\":\"SSENDI SAMUEL\",\"email\":\"sssendi@mubs.ac.ug\"},{\"name\":\"Migadde Elias\",\"email\":\"emigadde@mubs.ac.ug\"}]"
    },
    {
        "id": "212",
        "subjectCode": "BUC3126",
        "subject": "BUC3126 - Information Security and Auditing (ISA) Virtual Classroom (Group A)",
        "description": null,
        "meetingDate": "2026-08-08",
        "startTime": "08:00",
        "endTime": "12:00",
        "joinUrl": "https:\/\/teams.microsoft.com\/l\/meetup-join\/19%3ameeting_MjYwZmUxOTMtNTE3Yy00ZjM5LTg5NGItMjBlODg2YjYzZmEw%40thread.v2\/0?context=%7b%22Tid%22%3a%22e9220e78-c793-4150-b529-c9bbb0e979d3%22%2c%22Oid%22%3a%2252195640-ba1b-4c1c-8eb1-a7092c8e73e6%22%7d",
        "teamsMeetingId": "MSo1MjE5NTY0MC1iYTFiLTRjMWMtOGViMS1hNzA5MmM4ZTczZTYqMCoqMTk6bWVldGluZ19Nall3Wm1VeE9UTXROVEUzWXkwMFpqTTVMVGc1TkdJdE1qQmxPRGcyWWpZelptRXdAdGhyZWFkLnYy",
        "creatorEmail": "bbc3@mubs.ac.ug",
        "createdAt": "2026-08-08 14:10:41",
        "updatedAt": "2026-08-08 14:10:41",
        "expiryDate": "2026-11-30",
        "coOrganizers": "[{\"name\":\"STELLA KYALIMPA\",\"email\":\"skyalimpa@mubs.ac.ug\"},{\"name\":\"Bryan Lugemwa\",\"email\":\"blugemwa@mubs.ac.ug\"},{\"name\":\"Robinah Nabafu\",\"email\":\"rnabafu@mubs.ac.ug\"},{\"name\":\"Samuel Ssendi\",\"email\":\"sssendi@mubs.ac.ug\"}]"
    },
    {
        "id": "214",
        "subjectCode": "BUC3126",
        "subject": "BUC3126 - Information Security and Auditing (ISA) Virtual Classroom (Group B)",
        "description": null,
        "meetingDate": "2026-08-08",
        "startTime": "13:00",
        "endTime": "17:00",
        "joinUrl": "https:\/\/teams.microsoft.com\/l\/meetup-join\/19%3ameeting_ZGRmNjg0ZmMtMzM1YS00ODdmLWFjNDEtNmUzMjhlMzgwZTMw%40thread.v2\/0?context=%7b%22Tid%22%3a%22e9220e78-c793-4150-b529-c9bbb0e979d3%22%2c%22Oid%22%3a%2252195640-ba1b-4c1c-8eb1-a7092c8e73e6%22%7d",
        "teamsMeetingId": "MSo1MjE5NTY0MC1iYTFiLTRjMWMtOGViMS1hNzA5MmM4ZTczZTYqMCoqMTk6bWVldGluZ19aR1JtTmpnMFptTXRNek0xWVMwME9EZG1MV0ZqTkRFdE5tVXpNamhsTXpnd1pUTXdAdGhyZWFkLnYy",
        "creatorEmail": "bbc3@mubs.ac.ug",
        "createdAt": "2026-08-08 14:11:22",
        "updatedAt": "2026-08-08 14:11:22",
        "expiryDate": "2026-11-30",
        "coOrganizers": "[{\"name\":\"STELLA KYALIMPA\",\"email\":\"skyalimpa@mubs.ac.ug\"},{\"name\":\"Bryan Lugemwa\",\"email\":\"blugemwa@mubs.ac.ug\"},{\"name\":\"Robinah Nabafu\",\"email\":\"rnabafu@mubs.ac.ug\"},{\"name\":\"Samuel Ssendi\",\"email\":\"sssendi@mubs.ac.ug\"}]"
    },
    {
        "id": "215",
        "subjectCode": "BUC3126",
        "subject": "BUC3126 - Information Security and Auditing (ISA) Virtual Classroom (Group C)",
        "description": null,
        "meetingDate": "2026-08-08",
        "startTime": "17:30",
        "endTime": "21:30",
        "joinUrl": "https:\/\/teams.microsoft.com\/l\/meetup-join\/19%3ameeting_YjEwOTAyNWMtYzcxYS00ZTAzLWIwOWMtMDcwMGZlMTdlM2Mz%40thread.v2\/0?context=%7b%22Tid%22%3a%22e9220e78-c793-4150-b529-c9bbb0e979d3%22%2c%22Oid%22%3a%2252195640-ba1b-4c1c-8eb1-a7092c8e73e6%22%7d",
        "teamsMeetingId": "MSo1MjE5NTY0MC1iYTFiLTRjMWMtOGViMS1hNzA5MmM4ZTczZTYqMCoqMTk6bWVldGluZ19ZakV3T1RBeU5XTXRZemN4WVMwMFpUQXpMV0l3T1dNdE1EY3dNR1psTVRkbE0yTXpAdGhyZWFkLnYy",
        "creatorEmail": "bbc3@mubs.ac.ug",
        "createdAt": "2026-08-08 14:12:02",
        "updatedAt": "2026-08-08 14:12:02",
        "expiryDate": "2026-11-30",
        "coOrganizers": "[{\"name\":\"STELLA KYALIMPA\",\"email\":\"skyalimpa@mubs.ac.ug\"},{\"name\":\"Bryan Lugemwa\",\"email\":\"blugemwa@mubs.ac.ug\"},{\"name\":\"Robinah Nabafu\",\"email\":\"rnabafu@mubs.ac.ug\"},{\"name\":\"Samuel Ssendi\",\"email\":\"sssendi@mubs.ac.ug\"}]"
    },
    {
        "id": "382",
        "subjectCode": "DBD 2106",
        "subject": "DBD 2106 - BIG DATA AND CLOUD COMPUTING Virtual Classroom",
        "description": null,
        "meetingDate": "2026-08-09",
        "startTime": "11:30",
        "endTime": "23:59",
        "joinUrl": "https:\/\/teams.microsoft.com\/l\/meetup-join\/19%3ameeting_Y2RmMzcxYTUtOTI1ZC00NmYwLTgyMmUtZTliYjhiM2U4ODQx%40thread.v2\/0?context=%7b%22Tid%22%3a%22e9220e78-c793-4150-b529-c9bbb0e979d3%22%2c%22Oid%22%3a%2256076804-efaf-4a6b-a08b-83ef19468f66%22%7d",
        "teamsMeetingId": "MSo1NjA3NjgwNC1lZmFmLTRhNmItYTA4Yi04M2VmMTk0NjhmNjYqMCoqMTk6bWVldGluZ19ZMlJtTXpjeFlUVXRPVEkxWkMwME5tWXdMVGd5TW1VdFpUbGlZamhpTTJVNE9EUXhAdGhyZWFkLnYy",
        "creatorEmail": "dbida@mubs.ac.ug",
        "createdAt": "2026-08-09 11:31:10",
        "updatedAt": "2026-08-18 19:49:18",
        "expiryDate": "2026-12-12",
        "coOrganizers": "[{\"name\":\"Benedict Ogot\",\"email\":\"bogot@mubs.ac.ug\"},{\"name\":\"Samuel Ssendi\",\"email\":\"sssendi@mubs.ac.ug\"},{\"name\":\"Hillary Mirember Nagawa\",\"email\":\"hnagawa@mubs.ac.ug\"}]"
    },
    {
        "id": "658",
        "subjectCode": "BUC3128",
        "subject": "BUC3128 - Routing and Switching (RS) Virtual Classroom (Group A)",
        "description": null,
        "meetingDate": "2026-08-10",
        "startTime": "10:37",
        "endTime": "23:59",
        "joinUrl": "https:\/\/teams.microsoft.com\/l\/meetup-join\/19%3ameeting_ZDlmMDhjYTMtNDA5MS00ZWNhLTkxNTEtZmYwMmY2MDYxMjQw%40thread.v2\/0?context=%7b%22Tid%22%3a%22e9220e78-c793-4150-b529-c9bbb0e979d3%22%2c%22Oid%22%3a%2252195640-ba1b-4c1c-8eb1-a7092c8e73e6%22%7d",
        "teamsMeetingId": "MSo1MjE5NTY0MC1iYTFiLTRjMWMtOGViMS1hNzA5MmM4ZTczZTYqMCoqMTk6bWVldGluZ19aRGxtTURoallUTXROREE1TVMwMFpXTmhMVGt4TlRFdFptWXdNbVkyTURZeE1qUXdAdGhyZWFkLnYy",
        "creatorEmail": "bbc3@mubs.ac.ug",
        "createdAt": "2026-08-10 10:37:38",
        "updatedAt": "2026-08-10 10:37:38",
        "expiryDate": "2026-12-12",
        "coOrganizers": "[{\"name\":\"Samuel Ssendi\",\"email\":\"sssendi@mubs.ac.ug\"},{\"name\":\"Dr. Abdul Ssentumbwe Male\",\"email\":\"assentumbwe@mubs.ac.ug\"}]"
    },
    {
        "id": "659",
        "subjectCode": "BUC3128",
        "subject": "BUC3128 - Routing and Switching (RS) Virtual Classroom (Group B)",
        "description": null,
        "meetingDate": "2026-08-10",
        "startTime": "10:37",
        "endTime": "23:59",
        "joinUrl": "https:\/\/teams.microsoft.com\/l\/meetup-join\/19%3ameeting_NDUyODk2M2UtMGQ0Yi00NjYwLTk4YzUtNzM0YjRhMjEzMDM1%40thread.v2\/0?context=%7b%22Tid%22%3a%22e9220e78-c793-4150-b529-c9bbb0e979d3%22%2c%22Oid%22%3a%2252195640-ba1b-4c1c-8eb1-a7092c8e73e6%22%7d",
        "teamsMeetingId": "MSo1MjE5NTY0MC1iYTFiLTRjMWMtOGViMS1hNzA5MmM4ZTczZTYqMCoqMTk6bWVldGluZ19ORFV5T0RrMk0yVXRNR1EwWWkwME5qWXdMVGs0WXpVdE56TTBZalJoTWpFek1ETTFAdGhyZWFkLnYy",
        "creatorEmail": "bbc3@mubs.ac.ug",
        "createdAt": "2026-08-10 10:38:37",
        "updatedAt": "2026-08-10 10:38:37",
        "expiryDate": "2026-12-12",
        "coOrganizers": "[{\"name\":\"Samuel Ssendi\",\"email\":\"sssendi@mubs.ac.ug\"},{\"name\":\"Dr. Abdul Ssentumbwe Male\",\"email\":\"assentumbwe@mubs.ac.ug\"}]"
    },
    {
        "id": "660",
        "subjectCode": "BUC3128",
        "subject": "BUC3128 - Routing and Switching (RS) Virtual Classroom (Group C)",
        "description": null,
        "meetingDate": "2026-08-10",
        "startTime": "10:38",
        "endTime": "23:59",
        "joinUrl": "https:\/\/teams.microsoft.com\/l\/meetup-join\/19%3ameeting_Y2VhYWQ0ZmYtMTQ2YS00YTY1LWE3MDktNDhlZGI0MzliZDM3%40thread.v2\/0?context=%7b%22Tid%22%3a%22e9220e78-c793-4150-b529-c9bbb0e979d3%22%2c%22Oid%22%3a%2252195640-ba1b-4c1c-8eb1-a7092c8e73e6%22%7d",
        "teamsMeetingId": "MSo1MjE5NTY0MC1iYTFiLTRjMWMtOGViMS1hNzA5MmM4ZTczZTYqMCoqMTk6bWVldGluZ19ZMlZoWVdRMFptWXRNVFEyWVMwMFlUWTFMV0UzTURrdE5EaGxaR0kwTXpsaVpETTNAdGhyZWFkLnYy",
        "creatorEmail": "bbc3@mubs.ac.ug",
        "createdAt": "2026-08-10 10:38:58",
        "updatedAt": "2026-08-10 10:38:58",
        "expiryDate": "2026-12-12",
        "coOrganizers": "[{\"name\":\"Samuel Ssendi\",\"email\":\"sssendi@mubs.ac.ug\"},{\"name\":\"Dr. Abdul Ssentumbwe Male\",\"email\":\"assentumbwe@mubs.ac.ug\"}]"
    }
];

// Parse co-organizers from JSON string
function parseCoOrganizers(coOrgsData) {
    if (!coOrgsData) return [];
    if (Array.isArray(coOrgsData)) return coOrgsData;
    try {
        const parsed = typeof coOrgsData === 'string' ? JSON.parse(coOrgsData) : coOrgsData;
        return Array.isArray(parsed) ? parsed : [];
    } catch {
        return [];
    }
}

// Extract emails from co-organizers array
function extractEmails(coOrgs) {
    return coOrgs
        .map(c => typeof c === 'string' ? c : c?.email)
        .filter(Boolean);
}

async function quickUpdateMeetings() {
    console.log('🚀 Quick update for meetings from JSON data\n');
    console.log(`📊 Found ${meetingsData.length} meetings to update\n`);

    let successCount = 0;
    let failCount = 0;
    let skipCount = 0;

    for (let i = 0; i < meetingsData.length; i++) {
        const meeting = meetingsData[i];
        const indexStr = `[${i + 1}/${meetingsData.length}]`;

        try {
            // Parse co-organizers
            const coOrgs = parseCoOrganizers(meeting.coOrganizers);
            const coOrgEmails = extractEmails(coOrgs)
                .filter(email => email.toLowerCase() !== meeting.creatorEmail.toLowerCase());

            console.log(`${indexStr} 📝 Updating: "${meeting.subject}"`);
            console.log(`   ID: ${meeting.id}`);
            console.log(`   📧 Co-Organizers: ${coOrgEmails.join(', ') || 'None'}`);

            if (coOrgEmails.length === 0) {
                console.log(`   ⏭️ No valid co-organizers (organizer removed), skipping\n`);
                skipCount++;
                continue;
            }

            // Update co-organizers using the working PATCH approach
            await updateCoOrganizersOnly({
                organizerEmail: meeting.creatorEmail,
                teamsMeetingId: meeting.teamsMeetingId,
                coOrganizers: coOrgEmails
            });

            console.log(`   ✅ Success!\n`);
            successCount++;

        } catch (err) {
            console.error(`   ❌ Failed: ${err.message}\n`);
            failCount++;
        }

        // Small delay to avoid rate limiting
        await new Promise(resolve => setTimeout(resolve, 500));
    }

    console.log('='.repeat(60));
    console.log('📊 QUICK UPDATE SUMMARY');
    console.log('='.repeat(60));
    console.log(`✅ Success: ${successCount}`);
    console.log(`❌ Failed:  ${failCount}`);
    console.log(`⏭️  Skipped: ${skipCount}`);
    console.log(`📊 Total:   ${meetingsData.length}`);
    console.log('='.repeat(60));

    if (successCount === meetingsData.length) {
        console.log('\n🎉 All meetings updated successfully!');
    } else if (successCount > 0) {
        console.log('\n⚠️ Some meetings failed. Check the errors above.');
    } else {
        console.log('\n❌ All meetings failed. Check your configuration.');
    }

    process.exit(0);
}

// Run the update
quickUpdateMeetings();