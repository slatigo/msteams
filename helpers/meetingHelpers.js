/**
 * Extract and unify co-organizers from both Moodle checkboxes and direct manual inputs.
 * Optionally filters out the meeting organizer's email.
 */
function extractCoOrganizers(body, organizerEmail = '') {
    const coOrganizers = [];
    const normalizedOrganizer = organizerEmail ? organizerEmail.trim().toLowerCase() : '';

    // 1. Process Moodle Checkbox selections
    if (body.moodleCoOrganizers) {
        const moodleItems = Array.isArray(body.moodleCoOrganizers) 
            ? body.moodleCoOrganizers 
            : [body.moodleCoOrganizers];

        moodleItems.forEach(itemStr => {
            try {
                const parsed = typeof itemStr === 'string' ? JSON.parse(itemStr) : itemStr;
                const cleanEmail = parsed.email ? parsed.email.trim() : '';
                
                // Exclude if email is missing or matches the organizer
                if (cleanEmail && cleanEmail.toLowerCase() !== normalizedOrganizer) {
                    coOrganizers.push({
                        name: parsed.name || cleanEmail,
                        email: cleanEmail
                    });
                }
            } catch (err) {
                console.error('Failed to parse Moodle co-organizer JSON:', err);
            }
        });
    }

    // 2. Process Direct Manual Entry fields
    const names = body['manualCoOrganizerNames[]'] || body.manualCoOrganizerNames || [];
    const emails = body['manualCoOrganizerEmails[]'] || body.manualCoOrganizerEmails || [];

    const nameArray = Array.isArray(names) ? names : [names];
    const emailArray = Array.isArray(emails) ? emails : [emails];

    emailArray.forEach((email, idx) => {
        const cleanEmail = email ? email.trim() : '';
        if (cleanEmail && cleanEmail.toLowerCase() !== normalizedOrganizer) {
            const rawName = nameArray[idx] ? nameArray[idx].trim() : '';
            coOrganizers.push({
                name: rawName || cleanEmail,
                email: cleanEmail
            });
        }
    });

    // Remove duplicates based on lowercased email
    const uniqueMap = new Map();
    coOrganizers.forEach(item => uniqueMap.set(item.email.toLowerCase(), item));
    return Array.from(uniqueMap.values());
}

/**
 * Fetch course participants/lecturers from Moodle REST API or Database.
 */
async function fetchMoodleLecturers(subjectCode, organizerEmail = '') {
    try {
        // Replace this block with your actual Moodle API fetch logic or DB query.
        // Returning an empty array prevents mock data like Jane Smith from appearing.
        const lecturers = []; 

        // Filter out the organizer so they cannot add themselves as a co-organizer
        const normalizedOrganizer = organizerEmail ? organizerEmail.trim().toLowerCase() : '';
        return lecturers.filter(lecturer => 
            lecturer.email && lecturer.email.trim().toLowerCase() !== normalizedOrganizer
        );
    } catch (err) {
        console.error('Error fetching Moodle lecturers:', err);
        return [];
    }
}

/**
 * Safely parse and normalize JSON coOrganizers field from DB model
 */
function parseCoOrganizersField(rawCoOrganizers) {
    if (!rawCoOrganizers) return [];
    let items = [];

    if (typeof rawCoOrganizers === 'string') {
        try {
            items = JSON.parse(rawCoOrganizers);
        } catch (e) {
            items = rawCoOrganizers.split(',').map(e => e.trim());
        }
    } else if (Array.isArray(rawCoOrganizers)) {
        items = rawCoOrganizers;
    }

    // Convert string array items ['a@b.com'] into objects [{ name: 'a@b.com', email: 'a@b.com' }]
    return items.map(item => {
        if (typeof item === 'string') {
            return { name: item, email: item };
        }
        if (item && typeof item === 'object') {
            return {
                name: item.name || item.email || '',
                email: item.email || ''
            };
        }
        return null;
    }).filter(Boolean);
}

module.exports = {
    extractCoOrganizers,
    fetchMoodleLecturers,
    parseCoOrganizersField
};