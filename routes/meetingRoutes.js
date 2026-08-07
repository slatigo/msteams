const express = require('express');
const router = express.Router();

const { Subject, Meeting } = require('../models');
const { createTeamsMeeting, updateTeamsMeeting, deleteTeamsMeeting } = require('../services/msGraphService');
const { requireAuth, requireTeamsRole } = require('../middleware/authMiddleware');
const { extractCoOrganizers, fetchMoodleLecturers, parseCoOrganizersField } = require('../helpers/meetingHelpers');

// NEW MEETING FORM
router.get('/subjects/:subject_code/meetings/new', requireAuth, requireTeamsRole, async (req, res) => {
    try {
        const subjectCode = req.params.subject_code.toUpperCase();
        const subject = await Subject.findByPk(subjectCode);
        if (!subject) return res.status(404).send('Subject not found');

        const moodleLecturers = (req.session.moodleLecturers && req.session.moodleLecturers[subjectCode]) || [];

        res.render('meeting-form', {
            user: req.session.user,
            subject: subject.get({ plain: true }),
            moodleLecturers,
            isEditMode: false,
            isMoodle: req.session.isMoodle || false,
            meeting: { coOrganizers: [] }
        });
    } catch (err) {
        res.status(500).send('Error rendering form');
    }
});

// CREATE MEETING ACTION
router.post('/api/subjects/:subject_code/meetings', requireAuth, requireTeamsRole, async (req, res) => {
    try {
        const subjectCode = req.params.subject_code.toUpperCase();
        const { subject, startDateTime, endDateTime, meetingDate, expiryDate, startTime, endTime } = req.body;
        const organizerEmail = req.session.user.email;

        const coOrganizersList = extractCoOrganizers(req.body);

        const rawStart = startDateTime || (meetingDate && startTime ? `${meetingDate}T${startTime}` : null);
        const rawEnd = endDateTime || (expiryDate && endTime ? `${expiryDate}T${endTime}` : rawStart);

        if (!rawStart) {
            return res.status(400).send('Invalid request: Start date and time are required.');
        }

        const startDateObj = new Date(rawStart);
        const endDateObj = new Date(rawEnd || rawStart);

        if (isNaN(startDateObj.getTime()) || isNaN(endDateObj.getTime())) {
            return res.status(400).send('Invalid date format received.');
        }

        const startISO = startDateObj.toISOString();
        const endISO = endDateObj.toISOString();

        const [parsedMeetingDate, parsedStartTime] = rawStart.split('T');
        const [parsedExpiryDate, parsedEndTime] = (rawEnd || rawStart).split('T');

        let realTeamsData;
        try {
            realTeamsData = await createTeamsMeeting({
                organizerEmail,
                subject,
                startDateTime: startISO,
                endDateTime: endISO,
                coOrganizers: coOrganizersList.map(c => c.email)
            });
        } catch (graphErr) {
            console.error('MS Graph API Error:', graphErr);
            return res.status(500).send('Failed to generate Microsoft Teams meeting via Graph API.');
        }

        await Meeting.create({
            subjectCode,
            subject,
            meetingDate: parsedMeetingDate,
            expiryDate: parsedExpiryDate || parsedMeetingDate,
            startTime: parsedStartTime ? parsedStartTime.slice(0, 5) : '08:00',
            endTime: parsedEndTime ? parsedEndTime.slice(0, 5) : '17:00',
            joinUrl: realTeamsData.joinUrl,
            teamsMeetingId: realTeamsData.teamsMeetingId,
            creatorEmail: organizerEmail,
            coOrganizers: JSON.stringify(coOrganizersList)
        });

        res.redirect(`/subjects/${subjectCode}`);
    } catch (err) {
        console.error('Create Action Error:', err);
        res.status(500).send('Error recording meeting schedule.');
    }
});

// EDIT MEETING FORM
// EDIT MEETING FORM
router.get('/subjects/:subject_code/meetings/:id/edit', requireAuth, requireTeamsRole, async (req, res) => {
    try {
        const subjectCode = req.params.subject_code.toUpperCase();
        
        // 1. Fetch Subject model to get subject.name
        const subject = await Subject.findByPk(subjectCode);
        const subjectData = subject ? subject.get({ plain: true }) : { code: subjectCode };

        const meeting = await Meeting.findByPk(req.params.id);
        if (!meeting) return res.status(404).send('Meeting not found');

        // STRICT CREATOR OR ADMIN CHECK
        const currentUserEmail = req.session.user.email.toLowerCase();
        const isCreator = meeting.creatorEmail && meeting.creatorEmail.toLowerCase() === currentUserEmail;
        const isAdmin = req.session.user.role === 'admin';

        if (!isCreator && !isAdmin) {
            return res.status(403).send('Forbidden: Only the meeting creator or an administrator can edit this meeting.');
        }

        const plainMeeting = meeting.get({ plain: true });
        plainMeeting.coOrganizers = parseCoOrganizersField(plainMeeting.coOrganizers);

        let rawLecturers = (req.session.moodleLecturers && req.session.moodleLecturers[subjectCode]) || [];
        if (rawLecturers.length === 0) {
            rawLecturers = await fetchMoodleLecturers(subjectCode, meeting.creatorEmail);
        }

        const moodleLecturers = rawLecturers.filter(
            l => l.email && l.email.toLowerCase() !== meeting.creatorEmail.toLowerCase()
        );

        res.render('meeting-form', {
            user: req.session.user,
            subject: subjectData, // Now includes code and name
            moodleLecturers,
            isEditMode: true,
            isMoodle: req.session.isMoodle || false,
            meeting: plainMeeting
        });
    } catch (err) {
        console.error('Error loading edit form:', err);
        res.status(500).send('Database Error');
    }
});

// UPDATE MEETING ACTION
router.post('/api/subjects/:subject_code/meetings/:id/edit', requireAuth, requireTeamsRole, async (req, res) => {
    try {
        const meeting = await Meeting.findByPk(req.params.id);
        if (!meeting) return res.status(404).send('Meeting not found');

        // STRICT CREATOR OR ADMIN CHECK
        const currentUserEmail = req.session.user.email.toLowerCase();
        const isCreator = meeting.creatorEmail && meeting.creatorEmail.toLowerCase() === currentUserEmail;
        const isAdmin = req.session.user.role === 'admin';

        if (!isCreator && !isAdmin) {
            return res.status(403).send('Forbidden: Only the meeting creator or an administrator can edit this meeting.');
        }

        const { subject, startDateTime, endDateTime, meetingDate, expiryDate, startTime, endTime } = req.body;
        const coOrganizersList = extractCoOrganizers(req.body);

        const rawStart = startDateTime || (meetingDate && startTime ? `${meetingDate}T${startTime}` : null);
        const rawEnd = endDateTime || (expiryDate && endTime ? `${expiryDate}T${endTime}` : rawStart);

        if (!rawStart) {
            return res.status(400).send('Invalid request: Start date and time are required.');
        }

        const startDateObj = new Date(rawStart);
        const endDateObj = new Date(rawEnd || rawStart);

        if (isNaN(startDateObj.getTime()) || isNaN(endDateObj.getTime())) {
            return res.status(400).send('Invalid date format received.');
        }

        const startISO = startDateObj.toISOString();
        const endISO = endDateObj.toISOString();

        const [parsedMeetingDate, parsedStartTime] = rawStart.split('T');
        const [parsedExpiryDate, parsedEndTime] = (rawEnd || rawStart).split('T');

        // 1. UPDATE MICROSOFT TEAMS MEETING VIA GRAPH API
        if (meeting.teamsMeetingId) {
            try {
                await updateTeamsMeeting({
                    organizerEmail: meeting.creatorEmail,
                    teamsMeetingId: meeting.teamsMeetingId,
                    subject,
                    startDateTime: startISO,
                    endDateTime: endISO,
                    coOrganizers: coOrganizersList.map(c => c.email)
                });
            } catch (graphErr) {
                console.error('MS Graph API Update Error:', graphErr.message);
            }
        }

        // 2. UPDATE LOCAL DATABASE RECORD
        await meeting.update({
            subject,
            meetingDate: parsedMeetingDate,
            expiryDate: parsedExpiryDate || parsedMeetingDate,
            startTime: parsedStartTime ? parsedStartTime.slice(0, 5) : '08:00',
            endTime: parsedEndTime ? parsedEndTime.slice(0, 5) : '17:00',
            coOrganizers: JSON.stringify(coOrganizersList)
        });

        res.redirect(`/subjects/${req.params.subject_code.toUpperCase()}`);
    } catch (err) {
        console.error('Update Action Error:', err);
        res.status(500).send('Error updating meeting details');
    }
});

// DELETE MEETING ACTION
router.delete('/api/meetings/:id', requireAuth, async (req, res) => {
    try {
        const meeting = await Meeting.findByPk(req.params.id);
        if (!meeting) return res.status(404).json({ success: false, error: 'Meeting not found' });

        // STRICT CREATOR OR ADMIN CHECK
        const currentUserEmail = req.session.user.email.toLowerCase();
        const isCreator = meeting.creatorEmail && meeting.creatorEmail.toLowerCase() === currentUserEmail;
        const isAdmin = req.session.user.role === 'admin';

        if (!isCreator && !isAdmin) {
            return res.status(403).json({ success: false, error: 'Forbidden: Only the meeting creator or an administrator can delete this meeting.' });
        }

        if (meeting.teamsMeetingId) {
            try {
                await deleteTeamsMeeting({
                    organizerEmail: meeting.creatorEmail,
                    teamsMeetingId: meeting.teamsMeetingId
                });
            } catch (graphErr) {
                console.warn('Could not delete meeting from Graph API:', graphErr.message);
            }
        }

        await meeting.destroy();
        res.json({ success: true });
    } catch (err) {
        res.status(500).json({ success: false, error: 'Database processing error' });
    }
});

module.exports = router;