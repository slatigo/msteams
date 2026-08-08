const express = require('express');
const router = express.Router();
const { Op } = require('sequelize');

const { Subject, Meeting,User } = require('../models');
const { createTeamsMeeting, updateTeamsMeeting, deleteTeamsMeeting } = require('../services/msGraphService');
const { requireAuth, requireTeamsRole } = require('../middleware/authMiddleware');
const { extractCoOrganizers, fetchMoodleLecturers, parseCoOrganizersField } = require('../helpers/meetingHelpers');
const { ALLOWED_MANAGEMENT_ROLES } = require('../config/constants');

// ==========================================
// 0. ALL MEETINGS OVERVIEW (GET /meetings)
// ==========================================

router.get('/meetings', requireAuth, async (req, res, next) => {
    try {
        const currentUserEmail = req.session.user.email.toLowerCase();
        const userRole = req.session.user.role;
        const isAdmin = userRole === 'admin';
        const isMoodle = !!req.session.isMoodle;

        // Management authorization: Strictly admin and ms_teams_account / msteam_account
        const canManage = ALLOWED_MANAGEMENT_ROLES.includes(userRole);

        const { subjectCode, creatorEmail } = req.query;

        // 1. Fetch Subject metadata if viewing a specific subject
        let subjectObj = null;
        if (subjectCode) {
            const foundSubject = await Subject.findByPk(subjectCode.toUpperCase());
            if (foundSubject) {
                subjectObj = foundSubject.get({ plain: true });
            } else {
                subjectObj = { code: subjectCode.toUpperCase(), name: subjectCode.toUpperCase() };
            }
        }

        // 2. Build Database Query Filter
        let whereClause = {};

        if (subjectCode) {
            // SUBJECT / MOODLE VIEW: Everyone sees ALL meetings for this subject
            whereClause.subjectCode = subjectCode.toUpperCase();

            // Admin can optionally filter by creator within a subject
            if (isAdmin && creatorEmail) {
                whereClause.creatorEmail = creatorEmail.toLowerCase();
            }
        } else {
            // GLOBAL ALL-MEETINGS VIEW:
            if (!isAdmin) {
                // Non-admins only see meetings they created
                whereClause.creatorEmail = currentUserEmail;
            } else if (creatorEmail) {
                whereClause.creatorEmail = creatorEmail.toLowerCase();
            }
        }

        const rawMeetings = await Meeting.findAll({
            where: whereClause,
            order: [['meetingDate', 'ASC'], ['startTime', 'ASC']]
        });

        // 3. Admin dropdown options for filter UI
        let subjectsList = [];
        let creatorsList = [];

        if (isAdmin) {
            subjectsList = await Subject.findAll({ order: [['name', 'ASC']] });
            const distinctCreators = await Meeting.findAll({
                attributes: ['creatorEmail'],
                group: ['creatorEmail'],
                raw: true
            });
            creatorsList = distinctCreators.map(m => m.creatorEmail).filter(Boolean);
        }

        const now = new Date();

        const meetings = rawMeetings.map(m => {
            const plain = m.get({ plain: true });
            const coOrganizers = parseCoOrganizersField(plain.coOrganizers);
            const isCreator = plain.creatorEmail && plain.creatorEmail.toLowerCase() === currentUserEmail;

            const dateStr = plain.expiryDate || plain.meetingDate;
            const timeStr = plain.endTime || '23:59';
            const endDateTime = new Date(`${dateStr}T${timeStr}:00`);
            const isPast = !isNaN(endDateTime.getTime()) && endDateTime < now;

            return {
                ...plain,
                coOrganizers,
                canEdit: isCreator || isAdmin,
                isPast
            };
        });

        const activeMeetings = meetings.filter(m => !m.isPast);
        const previousMeetings = meetings.filter(m => m.isPast).reverse();

        const isGlobalView = !subjectCode;

        res.render('schedule', {
            user: req.session.user,
            isAdmin,
            isMoodle,
            activePage: 'meetings',
            isGlobalView,
            subject: subjectObj,
            canManage, // True ONLY for admin & ms_teams_account / msteam_account
            sessions: meetings,
            activeMeetings,
            previousMeetings,
            subjectsList: subjectsList.map(s => s.get({ plain: true })),
            creatorsList,
            selectedSubject: subjectCode ? subjectCode.toUpperCase() : '',
            selectedCreator: creatorEmail || ''
        });
    } catch (err) {
        console.error('Fetch All Meetings Error:', err);
        next(err);
    }
});

// ==========================================
// 1. MEETING FORM RENDERING (GET)
// ==========================================

// GET /meetings/new — Unified New Meeting Form
router.get('/meetings/new', requireAuth, requireTeamsRole, async (req, res) => {
    try {
        const { subjectCode } = req.query;
        const currentUserEmail = req.session.user.email.toLowerCase();
        const subjectsList = await Subject.findAll({ order: [['name', 'ASC']] });
        const fixedEndDate = req.session.fixedEndDate || null;

        // --- FETCH CANDIDATES FOR CO-ORGANIZERS ---
        let moodleLecturers = [];

        if (req.session.isMoodle && Array.isArray(req.session.moodleLecturers) && req.session.moodleLecturers.length > 0) {
            // Moodle SSO context: Use lecturer list passed in JWT payload from Moodle
            moodleLecturers = req.session.moodleLecturers.filter(
                l => l.email && l.email.toLowerCase() !== currentUserEmail
            );
        } else {
            // Standalone Direct Login context: Query active accounts from DB
            const dbUsers = await User.findAll({
                attributes: ['id', 'name', 'email', 'role'],
                order: [['name', 'ASC']]
            });

            moodleLecturers = dbUsers
                .map(u => u.get({ plain: true }))
                .filter(u => u.email && u.email.toLowerCase() !== currentUserEmail);
        }

        let selectedSubject = null;
        if (subjectCode) {
            const foundSubject = await Subject.findByPk(subjectCode.toUpperCase());
            selectedSubject = foundSubject 
                ? foundSubject.get({ plain: true }) 
                : { code: subjectCode.toUpperCase(), name: subjectCode.toUpperCase() };
        }

        res.render('meeting-form', {
            user: req.session.user,
            isAdmin: req.session.user.role === 'admin',
            isMoodle: !!req.session.isMoodle,
            isEditMode: false,
            subject: selectedSubject,
            subjects: subjectsList.map(s => s.get({ plain: true })),
            moodleLecturers, // List of co-organizer options rendered in the form UI
            fixedEndDate,
            meeting: null,
            errorMessage: null
        });
    } catch (err) {
        console.error('New Meeting Form Error:', err);
        res.status(500).send('Error loading meeting creation form.');
    }
});

// Edit Meeting Form
router.get('/subjects/:subject_code/meetings/:id/edit', requireAuth, requireTeamsRole, async (req, res) => {
    try {
        const subjectCode = req.params.subject_code.toUpperCase();
        const meetingId = req.params.id;

        const subject = await Subject.findByPk(subjectCode);
        const meeting = await Meeting.findByPk(meetingId);

        if (!meeting || meeting.subjectCode !== subjectCode) {
            return res.status(404).send('Meeting record not found.');
        }

        const isCreator = meeting.creatorEmail && meeting.creatorEmail.toLowerCase() === req.session.user.email.toLowerCase();
        const isAdmin = req.session.user.role === 'admin';

        if (!isCreator && !isAdmin) {
            return res.status(403).send('You do not have permission to edit this meeting.');
        }

        let moodleLecturers = (req.session.moodleLecturers && req.session.moodleLecturers[subjectCode]) || [];

        if (moodleLecturers.length === 0 && req.session.isMoodle && typeof fetchMoodleLecturers === 'function') {
            try {
                const fetched = await fetchMoodleLecturers(req.session, subjectCode);
                if (Array.isArray(fetched) && fetched.length > 0) {
                    moodleLecturers = fetched;
                }
            } catch (lecturerErr) {
                console.warn('Could not fetch Moodle lecturers:', lecturerErr.message);
            }
        }

        const fixedEndDate = (req.session.moodleSettings && req.session.moodleSettings[subjectCode]?.fixed_end_date) 
            || req.session.fixedEndDate 
            || null;

        const plainMeeting = meeting.get({ plain: true });
        plainMeeting.coOrganizers = parseCoOrganizersField(plainMeeting.coOrganizers);

        res.render('meeting-form', {
            user: req.session.user,
            isAdmin,
            isMoodle: !!req.session.isMoodle,
            isEditMode: true,
            subject: subject ? subject.get({ plain: true }) : { code: subjectCode },
            moodleLecturers,
            fixedEndDate,
            meeting: plainMeeting,
            errorMessage: null
        });
    } catch (err) {
        console.error('Edit Meeting Form Error:', err);
        res.status(500).send('Error loading edit form.');
    }
});

// ==========================================
// 2. MEETING ACTIONS (POST / DELETE)
// ==========================================

// Action: Create Meeting
router.post('/api/subjects/:subject_code/meetings', requireAuth, requireTeamsRole, async (req, res) => {
    const subjectCode = req.params.subject_code.toUpperCase();
    const { subject: customSubject, description, startDateTime, endDateTime, meetingDate, expiryDate, startTime, endTime } = req.body;
    const organizerEmail = req.session.user.email;
    const isMoodle = req.session.isMoodle || false;

    const renderFormWithError = async (errorMsg) => {
        const subjectObj = await Subject.findByPk(subjectCode);
        let moodleLecturers = (req.session.moodleLecturers && req.session.moodleLecturers[subjectCode]) || [];
        if (isMoodle && fetchMoodleLecturers && moodleLecturers.length === 0) {
            try { moodleLecturers = await fetchMoodleLecturers(req.session, subjectCode); } catch (e) {}
        }
        return res.status(400).render('meeting-form', {
            user: req.session.user,
            isAdmin: req.session.user.role === 'admin',
            isMoodle,
            isEditMode: false,
            subject: subjectObj ? subjectObj.get({ plain: true }) : { code: subjectCode },
            moodleLecturers,
            meeting: req.body,
            errorMessage: errorMsg
        });
    };

    try {
        const coOrganizersList = extractCoOrganizers(req.body);
        const fixedEndDateStr = (req.session.moodleSettings && req.session.moodleSettings[subjectCode]?.fixed_end_date) 
            || req.session.fixedEndDate 
            || expiryDate 
            || (endDateTime ? endDateTime.split('T')[0] : null);

        let startDateObj;
        let endDateObj;

        if (isMoodle && fixedEndDateStr) {
            const rawStart = startDateTime || (meetingDate && startTime ? `${meetingDate}T${startTime}` : new Date().toISOString());
            startDateObj = new Date(rawStart);
            endDateObj = new Date(`${fixedEndDateStr}T23:59:59.000Z`);

            if (isNaN(endDateObj.getTime())) {
                return renderFormWithError(`Invalid fixed end date format (${fixedEndDateStr}).`);
            }
        } else {
            const rawStart = startDateTime || (meetingDate && startTime ? `${meetingDate}T${startTime}` : null);
            const rawEnd = endDateTime || (expiryDate && endTime ? `${expiryDate}T${endTime}` : rawStart);

            if (!rawStart) {
                return renderFormWithError('Start date and time are required.');
            }

            startDateObj = new Date(rawStart);
            endDateObj = new Date(rawEnd || rawStart);

            if (isNaN(startDateObj.getTime()) || isNaN(endDateObj.getTime())) {
                return renderFormWithError('Invalid date format received.');
            }
        }

        const rawStartString = startDateTime || `${meetingDate}T${startTime}`;
        const rawEndString = endDateTime || `${expiryDate}T${endTime}`;
        const [pDate, pTime] = rawStartString.includes('T') ? rawStartString.split('T') : [meetingDate, startTime];
        const [pExpDate, pEndTime] = rawEndString.includes('T') ? rawEndString.split('T') : [expiryDate, endTime];

        let realTeamsData;
        try {
            realTeamsData = await createTeamsMeeting({
                organizerEmail,
                subject: customSubject,
                description,
                startDateTime: startDateObj.toISOString(),
                endDateTime: endDateObj.toISOString(),
                coOrganizers: coOrganizersList.map(c => c.email)
            });
        } catch (graphErr) {
            console.error('MS Graph API Error:', graphErr);
            return renderFormWithError('Failed to generate Microsoft Teams meeting via Graph API.');
        }

        await Meeting.create({
            subjectCode,
            subject: customSubject,
            description: description || null,
            meetingDate: pDate,
            expiryDate: pExpDate || pDate,
            startTime: pTime ? pTime.slice(0, 5) : '08:00',
            endTime: pEndTime ? pEndTime.slice(0, 5) : '17:00',
            joinUrl: realTeamsData.joinUrl,
            teamsMeetingId: realTeamsData.teamsMeetingId,
            creatorEmail: organizerEmail,
            coOrganizers: JSON.stringify(coOrganizersList)
        });

        return res.redirect(`/subjects/${subjectCode}`);
    } catch (err) {
        console.error('Create Action Error:', err);
        return renderFormWithError('Error recording meeting schedule.');
    }
});

// Action: Edit Meeting
router.post('/api/subjects/:subject_code/meetings/:id/edit', requireAuth, requireTeamsRole, async (req, res) => {
    const subjectCode = req.params.subject_code.toUpperCase();
    const meetingId = req.params.id;
    const isMoodle = req.session.isMoodle || false;

    const renderFormWithError = async (errorMsg) => {
        const subjectObj = await Subject.findByPk(subjectCode);
        let moodleLecturers = (req.session.moodleLecturers && req.session.moodleLecturers[subjectCode]) || [];
        if (isMoodle && fetchMoodleLecturers && moodleLecturers.length === 0) {
            try { moodleLecturers = await fetchMoodleLecturers(req.session, subjectCode); } catch (e) {}
        }
        return res.status(400).render('meeting-form', {
            user: req.session.user,
            isAdmin: req.session.user.role === 'admin',
            isMoodle,
            isEditMode: true,
            subject: subjectObj ? subjectObj.get({ plain: true }) : { code: subjectCode },
            moodleLecturers,
            meeting: { ...req.body, id: meetingId },
            errorMessage: errorMsg
        });
    };

    try {
        const meeting = await Meeting.findByPk(meetingId);
        if (!meeting) return renderFormWithError('Meeting not found.');

        const currentUserEmail = req.session.user.email.toLowerCase();
        const isCreator = meeting.creatorEmail && meeting.creatorEmail.toLowerCase() === currentUserEmail;
        const isAdmin = req.session.user.role === 'admin';

        if (!isCreator && !isAdmin) {
            return renderFormWithError('Forbidden: Only the meeting creator or an administrator can edit this meeting.');
        }

        const { subject: customSubject, description, startDateTime, endDateTime, meetingDate, expiryDate, startTime, endTime } = req.body;
        const coOrganizersList = extractCoOrganizers(req.body);

        const fixedEndDateStr = (req.session.moodleSettings && req.session.moodleSettings[subjectCode]?.fixed_end_date) 
            || req.session.fixedEndDate 
            || expiryDate 
            || (endDateTime ? endDateTime.split('T')[0] : null);

        let startDateObj;
        let endDateObj;

        if (isMoodle && fixedEndDateStr) {
            const rawStart = startDateTime || (meetingDate && startTime ? `${meetingDate}T${startTime}` : `${meeting.meetingDate}T${meeting.startTime}`);
            startDateObj = new Date(rawStart);
            endDateObj = new Date(`${fixedEndDateStr}T23:59:59.000Z`);

            if (isNaN(endDateObj.getTime())) {
                return renderFormWithError(`Invalid fixed end date format (${fixedEndDateStr}).`);
            }
        } else {
            const rawStart = startDateTime || (meetingDate && startTime ? `${meetingDate}T${startTime}` : null);
            const rawEnd = endDateTime || (expiryDate && endTime ? `${expiryDate}T${endTime}` : rawStart);

            if (!rawStart) {
                return renderFormWithError('Start date and time are required.');
            }

            startDateObj = new Date(rawStart);
            endDateObj = new Date(rawEnd || rawStart);

            if (isNaN(startDateObj.getTime()) || isNaN(endDateObj.getTime())) {
                return renderFormWithError('Invalid date format received.');
            }
        }

        const rawStartString = startDateTime || `${meetingDate}T${startTime}`;
        const rawEndString = endDateTime || `${expiryDate}T${endTime}`;
        const [parsedMeetingDate, parsedStartTime] = rawStartString.includes('T') ? rawStartString.split('T') : [meetingDate, startTime];
        const [parsedExpiryDate, parsedEndTime] = rawEndString.includes('T') ? rawEndString.split('T') : [expiryDate, endTime];

        if (meeting.teamsMeetingId) {
            try {
                await updateTeamsMeeting({
                    organizerEmail: meeting.creatorEmail,
                    teamsMeetingId: meeting.teamsMeetingId,
                    subject: customSubject,
                    description,
                    startDateTime: startDateObj.toISOString(),
                    endDateTime: endDateObj.toISOString(),
                    coOrganizers: coOrganizersList.map(c => c.email)
                });
            } catch (graphErr) {
                console.error('MS Graph API Update Error:', graphErr.message);
            }
        }

        await meeting.update({
            subject: customSubject,
            description: description || null,
            meetingDate: parsedMeetingDate,
            expiryDate: parsedExpiryDate,
            startTime: parsedStartTime ? parsedStartTime.slice(0, 5) : '08:00',
            endTime: parsedEndTime ? parsedEndTime.slice(0, 5) : '17:00',
            coOrganizers: JSON.stringify(coOrganizersList)
        });

        return res.redirect(`/subjects/${subjectCode}`);
    } catch (err) {
        console.error('Update Action Error:', err);
        return renderFormWithError('Error updating meeting details.');
    }
});

// Action: Delete Meeting
router.delete('/api/meetings/:id', requireAuth, requireTeamsRole, async (req, res) => {
    try {
        const meeting = await Meeting.findByPk(req.params.id);
        if (!meeting) {
            return res.status(404).json({ success: false, error: 'Meeting not found.' });
        }

        const currentUserEmail = req.session.user.email.toLowerCase();
        const isCreator = meeting.creatorEmail && meeting.creatorEmail.toLowerCase() === currentUserEmail;
        const isAdmin = req.session.user.role === 'admin';

        if (!isCreator && !isAdmin) {
            return res.status(403).json({ 
                success: false, 
                error: 'Forbidden: You do not have authorization to delete this meeting.' 
            });
        }

        if (meeting.teamsMeetingId) {
            try {
                await deleteTeamsMeeting({
                    organizerEmail: meeting.creatorEmail,
                    teamsMeetingId: meeting.teamsMeetingId
                });
            } catch (graphErr) {
                console.error('MS Graph API Delete Error:', graphErr.message);
            }
        }

        await meeting.destroy();
        return res.json({ success: true, message: 'Meeting successfully deleted.' });
    } catch (err) {
        console.error('Delete Action Error:', err);
        return res.status(500).json({ success: false, error: 'Server error processing deletion.' });
    }
});

module.exports = router;