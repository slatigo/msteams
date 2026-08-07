const express = require('express');
const router = express.Router();

const { Subject, Meeting } = require('../models');
const { requireAuth, requireTeamsRole } = require('../middleware/authMiddleware');
const { parseCoOrganizersField } = require('../helpers/meetingHelpers');
const { ALLOWED_MANAGEMENT_ROLES } = require('../config/constants');

// ROOT LANDING
router.get('/', (req, res) => {
    if (!req.session.user) {
        return res.redirect('/login');
    }
    res.redirect('/subjects');
});

// LIST ALL SUBJECTS
router.get('/subjects', requireAuth, async (req, res) => {
    try {
        const subjects = await Subject.findAll({ order: [['type', 'ASC'], ['code', 'ASC']] });
        res.render('subjects-list', {
            user: req.session.user,
            canManage: ALLOWED_MANAGEMENT_ROLES.includes(req.session.user.role),
            subjects: subjects.map(s => s.get({ plain: true }))
        });
    } catch (err) {
        console.error('Fetch Subjects Error:', err);
        res.status(500).send('Database Error loading subjects.');
    }
});

// NEW SUBJECT FORM
router.get('/subjects/new', requireAuth, requireTeamsRole, (req, res) => {
    res.render('subject-form', { user: req.session.user });
});

// CREATE SUBJECT ACTION
router.post('/api/subjects', requireAuth, requireTeamsRole, async (req, res) => {
    try {
        const { code, name, type, description } = req.body;
        await Subject.create({
            code: code.trim().toUpperCase(),
            name,
            type: type || 'class',
            description
        });
        res.redirect('/subjects');
    } catch (err) {
        console.error('Create Subject Error:', err);
        res.status(500).send('Failed to create subject. Ensure code is unique.');
    }
});

// VIEW SUBJECT MEETINGS (SCHEDULE)
router.get('/subjects/:subject_code', requireAuth, async (req, res) => {
    try {
        const subjectCode = req.params.subject_code.toUpperCase();
        const currentUserEmail = req.session.user.email.toLowerCase();
        
        let subject = await Subject.findByPk(subjectCode);
        if (!subject) {
            subject = await Subject.create({
                code: subjectCode,
                name: req.query.course_name || subjectCode,
                type: 'class',
                description: 'Auto-created from course integration'
            });
        }

        const rawMeetings = await Meeting.findAll({
            where: { subjectCode: subjectCode },
            order: [['meetingDate', 'ASC'], ['startTime', 'ASC']]
        });

        const meetings = rawMeetings.map(m => {
            const plain = m.get({ plain: true });
            const coOrganizers = parseCoOrganizersField(plain.coOrganizers);
            
            // --- STRICT CREATOR & ADMIN CHECK ---
            const isCreator = plain.creatorEmail && plain.creatorEmail.toLowerCase() === currentUserEmail;
            const isAdmin = req.session.user.role === 'admin';

            return {
                ...plain,
                coOrganizers,
                startDateTimeISO: `${plain.meetingDate}T${plain.startTime || '08:00'}:00`,
                endDateTimeISO: `${plain.expiryDate || plain.meetingDate}T${plain.endTime || '17:00'}:00`,
                // Can edit/delete only if creator or admin
                canEdit: isCreator || isAdmin
            };
        });

        res.render('schedule', {
            user: req.session.user,
            isMoodle: !!req.session.isMoodle, // Pass iframe state
            subject: subject.get({ plain: true }),
            canManage: ALLOWED_MANAGEMENT_ROLES.includes(req.session.user.role),
            sessions: meetings
        });
    } catch (err) {
        console.error('View Subject Error:', err);
        res.status(500).send('Database Error loading schedule.');
    }
});

module.exports = router;