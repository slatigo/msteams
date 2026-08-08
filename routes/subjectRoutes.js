const express = require('express');
const router = express.Router();

const { Subject, Meeting } = require('../models');
const { requireAuth, requireTeamsRole } = require('../middleware/authMiddleware');
const { parseCoOrganizersField } = require('../helpers/meetingHelpers');
const { ALLOWED_MANAGEMENT_ROLES } = require('../config/constants');

// ==========================================
// ROOT & SUBJECT LIST ROUTES
// ==========================================

// Landing Page Redirect
router.get('/', (req, res) => {
    if (!req.session.user) {
        return res.redirect('/login');
    }
    res.redirect('/subjects');
});

// List All Subjects
router.get('/subjects', requireAuth, async (req, res, next) => {
    try {
        const subjects = await Subject.findAll({ order: [['type', 'ASC'], ['code', 'ASC']] });
        res.render('subjects-list', {
            user: req.session.user,
            canManage: ALLOWED_MANAGEMENT_ROLES.includes(req.session.user.role),
            subjects: subjects.map(s => s.get({ plain: true }))
        });
    } catch (err) {
        console.error('Fetch Subjects Error:', err);
        next(err); // Handled by global Express error middleware
    }
});

// Render Form: Create New Subject
router.get('/subjects/new', requireAuth, requireTeamsRole, (req, res) => {
    res.render('subject-form', { 
        user: req.session.user,
        error: null,
        formData: {}
    });
});

// Action: Create Subject API
router.post('/api/subjects', requireAuth, requireTeamsRole, async (req, res, next) => {
    const { code, name, type, description } = req.body;
    try {
        if (!code || !name) {
            return res.status(400).render('subject-form', {
                user: req.session.user,
                error: 'Subject code and title are required.',
                formData: { code, name, type, description }
            });
        }

        await Subject.create({
            code: code.trim().toUpperCase(),
            name: name.trim(),
            type: type || 'class',
            description
        });

        res.redirect('/subjects');
    } catch (err) {
        console.error('Create Subject Error:', err);
        
        let userMessage = 'Failed to create subject. Please check your inputs.';
        if (err.name === 'SequelizeUniqueConstraintError') {
            userMessage = `Subject Code "${code.toUpperCase()}" already exists.`;
        }

        res.status(400).render('subject-form', {
            user: req.session.user,
            error: userMessage,
            formData: { code, name, type, description }
        });
    }
});

router.get('/subjects/:subject_code', requireAuth, async (req, res, next) => {
    try {
        const subjectCode = req.params.subject_code.toUpperCase();

        // Ensure subject exists in DB
        let subject = await Subject.findByPk(subjectCode);
        if (!subject) {
            await Subject.create({
                code: subjectCode,
                name: req.query.course_name || subjectCode,
                type: 'class',
                description: 'Auto-created from course integration'
            });
        }

        // Redirect to unified meetings endpoint with query param
        res.redirect(`/meetings?subjectCode=${subjectCode}`);
    } catch (err) {
        console.error('Subject auto-creation error:', err);
        next(err);
    }
});

module.exports = router;