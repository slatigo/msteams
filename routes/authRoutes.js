const express = require('express');
const router = express.Router();
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');

const { User, Subject } = require('../models');

// MOODLE SSO HANDLER
router.get('/auth/moodle-sso', async (req, res) => {
    const { token } = req.query;

    if (!token) {
        return res.status(400).send('Authentication token missing from Moodle request.');
    }

    try {
        const sharedSecret = process.env.MOODLE_SHARED_SECRET || 'mubs-teams-sso-shared-secret-key-2026';
        const decoded = jwt.verify(token, sharedSecret);

        const { user, course, lecturers } = decoded;

        // Inside router.get('/auth/moodle-sso', ...)
        req.session.isMoodle = true; // Mark session as embedded in Moodle
        req.session.user = {
            email: user.email,
            name: user.fullname,
            role: user.role
        };

        req.session.moodleLecturers = req.session.moodleLecturers || {};
        req.session.moodleLecturers[course.code] = lecturers;

        await Subject.findOrCreate({
            where: { code: course.code },
            defaults: {
                code: course.code,
                name: course.name,
                type: 'class',
                description: 'Imported from Moodle Course'
            }
        });

        req.session.save((err) => {
            if (err) console.error('Session save error during Moodle SSO:', err);
            res.redirect(`/subjects/${course.code}`);
        });

    } catch (err) {
        console.error('Moodle SSO Verification Error:', err.message);
        res.status(401).send('Invalid or expired Moodle session token.');
    }
});

// LOGIN GET
router.get('/login', (req, res) => {
    if (req.session.user) {
        return res.redirect('/subjects');
    }
    res.render('login');
});

// LOGIN POST
router.post('/login', async (req, res) => {
    const { email, password } = req.body;
    try {
        const user = await User.findOne({ where: { email } });
        if (user && await bcrypt.compare(password, user.password)) {
            req.session.user = { email: user.email, role: user.role };
            // Inside router.post('/login', ...)
            req.session.isMoodle = false;
            return req.session.save((err) => {
                if (err) console.error('Session save error:', err);
                res.redirect('/subjects');
            });
        }
        res.render('login', { error: 'Invalid credentials.' });
    } catch (err) {
        res.status(500).render('login', { error: 'Authentication service error' });
    }
});

// LOGOUT
router.get('/logout', (req, res) => {
    req.session.destroy(() => res.redirect('/login'));
});

module.exports = router;