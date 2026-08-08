const express = require('express');
const router = express.Router();
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');

const { User, Subject } = require('../models');

router.get('/auth/moodle-sso', async (req, res) => {
    const { token } = req.query;

    if (!token) {
        return res.status(400).send('Authentication token missing from Moodle request.');
    }

    try {
        const sharedSecret = process.env.MOODLE_SHARED_SECRET;
        const decoded = jwt.verify(token, sharedSecret, { algorithms: ['HS256'] });

        const { user, course, lecturers, settings } = decoded;

        if (!user?.email || !course?.code) {
            return res.status(400).send('Malformed payload: User email or Course code missing.');
        }

        const courseCode = course.code.trim().toUpperCase();

        req.session.isMoodle = true;
        req.session.user = {
            email: user.email.toLowerCase(),
            name: user.fullname || user.name || 'Moodle User',
            role: user.role || 'teacher'
        };

        // Store lecturers directly as an array for the active session
        req.session.moodleLecturers = Array.isArray(lecturers) ? lecturers : [];
        req.session.fixedEndDate = settings?.fixed_end_date || null;

        await Subject.findOrCreate({
            where: { code: courseCode },
            defaults: {
                code: courseCode,
                name: course.name || courseCode,
                type: 'class',
                description: 'Imported from Moodle Course'
            }
        });

        req.session.save((err) => {
            if (err) {
                console.error('Session save error during Moodle SSO:', err);
                return res.status(500).send('Session initialization failure.');
            }
            return res.redirect(`/subjects/${courseCode}`);
        });

    } catch (err) {
        console.error('Moodle SSO Verification Error:', err.message);
        return res.status(401).send('Invalid or expired Moodle session token.');
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
            req.session.isMoodle = false;
            req.session.fixedEndDate = null; // Clear Moodle settings for direct logins
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