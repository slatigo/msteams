const express = require('express');
const router = express.Router();
const bcrypt = require('bcryptjs');
const { User } = require('../models');

// Middleware: Require Admin Privileges
function ensureAdmin(req, res, next) {
    if (req.session && req.session.user && req.session.user.role === 'admin') {
        return next();
    }
    return res.status(403).send('Access Denied: Administrator access required.');
}

// GET /users - Render User Management Page
router.get('/', ensureAdmin, async (req, res, next) => {
    try {
        const users = await User.findAll({
            attributes: ['id', 'name', 'email', 'role', 'authSource', 'isActive', 'createdAt'],
            order: [['createdAt', 'DESC']],
            raw: true
        });

        res.render('users', {
            user: req.session.user,
            usersList: users,
            isMoodle: req.session.isMoodle || false
        });
    } catch (err) {
        console.error('User Fetch Error:', err);
        next(err); // Passed to central Express error handler
    }
});

// POST /api/users - Create Direct User
router.post('/api/users', ensureAdmin, async (req, res) => {
    try {
        const { name, email, password, role } = req.body;

        if (!name || !email || !password || !role) {
            return res.status(400).json({ success: false, error: 'All fields are required.' });
        }

        const normalizedEmail = email.toLowerCase().trim();

        const existingUser = await User.findOne({ where: { email: normalizedEmail } });
        if (existingUser) {
            return res.status(400).json({ success: false, error: 'User with this email already exists.' });
        }

        const hashedPassword = await bcrypt.hash(password, 10);

        const newUser = await User.create({
            name: name.trim(),
            email: normalizedEmail,
            password: hashedPassword,
            role,
            authSource: 'local',
            isActive: true
        });

        return res.json({ success: true, user: { id: newUser.id, name: newUser.name, email: newUser.email } });
    } catch (err) {
        console.error('Create User Error:', err);
        return res.status(500).json({ success: false, error: 'Failed to create user.' });
    }
});

// PUT /api/users/:id - Update User / Password
router.put('/api/users/:id', ensureAdmin, async (req, res) => {
    try {
        const { name, role, isActive, password } = req.body;
        const targetUser = await User.findByPk(req.params.id);

        if (!targetUser) {
            return res.status(404).json({ success: false, error: 'User not found.' });
        }

        if (name) targetUser.name = name.trim();
        if (role) targetUser.role = role;
        if (typeof isActive === 'boolean') targetUser.isActive = isActive;

        if (password && password.trim() !== '') {
            targetUser.password = await bcrypt.hash(password, 10);
        }

        await targetUser.save();
        return res.json({ success: true });
    } catch (err) {
        console.error('Update User Error:', err);
        return res.status(500).json({ success: false, error: 'Failed to update user.' });
    }
});

// DELETE /api/users/:id - Delete User
router.delete('/api/users/:id', ensureAdmin, async (req, res) => {
    try {
        const targetUser = await User.findByPk(req.params.id);
        if (!targetUser) {
            return res.status(404).json({ success: false, error: 'User not found.' });
        }

        // Prevent self-deletion
        if (targetUser.id === req.session.user.id) {
            return res.status(400).json({ success: false, error: 'You cannot delete your own account.' });
        }

        await targetUser.destroy();
        return res.json({ success: true });
    } catch (err) {
        console.error('Delete User Error:', err);
        return res.status(500).json({ success: false, error: 'Failed to delete user.' });
    }
});

module.exports = router;