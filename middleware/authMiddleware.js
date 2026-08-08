const { ALLOWED_MANAGEMENT_ROLES } = require('../config/constants');

function requireAuth(req, res, next) {
    if (!req.session.user) {
        return res.redirect('/login');
    }
    next();
}

const requireTeamsRole = (req, res, next) => {
    if (!req.session || !req.session.user) {
        if (req.xhr || req.headers.accept?.includes('json')) {
            return res.status(401).json({ success: false, error: 'Authentication required.' });
        }
        return res.redirect('/login');
    }

    const isMoodle = !!req.session.isMoodle;
    const userRole = req.session.user.role;

    if (isMoodle) {
        // Strict Moodle rule: Only 'msteam_account' role can schedule
        if (userRole !== 'msteam_account') {
            if (req.xhr || req.headers.accept?.includes('json')) {
                return res.status(403).json({ 
                    success: false, 
                    error: 'Forbidden: Only MS Teams account holders can schedule meetings in Moodle.' 
                });
            }
            return res.status(403).send('Forbidden: Only MS Teams account holders are allowed to schedule meetings from Moodle.');
        }
    } else {
        // Standard application rule
        if (!ALLOWED_MANAGEMENT_ROLES.includes(userRole)) {
            if (req.xhr || req.headers.accept?.includes('json')) {
                return res.status(403).json({ success: false, error: 'Access denied: Scheduling permissions required.' });
            }
            return res.status(403).send('Forbidden: You do not have permission to schedule meetings.');
        }
    }

    next();
};
/**
 * Middleware to restrict route access strictly to Admin users.
 */
const requireAdmin = (req, res, next) => {
    // Ensure user is authenticated first
    if (!req.session || !req.session.user) {
        if (req.xhr || req.headers.accept?.includes('json')) {
            return res.status(401).json({ success: false, error: 'Authentication required.' });
        }
        return res.redirect('/login');
    }

    // Check if user has admin role
    if (req.session.user.role !== 'admin') {
        if (req.xhr || req.headers.accept?.includes('json')) {
            return res.status(403).json({ success: false, error: 'Access denied: Admin permissions required.' });
        }
        return res.status(403).send('Forbidden: Access is restricted to Administrators only.');
    }

    next();
};

module.exports = {
    // ... existing exports ...
    requireAdmin
};
module.exports = {
    requireAuth,
    requireTeamsRole,
    requireAdmin
};