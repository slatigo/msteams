const { ALLOWED_MANAGEMENT_ROLES } = require('../config/constants');

function requireAuth(req, res, next) {
    if (!req.session.user) {
        return res.redirect('/login');
    }
    next();
}

function requireTeamsRole(req, res, next) {
    if (!req.session.user) {
        return res.redirect('/login');
    }
    if (ALLOWED_MANAGEMENT_ROLES.includes(req.session.user.role)) {
        return next();
    }
    return res.status(403).send('Access Denied: Insufficient permissions.');
}

module.exports = {
    requireAuth,
    requireTeamsRole
};