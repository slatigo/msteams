function allowMoodleIframe(req, res, next) {
    res.setHeader(
        "Content-Security-Policy", 
        "frame-ancestors 'self' http://localhost:* http://127.0.0.1:* http://localhost:8000 http://localhost:3000 https://*.mubs.ac.ug http://*.mubs.ac.ug"
    );
    res.removeHeader("X-Frame-Options");
    next();
}

module.exports = {
    allowMoodleIframe
};