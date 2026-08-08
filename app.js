require('dotenv').config();

const express = require('express');
const session = require('express-session');
const SequelizeStore = require('connect-session-sequelize')(session.Store); // 1. Import store
const path = require('path');

const { sequelize } = require('./models');
const { allowMoodleIframe } = require('./middleware/securityMiddleware');

const authRoutes = require('./routes/authRoutes');
const subjectRoutes = require('./routes/subjectRoutes');
const meetingRoutes = require('./routes/meetingRoutes');
const userRoutes = require('./routes/userRoutes');
const app = express();

// View Engine & Static Setup
app.set('view engine', 'pug');
app.set('views', path.join(__dirname, 'views'));

// Core Middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// 2. Configure Persistent Session Store
const sessionStore = new SequelizeStore({
    db: sequelize,
    tableName: 'Sessions' // Creates/uses 'Sessions' table in database
});

app.use(session({
    secret: process.env.SESSION_SECRET || 'mubs-teams-secret-key',
    store: sessionStore, // Attach database store
    resave: false,
    saveUninitialized: false,
    cookie: {
        maxAge: 24 * 60 * 60 * 1000 // 1 day (Adjust as needed)
    }
}));

// Moodle Iframe Embedding Guard
app.use(allowMoodleIframe);

// Mounted Routes
app.use(authRoutes);
app.use(subjectRoutes);
app.use(meetingRoutes);
// Mount user management under /users
app.use('/users', userRoutes);
// Catch-all Handler
app.use((req, res) => {
    res.redirect(req.session.user ? '/subjects' : '/login');
});

// 3. Synchronize Database, Session Table, and Start Server
const PORT = process.env.PORT || 3000;
sequelize.sync().then(() => {
    sessionStore.sync(); // Creates the Sessions table automatically if it doesn't exist
    app.listen(PORT, () => console.log(`MUBS MS Teams App running on http://localhost:${PORT}`));
});