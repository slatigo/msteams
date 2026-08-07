require('dotenv').config(); // Guarantee environment variables are loaded first

const { Sequelize } = require('sequelize');

const sequelize = new Sequelize(
    process.env.DB_NAME || 'ms_teams',
    process.env.DB_USER || 'usery',
    process.env.DB_PASS || 'feblite1423',
    {
        host: process.env.DB_HOST || '127.0.0.1',
        port: process.env.DB_PORT || 3306,
        dialect: 'mysql',
        logging: false
    }
);

// Import model definitions
const User = require('./User')(sequelize);
const Subject = require('./Subject')(sequelize);
const Meeting = require('./Meeting')(sequelize);

// Define associations
Subject.hasMany(Meeting, { foreignKey: 'subjectCode', onDelete: 'CASCADE' });
Meeting.belongsTo(Subject, { foreignKey: 'subjectCode' });

module.exports = {
    sequelize,
    User,
    Subject,
    Meeting
};