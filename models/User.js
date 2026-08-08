const { DataTypes } = require('sequelize');

module.exports = (sequelize) => {
    return sequelize.define('User', {
        name: {
            type: DataTypes.STRING,
            allowNull: false,
            defaultValue: 'User'
        },
        email: { 
            type: DataTypes.STRING, 
            allowNull: false, 
            unique: true,
            validate: { isEmail: true }
        },
        password: { 
            type: DataTypes.STRING, 
            allowNull: true // Set to true so Moodle SSO users without local passwords can be saved
        },
        role: { 
            type: DataTypes.STRING, 
            defaultValue: 'student' 
        },
        authSource: {
            type: DataTypes.STRING,
            defaultValue: 'local' // 'local' or 'moodle'
        },
        isActive: {
            type: DataTypes.BOOLEAN,
            defaultValue: true
        }
    });
};