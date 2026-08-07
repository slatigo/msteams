const { DataTypes } = require('sequelize');

module.exports = (sequelize) => {
    return sequelize.define('User', {
        email: { 
            type: DataTypes.STRING, 
            allowNull: false, 
            unique: true 
        },
        password: { 
            type: DataTypes.STRING, 
            allowNull: false 
        },
        role: { 
            type: DataTypes.STRING, 
            defaultValue: 'student' 
        }
    });
};