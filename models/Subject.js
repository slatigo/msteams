const { DataTypes } = require('sequelize');

module.exports = (sequelize) => {
    return sequelize.define('Subject', {
        code: { 
            type: DataTypes.STRING, 
            primaryKey: true, 
            allowNull: false 
        },
        name: { 
            type: DataTypes.STRING, 
            allowNull: false 
        },
        type: { 
            type: DataTypes.ENUM('class', 'general'), 
            defaultValue: 'class' 
        },
        description: { 
            type: DataTypes.TEXT, 
            allowNull: true 
        }
    });
};