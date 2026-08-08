const { DataTypes } = require('sequelize');

module.exports = (sequelize) => {
    return sequelize.define('Meeting', {
        subjectCode: { 
            type: DataTypes.STRING, 
            allowNull: false 
        },
        subject: { 
            type: DataTypes.STRING, 
            allowNull: false 
        },
        description: { 
            type: DataTypes.TEXT, 
            allowNull: true 
        },
        meetingDate: { 
            type: DataTypes.DATEONLY, 
            allowNull: false 
        },
        expiryDate: { 
            type: DataTypes.DATEONLY, 
            allowNull: false 
        },
        startTime: { 
            type: DataTypes.STRING, 
            allowNull: false 
        },
        endTime: { 
            type: DataTypes.STRING, 
            allowNull: false 
        },
        joinUrl: { 
            type: DataTypes.TEXT, 
            allowNull: false 
        },
        teamsMeetingId: { 
            type: DataTypes.STRING, 
            allowNull: true 
        },
        creatorEmail: { 
            type: DataTypes.STRING, 
            allowNull: false 
        },
        coOrganizers: { 
            type: DataTypes.TEXT, 
            allowNull: true,
            defaultValue: '[]'
        }
    });
};