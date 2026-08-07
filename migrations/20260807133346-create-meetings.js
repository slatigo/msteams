'use strict';

module.exports = {
  async up(queryInterface, Sequelize) {
    await queryInterface.createTable('Meetings', {
      id: {
        allowNull: false,
        autoIncrement: true,
        primaryKey: true,
        type: Sequelize.INTEGER
      },
      subjectCode: {
        type: Sequelize.STRING,
        allowNull: false,
        references: {
          model: 'Subjects', // Matches exact table name for Ubuntu case-sensitivity
          key: 'code'
        },
        onUpdate: 'CASCADE',
        onDelete: 'CASCADE'
      },
      subject: {
        type: Sequelize.STRING,
        allowNull: false
      },
      meetingDate: {
        type: Sequelize.DATEONLY,
        allowNull: false
      },
      startTime: {
        type: Sequelize.STRING,
        allowNull: false
      },
      endTime: {
        type: Sequelize.STRING,
        allowNull: false
      },
      joinUrl: {
        type: Sequelize.TEXT,
        allowNull: false
      },
      teamsMeetingId: {
        type: Sequelize.STRING,
        allowNull: true
      },
      creatorEmail: {
        type: Sequelize.STRING,
        allowNull: false
      },
      recurrenceType: {
        type: Sequelize.STRING,
        defaultValue: 'none'
      },
      recurrenceEndDate: {
        type: Sequelize.DATEONLY,
        allowNull: true
      },
      createdAt: {
        allowNull: false,
        type: Sequelize.DATE
      },
      updatedAt: {
        allowNull: false,
        type: Sequelize.DATE
      }
    });
  },

  async down(queryInterface, Sequelize) {
    await queryInterface.dropTable('Meetings');
  }
};