'use strict';

module.exports = {
  async up(queryInterface, Sequelize) {
    // 1. Add expiryDate column
    await queryInterface.addColumn('Meetings', 'expiryDate', {
      type: Sequelize.DATEONLY,
      allowNull: true // Temporarily allow null to backfill existing rows
    });

    // 2. Backfill expiryDate with existing meetingDate for existing records (MySQL syntax)
    await queryInterface.sequelize.query(
      'UPDATE `Meetings` SET `expiryDate` = `meetingDate` WHERE `expiryDate` IS NULL;'
    );

    // 3. Set expiryDate to NOT NULL
    await queryInterface.changeColumn('Meetings', 'expiryDate', {
      type: Sequelize.DATEONLY,
      allowNull: false
    });

    // 4. Remove obsolete recurrence columns
    await queryInterface.removeColumn('Meetings', 'recurrenceType');
    await queryInterface.removeColumn('Meetings', 'recurrenceEndDate');
  },

  async down(queryInterface, Sequelize) {
    await queryInterface.addColumn('Meetings', 'recurrenceType', {
      type: Sequelize.STRING,
      defaultValue: 'none'
    });

    await queryInterface.addColumn('Meetings', 'recurrenceEndDate', {
      type: Sequelize.DATEONLY,
      allowNull: true
    });

    await queryInterface.removeColumn('Meetings', 'expiryDate');
  }
};