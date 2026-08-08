'use strict';

/** @type {import('sequelize-cli').Migration} */
module.exports = {
  async up(queryInterface, Sequelize) {
    // Note: Use your actual database table name (usually plural 'Meetings')
    await queryInterface.changeColumn('Meetings', 'subjectCode', {
      type: Sequelize.STRING,
      allowNull: true
    });

    await queryInterface.changeColumn('Meetings', 'subject', {
      type: Sequelize.STRING,
      allowNull: true
    });
  },

  async down(queryInterface, Sequelize) {
    await queryInterface.changeColumn('Meetings', 'subjectCode', {
      type: Sequelize.STRING,
      allowNull: false
    });

    await queryInterface.changeColumn('Meetings', 'subject', {
      type: Sequelize.STRING,
      allowNull: false
    });
  }
};