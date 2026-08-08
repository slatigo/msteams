'use strict';

/** @type {import('sequelize-cli').Migration} */
module.exports = {
  async up(queryInterface, Sequelize) {
    await queryInterface.addColumn('Meetings', 'description', {
      type: Sequelize.TEXT,
      allowNull: true,
      after: 'subject' // Places column after 'subject' in MySQL/MariaDB
    });
  },

  async down(queryInterface, Sequelize) {
    await queryInterface.removeColumn('Meetings', 'description');
  }
};