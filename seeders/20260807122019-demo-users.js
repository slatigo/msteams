'use strict';
const bcrypt = require('bcryptjs');

module.exports = {
  async up(queryInterface, Sequelize) {
    const defaultPassword = await bcrypt.hash('admin', 10);
    const studentPassword = await bcrypt.hash('student', 10);

    await queryInterface.bulkInsert('Users', [
      {
        email: 'slatigo@mubs.ac.ug',
        password: defaultPassword,
        role: 'ms_teams_account',
        createdAt: new Date(),
        updatedAt: new Date()
      },
      {
        email: 'lecturer@mubs.ac.ug',
        password: defaultPassword,
        role: 'ms_teams_account',
        createdAt: new Date(),
        updatedAt: new Date()
      },
      {
        email: 'student@mubs.ac.ug',
        password: studentPassword,
        role: 'student',
        createdAt: new Date(),
        updatedAt: new Date()
      }
    ], {});
  },

  async down(queryInterface, Sequelize) {
    await queryInterface.bulkDelete('Users', {
      email: ['slatigo@mubs.ac.ug', 'lecturer@mubs.ac.ug', 'student@mubs.ac.ug']
    }, {});
  }
};