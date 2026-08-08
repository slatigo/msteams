'use strict';

/** @type {import('sequelize-cli').Migration} */
module.exports = {
  async up(queryInterface, Sequelize) {
    // 1. Add 'name' column
    await queryInterface.addColumn('Users', 'name', {
      type: Sequelize.STRING,
      allowNull: false,
      defaultValue: 'User'
    });

    // 2. Modify 'password' column to allow NULL (for Moodle SSO users)
    await queryInterface.changeColumn('Users', 'password', {
      type: Sequelize.STRING,
      allowNull: true
    });

    // 3. Add 'authSource' column
    await queryInterface.addColumn('Users', 'authSource', {
      type: Sequelize.STRING,
      defaultValue: 'local'
    });

    // 4. Add 'isActive' column
    await queryInterface.addColumn('Users', 'isActive', {
      type: Sequelize.BOOLEAN,
      defaultValue: true
    });
  },

  async down(queryInterface, Sequelize) {
    // Revert changes in reverse order
    await queryInterface.removeColumn('Users', 'isActive');
    await queryInterface.removeColumn('Users', 'authSource');

    await queryInterface.changeColumn('Users', 'password', {
      type: Sequelize.STRING,
      allowNull: false
    });

    await queryInterface.removeColumn('Users', 'name');
  }
};