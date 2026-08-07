require('dotenv').config();

module.exports = {
  "development": {
    "username": process.env.DB_USER || "usery",
    "password": process.env.DB_PASS || "feblite1423",
    "database": process.env.DB_NAME || "ms_teams",
    "host": process.env.DB_HOST || "127.0.0.1",
    "dialect": "mysql"
  },
  "test": {
    "username": process.env.DB_USER || "root",
    "password": process.env.DB_PASS || null,
    "database": process.env.DB_NAME || "ms_teams",
    "host": process.env.DB_HOST || "127.0.0.1",
    "dialect": "mysql"
  },
  "production": {
    "username": process.env.DB_USER,
    "password": process.env.DB_PASS,
    "database": process.env.DB_NAME,
    "host": process.env.DB_HOST,
    "dialect": "mysql"
  }
};