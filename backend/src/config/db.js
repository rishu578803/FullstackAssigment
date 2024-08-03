require('dotenv').config()
const mongoose = require('mongoose');




mongoose
  .connect('mongodb://127.0.0.1:27017')
  .then((d) => {
    console.log("connected");
  })
  .catch((error) => {
    console.log("error", error);
  });

module.exports = { mongoose };