const path = require('path');
module.exports = {
  plugins: {
    'posthtml-include': {
      // Når en include starter med "/", vil vi at den tolkes relativt til prosjektroten
      root: __dirname
    }
  }
};
