let app;
try {
  app = require('../src/server');
} catch (err) {
  console.error('Fatal initialization error:', err);
  app = (req, res) => {
    res.status(500).json({
      success: false,
      error: 'Backend Initialization Failure',
      message: err && err.message,
      stack: err && err.stack
    });
  };
}

module.exports = app;


