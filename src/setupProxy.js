module.exports = function(app) {
    app.use((req, res, next) => {
      res.setHeader('X-Frame-Options', 'ALLOWALL');
      res.setHeader('Content-Security-Policy', 'frame-ancestors *');
      next();
    });
  };