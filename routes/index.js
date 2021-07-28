const express = require('express');
const router = express.Router();

/* GET home page. */
router.get('/', function(req, res, next) {
  let params = {
    active: { home: true }
  };
  res.render('index', params);
});
// b50c261f-6811-4670-a3b8-68cff7179738
// 24381955-1b80-4417-a47d-7c60a61811f8
//  IY05A0HoJfR-Dj6--EBP.-O1Hd2g5gQ392 -value
module.exports = router;
