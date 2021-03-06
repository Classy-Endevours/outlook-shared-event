const router = require('express-promise-router')();
const graph = require('../graph.js');
const addDays = require('date-fns/addDays');
const formatISO = require('date-fns/formatISO');
const startOfWeek = require('date-fns/startOfWeek');
const zonedTimeToUtc = require('date-fns-tz/zonedTimeToUtc');
const iana = require('windows-iana');
const { body, validationResult } = require('express-validator');
const validator = require('validator');

/* GET /calendar */
/* GET /calendar */
router.get('/',
  async function(req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      const params = {
        active: { calendar: true }
      };

      // Get the user
      const user = req.app.locals.users[req.session.userId];
      // Convert user's Windows time zone ("Pacific Standard Time")
      // to IANA format ("America/Los_Angeles")
      const timeZoneId = iana.findIana(user.timeZone)[0];
      console.log(`Time zone: ${timeZoneId?.valueOf()}`);

      // Calculate the start and end of the current week
      // Get midnight on the start of the current week in the user's timezone,
      // but in UTC. For example, for Pacific Standard Time, the time value would be
      // 07:00:00Z
      var weekStart = zonedTimeToUtc(startOfWeek(new Date()), timeZoneId?.valueOf());
      var weekEnd = addDays(weekStart, 7);
      console.log(`Start: ${formatISO(weekStart)}`);

      // Get the access token
      var accessToken;
      try {
        accessToken = await getAccessToken(req.session.userId, req.app.locals.msalClient);
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not get access token. Try signing out and signing in again.',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
        return;
      }

      if (accessToken && accessToken.length > 0) {
        try {
          // Get the events
          const events = await graph.getCalendarView(
            accessToken,
            formatISO(weekStart),
            formatISO(weekEnd),
            user.timeZone);
          
          const users = await graph.getUsers(accessToken);

          params.users = users.value;
          params.events = events.value;
        } catch (err) {
          req.flash('error_msg', {
            message: 'Could not fetch events',
            debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
          });
        }
      }
      else {
        req.flash('error_msg', 'Could not get an access token');
      }

      res.render('calendar', params);
    }
  }
);

router.get('/raw',
  async function(req, res) {
    if (!req.session.userId) {
      res.status(401).send('Unauthorized!')
    } else {
      const params = {
        active: { calendar: true }
      };

      // Get the user
      const user = req.app.locals.users[req.session.userId];
      // Convert user's Windows time zone ("Pacific Standard Time")
      // to IANA format ("America/Los_Angeles")
      const timeZoneId = iana.findIana(user.timeZone)[0];
      console.log(`Time zone: ${timeZoneId?.valueOf()}`);

      // Calculate the start and end of the current week
      // Get midnight on the start of the current week in the user's timezone,
      // but in UTC. For example, for Pacific Standard Time, the time value would be
      // 07:00:00Z
      var weekStart = zonedTimeToUtc(startOfWeek(new Date()), timeZoneId?.valueOf());
      var weekEnd = addDays(weekStart, 7);
      console.log(`Start: ${formatISO(weekStart)}`);

      // Get the access token
      var accessToken;
      try {
        accessToken = await getAccessToken(req.session.userId, req.app.locals.msalClient);
      } catch (err) {
        
        res.status(500).send('Could not get access token. Try signing out and signing in again.!')

        return;
      }

      if (accessToken && accessToken.length > 0) {
        try {
          // Get the events
          const events = await graph.getCalendarView(
            accessToken,
            formatISO(weekStart),
            formatISO(weekEnd),
            user.timeZone);
          
          const users = await graph.getUsers(accessToken);

          params.users = users.value;
          params.events = events.value;
        } catch (err) {
          res.status(500).send('Could not get calendars!')
        }
      }
      else {
        res.status(401).send('Could not get access token. Try signing out and signing in again.!')
      }
      res.status(200).send({
        ...params
      })
    }
  }
);

async function getAccessToken(userId, msalClient) {
  // Look up the user's account in the cache
  try {
    const accounts = await msalClient
      .getTokenCache()
      .getAllAccounts();

    const userAccount = accounts.find(a => a.homeAccountId === userId);

    // Get the token silently
    const response = await msalClient.acquireTokenSilent({
      scopes: process.env.OAUTH_SCOPES.split(','),
      redirectUri: process.env.OAUTH_REDIRECT_URI,
      account: userAccount
    });

    return response.accessToken;
  } catch (err) {
    console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
  }
}
/* GET /calendar/new */
router.get('/new',
  function(req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      res.locals.newEvent = {};
      res.render('newevent');
    }
  }
);
/* POST /calendar/new */
router.post('/new', [
    body('ev-subject').escape(),
    // Custom sanitizer converts ;-delimited string
    // to an array of strings
    body('ev-attendees').customSanitizer(value => {
      return value.split(';');
    // Custom validator to make sure each
    // entry is an email address
    }).custom(value => {
      value.forEach(element => {
        if (!validator.isEmail(element)) {
          throw new Error('Invalid email address');
        }
      });

      return true;
    }),
    // Ensure start and end are ISO 8601 date-time values
    body('ev-start').isISO8601(),
    body('ev-end').isISO8601(),
    body('ev-body').escape()
  ], async function(req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      // Build an object from the form values
      const formData = {
        subject: req.body['ev-subject'],
        attendees: req.body['ev-attendees'],
        start: req.body['ev-start'],
        end: req.body['ev-end'],
        body: req.body['ev-body']
      };

      // Check if there are any errors with the form values
      const formErrors = validationResult(req);
      if (!formErrors.isEmpty()) {

        let invalidFields = '';
        formErrors.errors.forEach(error => {
          invalidFields += `${error.param.slice(3, error.param.length)},`
        });

        // Preserve the user's input when re-rendering the form
        // Convert the attendees array back to a string
        formData.attendees = formData.attendees.join(';');
        return res.render('newevent', {
          newEvent: formData,
          error: [{ message: `Invalid input in the following fields: ${invalidFields}` }]
        });
      }

      // Get the access token
      var accessToken;
      try {
        accessToken = await getAccessToken(req.session.userId, req.app.locals.msalClient);
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not get access token. Try signing out and signing in again.',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
        return;
      }

      // Get the user
      const user = req.app.locals.users[req.session.userId];

      // Create the event
      try {
        await graph.createEvent(accessToken, formData, req.query.calendar ,user.timeZone);
      } catch (error) {
        req.flash('error_msg', {
          message: 'Could not create event',
          debug: JSON.stringify(error, Object.getOwnPropertyNames(error))
        });
      }

      // Redirect back to the calendar view
      return res.redirect('/calendar');
    }
  }
);

module.exports = router;