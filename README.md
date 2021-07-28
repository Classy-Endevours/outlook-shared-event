# outlook-shared-event
A sample project which will create an shared repo for outlook

## Inspiration
This project is heavily inspired from [Node integration of outlook](https://docs.microsoft.com/en-us/graph/tutorials/node?context=outlook%2Fcontext&tutorial-step=1)
But it has working on shared outlook calender and adding event on that

## Run this project
`npm i` or `yarn` to install the dependencies
and then do `npm i -g nodemon`

### Environment Variable
- OAUTH_APP_ID = **your app id**
- OAUTH_APP_SECRET = **your app secret**
- OAUTH_REDIRECT_URI=http://localhost:3000/auth/callback
- OAUTH_SCOPES='user.read,calendars.readwrite,mailboxsettings.read'
- OAUTH_AUTHORITY=https://login.microsoftonline.com/common/

## Run the APP
To run the app you can start server by `npm start`

### Setup authentication
If you want to use API from the Postman or internal then you need to setup authentication for that.
**Steps**
- Start the server
- Login into the system by clicking into the login button
- Go to the calender page
- Take the Cookie - connect.sid from the browser
- Set the postman header with the Cookie to the value
- You are ready to go

Note - You need to keep this process on repeat in case your server restart

### Endpoints
- GET - /calendar/ - To fetch the html for listing
- GET - /calendar/raw - To fetch raw JSON for calendar
- POST - /calendar/new?calendar=your-shared-calendar-id - To create any calender with write access

```
{
    "ev-subject" : "Event subject",
    "ev-attendees" : "test@test.com;Test@test.com",
    "ev-start" : "2021-07-28T00:43",
    "ev-end" : "2021-08-07T20:49",
    "ev-body" : "This is a body"
}
```

