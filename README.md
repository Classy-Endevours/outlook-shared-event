# outlook-shared-event
A sample project which will create an shared repo for outlook

## Inspiration
This project is heavily inspired from [Node integration of outlook](https://docs.microsoft.com/en-us/graph/tutorials/node?context=outlook%2Fcontext&tutorial-step=1)
But it has working on shared outlook calender and adding event on that

## Run this project
`npm i` or `yarn` to install the dependencies

### Environment Variable
- OAUTH_APP_ID = **your app id**
- OAUTH_APP_SECRET = **your app secret**
- OAUTH_REDIRECT_URI=http://localhost:3000/auth/callback
- OAUTH_SCOPES='user.read,calendars.readwrite,mailboxsettings.read'
- OAUTH_AUTHORITY=https://login.microsoftonline.com/common/
