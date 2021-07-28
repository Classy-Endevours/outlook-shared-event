var graph = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

module.exports = {
  getUserDetails: async function (accessToken) {
    const client = getAuthenticatedClient(accessToken);

    const user = await client
      .api("/me")
      .select("displayName,mail,mailboxSettings,userPrincipalName")
      .get();
    return user;
  },
  getCalendarView: async function (
    accessToken,
    start,
    end,
    timeZone = "Pacific Standard Time"
  ) {
    const client = getAuthenticatedClient(accessToken);
    try {
      const events = await client
        .api("/me/calendars")
        // Add Prefer header to get back times in user's timezone
        .header("Prefer", `outlook.timezone="${timeZone}"`)
        // Add the begin and end of the calendar window
        // .query({
        //   startDateTime: new Date(start).toISOString(),
        //   endDateTime: new Date(end).toISOString(),
        // })
        // Get just the properties used by the app
        // .select("subject,organizer,start,end")
        // Order by start time
        // .orderby("start/dateTime")
        // Get at most 50 results
        .top(50)
        .get();

      return events;
    } catch (error) {
      return [];
    }
  },
  getUsers: async function (accessToken){
    const client = getAuthenticatedClient(accessToken);
    try {
        // POST /me/events
        const users = await client.api(`/users`).get();
        return users;
      } catch (error) {
        console.log({ error });
        return [];
    }
  },
  createEvent: async function (
    accessToken,
    formData,
    calendar,
    timeZone = "Pacific Standard Time"
  ) {
    const client = getAuthenticatedClient(accessToken);

    // Build a Graph event
    const newEvent = {
      subject: formData.subject,
      start: {
        dateTime: formData.start,
        timeZone: timeZone,
      },
      end: {
        dateTime: formData.end,
        timeZone: timeZone,
      },
      body: {
        contentType: "text",
        content: formData.body,
      },
    };

    // Add attendees if present
    if (formData.attendees) {
      newEvent.attendees = [];
      formData.attendees.forEach((attendee) => {
        newEvent.attendees.push({
          type: "required",
          emailAddress: {
            address: attendee,
          },
        });
      });
    }
    try {
      // POST /me/events
      await client.api(`/me/calendars/${calendar}/events`).post(newEvent);
    } catch (error) {
      console.log({ error });
    }
  },
};

function getAuthenticatedClient(accessToken) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done) => {
      done(null, accessToken);
    },
  });

  return client;
}
