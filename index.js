require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const { Client } = require("@googlemaps/google-maps-services-js");
const readline = require('readline');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

const mapsClient = new Client({});

// OAuth2 configuration
const oauth2Client = new google.auth.OAuth2(
  process.env.GOOGLE_CLIENT_ID,
  process.env.GOOGLE_CLIENT_SECRET,
  process.env.GOOGLE_REDIRECT_URI
);

// Add these functions near the top of the file
function saveTokens(tokens) {
  try {
    fs.writeFileSync(path.join(__dirname, 'tokens.json'), JSON.stringify(tokens));
  } catch (error) {
    console.error('Error saving tokens:', error);
  }
}

function loadTokens() {
  try {
    const tokens = fs.readFileSync(path.join(__dirname, 'tokens.json'));
    return JSON.parse(tokens);
  } catch (error) {
    return null;
  }
}

// Add this function after the loadTokens function
async function verifyTokens(tokens) {
  try {
    oauth2Client.setCredentials(tokens);
    const calendar = google.calendar({ version: 'v3', auth: oauth2Client });
    
    // Try to make a minimal API call to verify tokens
    await calendar.calendarList.list({
      maxResults: 1,
    });
    
    return true;
  } catch (error) {
    console.log('Stored tokens are invalid or expired, reauthorizing...');
    return false;
  }
}

// Modify the callback route to save tokens
app.get('/oauth2callback', async (req, res) => {
  const { code } = req.query;
  try {
    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);
    saveTokens(tokens); // Save tokens after getting them
    res.send('Authorization successful! You can close this window and return to the terminal.');
    
    await processEvents();
  } catch (error) {
    res.send('Error during authorization: ' + error.message);
  }
});

// Modify the server start to check for existing tokens
app.listen(port, async () => {
  console.log(`Server is running on http://localhost:${port}`);
  
  // Check for existing tokens and verify them
  const tokens = loadTokens();
  if (tokens && await verifyTokens(tokens)) {
    console.log('Found existing valid authorization, using saved tokens');
    oauth2Client.setCredentials(tokens);
    processEvents();
  } else {
    // Start the OAuth flow if tokens are missing or invalid
    const authUrl = getAuthUrl();
    console.log('Please visit this URL to authorize the application:', authUrl);
  }
});

// Generate authentication URL
function getAuthUrl() {
  const scopes = ['https://www.googleapis.com/auth/calendar.readonly'];
  return oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: scopes,
  });
}

// Modify promptDateRange to handle timezone-aware dates
async function promptDateRange() {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  const question = (query) => new Promise((resolve) => rl.question(query, resolve));

  try {
    console.log('Please enter dates in YYYY-MM-DD format');
    const startDate = await question('Start date: ');
    const endDate = await question('End date: ');
    const maxDistance = await question('Maximum one-way distance to include (in miles): ');
    rl.close();
    
    // Create dates at start of day and end of day in local timezone
    const start = new Date(startDate);
    start.setHours(0, 0, 0, 0);
    
    const end = new Date(endDate);
    end.setHours(23, 59, 59, 999);
    
    const maxMiles = parseFloat(maxDistance);
    
    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      throw new Error('Invalid date format');
    }
    if (isNaN(maxMiles) || maxMiles <= 0) {
      throw new Error('Invalid distance. Please enter a positive number');
    }

    return {
      start: start.toISOString(),  // This will convert to UTC but preserve the correct local time
      end: end.toISOString(),
      maxMiles: maxMiles
    };
  } catch (error) {
    rl.close();
    throw new Error(error.message);
  }
}

// Add this helper function to extract company domain
function getCompanyFromEmail(email) {
  try {
    const domain = email.split('@')[1];
    return domain.split('.')[0]; // This is very basic, could be enhanced
  } catch {
    return null;
  }
}

// Modify getEventsWithMileage to handle pagination
async function getEventsWithMileage(auth, timeMin, timeMax, maxMiles) {
  const calendar = google.calendar({ version: 'v3', auth });
  const homeAddress = process.env.HOME_ADDRESS;
  const eventsWithMileage = [];
  let totalEvents = 0;
  
  try {
    let pageToken;
    do {
      // Add delay between page requests to avoid rate limiting
      if (pageToken) {
        await new Promise(resolve => setTimeout(resolve, 100)); // 100ms delay
      }

      const response = await calendar.events.list({
        calendarId: 'primary',
        timeMin: timeMin,
        timeMax: timeMax,
        singleEvents: true,
        orderBy: 'startTime',
        pageToken: pageToken,
        maxResults: 100  // Process in smaller chunks
      });

      const events = response.data.items;
      totalEvents += events.length;
      console.log(`Processing events... (${totalEvents} so far)`);

      for (const event of events) {
        // Process attendees regardless of location
        const attendees = event.attendees?.map(attendee => {
          const name = attendee.displayName ? {
            fullName: attendee.displayName,
            firstName: attendee.displayName.split(' ')[0],
            lastName: attendee.displayName.split(' ').slice(1).join(' ')
          } : null;
          
          return {
            ...name,
            email: attendee.email,
            company: getCompanyFromEmail(attendee.email)
          };
        }) || [];

        if (event.location) {
          try {
            const distance = await calculateDistance(homeAddress, event.location);
            
            // Process attendees
            const attendees = event.attendees?.map(attendee => {
              const name = attendee.displayName ? {
                fullName: attendee.displayName,
                // Basic name splitting - could be enhanced
                firstName: attendee.displayName.split(' ')[0],
                lastName: attendee.displayName.split(' ').slice(1).join(' ')
              } : null;
              
              return {
                ...name,
                email: attendee.email,
                company: getCompanyFromEmail(attendee.email)
              };
            }) || [];

            if (distance) {
              const oneWayMiles = parseFloat(distance.oneWayMiles.replace(' mi', ''));
              const includeInTotal = oneWayMiles <= maxMiles;
              
              eventsWithMileage.push({
                summary: event.summary,
                location: event.location,
                start: event.start.dateTime || event.start.date,
                distance: distance,
                includeInTotal: includeInTotal,
                attendees: attendees,
                organizer: {
                  email: event.organizer?.email,
                  company: getCompanyFromEmail(event.organizer?.email)
                }
              });
            } else {
              eventsWithMileage.push({
                summary: event.summary,
                location: event.location,
                start: event.start.dateTime || event.start.date,
                distance: null,
                includeInTotal: false,
                error: "Could not calculate distance",
                attendees: attendees,
                organizer: {
                  email: event.organizer?.email,
                  company: getCompanyFromEmail(event.organizer?.email)
                }
              });
            }
          } catch (error) {
            console.error(`Error processing "${event.summary}": ${error.message}`);
          }
        } else {
          console.log(`Skipping event (no location): "${event.summary}" on ${event.start.dateTime || event.start.date}`);
          eventsWithMileage.push({
            summary: event.summary,
            location: null,
            start: event.start.dateTime || event.start.date,
            distance: null,
            includeInTotal: false,
            error: "No location provided - likely virtual or internal meeting",
            attendees: attendees,
            organizer: {
              email: event.organizer?.email,
              company: getCompanyFromEmail(event.organizer?.email)
            }
          });
        }
      }

      pageToken = response.data.nextPageToken;
    } while (pageToken);

    console.log(`Processed ${totalEvents} total events`);
    return eventsWithMileage;

  } catch (error) {
    if (error.code === 429) {  // HTTP 429 is "Too Many Requests"
      console.error('Hit Google Calendar API rate limit. Please try again later.');
    }
    console.error('Error fetching events:', error);
    throw error;
  }
}

// Calculate distance between two addresses using Google Maps API
async function calculateDistance(origin, destination) {
  try {
    const response = await mapsClient.distancematrix({
      params: {
        origins: [origin],
        destinations: [destination],
        key: process.env.GOOGLE_MAPS_API_KEY,
      },
    });

    if (response.data.rows[0].elements[0].status === 'OK') {
      // Double the distance for round trip
      const oneWayMiles = response.data.rows[0].elements[0].distance.text;
      const roundTripMiles = parseFloat(oneWayMiles.replace(' mi', '')) * 2;
      return {
        miles: `${roundTripMiles.toFixed(1)} mi`,
        duration: response.data.rows[0].elements[0].duration.text,
        oneWayMiles: oneWayMiles
      };
    }
    return null;
  } catch (error) {
    console.error('Error calculating distance:', error);
    throw error;
  }
}

// Add this helper function to format date for filename
function formatDateForFilename(isoDate) {
  return isoDate.split('T')[0]; // Convert ISO date to YYYY-MM-DD
}

// Modify formatDateTime to respect local timezone
function formatDateTime(isoString) {
  const date = new Date(isoString);
  // Format in local timezone
  return date.toLocaleString('en-US', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
    timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone  // Use system timezone
  }).replace(',', '');
}

// Add this helper function to sanitize strings for CSV
function sanitizeForCSV(str) {
  if (!str) return '';
  // Remove emojis and other special characters
  return str
    .replace(/[\u{1F600}-\u{1F6FF}\u{2600}-\u{26FF}]/gu, '') // Remove emojis
    .replace(/[\u0000-\u001F\u007F-\u009F]/g, '')            // Remove control characters
    .replace(/,/g, ';')                                       // Replace commas with semicolons
    .trim();                                                  // Remove extra whitespace
}

// Modify the generateMileageCSV function
function generateMileageCSV(events, summary, filename) {
  const driveEvents = events.filter(e => e.distance && e.distance.miles);
  
  // Create summary section
  const summaryLines = [
    'MILEAGE SUMMARY',
    `Total Included Mileage: ${parseFloat(summary.totalMiles)} miles`,
    `Excluded Mileage: ${parseFloat(summary.excludedMiles)} miles`,
    `Maximum One-Way Distance: ${summary.maxOneWayMiles} miles`,
    `Total Drive Events: ${driveEvents.length}`,
    '',  // blank line
    'DETAILED MILEAGE LIST',
    '',  // blank line before headers
    'Time,Meeting Title,Location,Round Trip Miles,One-Way Miles,Included In Total'
  ];
  
  // Create event rows
  const rows = driveEvents.map(event => {
    const roundTripMiles = parseFloat(event.distance.miles.replace(' mi', ''));
    const oneWayMiles = parseFloat(event.distance.oneWayMiles.replace(' mi', ''));
    
    return [
      formatDateTime(event.start),
      `"${sanitizeForCSV(event.summary)}"`,
      `"${sanitizeForCSV(event.location)}"`,
      roundTripMiles.toFixed(1),
      oneWayMiles.toFixed(1),
      event.includeInTotal ? 'Yes' : 'No'
    ].join(',');
  });
  
  // Combine everything and write to file
  fs.writeFileSync(filename, [...summaryLines, ...rows].join('\n'));
}

// Modify the generateMeetingsCSV function
function generateMeetingsCSV(events, filename) {
  // Find max number of attendees for header row
  const maxAttendees = events.reduce((max, event) => 
    Math.max(max, (event.attendees || []).length), 0);
  
  // Create header row
  const headers = ['Time', 'Meeting Title', 'Location'];
  for (let i = 1; i <= maxAttendees; i++) {
    headers.push(`Attendee ${i}`);
  }
  
  // Create summary section
  const summaryLines = [
    'MEETING SUMMARY',
    `Total Meetings: ${events.length}`,
    '',  // blank line
    'DETAILED MEETING LIST',
    '',  // blank line before headers
    headers.join(',')
  ];
  
  // Create event rows
  const rows = events.map(event => {
    const baseData = [
      formatDateTime(event.start),
      `"${sanitizeForCSV(event.summary)}"`,
      `"${sanitizeForCSV(event.location || 'Virtual Meeting')}"`,
    ];
    
    // Add attendees, padding with empty strings if needed
    const attendeeData = (event.attendees || [])
      .map(a => `"${sanitizeForCSV(a.fullName || a.email)}"`)
      .concat(Array(maxAttendees).fill(''))
      .slice(0, maxAttendees);
    
    return [...baseData, ...attendeeData].join(',');
  });
  
  // Combine everything and write to file
  fs.writeFileSync(filename, [...summaryLines, ...rows].join('\n'));
}

// Add this helper function to ensure output directory exists
function ensureOutputDirectory() {
  const outputDir = path.join(__dirname, 'output');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
  }
  return outputDir;
}

// Modify the processEvents function
async function processEvents() {
  try {
    const outputDir = ensureOutputDirectory();
    const dateRange = await promptDateRange();
    const events = await getEventsWithMileage(
      oauth2Client, 
      dateRange.start, 
      dateRange.end, 
      dateRange.maxMiles
    );
    
    let totalMiles = 0;
    let excludedMiles = 0;
    let errorCount = 0;
    
    // Calculate totals (keeping the debug logs for now)
    events.forEach(event => {
      if (event.distance && event.distance.miles) {
        const miles = parseFloat(event.distance.miles.replace(' mi', ''));
        if (!isNaN(miles)) {
          if (event.includeInTotal) {
            totalMiles += miles;
          } else {
            excludedMiles += miles;
          }
        }
      } else if (event.error) {
        errorCount++;
      }
    });

    // Create output object
    const output = {
      summary: {
        totalMiles: totalMiles.toFixed(1),
        excludedMiles: excludedMiles.toFixed(1),
        maxOneWayMiles: dateRange.maxMiles,
        errorCount,
        dateRange: {
          start: dateRange.start,
          end: dateRange.end
        }
      },
      events
    };

    // Generate base filename without extension
    const baseFilename = `driving-deduction-${formatDateForFilename(dateRange.start)}-${formatDateForFilename(dateRange.end)}-${dateRange.maxMiles}mi`;
    
    // Write all files to output directory
    fs.writeFileSync(path.join(outputDir, `${baseFilename}.json`), JSON.stringify(output, null, 2));
    generateMeetingsCSV(events, path.join(outputDir, `${baseFilename}-meetings.csv`));
    generateMileageCSV(events, output.summary, path.join(outputDir, `${baseFilename}-mileage.csv`));
    
    // Update console output to show relative paths
    console.log('\nSummary:');
    console.log(`Total included mileage: ${totalMiles.toFixed(1)} miles`);
    console.log(`Excluded mileage (over ${dateRange.maxMiles} miles one-way): ${excludedMiles.toFixed(1)} miles`);
    if (errorCount > 0) {
      console.log(`${errorCount} event(s) had errors calculating distance`);
    }
    console.log(`\nFiles written to output directory:`);
    console.log(`- ${baseFilename}.json`);
    console.log(`- ${baseFilename}-meetings.csv`);
    console.log(`- ${baseFilename}-mileage.csv`);
    
    process.exit(0);
  } catch (error) {
    console.error('Error:', error.message);
    process.exit(1);
  }
} 