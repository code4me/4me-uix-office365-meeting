/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) 4me inc.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * https://github.com/code4me/4me-uix-office365-meeting/blob/master/LICENSE.txt
 * -------------------------------------------------------------------------------------------
 */

 var $ = ITRP.$;            // jQuery
 var $extension = $(this);  // The UI Extension container with custom HTML
 console.log( "MSAL TEST 3.0" );
 console.log(JSON.stringify(ITRP.context));
 console.log(JSON.stringify(ITRP.record));
 
 var msalInstance;
 var username;
 var authenticationID;
 var userEmail;
 var accessToken;
 
 // Config ------------------------------------------------------------------------
 var useRooms = false;
 var startDate = new Date();
 startDate.setDate(startDate.getDate() + 1); // Tomorrow
 var endDate = new Date();
 endDate.setDate(startDate.getDate() + 10); // One week timeframe
 var startDateISO = startDate.toISOString();
 var endDateISO = endDate.toISOString();
 // -------------------------------------------------------------------------------
 
 
 //const msal_src = "https://alcdn.msauth.net/browser/2.1.0/js/msal-browser.min.js";
 //const msal_sha = "sha384-EmYPwkfj+VVmL1brMS1h6jUztl4QMS8Qq8xlZNgIT/luzg7MAzDVrRa2JxbNmk/e";
 const msal_src = "https://alcdn.msauth.net/browser/2.1.0/js/msal-browser.js";
 const msal_sha = "sha384-M9bRB06LdiYadS+F9rPQnntFCYR3UJvtb2Vr4Tmhw9WBwWUfxH8VDRAFKNn3VTc/";
 const useSSO = true;
 
 const msalConfig = {
   auth: {
     clientId: "7ac3ce20-7a5c-400f-9618-6b0f25a433b8",
     authority: "https://login.microsoftonline.com/hdt-software.com",
     redirectUri: window.location.origin + "/404.html",
     navigateToLoginRequestUrl: false,
   },
   cache: {
     cacheLocation: "localStorage", // This configures where your cache will be stored - sessionStorage or localStorage
     storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
   }
 };
 
 const graphConfig = {
   meEndpoint: "https://graph.microsoft.com/v1.0/me",
   findRoomsEndpoint: "https://graph.microsoft.com/beta/me/findRooms",
   findMeetingTimesEndpoint: "https://graph.microsoft.com/beta/me/findMeetingTimes",
   eventEndpoint: "https://graph.microsoft.com/beta/me/events",
   onlineMeetings: "https://graph.microsoft.com/beta/me/onlineMeetings",
 };
 
 const loginRequest = {
   loginHint: null,
   scopes: ["User.Read"],
 };
 
 const tokenRequest = {
   account: null,
   scopes: ["User.Read"],
   forceRefresh: true, // Set this to "true" to skip a cached token and go to the server to get a new token
 };
 
 const meetingTimeSuggestionsRequest = {
   attendees: [],  
   locationConstraint: { 
     isRequired: "false",  
     suggestLocation: "false",  
     locations: []
   },  
   timeConstraint: {
     activityDomain:"work", 
     timeslots: [ 
       { 
         start: { 
           dateTime: startDateISO,
           timeZone: "UTC" 
         },  
         end: { 
           dateTime: endDateISO,
           timeZone: "UTC" 
         }
       } 
     ] 
   },  
   isOrganizerOptional: "false",
   meetingDuration: "PT1H",
   returnSuggestionReasons: "true",
   minimumAttendeePercentage: "100"
 };
 
 const onlineMeetingsRequest = {
   startDateTime:"",
   endDateTime:"",
   subject:"Meeting"
 };
 
 const createEventRequest = {
   subject: null,
   body: {
     contentType: "HTML",
     content: null
   },
   start: {
     dateTime: null,
     timeZone: null
   },
   end: {
     dateTime: null,
     timeZone: null
   },
   location: {
     displayName:null
   },
   attendees: [
     /*{
       emailAddress: {
           address:"samanthab@contoso.onmicrosoft.com",
           name: "Samantha Booth"
       },
       "type": "required"
     }*/
   ],
   "allowNewTimeProposals": true,
   //"transactionId":"7E163156-7762-4BEB-A1C6-729EA81755A7"
 };
 
 function addAttendee ( email ) {
   if (email != null) {
     var attendee = { 
       type: "required",  
       emailAddress: {
         address: email 
       }
     };
     meetingTimeSuggestionsRequest.attendees.push(attendee);
   }
 }
 
 function callMSGraph (graphEndpoint, accessToken, callback, failureCallback) {
   var headers = new Headers({
     'Content-Type': 'application/json'
   });
   var bearer = "Bearer " + accessToken;
   headers.append("Authorization", bearer);
   var options = {
     method: "GET",
     headers: headers
   };
   fetch(graphEndpoint, options).then(function (response) {
     if (response.ok) return response.json();
     failureCallback(response.json());
   }).then(function (json) {
     //do something with response
     callback(json, accessToken);
   }).catch(failureCallback);
 }
 
 function postMSGraph (graphEndpoint, accessToken, payload, callback, failureCallback) {
   var headers = new Headers({
     'Content-Type': 'application/json'
   });
   var bearer = "Bearer " + accessToken;
   headers.append("Authorization", bearer);
   var options = {
     method: "POST",
     headers: headers,
     body: JSON.stringify(payload)
   };
   fetch(graphEndpoint, options).then(function (response) {
     return response.text();
   }).then(function (text) {
     if (text.length) {
       var json = JSON.parse(text);
       if (json.error) {
         failureCallback(json.error);
       } else {
         callback(json, accessToken);
       }
     } else {
       callback(null, accessToken);
     }
   }).catch(failureCallback);
 }
 
 function ValidateCachedMSALToken() {
   var timestamp = Math.floor((new Date()).getTime() / 1000);
 
   for (var _i = 0, _a = Object.keys(localStorage); _i < _a.length; _i++) {
     var key = _a[_i];
     if (key.includes('accesstoken')) {
       var val = JSON.parse(localStorage.getItem(key));
       if (val.expiresOn) {
         console.log("Access Token Expiration Dates: " + new Date(val.expiresOn * 1000) + " - " + new Date(val.extendedExpiresOn * 1000));
 
         // We have a (possibly expired) token
         if (val.expiresOn < timestamp) {
           console.warn("Access Token has expired!");
         }
         return;
       }
     }
   }
   throw new Error('No valid token found');
 }
 
 function msalCallback(successCallback) {
   console.log( "msal-browser.js Loaded" );
 
   // Create the main msal instance
   // configuration parameters are located at msalConfig.js
   msalInstance = new msal.PublicClientApplication(msalConfig);
 
   /*
   // Handle the redirect flows
   msalInstance.handleRedirectPromise().then(function (tokenResponse) {
     // Handle redirect response
     console.log("handleRedirectPromise tokenResponse: " + JSON.stringify(tokenResponse));
   }).catch(function (error) {
     // Handle redirect error
     console.error(error);
     onError(error);
   });
   */
 
   const currentAccounts = msalInstance.getAllAccounts();
   if (currentAccounts === null) {
     console.warn("No Account registered yet");
   } else if (currentAccounts.length > 1) {
     // Add choose account code here
     console.warn("Multiple accounts detected.");
   } else if (currentAccounts.length === 1) {
     console.log("Using Account: " + currentAccounts[0].username);
     //console.log(JSON.stringify(currentAccounts[0]));
     username = currentAccounts[0].username;
     loginRequest.loginHint = currentAccounts[0].username;
     tokenRequest.account = currentAccounts[0];
 
     try {
       var token = ValidateCachedMSALToken();
     } catch (error) {
       console.warn("MSALToken: " + error);
     }
     acquireToken(successCallback);
     return;
   }
 
   doLogin(successCallback);
 }
 
 function doLogin(successCallback) {
   if (loginRequest.loginHint == null) {
     getCurrentUser(function() { doLoginPopup(successCallback); });
   } else {
     doLoginPopup(successCallback);
   }
 }
 
 function doLoginPopup(successCallback) {
   if (useSSO == false) {
     console.log("SSO is Off");
     console.log("loginHint: " + loginRequest.loginHint);
     msalInstance.loginPopup(loginRequest).then(function(loginResponse) { acquireTokenFromLoginResponse(loginResponse, successCallback); })
       .catch(function (error) {
       console.error("login error: " + error);
     });
   } else {
     console.log("ssoSilent: " + loginRequest.loginHint);
     msalInstance.ssoSilent(loginRequest).then(function(loginResponse) { acquireTokenFromLoginResponse(loginResponse, successCallback); })
       .catch(function (error) {
       console.warn("ssoSilent: " + error);
       if (error instanceof msal.InteractionRequiredAuthError) {
         msalInstance.loginPopup(loginRequest)
           .then(function(loginResponse) { acquireTokenFromLoginResponse(loginResponse, successCallback); })
           .catch(function (error) {
           console.warn("ssoSilent: " + error);
           onError(error);
         });
       } else {
         onError(error);
       }
     }).catch(function (error) {
       console.error("loginPopup: " + error);
       onError(error);
     });
   }
 }
 
 function acquireTokenFromLoginResponse( loginResponse, successCallback ) {
   console.log("Getting Account from Login Response");
   //console.log(JSON.stringify(loginResponse));
   tokenRequest.account = loginResponse.account;
   acquireToken(successCallback);
 }
 
 function acquireToken(successCallback) {
   console.log("Acquiring Token");
   msalInstance.acquireTokenSilent(tokenRequest).then(function(tokenResponse) { handleTokenResponse(tokenResponse, successCallback); })
     .catch(function (error) {
     console.warn("silent token acquisition fails. acquiring token using redirect");
     console.log("acquireTokenSilent: " + error);
     if (error instanceof msal.InteractionRequiredAuthError) {
       // fallback to interaction when silent call fails
       msalInstance.acquireTokenPopup(tokenRequest).then(function(tokenResponse) { handleTokenResponse(tokenResponse, successCallback); });
     } else {
       doLogin();
     }
   }).catch(function (error) {
     console.log("acquireToken: " + error);
     onError(error);
   });
 
 }
 
 
 function handleTokenResponse ( tokenResponse, successCallback ) {
 
   //console.log(tokenResponse.accessToken);
   accessToken = tokenResponse.accessToken;
 
   // Check if the tokenResponse is null
   // If the tokenResponse !== null, then you are coming back from a successful authentication redirect. 
   // If the tokenResponse === null, you are not coming back from an auth redirect.
   //}.catch(function (error) {
   // handle error, either in the library or coming back from the server
 
   //getUser(ITRP.record, function(ITRP.record.member.id);
 
   successCallback(accessToken);
 }
 
 
 function findMeetingTimes(rooms, accessToken) {
   if (rooms) {
     $.each(rooms.value, function( i, r ) {
       var loc = {
         "resolveAvailability": true,
         "displayName": r.name,
         "locationEmailAddress": r.address
       };
       meetingTimeSuggestionsRequest.locationConstraint.locations.push(loc);
     });  
   }
   postMSGraph(graphConfig.findMeetingTimesEndpoint, accessToken, meetingTimeSuggestionsRequest, graphAPICallback, onError);
 }
 
 
 function onDemandScript( url, nonce, callback, successCallback ) {
   callback = (typeof callback != 'undefined') ? callback : {};
   $.ajax({
     type: "GET",
     url: url,
     attrs: { nonce: nonce, crossorigin: 'anonymous' },	// Important: crossorigin needed for ssoSilent to work
     success: function() { callback(successCallback); },
     dataType: "script",
     cache: true
   });
 }
 
 function getCurrentUser( callback ) { 
   var data = {
     url: '/v1/me',
     type: 'GET',
     dataType: 'json',
     success: function(jsondata) {
       userEmail = jsondata.primary_email;
       authenticationID = jsondata.authenticationID;
       loginRequest.loginHint = userEmail || authenticationID;
       callback();
     },
     error: onError,
   };
   $.ajax(data);
 }
 
 function getUserReq(ids, callback) { 
   var filtered = ids.filter(function (el) {
     return el != null && el != '';
   });
   var data = {
     url: '/v1/people?ids=' + filtered.join(',') + '&fields=primary_email',
     type: 'GET',
     dataType: 'json',
     success : callback,
     error: onError
   };
   $.ajax(data);
 }
 
 function onError(context)
 {
   console.warn("onError: " + context);
   if (context.code == "ErrorItemNotFound") {
     cancelMeeting();
   } else if (context.message) {
     errorMsg.html(context.message);
   } else {
     errorMsg.html(context);
   }
   interactionOn(false);
 }
 
 const dateDisplayOptions = { weekday: 'long', year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' };
 
 function isSameDay(d1, d2) {
   return d1.getFullYear() === d2.getFullYear() &&
     d1.getMonth() === d2.getMonth() &&
     d1.getDate() === d2.getDate();
 }
 function isPreferredTimeframe(d, ptf) {
   return (ptf === 'morning' && d.getHours() <= 13) || (ptf === 'afternoon' && d.getHours() > 13);
 }
 
 function meetingSlotTemplate( i, slot ) {
   var room = slot.locations[0];
   var roomDisplayName = '';
   var slotDate = new Date(slot.meetingTimeSlot.start.dateTime + 'Z');
   var preferedDate = $('#preferred_date').dateEntry('getDate');
   var preferredTimeframe = $('#preferred_timeframe').val();
   if (room) { roomDisplayName = room.displayName; }
   var html = '<div class="list-group-item meetingTimeSuggestion row">';
   html = html + '<input type="radio" name="slot" value="' + i + '">';
   html = html + '<div class="slotDate">';
   html = html +	'<p class="list-group-item-text">' + slotDate.toLocaleString(undefined, dateDisplayOptions) + '</p>';
   if (isSameDay(slotDate, preferedDate)) {
     html = html + '<span class="preferedDate">Prefered Date</span>';
   }
   if (isPreferredTimeframe(slotDate, preferredTimeframe)) {
     html = html + '<span class="preferedTime">Prefered Time</span>';
   }
   html = html + '</div>';
   html = html + '<div class="room">' + roomDisplayName + '</div>';
   html = html + '</div>';
   return html;
 }
 
 var scheduleButton = $extension.find('div.schedule-appointment');
 var confirmAppointmentButton = $extension.find('div.confirm-appointment');
 var cancelAppointmentButton = $extension.find('#cancel_appointment');
 var loading_spinner = $extension.find('#loading_spinner');
 var meetingSlots = $extension.find('#meetingSlots');
 var errorMsg = $extension.find('#errorMsg');
 var EventHtml = $extension.find('#EventHtml');
 var meetingDurationHolder = $extension.find('#meetingDurationHolder');
 var meetingTimeSuggestions;
 var selectedSlot;
 
 function graphAPICallback(data) {
   //alert(JSON.stringify(data, null, 2));
   $('#json').append(JSON.stringify(data, null, 2));
 
   var $meetingSlots = $extension.find('#meetingSlots');
   if (data.meetingTimeSuggestions.length === 0) {
     if (data.emptySuggestionsReason == 'AttendeesUnavailable') {
       errorMsg.html("Some attendees are not available and we couldn’t find any suitable meeting suggestions.<br>We apology for the inconvenience and kindly ask you to proceed with scheduling the appointment manually.");
     } else if (data.emptySuggestionsReason == 'AttendeesUnavailableOrUnknown') {
       errorMsg.html("Some attendees have unknown availability. Attendee availability can become unknown if the attendee is outside of the organization, or if there is an error obtaining free/busy information.<br> We apology for the inconvenience and kindly ask you to proceed with scheduling the appointment manually.");
     } else if (data.emptySuggestionsReason == 'LocationsUnavailable') {
       errorMsg.html("There are no suitable locations available at the calculated time slots. You could try to select a diferent location or proceed with scheduling the appointment manually.");
     } else if (data.emptySuggestionsReason == 'OrganizerUnavailable') {
       errorMsg.html("The organizer is not available during the requested time window.");
     } else {
       errorMsg.html("An unknown error prevented us from retrieving any suitable suggestions.<br>We apology for the inconvenience and kindly ask you to proceed with scheduling the appointment manually.");
     }
     interactionOn(false);
   } else {
     meetingDurationHolder.hide();
     meetingTimeSuggestions = data.meetingTimeSuggestions;
     $.each(data.meetingTimeSuggestions, function( i, slot ) {
       $meetingSlots.append(meetingSlotTemplate(i, slot));
     });
     interactionOn(true);
   }
 }
 
 scheduleButton.on('click', function() {
   console.log('Calling meetingTimeSuggestions API');
   interactionOff();
 
   var detail_content = JSON.parse($('#edit_content > .detail_content').attr('data'));
   //console.log(detail_content.customLinkParameters['person.primary_email']);
 
   var ids = [$('#req_requested_by_id').val(), $('#req_requested_for_id').val(), $('#req_member_id').val()];
   meetingTimeSuggestionsRequest.attendees = [];
   meetingTimeSuggestionsRequest.meetingDuration = $("input[type='radio'][name='meetingDuration']:checked").val();
   getUserReq(ids, function (jsondata) {
     //console.log(JSON.stringify(jsondata));
     $.each(jsondata, function( i, people ) {
       addAttendee(people.primary_email);
     });
     addAttendee(detail_content.customLinkParameters['person.primary_email']);
     console.log(JSON.stringify(meetingTimeSuggestionsRequest.attendees));
     onDemandScript(msal_src, msal_sha, msalCallback, function(accessToken){
       if (meetingTimeSuggestionsRequest.attendees.length >= 2) {
         if (useRooms) {
           callMSGraph(graphConfig.findRoomsEndpoint, accessToken, findMeetingTimes, onError);
         } else {
           postMSGraph(graphConfig.findMeetingTimesEndpoint, accessToken, meetingTimeSuggestionsRequest, graphAPICallback, onError);
         }
       } else {
         onError("Some attendees information is not available.");
       }	
     });
   });  
 });
 
 
 meetingSlots.on('change', function() {
   selectedSlot = $('input[name=slot]:checked', '#meetingSlots').val();
   $('#json').append(JSON.stringify(selectedSlot, null, 2));
   if (selectedSlot != null) {
     confirmAppointmentButton.css("display", "inline-block");
   }
 });
 
 confirmAppointmentButton.on('click', function() {
   interactionOff();
 
   var slot = meetingTimeSuggestions[selectedSlot];
   //console.log(JSON.stringify(slot));
 
   createEventRequest.subject = 'Request #' + ITRP.record.initialValues.id + ' ' + $('#req_subject').val();
   createEventRequest.start = slot.meetingTimeSlot.start;
   createEventRequest.end = slot.meetingTimeSlot.end;
 
   $.each(slot.attendeeAvailability, function( i, attendeeAvailability ) {
     createEventRequest.attendees.push(
       {
         emailAddress: attendeeAvailability.attendee.emailAddress,
         type: "required"
       }
     );
   });
 
   if ($('#modality').val() == "ms_teams") {
     createEventRequest.isOnlineMeeting = true;
     createEventRequest.onlineMeetingProvider = "teamsForBusiness";
   }
 
   createEventRequest.body.content = $extension.find('#invite_message').val();
 
   postMSGraph(graphConfig.eventEndpoint, accessToken, createEventRequest, function (data)
               {
     console.log(JSON.stringify(data));
     $("#meetingId").val(data.id);
     
     //ITRP.field('note').val({ html: data.body.content });
 
 
     // Meeting Details
     var body = '<p><b>Start Time:</b>&nbsp;' + new Date(data.start.dateTime + 'Z').toLocaleString(undefined, dateDisplayOptions) + '<br>';
     body = body + '<b>End Time:</b>&nbsp;' + new Date(data.end.dateTime + 'Z').toLocaleString(undefined, dateDisplayOptions) + '<br>';
     body = body + '<b>Organizer:</b>&nbsp;' + data.organizer.emailAddress.name + '<br>';
     $.each(data.attendees, function( i, attendee ) {
       if (i == 1) {
         body = body + '<b>Attendees:</b>&nbsp;' + attendee.emailAddress.name + '<br>';
       } else {
         body = body + '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' + attendee.emailAddress.name + '<br>';
       }
     });
     if (data.onlineMeeting) {
       body = body + '<a href="' + data.onlineMeeting.joinUrl + '">Click here to join the meeting</a><br>';
     }
     body = body + '</p>';
       
     $("#meeting").val({ html: body });
     EventHtml.html(body + data.body.content);
     $("#meetingHtml").val(body + data.body.content);
 
     meetingSlots.hide();
     interactionOn(true);
     confirmAppointmentButton.hide();
     $("#meeting-scheduler-wrapper").css("display", "none");
   }, onError);
 
 
 });
 
 cancelAppointmentButton.on('click', function() {
   interactionOff();
 
   var mid = $("#meetingId").val();
   if (mid) {
     onDemandScript(msal_src, msal_sha, msalCallback, function(accessToken) {
       postMSGraph(graphConfig.eventEndpoint + '/' + mid + '/cancel', accessToken, { Comment: "Appointment canceled." }, function (data) {
         cancelMeeting();
       }, onError);
     });
   }
 
 });
 
 function cancelMeeting() {
   $("#meetingId").val("");
   $("#meeting").val("");
   $("#meetingHtml").val("");
   EventHtml.html(null);
 
   meetingSlots.hide();
   interactionOn(true);
   confirmAppointmentButton.hide();
   $("#meeting-scheduler-wrapper").css("display", "none");
   cancelAppointmentButton.css("display", "none");
 }
 
 function interactionOff()
 {
   errorMsg.hide();
   scheduleButton.disabled = true;
   confirmAppointmentButton.disabled = true;
   loading_spinner.show();
 }
 function interactionOn(success)
 {
   loading_spinner.hide();
   if (success) {
     scheduleButton.hide();
   } else {
     errorMsg.show();
   }
   scheduleButton.disabled = false;
   confirmAppointmentButton.disabled = false;
 }
 
 ITRP.hooks.register('after-prefill', function() {
   //console.log($("#req_member_id").val());
   //console.log($("#meetingId").val());
   EventHtml.html($("#meetingHtml").val());
   if ($("#req_member_id").val() && !($("#meetingId").val())) { 
     $("#meeting-scheduler-wrapper").css("display", "block");
     scheduleButton.css("display", "inline-block");
   }
   // if a meeting ID is set we display the cancel button
   if ($("#meetingId").val()) {
     cancelAppointmentButton.css("display", "inline-block");
   }
   // Always hide in edit mode
   errorMsg.hide();
   $("#scheduled-meeting-details").css("display", "none");
   
 });