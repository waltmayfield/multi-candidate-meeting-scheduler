// ==UserScript==
// @name         Phonetool Calendar
// @namespace    https://amazon.com/
// @version      1.17
// @updateURL    https://axzile.corp.amazon.com/-/carthamus/download_script/phonetool-calendar.user.js
// @downloadURL  https://axzile.corp.amazon.com/-/carthamus/download_script/phonetool-calendar.user.js
// @description  Display availability information, send OOTO invites, create meetings - all on the phonetool page
// @author       @novakin, @ionutfa
// @match        https://phonetool.amazon.com/users/*
// @grant        GM.xmlHttpRequest
// @grant        unsafeWindow
// ==/UserScript==

/* globals $, moment */

(function() {
    'use strict';

    const ExchangeDataProvider = {
        name: 'exchange',
        onajaxerror (reject) {
            return _ => {
                console.error(_);
                log('error', this.name + ': ' + _.response);
                reject(_);
            };
        },

        fetchEmail (username, isRetry) {
            return new Promise((resolve, reject) => {
                GM.xmlHttpRequest({
                    method: 'POST',
                    url: 'https://ballard.amazon.com/owa/service.svc',
                    headers: {
                        action: 'FindPeople',
                        'content-type': 'application/json; charset=UTF-8',
                        'x-owa-actionname': 'OwaOptionPage',
                        'x-owa-canary': this.getToken()
                    },
                    data: JSON.stringify({
                        Header: {
                            RequestServerVersion: 'Exchange2013'
                        },
                        Body: {
                            IndexedPageItemView: {
                                __type: 'IndexedPageView:#Exchange',
                                BasePoint: 'Beginning'
                            },
                            QueryString: username + '@'
                        }
                    }),
                    onerror: this.onajaxerror(reject),
                    onload: _ => {
                        this.updateToken(_);
                        if (_.status == 200) {
                            const responseBody = JSON.parse(_.response).Body;
                            resolve(responseBody.ResultSet
                                .filter(_ => Object.keys(_).includes('EmailAddress'))
                                .map(_ => _.EmailAddress.EmailAddress)
                                .find(_ => _.startsWith(username + '@'))
                            );
                        } else {
                            this.onajaxerror(reject)(_);
                        }
                    }
                });
            }).catch(e => {
                if (isRetry) {
                    throw e;
                }
                return this.fetchEmail(username, true);
            });
        },

        getAvailability (mailboxes, start, end, currentUserEmail) {
            return new Promise((resolve, reject) => {
                GM.xmlHttpRequest({
                    method: 'POST',
                    url: 'https://ballard.amazon.com/owa/service.svc',
                    headers: {
                        action: 'GetUserAvailabilityInternal',
                        'content-type': 'application/json; charset=UTF-8',
                        'x-owa-actionname': 'GetUserAvailabilityInternal_FetchWorkingHours',
                        'x-owa-canary': this.getToken()
                    },
                    data: JSON.stringify({
                        request: {
                            Header: {
                                RequestServerVersion: 'Exchange2013',
                                TimeZoneContext: {
                                    TimeZoneDefinition: { Id: 'UTC' }
                                }
                            },
                            Body: {
                                MailboxDataArray: mailboxes.map(_ => ({ Email: { Address: _ }})),
                                FreeBusyViewOptions: {
                                    RequestedView: 'Detailed',
                                    TimeWindow: { StartTime: start.toISOString(), EndTime: end.toISOString() }
                                }
                            }
                        }
                    }),
                    onerror: this.onajaxerror(reject),
                    onload: _ => {
                        this.updateToken(_);
                        if (_.status == 200) {
                            const response = JSON.parse(_.response);
                            const events = [];
                            response.Body.Responses.map(_ => _.CalendarView.Items.filter(_ => _.FreeBusyType != 'Free')).forEach(group => {
                                for (let i = 0; i < group.length; i++) {
                                    // Merge together events of the same type
                                    // TODO: merge together events of all types
                                    if (group[i + 1] && group[i].End >= group[i + 1].Start) {
                                        if (group[i].FreeBusyType == group[i + 1].FreeBusyType) {
                                            group[i + 1].Start = group[i].Start;
                                            if (group[i].End > group[i + 1].End) {
                                                group[i + 1].End = group[i].End;
                                            }
                                            continue;
                                        }
                                    }
                                    events.push(group[i]);
                                }
                            });
                            resolve(events.map(_ => {
                                let status = _.FreeBusyType.toLowerCase();
                                if (status == 'oof') {
                                    status = 'out-of-office';
                                }
                                return {
                                    className: 'calendar-' + (_.ParentFolderId.Id == currentUserEmail ? 'my-' : '') + status,
                                    title: status,
                                    start: moment.utc(_.Start),
                                    end: moment.utc(_.End)
                                };
                            }));
                        } else {
                            this.onajaxerror(reject)(_);
                        }
                    }
                });
            });
        },

        createMeeting ({ subject, organizer, requiredAttendees, start, end }) {
            return new Promise((resolve, reject) => {
                GM.xmlHttpRequest({
                    method: 'POST',
                    url: 'https://ballard.amazon.com/owa/service.svc',
                    headers: {
                        action: 'CreateItem',
                        'content-type': 'application/json; charset=UTF-8',
                        'x-owa-actionname': 'CreateCalendarItemAction',
                        'x-owa-canary': this.getToken()
                    },
                    data: JSON.stringify({
                        Header: {
                            RequestServerVersion: 'Exchange2013',
                            TimeZoneContext: {
                                TimeZoneDefinition: { Id: 'UTC' }
                            }
                        },
                        Body: {
                            Items: [{
                                __type: 'CalendarItem:#Exchange',
                                Subject: subject,
                                Body: {
                                    BodyType: 'HTML',
                                    Value: getBody()
                                },
                                Sensitivity: 'Normal',
                                IsResponseRequested: true,
                                Start: start.toISOString(),
                                End: end.toISOString(),
                                FreeBusyType: 'Busy',
                                RequiredAttendees: requiredAttendees.map(_ => ({
                                    // TODO: figure out which parameters are required
                                    __type: 'AttendeeType:#Exchange',
                                    Mailbox: {
                                        EmailAddress: _.email,
                                        RoutingType: 'SMTP',
                                        MailboxType: 'Mailbox',
                                        OriginalDisplayName: _.email
                                    }
                                }))
                            }],
                            SendMeetingInvitations: 'SendToAllAndSaveCopy'
                        }
                    }),
                    onerror: _ => this.onajaxerror(reject),
                    onload: _ => {
                        if (_.status == 200) {
                            resolve();
                        } else {
                            this.onajaxerror(reject)(_);
                        }
                    }
                });
            });
        },

        getToken () {
            return localStorage.phonetoolCalendarOutlookToken;
        },

        updateToken (response) {
            const match = response.responseHeaders.match(/x-owa-canary=(.*?);/i);
            if (match) {
                localStorage.phonetoolCalendarOutlookToken = match[1];
            }
        }
    }

    function getBody () {
        return `
            ==============Conference Bridge Information==============
            <br>
            You have been invited to an online meeting, powered by Amazon Chime.
            <br>
            Join via Chime clients (auto-call): This is an auto-call meeting, Chime will call you when the meeting starts, select 'Answer'
            <br>
            This meeting has been created by using the <b><u>1-click meeting</u></b> feature directly from the Phonetool page by using the <a href="https://w.amazon.com/bin/view/Scrat/Tools/PhonetoolCalendar/">Phonetool Calendar extension</a>.
            <br>
            =====================================================
            `;
    }

    let currentUser;
    function getCurrentUser () {
        if (!currentUser) {
            currentUser = JSON.parse(document.querySelector('[data-react-class=NavBar]').dataset.reactProps).currentUser;
        }
        return currentUser;
    }

    let targetUser;
    function getTargetUser () {
        if (!targetUser) {
            const userDetails = document.querySelector('[data-react-class="UserDetails"]');
            targetUser = userDetails && JSON.parse(userDetails.dataset.reactProps).targetUser;
        }
        return targetUser;
    }

    function isSundayToThursdayWeek() {
      return getTargetUser().targetUserBuilding.match(/^(TLV|HFA|AMM|CAI|JED|RUH|DMM|AHB|ELQ|HOF|ALY|LXR|KWI)/);
    }

    const Days = {
        Monday: 1,
        Tuesday: 2,
        Wednesday: 3,
        Thursday: 4,
        Friday: 5,
        Saturday: 6,
        Sunday: 0
    };

    function getWeekend() {
        return isSundayToThursdayWeek() ? [Days.Friday, Days.Saturday] : [Days.Saturday, Days.Sunday];
    }

    function getBusinessHours () {
        const commonWeekdays = [Days.Monday, Days.Tuesday, Days.Wednesday, Days.Thursday];
        const daysOfWeek = isSundayToThursdayWeek() ? [Days.Sunday, ...commonWeekdays] : [...commonWeekdays, Days.Friday]; // Days of week

        const UTCOffset = getTargetUser().targetUserUTCOffset;
        const offsetInMinutes = moment().utcOffset() - moment().utcOffset(UTCOffset).utcOffset();

        const startTime = moment().hours(9).minutes(offsetInMinutes).format('HH:mm');
        const endTime = moment().hours(18).minutes(offsetInMinutes).format('HH:mm');

        // This can be simplified when https://github.com/fullcalendar/fullcalendar/issues/4440 is resolved
        return startTime < endTime
            ? { daysOfWeek, startTime, endTime }
            : [{ daysOfWeek, startTime: '00:00', endTime }, { daysOfWeek, startTime, endTime: '24:00' }];
    }

    async function initCredentials(dataProvider = ExchangeDataProvider) {
        const user = getTargetUser().targetUserLogin;
        const currentUser = getCurrentUser();
        const email = dataProvider.fetchEmail(user);
        const currentEmail = user == currentUser ? email : dataProvider.fetchEmail(currentUser);
        return Promise.all([dataProvider, user, currentUser, email, currentEmail]);
    }

    const sessionId = new Date % 1e9 + Math.random().toPrecision(3);
    function log (event, message) {
        new Image().src = 'http://ptcalendar.corp.amazon.com/ptcalendar.cyder?' + [
            GM_info.script.name.replace(/\s+/g, '_'),
            GM_info.script.version,
            getCurrentUser(),
            event,
            message,
            sessionId,
            Math.random().toPrecision(3)
        ].join('&');
    }

    function shouldShowMyEvents() {
        const val = localStorage.phonetoolCalendarShowMyEvents;
        return val && JSON.parse(val);
    }

    function convertEvents(events) {
        for(let i in events) {
            let event = events[i];
            event.end = event.end._d;
            event.start = event.start._d;
            if (event.title === 'tentative') {
                event.textColor = '#555';
            }
        }
        return events;
    }

    /**
     * Allows dragging the widget to the widgets area and back. The position is saved in the localStorage
     *
     * This is copied and pasted from `enable_widget_movements` method with few overrides
     */
    function enableContainerDragging (container) {
        const $slots = $('#widgets1, #widgets2, #calendar-container');
        if (!$slots.sortable) {
            // For unknown reason, some users have an issue when $(...).sortable doesn't work which breaks the tool
            // This workaround disables container dragging, however, the main functionality remains unimpared
            console.error('$.fn.sortable method is not available');
            log('error', '$.fn.sortable method is not available');
            return;
        }
        $slots.sortable({
            handle: '.widget-move-handle',
            connectWith: '#widgets1, #widgets2, #calendar-container',
            start: function (e, ui) {
                ui.placeholder.height(ui.helper.outerHeight());
                $('#calendar-container').css('min-height', '100px');
            },
            placeholder: 'widget-placeholder',
            forceHelperSize: true,
            forcePlaceHolderSize: true,
            tolerance: 'pointer', // Overriden
            over: unsafeWindow.keep_equal_heights,
            update: function () {
                unsafeWindow.update_positions();
                // Overriden behavior
                // Every time after resorting, we calculate and save container position
                const parentElement = container.parentElement;
                const parentSelector = '#' + parentElement.id;
                const itemIndex = [...parentElement.children].findIndex(_ => _ == container);
                localStorage.phonetoolCalendarContainer = [parentSelector, itemIndex].join('%');

                if (parentElement.id == 'calendar-container') {
                    container.classList.remove('well');
                    document.querySelector('.SecondaryDetails').classList.add('calendar-atf-container');
                } else {
                    container.classList.add('well');
                    document.querySelector('.SecondaryDetails').classList.remove('calendar-atf-container');
                }
            },
            sort: function (e, ui) {
                // Overriden behavior
                // Prevent dropping the default widgets to above the fold area
                if (ui.item[0] != container && ui.placeholder.parent()[0].id == 'calendar-container') {
                    return false;
                }
            },
            stop: function () {
                // Overriden behavior
                $('#calendar-container').css('min-height', '0');
            }
        });
    }

    function initCss() {
        const css = document.createElement('link');
        css.href = 'https://cdn.jsdelivr.net/npm/fullcalendar@5.9.0/main.min.css';
        css.rel = 'stylesheet';
        document.head.appendChild(css);

        document.head.appendChild(document.createElement('style')).textContent = `
            .SecondaryDetails.calendar-atf-container {
                align-items: stretch;
            }
            .calendar-atf-container .UserLinks, .calendar-atf-container .ResolverRow { white-space: nowrap }

            .calendar-atf-container .SharePassion {
                flex-grow: 1;
                max-width: 1000px;
            }

            .calendar-widget .widget-move-handle {
                position: relative;
                height: 17px;
                padding-left: 20px;
                background: url('data:image/svg+xml;utf8,<svg viewBox="0 0 25 30" xmlns="http://www.w3.org/2000/svg" fill="grey"><circle cx="5" cy="5" r="3"/><circle cx="18" cy="5" r="3"/><circle cx="5" cy="15" r="3"/><circle cx="18" cy="15" r="3"/><circle cx="5" cy="25" r="3"/><circle cx="18" cy="25" r="3"/></svg>') no-repeat;
            }

            .script-warning { display: none; }
            .error-unauthorized .script-warning { display: block; font-weight: bold; }
            .error-unauthorized .calendar-my-events-label { display: none; }

            .calendar-tentative, .calendar-tentative:hover {
              background: repeating-linear-gradient(-48deg, white, white 1px, #99c8e9 3px, #99c8e9 3px, white 6px);
              color: #555;
            }
            .calendar-out-of-office { border-color: #800080; background: #800080; }
            .calendar-no-data { border-color: #888; background: #888; }

            .calendar-my-busy { border-color: #86ac39; background: #86ac39; }
            .calendar-my-tentative, .calendar-my-tentative:hover {
              border-color: #86ac39;
              background: repeating-linear-gradient(-48deg, white, white 1px, #cfea9a 3px, #cfea9a 3px, white 6px);
              color: #555;
            }
            .calendar-my-out-of-office { border-color: #af6aaf; background: #af6aaf; }

            .fc-event.fc-v-event { cursor: default; }

            .fc-nonbusiness { background: #bbb; } /* Darker background for non-business hours */

            .calendar-container { margin-top: 7px; display: inline-flex; width: 100%; min-width: 540px; height: 400px; }
            .calendar-my-events-label { float: right; }
            .calendar-my-events-checkbox { transform: translateX(-5px); vertical-align: baseline; }

            .calendar-container + hr { display: none; }
            .SharePassion .calendar-container + hr { display: block; }
        `;
    }

    async function phonetoolScript () {
        if (!getTargetUser() || !getTargetUser().targetUserActive) {
            return; // The user is inactive or doesn't exist
        }

        log('init');

        const credentials = initCredentials();

        initCss();

        await Promise.all([
            'https://cdn.jsdelivr.net/npm/fullcalendar@5.9.0/main.min.js',
            'https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.22.2/moment.min.js'
        ].map(src => new Promise(resolve => {
            const script = document.createElement('script');
            script.src = src;
            script.onload = resolve;
            document.body.appendChild(script);
        })));

        const container = document.createElement('div');
        container.className = 'calendar-widget';
        const droppableContainer = document.createElement('div');
        droppableContainer.id = 'calendar-container';
        droppableContainer.style = 'min-width: 50px';
        document.querySelector('.SharePassion').prepend(droppableContainer);
        const [parentSelector, containerIndex] = (localStorage.phonetoolCalendarContainer || '#calendar-container%0').split('%');
        const parentContainer = document.querySelector(parentSelector);
        parentContainer.insertBefore(container, parentContainer.children[containerIndex]);
        let hasError = false;

        // Awaiting the credentials __before__ actual changes to html in order to fail gracefully
        let dataProvider, user, currentUser, email, currentEmail;
        try {
            [dataProvider, user, currentUser, email, currentEmail] = await credentials;
        } catch(e) {
            // high chances of missing Microsoft Exchange authentication token: show the error message
            container.classList.add('error-unauthorized');
            hasError = true;
        }

        const showMyEventsButton = user !== currentUser;
        const myEventsButton = showMyEventsButton ? `<label class="calendar-my-events-label"><input class="calendar-my-events-checkbox" type="checkbox">Display my events</label>` : '';

        container.innerHTML = `
          <div class="widget-move-handle">
            ${myEventsButton}
            <div style="position: absolute; left: 50%; transform: translateX(-50%);">
              <a href="https://w.amazon.com/bin/view/Scrat/Tools/PhonetoolCalendar/#HChangelog">
                Calendar Wiki
                ${localStorage.phonetoolCalendarVersion == GM_info.script.version ? '' : '<b>(see what is new in version ' + GM_info.script.version + ')</b>'}
              </a>
            </div>
          </div>
          <p class="script-warning">
            Unable to fetch calendar data, usually due to missing Exchange authentication token.<br>
            To solve this, try to login on <a href="https://ballard.amazon.com/owa/">ballard.amazon.com</a> and then refresh this page.<br>
            For more information please check out <a href="https://w.amazon.com/bin/view/Scrat/Tools/PhonetoolCalendar#HTroubleshooting2FFAQ">the Calendar's wiki page</a> that is being kept updated.<br>
          </p>
          <div id="cc-content" class="calendar-container"></div>
          <hr>
        `;

        if (hasError) {
            document.querySelector('#cc-content').remove();
        }

        if (parentContainer.classList.contains('widgets')) {
            container.classList.add('well');
        } else {
            document.querySelector('.SecondaryDetails').classList.add('calendar-atf-container');
        }

        if (showMyEventsButton) {
            container.querySelector('.calendar-my-events-checkbox').checked = shouldShowMyEvents();
        }
        localStorage.phonetoolCalendarVersion = GM_info.script.version;

        enableContainerDragging(container);

        // Stop rendering the calendar, if the user is missing (due to some load error)
        if (!user) {
            return;
        }

        const calendarEl = container.querySelector('.calendar-container');

        if (showMyEventsButton) {
            $(container.querySelector('.calendar-my-events-checkbox')).on('change', _ => {
                localStorage.phonetoolCalendarShowMyEvents = _.target.checked;
                calendar.destroy();
                calendar = new FullCalendar.Calendar(calendarEl, calendarParams);
                calendar.render();
            });
        }

        const subjectText = user === currentUser
            ? `This will create an Out of the Office calendar entry`
            : `This will create a quick meeting with @${user}`;

        const calendarParams = {
            allDaySlot: false,
            headerToolbar: {
                left: 'prev,next today',
                center: 'title',
                right: 'dayGridMonth,timeGridWeek,timeGridDay,listWeek'
            },
            businessHours: getBusinessHours(),
            nowIndicator: true,
            hiddenDays: getWeekend(),
            initialView: 'timeGridWeek',
            navLinks: true, // can click day/week names to navigate views
            // editable: true,
            scrollTime: '09:00:00',
            selectable: true,
            selectMirror: true,
            select: function(selectionInfo) {
                const [start, end] = [selectionInfo.start, selectionInfo.end];
                log('select');
                if (start < moment()._d) {
                    alert('Cannot create a meeting in the past');
                } else {
                    const subject = prompt(`${subjectText}. If agree, please enter the subject.`);
                    if (subject !== null) {
                        log('new_meeting');
                        const requiredAttendees = [{ email: 'meet@chime.aws' }];
                        if (email != currentEmail) {
                            requiredAttendees.push({ email });
                        }
                        dataProvider.createMeeting({ subject, organizer: currentEmail, requiredAttendees, start, end })
                            .then(_ => {
                                alert(`The meeting "${subject}" is successfully created!`);
                                log('create_meeting', dataProvider.name);
                                calendar.addEvent({ start, end });
                            })
                            .catch(_ => {
                                console.error(_);
                                alert('Error');
                            });
                    }
                }
                calendar.unselect();
            },
            dayMaxEventRows: true, // allow "more" link when too many events
            events: function (fetchInfo, successCallback, failureCallback) {
                const mailboxes = [email];
                if (email != currentEmail && shouldShowMyEvents()) {
                    mailboxes.push(currentEmail);
                }

                dataProvider.getAvailability(mailboxes, fetchInfo.start, fetchInfo.end, mailboxes[1]).then(events => {
                    log('fetch', JSON.stringify([dataProvider.name, mailboxes, fetchInfo.start, fetchInfo.end]));
                    const convertedEvents = convertEvents(events);
                    successCallback([].concat(...convertedEvents));
                }).catch(e => {
                    console.error(e);
                    log('error', dataProvider.name + ': ' + e.response);
                });
            }
        };

        var calendar = new FullCalendar.Calendar(calendarEl, calendarParams);

        calendar.render();

    }

    phonetoolScript();

})();
