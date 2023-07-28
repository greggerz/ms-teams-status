/* eslint-disable no-console */
const alarmName = 'checkTeamsAvailability';
let currentIdleState: chrome.idle.IdleState = 'active';

chrome.runtime.onInstalled.addListener(() => {
  chrome.alarms.create(alarmName, { periodInMinutes: 0.2 });
});

/**
 * Set the teams status based on current status and idle status
 *
 * @param {chrome.idle.IdleState} idleState The current idle state
 */
const setStatus = async (idleState: chrome.idle.IdleState) => {
  const status = document.querySelector('span.ts-skype-status') as Element;
  let statusTitle: any;
  if ('title' in status) {
    console.log('current status:', status?.title);
    statusTitle = status?.title;
  }

  /**
   * Get authentication token from local storage
   *
   * @returns the authentication token
   */
  const getAuthToken = (): string => {
    let token: string = '';

    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (key?.startsWith('ts.') && key?.endsWith('cache.token.https://presence.teams.microsoft.com/')) {
        const storageValue = localStorage.getItem(key);
        ({ token } = JSON.parse(storageValue as string));
      }
    }

    return token;
  };

  /**
   * Get the current endpointId from local storage
   *
   * @returns the current endpointId
   */
  const getEndpointId = () => {
    let endpointId: string = '';

    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (key?.startsWith('ts.') && key?.endsWith('userUpsEndpointIdsMap')) {
        const storageValue = localStorage.getItem(key);
        const storageObject = JSON.parse(storageValue as string);
        const endpointKey = Object.keys(storageObject)[0];

        // desc sort endpoint ids by lastActiveTimestamp
        storageObject[endpointKey].sort(
          (a: any, b: any) => b.lastActiveTimestamp - a.lastActiveTimestamp,
        );

        // eslint-disable-next-line prefer-destructuring
        ({ endpointId } = storageObject[endpointKey][0]);
      }
    }

    return endpointId;
  };

  /**
   * Submit teams requests to reflect status
   * @param {string} url The url to submit the request to
   * @param {string} body The payload of the request
   * @param {string} method The request method
   * @returns {Response} The response from the request
   */
  const submitRequest = async (
    url: string,
    body: string,
    method: string,
  ): Promise<Response> => {
    const response = await fetch(url, {
      headers: {
        Accept: 'json',
        'Content-Type': 'application/json',
        Authorization: `Bearer ${getAuthToken()}`,
      },
      body,
      method,
    });

    return response;
  };

  if (statusTitle === 'Away' && idleState === 'active') {
    const url = 'https://presence.teams.microsoft.com/v1/me/forceavailability/';
    const body = '{"availability":"Available"}';
    const method = 'PUT';

    const response = await submitRequest(url, body, method);

    if (response.status === 200) {
      console.log('availability set to Available');
    } else {
      console.log(
        `failed to submit available availability request: HTTP ${response.status}`,
      );
    }
  }

  if (
    idleState === 'locked'
    && status != null
    && statusTitle !== 'In a call'
    && statusTitle !== 'Busy'
    && statusTitle !== 'In a meeting'
    && statusTitle !== 'Out of Office'
  ) {
    const url = 'https://presence.teams.microsoft.com/v1/me/forceavailability/';
    const body = '{"availability":"Away"}';
    const method = 'PUT';

    const response = await submitRequest(url, body, method);

    if (response.status === 200) {
      console.log('availability set to Away');
    } else {
      console.log(
        `failed to submit away availability request: HTTP ${response.status}`,
      );
    }
  }

  if (
    status != null
    && statusTitle !== 'In a call'
    && statusTitle !== 'Busy'
    && statusTitle !== 'In a meeting'
    && statusTitle !== 'Out of Office'
    && idleState !== 'locked'
    && idleState !== 'idle'
  ) {
    const url = 'https://presence.teams.microsoft.com/v1/me/reportmyactivity/';
    const body = `{"endpointId": "${getEndpointId()}","isActive": true}`;
    const method = 'POST';

    const response = await submitRequest(url, body, method);

    if (response.status === 200) {
      console.log('sent reportmyactivity request');
    } else {
      console.log(
        `failed to submit report activity request: HTTP ${response.status}`,
      );
    }
  }
};

/**
 * Get the teams tab and set the status for teams
 *
 * @param {chrome.idle.IdleState} idleState The current idle state
 */
const getTeamsTabAndSetStatus = (idleState: chrome.idle.IdleState) => {
  chrome.tabs.query({ url: 'https://teams.microsoft.com/*' }, (tabs) => {
    tabs.forEach((tab) => {
      chrome.scripting.executeScript(
        {
          target: { tabId: tab.id as number },
          func: setStatus,
          args: [idleState],
        },
        () => {},
      );
    });
  });
};

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === alarmName) {
    chrome.idle.queryState(300, (newState: chrome.idle.IdleState) => {
      console.log('idle state:', newState);
      currentIdleState = newState;
    });

    getTeamsTabAndSetStatus(currentIdleState);
  }
});
