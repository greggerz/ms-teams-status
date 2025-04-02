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
const setStatus = async (idleState: chrome.idle.IdleState): Promise<void> => {
  const status = document.querySelector('[data-tid="me-control-avatar-trigger"]') as Element;
  const statusTitle = (status.getAttribute('aria-label')!).slice(46, -2);
  console.log(`current status: ${statusTitle}`);

  /**
   * Mimic a mouse move event to keep user presence Available
   */
  const mimicMouseEvent = () => {
    const clientX = 20;
    const clientY = window.innerHeight / 2;
    const movementX = 5;
    const movementY = 2;

    const mouseEvent = new MouseEvent('mousemove', {
      view: window,
      bubbles: true,
      clientX,
      clientY,
      movementX,
      movementY,
    });

    document.dispatchEvent(mouseEvent);
  };

  if (idleState === 'active') {
    console.log('calling mimicMouseEvent');
    mimicMouseEvent();
  }
};

/**
 * Get the teams tab and set the status for teams
 *
 * @param {chrome.idle.IdleState} idleState The current idle state
 */
const getTeamsTabAndSetStatus = (idleState: chrome.idle.IdleState): void => {
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

chrome.alarms.onAlarm.addListener((alarm): void => {
  if (alarm.name === alarmName) {
    chrome.idle.queryState(300, (newState: chrome.idle.IdleState) => {
      console.log('idle state:', newState);
      currentIdleState = newState;
    });

    getTeamsTabAndSetStatus(currentIdleState);
  }
});
