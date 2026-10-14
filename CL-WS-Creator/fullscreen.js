/*!
 * Copyright © 2026 Eithan Meyer. Full-screen UI extension for CL - WS Creator.
 * All rights reserved to the extent of the rights owned by Eithan Meyer.
 */
(function () {
  'use strict';

  var inFullscreen = new URLSearchParams(window.location.search).get('fullscreen') === '1';
  var fullscreenDialog = null;
  var reportSequence = 0;
  var pendingReports = Object.create(null);

  function setStatus(message, bad) {
    var status = document.getElementById('buildStatus');
    var pill = document.getElementById('statusPill');
    if (status) status.textContent = message;
    if (pill) {
      pill.textContent = bad ? 'Full-screen error' : 'Full-screen ready';
      pill.className = 'pill ' + (bad ? 'bad' : 'ok');
    }
  }

  function safeMessageParent(payload) {
    if (!window.Office || !Office.context || !Office.context.ui ||
        typeof Office.context.ui.messageParent !== 'function') {
      throw new Error('The Office dialog message channel is unavailable.');
    }
    Office.context.ui.messageParent(JSON.stringify(payload));
  }

  function handleDialogParentMessage(arg) {
    var message;
    try { message = JSON.parse(arg.message); } catch (e) { return; }
    if (message.type !== 'CLWSC_REPORT_ACK') return;
    var pending = pendingReports[message.requestId];
    if (!pending) return;
    clearTimeout(pending.timer);
    delete pendingReports[message.requestId];
    if (message.success) pending.resolve();
    else pending.reject(new Error(message.error || 'Excel report sheets could not be updated from full-screen mode.'));
  }

  function sendReportToParent(result, job, selectedSheetNames, actionName) {
    return new Promise(function (resolve, reject) {
      var requestId = 'report-' + Date.now() + '-' + (++reportSequence);
      var pending = {
        resolve: resolve,
        reject: reject,
        timer: window.setTimeout(function () {
          if (!pendingReports[requestId]) return;
          delete pendingReports[requestId];
          reject(new Error('Timed out waiting for Excel to finish writing the report sheets. The operation may have completed in SigmaNEST; check the workbook before retrying.'));
        }, 120000)
      };
      pendingReports[requestId] = pending;
      try {
        safeMessageParent({
          type: 'CLWSC_REPORT',
          requestId: requestId,
          result: result,
          job: job || '',
          selectedSheetNames: selectedSheetNames || [],
          actionName: actionName || 'Operation',
          clFileName: window.clFile && window.clFile.name || '',
          wsPath: result && result.wsPath || window.currentWsPath || ''
        });
      } catch (error) {
        clearTimeout(pending.timer);
        delete pendingReports[requestId];
        reject(error);
      }
    });
  }

  if (typeof window.writeReportSheets === 'function') {
    var originalWriteReportSheets = window.writeReportSheets;
    window.writeReportSheets = async function (result, job, selectedSheetNames, actionName) {
      if (!inFullscreen) return originalWriteReportSheets.apply(this, arguments);
      return sendReportToParent(result, job, selectedSheetNames, actionName);
    };
  }

  function closeFullscreen() {
    try { safeMessageParent({ type: 'CLWSC_FULLSCREEN_CLOSE' }); }
    catch (e) { setStatus(e.message || String(e), true); }
  }

  function handleDialogMessage(arg) {
    var message;
    try { message = JSON.parse(arg.message); } catch (e) { return; }
    if (arg.origin && arg.origin !== window.location.origin) return;
    if (message.type === 'CLWSC_FULLSCREEN_CLOSE') {
      if (fullscreenDialog) fullscreenDialog.close();
      return;
    }
    if (message.type !== 'CLWSC_REPORT') return;

    (async function () {
      var oldFile = window.clFile;
      var oldPath = window.currentWsPath;
      try {
        window.clFile = message.clFileName ? { name: message.clFileName } : null;
        window.currentWsPath = message.wsPath || oldPath || '';
        await window.writeReportSheets(message.result || {}, message.job || '', message.selectedSheetNames || [], message.actionName || 'Operation');
        fullscreenDialog.messageChild(JSON.stringify({
          type: 'CLWSC_REPORT_ACK',
          requestId: message.requestId,
          success: true
        }));
      } catch (error) {
        try {
          fullscreenDialog.messageChild(JSON.stringify({
            type: 'CLWSC_REPORT_ACK',
            requestId: message.requestId,
            success: false,
            error: error && error.message ? error.message : String(error)
          }));
        } catch (e) {}
      } finally {
        window.clFile = oldFile;
        window.currentWsPath = oldPath;
      }
    })();
  }

  function attachOpenButton() {
    var openButton = document.getElementById('openFullscreen');
    if (!openButton) return;
    openButton.addEventListener('click', function () {
      if (!window.Office || !Office.context || !Office.context.ui ||
          typeof Office.context.ui.displayDialogAsync !== 'function') {
        setStatus('Full-screen mode requires the add-in to be opened inside Excel.', true);
        return;
      }
      var target = window.location.origin + window.location.pathname + '?v=2.32.36&fullscreen=1';
      try {
        Office.context.ui.displayDialogAsync(target, { height: 99.5, width: 99.5 }, function (asyncResult) {
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            setStatus('Could not open the full-screen workspace: ' +
              (asyncResult.error && asyncResult.error.message || 'Office did not create the dialog.'), true);
            return;
          }
          fullscreenDialog = asyncResult.value;
          fullscreenDialog.addEventHandler(Office.EventType.DialogMessageReceived, handleDialogMessage);
          fullscreenDialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
            fullscreenDialog = null;
          });
        });
      } catch (error) {
        setStatus('Could not open the full-screen workspace: ' + (error.message || String(error)), true);
      }
    });
  }

  function enterFullscreenMode() {
    document.body.classList.add('fullscreenMode');
    var openButton = document.getElementById('openFullscreen');
    var closeButton = document.getElementById('closeFullscreen');
    if (openButton) openButton.hidden = true;
    if (closeButton) {
      closeButton.hidden = false;
      closeButton.addEventListener('click', closeFullscreen);
    }
    if (!window.Office || !Office.context || !Office.context.ui ||
        typeof Office.context.ui.addHandlerAsync !== 'function') {
      setStatus('Full-screen mode could not connect to Excel. Close this window and reopen it from the add-in.', true);
      return;
    }
    Office.context.ui.addHandlerAsync(
      Office.EventType.DialogParentMessageReceived,
      handleDialogParentMessage,
      function (asyncResult) {
        if (asyncResult && asyncResult.status === Office.AsyncResultStatus.Failed) {
          setStatus('Full-screen report hand-off could not be registered: ' +
            (asyncResult.error && asyncResult.error.message || 'Unknown error.'), true);
        }
      }
    );
  }

  if (window.Office && Office.onReady) {
    Office.onReady(function () {
      if (inFullscreen) enterFullscreenMode();
      else attachOpenButton();
    });
  } else {
    window.setTimeout(function () {
      if (inFullscreen) enterFullscreenMode();
      else attachOpenButton();
    }, 2500);
  }
})();