'use strict';

(function () {
  const DEFAULT_INTERVAL_SECONDS = 30;
  let refreshInterval = null;
  let activeDatasourceIdList = [];
  let uniqueDataSources = [];
  let isPaused = false;

  $(document).ready(function () {
    tableau.extensions.initializeAsync({ configure }).then(() => {
      loadSettings();
      tableau.extensions.settings.addEventListener(tableau.TableauEventType.SettingsChanged, (event) => {
        updateFromSettings(event.newSettings);
      });

      if (tableau.extensions.settings.get("configured") !== "1") {
        configure();
      }
    });
  });

  function loadSettings() {
    const settings = tableau.extensions.settings.getAll();

    if (settings.selectedDatasources) {
      activeDatasourceIdList = JSON.parse(settings.selectedDatasources);
    }

    const interval = settings.intervalkey ? parseInt(settings.intervalkey, 10) : DEFAULT_INTERVAL_SECONDS;

    if (settings.selectedDatasources) {
      $('#inactive').hide();
      $('#active').show();
      setupRefreshInterval(interval);
    }
  }

  function configure() {
    const popupUrl = `${window.location.origin}/AutoRefreshDialog.html`;
    const currentInterval = DEFAULT_INTERVAL_SECONDS;

    tableau.extensions.ui.displayDialogAsync(popupUrl, currentInterval.toString(), {
      height: 500,
      width: 500
    })
      .then((newInterval) => {
        $('#inactive').hide();
        $('#active').show();
        setupRefreshInterval(parseInt(newInterval, 10));
      })
      .catch((error) => {
        if (error.errorCode !== tableau.ErrorCodes.DialogClosedByUser) {
          console.error("Dialog error:", error.message);
        }
      });
  }

  function setupRefreshInterval(intervalSeconds) {
    if (refreshInterval) {
      clearTimeout(refreshInterval);
    }

    if (isPaused) return;

    startCircularTimer(intervalSeconds);

    function collectUniqueDataSources() {
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const seen = new Set();
      uniqueDataSources = [];

      const promises = dashboard.worksheets.map(worksheet =>
        worksheet.getDataSourcesAsync().then(datasources => {
          datasources.forEach(ds => {
            if (!seen.has(ds.id) && activeDatasourceIdList.includes(ds.id)) {
              seen.add(ds.id);
              uniqueDataSources.push(ds);
            }
          });
        })
      );

      return Promise.all(promises);
    }

    function executeRefresh() {
      if (uniqueDataSources.length === 0) {
        console.warn("No matching datasources to refresh.");
        scheduleNext();
        return;
      }

      const promises = uniqueDataSources.map(ds => ds.refreshAsync());

      Promise.all(promises)
        .then(() => {
          console.log(`✅ Refreshed ${uniqueDataSources.length} datasources.`);
          scheduleNext();
        })
        .catch(err => {
          console.error("Refresh failed:", err);
          scheduleNext();
        });
    }

    function scheduleNext() {
      if (!isPaused) {
        refreshInterval = setTimeout(executeRefresh, intervalSeconds * 1000);
      }
    }

    collectUniqueDataSources().then(() => {
      executeRefresh();
    });
  }

  function startCircularTimer(totalSeconds) {
    const timerRing = document.getElementById('progressRing');
    const timerText = document.getElementById('timerText');
    const circumference = 2 * Math.PI * 25;

    let secondsLeft = totalSeconds;

    timerText.textContent = secondsLeft;
    timerRing.style.strokeDasharray = `${circumference} ${circumference}`;
    timerRing.style.strokeDashoffset = circumference;

    const interval = setInterval(() => {
      secondsLeft--;

      timerText.textContent = secondsLeft;

      const offset = (secondsLeft / totalSeconds) * circumference;
      timerRing.style.strokeDashoffset = offset;

      if (secondsLeft <= 5) {
        timerText.style.color = '#dc3545';
      } else {
        timerText.style.color = '#495057';
      }

      if (secondsLeft <= 0) {
        clearInterval(interval);
        timerText.textContent = totalSeconds;
        timerText.style.color = '#495057';
        timerRing.style.strokeDashoffset = circumference;
      }
    }, 1000);
  }

  function updateFromSettings(settings) {
    if (settings.selectedDatasources) {
      activeDatasourceIdList = JSON.parse(settings.selectedDatasources);
    }
    const currentInterval = parseInt($('#interval').text?.(), 10) || DEFAULT_INTERVAL_SECONDS;
    setupRefreshInterval(currentInterval);
  }

  $(document).on('click', '#toggleRefresh', function () {
    isPaused = !isPaused;
    $(this).text(isPaused ? '▶️ Resume' : '⏸️ Pause');

    if (isPaused) {
      if (refreshInterval) {
        clearTimeout(refreshInterval);
        refreshInterval = null;
      }
    } else {
      const interval = DEFAULT_INTERVAL_SECONDS;
      setupRefreshInterval(interval);
    }
  });
})();