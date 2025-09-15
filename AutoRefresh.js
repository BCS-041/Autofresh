'use strict';
(function () {
  const DEFAULT_INTERVAL_SECONDS = 60;
  const SETTINGS_KEY_DATASOURCES = 'selectedDatasources';
  const SETTINGS_KEY_INTERVAL = 'intervalkey';
  const SETTINGS_KEY_CONFIGURED = 'configured';

  let refreshTimeout = null;
  let activeDatasourceIdList = [];
  let uniqueDataSources = [];

  $(document).ready(function () {
    tableau.extensions.initializeAsync({ configure }).then(() => {
      loadSettings();

      tableau.extensions.settings.addEventListener(
        tableau.TableauEventType.SettingsChanged,
        (event) => {
          updateFromSettings(event.newSettings);
        }
      );

      if (tableau.extensions.settings.get(SETTINGS_KEY_CONFIGURED) !== "1") {
        configure();
      }
    });
  });

  // ---------------------------
  // Load saved settings
  // ---------------------------
  function loadSettings() {
    const settings = tableau.extensions.settings.getAll();

    if (settings[SETTINGS_KEY_DATASOURCES]) {
      activeDatasourceIdList = JSON.parse(settings[SETTINGS_KEY_DATASOURCES]);
    }

    const interval = settings[SETTINGS_KEY_INTERVAL]
      ? parseInt(settings[SETTINGS_KEY_INTERVAL], 10)
      : DEFAULT_INTERVAL_SECONDS;

    if (activeDatasourceIdList.length > 0) {
      $('#inactive').hide();
      $('#active').show();
      setupRefreshLogic(interval);
    }
  }

  // ---------------------------
  // Configure dialog
  // ---------------------------
  function configure() {
    const popupUrl = `${window.location.origin}/AutoRefreshDialog.html`;
    const currentInterval = tableau.extensions.settings.get(SETTINGS_KEY_INTERVAL) || DEFAULT_INTERVAL_SECONDS;

    console.log("Opening configuration dialog:", popupUrl);

    tableau.extensions.ui.displayDialogAsync(
      popupUrl,
      currentInterval.toString(),
      { height: 500, width: 500 }
    )
    .then((newInterval) => {
      console.log("Dialog closed with interval:", newInterval);

      $('#inactive').hide();
      $('#active').show();

      setupRefreshLogic(parseInt(newInterval, 10));
    })
    .catch((error) => {
      if (error.errorCode === tableau.ErrorCodes.DialogClosedByUser) {
        console.log("Dialog was closed by user");
      } else {
        console.error("Dialog error:", error.message);
      }
    });
  }

  // ---------------------------
  // Setup the refresh logic
  // ---------------------------
  function setupRefreshLogic(intervalSeconds) {
    if (refreshTimeout) {
      clearTimeout(refreshTimeout);
    }

    function collectUniqueDataSources() {
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const seen = new Set();
      uniqueDataSources = [];

      const promises = dashboard.worksheets.map((worksheet) =>
        worksheet.getDataSourcesAsync().then((datasources) => {
          datasources.forEach((ds) => {
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
        console.warn("⚠️ No matching datasources to refresh. Scheduling next.");
        scheduleNextRefresh();
        return;
      }

      console.log(`Starting refresh for ${uniqueDataSources.length} datasource(s).`);

      const promises = uniqueDataSources.map((ds) => ds.refreshAsync());

      Promise.all(promises)
        .then(() => {
          console.log(`✅ Refreshed ${uniqueDataSources.length} datasource(s).`);
          scheduleNextRefresh();
        })
        .catch((err) => {
          console.error("❌ Refresh failed:", err);
          scheduleNextRefresh();
        });
    }

    function scheduleNextRefresh() {
      // Start the visual timer countdown
      if (typeof window.startTimer === "function") {
        window.startTimer(intervalSeconds);
      }
      
      // Schedule the data refresh to happen after the interval
      refreshTimeout = setTimeout(executeRefresh, intervalSeconds * 1000);
    }

    collectUniqueDataSources().then(() => {
      // Begin the process by scheduling the first refresh
      scheduleNextRefresh();
    });
  }

  // ---------------------------
  // Update when settings change
  // ---------------------------
  function updateFromSettings(settings) {
    if (settings[SETTINGS_KEY_DATASOURCES]) {
      activeDatasourceIdList = JSON.parse(settings[SETTINGS_KEY_DATASOURCES]);
    }

    const interval = settings[SETTINGS_KEY_INTERVAL]
      ? parseInt(settings[SETTINGS_KEY_INTERVAL], 10)
      : DEFAULT_INTERVAL_SECONDS;

    if (activeDatasourceIdList.length > 0) {
      $('#inactive').hide();
      $('#active').show();
      setupRefreshLogic(interval);
    } else {
      // Stop the timer if no datasources are selected
      if (refreshTimeout) {
        clearTimeout(refreshTimeout);
      }
      // Hide the timer and show the inactive message
      $('#active').hide();
      $('#inactive').show();
    }
  }
})();
