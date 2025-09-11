'use strict';

(function () {
  const DEFAULT_INTERVAL_SECONDS = 30;
  let refreshInterval = null;
  let activeDatasourceIdList = [];
  let uniqueDataSources = [];

  $(document).ready(function () {
    tableau.extensions.initializeAsync({ configure }).then(() => {
      loadSettings();
      tableau.extensions.settings.addEventListener(tableau.TableauEventType.SettingsChanged, (event) => {
        updateFromSettings(event.newSettings);
      });

      // If not configured yet → open dialog
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

    const interval = settings.intervalkey
      ? parseInt(settings.intervalkey, 10)
      : DEFAULT_INTERVAL_SECONDS;

    if (activeDatasourceIdList.length > 0) {
      $('#inactive').hide();
      $('#active').show();
      setupRefreshInterval(interval);
    }
  }

  function configure() {
    const popupUrl =`AutoRefreshDialog.html`;
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

    // Start circular timer in UI
    if (typeof window.startTimer === "function") {
      window.startTimer(intervalSeconds);
    }

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
        console.warn("⚠️ No matching datasources to refresh.");
        scheduleNext();
        return;
      }

      const promises = uniqueDataSources.map(ds => ds.refreshAsync());

      Promise.all(promises)
        .then(() => {
          console.log(`✅ Refreshed ${uniqueDataSources.length} datasource(s).`);
          scheduleNext();
        })
        .catch(err => {
          console.error("❌ Refresh failed:", err);
          scheduleNext();
        });
    }

    function scheduleNext() {
      refreshInterval = setTimeout(executeRefresh, intervalSeconds * 1000);
    }

    collectUniqueDataSources().then(() => {
      executeRefresh();
    });
  }

  function updateFromSettings(settings) {
    if (settings.selectedDatasources) {
      activeDatasourceIdList = JSON.parse(settings.selectedDatasources);
    }
    const interval = settings.intervalkey
      ? parseInt(settings.intervalkey, 10)
      : DEFAULT_INTERVAL_SECONDS;
    setupRefreshInterval(interval);
  }
})();
