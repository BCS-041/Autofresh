'use strict';
(function () {
  const SETTINGS_KEY_DATASOURCES = 'selectedDatasources';
  const SETTINGS_KEY_INTERVAL = 'intervalkey';
  const SETTINGS_KEY_CONFIGURED = 'configured';

  let selectedDatasources = [];

  $(document).ready(function () {
    tableau.extensions.initializeDialogAsync().then((openPayload) => {
      // Load interval from settings, payload, or fallback to 60
      const savedInterval = tableau.extensions.settings.get(SETTINGS_KEY_INTERVAL);
      const intervalValue = savedInterval || openPayload || 60;
      $('#interval').val(intervalValue);

      // Load datasources across all worksheets
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      selectedDatasources = getSavedDatasources();

      const seen = new Set();
      const promises = dashboard.worksheets.map(ws =>
        ws.getDataSourcesAsync().then(datasources => {
          datasources.forEach(ds => {
            if (!seen.has(ds.id)) {
              seen.add(ds.id);
              addDatasourceToUI(ds, selectedDatasources.includes(ds.id));
            }
          });
        })
      );

      Promise.all(promises).then(() => {
        console.log("✅ Datasources listed in config dialog");
      });

      // Close dialog handler
      $('#closeButton').on('click', saveAndClose);
    });
  });

  // ---------------------------
  // Helpers
  // ---------------------------
  function getSavedDatasources() {
    const settings = tableau.extensions.settings.getAll();
    return settings[SETTINGS_KEY_DATASOURCES]
      ? JSON.parse(settings[SETTINGS_KEY_DATASOURCES])
      : [];
  }

  function toggleDatasource(id) {
    const idx = selectedDatasources.indexOf(id);
    if (idx === -1) {
      selectedDatasources.push(id);
    } else {
      selectedDatasources.splice(idx, 1);
    }
  }

  function addDatasourceToUI(datasource, isActive) {
    const container = $('<div />').css({ marginBottom: '8px' });

    const checkbox = $('<input />', {
      type: 'checkbox',
      id: datasource.id,
      checked: isActive
    }).on('change', () => toggleDatasource(datasource.id));

    const label = $('<label />', {
      for: datasource.id,
      text: datasource.name,
      css: { marginLeft: '6px', cursor: 'pointer' }
    });

    container.append(checkbox).append(label);
    $('#datasources').append(container);
  }

  // ---------------------------
  // Save and close dialog
  // ---------------------------
  function saveAndClose() {
    const interval = $('#interval').val().trim();
    const intervalNum = parseInt(interval, 10);

    if (isNaN(intervalNum) || intervalNum < 15 || intervalNum > 3600) {
      alert("Please enter a valid interval between 15 and 3600 seconds.");
      $('#interval').focus();
      return;
    }

    tableau.extensions.settings.set(
      SETTINGS_KEY_DATASOURCES,
      JSON.stringify([...new Set(selectedDatasources)]) // dedupe
    );
    tableau.extensions.settings.set(SETTINGS_KEY_INTERVAL, intervalNum.toString());
    tableau.extensions.settings.set(SETTINGS_KEY_CONFIGURED, "1");

    tableau.extensions.settings.saveAsync()
      .then(() => {
        tableau.extensions.ui.closeDialog(intervalNum.toString());
      })
      .catch(err => {
        console.error("❌ Failed to save settings:", err);
        alert("Error saving configuration. Please try again.");
      });
  }
})();
