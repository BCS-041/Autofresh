'use strict';

(function () {
  const SETTINGS_KEY_DATASOURCES = 'selectedDatasources';
  const SETTINGS_KEY_INTERVAL = 'intervalkey';
  const SETTINGS_KEY_CONFIGURED = 'configured';

  const DEFAULT_INTERVAL = 30;
  let selectedDatasources = [];

  $(document).ready(function () {
    tableau.extensions.initializeDialogAsync().then((openPayload) => {
      // Set interval input
      const savedInterval = tableau.extensions.settings.get(SETTINGS_KEY_INTERVAL);
      const initialInterval = openPayload && !isNaN(parseInt(openPayload, 10))
        ? parseInt(openPayload, 10)
        : (savedInterval || DEFAULT_INTERVAL);

      $('#interval').val(initialInterval);

      // Load previously selected datasources
      selectedDatasources = loadSelectedDatasources();

      // Collect datasources across worksheets
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const seenDatasourceIds = new Set();
      let datasourceList = [];

      const promises = dashboard.worksheets.map(worksheet =>
        worksheet.getDataSourcesAsync().then(datasources => {
          datasources.forEach(ds => {
            if (!seenDatasourceIds.has(ds.id)) {
              seenDatasourceIds.add(ds.id);
              datasourceList.push(ds);
            }
          });
        })
      );

      Promise.all(promises).then(() => {
        // Sort alphabetically
        datasourceList.sort((a, b) => a.name.localeCompare(b.name));
        datasourceList.forEach(ds => {
          addDatasourceToUI(ds, selectedDatasources.includes(ds.id));
        });
      });

      // Save & Close
      $('#closeButton').on('click', saveAndClose);
    });
  });

  function loadSelectedDatasources() {
    const settings = tableau.extensions.settings.getAll();
    if (settings[SETTINGS_KEY_DATASOURCES]) {
      try {
        return JSON.parse(settings[SETTINGS_KEY_DATASOURCES]);
      } catch {
        return [];
      }
    }
    return [];
  }

  function addDatasourceToUI(datasource, isActive) {
    const container = $('<div />', {
      class: 'datasource-item',
      css: { marginBottom: '8px' }
    });

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

  function toggleDatasource(id) {
    const index = selectedDatasources.indexOf(id);
    if (index === -1) {
      selectedDatasources.push(id);
    } else {
      selectedDatasources.splice(index, 1);
    }
  }

  function saveAndClose() {
    const intervalValue = $('#interval').val().trim();
    const intervalNum = parseInt(intervalValue, 10);

    if (isNaN(intervalNum) || intervalNum < 15 || intervalNum > 3600) {
      $('#interval').css('border', '1px solid red').focus();
      alert("Please enter a valid interval between 15 and 3600 seconds.");
      return;
    }

    // Save settings
    tableau.extensions.settings.set(
      SETTINGS_KEY_DATASOURCES,
      JSON.stringify([...new Set(selectedDatasources)]) // remove duplicates
    );
    tableau.extensions.settings.set(SETTINGS_KEY_INTERVAL, intervalValue);
    tableau.extensions.settings.set(SETTINGS_KEY_CONFIGURED, "1");

    tableau.extensions.settings.saveAsync()
      .then(() => tableau.extensions.ui.closeDialog(intervalValue))
      .catch(err => {
        console.error("Failed to save settings:", err);
        alert("Error saving configuration. Please try again.");
      });
  }
})();
