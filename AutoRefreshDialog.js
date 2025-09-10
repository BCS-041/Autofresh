'use strict';

(function () {
  const SETTINGS_KEY_DATASOURCES = 'selectedDatasources';
  const SETTINGS_KEY_INTERVAL = 'intervalkey';
  const SETTINGS_KEY_CONFIGURED = 'configured';

  let selectedDatasources = [];

  $(document).ready(function () {
    tableau.extensions.initializeDialogAsync().then((openPayload) => {
      const defaultInterval = 30;

      if (openPayload && !isNaN(parseInt(openPayload))) {
        $('#interval').val(openPayload);
      } else {
        const saved = tableau.extensions.settings.get(SETTINGS_KEY_INTERVAL);
        $('#interval').val(saved || defaultInterval);
      }

      selectedDatasources = loadSelectedDatasources();

      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const seenDatasourceIds = new Set();

      dashboard.worksheets.forEach(worksheet => {
        worksheet.getDataSourcesAsync().then(datasources => {
          datasources.forEach(datasource => {
            if (seenDatasourceIds.has(datasource.id)) return;
            seenDatasourceIds.add(datasource.id);

            const isActive = selectedDatasources.includes(datasource.id);
            addDatasourceToUI(datasource, isActive);
          });
        });
      });

      $('#closeButton').click(saveAndClose);
    });
  });

  function loadSelectedDatasources() {
    const settings = tableau.extensions.settings.getAll();
    if (settings[SETTINGS_KEY_DATASOURCES]) {
      return JSON.parse(settings[SETTINGS_KEY_DATASOURCES]);
    }
    return [];
  }

  function addDatasourceToUI(datasource, isActive) {
    const container = $('<div />');

    $('<input />', {
      type: 'checkbox',
      id: datasource.id,
      checked: isActive,
      click: () => toggleDatasource(datasource.id)
    }).appendTo(container);

    $('<label />', {
      'for': datasource.id,
      text: datasource.name
    }).appendTo(container);

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
      alert("Please enter a valid interval between 15 and 3600 seconds.");
      return;
    }

    tableau.extensions.settings.set(SETTINGS_KEY_DATASOURCES, JSON.stringify(selectedDatasources));
    tableau.extensions.settings.set(SETTINGS_KEY_INTERVAL, intervalValue);
    tableau.extensions.settings.set(SETTINGS_KEY_CONFIGURED, "1");

    tableau.extensions.settings.saveAsync()
      .then(() => {
        tableau.extensions.ui.closeDialog(intervalValue);
      })
      .catch(err => {
        console.error("Failed to save settings:", err);
        alert("Error saving configuration. Please try again.");
      });
  }
})();