'use strict';
(function(){
  const DEFAULT_INTERVAL_SECONDS = 60;
  const SETTINGS_KEY_DATASOURCES = 'selectedDatasources';
  const SETTINGS_KEY_INTERVAL = 'intervalkey';
  const SETTINGS_KEY_CONFIGURED = 'configured';

  let refreshTimeout = null;
  let activeDatasourceIdList = [];
  let uniqueDataSources = [];

  $(document).ready(function(){
    tableau.extensions.initializeAsync({configure}).then(()=>{
      loadSettings();
      tableau.extensions.settings.addEventListener(
        tableau.TableauEventType.SettingsChanged,
        (event) => updateFromSettings(event.newSettings)
      );

      if(tableau.extensions.settings.get(SETTINGS_KEY_CONFIGURED)!=="1"){
        configure();
      }
    });
  });

  function loadSettings(){
    const settings = tableau.extensions.settings.getAll();
    if(settings[SETTINGS_KEY_DATASOURCES]){
      activeDatasourceIdList = JSON.parse(settings[SETTINGS_KEY_DATASOURCES]);
    }
    const interval = settings[SETTINGS_KEY_INTERVAL]?parseInt(settings[SETTINGS_KEY_INTERVAL],10):DEFAULT_INTERVAL_SECONDS;
    if(activeDatasourceIdList.length>0){
      $('#inactive').hide();
      $('#active').show();
      setupRefreshLogic(interval);
    }
  }

  function configure(){
    const popupUrl = `${window.location.origin}/AutoRefreshDialog.html`;
    const currentInterval = tableau.extensions.settings.get(SETTINGS_KEY_INTERVAL) || DEFAULT_INTERVAL_SECONDS;

    tableau.extensions.ui.displayDialogAsync(popupUrl,currentInterval.toString(),{height:500,width:500})
      .then((newInterval)=>{
        $('#inactive').hide();
        $('#active').show();
        setupRefreshLogic(parseInt(newInterval,10));
      })
      .catch((error)=>{
        if(error.errorCode!==tableau.ErrorCodes.DialogClosedByUser){
          console.error("Dialog error:", error.message);
        }
      });
  }

  function setupRefreshLogic(intervalSeconds){
    if(refreshTimeout) clearTimeout(refreshTimeout);

    function collectUniqueDataSources(){
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const seen = new Set();
      uniqueDataSources=[];
      const promises = dashboard.worksheets.map(ws =>
        ws.getDataSourcesAsync().then(dsList=>{
          dsList.forEach(ds=>{
            if(!seen.has(ds.id) && activeDatasourceIdList.includes(ds.id)){
              seen.add(ds.id);
              uniqueDataSources.push(ds);
            }
          });
        })
      );
      return Promise.all(promises);
    }

    function executeRefresh(){
      if(uniqueDataSources.length===0){
        console.warn("No datasources selected.");
        scheduleNextRefresh();
        return;
      }

      const promises = uniqueDataSources.map(ds=>ds.refreshAsync());

      Promise.all(promises)
        .then(()=>{
          console.log("Refresh done.");
          if(typeof window.startTimer==="function") window.startTimer(intervalSeconds);
          if(typeof window.triggerPulse==="function") window.triggerPulse();
          scheduleNextRefresh();
        })
        .catch(err=>{
          console.error("Refresh failed:",err);
          if(typeof window.startTimer==="function") window.startTimer(intervalSeconds);
          if(typeof window.triggerPulse==="function") window.triggerPulse();
          scheduleNextRefresh();
        });
    }

    function scheduleNextRefresh(){
      refreshTimeout = setTimeout(executeRefresh, intervalSeconds*1000);
    }

    collectUniqueDataSources().then(()=>executeRefresh());
  }

  function updateFromSettings(settings){
    if(settings[SETTINGS_KEY_DATASOURCES]){
      activeDatasourceIdList = JSON.parse(settings[SETTINGS_KEY_DATASOURCES]);
    }
    const interval = settings[SETTINGS_KEY_INTERVAL]?parseInt(settings[SETTINGS_KEY_INTERVAL],10):DEFAULT_INTERVAL_SECONDS;
    if(activeDatasourceIdList.length>0){
      $('#inactive').hide();
      $('#active').show();
      setupRefreshLogic(interval);
    } else {
      if(refreshTimeout) clearTimeout(refreshTimeout);
      $('#active').hide();
      $('#inactive').show();
    }
  }

})();
