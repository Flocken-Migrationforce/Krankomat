

window.Krankomat = window.Krankomat || {};

Krankomat.App = {
    init: function() {
        console.log('Initializing Krankomat...');
        // Initialize modules
        Krankomat.State.init();
        Krankomat.Colors.init();
        Krankomat.Builder.init();
        Krankomat.Preview.init();

        // Initial Render
        this.renderAll();

        // Subscribe to state changes
        Krankomat.State.subscribe(() => {
            this.renderAll();
        });

        this.setupModals();
        this.checkDisclaimer();
        this.checkConfigBanner();
        console.log('Krankomat initialized successfully.');
    },

    renderAll: function() {
        Krankomat.Builder.render();
        Krankomat.Preview.render();
        this.renderHeaderVisibility();
        this.renderHeaderTitle();
    },

    renderHeaderTitle: function() {
        const config = Krankomat.State.get('config') || {};
        const titleSuffixEl = document.getElementById('app-title-suffix');
        if (titleSuffixEl) {
            titleSuffixEl.innerText = config.profileName || 'WebApp';
        }
    },

    renderHeaderVisibility: function() {
        const config = Krankomat.State.get('config') || {};
        const buttons = config.headerButtons || { fileshare: true, calendar: true, mensa: true };

        const setVisibility = (id, isVisible) => {
            const el = document.getElementById(id);
            if (el) {
                if (isVisible) el.classList.remove('hidden');
                else el.classList.add('hidden');
            }
        };

        setVisibility('fileshare-btn', buttons.fileshare);
        setVisibility('calendar-toggle-btn', buttons.calendar);
        setVisibility('mensa-toggle-btn', buttons.mensa);
        // Theme toggle is now in settings modal only
    },

    checkDisclaimer: function() {
        const dismissed = localStorage.getItem('krankomat_disclaimerDismissed');
        const container = document.getElementById('disclaimer-container');
        if (container && dismissed !== 'true') {
            container.classList.remove('hidden');
            const btn = document.getElementById('dismiss-disclaimer-btn');
            if (btn) {
                btn.addEventListener('click', () => {
                    localStorage.setItem('krankomat_disclaimerDismissed', 'true');
                    container.classList.add('hidden');
                });
            }
        }
    },

    checkConfigBanner: function() {
        const dismissed = localStorage.getItem('krankomat_configBannerDismissed');
        const container = document.getElementById('config-banner-container');
        if (container && dismissed !== 'true') {
            container.classList.remove('hidden');
            
            const dismissBtn = document.getElementById('dismiss-config-banner-btn');
            if (dismissBtn) {
                dismissBtn.addEventListener('click', () => {
                    localStorage.setItem('krankomat_configBannerDismissed', 'true');
                    container.classList.add('hidden');
                });
            }

            const settingsLink = document.getElementById('open-settings-shortcut');
            if (settingsLink) {
                settingsLink.addEventListener('click', () => {
                    const settingsContainer = document.getElementById('settings-modal-container');
                    this.renderSettingsModal(settingsContainer);
                });
            }
        }
    },

    setupModals: function() {
        this.setupTermsModal();
        this.setupSettingsModal();
        this.setupMensaModal();
        this.setupCalendarModal();
    },

    setupTermsModal: function() {
        const SKIP_TERMS_CHECK = false; // Set to true if you want to skip
        const TERMS_VERSION = '20.11.2025';
        const STORAGE_KEY = 'krankomat_agb_accepted_version';
        
        if (SKIP_TERMS_CHECK) return;

        const accepted = localStorage.getItem(STORAGE_KEY);
        if (accepted === TERMS_VERSION) return;

        // Render modal
        const container = document.getElementById('terms-modal-container');
        if (!container) return;

        container.innerHTML = `
          <div class="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-md transition-all duration-500">
             <div class="bg-white dark:bg-slate-800 rounded-2xl shadow-2xl max-w-md w-full flex flex-col border border-slate-200 dark:border-slate-700 animate-scale-in">
                <div class="p-6 border-b border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 rounded-t-2xl text-center">
                   <div class="mx-auto h-12 w-12 bg-indigo-100 dark:bg-indigo-900/50 rounded-full flex items-center justify-center mb-4">
                     <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" class="w-6 h-6 text-indigo-600 dark:text-indigo-400">
                       <path fillRule="evenodd" d="M2.25 12c0-5.385 4.365-9.75 9.75-9.75s9.75 4.365 9.75 9.75-4.365 9.75-9.75 9.75S2.25 17.385 2.25 12zm8.706-1.442c1.146-.573 2.437.463 2.126 1.706l-.709 2.836.042-.02a.75.75 0 01.67 1.34l-.04.022c-1.147.573-2.438-.463-2.127-1.706l.71-2.836-.042.02a.75.75 11-.671-1.34l.041-.022zM12 9a.75.75 100-1.5.75.75 0 000 1.5z" clipRule="evenodd" />
                     </svg>
                   </div>
                   <h2 class="text-xl font-bold text-slate-800 dark:text-white">Nutzungsbedingungen</h2>
                   <p class="text-xs text-slate-500 dark:text-slate-400 font-mono mt-1">${TERMS_VERSION}</p>
                </div>
                <div class="p-8 bg-white dark:bg-slate-800 text-center space-y-6">
                   <p class="text-slate-600 dark:text-slate-300 text-sm leading-relaxed">
                     Bitte lesen Sie unsere Nutzungsbedingungen sorgfältig durch, bevor Sie fortfahren. Mit der Nutzung dieser Software akzeptieren Sie die Bedingungen.
                   </p>
                   <a href="./NUTZUNGSBEDINGUNGEN.txt" target="_blank" class="inline-flex items-center justify-center w-full px-4 py-3 bg-slate-100 dark:bg-slate-700 hover:bg-slate-200 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 rounded-lg transition-colors group border border-slate-200 dark:border-slate-600">
                     <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" class="w-5 h-5 mr-3 text-slate-500 group-hover:text-indigo-500 transition-colors">
                       <path d="M5.625 1.5c-1.036 0-1.875.84-1.875 1.875v17.25c0 1.035.84 1.875 1.875 1.875h12.75c1.035 0 1.875-.84 1.875-1.875V12.75A3.75 3.75 0 0016.5 9h-1.875a1.875 1.875 0 01-1.875-1.875V5.25A3.75 3.75 0 009 1.5H5.625z" />
                       <path d="M12.971 1.816A5.23 5.23 0 0114.25 5.25v1.875c0 .207.168.375.375.375H16.5a5.23 5.23 0 013.434 1.279 9.768 9.768 0 00-6.963-6.963z" />
                     </svg>
                     Nutzungsbedingungen lesen
                   </a>
                </div>
                <div class="p-6 border-t border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/50 rounded-b-2xl">
                   <button id="accept-terms-btn" class="w-full px-6 py-3 bg-indigo-600 hover:bg-indigo-700 text-white font-medium rounded-lg shadow-lg transition-all transform hover:-translate-y-0.5">
                     Gelesen & Akzeptieren
                   </button>
                </div>
             </div>
          </div>
        `;
        
        // Blur main
        const main = document.getElementById('main-content');
        if (main) main.classList.add('blur-sm', 'pointer-events-none');

        const btn = document.getElementById('accept-terms-btn');
        if (btn) {
            btn.addEventListener('click', () => {
                localStorage.setItem(STORAGE_KEY, TERMS_VERSION);
                container.innerHTML = '';
                if (main) main.classList.remove('blur-sm', 'pointer-events-none');
            });
        }
    },

    setupSettingsModal: function() {
        const toggle = () => {
            const container = document.getElementById('settings-modal-container');
            if (container.innerHTML) {
                container.innerHTML = '';
            } else {
                this.renderSettingsModal(container);
            }
        };

        const btn = document.getElementById('settings-toggle-btn');
        if (btn) btn.addEventListener('click', toggle);
    },

    renderSettingsModal: function(container) {
        // Retrieve current config
        const config = Krankomat.State.get('config') || {};
        const buttons = config.headerButtons || { fileshare: true, calendar: true, mensa: true };
        const supportEmail = config.supportEmail || "support@krankomat.cloud";
        const showAll = config.showAllRecipients || false;
        const currentColor = config.colorTheme || 'indigo';
        
        const palettes = ['indigo', 'blue', 'rose', 'emerald', 'amber', 'violet'];

        container.innerHTML = `
          <div class="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-sm transition-all" id="expert-backdrop">
             <div class="bg-white dark:bg-slate-800 rounded-xl shadow-2xl max-w-lg w-full border border-indigo-500/30 animate-scale-in" onclick="event.stopPropagation()">
                <div class="p-5 border-b border-slate-200 dark:border-slate-700 flex justify-between items-center">
                   <h2 class="text-lg font-bold text-slate-800 dark:text-white flex items-center">
                     <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6 mr-2 text-indigo-600 dark:text-indigo-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                     </svg>
                     Einstellungen
                   </h2>
                   <button id="expert-close-btn" class="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200">
                     <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" class="w-6 h-6">
                       <path fillRule="evenodd" d="M5.47 5.47a.75.75 0 011.06 0L12 10.94l5.47-5.47a.75.75 0 111.06 1.06L13.06 12l5.47 5.47a.75.75 0 11-1.06 1.06L12 13.06l-5.47 5.47a.75.75 0 01-1.06-1.06L10.94 12 5.47 6.53a.75.75 0 010-1.06z" clipRule="evenodd" />
                     </svg>
                   </button>
                </div>
                <div class="p-6 space-y-6 max-h-[70vh] overflow-y-auto">
                
                  <!-- App Theme -->
                  <div class="p-4 bg-slate-50 dark:bg-slate-700/50 rounded-lg border border-slate-200 dark:border-slate-600">
                     <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200 mb-3">Darstellung</h3>
                     
                     <div class="flex items-center justify-between mb-4">
                         <span class="text-sm text-slate-600 dark:text-slate-300">Dark Mode</span>
                         <button id="theme-toggle-btn" type="button" class="relative inline-flex items-center h-8 w-14 cursor-pointer rounded-full bg-slate-200 dark:bg-slate-700 transition-colors duration-300">
                            <span id="theme-toggle-indicator" class="translate-x-1 absolute left-0 inline-flex h-6 w-6 transform items-center justify-center rounded-full bg-white shadow-lg transition-transform duration-300 ease-in-out">
                                <!-- Icon injected by JS -->
                            </span>
                         </button>
                     </div>
                     
                     <hr class="my-3 border-slate-200 dark:border-slate-600">
                     
                     <h4 class="text-xs font-semibold text-slate-600 dark:text-slate-400 mb-2 uppercase tracking-wide">Farbthema</h4>
                     <div class="flex flex-wrap gap-3">
                        ${palettes.map(color => `
                            <button class="w-8 h-8 rounded-full border-2 transition-all duration-200 flex items-center justify-center color-theme-btn" 
                                style="background-color: var(--color-${color}); ${currentColor === color ? 'border-color: var(--color-slate-400); transform: scale(1.1);' : 'border-color: transparent;'}"
                                data-color="${color}" 
                                title="${color.charAt(0).toUpperCase() + color.slice(1)}">
                                ${currentColor === color ? '<svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4 text-white drop-shadow-md" viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clip-rule="evenodd" /></svg>' : ''}
                            </button>
                        `).join('')}
                     </div>
                     <!-- Hidden styles to help render the buttons above correctly using hardcoded values if vars fail, but mostly relying on colors.js -->
                     <style>
                        :root {
                            --color-indigo: #4f46e5; --color-blue: #2563eb; --color-rose: #e11d48; 
                            --color-emerald: #059669; --color-amber: #d97706; --color-violet: #7c3aed;
                        }
                     </style>
                  </div>

                  <!-- UI Settings Section (Merged) -->
                  <div class="p-4 bg-slate-50 dark:bg-slate-700/50 rounded-lg border border-slate-200 dark:border-slate-600">
                     <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200 mb-3">Benutzeroberfläche</h3>
                     <label class="flex items-center space-x-2 mb-2">
                        <input type="checkbox" id="setting-show-all-recipients" class="rounded border-slate-300 text-indigo-600 shadow-sm focus:border-indigo-300 focus:ring focus:ring-indigo-200 focus:ring-opacity-50" ${showAll ? 'checked' : ''}>
                        <span class="text-sm text-slate-600 dark:text-slate-300">Alle bekannten Empfänger anzeigen (statt nur Tagesauswahl)</span>
                     </label>

                     <hr class="my-3 border-slate-200 dark:border-slate-600">
                     
                     <h4 class="text-xs font-semibold text-slate-600 dark:text-slate-400 mb-2 uppercase tracking-wide">Header Buttons</h4>
                     <div class="space-y-2">
                        <label class="flex items-center space-x-2">
                            <input type="checkbox" class="header-toggle rounded border-slate-300 text-indigo-600 shadow-sm focus:border-indigo-300 focus:ring focus:ring-indigo-200 focus:ring-opacity-50" data-target="fileshare" ${buttons.fileshare ? 'checked' : ''}>
                            <span class="text-sm text-slate-600 dark:text-slate-300">Dateifreigabe (Cloud-Link)</span>
                        </label>
                        <label class="flex items-center space-x-2">
                            <input type="checkbox" class="header-toggle rounded border-slate-300 text-indigo-600 shadow-sm focus:border-indigo-300 focus:ring focus:ring-indigo-200 focus:ring-opacity-50" data-target="calendar" ${buttons.calendar ? 'checked' : ''}>
                            <span class="text-sm text-slate-600 dark:text-slate-300">Stundenplan (Kalender)</span>
                        </label>
                        <label class="flex items-center space-x-2">
                            <input type="checkbox" class="header-toggle rounded border-slate-300 text-indigo-600 shadow-sm focus:border-indigo-300 focus:ring focus:ring-indigo-200 focus:ring-opacity-50" data-target="mensa" ${buttons.mensa ? 'checked' : ''}>
                            <span class="text-sm text-slate-600 dark:text-slate-300">Mensa Menü</span>
                        </label>
                    </div>
                  </div>
                  
                  <!-- Support Section -->
                  <div class="p-4 bg-slate-50 dark:bg-slate-700/50 rounded-lg border border-slate-200 dark:border-slate-600">
                     <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200 mb-2">Hilfe & Support</h3>
                     <p class="text-xs text-slate-500 dark:text-slate-400 mb-3">
                        Haben Sie Fragen oder Feedback? Kontaktieren Sie uns.
                     </p>
                     <a href="mailto:${supportEmail}?subject=Feedback%20Krankomat%20WebApp" class="w-full inline-flex items-center justify-center px-4 py-2 border border-slate-300 dark:border-slate-500 text-sm font-medium rounded-md shadow-sm text-slate-700 dark:text-slate-200 bg-white dark:bg-slate-600 hover:bg-slate-50 dark:hover:bg-slate-500 transition-colors">
                        <svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4 mr-2" viewBox="0 0 20 20" fill="currentColor">
                           <path d="M2.003 5.884L10 9.882l7.997-3.998A2 2 0 0016 4H4a2 2 0 00-1.997 1.884z" />
                           <path d="M18 8.118l-8 4-8-4V14a2 2 0 002 2h12a2 2 0 002-2V8.118z" />
                        </svg>
                        E-Mail senden
                     </a>
                  </div>

                  <!-- Backup Section -->
                  <div class="p-4 bg-slate-50 dark:bg-slate-700/50 rounded-lg border border-slate-200 dark:border-slate-600">
                    <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200 mb-2">Sicherung & Vorlage</h3>
                    <p class="text-xs text-slate-500 dark:text-slate-400 mb-3">
                      Laden Sie Ihre aktuelle Konfiguration (inkl. Kalender & UI-Einstellungen) als JSON-Datei herunter. Nutzen Sie dies als Backup.
                    </p>
                    <button id="expert-export-btn" class="w-full py-2 bg-white dark:bg-slate-600 border border-slate-300 dark:border-slate-500 text-slate-700 dark:text-slate-200 rounded-md text-sm font-medium hover:bg-slate-50 dark:hover:bg-slate-500 transition-colors flex items-center justify-center">
                      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor" class="w-4 h-4 mr-2">
                        <path fillRule="evenodd" d="M10 3a.75.75 0 01.75.75v10.638l3.96-4.158a.75.75 0 111.08 1.04l-5.25 5.5a.75.75 0 01-1.08 0l-5.25-5.5a.75.75 0 111.08-1.04l3.96 4.158V3.75A.75.75 0 0110 3z" clipRule="evenodd" />
                      </svg>
                      Konfiguration herunterladen (JSON)
                    </button>
                  </div>

                  <!-- Import Section -->
                  <div class="p-4 bg-indigo-50 dark:bg-indigo-900/20 rounded-lg border border-indigo-100 dark:border-indigo-800/30">
                    <h3 class="text-sm font-bold text-indigo-800 dark:text-indigo-300 mb-2">Konfiguration importieren</h3>
                    <p class="text-xs text-indigo-600 dark:text-indigo-400 mb-3">
                      Laden Sie eine JSON-Datei hoch, um den aktuellen Status zu überschreiben. <strong>Warnung: Die Seite wird neu geladen.</strong>
                    </p>
                    <label class="w-full flex flex-col items-center px-4 py-4 bg-white dark:bg-slate-700 text-indigo-600 dark:text-indigo-300 rounded-lg shadow-sm tracking-wide uppercase border border-indigo-200 dark:border-slate-600 cursor-pointer hover:bg-indigo-50 dark:hover:bg-slate-600 transition-colors">
                        <svg class="w-6 h-6" fill="currentColor" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path d="M16.88 9.1A4 4 0 0 1 16 17H5a5 5 0 0 1-1-9.9V7a3 3 0 0 1 4.52-2.59A4.98 4.98 0 0 1 17 8c0 .38-.04.74-.12 1.1zM11 11h3l-4-4-4 4h3v3h2v-3z" /></svg>
                        <span class="mt-2 text-xs font-bold">JSON Datei auswählen</span>
                        <input type='file' id="expert-file-input" class="hidden" accept=".json" />
                    </label>
                  </div>

                  <!-- ICS Import (Moved to bottom) -->
                  <div class="p-4 bg-slate-50 dark:bg-slate-700/50 rounded-lg border border-slate-200 dark:border-slate-600">
                     <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200 mb-3">Stundenplan Import</h3>
                     <p class="text-xs text-slate-500 dark:text-slate-400 mb-3">
                        Laden Sie eine .ics Datei Ihres Stundenplans hoch.
                     </p>
                     <label class="w-full flex flex-col items-center px-4 py-3 bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 rounded-lg shadow-sm tracking-wide uppercase border border-indigo-200 dark:border-indigo-900/50 cursor-pointer hover:bg-indigo-50 dark:hover:bg-slate-700 transition-colors">
                        <svg class="w-6 h-6" fill="currentColor" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zM6.293 6.707a1 1 0 010-1.414l3-3a1 1 0 011.414 0l3 3a1 1 0 01-1.414 1.414L11 5.414V13a1 1 0 11-2 0V5.414L7.707 6.707a1 1 0 01-1.414 0z" clip-rule="evenodd" /></svg>
                        <span class="mt-2 text-xs font-bold">.ics Datei auswählen</span>
                        <input type='file' id="settings-ics-upload" class="hidden" accept=".ics" />
                    </label>
                  </div>
                </div>
             </div>
          </div>
        `;

        document.getElementById('expert-close-btn').onclick = () => container.innerHTML = '';
        document.getElementById('expert-backdrop').onclick = () => container.innerHTML = '';

        // Bind Theme Toggle
        const themeBtn = document.getElementById('theme-toggle-btn');
        if (themeBtn) {
            themeBtn.onclick = (e) => {
                e.preventDefault();
                Krankomat.Colors.toggle();
            };
        }
        Krankomat.Colors.renderToggle(); // Update icon state immediately
        
        // Bind Color Theme Buttons
        container.querySelectorAll('.color-theme-btn').forEach(btn => {
            btn.onclick = (e) => {
                const color = e.currentTarget.dataset.color;
                Krankomat.Colors.setPalette(color);
                // Re-render settings modal to update selection UI
                this.renderSettingsModal(container);
            };
        });

        // Bind Show All Recipients Toggle
        const showAllCheck = document.getElementById('setting-show-all-recipients');
        if (showAllCheck) {
            showAllCheck.addEventListener('change', (e) => {
                Krankomat.State.updateNested('config', 'showAllRecipients', e.target.checked);
                Krankomat.Builder.recalculateRecipients(); // Force recalculation to update list
            });
        }

        // Bind ICS Upload in Settings
        const icsInput = document.getElementById('settings-ics-upload');
        if (icsInput) {
            icsInput.addEventListener('change', (e) => {
                const file = e.target.files[0];
                if (file) {
                    Krankomat.Builder.handleIcsUpload(file);
                }
            });
        }

        // Bind Toggle Events
        container.querySelectorAll('.header-toggle').forEach(toggle => {
            toggle.addEventListener('change', (e) => {
                const target = e.target.dataset.target;
                const isChecked = e.target.checked;
                // Update State
                Krankomat.State.updateNested('config', 'headerButtons', { 
                    ...config.headerButtons,
                    [target]: isChecked 
                });
            });
        });

        document.getElementById('expert-export-btn').onclick = () => {
            const config = Krankomat.State.data;
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(config, null, 2));
            const node = document.createElement('a');
            node.setAttribute("href", dataStr);
            node.setAttribute("download", "krankomat_config.json");
            document.body.appendChild(node);
            node.click();
            node.remove();
        };

        document.getElementById('expert-file-input').onchange = (e) => {
            const file = e.target.files[0];
            if (!file) return;

            // Check filename for [Name]_config.json pattern to update profile name
            const filename = file.name;
            const nameMatch = filename.match(/^(.*)_config\.json$/);
            
            const reader = new FileReader();
            reader.onload = (ev) => {
                try {
                    const json = JSON.parse(ev.target.result);
                    if (json.userData) localStorage.setItem('krankomat_userData', JSON.stringify(json.userData));
                    if (json.recipients) localStorage.setItem('krankomat_recipients', JSON.stringify(json.recipients));
                    if (json.sicknessStartDate) localStorage.setItem('krankomat_sicknessStartDate', JSON.stringify(json.sicknessStartDate));
                    if (json.sicknessEndDate) localStorage.setItem('krankomat_sicknessEndDate', JSON.stringify(json.sicknessEndDate));
                    if (json.absenceReasons) localStorage.setItem('krankomat_absenceReasons', JSON.stringify(json.absenceReasons));
                    if (json.details) localStorage.setItem('krankomat_details', JSON.stringify(json.details));
                    if (json.calendarEvents) localStorage.setItem('krankomat_calendarEvents', JSON.stringify(json.calendarEvents));
                    if (json.emailDirectory) localStorage.setItem('krankomat_emailDirectory', JSON.stringify(json.emailDirectory));
                    
                    if (json.config) {
                        // If file matched pattern, override/set profileName
                        if (nameMatch && nameMatch[1]) {
                            json.config.profileName = nameMatch[1];
                        }
                        localStorage.setItem('krankomat_config', JSON.stringify(json.config));
                    } else if (nameMatch && nameMatch[1]) {
                        // Create config if missing but name is present
                        const defaultConfig = Krankomat.State.defaults.config;
                        defaultConfig.profileName = nameMatch[1];
                        localStorage.setItem('krankomat_config', JSON.stringify(defaultConfig));
                    }
                    
                    window.location.reload();
                } catch (err) {
                    alert("Fehler beim Verarbeiten der JSON Datei");
                }
            };
            reader.readAsText(file);
        };
    },

    setupMensaModal: function() {
        // ... (rest of the file remains unchanged)
        const toggleBtn = document.getElementById('mensa-toggle-btn');
        if (toggleBtn) {
            toggleBtn.addEventListener('click', () => {
                const container = document.getElementById('mensa-modal-container');
                if (container.innerHTML) {
                    container.innerHTML = '';
                } else {
                    this.renderMensaModal(container);
                }
            });
        }
    },

    renderMensaModal: function(container) {
        container.innerHTML = `
            <div class="fixed inset-0 bg-black bg-opacity-60 z-50 flex justify-center items-center p-4 transition-opacity duration-300" id="mensa-backdrop">
                <div class="bg-white dark:bg-slate-800 rounded-2xl shadow-2xl w-full max-w-3xl max-h-[90vh] flex flex-col animate-scale-in" onclick="event.stopPropagation()">
                    <header class="p-4 sm:p-5 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between flex-shrink-0">
                        <div class="flex items-center space-x-4">
                            <button id="mensa-prev" class="p-2 rounded-full hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors">
                                <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M12.707 5.293a1 1 0 010 1.414L9.414 10l3.293 3.293a1 1 0 01-1.414 1.414l-4-4a1 1 0 010-1.414l4-4a1 1 0 011.414 0z" clipRule="evenodd" /></svg>
                            </button>
                            <div class="text-center">
                                <h2 class="text-lg font-bold text-slate-800 dark:text-slate-100">Mensa Berliner Tor</h2>
                                <p id="mensa-date-display" class="text-sm text-slate-500 dark:text-slate-400 font-medium"></p>
                            </div>
                            <button id="mensa-next" class="p-2 rounded-full hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors">
                                <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z" clipRule="evenodd" /></svg>
                            </button>
                        </div>
                        <button id="mensa-close" class="p-2 rounded-full text-slate-500 hover:bg-slate-200 dark:text-slate-400 dark:hover:bg-slate-700 transition-colors">
                             <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
                        </button>
                    </header>
                    <main id="mensa-content" class="p-4 sm:p-6 overflow-y-auto flex-grow">
                        <div class="flex justify-center items-center h-full"><div class="w-8 h-8 border-4 border-indigo-500 border-t-transparent rounded-full animate-spin"></div></div>
                    </main>
                    <footer id="mensa-footer" class="p-4 sm:p-5 border-t border-slate-200 dark:border-slate-700 flex-shrink-0 max-h-40 overflow-y-auto hidden">
                        <!-- Legend will be injected here -->
                    </footer>
                </div>
            </div>
        `;
        
        let currentDate = new Date();
        
        const fetchMenu = async () => {
            const content = document.getElementById('mensa-content');
            const footer = document.getElementById('mensa-footer');
            const dateDisplay = document.getElementById('mensa-date-display');
            
            dateDisplay.innerText = currentDate.toLocaleDateString('de-DE', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
            content.innerHTML = '<div class="flex justify-center items-center h-full"><div class="w-8 h-8 border-4 border-indigo-500 border-t-transparent rounded-full animate-spin"></div></div>';
            footer.classList.add('hidden');

            const year = currentDate.getFullYear();
            const month = String(currentDate.getMonth() + 1).padStart(2, '0');
            const day = String(currentDate.getDate()).padStart(2, '0');
            const url = `https://raw.githubusercontent.com/HAWHHCalendarBot/mensa-data/main/Mensa%20Berliner%20Tor/${year}/${month}/${day}.json`;
            
            try {
                const response = await fetch(url);
                if (!response.ok) throw new Error("Keine Daten");
                const data = await response.json();
                this.renderMensaContent(content, footer, data);
            } catch (e) {
                content.innerHTML = '<div class="text-center text-red-500 dark:text-red-400 p-8">Heute ist leider kein Menü verfügbar oder die Mensa ist geschlossen.</div>';
            }
        };

        document.getElementById('mensa-backdrop').onclick = () => container.innerHTML = '';
        document.getElementById('mensa-close').onclick = () => container.innerHTML = '';
        document.getElementById('mensa-prev').onclick = () => { currentDate.setDate(currentDate.getDate() - 1); fetchMenu(); };
        document.getElementById('mensa-next').onclick = () => { currentDate.setDate(currentDate.getDate() + 1); fetchMenu(); };

        fetchMenu();
    },

    renderMensaContent: function(contentEl, footerEl, menuData) {
        // Group by Category
        const grouped = menuData.reduce((acc, item) => {
            const cat = item.Category || 'Sonstiges';
            if (!acc[cat]) acc[cat] = [];
            acc[cat].push(item);
            return acc;
        }, {});
        
        const allAdditives = {};

        // Sort categories
        const order = ['CampusKlassiker', 'Highlight', 'CampusVital', 'CampusWorld', 'Pottkieker', 'Pastabar', 'Gemüsebar'];
        const sortedKeys = Object.keys(grouped).sort((a, b) => {
             const ia = order.indexOf(a), ib = order.indexOf(b);
             if (ia === -1 && ib === -1) return a.localeCompare(b);
             if (ia === -1) return 1;
             if (ib === -1) return -1;
             return ia - ib;
        });

        let html = '';
        sortedKeys.forEach(cat => {
            html += `<div class="mb-6"><h3 class="text-xl font-bold text-indigo-600 dark:text-indigo-400 mb-3 pb-2 border-b-2 border-indigo-200 dark:border-indigo-500/30">${cat}</h3><div class="space-y-4">`;
            grouped[cat].forEach(item => {
                // Collect additives
                if (item.Additives) {
                    Object.entries(item.Additives).forEach(([k, v]) => {
                         const cleanKey = k.replace(/<\/?strong>/g, '');
                         allAdditives[cleanKey] = v;
                    });
                }

                html += `
                <div class="bg-slate-50 dark:bg-slate-700/40 p-4 rounded-lg">
                    <h4 class="font-semibold text-slate-800 dark:text-slate-100">${item.Name.replace(/\s*\([^)]+\)/g, '').trim()}</h4>
                    <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between mt-2 text-sm gap-2">
                        <div class="flex flex-wrap gap-2">
                            ${item.Vegan ? '<span class="px-2 py-0.5 bg-green-100 text-green-800 dark:bg-green-500/20 dark:text-green-300 rounded-full text-xs font-semibold">Vegan</span>' : ''}
                            ${item.Vegetarian && !item.Vegan ? '<span class="px-2 py-0.5 bg-emerald-100 text-emerald-800 dark:bg-emerald-500/20 dark:text-emerald-300 rounded-full text-xs font-semibold">Vegetarisch</span>' : ''}
                            ${item.LactoseFree ? '<span class="px-2 py-0.5 bg-blue-100 text-blue-800 dark:bg-blue-500/20 dark:text-blue-300 rounded-full text-xs font-semibold">Laktosefrei</span>' : ''}
                        </div>
                        <div class="text-left sm:text-right font-mono text-xs text-slate-600 dark:text-slate-300 flex-shrink-0">
                            <span>Stud: <strong>${item.PriceStudent.toFixed(2)}€</strong></span> / <span>Bed: ${item.PriceAttendant.toFixed(2)}€</span> / <span>Gast: ${item.PriceGuest.toFixed(2)}€</span>
                        </div>
                    </div>
                    ${item.Additives && Object.keys(item.Additives).length > 0 ? 
                        `<div class="mt-2 text-xs text-slate-500 dark:text-slate-400">
                            Zusatzstoffe: ${Object.keys(item.Additives).map(k => k.replace(/<\/?strong>/g, '')).join(', ')}
                        </div>` : ''}
                </div>`;
            });
            html += `</div></div>`;
        });
        contentEl.innerHTML = html;

        // Render Footer
        if (Object.keys(allAdditives).length > 0) {
            let footerHtml = '<h4 class="text-sm font-bold mb-2">Zusatzstoffe-Legende</h4><div class="text-xs text-slate-600 dark:text-slate-400 grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-x-4 gap-y-1">';
            Object.entries(allAdditives).forEach(([code, desc]) => {
                footerHtml += `<div><strong>${code}:</strong> ${desc}</div>`;
            });
            footerHtml += '</div>';
            footerEl.innerHTML = footerHtml;
            footerEl.classList.remove('hidden');
        }
    },

    setupCalendarModal: function() {
        const toggleBtn = document.getElementById('calendar-toggle-btn');
        if (toggleBtn) {
            toggleBtn.addEventListener('click', () => {
                const container = document.getElementById('calendar-modal-container');
                if (container.innerHTML) {
                    container.innerHTML = '';
                } else {
                    this.renderCalendarModal(container);
                }
            });
        }
    },

    renderCalendarModal: function(container) {
        // Count base events (not expanded)
        const baseEvents = Krankomat.State.data.calendarEvents ? Krankomat.State.data.calendarEvents.length : 0;
        const countText = baseEvents > 0 ? `(${baseEvents} Einträge geladen)` : '';

        container.innerHTML = `
            <div class="fixed inset-0 bg-black bg-opacity-60 z-50 flex justify-center items-center p-4 transition-opacity duration-300" id="calendar-backdrop">
                <div class="bg-white dark:bg-slate-800 rounded-2xl shadow-2xl w-full max-w-lg flex flex-col animate-scale-in" onclick="event.stopPropagation()">
                    <header class="p-4 sm:p-5 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between flex-shrink-0">
                        <div class="flex items-center space-x-4">
                            <button id="calendar-prev" class="p-2 rounded-full hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors">
                                <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M12.707 5.293a1 1 0 010 1.414L9.414 10l3.293 3.293a1 1 0 01-1.414 1.414l-4-4a1 1 0 010-1.414l4-4a1 1 0 011.414 0z" clipRule="evenodd" /></svg>
                            </button>
                            <div class="text-center">
                                <h2 class="text-lg font-bold text-slate-800 dark:text-slate-100">Stundenplan</h2>
                                <p class="text-xs text-indigo-600 dark:text-indigo-400 mb-0.5">${countText}</p>
                                <p id="calendar-date-display" class="text-sm text-slate-500 dark:text-slate-400 font-medium"></p>
                            </div>
                            <button id="calendar-next" class="p-2 rounded-full hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors">
                                <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z" clipRule="evenodd" /></svg>
                            </button>
                        </div>
                        <button id="calendar-close" class="p-2 rounded-full text-slate-500 hover:bg-slate-200 dark:text-slate-400 dark:hover:bg-slate-700 transition-colors">
                             <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
                        </button>
                    </header>
                    <main id="calendar-content" class="p-4 sm:p-6 overflow-y-auto max-h-[60vh]">
                        <!-- Calendar events injected here -->
                    </main>
                </div>
            </div>
        `;

        let currentDate = new Date();
        const events = Krankomat.State.data.calendarEvents || [];

        const renderEvents = () => {
            const dateDisplay = document.getElementById('calendar-date-display');
            const content = document.getElementById('calendar-content');
            
            dateDisplay.innerText = currentDate.toLocaleDateString('de-DE', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
            
            // Format current date to string for filtering
            const day = String(currentDate.getDate()).padStart(2, '0');
            const month = String(currentDate.getMonth() + 1).padStart(2, '0');
            const year = currentDate.getFullYear();
            const dateStr = `${day}.${month}.${year}`;

            // Smart filter that handles RRULEs on the fly
            const daysEvents = events.filter(evt => Krankomat.Utils.isEventOnDate(evt, dateStr));

            if (daysEvents.length === 0) {
                 if (events.length === 0) {
                     content.innerHTML = `
                        <div class="text-center py-8">
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-12 w-12 mx-auto text-slate-300 dark:text-slate-600 mb-3" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" />
                            </svg>
                            <p class="text-slate-500 dark:text-slate-400">Keine Kalenderdatei geladen.</p>
                            <p class="text-xs text-slate-400 mt-1">Bitte laden Sie eine .ics Datei in den Einstellungen hoch.</p>
                        </div>`;
                 } else {
                     content.innerHTML = `
                        <div class="text-center py-8">
                            <p class="text-slate-500 dark:text-slate-400">Keine Termine an diesem Tag.</p>
                        </div>`;
                 }
            } else {
                let html = '<div class="space-y-3">';
                // Sorting logic for time
                daysEvents.sort((a,b) => (a.start || '').localeCompare(b.start || '')).forEach(evt => {
                    // Extract time if present in start string (YYYYMMDDTHHMMSS)
                    let timeStr = '';
                    if (evt.start && evt.start.includes('T')) {
                        const timePart = evt.start.split('T')[1];
                        if (timePart && timePart.length >= 4) {
                            timeStr = `${timePart.substring(0,2)}:${timePart.substring(2,4)}`;
                        }
                    }
                    
                    // Style logic for Location - Bounding Box changes
                    // Default
                    let containerClasses = "bg-indigo-50 dark:bg-slate-700/50 border-indigo-500";
                    let locIconClass = "text-slate-500 dark:text-slate-400";
                    let timeColorClass = "text-indigo-600 dark:text-indigo-400";

                    if (evt.location) {
                        const locLower = evt.location.toLowerCase();
                        if (locLower.includes('asynchron')) {
                            // Grey
                            containerClasses = "bg-slate-100 dark:bg-slate-800 border-slate-400 dark:border-slate-500";
                            locIconClass = "text-slate-500 italic";
                            timeColorClass = "text-slate-600 dark:text-slate-400";
                        } else if (locLower.includes('synchron')) {
                            // Green
                            containerClasses = "bg-green-50 dark:bg-green-900/20 border-green-500";
                            locIconClass = "text-green-600 dark:text-green-400";
                            timeColorClass = "text-green-700 dark:text-green-400";
                        }
                    }
                    
                    html += `
                        <div class="${containerClasses} p-4 rounded-lg border-l-4 transition-colors">
                            <h4 class="font-bold text-slate-800 dark:text-white break-words">${evt.summary}</h4>
                            ${timeStr ? `<div class="flex items-center mt-2 text-sm ${timeColorClass}">
                                <svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4 mr-1.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" />
                                </svg>
                                ${timeStr} Uhr
                            </div>` : ''}
                            ${evt.location ? `<div class="flex items-center mt-1 text-xs ${locIconClass}">
                                <svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5 mr-1.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z" />
                                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 11a3 3 0 11-6 0 3 3 0 016 0z" />
                                </svg>
                                ${evt.location}
                            </div>` : ''}
                        </div>
                    `;
                });
                html += '</div>';
                content.innerHTML = html;
            }
        };

        document.getElementById('calendar-backdrop').onclick = () => container.innerHTML = '';
        document.getElementById('calendar-close').onclick = () => container.innerHTML = '';
        
        document.getElementById('calendar-prev').onclick = () => { 
            currentDate.setDate(currentDate.getDate() - 1); 
            renderEvents(); 
        };
        
        document.getElementById('calendar-next').onclick = () => { 
            currentDate.setDate(currentDate.getDate() + 1); 
            renderEvents(); 
        };

        renderEvents();
    }
};

// Start the app
document.addEventListener('DOMContentLoaded', () => {
    Krankomat.App.init();
});