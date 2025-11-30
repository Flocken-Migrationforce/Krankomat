
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
        console.log('Krankomat initialized successfully.');
    },

    renderAll: function() {
        Krankomat.Builder.render();
        Krankomat.Preview.render();
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

    setupModals: function() {
        this.setupTermsModal();
        this.setupExpertModal();
        this.setupMensaModal();
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

    setupExpertModal: function() {
        const toggle = () => {
            const container = document.getElementById('expert-modal-container');
            if (container.innerHTML) {
                container.innerHTML = '';
            } else {
                this.renderExpertModal(container);
            }
        };

        const btn = document.getElementById('expert-mode-btn');
        if (btn) btn.addEventListener('click', toggle);
        
        window.addEventListener('keydown', (e) => {
            if (e.ctrlKey && e.shiftKey && (e.key === 'X' || e.key === 'x')) {
                e.preventDefault();
                toggle();
            }
        });
    },

    renderExpertModal: function(container) {
        container.innerHTML = `
          <div class="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-sm transition-all" id="expert-backdrop">
             <div class="bg-white dark:bg-slate-800 rounded-xl shadow-2xl max-w-lg w-full border border-indigo-500/30 animate-scale-in" onclick="event.stopPropagation()">
                <div class="p-5 border-b border-slate-200 dark:border-slate-700 flex justify-between items-center">
                   <h2 class="text-lg font-bold text-slate-800 dark:text-white flex items-center">
                     <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" class="w-5 h-5 mr-2 text-indigo-600 dark:text-indigo-400">
                       <path fillRule="evenodd" d="M12 2.25c-5.385 0-9.75 4.365-9.75 9.75s4.365 9.75 9.75 9.75 9.75-4.365 9.75-9.75S17.385 2.25 12 2.25zm-1.72 6.97a.75.75 0 10-1.06 1.06L10.94 12l-1.72 1.72a.75.75 0 101.06 1.06L12 13.06l1.72 1.72a.75.75 0 101.06-1.06L13.06 12l1.72-1.72a.75.75 0 10-1.06-1.06L12 10.94l-1.72-1.72z" clipRule="evenodd" />
                     </svg>
                     Expert Mode: Configuration
                   </h2>
                   <button id="expert-close-btn" class="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200">
                     <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" class="w-6 h-6">
                       <path fillRule="evenodd" d="M5.47 5.47a.75.75 0 011.06 0L12 10.94l5.47-5.47a.75.75 0 111.06 1.06L13.06 12l5.47 5.47a.75.75 0 11-1.06 1.06L12 13.06l-5.47 5.47a.75.75 0 01-1.06-1.06L10.94 12 5.47 6.53a.75.75 0 010-1.06z" clipRule="evenodd" />
                     </svg>
                   </button>
                </div>
                <div class="p-6 space-y-6">
                  <div class="p-4 bg-slate-50 dark:bg-slate-700/50 rounded-lg border border-slate-200 dark:border-slate-600">
                    <h3 class="text-sm font-bold text-slate-700 dark:text-slate-200 mb-2">Backup & Template</h3>
                    <p class="text-xs text-slate-500 dark:text-slate-400 mb-3">
                      Download your current configuration as a JSON file. You can use this as a template for making changes or as a backup.
                    </p>
                    <button id="expert-export-btn" class="w-full py-2 bg-white dark:bg-slate-600 border border-slate-300 dark:border-slate-500 text-slate-700 dark:text-slate-200 rounded-md text-sm font-medium hover:bg-slate-50 dark:hover:bg-slate-500 transition-colors flex items-center justify-center">
                      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor" class="w-4 h-4 mr-2">
                        <path fillRule="evenodd" d="M10 3a.75.75 0 01.75.75v10.638l3.96-4.158a.75.75 0 111.08 1.04l-5.25 5.5a.75.75 0 01-1.08 0l-5.25-5.5a.75.75 0 111.08-1.04l3.96 4.158V3.75A.75.75 0 0110 3z" clipRule="evenodd" />
                      </svg>
                      Download Current Config (JSON)
                    </button>
                  </div>

                  <div class="p-4 bg-indigo-50 dark:bg-indigo-900/20 rounded-lg border border-indigo-100 dark:border-indigo-800/30">
                    <h3 class="text-sm font-bold text-indigo-800 dark:text-indigo-300 mb-2">Import Configuration</h3>
                    <p class="text-xs text-indigo-600 dark:text-indigo-400 mb-3">
                      Upload a JSON config file to overwrite the current application state. <strong>Warning: This will reload the page.</strong>
                    </p>
                    <label class="w-full flex flex-col items-center px-4 py-4 bg-white dark:bg-slate-700 text-indigo-600 dark:text-indigo-300 rounded-lg shadow-sm tracking-wide uppercase border border-indigo-200 dark:border-slate-600 cursor-pointer hover:bg-indigo-50 dark:hover:bg-slate-600 transition-colors">
                        <svg class="w-6 h-6" fill="currentColor" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path d="M16.88 9.1A4 4 0 0 1 16 17H5a5 5 0 0 1-1-9.9V7a3 3 0 0 1 4.52-2.59A4.98 4.98 0 0 1 17 8c0 .38-.04.74-.12 1.1zM11 11h3l-4-4-4 4h3v3h2v-3z" /></svg>
                        <span class="mt-2 text-xs font-bold">Select JSON File</span>
                        <input type='file' id="expert-file-input" class="hidden" accept=".json" />
                    </label>
                  </div>
                </div>
             </div>
          </div>
        `;

        document.getElementById('expert-close-btn').onclick = () => container.innerHTML = '';
        document.getElementById('expert-backdrop').onclick = () => container.innerHTML = '';

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
                    window.location.reload();
                } catch (err) {
                    alert("Error parsing JSON");
                }
            };
            reader.readAsText(file);
        };
    },

    setupMensaModal: function() {
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
            const month = currentDate.getMonth() + 1;
            const day = currentDate.getDate();
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
    }
};

// Start the app
document.addEventListener('DOMContentLoaded', () => {
    Krankomat.App.init();
});