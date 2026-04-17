const KrankomatHVV = {
    init: function() {
        this.bindEvents();
    },

    bindEvents: function() {
        const toggleBtn = document.getElementById('hvv-toggle-btn');
        if (toggleBtn) {
            toggleBtn.addEventListener('click', () => this.openModal());
        }
        
        // Close modal when clicking outside
        document.addEventListener('click', (e) => {
            const modal = document.getElementById('hvv-modal');
            if (modal && e.target === modal) {
                this.closeModal();
            }
        });
    },

    openModal: function() {
        this.renderModal();
        this.fetchData();
    },

    closeModal: function() {
        const container = document.getElementById('hvv-modal-container');
        if (container) {
            container.innerHTML = '';
        }
    },

    renderModal: function() {
        const container = document.getElementById('hvv-modal-container');
        if (!container) return;

        const hvvUrl = Krankomat.State.data.config?.hvvUrl || '';

        container.innerHTML = `
            <div id="hvv-modal" class="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4 animate-fade-in">
                <div class="bg-white dark:bg-slate-800 rounded-2xl shadow-xl w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden animate-slide-up">
                    <div class="p-4 lg:p-6 border-b border-slate-200 dark:border-slate-700 flex justify-between items-center bg-slate-50 dark:bg-slate-800/50">
                        <div class="flex items-center space-x-3">
                            <div class="w-10 h-10 flex items-center justify-center text-[#00962c]">
                                <svg xmlns="http://www.w3.org/2000/svg" viewBox="100.7025 87.8992 323.9685 323.9681" class="w-full h-full">
                                  <defs>
                                    <clipPath clipPathUnits="userSpaceOnUse" id="clipPath188_modal">
                                      <path d="m 482.129,10.747 98.903,0 0,21.033 -98.903,0 0,-21.033 z"/>
                                    </clipPath>
                                  </defs>
                                  <g transform="matrix(12.322287559509277, 0, 0, 12.322287559509277, -2686.6875, -5596.1552734375)">
                                    <g transform="matrix(1, 0, 0, 1, 33.668869, -27.289463)">
                                      <g transform="matrix(1.25,0,0,-1.25,-410.12253,528.29722)">
                                        <g clip-path="url(#clipPath188_modal)">
                                          <g transform="translate(482.1287,21.2643)">
                                            <path d="m 0,0 c 0,5.808 4.708,10.516 10.516,10.516 5.808,0 10.517,-4.708 10.517,-10.516 0,-5.808 -4.709,-10.517 -10.517,-10.517 C 4.708,-10.517 0,-5.808 0,0" style="fill:currentColor;fill-rule:evenodd;stroke:none"/>
                                          </g>
                                          <g transform="translate(498.4521,27.3232)">
                                            <path d="m 0,0 0,-2.729 c -1.449,1.563 -3.633,2.936 -6.193,2.936 -1.45,0 -2.278,-0.772 -2.278,-1.675 0,-3.802 9.468,-1.129 9.468,-7.378 0,-2.523 -2.446,-5.459 -6.569,-5.459 -2.297,0 -4.95,1.035 -6.4,2.334 l 0,3.276 c 1.374,-1.977 3.878,-3.501 6.4,-3.501 1.544,0 2.786,0.884 2.786,1.826 0,3.556 -9.317,1.354 -9.317,7.49 0,3.219 3.143,4.951 5.966,4.951 C -3.784,2.071 -1.675,1.299 0,0" style="fill:#ffffff;fill-rule:evenodd;stroke:none"/>
                                          </g>
                                        </g>
                                      </g>
                                    </g>
                                  </g>
                                </svg>
                            </div>
                            <h2 class="text-xl font-bold text-slate-800 dark:text-white">HVV Abfahrten</h2>
                        </div>
                        <button id="hvv-close-btn" class="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 transition-colors p-2 rounded-full hover:bg-slate-200 dark:hover:bg-slate-700">
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12" />
                            </svg>
                        </button>
                    </div>
                    
                    <div class="p-4 lg:p-6 overflow-y-auto flex-grow" id="hvv-content">
                        ${!hvvUrl ? `
                            <div class="text-center py-8">
                                <p class="text-slate-500 dark:text-slate-400 mb-4">Keine HVV URL in der Konfiguration gefunden.</p>
                                <p class="text-sm text-slate-400 dark:text-slate-500">Bitte füge eine HVV Abfahrtsmonitor-URL in der Konfiguration (hvvUrl) hinzu.</p>
                            </div>
                        ` : `
                            <div class="flex justify-center items-center py-12">
                                <div class="animate-spin rounded-full h-12 w-12 border-b-2 border-indigo-600"></div>
                            </div>
                        `}
                    </div>
                </div>
            </div>
        `;

        document.getElementById('hvv-close-btn')?.addEventListener('click', () => this.closeModal());
    },

    fetchData: async function() {
        const hvvUrl = Krankomat.State.data.config?.hvvUrl;
        if (!hvvUrl) return;

        const contentEl = document.getElementById('hvv-content');
        if (!contentEl) return;

        try {
            // Extract show ID
            const urlObj = new URL(hvvUrl);
            const showId = urlObj.searchParams.get('show');
            
            if (!showId) {
                throw new Error('Keine "show" ID in der URL gefunden.');
            }

            let configData = null;
            const cacheKey = `krankomat_hvv_config_${showId}`;
            const cached = localStorage.getItem(cacheKey);
            
            if (cached) {
                try {
                    configData = JSON.parse(cached);
                } catch (e) {
                    console.error('Cache parse error', e);
                }
            }

            if (!configData) {
                // Fetch config via proxy
                const proxyUrl = `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(`https://www.hvv.de/linking-service/abfahrten/show/${showId}`)}`;
                const configRes = await fetch(proxyUrl);
                
                if (!configRes.ok) {
                    throw new Error('Fehler beim Laden der HVV Konfiguration über den Proxy. Bitte versuche es später erneut.');
                }
                
                configData = await configRes.json();
                
                if (!configData.stationList || configData.stationList.length === 0) {
                    throw new Error('Keine Stationen in der Konfiguration gefunden.');
                }
                
                // Save to cache
                localStorage.setItem(cacheKey, JSON.stringify(configData));
            }

            // Hardcode Berliner Tor station to bypass HVV config url if needed
            configData.stationList = [{
                name: 'Berliner Tor',
                city: 'Hamburg',
                combinedName: 'Berliner Tor',
                id: 'Master:10952',
                globalId: 'de:02000:10952',
                type: 'STATION'
            }];
            configData.filterList = [];
            
            const stationName = configData.stationList[0].name;
            
            // Fetch departures from HVV API directly (supports CORS for file://)
            const now = new Date();
            const dateStr = `${String(now.getDate()).padStart(2, '0')}.${String(now.getMonth() + 1).padStart(2, '0')}.${now.getFullYear()}`;
            const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

            const payload = {
                version: 51,
                station: configData.stationList[0],
                time: { date: dateStr, time: timeStr },
                maxList: 40,
                allStationsInChangingNode: true,
                returnFilters: true,
                filter: configData.filterList || [],
                maxTimeOffset: 200,
                useRealtime: true
            };

            const depRes = await fetch('https://www.hvv.de/geofox/departureList', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                body: JSON.stringify(payload)
            });

            if (!depRes.ok) {
                throw new Error('Fehler beim Laden der Abfahrten (HVV API).');
            }

            const depData = await depRes.json();
            
            let departures = depData.departures || [];
            
            // Remove Client-Side logic, fetch all departures for Berliner Tor
            // Limit to 20 departures to avoid overwhelming the popup
            departures = departures.slice(0, 20);
            
            this.renderDepartures(departures, configData.stationList[0].name);

        } catch (error) {
            console.error('HVV Error:', error);
            contentEl.innerHTML = `
                <div class="bg-red-50 dark:bg-red-900/20 text-red-600 dark:text-red-400 p-4 rounded-lg text-center">
                    <p class="font-medium">Fehler beim Laden der Daten</p>
                    <p class="text-sm mt-1">${error.message}</p>
                </div>
            `;
        }
    },

    renderDepartures: function(departures, stationName) {
        const contentEl = document.getElementById('hvv-content');
        if (!contentEl) return;

        const validDepartures = departures; // HVV API returns valid departures

        if (validDepartures.length === 0) {
            contentEl.innerHTML = `
                <div class="text-center py-8">
                    <p class="text-slate-500 dark:text-slate-400">Keine Abfahrten in nächster Zeit gefunden.</p>
                </div>
            `;
            return;
        }

        let html = `
            <h3 class="font-semibold text-slate-700 dark:text-slate-300 mb-4 text-lg">Abfahrten ab ${stationName}</h3>
            <div class="space-y-3">
        `;

        validDepartures.forEach(dep => {
            const lineName = dep.line && dep.line.name ? dep.line.name : 'Unbekannt';
            const direction = dep.line && dep.line.direction ? dep.line.direction : 'Unbekannt';
            const isCancelled = dep.cancelled === true;
            
            const timeOffset = Math.max(0, dep.timeOffset || 0);
            const delay = dep.delay ? Math.floor(dep.delay / 60) : 0;
            
            let timeText = timeOffset === 0 ? 'Jetzt' : `in ${timeOffset} Min`;
            if (isCancelled) {
                timeText = 'Fällt aus';
            }
            
            let delayHtml = '';
            
            if (!isCancelled) {
                if (delay > 0) {
                    delayHtml = `<span class="text-red-500 text-xs font-medium ml-2">+${delay}</span>`;
                } else if (delay < 0) {
                    delayHtml = `<span class="text-green-500 text-xs font-medium ml-2">${delay}</span>`;
                }
            }

            // Determine color based on line type
            let bgColor = 'bg-slate-200 dark:bg-slate-700';
            let textColor = 'text-slate-800 dark:text-slate-200';
            let iconHtml = lineName;
            
            if (lineName === 'S1' || lineName === 'S 1') {
                iconHtml = `
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 661.6699 260.165" class="h-6 w-auto">
                        <path d="M 137.5635,0 L 524.1065,0 C 600.0899,0 661.67,58.2461 661.67,130.0889 C 661.67,201.9307 600.0899,260.1651 524.1065,260.1651 L 137.5635,260.1651 C 61.5801,260.165 0,201.9307 0,130.0889 C 0,58.2461 61.5801,0 137.5635,0 z " style="fill:#00962c;fill-rule:evenodd"/>
                        <path d="M 192.0244,200.5312 C 200.9951,205.2763 219.4805,210.2724 238.7002,210.2724 C 285.3789,210.2724 308.0889,186.0556 308.0889,156.5947 C 308.0889,134.3847 295.3711,117.9023 266.3965,106.4306 C 245.4414,98.1826 237.4444,94.4452 237.4444,85.7001 C 237.4444,78.2138 244.9434,72.7216 258.6583,72.7216 C 274.1554,72.7216 285.8761,77.205 292.6183,80.4677 L 301.3507,45.2646 C 291.3585,40.7666 277.6437,37.0302 258.921,37.0302 C 217.2276,37.0302 193.5206,59.9892 193.5206,89.4501 C 193.5206,112.4101 209.9884,128.3925 238.2032,139.3778 C 256.4288,146.3544 263.9024,151.1034 263.9024,159.8456 C 263.9024,168.579 256.4288,174.3222 240.9561,174.3222 C 225.46,174.3222 209.7286,169.3271 199.9991,164.3447 L 192.0244,200.5312 z " style="fill:#ffffff;fill-rule:evenodd"/>
                        <polygon points="406.877,208.1123 450.208,208.1123 450.208,40.4834 413.0713,40.4834 366.8916,61.3779 374.1064,95.1582 406.3564,80.1963 406.877,80.1963 406.877,208.1123 " style="fill:#ffffff;fill-rule:evenodd"/>
                    </svg>
                `;
                bgColor = 'bg-transparent';
                textColor = '';
            } else if (lineName.startsWith('S')) {
                bgColor = 'bg-[#00962c]';
                textColor = 'text-white';
            } else if (lineName.startsWith('U')) {
                bgColor = 'bg-[#003087]';
                textColor = 'text-white';
            } else if (dep.line && dep.line.type && dep.line.type.simpleType === 'BUS') {
                bgColor = 'bg-[#E3001B]';
                textColor = 'text-white';
            }

            const platform = dep.realtimePlatform || dep.platform || '';

            html += `
                <div class="flex items-center justify-between p-3 bg-slate-50 dark:bg-slate-800/50 rounded-lg border border-slate-100 dark:border-slate-700/50 hover:border-indigo-200 dark:hover:border-indigo-800 transition-colors ${isCancelled ? 'opacity-60' : ''}">
                    <div class="flex items-center space-x-4 overflow-hidden">
                        <div class="${bgColor} ${textColor} font-bold ${(lineName === 'S1' || lineName === 'S 1') ? '' : 'px-3 py-1'} rounded min-w-[3rem] text-center flex-shrink-0 flex items-center justify-center">
                            ${iconHtml}
                        </div>
                        <div class="truncate">
                            <p class="font-medium text-slate-800 dark:text-slate-200 truncate ${isCancelled ? 'line-through' : ''}" title="${direction}">${direction}</p>
                            <p class="text-xs text-slate-500 dark:text-slate-400">${platform ? 'Gleis ' + platform : ''}</p>
                        </div>
                    </div>
                    <div class="text-right flex-shrink-0 ml-4">
                        <p class="font-bold ${isCancelled ? 'text-red-500' : 'text-slate-800 dark:text-slate-200'}">${timeText}</p>
                        ${delayHtml}
                    </div>
                </div>
            `;
        });

        html += `</div>`;
        contentEl.innerHTML = html;
    }
};

// Expose globally
window.KrankomatHVV = KrankomatHVV;
