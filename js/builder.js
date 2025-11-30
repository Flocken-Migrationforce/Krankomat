
window.Krankomat = window.Krankomat || {};

Krankomat.Builder = {
    init: function() {
        this.bindEvents();
        this.render();
        
        // Bind recipients events once here instead of inside render
        const recipientsContainer = document.getElementById('recipients-list-container');
        if (recipientsContainer) {
             this.bindRecipientEvents(recipientsContainer);
        }
    },

    bindEvents: function() {
        // User Data Inputs
        ['firstName', 'lastName', 'studentId'].forEach(field => {
            const el = document.getElementById(field);
            if (el) {
                el.addEventListener('input', (e) => {
                    Krankomat.State.updateNested('userData', field, e.target.value);
                });
            }
        });

        // Dates
        const startEl = document.getElementById('startDate');
        if (startEl) {
            startEl.addEventListener('input', (e) => {
                Krankomat.State.set('sicknessStartDate', e.target.value);
                this.updateRecipientsBasedOnDate(e.target.value);
            });
        }
        
        const endEl = document.getElementById('endDate');
        if (endEl) {
            endEl.addEventListener('input', (e) => {
                Krankomat.State.set('sicknessEndDate', e.target.value);
            });
        }

        // Buttons for type
        document.getElementById('btn-type-krank').addEventListener('click', () => Krankomat.State.set('sicknessEndDate', ''));
        document.getElementById('btn-type-gesund').addEventListener('click', () => {
             const endDate = Krankomat.State.get('sicknessEndDate');
             if (!endDate) Krankomat.State.set('sicknessEndDate', Krankomat.Utils.todayFormatted());
        });
        document.getElementById('btn-today-healthy').addEventListener('click', () => Krankomat.State.set('sicknessEndDate', Krankomat.Utils.todayFormatted()));

        // Comments
        document.getElementById('comments').addEventListener('input', (e) => {
            Krankomat.State.updateNested('details', 'comments', e.target.value);
        });

        // Manual Save
        const saveBtn = document.getElementById('manual-save-btn');
        if (saveBtn) {
            saveBtn.addEventListener('click', () => {
                Krankomat.State.save();
                const originalContent = saveBtn.innerHTML;
                saveBtn.innerHTML = Krankomat.Utils.Icons.check + ' Gespeichert!';
                saveBtn.classList.add('bg-green-500', 'text-white');
                saveBtn.classList.remove('bg-indigo-100', 'text-indigo-700');
                setTimeout(() => {
                    saveBtn.innerHTML = originalContent;
                    saveBtn.classList.remove('bg-green-500', 'text-white');
                    saveBtn.classList.add('bg-indigo-100', 'text-indigo-700');
                }, 2000);
            });
        }
    },
    
    bindRecipientEvents: function(container) {
        container.addEventListener('change', (e) => {
            if (e.target.dataset.recipientId) {
                const id = parseInt(e.target.dataset.recipientId);
                const currentRecipients = Krankomat.State.get('recipients');
                const newRecipients = currentRecipients.map(r => r.id === id ? { ...r, isSelected: e.target.checked } : r);
                Krankomat.State.set('recipients', newRecipients);
            }
        });

        container.addEventListener('input', (e) => {
             if (e.target.dataset.recipientEmailId) {
                const id = parseInt(e.target.dataset.recipientEmailId);
                const currentRecipients = Krankomat.State.get('recipients');
                // We update state which triggers notify -> render. 
                // The updated renderRecipientsList handles focus preservation.
                const newRecipients = currentRecipients.map(r => r.id === id ? { ...r, email: e.target.value } : r);
                Krankomat.State.set('recipients', newRecipients);
             }
        });
    },

    updateRecipientsBasedOnDate: function(dateStr) {
        const day = Krankomat.Utils.getDayOfWeek(dateStr);
        if (!day) return;
        
        const data = Krankomat.State.data;
        const timetable = Krankomat.State.defaults.timetable;
        const todaysModules = timetable.filter(e => e.day.toLowerCase() === day.toLowerCase()).map(e => e.module.toLowerCase());
        
        const newRecipients = data.recipients.map(r => ({
            ...r,
            isSelected: r.module.includes('ZPD') || todaysModules.some(mod => r.module.toLowerCase().includes(mod))
        }));
        
        Krankomat.State.set('recipients', newRecipients);
    },

    render: function() {
        const data = Krankomat.State.data;
        if (!data.userData) return; // Safety check

        // Input values
        const firstNameInput = document.getElementById('firstName');
        if (firstNameInput && document.activeElement !== firstNameInput) firstNameInput.value = data.userData.firstName || '';
        
        const lastNameInput = document.getElementById('lastName');
        if (lastNameInput && document.activeElement !== lastNameInput) lastNameInput.value = data.userData.lastName || '';
        
        const studentIdInput = document.getElementById('studentId');
        if (studentIdInput && document.activeElement !== studentIdInput) studentIdInput.value = data.userData.studentId || '';
        
        const startDateInput = document.getElementById('startDate');
        if (startDateInput && document.activeElement !== startDateInput) startDateInput.value = data.sicknessStartDate || '';
        
        const endDateInput = document.getElementById('endDate');
        if (endDateInput && document.activeElement !== endDateInput) endDateInput.value = data.sicknessEndDate || '';
        
        const commentsInput = document.getElementById('comments');
        if (commentsInput && document.activeElement !== commentsInput) commentsInput.value = data.details.comments || '';

        // Type Buttons styling
        const isGesund = !!data.sicknessEndDate;
        const krankBtn = document.getElementById('btn-type-krank');
        const gesundBtn = document.getElementById('btn-type-gesund');
        
        if (!isGesund) {
            krankBtn.className = "w-1/2 py-2 text-sm font-semibold rounded-full transition-all duration-300 bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 shadow";
            gesundBtn.className = "w-1/2 py-2 text-sm font-semibold rounded-full transition-all duration-300 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600";
        } else {
            gesundBtn.className = "w-1/2 py-2 text-sm font-semibold rounded-full transition-all duration-300 bg-white dark:bg-slate-800 text-green-600 dark:text-green-400 shadow";
            krankBtn.className = "w-1/2 py-2 text-sm font-semibold rounded-full transition-all duration-300 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600";
        }

        this.renderAbsenceReasons(data.absenceReasons);
        this.renderRecipientsList(data.recipients);
        this.renderCommentPresets(data.details.comments);
    },

    renderAbsenceReasons: function(reasons) {
        const container = document.getElementById('absence-reasons-container');
        if (!container) return;
        
        const definitions = [
            { id: 'lecture', label: 'Vorlesungszeit' },
            { id: 'internship', label: 'Berufspraxis' },
            { id: 'exam', label: 'Prüfungsleistung / Klausur', warning: true },
            { id: 'partialDay', label: 'Rest des Tages (war schon da)' }
        ];

        const html = definitions.map(def => `
            <div>
                <label class="flex items-center space-x-3 cursor-pointer group">
                    <div class="relative flex items-center">
                        <input type="checkbox" data-reason="${def.id}" ${reasons[def.id] ? 'checked' : ''} class="peer relative h-5 w-5 cursor-pointer appearance-none rounded-md border border-slate-400 dark:border-slate-500 transition-all checked:border-indigo-600 checked:bg-indigo-600 dark:checked:border-indigo-500 dark:checked:bg-indigo-500" />
                        <div class="pointer-events-none absolute top-2/4 left-2/4 -translate-y-2/4 -translate-x-2/4 text-white opacity-0 transition-opacity peer-checked:opacity-100">
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" viewBox="0 0 20 20" fill="currentColor" stroke="currentColor" strokeWidth="1">
                                <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd"></path>
                            </svg>
                        </div>
                    </div>
                    <span class="text-sm text-slate-700 dark:text-slate-300 group-hover:text-indigo-600 dark:group-hover:text-indigo-400 transition-colors">
                        ${def.label}
                    </span>
                </label>
                ${def.warning && reasons[def.id] ? `
                <div class="mt-3 ml-8 p-3 bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-500/30 rounded-lg text-sm text-blue-800 dark:text-blue-200 flex items-start space-x-3 transition-all duration-300 animate-fade-in">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" fill="currentColor" class="h-5 w-5 text-blue-500 dark:text-blue-400 mt-1 flex-shrink-0">
                        <path d="M256 512A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM216 336h24V272H216c-13.3 0-24-10.7-24-24s10.7-24 24-24h48c13.3 0 24 10.7 24 24v88h8c13.3 0 24 10.7 24 24s-10.7 24-24 24H216c-13.3 0-24-10.7-24-24s10.7-24 24-24zm40-208a32 32 0 1 1 0 64 32 32 0 1 1 0-64z"/>
                    </svg>
                    <p><strong>Wichtiger Hinweis:</strong> Diese E-Mail ersetzt nicht die offizielle Prüfungsunfähigkeitsbescheinigung, die beim Prüfungsamt eingereicht werden muss.</p>
                </div>` : ''}
            </div>
        `).join('');
        
        // Only update innerHTML if it has changed to prevent unnecessary DOM thrashing
        if (container.innerHTML !== html) {
            container.innerHTML = html;
            // Re-bind events for reasons
            container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                cb.addEventListener('change', (e) => {
                    Krankomat.State.updateNested('absenceReasons', e.target.dataset.reason, e.target.checked);
                });
            });
        }
    },

    renderRecipientsList: function(recipients) {
        const container = document.getElementById('recipients-list-container');
        if (!recipients || !container) return;
        
        // If container is empty, build from scratch
        if (container.children.length === 0) {
             container.innerHTML = recipients.map(r => this.buildRecipientRow(r)).join('');
             return;
        }

        // Update existing DOM to preserve focus
        // This requires row count to be consistent, which it is in this app structure (fixed recipient list)
        recipients.forEach(r => {
            const checkbox = container.querySelector(`input[data-recipient-id="${r.id}"]`);
            const emailInput = container.querySelector(`input[data-recipient-email-id="${r.id}"]`);

            if (checkbox && checkbox.checked !== r.isSelected) {
                checkbox.checked = r.isSelected;
            }

            // Only update value if it differs AND it is NOT the active element (to prevent cursor jumps)
            if (emailInput && emailInput.value !== (r.email || '') && document.activeElement !== emailInput) {
                emailInput.value = r.email || '';
            }
        });
    },

    buildRecipientRow: function(r) {
        return `
            <div class="p-3 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors">
                <label class="flex items-center space-x-3 cursor-pointer group">
                    <div class="relative flex items-center">
                        <input type="checkbox" data-recipient-id="${r.id}" ${r.isSelected ? 'checked' : ''} class="peer relative h-5 w-5 cursor-pointer appearance-none rounded-md border border-slate-400 dark:border-slate-500 transition-all checked:border-indigo-600 checked:bg-indigo-600 dark:checked:border-indigo-500 dark:checked:bg-indigo-500" />
                        <div class="pointer-events-none absolute top-2/4 left-2/4 -translate-y-2/4 -translate-x-2/4 text-white opacity-0 transition-opacity peer-checked:opacity-100">
                             <svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" viewBox="0 0 20 20" fill="currentColor" stroke="currentColor" strokeWidth="1">
                                <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd"></path>
                             </svg>
                        </div>
                    </div>
                    <span class="text-sm text-slate-700 dark:text-slate-300 group-hover:text-indigo-600 dark:group-hover:text-indigo-400 transition-colors">
                        ${r.module}
                    </span>
                </label>
                <input type="email" data-recipient-email-id="${r.id}" value="${r.email || ''}" placeholder="E-Mail-Adresse eingeben" class="ml-8 mt-2 w-[calc(100%-2rem)] px-3 py-1.5 bg-white dark:bg-slate-700 text-slate-900 dark:text-white border border-slate-300 dark:border-slate-600 rounded-md shadow-sm text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 transition" />
            </div>
        `;
    },

    renderCommentPresets: function(currentComments) {
        const presets = ["nur 1 Tag krank", "vsl. ungefähr 3 Tage krank", "vsl. länger als 3 Tage, Arzt wird aufgesucht"];
        const container = document.getElementById('comment-presets-container');
        if (!container) return;
        
        const html = presets.map(preset => {
            const isChecked = currentComments && currentComments.includes(preset);
            return `
            <label class="flex items-center space-x-3 cursor-pointer group">
                 <div class="relative flex items-center">
                    <input type="checkbox" data-preset="${preset}" ${isChecked ? 'checked' : ''} class="peer relative h-5 w-5 cursor-pointer appearance-none rounded-md border border-slate-400 dark:border-slate-500 transition-all checked:border-indigo-600 checked:bg-indigo-600 dark:checked:border-indigo-500 dark:checked:bg-indigo-500" />
                    <div class="pointer-events-none absolute top-2/4 left-2/4 -translate-y-2/4 -translate-x-2/4 text-white opacity-0 transition-opacity peer-checked:opacity-100">
                        <svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" viewBox="0 0 20 20" fill="currentColor" stroke="currentColor" strokeWidth="1">
                            <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd"></path>
                        </svg>
                    </div>
                </div>
                <span class="text-sm text-slate-700 dark:text-slate-300 group-hover:text-indigo-600 dark:group-hover:text-indigo-400 transition-colors">
                    ${preset}
                </span>
            </label>`;
        }).join('');

        if (container.innerHTML !== html) {
            container.innerHTML = html;
            container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                cb.addEventListener('change', (e) => {
                    const text = e.target.dataset.preset;
                    const current = Krankomat.State.get('details').comments || '';
                    const lines = current.split('\n').filter(line => line.trim() !== '');
                    let newComments;
                    if (lines.includes(text)) {
                        newComments = lines.filter(c => c !== text).join('\n');
                    } else {
                        newComments = [...lines, text].join('\n');
                    }
                    Krankomat.State.updateNested('details', 'comments', newComments);
                });
            });
        }
    }
};