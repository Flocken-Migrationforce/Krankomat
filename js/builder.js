

window.Krankomat = window.Krankomat || {};

Krankomat.Builder = {
    init: function() {
        this.bindEvents();
        this.render();
        
        // Immediately calculate recipients based on the loaded date/calendar to auto-populate
        this.recalculateRecipients();
        
        // Bind recipients events once here instead of inside render
        const recipientsContainer = document.getElementById('recipients-list-container');
        if (recipientsContainer) {
             this.bindRecipientEvents(recipientsContainer);
        }
    },

    bindEvents: function() {
        // User Data Inputs
        ['firstName', 'lastName'].forEach(field => {
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
                this.recalculateRecipients();
            });
        }
        
        const endEl = document.getElementById('endDate');
        if (endEl) {
            endEl.addEventListener('input', (e) => {
                Krankomat.State.set('sicknessEndDate', e.target.value);
            });
        }

        // ICS Upload Handling is now in App.js / Settings Modal
        // but we expose the logic handler here
        
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

        // Reset Date Button (Formerly Manual Save)
        const resetBtn = document.getElementById('btn-reset-date');
        if (resetBtn) {
            resetBtn.addEventListener('click', () => {
                const today = Krankomat.Utils.todayFormatted();
                Krankomat.State.set('sicknessStartDate', today);
                this.recalculateRecipients();
                
                // Visual feedback
                const originalContent = resetBtn.innerHTML;
                resetBtn.innerHTML = Krankomat.Utils.Icons.check + ' Zurückgesetzt!';
                resetBtn.classList.add('bg-green-100', 'text-green-700');
                resetBtn.classList.remove('bg-indigo-100', 'text-indigo-700');
                setTimeout(() => {
                    resetBtn.innerHTML = originalContent;
                    resetBtn.classList.remove('bg-green-100', 'text-green-700');
                    resetBtn.classList.add('bg-indigo-100', 'text-indigo-700');
                }, 1500);
            });
        }
    },

    handleIcsUpload: function(file) {
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
            try {
                // Modified to return compacted events (not expanded)
                const events = Krankomat.Utils.parseICS(ev.target.result);
                Krankomat.State.set('calendarEvents', events);
                
                // Auto-expand Email Directory with new courses
                const directory = Krankomat.State.get('emailDirectory') || {};
                let dirChanged = false;
                events.forEach(evt => {
                    if (evt.summary && !directory.hasOwnProperty(evt.summary)) {
                        directory[evt.summary] = ""; // Initialize with empty email
                        dirChanged = true;
                    }
                });
                if (dirChanged) {
                    Krankomat.State.set('emailDirectory', directory);
                }

                // Re-evaluate recipients with new data
                this.recalculateRecipients();
                
                alert(`Erfolg: ${events.length} Kalendereinträge importiert.`);
            } catch (err) {
                alert("Fehler beim Lesen der Kalenderdatei.");
                console.error(err);
            }
        };
        reader.readAsText(file);
    },
    
    bindRecipientEvents: function(container) {
        container.addEventListener('change', (e) => {
            if (e.target.dataset.recipientId) {
                const id = e.target.dataset.recipientId; // String match for non-numeric IDs like 'exam-office'
                const currentRecipients = Krankomat.State.get('recipients');
                const newRecipients = currentRecipients.map(r => r.id.toString() === id ? { ...r, isSelected: e.target.checked } : r);
                
                // Sort after toggle to keep selected at top
                this.sortAndSetRecipients(newRecipients);
            }
        });

        container.addEventListener('input', (e) => {
             if (e.target.dataset.recipientEmailId) {
                const id = e.target.dataset.recipientEmailId;
                const currentRecipients = Krankomat.State.get('recipients');
                const newRecipients = currentRecipients.map(r => r.id.toString() === id ? { ...r, email: e.target.value } : r);
                
                // If user manually enters email, save it to directory for future use
                const recipient = currentRecipients.find(r => r.id.toString() === id);
                if (recipient && recipient.module) {
                    const dir = Krankomat.State.get('emailDirectory') || {};
                    dir[recipient.module] = e.target.value;
                    Krankomat.State.set('emailDirectory', dir);
                }

                Krankomat.State.set('recipients', newRecipients);
             }
        });
    },

    sortAndSetRecipients: function(recipients) {
        // Sort: ZPD (id 1), Exam (id exam-office), then Selected, then Alphabetical
        const sorted = [...recipients].sort((a, b) => {
            // 1. ZPD always top
            if (a.id === 1) return -1;
            if (b.id === 1) return 1;
            
            // 2. Exam Office second
            if (a.id === 'exam-office') return -1;
            if (b.id === 'exam-office') return 1;

            // 3. Selected checked boxes first
            if (a.isSelected && !b.isSelected) return -1;
            if (!a.isSelected && b.isSelected) return 1;

            // 4. Alphabetical by module name
            return (a.module || '').localeCompare(b.module || '');
        });
        Krankomat.State.set('recipients', sorted);
    },

    capitalizeName: function(str) {
        if (!str) return '';
        // Handles hyphenated names like "Ann-Kathrin"
        return str.split('-').map(part => 
            part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()
        ).join('-');
    },

    extractNameFromEmail: function(email) {
        if (!email || !email.includes('@')) return null;
        
        const localPart = email.split('@')[0];
        // Remove numbers (e.g. m.muster2)
        const cleanLocal = localPart.replace(/[0-9]/g, '');
        
        const parts = cleanLocal.split('.');
        // require at least two parts (first.last) to form a name
        if (parts.length < 2) return null; 

        const nameParts = parts
            .filter(p => p.length > 0)
            .map(p => this.capitalizeName(p));
            
        if (nameParts.length === 0) return null;
        
        return nameParts.join(' ');
    },

    // Central function to rebuild the recipients list based on Date AND Exam status
    recalculateRecipients: function() {
        const data = Krankomat.State.data;
        const config = data.config || {};
        const showAll = config.showAllRecipients;

        const rawDateStr = data.sicknessStartDate || '';
        
        // Normalize Date String to DD.MM.YYYY (ensure padding)
        const normalizeDate = (str) => {
            if (!str) return '';
            const parts = str.split('.');
            if (parts.length !== 3) return str;
            return `${parts[0].padStart(2, '0')}.${parts[1].padStart(2, '0')}.${parts[2]}`;
        };
        const dateStr = normalizeDate(rawDateStr);

        const calendarEvents = data.calendarEvents || [];
        const emailDirectory = data.emailDirectory || {};
        const isExam = data.absenceReasons && data.absenceReasons.exam;

        const zpd = { id: 1, anrede: 'Sehr geehrte Damen und Herren vom ZPD', module: 'ZPD (Zentrum für Personaldienste)', isSelected: true, email: '' };
        
        // Preserve ZPD email if typed
        const currentRecipients = data.recipients || [];
        const currentZPD = currentRecipients.find(r => r.id === 1);
        if (currentZPD && currentZPD.email) zpd.email = currentZPD.email;

        let newRecipients = [zpd];

        // 1. Exam / Prüfungsamt
        if (isExam) {
            const examOffice = {
                id: 'exam-office',
                anrede: 'Sehr geehrte Damen und Herren vom Prüfungsamt',
                module: 'Prüfungsamt (FSB PuMa)',
                isSelected: true,
                email: 'pruefungsamt-berlinertor-sp2@haw-hamburg.de'
            };
            newRecipients.push(examOffice);
        }

        // 2. Calendar / Timetable / All Recipients
        // Determine "Today's Modules" for auto-selection purposes
        let todaysModules = [];
        if (calendarEvents.length > 0) {
             const dayEvents = calendarEvents.filter(evt => Krankomat.Utils.isEventOnDate(evt, dateStr));
             todaysModules = dayEvents.map(e => e.summary); 
        } else {
             const day = Krankomat.Utils.getDayOfWeek(dateStr);
             if (day) {
                const timetable = Krankomat.State.defaults.timetable;
                todaysModules = timetable.filter(e => e.day.toLowerCase() === day.toLowerCase()).map(e => e.module);
             }
        }

        let modulesToList = [];
        
        if (showAll) {
             // Show all unique keys from directory, plus today's modules if missing from directory
             const allKeys = new Set([...Object.keys(emailDirectory), ...todaysModules]);
             modulesToList = Array.from(allKeys).sort();
        } else {
             // Only show modules relevant for today
             modulesToList = todaysModules;
        }

        const generatedRecipients = modulesToList.map((modName, index) => {
            const knownEmail = emailDirectory[modName] || '';
            let anrede = `Sehr geehrte/r Dozent/in für ${modName}`;
            if (knownEmail) {
                const extractedName = this.extractNameFromEmail(knownEmail);
                if (extractedName) {
                    anrede = `Sehr geehrte/r ${extractedName}`;
                }
            }
            
            // Auto-select logic
            // If showAll=true: Select only if it matches today's modules.
            // If showAll=false: Select true (since it IS today's module).
            let shouldSelect = true;
            if (showAll) {
                shouldSelect = todaysModules.includes(modName);
            }

            return {
                id: index + 2, // Offset ID after ZPD
                anrede: anrede,
                module: modName,
                isSelected: shouldSelect,
                email: knownEmail
            };
        });
        
        newRecipients = [...newRecipients, ...generatedRecipients];
        
        // Apply sorting
        this.sortAndSetRecipients(newRecipients);
    },

    render: function() {
        const data = Krankomat.State.data;
        if (!data.userData) return; // Safety check

        // Input values
        const firstNameInput = document.getElementById('firstName');
        if (firstNameInput && document.activeElement !== firstNameInput) firstNameInput.value = data.userData.firstName || '';
        
        const lastNameInput = document.getElementById('lastName');
        if (lastNameInput && document.activeElement !== lastNameInput) lastNameInput.value = data.userData.lastName || '';
        
        // StudentID is handled dynamically in renderAbsenceReasons for Exam
        
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
        
        if (krankBtn && gesundBtn) {
            if (!isGesund) {
                krankBtn.className = "w-1/2 py-2 text-sm font-semibold rounded-full transition-all duration-300 bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 shadow";
                gesundBtn.className = "w-1/2 py-2 text-sm font-semibold rounded-full transition-all duration-300 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600";
            } else {
                gesundBtn.className = "w-1/2 py-2 text-sm font-semibold rounded-full transition-all duration-300 bg-white dark:bg-slate-800 text-green-600 dark:text-green-400 shadow";
                krankBtn.className = "w-1/2 py-2 text-sm font-semibold rounded-full transition-all duration-300 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600";
            }
        }

        this.renderAbsenceReasons(data.absenceReasons);
        this.renderRecipientsList(data.recipients);
        this.renderCommentPresets(data.details.comments);
        
        // Handle conditional student ID focus retention
        const condStudentId = document.getElementById('conditional-student-id');
        if (condStudentId && document.activeElement !== condStudentId) {
            condStudentId.value = data.userData.studentId || '';
        }
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
                <div class="mt-3 ml-8 p-3 bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-500/30 rounded-lg text-sm text-blue-800 dark:text-blue-200 transition-all duration-300 animate-fade-in">
                    <div class="flex items-start space-x-3 mb-3">
                        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" fill="currentColor" class="h-5 w-5 text-blue-500 dark:text-blue-400 mt-1 flex-shrink-0">
                            <path d="M256 512A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM216 336h24V272H216c-13.3 0-24-10.7-24-24s10.7-24 24-24h48c13.3 0 24 10.7 24 24v88h8c13.3 0 24 10.7 24 24s-10.7 24-24 24H216c-13.3 0-24-10.7-24-24s10.7-24 24-24zm40-208a32 32 0 1 1 0 64 32 32 0 1 1 0-64z"/>
                        </svg>
                        <div class="flex-1 min-w-0">
                            <p class="mb-2"><strong>Wichtiger Hinweis:</strong> Senden Sie bitte zusätzlich Ihre Prüfungsunfähigkeitsbescheinigung an das Prüfungsamt.</p>
                            <p class="text-xs mb-3">Bitte geben Sie hier Ihre Matrikelnummer an, damit sie im E-Mail-Text für das Prüfungsamt erscheint:</p>
                            <p class="text-xs mb-2 text-amber-600 dark:text-amber-400 italic">Die Eingabe einer Matrikelnummer deaktiviert automatisch Dozenten als Empfänger, um Ihre Anonymität zu wahren.</p>
                            <input id="conditional-student-id" type="text" placeholder="Matrikelnummer" class="form-input px-3 py-1.5 w-full bg-white dark:bg-slate-700 border border-slate-300 dark:border-slate-600 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 text-slate-900 dark:text-white" value="">
                        </div>
                    </div>
                </div>` : ''}
            </div>
        `).join('');
        
        // DOM update
        const currentHTML = container.innerHTML;
        if (currentHTML !== html) {
             container.innerHTML = html;
             
             // Bind Checkboxes
             container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                cb.addEventListener('change', (e) => {
                    const reason = e.target.dataset.reason;
                    // Update state, then recalculate recipients (specifically for Exam)
                    Krankomat.State.data.absenceReasons[reason] = e.target.checked;
                    this.recalculateRecipients(); // This also saves and notifies
                    // Force recipients list update if privacy logic changed state
                    this.updatePrivacyState();
                });
            });

            // Bind Conditional Input
            const condInput = document.getElementById('conditional-student-id');
            if (condInput) {
                condInput.value = Krankomat.State.data.userData.studentId || '';
                condInput.addEventListener('input', (e) => {
                    const value = e.target.value;
                    Krankomat.State.data.userData.studentId = value;
                    this.updatePrivacyState();
                });
            }
        }
    },

    updatePrivacyState: function() {
        const studentId = Krankomat.State.data.userData.studentId;
        const isExam = Krankomat.State.data.absenceReasons.exam;
        
        // Privacy Active only if ID present AND Exam checked
        const privacyActive = (studentId && studentId.length > 0) && isExam;
        
        const recipients = Krankomat.State.get('recipients');
        let changed = false;

        const updatedRecipients = recipients.map(r => {
            // Deselect professors (IDs not 1 and not exam-office) if privacy active
            if (privacyActive && r.id !== 1 && r.id !== 'exam-office' && r.isSelected) {
                changed = true;
                return { ...r, isSelected: false };
            }
            return r;
        });
        
        if (changed) {
            Krankomat.State.set('recipients', updatedRecipients);
        } else {
            Krankomat.State.save();
            // Force redraw of recipients list to apply disabled styling based on privacy logic
            this.renderRecipientsList(updatedRecipients);
        }
        
        // Manually trigger preview update to update email body
        Krankomat.Preview.render(); 
    },

    renderRecipientsList: function(recipients) {
        const container = document.getElementById('recipients-list-container');
        if (!recipients || !container) return;
        
        const studentId = Krankomat.State.data.userData.studentId;
        const isExam = Krankomat.State.data.absenceReasons.exam;
        
        // Disable state logic: ID present AND Exam checked
        const privacyActive = (studentId && studentId.length > 0) && isExam;
        
        // Always rebuild fully to ensure disabled state is applied correctly in HTML
        container.innerHTML = recipients.map(r => {
            // Disable if privacy mode is active AND recipient is a professor (not ZPD/Exam)
            const isDisabled = privacyActive && r.id !== 1 && r.id !== 'exam-office';
            return this.buildRecipientRow(r, isDisabled);
        }).join('');
    },

    buildRecipientRow: function(r, isDisabled) {
        // Disabled styling classes
        const opacityClass = isDisabled ? 'opacity-50 grayscale pointer-events-none' : '';
        const disabledAttr = isDisabled ? 'disabled' : '';

        return `
            <div class="p-3 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors ${opacityClass}">
                <label class="flex items-center space-x-3 cursor-pointer group">
                    <div class="relative flex items-center">
                        <input type="checkbox" data-recipient-id="${r.id}" ${r.isSelected ? 'checked' : ''} ${disabledAttr} class="peer relative h-5 w-5 cursor-pointer appearance-none rounded-md border border-slate-400 dark:border-slate-500 transition-all checked:border-indigo-600 checked:bg-indigo-600 dark:checked:border-indigo-500 dark:checked:bg-indigo-500 disabled:cursor-not-allowed disabled:opacity-50" />
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
                <input type="email" data-recipient-email-id="${r.id}" value="${r.email || ''}" ${disabledAttr} placeholder="E-Mail-Adresse eingeben" class="ml-8 mt-2 w-[calc(100%-2rem)] px-3 py-1.5 bg-white dark:bg-slate-700 text-slate-900 dark:text-white border border-slate-300 dark:border-slate-600 rounded-md shadow-sm text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 transition disabled:opacity-50 disabled:cursor-not-allowed" />
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