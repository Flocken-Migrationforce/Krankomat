

window.Krankomat = window.Krankomat || {};

Krankomat.State = {
    defaults: {
        userData: {
            firstName: 'Max',
            lastName: 'Mustermann',
            // StudentId is no longer part of the defaults visible on load
            email: 'max.mustermann@example.com',
            studyProgram: 'E-Government',
        },
        recipientsStructure: [
            { id: 1, anrede: 'Sehr geehrte Damen und Herren vom ZPD', module: 'ZPD (Zentrum für Personaldienste)', isSelected: true },
        ],
        // Default static timetable if no ICS is loaded
        timetable: [
            { day: 'Montag', module: 'IT-Projektmanagement' },
            { day: 'Dienstag', module: 'Grundlagen der BWL' },
            { day: 'Mittwoch', module: 'Datenbanken' },
            { day: 'Donnerstag', module: 'Software Engineering' },
            { day: 'Freitag', module: 'IT-Projektmanagement' },
        ],
        absenceReasons: { lecture: true, internship: false, exam: false, partialDay: false },
        calendarEvents: [], // Stores parsed ICS events
        emailDirectory: { // Maps Course Name to Email
            "IT-Projektmanagement": "prof.it@example.com",
            "Grundlagen der BWL": "prof.bwl@example.com"
        },
        config: {
            showAllRecipients: false,
            headerButtons: {
                fileshare: true,
                calendar: true,
                mensa: true
            },
            supportEmail: "support@krankomat.cloud",
            profileName: "WebApp"
        }
    },

    data: {},
    listeners: [],

    init: function() {
        this.load();
    },

    loadJSONState: function(key, defaultValue) {
        try {
            const item = localStorage.getItem(key);
            // Clone default value to avoid reference mutation of the defaults object
            const safeDefault = JSON.parse(JSON.stringify(defaultValue));
            return item ? JSON.parse(item) : safeDefault;
        } catch (error) {
            console.error(`Failed to load state for ${key}`, error);
            return JSON.parse(JSON.stringify(defaultValue));
        }
    },

    load: function() {
        this.data.userData = this.loadJSONState('krankomat_userData', this.defaults.userData);
        this.data.sicknessStartDate = this.loadJSONState('krankomat_sicknessStartDate', Krankomat.Utils.todayFormatted());
        this.data.sicknessEndDate = this.loadJSONState('krankomat_sicknessEndDate', '');
        this.data.absenceReasons = { ...this.defaults.absenceReasons, ...this.loadJSONState('krankomat_absenceReasons', {}) };
        this.data.details = this.loadJSONState('krankomat_details', { comments: '' });
        
        this.data.calendarEvents = this.loadJSONState('krankomat_calendarEvents', []);
        this.data.emailDirectory = this.loadJSONState('krankomat_emailDirectory', this.defaults.emailDirectory);
        this.data.config = { ...this.defaults.config, ...this.loadJSONState('krankomat_config', {}) };

        // Merge recipients logic
        const baseRecipients = this.defaults.recipientsStructure.map(r => ({ ...r, email: '' }));
        const savedRecipients = this.loadJSONState('krankomat_recipients', []);
        
        if (savedRecipients.length > 0) {
            this.data.recipients = baseRecipients.map(base => {
                const saved = savedRecipients.find(sr => sr.id === base.id);
                return { ...base, ...(saved || {}) };
            });
        } else {
            this.data.recipients = baseRecipients;
        }
    },

    save: function() {
        try {
            localStorage.setItem('krankomat_userData', JSON.stringify(this.data.userData));
            localStorage.setItem('krankomat_sicknessStartDate', JSON.stringify(this.data.sicknessStartDate));
            localStorage.setItem('krankomat_sicknessEndDate', JSON.stringify(this.data.sicknessEndDate));
            localStorage.setItem('krankomat_absenceReasons', JSON.stringify(this.data.absenceReasons));
            localStorage.setItem('krankomat_details', JSON.stringify(this.data.details));
            localStorage.setItem('krankomat_calendarEvents', JSON.stringify(this.data.calendarEvents));
            localStorage.setItem('krankomat_emailDirectory', JSON.stringify(this.data.emailDirectory));
            localStorage.setItem('krankomat_config', JSON.stringify(this.data.config));
            
            if (this.data.recipients && this.data.recipients.length > 0) {
                localStorage.setItem('krankomat_recipients', JSON.stringify(this.data.recipients));
            }
        } catch (e) {
            console.warn('Failed to save state to localStorage:', e);
        }
    },
    
    // API methods
    get: function(key) {
        return this.data[key];
    },

    set: function(key, value) {
        this.data[key] = value;
        this.notify();
    },

    updateNested: function(parentKey, childKey, value) {
        if (this.data[parentKey]) {
            this.data[parentKey][childKey] = value;
            this.notify();
        }
    },

    subscribe: function(callback) {
        this.listeners.push(callback);
    },

    notify: function() {
        this.listeners.forEach(callback => callback(this.data));
        // Auto-save on every change
        this.save();
    }
};