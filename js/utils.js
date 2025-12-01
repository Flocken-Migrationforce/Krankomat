

// Namespace setup
window.Krankomat = window.Krankomat || {};

Krankomat.Utils = {
    todayFormatted: () => {
        const today = new Date();
        const day = String(today.getDate()).padStart(2, '0');
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const year = today.getFullYear();
        return `${day}.${month}.${year}`;
    },

    getDayOfWeek: (dateStr) => {
        const parts = dateStr.match(/(\d{2})\.(\d{2})\.(\d{4})/);
        if (!parts) return null;
        const date = new Date(`${parts[3]}-${parts[2]}-${parts[1]}`);
        if (isNaN(date.getTime())) return null;
        const days = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
        return days[date.getDay()];
    },
    
    // Helper to parse ICS date string (YYYYMMDDTHHMMSS or YYYYMMDD) to JS Date
    parseIcsDateStringToJsDate: (str) => {
        if (!str) return null;
        const clean = str.replace('Z', '').trim();
        const y = parseInt(clean.substring(0,4));
        const m = parseInt(clean.substring(4,6)) - 1; // Months are 0-indexed
        const d = parseInt(clean.substring(6,8));
        
        let h = 0, min = 0, s = 0;
        if (clean.includes('T')) {
            const timePart = clean.split('T')[1];
            if (timePart.length >= 2) h = parseInt(timePart.substring(0,2));
            if (timePart.length >= 4) min = parseInt(timePart.substring(2,4));
            if (timePart.length >= 6) s = parseInt(timePart.substring(4,6));
        }
        return new Date(y, m, d, h, min, s);
    },

    // Helper to format JS Date back to ICS string YYYYMMDDTHHMMSS
    formatJsDateToIcsString: (date) => {
        const pad = n => n.toString().padStart(2, '0');
        return `${date.getFullYear()}${pad(date.getMonth()+1)}${pad(date.getDate())}T${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
    },

    // ICS Parsing Utility - COMPACT MODE
    // Stores only the unique events with their RRULE strings, does not expand them.
    parseICS: (icsContent) => {
        const events = [];
        
        // RFC 5545 Unfolding
        const lines = icsContent.split(/\r\n|\n|\r/);
        const unfoldedLines = [];
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            if (line.length === 0) continue;
            
            if (line.startsWith(' ') || line.startsWith('\t')) {
                if (unfoldedLines.length > 0) {
                    unfoldedLines[unfoldedLines.length - 1] += line.substring(1);
                }
            } else {
                unfoldedLines.push(line);
            }
        }
        
        let inEvent = false;
        let currentEvent = {};
        
        for (const rawLine of unfoldedLines) {
            const line = rawLine.trim();
            const upperLine = line.toUpperCase();
            
            // Check boundaries
            if (upperLine.startsWith('BEGIN:VEVENT')) {
                inEvent = true;
                currentEvent = {};
                continue;
            }
            if (upperLine.startsWith('END:VEVENT')) {
                inEvent = false;
                if (currentEvent.start && currentEvent.summary) {
                    events.push({...currentEvent});
                }
                continue;
            }

            if (inEvent) {
                if (upperLine.startsWith('DTSTART')) {
                    const idx = line.indexOf(':');
                    if (idx !== -1) {
                         currentEvent.start = line.substring(idx + 1).trim();
                    }
                } else if (upperLine.startsWith('SUMMARY')) {
                    const idx = line.indexOf(':');
                    if (idx !== -1) {
                        let summaryVal = line.substring(idx + 1).trim();
                        // Handle "Code:Name" convention
                        const codeIdx = summaryVal.indexOf(':');
                        if (codeIdx !== -1) {
                            currentEvent.summary = summaryVal.substring(codeIdx + 1).trim();
                        } else {
                            currentEvent.summary = summaryVal;
                        }
                    }
                } else if (upperLine.startsWith('LOCATION')) {
                    const idx = line.indexOf(':');
                    if (idx !== -1) {
                         // Extract and clean location (unescape commas)
                         currentEvent.location = line.substring(idx + 1).trim().replace(/\\,/g, ',');
                    }
                } else if (upperLine.startsWith('RRULE')) {
                    const idx = line.indexOf(':');
                    if (idx !== -1) {
                        // Store the RRULE string for on-demand calculation
                        currentEvent.rrule = line.substring(idx + 1).trim();
                    }
                }
            }
        }
        return events;
    },

    // Check if an event falls on a specific date (DD.MM.YYYY)
    // Handles base date matching and RRULE logic (Weekly only)
    isEventOnDate: (event, targetDateStr) => {
        if (!event.start || !targetDateStr) return false;

        // 1. Parse dates to comparable objects (midnight)
        const parts = targetDateStr.split('.');
        if (parts.length !== 3) return false;
        
        // Target date at midnight
        const targetDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
        
        // Event start date at midnight
        const startDateFull = Krankomat.Utils.parseIcsDateStringToJsDate(event.start);
        if (!startDateFull) return false;
        
        const startDate = new Date(startDateFull.getFullYear(), startDateFull.getMonth(), startDateFull.getDate());

        // Optimization: If target is before start, impossible
        if (targetDate < startDate) return false;

        // 2. Base match
        // Calculate difference in days
        const diffTime = targetDate.getTime() - startDate.getTime();
        const diffDays = Math.round(diffTime / (1000 * 3600 * 24));

        if (diffDays === 0) return true;

        // 3. RRULE logic
        if (!event.rrule) return false;

        const rrule = event.rrule;
        const rruleParts = rrule.split(';').reduce((acc, part) => {
            const [key, val] = part.split('=');
            acc[key] = val;
            return acc;
        }, {});

        // Only support WEEKLY for now
        if (rruleParts['FREQ'] !== 'WEEKLY') return false;

        const interval = rruleParts['INTERVAL'] ? parseInt(rruleParts['INTERVAL']) : 1;
        
        // Check if days diff aligns with weekly interval
        if (diffDays % (7 * interval) !== 0) return false;
        
        // Check UNTIL
        if (rruleParts['UNTIL']) {
            const untilStr = rruleParts['UNTIL'];
            const untilDate = Krankomat.Utils.parseIcsDateStringToJsDate(untilStr);
            
            // Fix for inclusive UNTIL date-only
            if (untilStr.replace('Z','').trim().length === 8) {
                untilDate.setHours(23, 59, 59, 999);
            }
            
            // Compare full target date (which is midnight) vs Until
            if (targetDate > untilDate) return false;
        }

        // Check COUNT
        if (rruleParts['COUNT']) {
            const count = parseInt(rruleParts['COUNT']);
            const occurrences = (diffDays / (7 * interval)) + 1; // +1 because base is #1
            if (occurrences > count) return false;
        }

        return true;
    },

    // Convert ICS date (YYYYMMDD or YYYYMMDDTHHMMSS) to app format (DD.MM.YYYY)
    convertICSDateToAppDate: (icsDateStr) => {
        if (!icsDateStr) return null;
        // Basic parsing for YYYYMMDD
        const year = icsDateStr.substring(0, 4);
        const month = icsDateStr.substring(4, 6);
        const day = icsDateStr.substring(6, 8);
        return `${day}.${month}.${year}`;
    },

    renderTemplate: (template, context) => {
        return Object.entries(context).reduce((acc, [key, value]) => {
            return acc.replace(new RegExp(`{${key}}`, 'g'), value || '');
        }, template);
    },

    copyToClipboard: (text, onComplete) => {
        navigator.clipboard.writeText(text).then(() => {
            if (onComplete) onComplete();
        }).catch(err => {
            console.error('Failed to copy text: ', err);
            // Fallback for local file access if clipboard API fails
            try {
                const textArea = document.createElement("textarea");
                textArea.value = text;
                document.body.appendChild(textArea);
                textArea.select();
                document.execCommand("copy");
                document.body.removeChild(textArea);
                if (onComplete) onComplete();
            } catch (e) {
                console.error('Fallback copy failed', e);
            }
        });
    },

    // Icons SVG Strings
    Icons: {
        copy: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" fill="currentColor" class="h-4 w-4"><path d="M208 0H332.1c12.7 0 24.9 5.1 33.9 14.1l67.9 67.9c9 9 14.1 21.2 14.1 33.9V336c0 26.5-21.5 48-48 48H208c-26.5 0-48-21.5-48-48V48c0-26.5 21.5-48 48-48zM48 128h80v64H64V448H256V416h64v48c0 26.5-21.5 48-48 48H48c-26.5 0-48-21.5-48-48V176c0-26.5 21.5-48 48-48z"/></svg>',
        check: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" fill="currentColor" class="h-4 w-4 text-green-500"><path d="M438.6 105.4c12.5 12.5 12.5 32.8 0 45.3l-256 256c-12.5 12.5-32.8 12.5-45.3 0l-128-128c-12.5-12.5-12.5-32.8 0-45.3s32.8-12.5 45.3 0L160 338.7 393.4 105.4c12.5-12.5 32.8-12.5 45.3 0z"/></svg>',
        moon: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 384 512" fill="currentColor" class="h-4 w-4 text-indigo-500 transition-opacity duration-200"><path d="M223.5 32C100 32 0 132.3 0 256S100 480 223.5 480c60.6 0 115.5-24.2 155.8-63.4c5-4.9 6.3-12.5 3.1-18.7s-10.1-9.7-17-8.5c-9.8 1.7-20 2.6-30.6 2.6c-96.8 0-175.5-78.8-175.5-176c0-65.8 36-123.1 89.3-153.3c6.1-3.5 9.2-10.5 7.7-17.3s-7.3-11.9-14.3-12.5c-6.3-.5-12.6-.8-19-.8z"/></svg>',
        sun: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" class="h-5 w-5 text-amber-500 transition-opacity duration-200"><path fillRule="evenodd" d="M12 2.25c.414 0 .75.336.75.75v2.25a.75.75 0 01-1.5 0V3a.75.75 0 01.75-.75zM7.5 12a4.5 4.5 0 119 0 4.5 4.5 0 01-9 0zM18.894 6.106a.75.75 0 010 1.06l-1.591 1.592a.75.75 0 01-1.061-1.061l1.591-1.591a.75.75 0 011.06 0zM21.75 12a.75.75 0 01-.75.75h-2.25a.75.75 0 010-1.5H21a.75.75 0 01.75.75zM17.836 17.894a.75.75 0 011.06 0l1.591 1.591a.75.75 0 01-1.06 1.061l-1.591-1.591a.75.75 0 010-1.06zM12 18a.75.75 0 01.75.75v2.25a.75.75 0 01-1.5 0V18.75a.75.75 0 01.75-.75zM4.106 17.894a.75.75 0 010-1.06l1.591-1.591a.75.75 0 111.061 1.06l-1.591 1.592a.75.75 0 01-1.06 0zM3 12a.75.75 0 01.75-.75h2.25a.75.75 0 010 1.5H3.75A.75.75 0 013 12zM6.106 6.106a.75.75 0 011.06 0l1.591 1.591a.75.75 0 01-1.06 1.061L6.106 7.167a.75.75 0 010-1.06z" clipRule="evenodd" /></svg>'
    }
};