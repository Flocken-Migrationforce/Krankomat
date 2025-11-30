
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