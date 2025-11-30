
window.Krankomat = window.Krankomat || {};

Krankomat.Colors = {
    isDarkMode: false,

    init: function() {
        // Safe localStorage access
        let savedTheme = null;
        try {
            savedTheme = localStorage.getItem('theme');
        } catch (e) {
            console.warn('Cannot access localStorage', e);
        }

        // Check preference
        const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
        this.isDarkMode = savedTheme === 'dark' || (!savedTheme && prefersDark);
        
        this.apply();
        this.renderToggle();

        const btn = document.getElementById('theme-toggle-btn');
        if (btn) {
            // Use arrow function or bind to preserve 'this' context
            btn.onclick = (e) => {
                e.preventDefault(); // Prevent potential form submission
                this.toggle();
            };
        }
    },

    toggle: function() {
        this.isDarkMode = !this.isDarkMode;
        
        try {
            localStorage.setItem('theme', this.isDarkMode ? 'dark' : 'light');
        } catch (e) {
            console.warn('Cannot save to localStorage', e);
        }
        
        this.apply();
        this.renderToggle();
    },

    apply: function() {
        if (this.isDarkMode) {
            document.documentElement.classList.add('dark');
        } else {
            document.documentElement.classList.remove('dark');
        }
    },

    renderToggle: function() {
        const indicator = document.getElementById('theme-toggle-indicator');
        if (!indicator) return;
        
        if (this.isDarkMode) {
            indicator.classList.remove('translate-x-1');
            indicator.classList.add('translate-x-7');
            if (Krankomat.Utils && Krankomat.Utils.Icons) {
                indicator.innerHTML = Krankomat.Utils.Icons.moon;
            }
        } else {
            indicator.classList.remove('translate-x-7');
            indicator.classList.add('translate-x-1');
            if (Krankomat.Utils && Krankomat.Utils.Icons) {
                indicator.innerHTML = Krankomat.Utils.Icons.sun;
            }
        }
    }
};
