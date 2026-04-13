
window.Krankomat = window.Krankomat || {};

Krankomat.Colors = {
    isDarkMode: false,
    
    // Define available palettes
    palettes: {
        indigo: { // Default
            50: '#eef2ff', 100: '#e0e7ff', 200: '#c7d2fe', 300: '#a5b4fc', 400: '#818cf8',
            500: '#6366f1', 600: '#4f46e5', 700: '#4338ca', 800: '#3730a3', 900: '#312e81', 950: '#1e1b4b'
        },
        blue: {
            50: '#eff6ff', 100: '#dbeafe', 200: '#bfdbfe', 300: '#93c5fd', 400: '#60a5fa',
            500: '#3b82f6', 600: '#2563eb', 700: '#1d4ed8', 800: '#1e40af', 900: '#1e3a8a', 950: '#172554'
        },
        rose: {
            50: '#fff1f2', 100: '#ffe4e6', 200: '#fecdd3', 300: '#fda4af', 400: '#fb7185',
            500: '#f43f5e', 600: '#e11d48', 700: '#be123c', 800: '#9f1239', 900: '#881337', 950: '#4c0519'
        },
        emerald: {
            50: '#ecfdf5', 100: '#d1fae5', 200: '#a7f3d0', 300: '#6ee7b7', 400: '#34d399',
            500: '#10b981', 600: '#059669', 700: '#047857', 800: '#065f46', 900: '#064e3b', 950: '#022c22'
        },
        amber: {
            50: '#fffbeb', 100: '#fef3c7', 200: '#fde68a', 300: '#fcd34d', 400: '#fbbf24',
            500: '#f59e0b', 600: '#d97706', 700: '#b45309', 800: '#92400e', 900: '#78350f', 950: '#451a03'
        },
        violet: {
            50: '#f5f3ff', 100: '#ede9fe', 200: '#ddd6fe', 300: '#c4b5fd', 400: '#a78bfa',
            500: '#8b5cf6', 600: '#7c3aed', 700: '#6d28d9', 800: '#5b21b6', 900: '#4c1d95', 950: '#2e1065'
        }
    },

    init: function() {
        // Safe localStorage access for Theme Mode
        let savedTheme = null;
        try {
            savedTheme = localStorage.getItem('theme');
        } catch (e) {
            console.warn('Cannot access localStorage', e);
        }

        // Check preference for dark mode
        const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
        this.isDarkMode = savedTheme === 'dark' || (!savedTheme && prefersDark);
        
        this.apply();
        this.renderToggle();
        
        // Initialize Color Palette from State
        const config = Krankomat.State.get('config') || {};
        const paletteName = config.colorTheme || 'indigo';
        this.setPalette(paletteName);

        const btn = document.getElementById('theme-toggle-btn');
        if (btn) {
            btn.onclick = (e) => {
                e.preventDefault();
                this.toggle();
            };
        }
    },
    
    setPalette: function(colorName) {
        const palette = this.palettes[colorName] || this.palettes.indigo;
        const root = document.documentElement;
        
        // Set CSS Variables for the tailwind config to pick up
        Object.entries(palette).forEach(([shade, value]) => {
            root.style.setProperty(`--color-primary-${shade}`, value);
        });
        
        // Update Favicon
        this.updateFavicon(colorName);

        // Update State if changed
        const currentConfig = Krankomat.State.get('config') || {};
        if (currentConfig.colorTheme !== colorName) {
            Krankomat.State.updateNested('config', 'colorTheme', colorName);
        }
    },

    updateFavicon: function(colorName) {
        const palette = this.palettes[colorName] || this.palettes.indigo;
        const color = palette[600];
        // Path matches the logo in index.html
        const svg = `<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='${color}'><path d='M19 14h-3v3h-2v-3h-3v-2h3V9h2v3h3m3-9H2c-1.11 0-2 .89-2 2v12a2 2 0 0 0 2 2h12.1c-.08-.33-.12-.66-.12-1c0-2.71 1.7-5.02 4-6.09V6l-8 5l-8-5h16v2.09c.72.2 1.39.57 2 .9V6c0-1.11-.89-2-2-2m-2 2L12 9L4 5h16Z'/></svg>`;
        
        let link = document.querySelector("link[rel~='icon']");
        if (!link) {
            link = document.createElement('link');
            link.rel = 'icon';
            document.head.appendChild(link);
        }
        link.href = `data:image/svg+xml,${encodeURIComponent(svg)}`;
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
