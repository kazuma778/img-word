// Jobie Dashboard Theme Handler - Rewritten for zero FOUC
document.addEventListener('DOMContentLoaded', () => {

    // --- Sidebar Toggle --- //
    const sidebarToggle = document.getElementById('sidebarToggle');
    const sidebar = document.querySelector('.sidebar');
    const mainWrapper = document.querySelector('.main-wrapper');

    if (sidebarToggle && sidebar) {
        sidebarToggle.addEventListener('click', () => {
            sidebar.classList.toggle('is-open');
            mainWrapper.classList.toggle('sidebar-open-overlay');
        });

        // Close sidebar when clicking on the overlay
        mainWrapper.addEventListener('click', (e) => {
            if (mainWrapper.classList.contains('sidebar-open-overlay') && e.target === mainWrapper) {
                sidebar.classList.remove('is-open');
                mainWrapper.classList.remove('sidebar-open-overlay');
            }
        });
    }

    // --- Theme Toggle - Rewritten for consistency --- //
    const darkModeCheckbox = document.getElementById('darkModeCheckbox');
    const htmlElement = document.documentElement;

    // Sync checkbox with current theme (already set by inline script)
    const syncCheckboxWithTheme = () => {
        const currentTheme = localStorage.getItem('theme') || 'dark';
        const isDarkMode = currentTheme === 'dark';

        if (isDarkMode) {
            htmlElement.classList.add('dark-mode');
            htmlElement.classList.remove('light-mode');
            if (darkModeCheckbox) darkModeCheckbox.checked = true;
        } else {
            htmlElement.classList.add('light-mode');
            htmlElement.classList.remove('dark-mode');
            if (darkModeCheckbox) darkModeCheckbox.checked = false;
        }
    };

    // Set initial state
    syncCheckboxWithTheme();

    // Handle theme toggle
    if (darkModeCheckbox) {
        darkModeCheckbox.addEventListener('change', () => {
            const newTheme = darkModeCheckbox.checked ? 'dark' : 'light';

            if (newTheme === 'dark') {
                htmlElement.classList.add('dark-mode');
                htmlElement.classList.remove('light-mode');
            } else {
                htmlElement.classList.add('light-mode');
                htmlElement.classList.remove('dark-mode');
            }

            localStorage.setItem('theme', newTheme);
        });
    }

});