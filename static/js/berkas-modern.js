// View Toggle Script (Grid, List)
document.addEventListener('DOMContentLoaded', () => {
    const viewToggleButtons = document.querySelectorAll('.toggle-btn');
    const fileGrid = document.getElementById('fileGrid');
    const searchInput = document.getElementById('searchInput');
    const fileCards = document.querySelectorAll('.file-card');

    // View Toggle
    viewToggleButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const view = btn.dataset.view;

            // Update active state
            viewToggleButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            // Toggle view
            if (view === 'list') {
                fileGrid.classList.add('list-view');
            } else {
                fileGrid.classList.remove('list-view');
            }

            // Save preference
            localStorage.setItem('fileViewPreference', view);
        });
    });

    // Restore view preference
    const savedView = localStorage.getItem('fileViewPreference');
    if (savedView === 'list') {
        document.querySelector('[data-view="list"]')?.click();
    }

    // Theme Toggle
    const themeToggle = document.getElementById('themeToggle');
    if (themeToggle) {
        themeToggle.addEventListener('click', () => {
            const currentTheme = document.documentElement.className;
            const newTheme = currentTheme === 'dark-mode' ? 'light' : 'dark';

            document.documentElement.className = newTheme === 'dark' ? 'dark-mode' : 'light-mode';
            localStorage.setItem('theme', newTheme);
        });
    }

    // Search Functionality with server-side recursive search
    let searchTimeout;
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            const searchTerm = e.target.value.trim();

            // Clear previous timeout
            clearTimeout(searchTimeout);

            if (!searchTerm) {
                // Reset to showing all cards
                fileCards.forEach(card => {
                    card.style.display = '';
                });
                const searchEmpty = document.querySelector('.search-empty');
                if (searchEmpty) {
                    searchEmpty.remove();
                    const grid = document.querySelector('.file-grid');
                    if (grid) grid.style.display = 'grid';
                }
                return;
            }

            // Debounce search
            searchTimeout = setTimeout(async () => {
                try {
                    // Get current path from URL
                    const currentPath = window.location.pathname.replace('/berkas/', '').replace(/\/$/, '');

                    // Fetch search results from server
                    const response = await fetch(`/berkas/search?q=${encodeURIComponent(searchTerm)}&path=${encodeURIComponent(currentPath)}`);
                    const data = await response.json();

                    // Hide all current cards
                    fileCards.forEach(card => card.style.display = 'none');

                    // Remove previous search results
                    document.querySelectorAll('.search-result-card').forEach(card => card.remove());

                    if (data.results && data.results.length > 0) {
                        const fileGrid = document.getElementById('fileGrid');

                        // Add search results
                        data.results.forEach(result => {
                            const card = document.createElement('div');
                            card.className = `file-card ${result.type === 'folder' ? 'folder-card' : 'file-card-item'} search-result-card`;

                            const url = result.type === 'folder'
                                ? `/berkas/${result.path}/`
                                : `/berkas/${result.path}`;

                            const icon = result.type === 'folder'
                                ? `<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z" />
                                   </svg>`
                                : `<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z" />
                                   </svg>`;

                            const parentPath = result.parent === '.' ? 'Root' : result.parent;
                            const meta = result.type === 'file'
                                ? `<span class="file-date">${result.modified}</span><span class="file-size">${result.size}</span>`
                                : `<span class="file-date">in: ${parentPath}</span>`;

                            card.innerHTML = `
                                <a href="${url}" class="file-link">
                                    <div class="file-icon">${icon}</div>
                                    <div class="file-info">
                                        <div class="file-name" title="${result.name}">${result.name}</div>
                                        <div class="file-meta">${meta}</div>
                                    </div>
                                </a>
                            `;

                            fileGrid.appendChild(card);
                        });

                        // Remove empty state if exists
                        const searchEmpty = document.querySelector('.search-empty');
                        if (searchEmpty) searchEmpty.remove();
                    } else {
                        // Show no results message
                        const container = document.querySelector('.file-container');
                        const grid = document.querySelector('.file-grid');

                        const noResults = document.createElement('div');
                        noResults.className = 'empty-state search-empty';
                        noResults.innerHTML = `
                            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
                            </svg>
                            <h3>No Results Found</h3>
                            <p>No files or folders match "${searchTerm}" in this directory and its subdirectories</p>
                        `;

                        if (grid) grid.style.display = 'none';
                        container.appendChild(noResults);
                    }
                } catch (error) {
                    console.error('Search error:', error);
                }
            }, 300); // 300ms debounce
        });
    }

    // Keyboard Navigation
    document.addEventListener('keydown', (e) => {
        // Focus search on '/' key
        if (e.key === '/' && !e.ctrlKey && !e.metaKey) {
            e.preventDefault();
            searchInput?.focus();
        }

        // Clear search on Escape
        if (e.key === 'Escape' && document.activeElement === searchInput) {
            searchInput.value = '';
            searchInput.dispatchEvent(new Event('input'));
            searchInput.blur();
        }
    });

    // Add hover effect enhancements
    fileCards.forEach(card => {
        card.addEventListener('mouseenter', function () {
            this.style.transform = 'translateY(-4px)';
        });

        card.addEventListener('mouseleave', function () {
            this.style.transform = 'translateY(0)';
        });
    });

    // Add ripple effect on click
    fileCards.forEach(card => {
        card.addEventListener('click', function (e) {
            const ripple = document.createElement('div');
            ripple.style.position = 'absolute';
            ripple.style.borderRadius = '50%';
            ripple.style.background = 'rgba(255, 255, 255, 0.2)';
            ripple.style.width = '20px';
            ripple.style.height = '20px';
            ripple.style.animation = 'ripple 0.6s ease-out';
            ripple.style.pointerEvents = 'none';

            const rect = this.getBoundingClientRect();
            ripple.style.left = (e.clientX - rect.left - 10) + 'px';
            ripple.style.top = (e.clientY - rect.top - 10) + 'px';

            this.style.position = 'relative';
            this.appendChild(ripple);

            setTimeout(() => ripple.remove(), 600);
        });
    });

    // Add CSS for ripple animation dynamically
    const style = document.createElement('style');
    style.textContent = `
        @keyframes ripple {
            to {
                transform: scale(4);
                opacity: 0;
            }
        }
    `;
    document.head.appendChild(style);

    // Smooth scroll for better UX
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            const target = document.querySelector(this.getAttribute('href'));
            if (target) {
                target.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        });
    });

    console.log('File browser initialized with modern UI and recursive search');
});
