(() => {
    'use strict';

    const SELECTORS = {
        CUSTOM_BUTTON: '#customButton',
        TARGET_TOOLBAR: '.TaskPaneToolbar',
        SVG_ELEMENT: 'svg'
    };

    const CSS_CLASSES = {
        BUTTON_BASE: 'ThemeableRectangularButtonPresentation--isEnabled ThemeableRectangularButtonPresentation--withNoLabel ThemeableRectangularButtonPresentation ThemeableRectangularButtonPresentation--medium TaskPaneToolbar-button TaskPaneToolbar-fullScreenButton BaseLink',
        ICON_CONTAINER: 'IconButtonThemeablePresentation--isEnabled IconButtonThemeablePresentation IconButtonThemeablePresentation--medium SubtleIconButton--standardTheme SubtleIconButton TaskPaneExtraActionsButton TaskPaneToolbar-button HighlightSol HighlightSol--core HighlightSol--buildingBlock Stack Stack--align-center Stack--direction-row Stack--display-inline Stack--justify-center',
        SVG_ICON: 'Icon ButtonThemeablePresentation-leftIcon ExpandIcon HighlightSol HighlightSol--core',
        CLICKED_EFFECT: 'clicked'
    };

    const TIMINGS = {
        CLICK_ANIMATION: 200,
        RETRY_DELAY: 1000,
        URL_CHECK_INTERVAL: 100,
        DEBOUNCE_DELAY: 100
    };

    const SVG_PATH = 'M80 104a24 24 0 1 0 0-48 24 24 0 1 0 0 48zm80-24c0 32.8-19.7 61-48 73.3v87.8c18.8-10.9 40.7-17.1 64-17.1h96c35.3 0 64-28.7 64-64v-6.7C307.7 141 288 112.8 288 80c0-44.2 35.8-80 80-80s80 35.8 80 80c0 32.8-19.7 61-48 73.3V160c0 70.7-57.3 128-128 128H176c-35.3 0-64 28.7-64 64v6.7c28.3 12.3 48 40.5 48 73.3c0 44.2-35.8 80-80 80s-80-35.8-80-80c0-32.8 19.7-61 48-73.3V352 153.3C19.7 141 0 112.8 0 80C0 35.8 35.8 0 80 0s80 35.8 80 80zm232 0a24 24 0 1 0 -48 0 24 24 0 1 0 48 0zM80 456a24 24 0 1 0 0-48 24 24 0 1 0 0 48z';

    /**
     * Extracts the branch number from the current URL
     * @returns {string|null} The branch number or null if not found
     */
    function extractBranchNumber() {
        const url = window.location.href;
        const lastDigitMatch = url.match(/\/(\d+)\/?$/);
        return lastDigitMatch ? lastDigitMatch[1] : null;
    }

    /**
     * Applies visual feedback to the button
     * @param {HTMLElement} button - The button element
     */
    function applyButtonFeedback(button) {
        button.classList.add(CSS_CLASSES.CLICKED_EFFECT);
        setTimeout(() => {
            button.classList.remove(CSS_CLASSES.CLICKED_EFFECT);
        }, TIMINGS.CLICK_ANIMATION);
    }

    /**
     * Handles the button click event
     * @param {Event} event - The click event
     */
    async function handleButtonClick(event) {
        const button = event.currentTarget;
        const branchNumber = extractBranchNumber();

        if (branchNumber) {
            try {
                await navigator.clipboard.writeText(branchNumber);
                applyButtonFeedback(button);
            } catch (error) {
                applyButtonFeedback(button);
            }
        } else {
            applyButtonFeedback(button);
        }
    }

    /**
     * Creates the custom button element
     * @returns {HTMLElement} The created button element
     */
    function createCustomButton() {
        const button = document.createElement('div');
        button.id = 'customButton';
        button.setAttribute('role', 'button');
        button.setAttribute('tabindex', '0');
        button.setAttribute('aria-label', 'Copy branch number');
        button.className = CSS_CLASSES.BUTTON_BASE;

        button.innerHTML = `
            <div class="${CSS_CLASSES.ICON_CONTAINER}">
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" class="${CSS_CLASSES.SVG_ICON}">
                    <path d="${SVG_PATH}"/>
                </svg>
            </div>
        `;

        // Add event listeners
        button.addEventListener('click', handleButtonClick);
        button.addEventListener('keydown', (event) => {
            if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault();
                handleButtonClick(event);
            }
        });

        return button;
    }

    /**
     * Inserts the button into the target toolbar
     * @param {HTMLElement} button - The button element to insert
     */
    function insertButtonIntoToolbar(button) {
        const targetDiv = document.querySelector(SELECTORS.TARGET_TOOLBAR);

        if (!targetDiv) {
            return false;
        }

        const children = targetDiv.children;
        if (children.length > 0) {
            targetDiv.insertBefore(button, children[children.length - 1]);
        } else {
            targetDiv.appendChild(button);
        }

        return true;
    }

    /**
     * Waits for the target toolbar to be available in the DOM
     * @param {number} maxAttempts - Maximum number of attempts
     * @param {number} delay - Delay between attempts in milliseconds
     * @returns {Promise<HTMLElement|null>} The toolbar element or null if not found
     */
    async function waitForToolbar(maxAttempts = 20, delay = 250) {
        for (let attempt = 0; attempt < maxAttempts; attempt++) {
            const toolbar = document.querySelector(SELECTORS.TARGET_TOOLBAR);
            if (toolbar) {
                return toolbar;
            }
            await new Promise(resolve => setTimeout(resolve, delay));
        }
        return null;
    }

    // Global state to prevent multiple button additions
    let isAddingButton = false;
    let debounceTimeout = null;

    /**
     * Debounced version of addCustomButton to prevent multiple calls
     */
    function debouncedAddCustomButton() {
        if (debounceTimeout) {
            clearTimeout(debounceTimeout);
        }

        debounceTimeout = setTimeout(() => {
            addCustomButton();
        }, TIMINGS.DEBOUNCE_DELAY);
    }

    /**
     * Adds the custom button to the page if it doesn't already exist
     */
    async function addCustomButton() {
        // Prevent multiple simultaneous attempts
        if (isAddingButton) {
            return;
        }

        // Check if button already exists
        if (document.querySelector(SELECTORS.CUSTOM_BUTTON)) {
            return;
        }

        isAddingButton = true;

        try {
            // Wait for toolbar to be available
            const toolbar = await waitForToolbar();
            if (!toolbar) {
                return;
            }

            // Double-check button doesn't exist (race condition protection)
            if (document.querySelector(SELECTORS.CUSTOM_BUTTON)) {
                return;
            }

            const button = createCustomButton();
            const inserted = insertButtonIntoToolbar(button);

            if (!inserted) {
                // Retry after delay if insertion failed
                setTimeout(() => {
                    isAddingButton = false;
                    addCustomButton();
                }, TIMINGS.RETRY_DELAY);
            } else {
            }
        } finally {
            isAddingButton = false;
        }
    }

    function initializeUrlMonitoring() {
        let lastUrl = window.location.href;
        let intervalId;

        const checkUrlChange = () => {
            const currentUrl = window.location.href;
            if (currentUrl !== lastUrl) {
                lastUrl = currentUrl;
                setTimeout(() => debouncedAddCustomButton(), 50);
            }
        };

        if (window.MutationObserver) {
            const observer = new MutationObserver(() => {
                checkUrlChange();
            });

            observer.observe(document.body, {
                childList: true,
                subtree: true
            });

            intervalId = setInterval(checkUrlChange, TIMINGS.URL_CHECK_INTERVAL * 10);
        } else {
            intervalId = setInterval(checkUrlChange, TIMINGS.URL_CHECK_INTERVAL);
        }

        return () => {
            if (intervalId) {
                clearInterval(intervalId);
            }
        };
    }

    function initialize() {
        // Use debounced version for all initialization calls
        debouncedAddCustomButton();

        initializeUrlMonitoring();

        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', () => {
                setTimeout(() => debouncedAddCustomButton(), 100);
            });
        } else if (document.readyState === 'interactive') {
            setTimeout(() => debouncedAddCustomButton(), 100);
        }

        window.addEventListener('load', () => {
            setTimeout(() => debouncedAddCustomButton(), 200);
        });
    }

    // Start the extension
    initialize();

})();
