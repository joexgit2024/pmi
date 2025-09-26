/* Main JavaScript for PMI PMDoS Management System */

// Global variables
let loadingOverlay = null;

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

function initializeApp() {
    // Get references to common elements
    loadingOverlay = document.getElementById('loadingOverlay');
    
    // Initialize Bootstrap tooltips and popovers
    initializeBootstrapComponents();
    
    // Set up HTMX event handlers
    setupHTMXHandlers();
    
    // Initialize page-specific functionality
    initializePageSpecific();
    
    // Add fade-in animation to main content
    document.querySelector('main').classList.add('fade-in');
}

function initializeBootstrapComponents() {
    // Initialize tooltips
    var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });
    
    // Initialize popovers
    var popoverTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="popover"]'));
    var popoverList = popoverTriggerList.map(function (popoverTriggerEl) {
        return new bootstrap.Popover(popoverTriggerEl);
    });
}

function setupHTMXHandlers() {
    // Show loading on HTMX requests
    document.addEventListener('htmx:beforeRequest', function(evt) {
        showLoading();
    });
    
    // Hide loading when HTMX request completes
    document.addEventListener('htmx:afterRequest', function(evt) {
        hideLoading();
    });
    
    // Handle HTMX errors
    document.addEventListener('htmx:responseError', function(evt) {
        hideLoading();
        showAlert('Request failed. Please try again.', 'error');
    });
    
    // Handle successful responses
    document.addEventListener('htmx:afterSwap', function(evt) {
        // Re-initialize components in swapped content
        initializeBootstrapComponents();
    });
}

function initializePageSpecific() {
    const path = window.location.pathname;
    
    if (path === '/' || path === '/dashboard') {
        initializeDashboard();
    } else if (path.includes('/upload')) {
        initializeUpload();
    } else if (path.includes('/matching')) {
        initializeMatching();
    } else if (path.includes('/email')) {
        initializeEmail();
    }
}

function initializeDashboard() {
    // Dashboard-specific initialization
    console.log('Dashboard initialized');
    
    // Auto-refresh statistics every 5 minutes
    setInterval(refreshDashboardStats, 300000);
}

function initializeUpload() {
    // Upload-specific initialization
    console.log('Upload page initialized');
}

function initializeMatching() {
    // Matching-specific initialization
    console.log('Matching page initialized');
}

function initializeEmail() {
    // Email-specific initialization
    console.log('Email page initialized');
}

// Utility Functions

function showLoading(message = 'Processing...') {
    if (loadingOverlay) {
        const loadingText = loadingOverlay.querySelector('div:last-child');
        if (loadingText) {
            loadingText.textContent = message;
        }
        loadingOverlay.classList.remove('d-none');
    }
}

function hideLoading() {
    if (loadingOverlay) {
        loadingOverlay.classList.add('d-none');
    }
}

function showAlert(message, type = 'info', duration = 5000) {
    // Create alert element
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type === 'error' ? 'danger' : type} alert-dismissible fade show`;
    alertDiv.setAttribute('role', 'alert');
    
    alertDiv.innerHTML = `
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
    `;
    
    // Insert alert at the top of the main content
    const container = document.querySelector('.container-fluid');
    const firstChild = container.firstElementChild;
    container.insertBefore(alertDiv, firstChild);
    
    // Auto-remove after duration
    if (duration > 0) {
        setTimeout(() => {
            if (alertDiv && alertDiv.parentNode) {
                alertDiv.remove();
            }
        }, duration);
    }
    
    // Scroll to top to show alert
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function refreshSystem() {
    showLoading('Refreshing system data...');
    
    fetch('/api/system/refresh', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        }
    })
    .then(response => response.json())
    .then(data => {
        hideLoading();
        if (data.success) {
            showAlert('System data refreshed successfully!', 'success');
            setTimeout(() => location.reload(), 1500);
        } else {
            showAlert('Error refreshing system: ' + data.message, 'error');
        }
    })
    .catch(error => {
        hideLoading();
        showAlert('Error refreshing system: ' + error.message, 'error');
    });
}

function refreshDashboardStats() {
    // Refresh dashboard statistics without showing loading
    fetch('/api/dashboard/stats')
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            updateDashboardStats(data.stats);
        }
    })
    .catch(error => {
        console.error('Error refreshing dashboard stats:', error);
    });
}

function updateDashboardStats(stats) {
    // Update statistics cards
    const statElements = {
        'total_registrations': document.querySelector('[data-stat="total_registrations"]'),
        'successful_matches': document.querySelector('[data-stat="successful_matches"]'),
        'emails_sent': document.querySelector('[data-stat="emails_sent"]'),
        'pending_emails': document.querySelector('[data-stat="pending_emails"]')
    };
    
    for (const [key, element] of Object.entries(statElements)) {
        if (element && stats[key] !== undefined) {
            animateCounter(element, parseInt(element.textContent) || 0, stats[key]);
        }
    }
}

function animateCounter(element, startValue, endValue, duration = 1000) {
    const startTime = performance.now();
    const difference = endValue - startValue;
    
    function updateCounter(currentTime) {
        const elapsed = currentTime - startTime;
        const progress = Math.min(elapsed / duration, 1);
        
        // Use easing function for smooth animation
        const easeOutCubic = 1 - Math.pow(1 - progress, 3);
        const currentValue = Math.round(startValue + (difference * easeOutCubic));
        
        element.textContent = currentValue;
        
        if (progress < 1) {
            requestAnimationFrame(updateCounter);
        }
    }
    
    requestAnimationFrame(updateCounter);
}

// File handling utilities
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function getFileIcon(filename) {
    const extension = filename.split('.').pop().toLowerCase();
    
    switch (extension) {
        case 'xlsx':
        case 'xls':
            return 'bi-file-earmark-excel';
        case 'csv':
            return 'bi-file-earmark-text';
        case 'pdf':
            return 'bi-file-earmark-pdf';
        default:
            return 'bi-file-earmark';
    }
}

// API request helper
function apiRequest(url, options = {}) {
    const defaultOptions = {
        headers: {
            'Content-Type': 'application/json',
        }
    };
    
    const mergedOptions = { ...defaultOptions, ...options };
    
    return fetch(url, mergedOptions)
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .catch(error => {
            console.error('API request failed:', error);
            throw error;
        });
}

// Form validation helpers
function validateEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

function validateRequired(value) {
    return value !== null && value !== undefined && value.toString().trim() !== '';
}

function validateForm(formElement) {
    const requiredFields = formElement.querySelectorAll('[required]');
    let isValid = true;
    
    requiredFields.forEach(field => {
        if (!validateRequired(field.value)) {
            field.classList.add('is-invalid');
            isValid = false;
        } else {
            field.classList.remove('is-invalid');
            field.classList.add('is-valid');
        }
        
        // Special validation for email fields
        if (field.type === 'email' && field.value && !validateEmail(field.value)) {
            field.classList.add('is-invalid');
            isValid = false;
        }
    });
    
    return isValid;
}

// Local storage helpers
function saveToLocalStorage(key, data) {
    try {
        localStorage.setItem(key, JSON.stringify(data));
        return true;
    } catch (error) {
        console.error('Error saving to localStorage:', error);
        return false;
    }
}

function loadFromLocalStorage(key) {
    try {
        const data = localStorage.getItem(key);
        return data ? JSON.parse(data) : null;
    } catch (error) {
        console.error('Error loading from localStorage:', error);
        return null;
    }
}

// Theme management
function toggleTheme() {
    const currentTheme = document.documentElement.getAttribute('data-theme') || 'light';
    const newTheme = currentTheme === 'light' ? 'dark' : 'light';
    
    document.documentElement.setAttribute('data-theme', newTheme);
    saveToLocalStorage('theme', newTheme);
}

function loadSavedTheme() {
    const savedTheme = loadFromLocalStorage('theme');
    if (savedTheme) {
        document.documentElement.setAttribute('data-theme', savedTheme);
    }
}

// Initialize theme on load
document.addEventListener('DOMContentLoaded', loadSavedTheme);

// Keyboard shortcuts
document.addEventListener('keydown', function(event) {
    // Ctrl+R or F5 - Refresh
    if ((event.ctrlKey && event.key === 'r') || event.key === 'F5') {
        event.preventDefault();
        refreshSystem();
    }
    
    // Escape - Close modals/loading
    if (event.key === 'Escape') {
        hideLoading();
        // Close any open modals
        const modals = document.querySelectorAll('.modal.show');
        modals.forEach(modal => {
            const modalInstance = bootstrap.Modal.getInstance(modal);
            if (modalInstance) {
                modalInstance.hide();
            }
        });
    }
});

// Export functions for global use
window.PMIApp = {
    showLoading,
    hideLoading,
    showAlert,
    refreshSystem,
    apiRequest,
    validateForm,
    formatFileSize,
    getFileIcon,
    toggleTheme
};