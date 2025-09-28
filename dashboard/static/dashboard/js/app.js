/**
 * Dashboard App JavaScript
 */

// Configuration
const API_BASE = '/api/';
const REFRESH_INTERVAL = 60000; // 60 seconds

// Dashboard App Object
const DashboardApp = {

    init: function() {
        console.log('Dashboard initialized');
        this.bindEvents();
        this.initTooltips();
        this.startAutoRefresh();
    },

    bindEvents: function() {
        // Search form
        $('#searchForm').on('submit', this.handleSearch);

        // Refresh button
        $('.btn-refresh').on('click', this.refreshData);

        // Task checkboxes
        $('.task-checkbox').on('change', this.updateTask);
    },

    initTooltips: function() {
        $('[data-bs-toggle="tooltip"]').tooltip();
    },

    startAutoRefresh: function() {
        setInterval(() => {
            this.loadStatistics();
        }, REFRESH_INTERVAL);
    },

    handleSearch: function(e) {
        e.preventDefault();
        const query = $('#searchInput').val();
        if (query.length < 3) {
            DashboardApp.showNotification('لطفا حداقل 3 کاراکتر وارد کنید', 'warning');
            return;
        }
        window.location.href = `/search/?q=${encodeURIComponent(query)}`;
    },

    refreshData: function() {
        DashboardApp.showNotification('در حال بروزرسانی...', 'info');
        location.reload();
    },

    updateTask: function() {
        const taskId = $(this).data('task-id');
        const isChecked = $(this).is(':checked');

        // Update task status via API
        $.ajax({
            url: `${API_BASE}tasks/${taskId}/`,
            method: 'PATCH',
            data: {
                status: isChecked ? 'Completed' : 'In Progress'
            },
            success: function() {
                DashboardApp.showNotification('وضعیت وظیفه بروز شد', 'success');
            },
            error: function() {
                DashboardApp.showNotification('خطا در بروزرسانی وظیفه', 'error');
            }
        });
    },

    loadStatistics: function() {
        $.ajax({
            url: `${API_BASE}statistics/`,
            method: 'GET',
            success: function(response) {
                if (response.success) {
                    DashboardApp.updateStatistics(response.data);
                }
            }
        });
    },

    updateStatistics: function(data) {
        // Animate number changes
        $('.stat-value').each(function() {
            const $el = $(this);
            const key = $el.data('stat-key');
            if (data[key]) {
                $el.animateNumber({
                    number: data[key]
                });
            }
        });
    },

    showNotification: function(message, type = 'info') {
        const alertClass = {
            'info': 'alert-info',
            'success': 'alert-success',
            'warning': 'alert-warning',
            'error': 'alert-danger'
        }[type];

        const alert = `
            <div class="alert ${alertClass} alert-dismissible fade show" role="alert">
                ${message}
                <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
            </div>
        `;

        $('.alerts-container').append(alert);

        // Auto remove after 5 seconds
        setTimeout(() => {
            $('.alert').fadeOut();
        }, 5000);
    }
};

// Initialize on DOM ready
$(document).ready(function() {
    DashboardApp.init();
});

// jQuery Number Animation Plugin (simplified)
$.fn.animateNumber = function(options) {
    return this.each(function() {
        const $el = $(this);
        const start = parseInt($el.text()) || 0;
        const end = options.number;
        const duration = 1000;
        const range = end - start;
        const stepTime = Math.abs(Math.floor(duration / range));

        let current = start;
        const timer = setInterval(() => {
            current += 1;
            $el.text(current);

            if (current >= end) {
                clearInterval(timer);
            }
        }, stepTime);
    });
};