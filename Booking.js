// This script will be included in the main page, so we need to wrap our logic
// in a function that can be called when the booking page is loaded.

function initBookingPage() {
    const resourceSelector = document.getElementById('resource-selector');
    const calendarContainer = document.getElementById('calendar-container');
    const calendarGrid = document.getElementById('calendar-grid');
    const monthYearEl = document.getElementById('calendar-month-year');
    const prevMonthBtn = document.getElementById('prev-month-btn');
    const nextMonthBtn = document.getElementById('next-month-btn');
    const modal = document.getElementById('booking-modal');
    const closeModalBtn = modal.querySelector('.close-btn');
    const bookingForm = document.getElementById('booking-form');
    const existingBookingsList = document.getElementById('existing-bookings-list');

    let currentYear = new Date().getFullYear();
    let currentMonth = new Date().getMonth();
    let selectedResource = null;
    let bookings = [];

    function loadResources() {
        google.script.run.withSuccessHandler(response => {
            if (response.ok) {
                resourceSelector.innerHTML = '<option value="">-- Vennligst velg --</option>';
                response.resources.forEach(resource => {
                    const option = document.createElement('option');
                    option.value = resource.id;
                    option.textContent = resource.name;
                    option.dataset.resource = JSON.stringify(resource);
                    resourceSelector.appendChild(option);
                });
            } else {
                alert('Feil ved lasting av ressurser: ' + response.message);
            }
        }).listResources(); // CORRECTED: Was getCommonResources
    }

    function fetchAndRenderCalendar() {
        if (!selectedResource) return;
        google.script.run.withSuccessHandler(response => {
            if (response.ok) {
                bookings = response.bookings;
                renderCalendar();
            } else {
                alert('Feil ved lasting av bookinger: ' + response.message);
            }
        }).getBookings(selectedResource.id, currentYear, currentMonth);
    }

    function renderCalendar() {
        calendarGrid.innerHTML = '';
        monthYearEl.textContent = `${new Date(currentYear, currentMonth).toLocaleString('no-NO', { month: 'long', year: 'numeric' })}`;

        const firstDayOfMonth = new Date(currentYear, currentMonth, 1).getDay();
        const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();

        // Add day headers
        ['Søn', 'Man', 'Tir', 'Ons', 'Tor', 'Fre', 'Lør'].forEach(day => {
            const dayEl = document.createElement('div');
            dayEl.classList.add('calendar-day', 'header');
            dayEl.textContent = day;
            calendarGrid.appendChild(dayEl);
        });

        // Add empty cells for days before the 1st
        for (let i = 0; i < (firstDayOfMonth === 0 ? 6 : firstDayOfMonth - 1); i++) {
            const emptyCell = document.createElement('div');
            emptyCell.classList.add('calendar-day', 'other-month');
            calendarGrid.appendChild(emptyCell);
        }

        // Add day cells
        for (let day = 1; day <= daysInMonth; day++) {
            const dayEl = document.createElement('div');
            dayEl.classList.add('calendar-day');
            dayEl.textContent = day;
            const cellDateStr = new Date(currentYear, currentMonth, day).toISOString().split('T')[0];
            dayEl.dataset.date = cellDateStr;

            // Check if this day has passed
            const today = new Date();
            today.setHours(0,0,0,0);
            const cellDate = new Date(cellDateStr);
            if(cellDate < today) {
                dayEl.classList.add('other-month'); // Style as past day
            } else {
                dayEl.classList.add('available');
                dayEl.addEventListener('click', () => openBookingModal(cellDateStr));
            }

            // --- NEW: Add visual indicator for booked days ---
            const bookingsForDay = bookings.filter(b => b.startTime.startsWith(cellDateStr));
            if (bookingsForDay.length > 0) {
                dayEl.classList.add('has-bookings');
            }

            calendarGrid.appendChild(dayEl);
        }
    }

    function openBookingModal(date) {
        document.getElementById('modal-resource-name').textContent = selectedResource.name;
        document.getElementById('modal-booking-date').textContent = new Date(date).toLocaleDateString('no-NO');
        bookingForm.dataset.date = date;

        // --- NEW: Display existing bookings for the selected day ---
        existingBookingsList.innerHTML = '';
        const bookingsForDay = bookings.filter(b => b.startTime.startsWith(date));

        if (bookingsForDay.length > 0) {
            bookingsForDay.forEach(b => {
                const li = document.createElement('li');
                const startTime = new Date(b.startTime).toLocaleTimeString('no-NO', { hour: '2-digit', minute: '2-digit' });
                const endTime = new Date(b.endTime).toLocaleTimeString('no-NO', { hour: '2-digit', minute: '2-digit' });
                li.textContent = `Opptatt: ${startTime} - ${endTime}`;
                existingBookingsList.appendChild(li);
            });
        } else {
            existingBookingsList.innerHTML = '<li>Ingen bookinger for denne dagen.</li>';
        }

        modal.style.display = 'block';
    }

    function closeModal() {
        modal.style.display = 'none';
        bookingForm.reset();
        document.getElementById('booking-error').textContent = '';
    }

    // --- Event Listeners ---
    resourceSelector.addEventListener('change', () => {
        const selectedOption = resourceSelector.options[resourceSelector.selectedIndex];
        if (selectedOption.value) {
            selectedResource = JSON.parse(selectedOption.dataset.resource);
            calendarContainer.style.display = 'block';
            fetchAndRenderCalendar();
        } else {
            selectedResource = null;
            calendarContainer.style.display = 'none';
        }
    });

    prevMonthBtn.addEventListener('click', () => {
        currentMonth--;
        if (currentMonth < 0) {
            currentMonth = 11;
            currentYear--;
        }
        fetchAndRenderCalendar();
    });

    nextMonthBtn.addEventListener('click', () => {
        currentMonth++;
        if (currentMonth > 11) {
            currentMonth = 0;
            currentYear++;
        }
        fetchAndRenderCalendar();
    });

    closeModalBtn.addEventListener('click', closeModal);
    window.addEventListener('click', (event) => {
        if (event.target == modal) {
            closeModal();
        }
    });

    bookingForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const date = bookingForm.dataset.date;
        const startTime = document.getElementById('booking-start-time').value;
        const endTime = document.getElementById('booking-end-time').value;

        const bookingDetails = {
            resourceId: selectedResource.id,
            startTime: `${date}T${startTime}:00`,
            endTime: `${date}T${endTime}:00`,
        };

        google.script.run.withSuccessHandler(response => {
            if (response.ok) {
                alert('Booking bekreftet!');
                closeModal();
                fetchAndRenderCalendar(); // Refresh calendar to show new booking
            } else {
                document.getElementById('booking-error').textContent = response.message;
            }
        }).createBooking(bookingDetails);
    });

    // Initial load
    loadResources();
}