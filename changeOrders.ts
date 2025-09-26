/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/

// --- TYPE DEFINITIONS ---
type ApproverStatus = 'Pending' | 'Approved' | 'Rejected';

interface Approver {
    name: string;
    title: string;
    company: string;
    status: ApproverStatus;
    approved: boolean;
}

// NOTE: This interface is duplicated in index.tsx.
// It's kept here for TypeScript type safety within this module.
interface ChangeOrder {
    id: number;
    name: string;
    date: string;
    amount: number;
    status: 'Pending' | 'Approved' | 'Rejected';
    description: string;
    scheduleImpact?: number;
    comments?: string;
    approvers?: Approver[];
}

interface ChangeOrdersTabProps {
    changeOrders: ChangeOrder[];
    onAdd: (data: Omit<ChangeOrder, 'id'>) => void;
    onUpdate: (data: ChangeOrder) => void;
    onDelete: (id: number) => void;
}


/**
 * A functional component that renders a complete UI for managing Change Orders.
 * This includes a form for adding/editing and a table for displaying items.
 *
 * @returns {HTMLDivElement} The container element for the Change Orders tab.
 */
const ChangeOrdersTab = (props: ChangeOrdersTabProps): HTMLDivElement => {
    // --- STATE MANAGEMENT ---
    let editingId: number | null = null;
    let currentApprovers: Approver[] = [];

    // --- UI SETUP ---
    const container = document.createElement('div');
    container.className = 'change-orders-container';

    container.innerHTML = `
        <div class="change-orders-content">
            <h2>Change Orders</h2>
            <p>Track and manage all change orders for this project.</p>
            
            <form id="change-order-form" class="add-change-order-form">
                <h3 id="change-order-form-title">Add New Change Order</h3>
                <input type="hidden" id="change-order-id">
                <div class="form-group form-group-full">
                    <label for="change-order-name">Name / Title</label>
                    <input type="text" id="change-order-name" placeholder="e.g., Additional Sensor Installation" required>
                </div>
                 <div class="form-group form-group-full">
                    <label for="change-order-description">Description</label>
                    <textarea id="change-order-description" rows="3" placeholder="Describe the reason and scope of the change..."></textarea>
                </div>
                <div class="form-group">
                    <label for="change-order-date">Date</label>
                    <input type="date" id="change-order-date" required>
                </div>
                <div class="form-group">
                    <label for="change-order-amount">Amount</label>
                    <input type="number" id="change-order-amount" step="0.01" placeholder="Enter amount" required>
                </div>
                <div class="form-group">
                    <label for="change-order-schedule-impact">Impact to Schedule (days)</label>
                    <input type="number" id="change-order-schedule-impact" step="1" placeholder="e.g., 5">
                </div>
                <div class="form-group form-group-full">
                    <label for="change-order-comments">Comments</label>
                    <textarea id="change-order-comments" rows="3" placeholder="Add any relevant comments..."></textarea>
                </div>
                
                <fieldset class="approvers-fieldset">
                    <legend>Approvers</legend>
                    <div class="approver-inputs">
                        <div class="form-group">
                            <label for="approver-name">Approver Name</label>
                            <input type="text" id="approver-name" placeholder="John Doe">
                        </div>
                        <div class="form-group">
                            <label for="approver-title">Title</label>
                            <input type="text" id="approver-title" placeholder="Project Manager">
                        </div>
                        <div class="form-group">
                            <label for="approver-company">Company</label>
                            <input type="text" id="approver-company" placeholder="Contractor Inc.">
                        </div>
                        <button type="button" id="add-approver-btn" class="secondary-btn">Add Approver</button>
                    </div>
                    <div id="approvers-list" class="approvers-list">
                        <!-- Approvers will be listed here -->
                    </div>
                </fieldset>

                <div class="form-actions">
                    <button type="submit" id="change-order-submit-btn" class="save-btn">Add Change Order</button>
                    <button type="button" id="change-order-cancel-btn" class="secondary-btn" style="display: none;">Cancel Edit</button>
                </div>
            </form>

            <div class="change-orders-list-container">
                <h3>Existing Change Orders</h3>
                <div class="table-wrapper">
                    <table id="change-orders-table" class="data-table">
                        <thead>
                            <tr>
                                <th class="col-toggle"></th>
                                <th class="col-name">Name / Title</th>
                                <th class="col-date">Date</th>
                                <th class="col-amount">Amount</th>
                                <th class="col-impact">Impact (days)</th>
                                <th class="col-status">Status</th>
                                <th class="col-actions">Actions</th>
                            </tr>
                        </thead>
                        <tbody id="change-orders-tbody">
                            <!-- Rows will be injected here -->
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    `;

    // --- DOM ELEMENT REFERENCES ---
    const form = container.querySelector('#change-order-form') as HTMLFormElement;
    const formTitle = container.querySelector('#change-order-form-title') as HTMLHeadingElement;
    const idInput = container.querySelector('#change-order-id') as HTMLInputElement;
    const nameInput = container.querySelector('#change-order-name') as HTMLInputElement;
    const descriptionTextarea = container.querySelector('#change-order-description') as HTMLTextAreaElement;
    const dateInput = container.querySelector('#change-order-date') as HTMLInputElement;
    const amountInput = container.querySelector('#change-order-amount') as HTMLInputElement;
    const scheduleImpactInput = container.querySelector('#change-order-schedule-impact') as HTMLInputElement;
    const commentsTextarea = container.querySelector('#change-order-comments') as HTMLTextAreaElement;
    const submitBtn = container.querySelector('#change-order-submit-btn') as HTMLButtonElement;
    const cancelBtn = container.querySelector('#change-order-cancel-btn') as HTMLButtonElement;
    const tableBody = container.querySelector('#change-orders-tbody') as HTMLTableSectionElement;
    
    // Approver form elements
    const approverNameInput = container.querySelector('#approver-name') as HTMLInputElement;
    const approverTitleInput = container.querySelector('#approver-title') as HTMLInputElement;
    const approverCompanyInput = container.querySelector('#approver-company') as HTMLInputElement;
    const addApproverBtn = container.querySelector('#add-approver-btn') as HTMLButtonElement;
    const approversListDiv = container.querySelector('#approvers-list') as HTMLDivElement;

    // --- HELPER FUNCTIONS ---

    /**
     * Calculates the overall status of a change order based on its approvers.
     * @param approvers The list of approvers.
     * @returns The calculated status.
     */
    const calculateOverallStatus = (approvers: Approver[]): 'Pending' | 'Approved' | 'Rejected' => {
        if (!approvers || approvers.length === 0) {
            return 'Pending'; // Default to pending if no approvers
        }
        if (approvers.some(a => a.status === 'Rejected')) {
            return 'Rejected';
        }
        if (approvers.every(a => a.status === 'Approved')) {
            return 'Approved';
        }
        return 'Pending';
    };

    /** Renders the list of current approvers in the form. */
    const renderApproversList = () => {
        approversListDiv.innerHTML = '';
        if (currentApprovers.length === 0) {
            approversListDiv.innerHTML = '<p class="no-approvers-msg">No approvers added for this change order.</p>';
            return;
        }

        const statusOptions: ApproverStatus[] = ['Pending', 'Approved', 'Rejected'];

        currentApprovers.forEach((approver, index) => {
            const approverEl = document.createElement('div');
            approverEl.className = 'approver-item';
            approverEl.dataset.index = index.toString();

            const selectOptions = statusOptions.map(s =>
                `<option value="${s}" ${approver.status === s ? 'selected' : ''}>${s}</option>`
            ).join('');
            
            approverEl.innerHTML = `
                <div class="approver-info">
                    <strong>${approver.name}</strong> (${approver.title}, ${approver.company})
                </div>
                <div class="approver-controls">
                    <select class="approver-status-select">
                        ${selectOptions}
                    </select>
                    <div class="approver-checkbox-container">
                        <input type="checkbox" id="approver-check-${index}" class="approver-checkbox" ${approver.approved ? 'checked' : ''}>
                        <label for="approver-check-${index}">Approved</label>
                    </div>
                    <button type="button" class="delete-btn-small" data-action="delete-approver" aria-label="Remove ${approver.name}">Remove</button>
                </div>
            `;
            approversListDiv.appendChild(approverEl);
        });
    };
    
    /** Resets the form to its initial state for adding a new item. */
    const resetForm = () => {
        form.reset();
        editingId = null;
        idInput.value = '';
        formTitle.textContent = 'Add New Change Order';
        submitBtn.textContent = 'Add Change Order';
        cancelBtn.style.display = 'none';
        scheduleImpactInput.value = '';
        commentsTextarea.value = '';
        currentApprovers = [];
        renderApproversList();
        nameInput.focus();
    };
    
    /** Renders the table with the current change orders from props. */
    const renderTable = () => {
        tableBody.innerHTML = '';
        if (props.changeOrders.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="7">No change orders have been added yet.</td></tr>';
            return;
        }

        props.changeOrders.forEach(co => {
            const formattedAmount = new Intl.NumberFormat(undefined, { style: 'currency', currency: 'USD' }).format(co.amount);
            const scheduleImpact = co.scheduleImpact !== undefined ? `${co.scheduleImpact}` : 'N/A';
            
            const approversHtml = co.approvers && co.approvers.length > 0
                ? `<ul class="details-approvers-list">
                    ${co.approvers.map(a => `
                        <li>
                           ${a.name} (${a.title}, ${a.company}) - <strong>${a.status}</strong> ${a.approved ? '<span class="approved-icon" title="Approved">&#10004;</span>' : ''}
                        </li>
                    `).join('')}
                   </ul>`
                : '<p>N/A</p>';

            // Main Row
            const mainRow = tableBody.insertRow();
            mainRow.className = 'main-row';
            mainRow.dataset.id = co.id.toString();
            mainRow.innerHTML = `
                <td class="col-toggle">
                    <button class="toggle-details-btn" data-action="toggle-details" aria-expanded="false" aria-label="Show details for ${co.name}">
                        <svg viewBox="0 0 24 24" width="20" height="20"><path d="M7 10l5 5 5-5z"></path></svg>
                    </button>
                </td>
                <td class="col-name">${co.name}</td>
                <td class="col-date">${co.date}</td>
                <td class="col-amount">${formattedAmount}</td>
                <td class="col-impact">${scheduleImpact}</td>
                <td class="col-status">
                    <span class="status-badge status-${co.status.toLowerCase()}">${co.status}</span>
                </td>
                <td class="col-actions">
                    <button class="edit-btn" data-action="edit">Edit</button>
                    <button class="delete-btn-list" data-action="delete">Delete</button>
                </td>
            `;

            // Details Row (initially hidden)
            const detailsRow = tableBody.insertRow();
            detailsRow.className = 'details-row';
            detailsRow.dataset.detailsForId = co.id.toString();
            detailsRow.style.display = 'none';
            detailsRow.innerHTML = `
                <td colspan="7">
                    <div class="details-content">
                        <dl>
                            <dt>Description</dt>
                            <dd>${co.description.replace(/</g, "&lt;").replace(/>/g, "&gt;") || 'N/A'}</dd>
                            
                            <dt>Comments</dt>
                            <dd>${co.comments?.replace(/</g, "&lt;").replace(/>/g, "&gt;") || 'N/A'}</dd>

                            <dt>Approvers</dt>
                            <dd>${approversHtml}</dd>
                        </dl>
                    </div>
                </td>
            `;
        });
    };


    /** Handles form submission for both adding and updating by calling parent handlers. */
    const handleSubmit = (e: Event) => {
        e.preventDefault();
        const scheduleImpactValue = scheduleImpactInput.value ? parseInt(scheduleImpactInput.value, 10) : undefined;
        
        const changeOrderData = {
            name: nameInput.value.trim(),
            description: descriptionTextarea.value.trim(),
            date: dateInput.value,
            amount: parseFloat(amountInput.value),
            status: calculateOverallStatus(currentApprovers),
            scheduleImpact: scheduleImpactValue,
            comments: commentsTextarea.value.trim(),
            approvers: [...currentApprovers],
        };

        if (editingId !== null) {
            props.onUpdate({ id: editingId, ...changeOrderData });
        } else {
            props.onAdd(changeOrderData);
        }
        
        resetForm();
    };

    /** Handles clicks within the table body for edit/delete actions. */
    const handleTableClick = (e: Event) => {
        const target = e.target as HTMLElement;
        const button = target.closest<HTMLButtonElement>('button');
        if (!button) return;

        const action = button.dataset.action;
        const row = target.closest<HTMLTableRowElement>('tr.main-row');
        if (!action || !row || !row.dataset.id) return;
        
        const id = parseInt(row.dataset.id, 10);

        if (action === 'delete') {
            if (confirm('Are you sure you want to delete this change order?')) {
                props.onDelete(id);
                if (editingId === id) {
                    resetForm();
                }
            }
        } else if (action === 'edit') {
            const coToEdit = props.changeOrders.find(co => co.id === id);
            if (coToEdit) {
                formTitle.textContent = 'Edit Change Order';
                submitBtn.textContent = 'Update Change Order';
                cancelBtn.style.display = 'inline-block';

                editingId = coToEdit.id;
                idInput.value = coToEdit.id.toString();
                nameInput.value = coToEdit.name;
                descriptionTextarea.value = coToEdit.description;
                dateInput.value = coToEdit.date;
                amountInput.value = coToEdit.amount.toString();
                scheduleImpactInput.value = coToEdit.scheduleImpact?.toString() || '';
                commentsTextarea.value = coToEdit.comments || '';
                
                const approvers = coToEdit.approvers ? JSON.parse(JSON.stringify(coToEdit.approvers)) : [];
                currentApprovers = approvers.map((a: Partial<Approver> & { approved: boolean }) => ({
                    name: a.name || '',
                    title: a.title || '',
                    company: a.company || '',
                    status: a.status || (a.approved ? 'Approved' : 'Pending'),
                    approved: !!a.approved 
                }));
                renderApproversList();
                
                nameInput.focus();
                form.scrollIntoView({ behavior: 'smooth' });
            }
        } else if (action === 'toggle-details') {
            const detailsRow = tableBody.querySelector(`tr[data-details-for-id="${id}"]`) as HTMLTableRowElement | null;
            if (detailsRow) {
                const isExpanded = detailsRow.style.display !== 'none';
                detailsRow.style.display = isExpanded ? 'none' : 'table-row';
                row.classList.toggle('is-expanded', !isExpanded);
                button.setAttribute('aria-expanded', String(!isExpanded));
                button.classList.toggle('expanded', !isExpanded);
            }
        }
    };

    // --- EVENT LISTENERS ---
    form.addEventListener('submit', handleSubmit);
    tableBody.addEventListener('click', handleTableClick);
    cancelBtn.addEventListener('click', resetForm);

    addApproverBtn.addEventListener('click', () => {
        const name = approverNameInput.value.trim();
        const title = approverTitleInput.value.trim();
        const company = approverCompanyInput.value.trim();

        if (name && title && company) {
            currentApprovers.push({ name, title, company, status: 'Pending', approved: false });
            renderApproversList();
            approverNameInput.value = '';
            approverTitleInput.value = '';
            approverCompanyInput.value = '';
            approverNameInput.focus();
        } else {
            alert('Please fill out all three approver fields (Name, Title, Company).');
        }
    });

    approversListDiv.addEventListener('click', (e) => {
        const target = e.target as HTMLElement;
        if (target.dataset.action === 'delete-approver') {
            const itemEl = target.closest('.approver-item') as HTMLElement | null;
            if(itemEl) {
                const index = parseInt(itemEl.dataset.index!, 10);
                if (!isNaN(index)) {
                    currentApprovers.splice(index, 1);
                    renderApproversList();
                }
            }
        }
    });

    approversListDiv.addEventListener('change', (e) => {
        const target = e.target as HTMLInputElement | HTMLSelectElement;
        const itemEl = target.closest('.approver-item') as HTMLElement | null;
        if (!itemEl) return;

        const index = parseInt(itemEl.dataset.index!, 10);
        if (isNaN(index) || !currentApprovers[index]) return;

        const approver = currentApprovers[index];
        const checkbox = itemEl.querySelector('.approver-checkbox') as HTMLInputElement;
        const statusSelect = itemEl.querySelector('.approver-status-select') as HTMLSelectElement;

        if (target.matches('.approver-checkbox')) {
            approver.approved = checkbox.checked;
            approver.status = checkbox.checked ? 'Approved' : 'Pending';
            statusSelect.value = approver.status;
        }

        if (target.matches('.approver-status-select')) {
            approver.status = statusSelect.value as ApproverStatus;
            approver.approved = approver.status === 'Approved';
            checkbox.checked = approver.approved;
        }
    });

    // --- INITIAL RENDER ---
    renderTable();
    renderApproversList(); // Render empty list initially
    
    // --- STYLES ---
    const style = document.createElement('style');
    style.textContent = `
        .change-orders-content {
            max-width: 1000px;
            margin: 0 auto;
        }
        .add-change-order-form {
            background-color: var(--primary-bg);
            padding: 1.5rem;
            border-radius: 8px;
            margin-bottom: 2rem;
            display: flex;
            flex-wrap: wrap;
            gap: 1rem;
        }
        .add-change-order-form h3 {
            width: 100%;
            margin-top: 0;
            border: none;
            padding: 0;
        }
        .add-change-order-form .form-group {
            flex: 1 1 200px;
        }
        .add-change-order-form .form-actions {
            flex-basis: 100%;
            display: flex;
            justify-content: flex-start;
            gap: 1rem;
            align-items: center;
        }
        .approvers-fieldset {
            flex-basis: 100%;
            border: 1px solid var(--border-color);
            padding: 1rem;
            border-radius: 6px;
            margin-top: 0.5rem;
        }
        .approvers-fieldset legend {
            font-weight: 500;
            padding: 0 0.5rem;
            color: #444;
        }
        .approver-inputs {
            display: flex;
            flex-wrap: wrap;
            gap: 1rem;
            align-items: flex-end;
            margin-bottom: 1rem;
        }
        .approver-inputs .form-group {
            flex: 1 1 150px;
            margin-bottom: 0;
        }
        .approver-inputs button {
            height: 48px; /* Match input height */
            flex-shrink: 0;
        }
        .approvers-list {
            display: flex;
            flex-direction: column;
            gap: 0.5rem;
            max-height: 150px;
            overflow-y: auto;
        }
        .approver-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            background-color: var(--secondary-bg);
            padding: 0.5rem 0.75rem;
            border-radius: 4px;
            border: 1px solid var(--border-color);
            font-size: 0.9rem;
            gap: 1rem;
        }
        .approver-info {
            flex-grow: 1;
            word-break: break-word;
        }
        .approver-controls {
            display: flex;
            align-items: center;
            gap: 1rem;
            flex-shrink: 0;
        }
        .approver-status-select {
            padding: 0.25rem 0.5rem;
            border-radius: 4px;
            border: 1px solid var(--border-color);
            background-color: var(--secondary-bg);
        }
        .approver-checkbox-container {
            display: flex;
            align-items: center;
            gap: 0.25rem;
        }
        .approver-checkbox-container label {
            font-size: 0.8rem;
            color: #555;
            margin: 0;
            cursor: pointer;
            white-space: nowrap;
        }
        .approver-checkbox {
            width: 16px;
            height: 16px;
            flex-shrink: 0;
            cursor: pointer;
        }
        .no-approvers-msg {
            font-style: italic;
            color: var(--placeholder-text);
            padding: 0.5rem 0;
        }
        .approved-icon {
            color: var(--status-on-track);
            font-weight: bold;
            margin-left: 4px;
        }
        
        /* New Table Styles */
        #change-orders-table {
            border-top: 1px solid var(--border-color);
        }
        #change-orders-table .main-row td {
            vertical-align: middle;
        }
        #change-orders-table .main-row.is-expanded {
            border-bottom-color: transparent; /* Visually connect main row to details row */
        }
        #change-orders-table .col-toggle { width: 40px; text-align: center; }
        #change-orders-table .col-date { width: 110px; }
        #change-orders-table .col-amount { width: 130px; }
        #change-orders-table .col-impact { width: 90px; text-align: center; }
        #change-orders-table .col-status { width: 120px; }
        #change-orders-table .col-actions { width: 150px; text-align: right; }

        .toggle-details-btn {
            background: transparent;
            border: 1px solid var(--border-color);
            border-radius: 50%;
            width: 28px;
            height: 28px;
            cursor: pointer;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 0;
        }
        .toggle-details-btn svg {
            fill: #555;
            transition: transform 0.2s ease-in-out;
        }
        .toggle-details-btn.expanded svg {
            transform: rotate(180deg);
        }

        .status-badge {
            display: inline-block;
            padding: 0.25rem 0.75rem;
            border-radius: 12px;
            font-weight: 500;
            font-size: 0.8rem;
            color: #fff;
        }
        .status-badge.status-pending { background-color: var(--status-at-risk); }
        .status-badge.status-approved { background-color: var(--status-on-track); }
        .status-badge.status-rejected { background-color: var(--status-delayed); }

        .details-row td {
            padding: 0 !important;
            border-bottom: 1px solid var(--border-color);
        }
        .details-content {
            background-color: #f8f9fa;
            padding: 1.5rem 1.5rem 1.5rem 50px;
        }
        .details-content dl {
            margin: 0;
        }
        .details-content dt {
            font-weight: 600;
            color: var(--primary-accent);
            margin-top: 1rem;
            font-size: 0.9rem;
        }
        .details-content dd {
            margin-left: 0;
            margin-top: 0.25rem;
            white-space: pre-wrap;
            word-break: break-word;
        }
        .details-approvers-list {
            list-style: none;
            padding: 0;
        }
        .details-approvers-list li {
            padding: 0.25rem 0;
        }
    `;
    container.appendChild(style);

    return container;
};

export default ChangeOrdersTab;
