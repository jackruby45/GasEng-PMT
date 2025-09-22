

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/
// Add declarations for CDN libraries
declare const jspdf: any;
declare const html2canvas: any;
declare const flatpickr: any;

type ViewMode = 'day' | 'week' | 'month' | 'quarter' | 'year';
type TaskStatus = 'On Track' | 'At Risk' | 'Delayed' | 'Complete';
type ActionItemStatus = 'Not Started' | 'In Progress' | 'Complete';

interface Task {
    id: number;
    name: string;
    description?: string;
    startDate: string;
    endDate: string;
    parentId: number | null;
    percentComplete: number;
    dependencies?: number[];
    resources?: string[];
    notes?: string;
    status: TaskStatus | 'auto';
    customFields?: { [key: string]: string };
}

interface CustomFieldSchema {
    name: string;
    type: 'text'; // For future expansion (e.g., 'number', 'date')
}

// FIX: Renamed Notification to AppNotification to avoid conflict with the browser's built-in Notification API.
interface AppNotification {
    id: string; // e.g., 'overdue-5' or 'due-soon-5'
    message: string;
    taskId: number;
}

interface ProjectInfo {
    name: string;
    description: string;
    deliverables: string;
    location: string;
    doc: string;
    owner: string;
    engineeringManager: string;
    manager: string;
    engineers: string;
    startDate: string;
    endDate: string;
    budget: number | string; // Use string for empty input, number otherwise
    currency: string;
}

interface MeetingMinute {
    id: number;
    meetingDate: string;
    addedDate: string;
    minutes: string;
}

interface ActionItem {
    id: number;
    what: string;
    assignedTo: string[];
    assignedBy: string;
    assignedDate: string;
    dueDate: string;
    status: ActionItemStatus;
}

interface PreviousReport {
    id: number;
    timestamp: string;
    content: string;
}

interface ProjectState {
    projectInfo: ProjectInfo;
    tasks: Task[];
    customFieldsSchema: CustomFieldSchema[];
    baselineTasks: Task[];
    meetingMinutes: MeetingMinute[];
    actionItems: ActionItem[];
    previousReports: PreviousReport[];
    nextMeetingMinuteId: number;
    nextActionItemId: number;
    nextPreviousReportId: number;
    nextTaskId: number;
    expandedTaskIds: number[];
    currentView: ViewMode;
    showBaseline: boolean;
    colorCodeCriteria: string;
    dayWidth: number;
}


// --- CONFIGURATION & STATE ---

// The width of a single day on the timeline in pixels.
// This is now a dynamic variable instead of a constant.
let dayWidth = 50;

// State
let projectInfo: ProjectInfo = {
    name: '',
    description: '',
    deliverables: '',
    location: '',
    doc: '',
    owner: '',
    engineeringManager: '',
    manager: '',
    engineers: '',
    startDate: '',
    endDate: '',
    budget: '',
    currency: 'USD'
};
let tasks: Task[] = [];
let baselineTasks: Task[] = [];
let customFieldsSchema: CustomFieldSchema[] = [];
let meetingMinutes: MeetingMinute[] = [];
let actionItems: ActionItem[] = [];
let previousReports: PreviousReport[] = [];
// FIX: Updated the type of the notifications array to use the new AppNotification interface.
let notifications: AppNotification[] = [];
let nextTaskId = 1;
let nextMeetingMinuteId = 1;
let nextActionItemId = 1;
let nextPreviousReportId = 1;
let editingTaskId: number | null = null;
let subtaskParentId: number | null = null;
let expandedTaskIds: Set<number> = new Set();
let currentView: ViewMode = 'day';
let showBaseline = false;
let filters = { name: '', resource: '', status: '' };
let sortCriteria = { key: 'startDate', order: 'asc' };
let colorCodeCriteria: string = 'status';
const flatpickrInstances: { [key: string]: any } = {};

// DOM Elements
const addTaskForm = document.getElementById('add-task-form') as HTMLFormElement;
const taskNameInput = document.getElementById('task-name-input') as HTMLInputElement;
const startDateInput = document.getElementById('start-date-input') as HTMLInputElement;
const endDateInput = document.getElementById('end-date-input') as HTMLInputElement;
const parentTaskSelect = document.getElementById('parent-task-select') as HTMLSelectElement;
const percentCompleteInput = document.getElementById('percent-complete-input') as HTMLInputElement;
const resourcesInput = document.getElementById('resources-input') as HTMLInputElement;
const addCustomFieldsContainer = document.getElementById('add-custom-fields-container') as HTMLDivElement;
const subtaskFieldsContainer = document.getElementById('subtask-fields-container') as HTMLDivElement;
const taskList = document.getElementById('task-list') as HTMLUListElement;
const ganttChartArea = document.getElementById('gantt-chart') as HTMLDivElement;
const dependencyLinesSvg = document.getElementById('dependency-lines') as unknown as SVGSVGElement;
const viewSwitcher = document.getElementById('view-switcher') as HTMLDivElement;
const filterNameInput = document.getElementById('filter-name') as HTMLInputElement;
const filterResourceInput = document.getElementById('filter-resource') as HTMLInputElement;
const filterStatusSelect = document.getElementById('filter-status') as HTMLSelectElement;
const sortTasksSelect = document.getElementById('sort-tasks') as HTMLSelectElement;
const zoomSlider = document.getElementById('zoom-slider') as HTMLInputElement;
const zoomValue = document.getElementById('zoom-value') as HTMLSpanElement;

// Project Info DOM Elements
const projectInfoForm = document.getElementById('project-info-form') as HTMLDivElement;
const projectNameInput = document.getElementById('project-name') as HTMLInputElement;
const projectDescriptionTextarea = document.getElementById('project-description') as HTMLTextAreaElement;
const projectDeliverablesTextarea = document.getElementById('project-deliverables') as HTMLTextAreaElement;
const projectLocationInput = document.getElementById('project-location') as HTMLInputElement;
const projectDocSelect = document.getElementById('project-doc') as HTMLSelectElement;
const projectOwnerInput = document.getElementById('project-owner') as HTMLInputElement;
const projectEngineeringManagerInput = document.getElementById('project-engineering-manager') as HTMLInputElement;
const projectManagerInput = document.getElementById('project-manager') as HTMLInputElement;
const projectEngineersInput = document.getElementById('project-engineers') as HTMLInputElement;
const projectStartDateInput = document.getElementById('project-start-date') as HTMLInputElement;
const projectEndDateInput = document.getElementById('project-end-date') as HTMLInputElement;
const projectBudgetInput = document.getElementById('project-budget') as HTMLInputElement;
const projectCurrencyInput = document.getElementById('project-currency') as HTMLInputElement;


// Project Control Buttons
const setBaselineBtn = document.getElementById('set-baseline-btn') as HTMLButtonElement;
const showBaselineToggle = document.getElementById('show-baseline-toggle') as HTMLInputElement;

// Color Coding Elements
const colorCodeSelect = document.getElementById('color-code-by') as HTMLSelectElement;
const ganttLegend = document.getElementById('gantt-legend') as HTMLDivElement;

// Data Control Buttons
const newProjectBtn = document.getElementById('new-project-btn') as HTMLButtonElement;
const saveProjectBtn = document.getElementById('save-project-btn') as HTMLButtonElement;
const loadProjectInput = document.getElementById('load-project-input') as HTMLInputElement;


// Export Buttons
const exportCsvBtn = document.getElementById('export-csv-btn');
const exportPdfChartBtn = document.getElementById('export-pdf-chart-btn') as HTMLButtonElement;

// Notification Elements
const notificationBell = document.getElementById('notification-bell') as HTMLButtonElement;
const notificationCount = document.getElementById('notification-count') as HTMLSpanElement;
const notificationPanel = document.getElementById('notification-panel') as HTMLDivElement;
const notificationList = document.getElementById('notification-list') as HTMLUListElement;


// Edit Modal DOM Elements
const editTaskModal = document.getElementById('edit-task-modal') as HTMLDivElement;
const editTaskForm = document.getElementById('edit-task-form') as HTMLFormElement;
const editTaskNameInput = document.getElementById('edit-task-name-input') as HTMLInputElement;
const editTaskDescriptionTextarea = document.getElementById('edit-task-description-textarea') as HTMLTextAreaElement;
const editStartDateInput = document.getElementById('edit-start-date-input') as HTMLInputElement;
const editEndDateInput = document.getElementById('edit-end-date-input') as HTMLInputElement;
const editPercentCompleteInput = document.getElementById('edit-percent-complete-input') as HTMLInputElement;
const editStatusSelect = document.getElementById('edit-status-select') as HTMLSelectElement;
const editResourcesInput = document.getElementById('edit-resources-input') as HTMLInputElement;
const editDependenciesSelect = document.getElementById('edit-dependencies-select') as HTMLSelectElement;
const editDependenciesContainer = document.getElementById('edit-dependencies-container') as HTMLDivElement;
const editNotesTextarea = document.getElementById('edit-notes-textarea') as HTMLTextAreaElement;
const editCustomFieldsContainer = document.getElementById('edit-custom-fields-container') as HTMLDivElement;
const closeModalButton = document.getElementById('close-modal') as HTMLButtonElement;
const deleteTaskButton = document.getElementById('delete-task-btn') as HTMLButtonElement;

// Custom Fields Modal DOM Elements
const customFieldsModal = document.getElementById('custom-fields-modal') as HTMLDivElement;
const manageCustomFieldsBtn = document.getElementById('manage-custom-fields-btn') as HTMLButtonElement;
const closeCustomFieldsModalBtn = document.getElementById('close-custom-fields-modal') as HTMLButtonElement;
const addCustomFieldForm = document.getElementById('add-custom-field-form') as HTMLFormElement;
const customFieldNameInput = document.getElementById('custom-field-name-input') as HTMLInputElement;
const customFieldsList = document.getElementById('custom-fields-list') as HTMLDivElement;

// Multiple Sub-tasks Modal DOM Elements
const multipleSubtasksModal = document.getElementById('multiple-subtasks-modal') as HTMLDivElement;
const closeMultipleSubtasksModalBtn = document.getElementById('close-multiple-subtasks-modal') as HTMLButtonElement;
const cancelMultipleSubtasksBtn = document.getElementById('cancel-multiple-subtasks-btn') as HTMLButtonElement;
const submitMultipleSubtasksBtn = document.getElementById('submit-multiple-subtasks-btn') as HTMLButtonElement;
const multipleSubtasksTextarea = document.getElementById('multiple-subtasks-textarea') as HTMLTextAreaElement;
const multipleSubtasksParentName = document.getElementById('multiple-subtasks-parent-name') as HTMLSpanElement;

// Page & Tab Elements
const tabNavigation = document.querySelector('.tab-navigation') as HTMLElement;
const ganttViewContainer = document.getElementById('gantt-view-container') as HTMLDivElement;
const reportPage = document.getElementById('report-page') as HTMLDivElement;
const instructionsPage = document.getElementById('instructions-page') as HTMLDivElement;
const meetingMinutesPage = document.getElementById('meeting-minutes-page') as HTMLDivElement;
const actionItemsPage = document.getElementById('action-items-page') as HTMLDivElement;
const previousReportsPage = document.getElementById('previous-reports-page') as HTMLDivElement;
const tabContents = document.querySelectorAll('.tab-content') as NodeListOf<HTMLElement>;
const exportPdfBtn = document.getElementById('export-pdf-btn') as HTMLButtonElement;
const saveToArchiveBtn = document.getElementById('save-to-archive-btn') as HTMLButtonElement;
const upcomingDaysInput = document.getElementById('upcoming-days-input') as HTMLInputElement;

// Meeting Minutes DOM Elements
const addMeetingMinuteForm = document.getElementById('add-meeting-minute-form') as HTMLFormElement;
const meetingDateInput = document.getElementById('meeting-date') as HTMLInputElement;
const meetingMinutesTextarea = document.getElementById('meeting-minutes-textarea') as HTMLTextAreaElement;
const meetingMinutesList = document.getElementById('meeting-minutes-list') as HTMLDivElement;

// Action Items DOM Elements
const addActionItemForm = document.getElementById('add-action-item-form') as HTMLFormElement;
const actionItemWhatInput = document.getElementById('action-item-what') as HTMLInputElement;
const actionItemAssignedToInput = document.getElementById('action-item-assigned-to') as HTMLInputElement;
const actionItemAssignedByInput = document.getElementById('action-item-assigned-by') as HTMLInputElement;
const actionItemAssignedDateInput = document.getElementById('action-item-assigned-date') as HTMLInputElement;
const actionItemDueDateInput = document.getElementById('action-item-due-date') as HTMLInputElement;
const actionItemStatusSelect = document.getElementById('action-item-status') as HTMLSelectElement;
const actionItemsTableBody = document.getElementById('action-items-list-tbody') as HTMLTableSectionElement;

// Previous Reports DOM Elements
const previousReportsList = document.getElementById('previous-reports-list') as HTMLDivElement;


// PDF Export Modal Elements
const pdfExportModal = document.getElementById('pdf-export-modal') as HTMLDivElement;
const closePdfModalBtn = document.getElementById('close-pdf-modal') as HTMLButtonElement;
const cancelPdfExportBtn = document.getElementById('cancel-pdf-export-btn') as HTMLButtonElement;
const generateGanttPdfBtn = document.getElementById('generate-gantt-pdf-btn') as HTMLButtonElement;

// Confirmation Modal Elements
const confirmationModal = document.getElementById('confirmation-modal') as HTMLDivElement;
const confirmationMessage = document.getElementById('confirmation-message') as HTMLParagraphElement;
const confirmationCancelBtn = document.getElementById('confirmation-cancel-btn') as HTMLButtonElement;
const confirmationConfirmBtn = document.getElementById('confirmation-confirm-btn') as HTMLButtonElement;
let onConfirmCallback: (() => void) | null = null;
let onCancelCallback: (() => void) | null = null;

// --- COLOR GENERATION & MANAGEMENT ---
const statusColorMap: { [key in TaskStatus]: string } = {
    'On Track': '#2ecc71',
    'At Risk': '#f39c12',
    'Delayed': '#e74c3c',
    'Complete': '#95a5a6'
};

const colorPalette = [
    '#3498db', '#1abc9c', '#9b59b6', '#34495e', '#f1c40f',
    '#e67e22', '#16a085', '#27ae60', '#2980b9', '#8e44ad',
    '#2c3e50', '#c0392b', '#d35400', '#7f8c8d'
];

const colorCache = new Map<string, string>();

/**
 * Generates a consistent color from a string.
 * @param str The input string (e.g., resource name).
 * @returns A hex color string.
 */
function generateColorFromString(str: string): string {
    if (colorCache.has(str)) {
        return colorCache.get(str)!;
    }
    if (!str) {
        return '#bdc3c7'; // Default color for unassigned
    }

    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = str.charCodeAt(i) + ((hash << 5) - hash);
    }
    const index = Math.abs(hash % colorPalette.length);
    const color = colorPalette[index];
    colorCache.set(str, color);
    return color;
}

/**
 * Darkens a hex color by a given percentage.
 * @param color The hex color string.
 * @param percent The percentage to darken (0-100).
 * @returns The darkened hex color string.
 */
function darkenColor(color: string, percent: number): string {
    const num = parseInt(color.replace("#", ""), 16);
    const amt = Math.round(2.55 * percent);
    const R = (num >> 16) - amt;
    const G = (num >> 8 & 0x00FF) - amt;
    const B = (num & 0x0000FF) - amt;
    return "#" + (0x1000000 + (R<255?R<1?0:R:255)*0x10000 + (G<255?G<1?0:G:255)*0x100 + (B<255?B<1?0:B:255)).toString(16).slice(1);
}

// --- DATE UTILITIES (UTC-SAFE) ---

/**
 * Parses a 'YYYY-MM-DD' string into a Date object at UTC midnight.
 * This prevents timezone-related off-by-one errors.
 * @param {string} dateStr The date string to parse.
 * @returns {Date} A Date object set to midnight UTC.
 */
function parseDateUTC(dateStr: string): Date {
    // Appending 'T00:00:00.000Z' forces the parser to interpret the date as UTC.
    return new Date(`${dateStr}T00:00:00.000Z`);
}

/**
 * Formats a Date object into a 'YYYY-MM-DD' string based on its UTC values.
 * @param {Date} date The date object to format.
 * @returns {string} The formatted date string.
 */
function formatDateUTC(date: Date): string {
    const year = date.getUTCFullYear();
    const month = (date.getUTCMonth() + 1).toString().padStart(2, '0');
    const day = date.getUTCDate().toString().padStart(2, '0');
    return `${year}-${month}-${day}`;
}


/**
 * Calculates the difference in days between two dates, ignoring timezone offsets.
 * @param {string} dateStr1 - The first date string (YYYY-MM-DD).
 * @param {string} dateStr2 - The second date string (YYYY-MM-DD).
 * @returns {number} The difference in days.
 */
function dateDiffInDays(dateStr1: string, dateStr2: string): number {
    const dt1 = parseDateUTC(dateStr1);
    const dt2 = parseDateUTC(dateStr2);
    // getTime() returns milliseconds since epoch, which is timezone-independent.
    return Math.round((dt2.getTime() - dt1.getTime()) / (1000 * 60 * 60 * 24));
}

/**
 * Adds a specified number of days to a date string.
 * @param {string} dateStr - The date string (YYYY-MM-DD).
 * @param {number} days - The number of days to add.
 * @returns {string} The new date string.
 */
function addDays(dateStr: string, days: number): string {
    const date = parseDateUTC(dateStr);
    date.setUTCDate(date.getUTCDate() + days);
    return formatDateUTC(date);
}

/**
 * Initializes flatpickr calendar pickers on all date inputs.
 */
function initializeDatePickers(): void {
    const dateInputs = document.querySelectorAll('input[type="date"]');
    dateInputs.forEach(input => {
        const instance = flatpickr(input, {
            dateFormat: "Y-m-d", // Keep backend format consistent
            altInput: true,      // Show a user-friendly format
            altFormat: "M j, Y", // e.g., "Aug 1, 2024"
            disableMobile: true, // Use custom UI on mobile for consistency
        });
        flatpickrInstances[input.id] = instance;
    });
}

/**
 * Renders the project information form with data from the state.
 */
function renderProjectInfo(): void {
    if (!projectInfo) return;
    projectNameInput.value = projectInfo.name || '';
    projectDescriptionTextarea.value = projectInfo.description || '';
    projectDeliverablesTextarea.value = projectInfo.deliverables || '';
    projectLocationInput.value = projectInfo.location || '';
    projectDocSelect.value = projectInfo.doc || '';
    projectOwnerInput.value = projectInfo.owner || '';
    projectEngineeringManagerInput.value = projectInfo.engineeringManager || '';
    projectManagerInput.value = projectInfo.manager || '';
    projectEngineersInput.value = projectInfo.engineers || '';
    projectStartDateInput.value = projectInfo.startDate || '';
    projectEndDateInput.value = projectInfo.endDate || '';
    projectBudgetInput.value = String(projectInfo.budget || '');
    projectCurrencyInput.value = projectInfo.currency || 'USD';
}


/**
 * Recalculates the status for a given task based on predefined rules.
 * @param {Task} task - The task to calculate the status for.
 * @returns {TaskStatus} The calculated status.
 */
function getCalculatedStatus(task: Task): TaskStatus {
    if (task.percentComplete === 100) return 'Complete';
    
    const now = new Date();
    const todayUTC = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
    const endDate = parseDateUTC(task.endDate);

    if (endDate < todayUTC) return 'Delayed';

    const threeDaysFromNow = new Date(todayUTC);
    threeDaysFromNow.setUTCDate(todayUTC.getUTCDate() + 3);

    if (endDate <= threeDaysFromNow && task.percentComplete < 75) return 'At Risk';
    
    return 'On Track';
}

/**
 * Updates the status of all tasks that are set to 'auto'.
 */
function updateAllTaskStatuses(): void {
    tasks.forEach(task => {
        if (task.status === 'auto') {
            // This is a temporary update for rendering, don't save it back to the task object
            // The real status is calculated on-the-fly when needed.
        }
    });
}


/**
 * Recursively updates the start/end dates and completion percentage of parent tasks.
 */
function updateParentTasks(): void {
  let changed = true;
  
  // Loop until no more changes occur, ensuring propagation up the hierarchy
  while (changed) {
    changed = false;
    
    // Create a copy to iterate over, as we modify the original tasks array
    const taskIds = tasks.map(t => t.id);

    taskIds.forEach(taskId => {
      const task = tasks.find(t => t.id === taskId);
      if (!task) return;

      const children = tasks.filter(t => t.parentId === task.id);
      if (children.length > 0) {
        // --- Date Roll-up ---
        const startDates = children.map(c => parseDateUTC(c.startDate));
        const endDates = children.map(c => parseDateUTC(c.endDate));
        const newMinDate = new Date(Math.min(...startDates.map(d => d.getTime())));
        const newMaxDate = new Date(Math.max(...endDates.map(d => d.getTime())));
        const newStartDateStr = formatDateUTC(newMinDate);
        const newEndDateStr = formatDateUTC(newMaxDate);

        if (task.startDate !== newStartDateStr || task.endDate !== newEndDateStr) {
          task.startDate = newStartDateStr;
          task.endDate = newEndDateStr;
          changed = true;
        }

        // --- Percentage Complete Roll-up (Weighted by Duration) ---
        let totalWeightedCompletion = 0;
        let totalDuration = 0;
        children.forEach(child => {
            const duration = dateDiffInDays(child.startDate, child.endDate) + 1;
            if(duration > 0) {
              totalWeightedCompletion += (child.percentComplete || 0) * duration;
              totalDuration += duration;
            }
        });

        const newPercentComplete = totalDuration > 0
            ? Math.round(totalWeightedCompletion / totalDuration)
            : 0;
        
        if (task.percentComplete !== newPercentComplete) {
            task.percentComplete = newPercentComplete;
            changed = true;
        }
      }
    });
  }
}

/**
 * Applies current filters and sorting to the tasks.
 * @returns {Task[]} The filtered and sorted array of tasks.
 */
function getFilteredAndSortedTasks(): Task[] {
    let filteredTasks = tasks.filter(task => {
        const nameMatch = filters.name ? task.name.toLowerCase().includes(filters.name.toLowerCase()) : true;
        const resourceMatch = filters.resource ? task.resources?.some(r => r.toLowerCase().includes(filters.resource.toLowerCase())) : true;
        const currentStatus = task.status === 'auto' ? getCalculatedStatus(task) : task.status;
        const statusMatch = filters.status ? currentStatus === filters.status : true;
        return nameMatch && resourceMatch && statusMatch;
    });

    filteredTasks.sort((a, b) => {
        let valA, valB;
        const key = sortCriteria.key;

        switch (key) {
            case 'name':
                valA = a.name.toLowerCase();
                valB = b.name.toLowerCase();
                break;
            case 'startDate':
            case 'endDate':
                valA = parseDateUTC(a[key]).getTime();
                valB = parseDateUTC(b[key]).getTime();
                break;
            default:
                return 0;
        }
        
        if (valA < valB) return sortCriteria.order === 'asc' ? -1 : 1;
        if (valA > valB) return sortCriteria.order === 'asc' ? 1 : -1;
        return 0;
    });

    return filteredTasks;
}

/**
 * Renders the tasks in the task list recursively.
 */
function renderTaskList(): void {
    if (!taskList) return;
    taskList.innerHTML = '';
    const filteredAndSortedTasks = getFilteredAndSortedTasks();
    const taskTree = filteredAndSortedTasks.filter(task => {
        // Only include top-level tasks or children of visible parents
        return task.parentId === null || filteredAndSortedTasks.some(t => t.id === task.parentId);
    });

    const topLevelTasks = taskTree.filter(task => task.parentId === null);
    topLevelTasks.forEach(task => renderTaskListItem(task, 0, taskTree));
}

/**
 * Renders a single task list item and its children.
 * @param {Task} task - The task to render.
 * @param {number} level - The indentation level.
 * @param {Task[]} availableTasks - The pre-filtered and sorted list of tasks.
 */
function renderTaskListItem(task: Task, level: number, availableTasks: Task[]): void {
    const children = availableTasks.filter(t => t.parentId === task.id);
    const isParent = tasks.some(t => t.parentId === task.id); // Check original tasks for parent status
    const isExpanded = expandedTaskIds.has(task.id);
    const effectiveStatus = task.status === 'auto' ? getCalculatedStatus(task) : task.status;

    const listItem = document.createElement('li');
    listItem.className = 'task-item-container';
    listItem.id = `task-list-item-${task.id}`;

    const itemContent = document.createElement('div');
    itemContent.className = 'task-item';
    itemContent.setAttribute('data-task-id', task.id.toString());
    itemContent.style.paddingLeft = `${level * 25}px`;
    itemContent.classList.add(`status-${effectiveStatus.toLowerCase().replace(' ', '-')}`);
    itemContent.onclick = () => openEditModal(task.id);

    const taskInfo = document.createElement('div');
    taskInfo.className = 'task-info';
    
    const taskNameContainer = document.createElement('div');
    taskNameContainer.className = 'task-name-container';

    if (isParent) {
        const toggleButton = document.createElement('button');
        toggleButton.className = 'toggle-btn';
        toggleButton.innerHTML = isExpanded ? '&#9662;' : '&#9656;'; // Down/Right arrow
        toggleButton.setAttribute('aria-expanded', String(isExpanded));
        toggleButton.onclick = (e) => {
            e.stopPropagation();
            toggleExpand(task.id);
        };
        taskNameContainer.appendChild(toggleButton);
    }

    const taskNameSpan = document.createElement('span');
    taskNameSpan.className = 'task-name';
    taskNameSpan.textContent = task.name;
    taskNameContainer.appendChild(taskNameSpan);

    if (task.notes) {
        const notesIcon = document.createElement('span');
        notesIcon.className = 'notes-icon';
        notesIcon.innerHTML = '&#128172;'; // Speech bubble emoji
        notesIcon.title = task.notes;
        taskNameContainer.appendChild(notesIcon);
    }

    const duration = dateDiffInDays(task.startDate, task.endDate) + 1;

    const taskDetails = document.createElement('div');
    taskDetails.className = 'task-details';
    taskDetails.innerHTML = `
        <span class="task-status"><span class="status-dot"></span>${effectiveStatus}</span>
        <span>${task.startDate}</span>
        <span class="date-arrow" aria-hidden="true">&rarr;</span>
        <span>${task.endDate}</span>
        <span class="task-duration">(${duration} ${duration > 1 ? 'days' : 'day'})</span>
        <span class="task-completion">${Math.round(task.percentComplete)}%</span>
    `;

    const resourcesContainer = document.createElement('div');
    resourcesContainer.className = 'resources-container';
    if (task.resources && task.resources.length > 0) {
        task.resources.forEach(resource => {
            const resourceTag = document.createElement('span');
            resourceTag.className = 'resource-tag';
            resourceTag.textContent = resource;
            resourcesContainer.appendChild(resourceTag);
        });
    }

    const customFieldsContainer = document.createElement('div');
    customFieldsContainer.className = 'custom-fields-display-container';
    if (task.customFields) {
        Object.entries(task.customFields).forEach(([fieldName, fieldValue]) => {
            if (fieldValue) {
                const fieldDiv = document.createElement('div');
                fieldDiv.className = 'custom-field-display';
                fieldDiv.innerHTML = `<span class="custom-field-name">${fieldName}:</span> <span class="custom-field-value">${fieldValue}</span>`;
                customFieldsContainer.appendChild(fieldDiv);
            }
        });
    }
    
    taskInfo.appendChild(taskNameContainer);
    taskInfo.appendChild(taskDetails);
    taskInfo.appendChild(resourcesContainer);
    if (customFieldsContainer.hasChildNodes()) {
        taskInfo.appendChild(customFieldsContainer);
    }


    const taskActions = document.createElement('div');
    taskActions.className = 'task-actions';

    const addSubtasksButton = document.createElement('button');
    addSubtasksButton.className = 'add-subtasks-btn';
    addSubtasksButton.textContent = 'Add Sub-tasks';
    addSubtasksButton.onclick = () => openMultipleSubtasksModal(task.id);
    
    const editButton = document.createElement('button');
    editButton.className = 'edit-btn';
    editButton.textContent = 'Edit';
    editButton.onclick = () => openEditModal(task.id);

    const deleteButton = document.createElement('button');
    deleteButton.className = 'delete-btn-list';
    deleteButton.textContent = 'Delete';
    deleteButton.onclick = () => {
        showConfirmation(
            `Are you sure you want to delete "${task.name}" and all its sub-tasks?`,
            () => deleteTask(task.id)
        );
    };

    taskActions.appendChild(addSubtasksButton);
    taskActions.appendChild(editButton);
    taskActions.appendChild(deleteButton);

    itemContent.appendChild(taskInfo);
    itemContent.appendChild(taskActions);
    listItem.appendChild(itemContent);
    taskList.appendChild(listItem);

    if (isExpanded && isParent) {
        children.forEach(child => renderTaskListItem(child, level + 1, availableTasks));
    }
}

/**
 * Toggles the expanded/collapsed state of a parent task.
 * @param {number} taskId The ID of the task to toggle.
 */
function toggleExpand(taskId: number): void {
    if (expandedTaskIds.has(taskId)) {
        expandedTaskIds.delete(taskId);
    } else {
        expandedTaskIds.add(taskId);
    }
    render(); // Re-render the whole UI to reflect the change
}

/**
 * Builds a flat list of visible tasks for the Gantt chart.
 * This function respects the expanded/collapsed state of parent tasks.
 * @returns {Task[]} A flat array of tasks to be rendered.
 */
function getVisibleTasks(): Task[] {
    const availableTasks = getFilteredAndSortedTasks();
    const visible: Task[] = [];

    // Recursive function to add tasks and their expanded children
    function addTasks(parentId: number | null) {
        const children = availableTasks.filter(task => task.parentId === parentId);
        children.forEach(task => {
            visible.push(task);
            // If the task is expanded, recursively add its children
            if (expandedTaskIds.has(task.id)) {
                addTasks(task.id);
            }
        });
    }

    addTasks(null); // Start with top-level tasks
    return visible;
}

/**
 * Renders the Gantt chart visualization.
 */
function renderGanttChart(): void {
    if (!ganttChartArea) return;
    
    // Remove the old grid if it exists
    const oldGrid = ganttChartArea.querySelector('.gantt-grid');
    if (oldGrid) {
        oldGrid.remove();
    }

    const visibleTasks = getVisibleTasks();
    if (visibleTasks.length === 0) {
        ganttChartArea.classList.remove('has-content');
        dependencyLinesSvg.innerHTML = '';
        return;
    }

    ganttChartArea.classList.add('has-content');

    const allDates = tasks.flatMap(task => [parseDateUTC(task.startDate), parseDateUTC(task.endDate)]);
    if (allDates.length === 0) return;
    const minDate = new Date(Math.min(...allDates.map(d => d.getTime())));
    const maxDate = new Date(Math.max(...allDates.map(d => d.getTime())));

    const chartGrid = document.createElement('div');
    chartGrid.className = 'gantt-grid';

    const header = document.createElement('div');
    header.className = 'gantt-header';
    generateTimelineHeader(header, minDate, maxDate);
    chartGrid.appendChild(header);

    const minDateStr = formatDateUTC(minDate);
    const totalDurationInDays = dateDiffInDays(minDateStr, formatDateUTC(maxDate)) + 1;
    const totalWidth = totalDurationInDays * dayWidth;
    chartGrid.style.width = `${totalWidth}px`;

    let currentTopOffset = 40; // Initial offset for the timeline header
    const barPositions = new Map<number, { top: number, left: number, width: number }>();

    visibleTasks.forEach(task => {
        const isParent = tasks.some(t => t.parentId === task.id);
        const isExpanded = expandedTaskIds.has(task.id);
        const offsetInDays = dateDiffInDays(minDateStr, task.startDate);
        const durationInDays = dateDiffInDays(task.startDate, task.endDate) + 1;
        const barLeft = offsetInDays * dayWidth;
        const barWidth = durationInDays * dayWidth - 2;

        if (isParent) {
            // Render an interactive summary bar for parent tasks
            const summaryBar = document.createElement('div');
            summaryBar.className = 'gantt-summary-bar';
            summaryBar.style.left = `${barLeft}px`;
            summaryBar.style.width = `${barWidth}px`;
            summaryBar.style.top = `${currentTopOffset}px`;
            summaryBar.setAttribute('data-task-id', task.id.toString());
            summaryBar.title = `${task.name}\n${task.startDate} to ${task.endDate}`;
            summaryBar.onclick = () => openEditModal(task.id);

            const toggleButton = document.createElement('button');
            toggleButton.className = 'toggle-btn';
            toggleButton.innerHTML = isExpanded ? '&#9662;' : '&#9656;'; // Down/Right arrow
            toggleButton.setAttribute('aria-expanded', String(isExpanded));
            toggleButton.onclick = (e) => {
                e.stopPropagation();
                toggleExpand(task.id);
            };

            const label = document.createElement('span');
            label.className = 'gantt-bar-label';
            label.textContent = task.name;

            summaryBar.appendChild(toggleButton);
            summaryBar.appendChild(label);
            chartGrid.appendChild(summaryBar);
            barPositions.set(task.id, { top: currentTopOffset, left: barLeft, width: barWidth });

        } else {
            // Render a regular task bar for non-parent (leaf) tasks
            const effectiveStatus = task.status === 'auto' ? getCalculatedStatus(task) : task.status;

            const taskBar = document.createElement('div');
            taskBar.className = 'gantt-bar';
            taskBar.classList.add(`status-${effectiveStatus.toLowerCase().replace(' ', '-')}`);

            let barColor = '#4a90e2'; // Default color
            switch (colorCodeCriteria) {
                case 'status': barColor = statusColorMap[effectiveStatus]; break;
                case 'resource':
                    const resource = task.resources?.[0] || 'Unassigned';
                    barColor = generateColorFromString(resource);
                    break;
                default:
                    const customFieldValue = task.customFields?.[colorCodeCriteria] || 'Not Set';
                    barColor = generateColorFromString(customFieldValue);
                    break;
            }
            taskBar.style.backgroundColor = barColor;

            const taskBarLabel = document.createElement('span');
            taskBarLabel.className = 'gantt-bar-label';
            taskBarLabel.textContent = task.name;
            
            const progressBar = document.createElement('div');
            progressBar.className = 'gantt-progress';
            progressBar.style.width = `${task.percentComplete}%`;
            progressBar.style.backgroundColor = darkenColor(barColor, 20);

            taskBar.appendChild(progressBar);
            taskBar.appendChild(taskBarLabel);
            
            taskBar.style.left = `${barLeft}px`;
            taskBar.style.width = `${barWidth}px`;
            taskBar.style.top = `${currentTopOffset}px`;
            taskBar.setAttribute('data-task-id', task.id.toString());
            
            barPositions.set(task.id, { top: currentTopOffset, left: barLeft, width: barWidth });

            let tooltip = `${task.name}\n${task.startDate} to ${task.endDate}\n${task.percentComplete}% Complete\nStatus: ${effectiveStatus}`;
            if (task.description) tooltip += `\nDescription: ${task.description}`;
            if (task.resources && task.resources.length > 0) tooltip += `\nResources: ${task.resources.join(', ')}`;
            if (task.notes) tooltip += `\nNotes: ${task.notes}`;
            if (task.customFields) Object.entries(task.customFields).forEach(([key, value]) => { if (value) tooltip += `\n${key}: ${value}`; });
            taskBar.title = tooltip;

            taskBar.onclick = () => openEditModal(task.id);
            
            const leftHandle = document.createElement('div');
            leftHandle.className = 'resize-handle left';
            leftHandle.onmousedown = (e) => initResize(e, task.id, 'left', minDateStr);
            
            const rightHandle = document.createElement('div');
            rightHandle.className = 'resize-handle right';
            rightHandle.onmousedown = (e) => initResize(e, task.id, 'right', minDateStr);
            
            taskBar.appendChild(leftHandle);
            taskBar.appendChild(rightHandle);

            taskBar.onmousedown = (e) => {
                // Do not initiate move if the target is a resize handle
                const target = e.target as HTMLElement;
                if (target.classList.contains('resize-handle')) return;
                initMove(e, task.id);
            };

            if (showBaseline) {
                const baselineTask = baselineTasks.find(bt => bt.id === task.id);
                if (baselineTask) {
                    const baselineOffsetInDays = dateDiffInDays(minDateStr, baselineTask.startDate);
                    const baselineDurationInDays = dateDiffInDays(baselineTask.startDate, baselineTask.endDate) + 1;
                    
                    const baselineBar = document.createElement('div');
                    baselineBar.className = 'gantt-bar-baseline';
                    baselineBar.style.left = `${baselineOffsetInDays * dayWidth}px`;
                    baselineBar.style.width = `${baselineDurationInDays * dayWidth - 2}px`;
                    baselineBar.style.top = `${currentTopOffset + 14}px`;
                    baselineBar.title = `Baseline: ${baselineTask.startDate} to ${baselineTask.endDate}`;
                    chartGrid.appendChild(baselineBar);
                }
            }
            
            chartGrid.appendChild(taskBar);
        }
        currentTopOffset += 40; // Unified height for all rows
    });


    // Add Today Marker
    const now = new Date();
    const today = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
    if (today >= minDate && today <= maxDate) {
        const todayOffsetInDays = dateDiffInDays(minDateStr, formatDateUTC(today));
        const todayMarker = document.createElement('div');
        todayMarker.id = 'today-marker';
        todayMarker.className = 'today-marker';
        todayMarker.style.left = `${todayOffsetInDays * dayWidth}px`;
        todayMarker.title = `Today: ${today.toLocaleDateString()}`;
        chartGrid.appendChild(todayMarker);
    }

    const finalHeight = currentTopOffset + 20;
    chartGrid.style.height = `${finalHeight}px`; // Add some padding at the bottom
    ganttChartArea.appendChild(chartGrid);

    // Set SVG dimensions and draw dependency lines
    dependencyLinesSvg.setAttribute('width', totalWidth.toString());
    dependencyLinesSvg.setAttribute('height', finalHeight.toString());
    dependencyLinesSvg.innerHTML = `
        <defs>
            <marker id="arrowhead" markerWidth="10" markerHeight="7" 
            refX="9" refY="3.5" orient="auto">
                <polygon points="0 0, 10 3.5, 0 7" class="dependency-arrow" />
            </marker>
        </defs>
    `;

    visibleTasks.forEach(task => {
        if (task.dependencies && task.dependencies.length > 0) {
            const successorPos = barPositions.get(task.id);
            if (!successorPos) return;

            task.dependencies.forEach(depId => {
                const predecessorPos = barPositions.get(depId);
                if (!predecessorPos) return;

                const startX = predecessorPos.left + predecessorPos.width + 2; // from end of predecessor
                const startY = predecessorPos.top + 14;
                const endX = successorPos.left; // to start of successor
                const endY = successorPos.top + 14;

                const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
                const d = `M ${startX} ${startY} L ${startX + 10} ${startY} L ${startX + 10} ${endY} L ${endX} ${endY}`;
                path.setAttribute('d', d);
                path.setAttribute('class', 'dependency-line');
                path.setAttribute('marker-end', 'url(#arrowhead)');
                dependencyLinesSvg.appendChild(path);
            });
        }
    });
}

/**
 * Generates the timeline header based on the current view.
 */
function generateTimelineHeader(header: HTMLElement, minDate: Date, maxDate: Date): void {
    let currentDate = new Date(minDate.getTime());
    const now = new Date();
    const today = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
    header.innerHTML = '';

    const addHeader = (label: string, isToday: boolean = false, width: number) => {
        const headerEl = document.createElement('div');
        headerEl.className = 'gantt-day-header';
        if (isToday) headerEl.classList.add('today');
        headerEl.textContent = label;
        headerEl.style.width = `${width}px`;
        headerEl.style.flex = `0 0 ${width}px`; // Prevent shrinking/growing
        header.appendChild(headerEl);
    };
    
    switch(currentView) {
        case 'day':
            while (currentDate <= maxDate) {
                addHeader(`${currentDate.getUTCMonth() + 1}/${currentDate.getUTCDate()}`, currentDate.getTime() === today.getTime(), dayWidth);
                currentDate.setUTCDate(currentDate.getUTCDate() + 1);
            }
            break;
        case 'week':
            let weekCount = 1;
            while(currentDate <= maxDate) {
                const weekStart = new Date(currentDate.getTime());
                const nextWeek = new Date(currentDate.getTime());
                nextWeek.setUTCDate(currentDate.getUTCDate() + 7);

                const daysInChunk = dateDiffInDays(formatDateUTC(currentDate), formatDateUTC(new Date(Math.min(nextWeek.getTime() - 1, maxDate.getTime())))) + 1;
                const width = daysInChunk * dayWidth;

                addHeader(`W${weekCount}`, today >= weekStart && today < nextWeek, width);
                currentDate = nextWeek;
                weekCount++;
            }
            break;
        case 'month':
             currentDate = new Date(minDate.getTime());
             while (currentDate <= maxDate) {
                const monthStart = new Date(currentDate.getTime());
                const nextMonth = new Date(Date.UTC(monthStart.getUTCFullYear(), monthStart.getUTCMonth() + 1, 1));
                
                const daysInChunk = dateDiffInDays(formatDateUTC(currentDate), formatDateUTC(new Date(Math.min(nextMonth.getTime() - 1, maxDate.getTime())))) + 1;
                const width = daysInChunk * dayWidth;

                addHeader(
                    monthStart.toLocaleString('default', { timeZone: 'UTC', month: 'short', year: '2-digit' }),
                    today.getUTCFullYear() === monthStart.getUTCFullYear() && today.getUTCMonth() === monthStart.getUTCMonth(),
                    width
                );
                currentDate = nextMonth;
             }
             break;
        case 'quarter':
             currentDate = new Date(minDate.getTime());
             while(currentDate <= maxDate) {
                const quarter = Math.floor(currentDate.getUTCMonth() / 3) + 1;
                const year = currentDate.getUTCFullYear();
                const nextQuarter = new Date(Date.UTC(year, quarter * 3, 1));

                const daysInChunk = dateDiffInDays(formatDateUTC(currentDate), formatDateUTC(new Date(Math.min(nextQuarter.getTime() - 1, maxDate.getTime())))) + 1;
                const width = daysInChunk * dayWidth;

                addHeader(`Q${quarter} ${year}`, false, width);
                currentDate = nextQuarter;
            }
            break;
        case 'year':
            currentDate = new Date(minDate.getTime());
            while(currentDate <= maxDate) {
                const year = currentDate.getUTCFullYear();
                const nextYear = new Date(Date.UTC(year + 1, 0, 1));

                const daysInChunk = dateDiffInDays(formatDateUTC(currentDate), formatDateUTC(new Date(Math.min(nextYear.getTime() - 1, maxDate.getTime())))) + 1;
                const width = daysInChunk * dayWidth;

                addHeader(year.toString(), false, width);
                currentDate = nextYear;
            }
            break;
    }
}


/**
 * Initializes the drag-to-move functionality for a task bar.
 * This function distinguishes between a drag and a click to prevent
 * the click event from being cancelled by an unnecessary re-render.
 */
function initMove(event: MouseEvent, taskId: number) {
    const task = tasks.find(t => t.id === taskId);
    const bar = document.querySelector(`.gantt-bar[data-task-id="${taskId}"]`) as HTMLDivElement;
    if (!task || !bar) return;

    const startX = event.clientX;
    const initialLeft = bar.offsetLeft;
    const initialStartDate = task.startDate;
    const taskDuration = dateDiffInDays(task.startDate, task.endDate);
    let hasMoved = false;

    function onMouseMove(moveEvent: MouseEvent) {
        const deltaX = moveEvent.clientX - startX;
        // Update visual position directly for smooth feedback.
        bar.style.left = `${initialLeft + deltaX}px`;
        
        // If the mouse has moved more than a small threshold, consider it a drag.
        if (Math.abs(deltaX) > 4) {
            hasMoved = true;
            bar.classList.add('dragging'); // Add class only when dragging confirmed
        }
    }

    function onMouseUp(upEvent: MouseEvent) {
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);

        // Only update and re-render if the task was actually dragged.
        if (hasMoved) {
            const deltaX = upEvent.clientX - startX;
            const deltaDays = Math.round(deltaX / dayWidth);

            if (deltaDays !== 0) {
                const newStartDate = addDays(initialStartDate, deltaDays);
                const newEndDate = addDays(newStartDate, taskDuration);
                
                task.startDate = newStartDate;
                task.endDate = newEndDate;
                updateParentTasks();
            }
            bar.classList.remove('dragging');
            // Re-render to commit the changes and snap the bar to its final grid position.
            render();
        } else {
            // It was a click, not a drag.
            // Snap the bar back to its original position to correct any minor visual jitter.
            bar.style.left = `${initialLeft}px`;
            // Do NOT call render(), which would destroy the element and prevent the `onclick` event from firing.
        }
    }

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp, { once: true });
}

/**
 * Initializes the drag-to-resize functionality for a task bar handle.
 * This function distinguishes between a drag and a click to prevent
 * the click event from being cancelled by an unnecessary re-render.
 */
function initResize(event: MouseEvent, taskId: number, handle: 'left' | 'right', minDateStr: string) {
    event.preventDefault();
    event.stopPropagation(); // Prevent move action
    
    const task = tasks.find(t => t.id === taskId);
    const bar = document.querySelector(`.gantt-bar[data-task-id="${taskId}"]`) as HTMLDivElement;
    if (!task || !bar) return;

    const startX = event.clientX;
    const initialStartDate = task.startDate;
    const initialEndDate = task.endDate;
    const initialLeft = bar.offsetLeft;
    const initialWidth = bar.offsetWidth;
    let hasResized = false;

    function onMouseMove(moveEvent: MouseEvent) {
        const deltaX = moveEvent.clientX - startX;
        
        if (Math.abs(deltaX) > 4) {
            hasResized = true;
            bar.classList.add('dragging');
        }

        if (handle === 'left') {
            const newLeft = initialLeft + deltaX;
            const newWidth = initialWidth - deltaX;
            if(newWidth < dayWidth/2) return; // Prevent shrinking too small
            bar.style.left = `${newLeft}px`;
            bar.style.width = `${newWidth}px`;
        } else { // handle === 'right'
            const newWidth = initialWidth + deltaX;
            if(newWidth < dayWidth/2) return;
            bar.style.width = `${newWidth}px`;
        }
    }

    function onMouseUp(upEvent: MouseEvent) {
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
        
        // Only update and re-render if a resize actually occurred.
        if (hasResized) {
            const deltaX = upEvent.clientX - startX;
            const deltaDays = Math.round(deltaX / dayWidth);

            if (deltaDays !== 0) {
                if (handle === 'left') {
                    const newStartDate = addDays(initialStartDate, deltaDays);
                    if (parseDateUTC(newStartDate) <= parseDateUTC(initialEndDate)) {
                        task.startDate = newStartDate;
                        updateParentTasks();
                    }
                } else { // handle === 'right'
                    const newEndDate = addDays(initialEndDate, deltaDays);
                    if (parseDateUTC(newEndDate) >= parseDateUTC(initialStartDate)) {
                        task.endDate = newEndDate;
                        updateParentTasks();
                    }
                }
            }
            bar.classList.remove('dragging');
            // Re-render to commit the changes and snap the bar to its final grid size.
            render();
        } else {
            // It was a click, not a resize.
            // Snap the bar back to its original visual state.
            bar.style.left = `${initialLeft}px`;
            bar.style.width = `${initialWidth}px`;
            // Do NOT call render() to allow the parent bar's onclick event to fire.
        }
    }
    
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp, { once: true });
}

/**
 * Gets all descendant task IDs for a given task.
 * @param taskId The ID of the parent task.
 * @returns An array of descendant task IDs.
 */
function getDescendants(taskId: number): number[] {
    let descendants: number[] = [];
    const children = tasks.filter(t => t.parentId === taskId);
    children.forEach(child => {
        descendants.push(child.id);
        descendants = descendants.concat(getDescendants(child.id));
    });
    return descendants;
}

/**
 * Opens the edit modal and populates it with task data.
 * @param {number} taskId The ID of the task to edit.
 */
function openEditModal(taskId: number): void {
    const task = tasks.find(t => t.id === taskId);
    if (!task) return;

    editingTaskId = taskId;
    const isParent = tasks.some(t => t.parentId === taskId);
    const isSubtask = task.parentId !== null;

    editTaskNameInput.value = task.name;
    editTaskDescriptionTextarea.value = task.description || '';
    // Set the raw value for the hidden input first
    editStartDateInput.value = task.startDate;
    editEndDateInput.value = task.endDate;
    
    // Un-constrain date pickers for sub-tasks
    const startPicker = flatpickrInstances['edit-start-date-input'];
    if (startPicker) {
        startPicker.set({ minDate: undefined, maxDate: undefined });
        startPicker.setDate(task.startDate, false);
    }
    const endPicker = flatpickrInstances['edit-end-date-input'];
    if (endPicker) {
        endPicker.set({ minDate: undefined, maxDate: undefined });
        endPicker.setDate(task.endDate, false);
    }

    editPercentCompleteInput.value = task.percentComplete.toString();
    editResourcesInput.value = task.resources?.join(', ') || '';
    editNotesTextarea.value = task.notes || '';
    editStatusSelect.value = task.status || 'auto';
    
    renderCustomFieldsForm(editCustomFieldsContainer, task.customFields);

    // Populate dependencies only for main tasks, hide for subtasks
    if (isSubtask) {
        editDependenciesContainer.style.display = 'none';
    } else {
        editDependenciesContainer.style.display = 'block';
        editDependenciesSelect.innerHTML = '';
        const descendants = getDescendants(taskId);
        const possibleDependencies = tasks.filter(t => t.id !== taskId && !descendants.includes(t.id));
        
        const hierarchicalOptions = getHierarchicalTaskOptions(possibleDependencies, null, 0);

        hierarchicalOptions.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt.id.toString();
            option.textContent = opt.name;
            if (task.dependencies?.includes(opt.id)) {
                option.selected = true;
            }
            editDependenciesSelect.appendChild(option);
        });
    }

    // Disable date and percentage inputs for parent tasks
    editStartDateInput.disabled = isParent;
    editEndDateInput.disabled = isParent;
    editPercentCompleteInput.disabled = isParent;
    if (isParent) {
        editStartDateInput.title = "Parent task dates are calculated automatically.";
        editEndDateInput.title = "Parent task dates are calculated automatically.";
        editPercentCompleteInput.title = "Parent task completion is calculated automatically.";
    } else {
        editStartDateInput.title = "";
        editEndDateInput.title = "";
        editPercentCompleteInput.title = "";
    }


    editTaskModal.style.display = 'flex';
    editTaskNameInput.focus();
}

/**
 * Closes the edit modal.
 */
function closeModal(): void {
    editTaskModal.style.display = 'none';
    editingTaskId = null;
}

/**
 * Adds a new task to the list.
 */
function addTask(name: string, startDate: string, endDate: string, parentId: number | null = null, percentComplete: number = 0, resources: string[] = [], customFields: {[key: string]: string} = {}): void {
    const newTask: Task = {
        id: nextTaskId++,
        name,
        startDate,
        endDate,
        parentId,
        percentComplete,
        resources,
        customFields,
        dependencies: [],
        status: 'auto' // Always start with automatic status
    };
    tasks.push(newTask);

    // Check if the new task will be visible with current filters
    const nameMatch = filters.name ? newTask.name.toLowerCase().includes(filters.name.toLowerCase()) : true;
    const resourceMatch = filters.resource ? (newTask.resources || []).some(r => r.toLowerCase().includes(filters.resource.toLowerCase())) : true;
    const calculatedStatus = getCalculatedStatus(newTask); // New tasks are 'auto', so we must calculate it
    const statusMatch = filters.status ? calculatedStatus === filters.status : true;
    const isVisible = nameMatch && resourceMatch && statusMatch;

    // If any filter is active and the new task wouldn't be visible, alert the user.
    if (!isVisible && (filters.name || filters.resource || filters.status)) {
        alert(`Task "${name}" was added successfully, but it is currently hidden by your active filters. Clear the filters to see it.`);
    }

    if (parentId !== null) {
        expandedTaskIds.add(parentId);
    }
    scheduleTasks();
    updateParentTasks();
    render();
}

/**
 * Opens the modal to add multiple sub-tasks.
 * @param {number} parentId The ID of the parent task.
 */
function openMultipleSubtasksModal(parentId: number): void {
    const parentTask = tasks.find(t => t.id === parentId);
    if (!parentTask) return;

    subtaskParentId = parentId;
    multipleSubtasksParentName.textContent = `"${parentTask.name}"`;
    multipleSubtasksModal.style.display = 'flex';
    multipleSubtasksTextarea.focus();
}

/**
 * Handles the submission of the multiple sub-tasks modal.
 */
function handleMultipleSubtasksSubmit(): void {
    if (subtaskParentId === null) return;
    
    const parentTask = tasks.find(t => t.id === subtaskParentId);
    if (!parentTask) return;

    const text = multipleSubtasksTextarea.value.trim();
    if (!text) {
        alert("Text area is empty. Please enter sub-task names.");
        return;
    }

    const taskNames = text.split('\n').map(line => line.trim()).filter(line => line.length > 0);
    if (taskNames.length === 0) {
        alert("No valid sub-task names found.");
        return;
    }
    
    // Use the parent's start date as the default for all new sub-tasks
    const defaultStartDate = parentTask.startDate;

    taskNames.forEach(name => {
        const newTask: Task = {
            id: nextTaskId++,
            name,
            startDate: defaultStartDate,
            endDate: defaultStartDate,
            parentId: subtaskParentId,
            percentComplete: 0,
            status: 'auto'
        };
        tasks.push(newTask);
    });

    expandedTaskIds.add(subtaskParentId);
    
    scheduleTasks();
    updateParentTasks();
    render();

    // Reset and close the modal
    multipleSubtasksTextarea.value = '';
    subtaskParentId = null;
    multipleSubtasksModal.style.display = 'none';

    alert(`${taskNames.length} sub-task(s) added successfully.`);
}

/**
 * Updates an existing task.
 * @param {number} id - The ID of the task to update.
 * @param {Partial<Omit<Task, 'id' | 'parentId'>>} updates - The properties to update.
 */
function updateTask(id: number, updates: Partial<Omit<Task, 'id' | 'parentId'>>): void {
    const taskIndex = tasks.findIndex(t => t.id === id);
    if (taskIndex === -1) return;
    
    // Update task
    tasks[taskIndex] = { ...tasks[taskIndex], ...updates };
    
    scheduleTasks();
    updateParentTasks();
    render();
}


/**
 * Deletes a task and all its sub-tasks recursively.
 * @param {number} taskId The ID of the task to delete.
 */
function deleteTask(taskId: number): void {
    const children = tasks.filter(t => t.parentId === taskId);
    children.forEach(child => deleteTask(child.id)); // Recursively delete children

    // Remove from other tasks' dependencies
    tasks.forEach(t => {
        if (t.dependencies?.includes(taskId)) {
            t.dependencies = t.dependencies.filter(depId => depId !== taskId);
        }
    });

    // Remove the task itself
    tasks = tasks.filter(t => t.id !== taskId);
    
    updateParentTasks();
    render();
}

/**
 * Creates a hierarchical list of tasks for dropdown selectors.
 * @param {Task[]} tasksToProcess - The tasks to include in the list.
 * @param {number | null} parentId - The current parent ID to filter by.
 * @param {number} level - The current indentation level.
 * @returns { {id: number, name: string}[] } An array of objects for the options.
 */
function getHierarchicalTaskOptions(tasksToProcess: Task[], parentId: number | null, level: number): {id: number, name: string}[] {
    let options: {id: number, name: string}[] = [];
    const children = tasksToProcess.filter(t => t.parentId === parentId);

    children.forEach(task => {
        const prefix = '— '.repeat(level);
        options.push({ id: task.id, name: `${prefix}${task.name}` });
        options = options.concat(getHierarchicalTaskOptions(tasksToProcess, task.id, level + 1));
    });

    return options;
}

/**
 * Populates the parent task/category dropdowns in the UI.
 */
function populateCategorySelectors(): void {
    const currentParentId = parentTaskSelect.value;
    
    parentTaskSelect.innerHTML = '<option value="">None (Creates a New Main Task)</option>';
    
    const sortedTasks = [...tasks].sort((a, b) => {
        if (a.startDate < b.startDate) return -1;
        if (a.startDate > b.startDate) return 1;
        return 0;
    });

    const hierarchicalOptions = getHierarchicalTaskOptions(sortedTasks, null, 0);

    hierarchicalOptions.forEach(opt => {
        parentTaskSelect.add(new Option(opt.name, opt.id.toString()));
    });

    parentTaskSelect.value = currentParentId;
}


/**
 * Renders the color coding dropdown control.
 */
function renderColorCodeControls(): void {
    const currentValue = colorCodeSelect.value;
    colorCodeSelect.innerHTML = '';

    const createOption = (value: string, text: string) => {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = text;
        colorCodeSelect.appendChild(option);
    };

    createOption('status', 'Status');
    createOption('resource', 'Resource');
    customFieldsSchema.forEach(field => {
        createOption(field.name, `Custom: ${field.name}`);
    });

    colorCodeSelect.value = currentValue && colorCodeSelect.querySelector(`option[value="${currentValue}"]`) ? currentValue : 'status';
}

/**
 * Renders the dynamic Gantt chart legend.
 */
function renderGanttLegend(): void {
    ganttLegend.innerHTML = '';
    let items: { label: string, color: string }[] = [];
    const uniqueValues = new Set<string>();

    switch (colorCodeCriteria) {
        case 'status':
            items = Object.entries(statusColorMap).map(([label, color]) => ({ label, color }));
            break;
        case 'resource':
            tasks.forEach(task => task.resources?.forEach(r => uniqueValues.add(r)));
            if (tasks.some(t => !t.resources || t.resources.length === 0)) {
                uniqueValues.add('Unassigned');
            }
            items = Array.from(uniqueValues).sort().map(value => ({
                label: value,
                color: generateColorFromString(value)
            }));
            break;
        default: // Custom Field
            tasks.forEach(task => {
                const value = task.customFields?.[colorCodeCriteria];
                uniqueValues.add(value || 'Not Set');
            });
            items = Array.from(uniqueValues).sort().map(value => ({
                label: value,
                color: generateColorFromString(value)
            }));
            break;
    }
    
    if (items.length > 0) {
        items.forEach(item => {
            const legendItem = document.createElement('div');
            legendItem.className = 'legend-item';
            legendItem.innerHTML = `
                <span class="legend-color-swatch" style="background-color: ${item.color};"></span>
                <span class="legend-label">${item.label}</span>
            `;
            ganttLegend.appendChild(legendItem);
        });
    }
}

// --- MEETING MINUTES ---
function renderMeetingMinutesPage(): void {
    meetingMinutesList.innerHTML = '';
    if (meetingMinutes.length === 0) {
        meetingMinutesList.innerHTML = '<p>No meeting minutes have been added yet.</p>';
        return;
    }

    const sortedMinutes = [...meetingMinutes].sort((a, b) => parseDateUTC(b.meetingDate).getTime() - parseDateUTC(a.meetingDate).getTime());
    
    sortedMinutes.forEach(minute => {
        const card = document.createElement('div');
        card.className = 'meeting-minute-card';
        const cleanedMinutes = minute.minutes.replace(/[#*]/g, '');
        const escapedMinutes = cleanedMinutes.replace(/</g, "&lt;").replace(/>/g, "&gt;");

        card.innerHTML = `
            <div class="meeting-minute-card-header">
                <div class="meeting-minute-card-header-dates">
                    <span><strong>Meeting Date:</strong> ${minute.meetingDate}</span>
                    <span><strong>Date Added:</strong> ${minute.addedDate}</span>
                </div>
                <button class="delete-btn-list" data-id="${minute.id}">Delete</button>
            </div>
            <div class="meeting-minute-card-body">
                ${escapedMinutes}
            </div>
        `;
        meetingMinutesList.appendChild(card);
    });
}

function addMeetingMinute(): void {
    if (!meetingDateInput.value || !meetingMinutesTextarea.value.trim()) {
        alert("Please provide both a meeting date and minutes text.");
        return;
    }
    const newMinute: MeetingMinute = {
        id: nextMeetingMinuteId++,
        meetingDate: meetingDateInput.value,
        addedDate: formatDateUTC(new Date()),
        minutes: meetingMinutesTextarea.value.trim(),
    };
    meetingMinutes.push(newMinute);
    renderMeetingMinutesPage();
    addMeetingMinuteForm.reset();
    flatpickrInstances['meeting-date']?.clear();
}

function deleteMeetingMinute(id: number): void {
    const minute = meetingMinutes.find(m => m.id === id);
    if (!minute) return;

    showConfirmation(
        `Are you sure you want to delete the meeting minutes from ${minute.meetingDate}?`,
        () => {
            meetingMinutes = meetingMinutes.filter(m => m.id !== id);
            renderMeetingMinutesPage();
        }
    );
}

// --- ACTION ITEMS ---
function renderActionItemsPage(): void {
    actionItemsTableBody.innerHTML = '';
    if (actionItems.length === 0) {
        actionItemsTableBody.innerHTML = '<tr><td colspan="7">No action items have been added yet.</td></tr>';
        return;
    }

    const sortedItems = [...actionItems].sort((a, b) => parseDateUTC(a.dueDate).getTime() - parseDateUTC(b.dueDate).getTime());
    
    sortedItems.forEach(item => {
        const row = document.createElement('tr');
        const statusClass = item.status.toLowerCase().replace(' ', '-');
        row.innerHTML = `
            <td>${item.what}</td>
            <td>${item.assignedTo.join(', ')}</td>
            <td>${item.assignedBy}</td>
            <td>${item.assignedDate}</td>
            <td>${item.dueDate}</td>
            <td>
                <span class="status-indicator">
                    <span class="status-indicator-dot status-${statusClass}"></span>
                    ${item.status}
                </span>
            </td>
            <td><button class="delete-btn-list" data-id="${item.id}">Delete</button></td>
        `;
        actionItemsTableBody.appendChild(row);
    });
}

function addActionItem(): void {
    const what = actionItemWhatInput.value.trim();
    const assignedToValue = actionItemAssignedToInput.value.trim();
    const assignedBy = actionItemAssignedByInput.value.trim();
    const assignedDate = actionItemAssignedDateInput.value;
    const dueDate = actionItemDueDateInput.value;
    const status = actionItemStatusSelect.value as ActionItemStatus;

    if (!what || !assignedToValue || !assignedBy || !assignedDate || !dueDate) {
        alert("Please fill out all fields for the action item.");
        return;
    }

    const assignedTo = assignedToValue.split(',').map(name => name.trim()).filter(Boolean);

    const newItem: ActionItem = {
        id: nextActionItemId++,
        what,
        assignedTo,
        assignedBy,
        assignedDate,
        dueDate,
        status,
    };
    actionItems.push(newItem);
    renderActionItemsPage();
    addActionItemForm.reset();
    flatpickrInstances['action-item-assigned-date']?.clear();
    flatpickrInstances['action-item-due-date']?.clear();
    actionItemStatusSelect.value = 'Not Started';
}

function deleteActionItem(id: number): void {
    const item = actionItems.find(item => item.id === id);
    if (!item) return;

    showConfirmation(
        `Are you sure you want to delete the action item: "${item.what}"?`,
        () => {
            actionItems = actionItems.filter(item => item.id !== id);
            renderActionItemsPage();
        }
    );
}


// --- PREVIOUS REPORTS ---
function renderPreviousReportsPage(): void {
    if (!previousReportsList) return;
    previousReportsList.innerHTML = '';

    if (previousReports.length === 0) {
        previousReportsList.innerHTML = '<p>No reports have been generated and archived yet.</p>';
        return;
    }

    previousReports.forEach(report => {
        const card = document.createElement('div');
        card.className = 'previous-report-card';
        card.dataset.id = report.id.toString();

        const reportDate = new Date(report.timestamp).toLocaleString(undefined, {
            year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit'
        });

        card.innerHTML = `
            <div class="previous-report-header">
                <span class="previous-report-timestamp">Report from: <strong>${reportDate}</strong></span>
                <button class="delete-btn-list" data-id="${report.id}" aria-label="Delete this report">Delete</button>
            </div>
            <div class="previous-report-body">
                <div class="report-content-wrapper">${report.content}</div>
            </div>
        `;
        previousReportsList.appendChild(card);
    });
}

function deletePreviousReport(id: number): void {
    const report = previousReports.find(r => r.id === id);
    if (!report) return;

    const reportDate = new Date(report.timestamp).toLocaleString();

    showConfirmation(
        `Are you sure you want to delete the report from ${reportDate}? This action cannot be undone.`,
        () => {
            previousReports = previousReports.filter(r => r.id !== id);
            renderPreviousReportsPage();
        }
    );
}


/**
 * Renders the entire application UI.
 */
function render(): void {
    updateAllTaskStatuses();
    generateNotifications();
    populateCategorySelectors();
    renderTaskList();
    renderGanttLegend();
    renderGanttChart();
    renderNotifications();
}

// --- TASK SCHEDULING ---

/**
 * Performs a topological sort of tasks based on dependencies.
 * @returns {Task[] | null} A sorted array of tasks, or null if a cycle is detected.
 */
function topologicalSort(tasksToSort: Task[]): Task[] | null {
    const sorted: Task[] = [];
    const inDegree = new Map<number, number>();
    const adjList = new Map<number, number[]>();
    const taskMap = new Map(tasksToSort.map(t => [t.id, t]));

    for (const task of tasksToSort) {
        inDegree.set(task.id, 0);
        adjList.set(task.id, []);
    }

    for (const task of tasksToSort) {
        for (const depId of task.dependencies || []) {
            if (adjList.has(depId)) { // Ensure dependency exists in the current task list
                adjList.get(depId)!.push(task.id);
                inDegree.set(task.id, (inDegree.get(task.id) || 0) + 1);
            }
        }
    }

    const queue: number[] = [];
    for (const [taskId, degree] of inDegree.entries()) {
        if (degree === 0) {
            queue.push(taskId);
        }
    }

    while (queue.length > 0) {
        const u = queue.shift()!;
        sorted.push(taskMap.get(u)!);

        for (const v of adjList.get(u) || []) {
            inDegree.set(v, (inDegree.get(v) || 1) - 1);
            if (inDegree.get(v) === 0) {
                queue.push(v);
            }
        }
    }

    if (sorted.length !== tasksToSort.length) {
        return null; // Cycle detected
    }
    return sorted;
}


/**
 * Auto-schedules tasks based on their dependencies using a topological sort.
 */
function scheduleTasks(): void {
    const taskMap = new Map(tasks.map(t => [t.id, t]));
    const sortedTasks = topologicalSort(tasks);

    if (!sortedTasks) {
        console.error("Circular dependency detected in project. Auto-scheduling halted.");
        return;
    }
    
    const leafTasks = sortedTasks.filter(t => !tasks.some(other => other.parentId === t.id));

    for (const task of leafTasks) {
        if (!task.dependencies || task.dependencies.length === 0) continue;

        let latestPredecessorEndDate: Date | null = null;

        for (const depId of task.dependencies) {
            const predecessor = taskMap.get(depId);
            if (predecessor) {
                const predecessorEndDate = parseDateUTC(predecessor.endDate);
                if (!latestPredecessorEndDate || predecessorEndDate > latestPredecessorEndDate) {
                    latestPredecessorEndDate = predecessorEndDate;
                }
            }
        }
        
        if (latestPredecessorEndDate) {
            const newStartDate = new Date(latestPredecessorEndDate.getTime());
            newStartDate.setUTCDate(newStartDate.getUTCDate() + 1);
            
            const currentStartDate = parseDateUTC(task.startDate);
            
            if (newStartDate > currentStartDate) {
                const duration = dateDiffInDays(task.startDate, task.endDate);
                const newStartDateStr = formatDateUTC(newStartDate);
                const newEndDateStr = addDays(newStartDateStr, duration);
                
                task.startDate = newStartDateStr;
                task.endDate = newEndDateStr;
            }
        }
    }
}

/**
 * Checks if adding a dependency would create a circular reference.
 * @param taskId The ID of the task being edited.
 * @param dependencyId The ID of the proposed dependency.
 * @returns True if a circular dependency is detected, false otherwise.
 */
function isCircularDependency(taskId: number, dependencyId: number): boolean {
    const dependencyTask = tasks.find(t => t.id === dependencyId);
    if (!dependencyTask || !dependencyTask.dependencies) {
        return false;
    }

    const toCheck: number[] = [...dependencyTask.dependencies];
    const checked = new Set<number>();

    while (toCheck.length > 0) {
        const currentId = toCheck.pop()!;
        if (currentId === taskId) {
            return true;
        }
        if (checked.has(currentId)) {
            continue;
        }
        checked.add(currentId);
        const task = tasks.find(t => t.id === currentId);
        if (task?.dependencies) {
            toCheck.push(...task.dependencies);
        }
    }
    return false;
}

// --- DATA PERSISTENCE (SAVE/LOAD) ---

/**
 * Gathers the complete current project state into an object.
 * @returns {ProjectState} The current project state.
 */
function getProjectStateObject(): ProjectState {
    return {
        projectInfo,
        tasks,
        customFieldsSchema,
        baselineTasks,
        meetingMinutes,
        actionItems,
        previousReports,
        nextMeetingMinuteId,
        nextActionItemId,
        nextPreviousReportId,
        nextTaskId,
        expandedTaskIds: Array.from(expandedTaskIds),
        currentView,
        showBaseline,
        colorCodeCriteria,
        dayWidth,
    };
}

/**
 * Applies a loaded state object to the application.
 * @param {Partial<ProjectState>} state - The state object to apply.
 */
function applyState(state: Partial<ProjectState>): void {
    // Validate state object
    if (!state || typeof state !== 'object' || !Array.isArray(state.tasks)) {
        alert('Invalid project file format.');
        return;
    }

    projectInfo = state.projectInfo || { name: '', description: '', deliverables: '', location: '', doc: '', owner: '', engineeringManager: '', manager: '', engineers: '', startDate: '', endDate: '', budget: '', currency: 'USD' };
    tasks = state.tasks || [];
    customFieldsSchema = state.customFieldsSchema || [];
    baselineTasks = state.baselineTasks || [];
    meetingMinutes = state.meetingMinutes || [];
    actionItems = state.actionItems || [];
    previousReports = state.previousReports || [];
    nextMeetingMinuteId = state.nextMeetingMinuteId || 1;
    nextActionItemId = state.nextActionItemId || 1;
    nextPreviousReportId = state.nextPreviousReportId || 1;
    nextTaskId = state.nextTaskId || 1;
    expandedTaskIds = new Set(state.expandedTaskIds || []);
    currentView = state.currentView || 'day';
    showBaseline = state.showBaseline || false;
    colorCodeCriteria = state.colorCodeCriteria || 'status';
    dayWidth = state.dayWidth || 50;

    // Update UI elements to reflect loaded state
    renderProjectInfo();
    // After rendering the info, update the date pickers with the loaded values
    flatpickrInstances['project-start-date']?.setDate(projectInfo.startDate, false);
    flatpickrInstances['project-end-date']?.setDate(projectInfo.endDate, false);

    showBaselineToggle.checked = showBaseline;
    colorCodeSelect.value = colorCodeCriteria;
    zoomSlider.value = String(dayWidth);
    zoomValue.textContent = `${dayWidth}px`;
    document.querySelectorAll('.view-btn').forEach(btn => {
        btn.classList.toggle('active', btn.getAttribute('data-view') === currentView);
    });

    scheduleTasks();
    updateParentTasks();
    renderCustomFieldsForm(addCustomFieldsContainer);
    renderColorCodeControls();
    render();
    alert('Project loaded successfully.');
}


/**
 * Saves the current state to localStorage for session persistence.
 */
function saveState(): void {
    const state = getProjectStateObject();
    localStorage.setItem('ganttChartState', JSON.stringify(state));
}

/**
 * Initializes the application on page load.
 * Always starts with a new, empty project.
 */
function loadState(): void {
    // Always start with a fresh project. handleNewProject resets all state variables.
    handleNewProject();
}


/**
 * Saves the entire project state to a downloadable JSON file.
 */
function saveProjectToFile(): void {
    const state = getProjectStateObject();
    const formattedJson = JSON.stringify(state, null, 2);
    
    // FIX: Replaced prompt with an auto-generated filename to prevent browser blocking issues.
    const timestamp = new Date().toISOString().replace(/[:.T-]/g, '').slice(0, 14);
    const fileName = `gaseng-pmt-project-${timestamp}.json`;

    triggerDownload(formattedJson, fileName, 'application/json');
}

/**
 * Handles the file input change event to load a project from a JSON file.
 * @param {Event} event The file input change event.
 */
function handleProjectFileLoad(event: Event): void {
    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) {
        return;
    }
    const file = input.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
        try {
            const fileContent = e.target?.result;
            if (typeof fileContent === 'string') {
                const state = JSON.parse(fileContent);
                applyState(state);
            } else {
                throw new Error('File content is not a string.');
            }
        } catch (error) {
            console.error('Error loading project file:', error);
            alert('Failed to load project file. It may be corrupted or in an invalid format.');
        } finally {
            // Reset the input so the user can load the same file again if they wish
            input.value = '';
        }
    };
    
    reader.onerror = () => {
        alert(`Error reading file: ${reader.error}`);
        input.value = '';
    };

    reader.readAsText(file);
}


/**
 * Clears all current project data to start a new project.
 */
function handleNewProject(): void {
    projectInfo = {
        name: '', description: '', deliverables: '', location: '', doc: '', owner: '', engineeringManager: '', manager: '',
        engineers: '', startDate: '', endDate: '', budget: '', currency: 'USD'
    };
    tasks = [];
    baselineTasks = [];
    customFieldsSchema = [];
    meetingMinutes = [];
    actionItems = [];
    previousReports = [];
    nextTaskId = 1;
    nextMeetingMinuteId = 1;
    nextActionItemId = 1;
    nextPreviousReportId = 1;
    editingTaskId = null;
    expandedTaskIds.clear();
    showBaseline = false;
    filters = { name: '', resource: '', status: '' };
    sortCriteria = { key: 'startDate', order: 'asc' };
    colorCodeCriteria = 'status';
    dayWidth = 50;

    // Reset UI control values
    renderProjectInfo();
    // Also clear the flatpickr instances
    Object.values(flatpickrInstances).forEach(instance => instance.clear());

    showBaselineToggle.checked = false;
    filterNameInput.value = '';
    filterResourceInput.value = '';
    filterStatusSelect.value = '';
    sortTasksSelect.value = 'startDate-asc';
    zoomSlider.value = '50';
    zoomValue.textContent = '50px';
    addTaskForm.reset();
    percentCompleteInput.value = '0';

    // Update controls that are dynamically populated
    renderCustomFieldsForm(addCustomFieldsContainer);
    renderColorCodeControls();
    colorCodeSelect.value = 'status';

    render(); // Redraws everything and saves the new empty state
}

// --- CUSTOM FIELDS ---
function renderCustomFieldsForm(container: HTMLElement, values: {[key: string]: string} = {}): void {
    container.innerHTML = '';
    customFieldsSchema.forEach(field => {
        const fieldId = `custom-field-${field.name.replace(/\s+/g, '-')}`;
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        
        const label = document.createElement('label');
        label.setAttribute('for', fieldId);
        label.textContent = field.name;

        const input = document.createElement('input');
        input.type = 'text';
        input.id = fieldId;
        input.name = field.name;
        input.value = values[field.name] || '';
        input.placeholder = `Enter ${field.name}...`;

        formGroup.appendChild(label);
        formGroup.appendChild(input);
        container.appendChild(formGroup);
    });
}

function renderCustomFieldsList(): void {
    customFieldsList.innerHTML = '';
    if (customFieldsSchema.length === 0) {
        customFieldsList.innerHTML = '<p>No custom fields defined.</p>';
        return;
    }
    customFieldsSchema.forEach((field, index) => {
        const fieldEl = document.createElement('div');
        fieldEl.className = 'custom-field-item';
        fieldEl.textContent = field.name;

        const deleteBtn = document.createElement('button');
        deleteBtn.textContent = 'Remove';
        deleteBtn.className = 'delete-btn-small';
        deleteBtn.onclick = () => removeCustomField(index);
        
        fieldEl.appendChild(deleteBtn);
        customFieldsList.appendChild(fieldEl);
    });
}

function addCustomField(name: string): void {
    if (name && !customFieldsSchema.some(f => f.name.toLowerCase() === name.toLowerCase())) {
        customFieldsSchema.push({ name, type: 'text' });
        renderCustomFieldsList();
        renderCustomFieldsForm(addCustomFieldsContainer);
        renderColorCodeControls();
    } else {
        alert('Field name cannot be empty or already exist.');
    }
}

function removeCustomField(index: number): void {
    showConfirmation('Are you sure you want to remove this field? This will delete the corresponding data from all tasks.', () => {
        const fieldNameToRemove = customFieldsSchema[index].name;
        customFieldsSchema.splice(index, 1);
        
        // Remove data from all tasks
        tasks.forEach(task => {
            if (task.customFields && task.customFields[fieldNameToRemove]) {
                delete task.customFields[fieldNameToRemove];
            }
        });
        
        // If the removed field was the color criteria, reset to status
        if(colorCodeCriteria === fieldNameToRemove) {
            colorCodeCriteria = 'status';
        }

        renderCustomFieldsList();
        renderCustomFieldsForm(addCustomFieldsContainer);
        renderColorCodeControls();
        render(); // Full re-render to update task list
    });
}


// --- EXPORT FUNCTIONALITY ---

/**
 * Triggers a file download in the browser.
 * @param {string} content - The file content.
 * @param {string} fileName - The desired file name.
 * @param {string} mimeType - The MIME type of the file.
 */
function triggerDownload(content: string, fileName: string, mimeType: string): void {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

/**
 * Escapes a cell for CSV format, handling commas and quotes.
 * @param {any} cellData - The data for the cell.
 * @returns {string} The escaped CSV cell string.
 */
function escapeCsvCell(cellData: any): string {
    const stringData = String(cellData ?? '');
    if (stringData.includes(',') || stringData.includes('"') || stringData.includes('\n')) {
        return `"${stringData.replace(/"/g, '""')}"`;
    }
    return stringData;
}

/**
 * Exports the tasks as a CSV file.
 */
function exportCSV(): void {
    if (tasks.length === 0) {
        alert('No tasks to export.');
        return;
    }
    const customFieldNames = customFieldsSchema.map(f => f.name);
    const headers = ['id', 'name', 'description', 'startDate', 'endDate', 'parentId', 'percentComplete', 'status', 'resources', 'notes', 'dependencies', ...customFieldNames];
    const csvRows = [headers.join(',')];

    tasks.forEach(task => {
        const effectiveStatus = task.status === 'auto' ? getCalculatedStatus(task) : task.status;
        const customFieldValues = customFieldNames.map(fieldName => task.customFields?.[fieldName] || '');
        const row = [
            task.id,
            task.name,
            task.description || '',
            task.startDate,
            task.endDate,
            task.parentId ?? '',
            task.percentComplete,
            effectiveStatus,
            task.resources?.join(';') || '', // Use semicolon for multi-value cells
            task.notes || '',
            task.dependencies?.join(';') || '',
            ...customFieldValues
        ].map(escapeCsvCell);
        csvRows.push(row.join(','));
    });

    triggerDownload(csvRows.join('\n'), 'gaseng-pmt-tasks.csv', 'text/csv;charset=utf-8;');
}


// --- BASELINE ---

/**
 * Sets the current task schedule as the baseline.
 */
function setBaseline(): void {
    showConfirmation('Are you sure you want to set the current schedule as the new baseline? This will overwrite any existing baseline.', () => {
        baselineTasks = JSON.parse(JSON.stringify(tasks)); // Deep copy
        alert('Baseline has been set successfully.');
        render();
    });
}

// --- NOTIFICATIONS ---

/**
 * Generates notifications for overdue tasks and upcoming deadlines.
 */
function generateNotifications(): void {
    notifications = [];
    const now = new Date();
    const today = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));

    const threeDaysFromNow = new Date(today);
    threeDaysFromNow.setUTCDate(today.getUTCDate() + 3);

    tasks.forEach(task => {
        if (task.percentComplete < 100) {
            const endDate = parseDateUTC(task.endDate);
            // Check for overdue tasks
            if (endDate < today) {
                notifications.push({
                    id: `overdue-${task.id}`,
                    message: `<strong>Overdue:</strong> "${task.name}" was due on ${task.endDate}.`,
                    taskId: task.id
                });
            } 
            // Check for tasks due soon
            else if (endDate <= threeDaysFromNow) {
                notifications.push({
                    id: `due-soon-${task.id}`,
                    message: `<strong>Due Soon:</strong> "${task.name}" is due on ${task.endDate}.`,
                    taskId: task.id
                });
            }
        }
    });
}

/**
 * Renders the notifications in the UI.
 */
function renderNotifications(): void {
    if (notifications.length > 0) {
        notificationCount.textContent = notifications.length.toString();
        notificationCount.style.display = 'block';
    } else {
        notificationCount.style.display = 'none';
    }

    notificationList.innerHTML = '';
    if (notifications.length === 0) {
        notificationList.innerHTML = '<li class="no-notifications">No new notifications.</li>';
        return;
    }

    notifications.forEach(notification => {
        const li = document.createElement('li');
        li.innerHTML = notification.message;
        li.onclick = () => {
            const taskElement = document.getElementById(`task-list-item-${notification.taskId}`);
            taskElement?.scrollIntoView({ behavior: 'smooth', block: 'center' });
            taskElement?.classList.add('highlight');
            setTimeout(() => taskElement?.classList.remove('highlight'), 2000);
            notificationPanel.style.display = 'none';
        };
        notificationList.appendChild(li);
    });
}

// --- REPORTING ---

/**
 * Recursively builds the HTML for the detailed task list in the report.
 * @param taskId The ID of the task to render.
 * @param level The current indentation level.
 * @returns An HTML string representing the task and its children.
 */
function buildTaskDetailsHtml(taskId: number, level: number): string {
    const task = tasks.find(t => t.id === taskId);
    if (!task) return '';

    const effectiveStatus = task.status === 'auto' ? getCalculatedStatus(task) : task.status;
    const duration = dateDiffInDays(task.startDate, task.endDate) + 1;
    const resources = task.resources?.length ? task.resources.join(', ') : 'N/A';
    const statusClass = effectiveStatus.toLowerCase().replace(' ', '-');
    
    let customFieldsHtml = '';
    if (task.customFields && Object.keys(task.customFields).length > 0) {
        for (const [key, value] of Object.entries(task.customFields)) {
            if(value) {
                customFieldsHtml += `<div><strong>${key}:</strong> <span>${value}</span></div>`;
            }
        }
    }

    let descriptionHtml = '';
    if (task.description) {
        const escapedDescription = task.description.replace(/</g, "&lt;").replace(/>/g, "&gt;");
        descriptionHtml = `
            <div class="report-task-description">
                <strong>Description:</strong>
                <p>${escapedDescription}</p>
            </div>
        `;
    }

    let notesHtml = '';
    if (task.notes) {
        // Basic HTML escaping for notes
        const escapedNotes = task.notes.replace(/</g, "&lt;").replace(/>/g, "&gt;");
        notesHtml = `
            <div class="report-task-notes">
                <strong>Notes:</strong>
                <p>${escapedNotes}</p>
            </div>
        `;
    }

    let html = `
        <div class="report-task-item-detail" style="margin-left: ${level * 30}px;">
            <div class="report-task-header ${statusClass}">
                ${task.name}
            </div>
            <div class="report-task-body">
                <div><strong>Status:</strong> <span class="${statusClass}">${effectiveStatus}</span></div>
                <div><strong>Completion:</strong> <span>${task.percentComplete}%</span></div>
                <div><strong>Dates:</strong> <span>${task.startDate} to ${task.endDate}</span></div>
                <div><strong>Duration:</strong> <span>${duration} day(s)</span></div>
                <div><strong>Resources:</strong> <span>${resources}</span></div>
                ${customFieldsHtml}
                ${descriptionHtml}
                ${notesHtml}
            </div>
        </div>
    `;

    const children = tasks.filter(t => t.parentId === taskId);
    children.sort((a, b) => parseDateUTC(a.startDate).getTime() - parseDateUTC(b.startDate).getTime());
    children.forEach(child => {
        html += buildTaskDetailsHtml(child.id, level + 1);
    });

    return html;
}


/**
 * Calculates all metrics and populates the report page.
 */
function generateReport(): void {
    const N_A = 'N/A';
    
    // Timestamp
    const timestampEl = document.getElementById('report-timestamp');
    if (timestampEl) {
        const timestampFormat: Intl.DateTimeFormatOptions = { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' };
        timestampEl.textContent = new Date().toLocaleDateString(undefined, timestampFormat);
    }

    // 0. Project Information
    (document.getElementById('report-project-name') as HTMLElement).textContent = projectInfo.name || 'Untitled Project';
    (document.getElementById('report-project-description') as HTMLElement).textContent = projectInfo.description || N_A;
    (document.getElementById('report-project-deliverables') as HTMLElement).textContent = projectInfo.deliverables || N_A;
    (document.getElementById('report-project-owner') as HTMLElement).textContent = projectInfo.owner || N_A;
    (document.getElementById('report-project-engineering-manager') as HTMLElement).textContent = projectInfo.engineeringManager || N_A;
    (document.getElementById('report-project-manager') as HTMLElement).textContent = projectInfo.manager || N_A;
    (document.getElementById('report-project-start-date') as HTMLElement).textContent = projectInfo.startDate || N_A;
    (document.getElementById('report-project-end-date') as HTMLElement).textContent = projectInfo.endDate || N_A;
    (document.getElementById('report-project-location') as HTMLElement).textContent = projectInfo.location || N_A;
    (document.getElementById('report-project-doc') as HTMLElement).textContent = projectInfo.doc || N_A;
    (document.getElementById('report-project-engineers') as HTMLElement).textContent = projectInfo.engineers || N_A;
    const budgetValue = projectInfo.budget;
    const currency = projectInfo.currency || 'USD';
    let formattedBudget = N_A;
    if (typeof budgetValue === 'number' && !isNaN(budgetValue)) {
        formattedBudget = new Intl.NumberFormat(undefined, {
            style: 'currency',
            currency: currency,
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(budgetValue);
    } else if (budgetValue) {
        formattedBudget = `${budgetValue} ${currency}`;
    }
    (document.getElementById('report-project-budget') as HTMLElement).textContent = formattedBudget;

    // 1. Overall Progress (weighted by duration of top-level tasks)
    let totalWeightedCompletion = 0;
    let totalDuration = 0;
    tasks.filter(t => t.parentId === null).forEach(task => {
        const duration = dateDiffInDays(task.startDate, task.endDate) + 1;
        if (duration > 0) {
            totalWeightedCompletion += (task.percentComplete || 0) * duration;
            totalDuration += duration;
        }
    });
    const overallProgress = totalDuration > 0 ? Math.round(totalWeightedCompletion / totalDuration) : 0;
    const progressBarFill = reportPage.querySelector('#report-overall-progress .progress-bar-fill') as HTMLDivElement;
    const progressBarLabel = reportPage.querySelector('#report-overall-progress .progress-bar-label') as HTMLSpanElement;
    if (progressBarFill && progressBarLabel) {
        progressBarFill.style.width = `${overallProgress}%`;
        progressBarLabel.textContent = `${overallProgress}% Complete`;
    }

    const now = new Date();
    const today = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));

    // Helper to render lists of tasks in the report
    const renderReportTaskList = (ul: HTMLUListElement, taskList: Task[], emptyMessage: string) => {
        ul.innerHTML = '';
        if (taskList.length === 0) {
            ul.innerHTML = `<li>${emptyMessage}</li>`;
            return;
        }
        taskList.forEach(task => {
            const li = document.createElement('li');
            const effectiveStatus = task.status === 'auto' ? getCalculatedStatus(task) : task.status;
            li.innerHTML = `
                <span class="report-task-name">${task.name}</span>
                <span class="report-task-details">
                    (Due: ${task.endDate} - 
                    <span class="status-${effectiveStatus.toLowerCase().replace(' ', '-')}">${effectiveStatus}</span>)
                </span>
            `;
            ul.appendChild(li);
        });
    };
    
    // 2. At-Risk & Overdue Tasks
    const overdueTasks = tasks.filter(t => parseDateUTC(t.endDate) < today && t.percentComplete < 100);
    const atRiskTasks = tasks.filter(t => getCalculatedStatus(t) === 'At Risk' && !overdueTasks.includes(t));
    const atRiskAndOverdue = [...overdueTasks, ...atRiskTasks];
    const atRiskUl = reportPage.querySelector('#report-at-risk .report-task-list') as HTMLUListElement;
    if(atRiskUl) renderReportTaskList(atRiskUl, atRiskAndOverdue, 'No at-risk or overdue tasks.');

    // 3. Upcoming Deadlines
    const upcomingDays = parseInt(upcomingDaysInput.value, 10) || 7;
    const deadlineDate = new Date(today.getTime());
    deadlineDate.setUTCDate(today.getUTCDate() + upcomingDays);
    const upcomingTasks = tasks.filter(t => {
        const endDate = parseDateUTC(t.endDate);
        return endDate >= today && endDate <= deadlineDate && t.percentComplete < 100;
    });
    const upcomingUl = reportPage.querySelector('#report-upcoming .report-task-list') as HTMLUListElement;
    if(upcomingUl) renderReportTaskList(upcomingUl, upcomingTasks, `No upcoming deadlines in the next ${upcomingDays} days.`);

    // 4. Key Milestones (tasks with 0 duration)
    const milestones = tasks.filter(t => dateDiffInDays(t.startDate, t.endDate) === 0);
    const milestonesUl = reportPage.querySelector('#report-milestones .report-task-list') as HTMLUListElement;
    if (milestonesUl) {
         milestonesUl.innerHTML = '';
        if (milestones.length === 0) {
            milestonesUl.innerHTML = `<li>No milestones defined.</li>`;
        } else {
            milestones.forEach(task => {
                const li = document.createElement('li');
                const isComplete = task.percentComplete === 100;
                const isOverdue = parseDateUTC(task.endDate) < today && !isComplete;
                const status = isComplete ? 'Complete' : (isOverdue ? 'Delayed' : 'Pending');
                const statusClass = status.toLowerCase();
                 li.innerHTML = `
                    <span class="report-task-name">${task.name}</span>
                    <span class="report-task-details">
                        (Date: ${task.endDate} - 
                        <span class="status-${statusClass}">${status}</span>)
                    </span>
                `;
                milestonesUl.appendChild(li);
            });
        }
    }
    
    // 5. Resource Utilization
    const resourceStats: { [key: string]: { totalTasks: number, completedTasks: number, totalDuration: number } } = {};
    tasks.forEach(task => {
        const duration = dateDiffInDays(task.startDate, task.endDate) + 1;
        const resources = task.resources?.length ? task.resources : ['Unassigned'];
        resources.forEach(resource => {
            if (!resourceStats[resource]) {
                resourceStats[resource] = { totalTasks: 0, completedTasks: 0, totalDuration: 0 };
            }
            resourceStats[resource].totalTasks++;
            resourceStats[resource].totalDuration += duration;
            if (task.percentComplete === 100) {
                resourceStats[resource].completedTasks++;
            }
        });
    });
    const resourceTableBody = reportPage.querySelector('#resource-utilization-table tbody') as HTMLTableSectionElement;
    if (resourceTableBody) {
        resourceTableBody.innerHTML = '';
        const sortedResources = Object.entries(resourceStats).sort((a, b) => a[0].localeCompare(b[0]));
        if (sortedResources.length === 0) {
            resourceTableBody.innerHTML = '<tr><td colspan="4">No resources assigned to tasks.</td></tr>';
        } else {
            sortedResources.forEach(([resource, stats]) => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${resource}</td>
                    <td>${stats.totalTasks}</td>
                    <td>${stats.completedTasks}</td>
                    <td>${stats.totalDuration} days</td>
                `;
                resourceTableBody.appendChild(tr);
            });
        }
    }

    // 6. Detailed Task List
    const taskDetailsContainer = document.getElementById('report-task-details-list');
    if (taskDetailsContainer) {
        let detailsHtml = '';
        const topLevelTasks = tasks.filter(task => task.parentId === null).sort((a, b) => parseDateUTC(a.startDate).getTime() - parseDateUTC(b.startDate).getTime());
        if (topLevelTasks.length > 0) {
            topLevelTasks.forEach(task => {
                detailsHtml += buildTaskDetailsHtml(task.id, 0);
            });
        } else {
            detailsHtml = '<p>No tasks in the project.</p>';
        }
        taskDetailsContainer.innerHTML = detailsHtml;
    }

    // 7. Action Items
    const actionItemsTable = reportPage.querySelector('#action-items-table tbody') as HTMLTableSectionElement;
    if (actionItemsTable) {
        actionItemsTable.innerHTML = '';
        if (actionItems.length === 0) {
            actionItemsTable.innerHTML = '<tr><td colspan="6">No action items recorded.</td></tr>';
        } else {
            const sortedItems = [...actionItems].sort((a, b) => parseDateUTC(a.dueDate).getTime() - parseDateUTC(b.dueDate).getTime());
            sortedItems.forEach(item => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${item.what}</td>
                    <td>${item.assignedTo.join(', ')}</td>
                    <td>${item.assignedBy}</td>
                    <td>${item.assignedDate}</td>
                    <td>${item.dueDate}</td>
                    <td>${item.status}</td>
                `;
                actionItemsTable.appendChild(tr);
            });
        }
    }

    // 8. Meeting Minutes
    const reportMeetingMinutesList = document.getElementById('report-meeting-minutes-list');
    if (reportMeetingMinutesList) {
        reportMeetingMinutesList.innerHTML = '';
        if (meetingMinutes.length === 0) {
            reportMeetingMinutesList.innerHTML = '<p>No meeting minutes recorded.</p>';
        } else {
            const sortedMinutes = [...meetingMinutes].sort((a, b) => parseDateUTC(a.meetingDate).getTime() - parseDateUTC(b.meetingDate).getTime());
            sortedMinutes.forEach(minute => {
                const minuteEl = document.createElement('div');
                minuteEl.className = 'report-minute-item';
                const cleanedMinutes = minute.minutes.replace(/[#*]/g, '');
                const escapedMinutes = cleanedMinutes.replace(/</g, "&lt;").replace(/>/g, "&gt;");
                minuteEl.innerHTML = `
                    <div class="report-minute-header">Meeting Date: ${minute.meetingDate} (Added: ${minute.addedDate})</div>
                    <div class="report-minute-content">${escapedMinutes}</div>
                `;
                reportMeetingMinutesList.appendChild(minuteEl);
            });
        }
    }
}

/**
 * Captures the current state of the report, including the written summary,
 * and saves it to the previous reports archive.
 * @param {boolean} showAlert - If true, displays a success alert to the user.
 */
function archiveCurrentReport(showAlert: boolean = false): void {
    const reportContentEl = document.getElementById('report-content');
    const summaryTextarea = document.getElementById('report-written-summary-textarea') as HTMLTextAreaElement;
    if (!reportContentEl || !summaryTextarea) return;

    // Clone the report content to avoid modifying the live DOM
    const reportClone = reportContentEl.cloneNode(true) as HTMLElement;
    
    // Find the textarea and actions container within the clone
    const summaryTextareaClone = reportClone.querySelector('#report-written-summary-textarea') as HTMLTextAreaElement;
    const actionsClone = reportClone.querySelector('.report-section-actions');
    
    if (summaryTextareaClone && summaryTextareaClone.parentElement) {
        // Create a display div with the summary text
        const summaryDisplay = document.createElement('div');
        summaryDisplay.className = 'report-multiline-text'; // Use existing styles for consistency
        summaryDisplay.textContent = summaryTextarea.value || 'No summary was provided for this report.';
        
        // Replace the textarea in the clone with the static display div
        summaryTextareaClone.parentElement.replaceChild(summaryDisplay, summaryTextareaClone);
    }

    // Remove the actions container from the cloned/archived version
    if (actionsClone) {
        actionsClone.remove();
    }

    // Now get the innerHTML from the modified clone, which includes the summary as static text
    const reportHtml = reportClone.innerHTML;

    const newReport: PreviousReport = {
        id: nextPreviousReportId++,
        timestamp: new Date().toISOString(),
        content: reportHtml
    };
    previousReports.unshift(newReport); // Add to the top of the list

    if (showAlert) {
        alert('Report has been saved to the "Previous Reports" archive.');
    }
}


/**
 * Exports the generated report to a high-quality, multi-page PDF file.
 */
async function exportReportToPDF(): Promise<void> {
    // Archive the current report before generating PDF. Do not show a user alert here.
    archiveCurrentReport(false);

    const { jsPDF } = jspdf;
    exportPdfBtn.textContent = 'Generating...';
    exportPdfBtn.disabled = true;

    try {
        const pdf = new jsPDF({
            orientation: 'portrait',
            unit: 'mm',
            format: 'a4'
        });

        // --- PDF Generation Constants ---
        const MARGIN = { TOP: 20, LEFT: 15, RIGHT: 15, BOTTOM: 20 };
        const PAGE_WIDTH = pdf.internal.pageSize.getWidth();
        const CONTENT_WIDTH = PAGE_WIDTH - MARGIN.LEFT - MARGIN.RIGHT;
        const PAGE_HEIGHT = pdf.internal.pageSize.getHeight();
        let currentY = MARGIN.TOP;

        // --- Helper Functions ---
        const checkAndAddPage = (neededHeight: number) => {
            if (currentY + neededHeight > PAGE_HEIGHT - MARGIN.BOTTOM) {
                pdf.addPage();
                currentY = MARGIN.TOP;
            }
        };

        const drawSectionTitle = (title: string) => {
            checkAndAddPage(15);
            pdf.setFontSize(14);
            pdf.setFont(undefined, 'bold');
            pdf.setTextColor(40, 40, 40);
            pdf.text(title, MARGIN.LEFT, currentY);
            currentY += 8;
            pdf.setDrawColor(200, 200, 200);
            pdf.line(MARGIN.LEFT, currentY, MARGIN.LEFT + CONTENT_WIDTH, currentY);
            currentY += 8;
        };
        
        const drawWrappedText = (text: string, x: number, y: number, width: number, options: any = {}) => {
            const lines = pdf.splitTextToSize(text || 'N/A', width);
            pdf.text(lines, x, y, options);
            return lines.length * 5; // Estimate height
        };

        // --- Cover Page ---
        pdf.setFontSize(28);
        pdf.setFont(undefined, 'bold');
        pdf.setTextColor(40, 40, 40);
        pdf.text('Project Status Report', PAGE_WIDTH / 2, PAGE_HEIGHT / 3, { align: 'center' });
        pdf.setFontSize(18);
        pdf.setFont(undefined, 'normal');
        pdf.text(projectInfo.name || 'Untitled Project', PAGE_WIDTH / 2, PAGE_HEIGHT / 3 + 20, { align: 'center' });
        pdf.setFontSize(12);
        const reportDate = `Generated: ${new Date().toLocaleString()}`;
        pdf.text(reportDate, PAGE_WIDTH / 2, PAGE_HEIGHT / 3 + 35, { align: 'center' });

        // --- Start Content ---
        pdf.addPage();
        currentY = MARGIN.TOP;

        // --- Project Summary Section ---
        drawSectionTitle("Project Summary");
        pdf.setFontSize(10);
        pdf.setFont(undefined, 'normal');
        pdf.setTextColor(80, 80, 80);
        
        let summaryGridY = currentY;
        const col1X = MARGIN.LEFT;
        const col2X = MARGIN.LEFT + CONTENT_WIDTH / 2;
        const labelOffset = 40;

        // Row 1: Owner & Engineering Manager
        pdf.setFont(undefined, 'bold');
        pdf.text('Owner:', col1X, summaryGridY);
        pdf.setFont(undefined, 'normal');
        pdf.text(projectInfo.owner || 'N/A', col1X + labelOffset, summaryGridY);

        pdf.setFont(undefined, 'bold');
        pdf.text('Engineering Manager:', col2X, summaryGridY);
        pdf.setFont(undefined, 'normal');
        pdf.text(projectInfo.engineeringManager || 'N/A', col2X + labelOffset, summaryGridY);
        summaryGridY += 7;

        // Row 2: Project Manager & Start Date
        pdf.setFont(undefined, 'bold');
        pdf.text('Project Manager:', col1X, summaryGridY);
        pdf.setFont(undefined, 'normal');
        pdf.text(projectInfo.manager || 'N/A', col1X + labelOffset, summaryGridY);

        pdf.setFont(undefined, 'bold');
        pdf.text('Start Date:', col2X, summaryGridY);
        pdf.setFont(undefined, 'normal');
        pdf.text(projectInfo.startDate || 'N/A', col2X + labelOffset, summaryGridY);
        summaryGridY += 7;

        // Row 3: End Date & Location
        pdf.setFont(undefined, 'bold');
        pdf.text('End Date:', col1X, summaryGridY);
        pdf.setFont(undefined, 'normal');
        pdf.text(projectInfo.endDate || 'N/A', col1X + labelOffset, summaryGridY);

        pdf.setFont(undefined, 'bold');
        pdf.text('Location:', col2X, summaryGridY);
        pdf.setFont(undefined, 'normal');
        pdf.text(projectInfo.location || 'N/A', col2X + labelOffset, summaryGridY);
        summaryGridY += 7;

        // Row 4: DOC & Budget
        pdf.setFont(undefined, 'bold');
        pdf.text('DOC:', col1X, summaryGridY);
        pdf.setFont(undefined, 'normal');
        pdf.text(projectInfo.doc || 'N/A', col1X + labelOffset, summaryGridY);

        const budgetValue = projectInfo.budget;
        const currency = projectInfo.currency || 'USD';
        let formattedBudget = 'N/A';
        if (typeof budgetValue === 'number' && !isNaN(budgetValue)) {
            formattedBudget = new Intl.NumberFormat(undefined, {
                style: 'currency',
                currency: currency,
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            }).format(budgetValue);
        } else if (budgetValue) {
            formattedBudget = `${budgetValue} ${currency}`;
        }
        pdf.setFont(undefined, 'bold');
        pdf.text('Budget:', col2X, summaryGridY);
        pdf.setFont(undefined, 'normal');
        pdf.text(formattedBudget, col2X + labelOffset, summaryGridY);
        summaryGridY += 10;
        currentY = summaryGridY;

        // Engineers (full width)
        pdf.setFont(undefined, 'bold');
        pdf.text('Project Engineer(s):', MARGIN.LEFT, currentY);
        pdf.setFont(undefined, 'normal');
        pdf.setTextColor(80, 80, 80);
        currentY += 6;
        currentY += drawWrappedText(projectInfo.engineers || 'N/A', MARGIN.LEFT, currentY, CONTENT_WIDTH);
        currentY += 5;

        // Description
        pdf.setFont(undefined, 'bold');
        pdf.text('Description:', MARGIN.LEFT, currentY);
        pdf.setFont(undefined, 'normal');
        pdf.setTextColor(80, 80, 80);
        currentY += 6;
        currentY += drawWrappedText(projectInfo.description || 'N/A', MARGIN.LEFT, currentY, CONTENT_WIDTH);
        currentY += 5;

        // Deliverables
        pdf.setFont(undefined, 'bold');
        pdf.text('Deliverables:', MARGIN.LEFT, currentY);
        pdf.setFont(undefined, 'normal');
        pdf.setTextColor(80, 80, 80);
        currentY += 6;
        currentY += drawWrappedText(projectInfo.deliverables || 'N/A', MARGIN.LEFT, currentY, CONTENT_WIDTH);
        currentY += 10;
        
        // --- Written Summary Section ---
        const writtenSummary = (document.getElementById('report-written-summary-textarea') as HTMLTextAreaElement)?.value.trim();
        if (writtenSummary) {
            drawSectionTitle("Written Summary");
            pdf.setFontSize(10);
            pdf.setFont(undefined, 'normal');
            pdf.setTextColor(80, 80, 80);
            currentY += drawWrappedText(writtenSummary, MARGIN.LEFT, currentY, CONTENT_WIDTH);
            currentY += 5;
        }
        
        // --- Overall Progress ---
        drawSectionTitle("Overall Progress");
        const totalWeightedCompletion = tasks.filter(t => t.parentId === null).reduce((acc, task) => {
            const duration = dateDiffInDays(task.startDate, task.endDate) + 1;
            return acc + (task.percentComplete || 0) * (duration > 0 ? duration : 0);
        }, 0);
        const totalDuration = tasks.filter(t => t.parentId === null).reduce((acc, task) => {
            const duration = dateDiffInDays(task.startDate, task.endDate) + 1;
            return acc + (duration > 0 ? duration : 0);
        }, 0);
        const overallProgress = totalDuration > 0 ? Math.round(totalWeightedCompletion / totalDuration) : 0;
        
        pdf.setDrawColor(220, 220, 220);
        pdf.setFillColor(240, 240, 240);
        pdf.rect(MARGIN.LEFT, currentY, CONTENT_WIDTH, 8, 'FD');
        pdf.setFillColor(statusColorMap['On Track']);
        pdf.rect(MARGIN.LEFT, currentY, CONTENT_WIDTH * (overallProgress / 100), 8, 'F');
        pdf.setFontSize(12);
        pdf.setFont(undefined, 'bold');
        pdf.setTextColor(40, 40, 40);
        pdf.text(`${overallProgress}% Complete`, MARGIN.LEFT + CONTENT_WIDTH + 3, currentY + 6);
        currentY += 15;


        // --- Task Lists ---
        const now = new Date();
        const today = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
        const overdueTasks = tasks.filter(t => parseDateUTC(t.endDate) < today && t.percentComplete < 100);
        const atRiskTasks = tasks.filter(t => getCalculatedStatus(t) === 'At Risk' && !overdueTasks.includes(t));
        
        const renderTaskTable = (title: string, taskList: Task[]) => {
            drawSectionTitle(title);
            if (taskList.length === 0) {
                pdf.setFontSize(10);
                pdf.setFont(undefined, 'italic');
                pdf.setTextColor(150, 150, 150);
                pdf.text("No tasks in this category.", MARGIN.LEFT, currentY);
                currentY += 10;
                return;
            }
             (pdf as any).autoTable({
                startY: currentY,
                head: [['Task Name', 'Status', 'Due Date', '% Complete']],
                body: taskList.map(t => {
                    const status = getCalculatedStatus(t);
                    return [t.name, status, t.endDate, `${t.percentComplete}%`];
                }),
                theme: 'grid',
                headStyles: { fillColor: [60, 70, 80] },
                didDrawPage: (data: any) => { currentY = data.cursor.y + 5; },
                margin: { left: MARGIN.LEFT, right: MARGIN.RIGHT }
            });
            currentY = (pdf as any).lastAutoTable.finalY + 10;
        };

        renderTaskTable("At-Risk & Overdue Tasks", [...overdueTasks, ...atRiskTasks]);

        // --- Task Details ---
        drawSectionTitle("Task Details");
        
        const renderPdfTaskDetailRecursive = (taskId: number, level: number) => {
            const task = tasks.find(t => t.id === taskId);
            if (!task) return;
            
            // FIX: Use the correct variable 'task' instead of 't' which is out of scope.
            const effectiveStatus = getCalculatedStatus(task);
            const statusColor = statusColorMap[effectiveStatus];

            checkAndAddPage(25); // Minimum height for a task block
            
            pdf.setFillColor(248, 249, 250);
            pdf.setDrawColor(222, 226, 230);
            
            const indent = level * 10;
            pdf.setFontSize(11);
            pdf.setFont(undefined, 'bold');
            pdf.setTextColor(40, 40, 40);
            pdf.text(task.name, MARGIN.LEFT + indent, currentY);

            pdf.setFillColor(statusColor);
            pdf.circle(MARGIN.LEFT + CONTENT_WIDTH - 5, currentY - 1.5, 2, 'F');
            pdf.setFontSize(9);
            pdf.setFont(undefined, 'normal');
            pdf.setTextColor(80, 80, 80);
            pdf.text(effectiveStatus, MARGIN.LEFT + CONTENT_WIDTH - 10, currentY, { align: 'right' });

            currentY += 6;
            
            pdf.text(`Dates: ${task.startDate} to ${task.endDate}`, MARGIN.LEFT + indent, currentY);
            pdf.text(`Progress: ${task.percentComplete}%`, MARGIN.LEFT + CONTENT_WIDTH / 2, currentY);
            currentY += 5;

            if (task.description) {
                currentY += drawWrappedText(`Description: ${task.description}`, MARGIN.LEFT + indent, currentY, CONTENT_WIDTH - indent);
                 currentY += 3;
            }
            
            if (task.notes) {
                currentY += drawWrappedText(`Notes: ${task.notes}`, MARGIN.LEFT + indent, currentY, CONTENT_WIDTH - indent);
            }
            currentY += 8;

            const children = tasks.filter(t => t.parentId === taskId);
            children.forEach(child => renderPdfTaskDetailRecursive(child.id, level + 1));
        };
        
        const topLevelTasks = tasks.filter(t => t.parentId === null);
        topLevelTasks.forEach(task => renderPdfTaskDetailRecursive(task.id, 0));
        
        // --- Action Items Section ---
        if (actionItems.length > 0) {
            drawSectionTitle("Action Items");
            const sortedItems = [...actionItems].sort((a, b) => parseDateUTC(a.dueDate).getTime() - parseDateUTC(b.dueDate).getTime());
             (pdf as any).autoTable({
                startY: currentY,
                head: [['Action Item', 'Assigned To', 'Assigned By', 'Assigned', 'Due', 'Status']],
                body: sortedItems.map(item => [item.what, item.assignedTo.join(', '), item.assignedBy, item.assignedDate, item.dueDate, item.status]),
                theme: 'grid',
                headStyles: { fillColor: [60, 70, 80] },
                didDrawPage: (data: any) => { currentY = data.cursor.y + 5; },
                margin: { left: MARGIN.LEFT, right: MARGIN.RIGHT }
            });
            currentY = (pdf as any).lastAutoTable.finalY + 10;
        }

        // --- Meeting Minutes Section ---
        if (meetingMinutes.length > 0) {
            drawSectionTitle("Meeting Minutes");
            const sortedMinutes = [...meetingMinutes].sort((a, b) => parseDateUTC(a.meetingDate).getTime() - parseDateUTC(b.meetingDate).getTime());
            sortedMinutes.forEach(minute => {
                const cleanedMinutes = minute.minutes.replace(/[#*]/g, '');
                const textContent = `Meeting Date: ${minute.meetingDate} (Added: ${minute.addedDate})\n\n${cleanedMinutes}`;
                const lines = pdf.splitTextToSize(textContent, CONTENT_WIDTH);
                const neededHeight = (lines.length * 5) + 10; // 5mm per line + padding
                checkAndAddPage(neededHeight);

                pdf.setFontSize(10);
                pdf.setFont(undefined, 'bold');
                pdf.text(`Meeting Date: ${minute.meetingDate}`, MARGIN.LEFT, currentY);
                pdf.setFont(undefined, 'normal');
                pdf.text(`(Added: ${minute.addedDate})`, MARGIN.LEFT + 50, currentY);
                currentY += 8;

                pdf.setFontSize(10);
                pdf.setTextColor(80, 80, 80);
                const minutesHeight = drawWrappedText(cleanedMinutes, MARGIN.LEFT, currentY, CONTENT_WIDTH);
                currentY += minutesHeight + 8;
            });
        }


        // --- Page Numbers ---
        const pageCount = (pdf as any).internal.getNumberOfPages();
        for (let i = 1; i <= pageCount; i++) {
            pdf.setPage(i);
            pdf.setFontSize(9);
            pdf.setTextColor(150, 150, 150);
            const pageNumText = `Page ${i} of ${pageCount}`;
            pdf.text(pageNumText, PAGE_WIDTH / 2, PAGE_HEIGHT - MARGIN.BOTTOM / 2, { align: 'center' });
        }
        
        const fileName = `${projectInfo.name || 'gaseng-pmt-project'}-report.pdf`.replace(/[^a-z0-9]/gi, '_').toLowerCase();
        pdf.save(fileName);

    } catch (error) {
        console.error("Failed to export PDF:", error);
        alert("Could not export PDF. See console for details.");
    } finally {
        exportPdfBtn.textContent = 'Export to PDF';
        exportPdfBtn.disabled = false;
    }
}


/**
 * Opens the PDF export modal for the Gantt chart.
 */
function openPdfExportModal(): void {
    pdfExportModal.style.display = 'flex';
}

/**
 * Closes the PDF export modal for the Gantt chart.
 */
function closePdfExportModal(): void {
    pdfExportModal.style.display = 'none';
}

/**
 * Exports the current Gantt chart view to a PDF file with a custom page size
 * that matches the content to ensure high quality and no distortion.
 */
async function exportGanttToPdf(): Promise<void> {
    const exportContainer = document.querySelector('.gantt-scroll-wrapper') as HTMLElement;
    
    if (!exportContainer || !ganttChartArea.classList.contains('has-content')) {
        alert('Chart is empty. Nothing to export.');
        closePdfExportModal();
        return;
    }

    generateGanttPdfBtn.textContent = 'Generating...';
    generateGanttPdfBtn.disabled = true;
    
    const originalScrollLeft = exportContainer.scrollLeft;
    const originalScrollTop = exportContainer.scrollTop;
    exportContainer.scrollLeft = 0;
    exportContainer.scrollTop = 0;

    try {
        const canvas = await html2canvas(exportContainer, {
            scale: 2,
            useCORS: true,
            width: exportContainer.scrollWidth,
            height: exportContainer.scrollHeight,
            windowWidth: exportContainer.scrollWidth,
            windowHeight: exportContainer.scrollHeight,
        });

        const imgData = canvas.toDataURL('image/png');
        const { jsPDF } = jspdf;

        // --- PDF Generation Constants ---
        const DPI = 96;
        const MM_PER_INCH = 25.4;
        const TITLE_MARGIN_MM = 20;

        // --- Calculate Dimensions ---
        const imgWidthMM = (canvas.width / DPI) * MM_PER_INCH;
        const imgHeightMM = (canvas.height / DPI) * MM_PER_INCH;
        const totalPdfWidth = imgWidthMM;
        const totalPdfHeight = imgHeightMM + TITLE_MARGIN_MM;
        const orientation = totalPdfWidth > totalPdfHeight ? 'l' : 'p';

        // --- Create PDF ---
        const pdf = new jsPDF({
            orientation: orientation,
            unit: 'mm',
            format: [totalPdfWidth, totalPdfHeight]
        });

        // --- Add Title ---
        const projectName = projectInfo.name || 'Gantt Chart';
        pdf.setFontSize(16);
        pdf.setFont(undefined, 'bold');
        pdf.text(projectName, totalPdfWidth / 2, TITLE_MARGIN_MM / 2, {
            align: 'center',
            baseline: 'middle'
        });
        
        // --- Add Gantt Chart Image ---
        pdf.addImage(imgData, 'PNG', 0, TITLE_MARGIN_MM, imgWidthMM, imgHeightMM);
        
        // --- Save PDF ---
        const safeProjectName = (projectInfo.name || 'gaseng-pmt-chart')
            .trim()
            .replace(/[^a-z0-9]/gi, '_')
            .toLowerCase();
        pdf.save(`${safeProjectName}.pdf`);

    } catch (error) {
        console.error("Failed to export Gantt to PDF:", error);
        alert("Could not export Gantt chart to PDF. See console for details.");
    } finally {
        generateGanttPdfBtn.textContent = 'Generate PDF';
        generateGanttPdfBtn.disabled = false;
        closePdfExportModal();
        exportContainer.scrollLeft = originalScrollLeft;
        exportContainer.scrollTop = originalScrollTop;
    }
}


// --- CONFIRMATION MODAL ---

/**
 * Shows a confirmation modal.
 * @param message The message to display.
 * @param onConfirm The function to execute if the user confirms.
 * @param onCancel The function to execute if the user cancels.
 */
function showConfirmation(message: string, onConfirm: () => void, onCancel?: () => void): void {
    confirmationMessage.textContent = message;
    onConfirmCallback = onConfirm;
    onCancelCallback = onCancel || null;
    confirmationModal.style.display = 'flex';
}

/**
 * Hides the confirmation modal and clears callbacks.
 */
function hideConfirmation(): void {
    confirmationModal.style.display = 'none';
    onConfirmCallback = null;
    onCancelCallback = null;
}

/**
 * Toggles the visibility of form fields for sub-tasks based on parent selection.
 */
function toggleSubtaskFields(): void {
    const isSubtask = parentTaskSelect.value !== '';
    if (isSubtask) {
        subtaskFieldsContainer.style.display = 'contents';
        startDateInput.required = true;
        endDateInput.required = true;
        percentCompleteInput.required = true;
    } else {
        subtaskFieldsContainer.style.display = 'none';
        startDateInput.required = false;
        endDateInput.required = false;
        percentCompleteInput.required = false;
    }
}

// Event Listeners
parentTaskSelect.addEventListener('change', toggleSubtaskFields);

addTaskForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const parentId = parentTaskSelect.value ? parseInt(parentTaskSelect.value, 10) : null;
    
    if (parentId !== null) { // Logic for adding a sub-task
        if (taskNameInput.value && startDateInput.value && endDateInput.value) {
            if (parseDateUTC(startDateInput.value) > parseDateUTC(endDateInput.value)) {
                alert("Error: The start date cannot be later than the end date.");
                return;
            }

            const resources = resourcesInput.value.split(',').map(r => r.trim()).filter(Boolean);
            const customFields: { [key: string]: string } = {};
            customFieldsSchema.forEach(field => {
                const input = addTaskForm.querySelector(`[name="${field.name}"]`) as HTMLInputElement;
                if (input) customFields[field.name] = input.value.trim();
            });
            
            addTask(
                taskNameInput.value,
                startDateInput.value,
                endDateInput.value,
                parentId,
                parseInt(percentCompleteInput.value, 10) || 0,
                resources,
                customFields
            );
        }
    } else { // Logic for adding a new main task
        if (taskNameInput.value) {
            const today = formatDateUTC(new Date());
            addTask(
                taskNameInput.value,
                today,
                today,
                null,
                0,
                [],
                {}
            );
        }
    }

    addTaskForm.reset();
    parentTaskSelect.value = '';
    flatpickrInstances['start-date-input']?.clear();
    flatpickrInstances['end-date-input']?.clear();
    percentCompleteInput.value = '0';
    taskNameInput.focus();
    toggleSubtaskFields(); // Reset field visibility
});

editTaskForm.addEventListener('submit', (e) => {
    e.preventDefault();
    if (editingTaskId !== null) {
        const task = tasks.find(t => t.id === editingTaskId);
        if (!task) return;
        
        // Add validation to prevent saving if start date is after end date
        if (parseDateUTC(editStartDateInput.value) > parseDateUTC(endDateInput.value)) {
            alert("Error: The start date cannot be later than the end date.");
            return; // Stop the function, keeping the modal open for correction
        }

        const resources = editResourcesInput.value.split(',').map(r => r.trim()).filter(Boolean);
        const customFields: { [key: string]: string } = {};
        customFieldsSchema.forEach(field => {
            const input = editTaskForm.querySelector(`[name="${field.name}"]`) as HTMLInputElement;
            if (input) customFields[field.name] = input.value.trim();
        });

        const isSubtask = task.parentId !== null;
        const selectedDeps = isSubtask ? [] : Array.from(editDependenciesSelect.selectedOptions).map(opt => parseInt(opt.value, 10));

        // Check for circular dependencies, but only for main tasks
        if (!isSubtask) {
            for (const depId of selectedDeps) {
                if (isCircularDependency(editingTaskId, depId)) {
                    const depTask = tasks.find(t => t.id === depId);
                    alert(`Error: Creating a circular dependency. Task "${depTask?.name}" already depends on "${task.name}".`);
                    return;
                }
            }
        }

        updateTask(editingTaskId, {
            name: editTaskNameInput.value,
            description: editTaskDescriptionTextarea.value.trim(),
            startDate: editStartDateInput.value,
            endDate: editEndDateInput.value,
            percentComplete: parseInt(editPercentCompleteInput.value, 10) || 0,
            resources: resources,
            notes: editNotesTextarea.value.trim(),
            status: editStatusSelect.value as Task['status'],
            customFields: customFields,
            dependencies: selectedDeps
        });
        closeModal();
    }
});

viewSwitcher.addEventListener('click', (e) => {
    const target = e.target as HTMLButtonElement;
    if (target.classList.contains('view-btn')) {
        const view = target.dataset.view as ViewMode;
        if (view) {
            currentView = view;
            document.querySelectorAll('.view-btn').forEach(btn => btn.classList.remove('active'));
            target.classList.add('active');
            render();
        }
    }
});

closeModalButton.addEventListener('click', closeModal);
deleteTaskButton.addEventListener('click', () => {
    if (editingTaskId !== null) {
        showConfirmation('Are you sure you want to delete this task and all its sub-tasks?', () => {
            deleteTask(editingTaskId!);
            closeModal();
        });
    }
});

// Data Controls
newProjectBtn?.addEventListener('click', () => {
    showConfirmation(
        'Are you sure you want to create a new project? All unsaved changes will be lost.',
        handleNewProject
    );
});
saveProjectBtn?.addEventListener('click', saveProjectToFile);
loadProjectInput?.addEventListener('change', handleProjectFileLoad);

// Export Buttons
exportCsvBtn?.addEventListener('click', exportCSV);
exportPdfChartBtn?.addEventListener('click', openPdfExportModal);

setBaselineBtn?.addEventListener('click', setBaseline);
showBaselineToggle?.addEventListener('change', () => {
    showBaseline = showBaselineToggle.checked;
    render();
});

filterNameInput.addEventListener('input', () => {
    filters.name = filterNameInput.value;
    render();
});
filterResourceInput.addEventListener('input', () => {
    filters.resource = filterResourceInput.value;
    render();
});
filterStatusSelect.addEventListener('change', () => {
    filters.status = filterStatusSelect.value;
    render();
});
sortTasksSelect.addEventListener('change', () => {
    const [key, order] = sortTasksSelect.value.split('-');
    sortCriteria = { key, order };
    render();
});

colorCodeSelect.addEventListener('change', () => {
    colorCodeCriteria = colorCodeSelect.value;
    render();
});

zoomSlider.addEventListener('input', () => {
    dayWidth = parseInt(zoomSlider.value, 10);
    zoomValue.textContent = `${dayWidth}px`;
    renderGanttChart();
});


// Custom Fields Modal Listeners
manageCustomFieldsBtn.addEventListener('click', () => {
    renderCustomFieldsList();
    customFieldsModal.style.display = 'flex';
});
closeCustomFieldsModalBtn.addEventListener('click', () => {
    customFieldsModal.style.display = 'none';
});
customFieldsModal.addEventListener('click', (e) => {
    if (e.target === customFieldsModal) {
        customFieldsModal.style.display = 'none';
    }
});
addCustomFieldForm.addEventListener('submit', (e) => {
    e.preventDefault();
    if(customFieldNameInput.value) {
        addCustomField(customFieldNameInput.value.trim());
        addCustomFieldForm.reset();
    }
});

// Multiple Sub-tasks Modal Listeners
closeMultipleSubtasksModalBtn.addEventListener('click', () => {
    multipleSubtasksModal.style.display = 'none';
});
cancelMultipleSubtasksBtn.addEventListener('click', () => {
    multipleSubtasksModal.style.display = 'none';
});
submitMultipleSubtasksBtn.addEventListener('click', handleMultipleSubtasksSubmit);


// Notification Listeners
notificationBell.addEventListener('click', (e) => {
    e.stopPropagation();
    const isVisible = notificationPanel.style.display === 'block';
    notificationPanel.style.display = isVisible ? 'none' : 'block';
});

// Report Page & Tab Listeners
tabNavigation.addEventListener('click', (e) => {
    const target = e.target as HTMLButtonElement;
    if (target.classList.contains('tab-btn')) {
        const tabId = target.dataset.tab;
        if (!tabId) return;

        // Update active tab button
        document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
        target.classList.add('active');

        // Show/hide content
        tabContents.forEach(content => {
            let contentId = '';
            if (content.id === 'gantt-view-container') contentId = 'plan';
            if (content.id === 'report-page') contentId = 'report';
            if (content.id === 'instructions-page') contentId = 'instructions';
            if (content.id === 'meeting-minutes-page') contentId = 'minutes';
            if (content.id === 'action-items-page') contentId = 'action-items';
            if (content.id === 'previous-reports-page') contentId = 'previous-reports';
            
            content.style.display = contentId === tabId ? 'block' : 'none';
        });
        
        if (tabId === 'report') {
            generateReport();
        }
        if (tabId === 'minutes') {
            renderMeetingMinutesPage();
        }
        if (tabId === 'action-items') {
            renderActionItemsPage();
        }
        if (tabId === 'previous-reports') {
            renderPreviousReportsPage();
        }
    }
});

exportPdfBtn.addEventListener('click', exportReportToPDF);
saveToArchiveBtn.addEventListener('click', () => archiveCurrentReport(true));
upcomingDaysInput.addEventListener('input', generateReport);

// PDF Export Modal Listeners
closePdfModalBtn.addEventListener('click', closePdfExportModal);
cancelPdfExportBtn.addEventListener('click', closePdfExportModal);
generateGanttPdfBtn.addEventListener('click', exportGanttToPdf);

// Confirmation Modal Listeners
confirmationCancelBtn.addEventListener('click', () => {
    if (onCancelCallback) {
        onCancelCallback();
    }
    hideConfirmation();
});
confirmationConfirmBtn.addEventListener('click', () => {
    if (onConfirmCallback) {
        onConfirmCallback();
    }
    hideConfirmation();
});

// Project Info Form Listener
projectInfoForm?.addEventListener('input', (e) => {
    const target = e.target as HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement;
    
    switch (target.id) {
        case 'project-name':
            projectInfo.name = target.value;
            break;
        case 'project-description':
            projectInfo.description = target.value;
            break;
        case 'project-deliverables':
            projectInfo.deliverables = target.value;
            break;
        case 'project-location':
            projectInfo.location = target.value;
            break;
        case 'project-doc':
            projectInfo.doc = target.value;
            break;
        case 'project-owner':
            projectInfo.owner = target.value;
            break;
        case 'project-engineering-manager':
            projectInfo.engineeringManager = target.value;
            break;
        case 'project-manager':
            projectInfo.manager = target.value;
            break;
        case 'project-engineers':
            projectInfo.engineers = target.value;
            break;
        case 'project-start-date':
            projectInfo.startDate = target.value;
            break;
        case 'project-end-date':
            projectInfo.endDate = target.value;
            break;
        case 'project-budget':
            projectInfo.budget = (target as HTMLInputElement).valueAsNumber || '';
            break;
        case 'project-currency':
            projectInfo.currency = target.value;
            break;
    }
});

// Meeting Minutes Listeners
addMeetingMinuteForm.addEventListener('submit', (e) => {
    e.preventDefault();
    addMeetingMinute();
});
meetingMinutesList.addEventListener('click', (e) => {
    const target = e.target as HTMLElement;
    if (target.matches('.delete-btn-list')) {
        const id = parseInt(target.dataset.id!, 10);
        if (id) {
            deleteMeetingMinute(id);
        }
    }
});

// Action Items Listeners
addActionItemForm.addEventListener('submit', (e) => {
    e.preventDefault();
    addActionItem();
});
actionItemsTableBody.addEventListener('click', (e) => {
    const target = e.target as HTMLElement;
    if (target.matches('.delete-btn-list')) {
        const id = parseInt(target.dataset.id!, 10);
        if (id) {
            deleteActionItem(id);
        }
    }
});

// Previous Reports Listeners
previousReportsList.addEventListener('click', (e) => {
    const target = e.target as HTMLElement;

    // Handle delete button click
    const deleteBtn = target.closest('.delete-btn-list');
    if (deleteBtn) {
        const id = parseInt(deleteBtn.getAttribute('data-id')!, 10);
        if (id) {
            deletePreviousReport(id);
        }
        return; // Stop propagation to prevent accordion from toggling
    }

    // Handle accordion toggle
    const header = target.closest('.previous-report-header');
    if (header) {
        const card = header.closest('.previous-report-card');
        card?.classList.toggle('expanded');
    }
});

// Close modals/panels on outside click
document.addEventListener('click', (e) => {
    if (e.target !== notificationBell && !notificationBell.contains(e.target as Node)) {
        notificationPanel.style.display = 'none';
    }
    if (e.target === editTaskModal) {
        closeModal();
    }
    if (e.target === pdfExportModal) {
        closePdfExportModal();
    }
    if (e.target === confirmationModal) {
        hideConfirmation();
    }
    if (e.target === multipleSubtasksModal) {
        multipleSubtasksModal.style.display = 'none';
    }
});


// Initial Load
toggleSubtaskFields(); // Set initial form state
loadState();
document.querySelector(`.view-btn[data-view="${currentView}"]`)?.classList.add('active');
initializeDatePickers();