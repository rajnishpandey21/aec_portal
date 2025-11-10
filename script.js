// This function MUST be global, so we define it outside the DOMContentLoaded
// This is the function Google will call after a successful sign-in
function handleGoogleLogin(response) {
    // 'response.credential' is an encoded JWT (JSON Web Token)
    // We need to decode it to get the user's email
    const idToken = response.credential;
    const decodedToken = JSON.parse(atob(idToken.split('.')[1]));
    const userEmail = decodedToken.email;
    console.log("Google Sign-In successful. Email:", userEmail);

    // Manually trigger our app's login logic
    window.app.login(userEmail);
}


document.addEventListener("DOMContentLoaded", () => {

    // --- 1. Get All Our HTML Elements ---
    
    // --- NEW: Select login and app containers ---
    const loginContainer = document.getElementById("login-container");
    const appContainer = document.getElementById("app-container");
    
    const logoutBtn = document.getElementById("logout-btn");
    const userInfo = document.getElementById("user-info");
    const loginStatus = document.getElementById("login-status"); // For error messages
    
    const navLinks = document.querySelectorAll(".nav-link");
    const pages = document.querySelectorAll(".page");

    // (All your form and table elements remain the same...)
    const scheduleForm = document.getElementById("schedule-form");
    const formDate = document.getElementById("form-date");
    const formBatch = document.getElementById("form-batch");
    const formShift = document.getElementById("form-shift");
    const formFaculty = document.getElementById("form-faculty");
    const formSubmitBtn = document.getElementById("form-submit-btn");
    const formStatus = document.getElementById("form-status");
    const formIsDemo = document.getElementById("form-is-demo");
    const addRowBtn = document.getElementById("form-add-row-btn");
    const bulkStatus = document.getElementById("bulk-status");
    const pendingTable = document.getElementById("pending-entries-table");
    const bulkSection = document.getElementById("bulk-section");
    const pendingTbody = pendingTable ? pendingTable.querySelector("tbody") : null;
    const baStatus = document.getElementById("ba-status");
    const baTbody = document.getElementById("ba-tbody");
    const viewScheduleStatus = document.getElementById("schedule-view-status");
    const viewScheduleTable = document.getElementById("schedule-view-table");
    const viewScheduleTbody = document.getElementById("schedule-view-tbody");
    const queueJobModal = document.getElementById("queue-job-modal");
    const queueJobStatus = document.getElementById("queue-job-status");
    const queueJobRunBtn = document.getElementById("queue-job-run-btn");
    const queueJobCloseBtn = document.getElementById("queue-job-close-btn");
    // View Schedule pagination/status controls and modal elements
    const vsPrev = document.getElementById('vs-prev');
    const vsNext = document.getElementById('vs-next');
    const vsToday = document.getElementById('vs-today');
    const vsMarkAll = document.getElementById('vs-mark-all');
    const vsWindowLabel = document.getElementById('vs-window-label');
    const vsModal = document.getElementById('vs-modal');
    const vsModalSave = document.getElementById('vs-modal-save');
    const vsModalCancel = document.getElementById('vs-modal-cancel');
    const vsCancelFields = document.getElementById('vs-cancel-fields');
    const vsSubFields = document.getElementById('vs-sub-fields');
    const vsReason = document.getElementById('vs-reason');
    const vsSubFaculty = document.getElementById('vs-sub-faculty');
    const vsSubTopic = document.getElementById('vs-sub-topic');
    const vsNewHall = document.getElementById('vs-new-hall');
    const vsSubSubject = document.getElementById('vs-sub-subject');
    const vsSubReason = document.getElementById('vs-sub-reason');
    // Context map for view-schedule rows (used by actions/modals)
    let vsRowContext = {};
    const filterDateStart = document.getElementById("filter-date-start");
    const filterDateEnd = document.getElementById("filter-date-end");
    const filterBatch = document.getElementById("filter-batch");
    const filterFaculty = document.getElementById("filter-faculty");
    const filterHall = document.getElementById("filter-hall");
    const filterResetBtn = document.getElementById("filter-reset-btn");
    // Special lecture form elements
    const slForm = document.getElementById('sl-form');
    const slDate = document.getElementById('sl-date');
    const slBatches = document.getElementById('sl-batches');
    const slCourses = document.getElementById('sl-courses');
    const slFaculty = document.getElementById('sl-faculty');
    const slExpertise = document.getElementById('sl-expertise');
    const slHall = document.getElementById('sl-hall');
    const slTime = document.getElementById('sl-time');
    const slDuration = document.getElementById('sl-duration');
    const slTopic = document.getElementById('sl-topic');
    const slPayment = document.getElementById('sl-payment');
    const slReason = document.getElementById('sl-reason'); // <-- ADD THIS
    const slStatus = document.getElementById('sl-status');
    const slBatchTrigger = document.getElementById('sl-batch-trigger');
    const slBatchMenu = document.getElementById('sl-batch-menu');
    const slBatchSearch = document.getElementById('sl-batch-search');
    const slBatchOptions = document.getElementById('sl-batch-options');
    
    // --- NEW: Analysis Dashboard Elements ---
    const pageAnalysis = document.getElementById("page-analysis");
    const analysisStatus = document.getElementById("analysis-status");
    const analysisDateStart = document.getElementById("analysis-date-start");
    const analysisDateEnd = document.getElementById("analysis-date-end");
    const analysisBatch = document.getElementById("analysis-batch");
    const analysisApplyBtn = document.getElementById("analysis-apply-btn");
    
    // --- NEW: Hall Timeline Elements ---
    const pageHallTimeline = document.getElementById("page-hall-timeline");
    const timelineDate = document.getElementById("timeline-date");
    const timelineRefreshBtn = document.getElementById("timeline-refresh-btn");
    const timelineStatus = document.getElementById("timeline-status");
    const timelineContainer = document.getElementById("timeline-container");
    const timelineBatchFilter = document.getElementById("timeline-batch-filter");
    const timelineFacultyFilter = document.getElementById("timeline-faculty-filter");
    // --- END NEW ---

    // --- NEW: Admin Elements ---
    const adminControls = document.getElementById("admin-controls");
    const adminBranchSwitcher = document.getElementById("admin-branch-switcher");
    const arStatus = document.getElementById("ar-status");
    const arTbody = document.getElementById("ar-tbody");
    // --- END NEW ---

    const kpiPlannedClasses = document.getElementById("kpi-planned-classes");
    const kpiDoneClasses = document.getElementById("kpi-done-classes");
    const kpiFillRate = document.getElementById("kpi-fill-rate");
    const kpiCancelledClasses = document.getElementById("kpi-cancelled-classes");
    const kpiCancellationRate = document.getElementById("kpi-cancellation-rate");
    const kpiSubstitutedClasses = document.getElementById("kpi-substituted-classes");
    const kpiRunningBatches = document.getElementById("kpi-running-batches");
    const kpiTotalBatches = document.getElementById("kpi-total-batches");
    const kpiTotalChanges = document.getElementById("kpi-total-changes");
    const kpiTotalSpecial = document.getElementById("kpi-total-special");
    
    const analysisPivotTable = document.getElementById("analysis-pivot-table");
    const analysisTopFaculty = document.getElementById("analysis-top-faculty");
    const analysisCancelFaculty = document.getElementById("analysis-cancel-faculty");
    // --- END NEW ---

    // --- This object will hold our app's "state" ---
    const appState = {
        isLoggedIn: false,
        isAdmin: false, // <-- ADD THIS
        allBranches: [], // <-- ADD THIS
        branchName: null,
        branchId: null,
        batchLectures: [], 
        faculty: [],
        pendingEntries: [],
        existingSchedule: [],
        existingKeys: { duplicate: new Set(), facultyAtTime: new Set() },
        finalScheduleData: [],
        isScheduleLoaded: false,
        isDashboardLoaded: false, // --- NEW ---
        batchShiftToTime: new Map(),
        slStep: 1, // --- NEW: For Special Lecture Card Stack ---
        crStep: 1 // --- NEW: For Change Request Card Stack ---
    };

    // --- 2. Page Navigation Logic ---
    navLinks.forEach(link => {
        link.addEventListener("click", (e) => {
            
            // --- MODIFIED ---
            // Get the pageId first
            const pageId = link.getAttribute("data-page");

            // If it's the batch request link, do NOT prevent default.
            // Just return and let the browser handle the href link.
            if (pageId === "page-batch-request") {
                return;
            }
            // --- END MODIFICATION ---

            e.preventDefault(); // Stop the default link behavior for all *other* links
            if (!appState.isLoggedIn) {
                // User should not be able to click this if not logged in, but as a fallback.
                return;
            }
            // const pageId = link.getAttribute("data-page"); // Already defined above
            
            // --- UPDATED: Handle new page loads ---
            if (pageId === "page-analysis" && !appState.isDashboardLoaded) {
                loadDashboardAnalytics(); 
            }
            // --- NEW: Handle Hall Timeline ---
            if (pageId === "page-hall-timeline") {
                if (timelineDate && !timelineDate.value) {
                    timelineDate.value = formatYmdLocal(new Date());
                }
                loadHallTimeline();
            }
            // --- END NEW ---
            
            if (pageId === "page-view-schedule" && !appState.isScheduleLoaded) {
                loadViewScheduleData(); 
            }
            if (pageId === "page-change-request") {
                initChangeRequestForm();
            }
            if (pageId === "page-batches-analysis") {
                loadBatchAnalysis();
            }
            // --- NEW: Handle Approve Requests Page ---
            if (pageId === "page-approve-requests") {
                loadApproveRequestsPage();
            }
            // --- END NEW ---
            if (pageId === 'page-special-lecture') {
                initSpecialLectureForm();
            }
            pages.forEach(page => page.classList.remove("active"));
            const targetPage = document.getElementById(pageId);
            if (targetPage) {
                targetPage.classList.add("active");
            }
            // Update active nav link (skip for batch-request which opens in new tab)
            if (pageId !== "page-batch-request") {
                navLinks.forEach(nav => nav.classList.remove("active"));
                link.classList.add("active");
            }
            if (pageId === "page-batches-analysis") {
                // ensure status area visible
                if (baStatus) baStatus.style.display = 'block';
            }
        });
    });

    // --- 4. Login / Logout Logic ---

    // This is our new app-wide login function
    const appLogin = async (email) => {
        try {
            showLoginStatus("Verifying user: " + email, "loading");

            // Call our backend to get role and branch details
            const result = await jsonpRequest(API_URL, {
                action: "getUserDetails",
                email: email
            });

            if (result.status === "success") {
                appState.isLoggedIn = true;
                appState.userEmail = email;

                if (result.role === "admin") {
                    // --- ADMIN LOGIN ---
                    console.log("Logged in as ADMIN");
                    appState.isAdmin = true;
                    document.body.classList.add("admin-mode");
                    userInfo.textContent = `Admin: ${email.split('@')[0]}`;

                    // Persist admin session
                    try {
                        localStorage.setItem('aec_session', JSON.stringify({
                            email: email,
                            role: "admin"
                        }));
                    } catch (e) { /* ignore */ }

                    // Init the admin UI (fetch branches, show dropdown)
                    await initAdminMode();

                    // Don't load any branch data until admin selects one
                    showLoginStatus("", "idle");
                    if (loginContainer) loginContainer.style.display = "none";
                    if (appContainer) appContainer.style.display = "flex";

                } else if (result.role === "branch") {
                    // --- BRANCH LOGIN ---
                    console.log("Logged in as BRANCH user");
                    appState.isAdmin = false;
                    appState.branchName = result.branchName;
                    appState.branchId = result.branchId;

                    userInfo.textContent = `Branch: ${appState.branchName}`;

                    // Persist branch session
                    try {
                        localStorage.setItem('aec_session', JSON.stringify({
                            email: email,
                            role: "branch",
                            branchName: appState.branchName,
                            branchId: appState.branchId
                        }));
                    } catch (e) { /* ignore */ }

                    // Load data for this user's assigned branch
                    await loadAllBranchData();

                    showLoginStatus("", "idle");
                    if (loginContainer) loginContainer.style.display = "none";
                    if (appContainer) appContainer.style.display = "flex";

                } else {
                    throw new Error("Invalid user role received.");
                }

            } else {
                throw new Error(result.message);
            }

        } catch (err) {
            console.error("Login failed:", err);
            showLoginStatus("Login Failed: " + err.message, "error");
        }
    };
    
    // Attach the login function to the window so Google can call it
    window.app = {
        login: appLogin
    };

    // --- Try to restore session on refresh ---
    (async function tryRestoreSession() {
        try {
            const raw = localStorage.getItem('aec_session');
            if (!raw) return;
            const saved = JSON.parse(raw);
            if (!saved || !saved.email || !saved.role) return;

            appState.isLoggedIn = true;
            appState.userEmail = saved.email;

            if (saved.role === "admin") {
                // --- RESTORE ADMIN ---
                console.log("Restoring ADMIN session");
                appState.isAdmin = true;
                document.body.classList.add("admin-mode");
                userInfo.textContent = `Admin: ${saved.email.split('@')[0]}`;

                await initAdminMode(); // Fetch branches and set up dropdown

                if (loginContainer) loginContainer.style.display = "none";
                if (appContainer) appContainer.style.display = "flex";

            } else if (saved.role === "branch") {
                // --- RESTORE BRANCH ---
                console.log("Restoring BRANCH session");
                appState.isAdmin = false;
                appState.branchName = saved.branchName;
                appState.branchId = saved.branchId;

                if (!appState.branchId || !appState.branchName) return; // Invalid session

                userInfo.textContent = `Branch: ${appState.branchName}`;

                await loadAllBranchData(); // Load data for this branch

                if (loginContainer) loginContainer.style.display = "none";
                if (appContainer) appContainer.style.display = "flex";
            }

        } catch (e) {
            console.warn('Session restore failed:', e);
            localStorage.removeItem('aec_session'); // Clear bad session
        }
    })();

    // Logout Button
    logoutBtn.addEventListener("click", (e) => {
        e.preventDefault();
        
        // Reset all app state
        appState.isLoggedIn = false;
        appState.isAdmin = false; // <-- ADD THIS
        appState.allBranches = []; // <-- ADD THIS
        appState.branchName = null;
        appState.branchId = null;
        appState.batchLectures = [];
        appState.faculty = [];       
        appState.pendingEntries = [];
        appState.existingSchedule = [];
        appState.finalScheduleData = [];
        appState.isScheduleLoaded = false;
        appState.isDashboardLoaded = false; // --- NEW ---
        
        // This will disable the "auto sign-in"
        if (window.google && google.accounts && google.accounts.id) {
            google.accounts.id.disableAutoSelect();
        }

        userInfo.textContent = "Not Logged In";
        document.body.classList.remove("admin-mode"); // <-- ADD THIS
        
        // --- NEW: Hide app, show login screen ---
        if (appContainer) appContainer.style.display = "none";
        if (loginContainer) loginContainer.style.display = "flex";
        
        showLoginStatus("", "idle");
        try { localStorage.removeItem('aec_session'); } catch (e) { /* ignore */ }
        
        console.log("Logged out.");
    });

    function showLoginStatus(message, type) {
        if (!loginStatus) return;
        loginStatus.textContent = message;
        loginStatus.className = `form-status-message ${type}`;
        loginStatus.style.display = (type === "idle") ? "none" : "block";
    }

    /**
     * NEW: Fetches all branches and populates the admin dropdown.
     */
    async function initAdminMode() {
        if (!adminBranchSwitcher) return;
        try {
            // Call new backend action
            const result = await jsonpRequest(API_URL, {
                action: "getAdminsAndBranches"
            });

            if (result.status === "success") {
                appState.allBranches = result.branches || [];
                adminBranchSwitcher.innerHTML = '<option value="">-- Select a Branch --</option>';

                appState.allBranches.sort((a, b) => a.BranchName.localeCompare(b.BranchName));

                appState.allBranches.forEach(branch => {
                    const option = document.createElement("option");
                    option.value = branch.Sheet_ID;
                    option.textContent = branch.BranchName;
                    adminBranchSwitcher.appendChild(option);
                });

                // Add event listener
                adminBranchSwitcher.addEventListener("change", handleBranchSwitch);

            } else {
                throw new Error(result.message);
            }
        } catch (err) {
            console.error("Failed to init admin mode:", err);
            alert("Error: Could not load branch list for admin. " + err.message);
        }
    }

    /**
     * NEW: Called when an admin selects a branch from the dropdown.
     */
    async function handleBranchSwitch() {
        if (!adminBranchSwitcher) return;
        const selectedBranchId = adminBranchSwitcher.value;
        const selectedBranch = appState.allBranches.find(b => b.Sheet_ID === selectedBranchId);

        if (selectedBranch) {
            console.log(`Admin switching to branch: ${selectedBranch.BranchName}`);
            appState.branchId = selectedBranch.Sheet_ID;
            appState.branchName = selectedBranch.BranchName;

            // Reset flags
            appState.isScheduleLoaded = false;
            appState.isDashboardLoaded = false;

            // Activate the main dashboard page
            document.querySelectorAll(".page").forEach(p => p.classList.remove("active"));
            document.getElementById("page-analysis").classList.add("active");
            document.querySelectorAll(".nav-link").forEach(n => n.classList.remove("active"));
            document.querySelector('.nav-link[data-page="page-analysis"]').classList.add("active");

            // Load all data for the selected branch
            await loadAllBranchData();

        } else {
            // Reset if they selected "-- Select a Branch --"
            appState.branchId = null;
            appState.branchName = null;
        }
    }

    /**
     * NEW: A wrapper function to load all data for the active branch.
     * This is called on branch-user login, branch-user session restore,
     * and admin branch switch.
     */
    async function loadAllBranchData() {
        if (!appState.branchId || !appState.branchName) {
            console.warn("loadAllBranchData called without a branch selected.");
            return;
        }

        try {
            // Show loading indicators
            if (analysisStatus) {
                analysisStatus.innerHTML = '<span class="spinner"></span> Loading branch data...';
                analysisStatus.className = "form-status-message loading";
                analysisStatus.style.display = "block";
            }

            // Load essential form data and schedule keys first
            await loadFormData();
            await loadExistingSchedule();

            // Then load the dashboard for the active page
            // (page-analysis is the default active page)
            await loadDashboardAnalytics();

            if (analysisStatus) {
                analysisStatus.style.display = "none";
            }

        } catch (err) {
            console.error("Error loading all branch data:", err);
            if (analysisStatus) {
                analysisStatus.textContent = "Error loading branch data: " + err.message;
                analysisStatus.className = "form-status-message error";
            }
        }
    }

    // --- 5. Data Loading Logic (JSONP Helper) ---
    function jsonpRequest(url, params = {}) {
        return new Promise((resolve, reject) => {
            const callbackName = "jsonp_cb_" + Math.random().toString(36).slice(2);
            params.callback = callbackName;
            const qs = new URLSearchParams(params).toString();
            const script = document.createElement("script");
            script.src = `${url}?${qs}`;
            script.onerror = () => {
                delete window[callbackName];
                script.remove();
                reject(new Error("JSONP request failed"));
            };
            window[callbackName] = (data) => {
                resolve(data);
                delete window[callbackName];
                script.remove();
            };
            document.body.appendChild(script);
        });
    }

    // (This can go right after the jsonpRequest function, around line 340)

    /**
     * Converts a 12-hour AM/PM time string into a comparable number (minutes since midnight).
     * @param {string} timeStr - e.g., "11:15 AM" or "07:00 PM"
     * @returns {number}
     */
    function parseScheduleTime(timeStr) {
        try {
            const match = String(timeStr).match(/(\d+):(\d+)\s*(AM|PM)/i);
            if (!match) return 0; 

            let [_, hh, mm, ampm] = match;
            let hours = parseInt(hh, 10);
            
            if (ampm.toUpperCase() === 'PM' && hours !== 12) {
                hours += 12;
            }
            if (ampm.toUpperCase() === 'AM' && hours === 12) {
                hours = 0; // Midnight case
            }
            return (hours * 60) + parseInt(mm, 10);
        } catch (e) {
            return 0; // Fallback for any invalid time format
        }
    }

    function normalizeTimeToHHMM(timeStr) {
        if (!timeStr) return "";
        const m = String(timeStr).trim().toUpperCase().match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/);
        if (!m) return String(timeStr).trim();
        let hh = parseInt(m[1], 10);
        const mm = m[2];
        const ap = m[3];
        if (ap === 'AM') {
            if (hh === 12) hh = 0;
        } else {
            if (hh !== 12) hh += 12;
        }
        const hhStr = (hh < 10 ? '0' : '') + hh;
        return `${hhStr}:${mm}`;
    }
    
    function buildBatchShiftTimeMap() {
        appState.batchShiftToTime = new Map();
        (appState.batchLectures || []).forEach(row => {
            const batch = String(row.batchCode || row["Batch Code"] || "").trim();
            const shift = String(row.shiftId || row.shift_id || "").trim();
            const time = String(row.startTime || row.StartTime || "").trim();
            if (batch && shift && time) {
                const key = `${batch}|${shift}`;
                appState.batchShiftToTime.set(key, normalizeTimeToHHMM(time));
            }
        });
    }

    function getTimeCodeFor(batch, shift) {
        const key = `${String(batch || '').trim()}|${String(shift || '').trim()}`;
        return appState.batchShiftToTime.get(key) || "";
    }

    async function loadFormData() {
        try {
            console.log("Loading form data...");
            const [batchResult, facultyResult] = await Promise.all([
                jsonpRequest(API_URL, { action: "getBatchLectures", branchId: appState.branchId }),
                jsonpRequest(API_URL, { action: "getFaculty", branchName: appState.branchName })
            ]);

            if (batchResult.status === "success") {
                appState.batchLectures = batchResult.data;
                buildBatchShiftTimeMap();
                renderBatchDropdown();
                
                // --- NEW: Populate Analysis batch filter ---
                if (analysisBatch) {
                    analysisBatch.innerHTML = '<option value="">All Batches</option>';
                    const uniqueBatches = [...new Map(appState.batchLectures.map(item => [item.batchCode, item])).values()];
                    uniqueBatches.sort((a,b) => a.batchCode.localeCompare(b.batchCode)); // Sort them
                    uniqueBatches.forEach(batch => {
                        const option = document.createElement("option");
                        option.value = batch.batchCode;
                        option.textContent = `${batch.course} (${batch.batchCode})`;
                        analysisBatch.appendChild(option);
                    });
                }
                // --- END NEW ---
                
            } else {
                throw new Error("Batch Data Error: " + batchResult.message);
            }

            if (facultyResult.status === "success") {
                appState.faculty = facultyResult.data;
                renderFacultyDropdown();
            } else {
                throw new Error("Faculty Data Error: ".concat(facultyResult.message));
            }
            
            console.log("Form data loaded successfully.");
        } catch (err) {
            console.error("Error loading form data:", err);
            alert("Error loading form data. Please check the console. Message: " + err.message);
        }
    }

    // --- Change Request Page Logic ---
    const crForm = document.getElementById("cr-form");
    const crDate = document.getElementById("cr-date");
    const crBatch = document.getElementById("cr-batch");
    const crShift = document.getElementById("cr-shift");
    const crType = document.getElementById("cr-type");
    const crNewHallGroup = document.getElementById("cr-new-hall-group");
    const crNewHall = document.getElementById("cr-new-hall");
    const crNewDurationGroup = document.getElementById("cr-new-duration-group");
    const crNewDuration = document.getElementById("cr-new-duration");
    const crNewFacultyGroup = document.getElementById("cr-new-faculty-group");
    const crNewFaculty = document.getElementById("cr-new-faculty");
    const crRemarks = document.getElementById("cr-remarks");
    const crStatus = document.getElementById("cr-status");

    function initChangeRequestForm() {
        // --- This part is the same: Populate dropdowns ---
        if (crBatch) {
            crBatch.innerHTML = '<option value="">-- Select a Batch --</option>';
            const uniqueBatches = [...new Map(appState.batchLectures.map(item => [item.batchCode, item])).values()];
            uniqueBatches.sort((a,b) => a.batchCode.localeCompare(b.batchCode));
            uniqueBatches.forEach(batch => {
                const option = document.createElement("option");
                option.value = batch.batchCode;
                option.textContent = `${batch.course} (${batch.batchCode})`;
                crBatch.appendChild(option);
            });
        }
        if (crNewFaculty) {
            crNewFaculty.innerHTML = '<option value="">-- Select a Faculty --</option>';
            appState.faculty.forEach(faculty => {
                const option = document.createElement("option");
                option.value = faculty.code;
                option.textContent = `${faculty.name} (${faculty.expertise})`;
                crNewFaculty.appendChild(option);
            });
        }

        // --- NEW: Card Stack Logic ---
        appState.crStep = 1; // Reset to step 1
        updateCrCardStackView(); // Set initial card positions
        updateCrSummary(); // Clear summary

        // Add 'input' listener to the whole form for real-time updates
        if (crForm) {
            crForm.addEventListener('input', updateCrSummary);
        }

        // Add 'click' listeners to all "Next" buttons
        document.querySelectorAll('.cr-next-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const currentStep = appState.crStep;
                // --- Simple Validation ---
                let isValid = true;
                if (currentStep === 1 && !crDate.value) isValid = false;
                if (currentStep === 2 && !crBatch.value) isValid = false;
                if (currentStep === 3 && !crShift.value) isValid = false;
                if (currentStep === 4 && !crType.value) isValid = false;
                if (currentStep === 5) {
                    const type = crType.value;
                    if (type === 'HALL' && !crNewHall.value) isValid = false;
                    if (type === 'DURATION' && !crNewDuration.value) isValid = false;
                    if (type === 'TEACHER' && !crNewFaculty.value) isValid = false;
                }
                if (!isValid) {
                    showCrStatus("Please complete the current field.", "error");
                    return;
                }
                // --- End Validation ---
                showCrStatus("", "idle"); // Clear error
                const nextStep = parseInt(e.currentTarget.dataset.next, 10);
                appState.crStep = nextStep;
                updateCrCardStackView();
            });
        });

        // Handle batch/shift dependency
        if (crBatch) {
            crBatch.addEventListener('change', () => {
                const selectedBatchCode = crBatch.value;
                crShift.innerHTML = '<option value="">-- Select a Shift/Time --</option>';
                if (!selectedBatchCode) {
                    crShift.disabled = true;
                    return;
                }
                const shiftsForBatch = appState.batchLectures.filter(lecture => lecture.batchCode === selectedBatchCode);
                shiftsForBatch.sort((a, b) => a.lectureSerial - b.lectureSerial);
                shiftsForBatch.forEach(lecture => {
                    const option = document.createElement("option");
                    option.value = lecture.shiftId;
                    option.textContent = `Lec ${lecture.lectureSerial} (${lecture.startTime})`;
                    crShift.appendChild(option);
                });
                crShift.disabled = false;
            });
        }
    }

    function updateCRTypeUI() {
        if (!crType) return;
        const type = crType.value;
        crNewHallGroup.style.display = (type === 'HALL') ? 'block' : 'none';
        crNewDurationGroup.style.display = (type === 'DURATION') ? 'block' : 'none';
        crNewFacultyGroup.style.display = (type === 'TEACHER') ? 'block' : 'none';
    }

    if (crType) {
        crType.addEventListener('change', () => { updateCRTypeUI(); updateCrSummary(); });
    }

    if (crBatch) {
        crBatch.addEventListener('change', () => {
            const selectedBatchCode = crBatch.value;
            crShift.innerHTML = '<option value="">-- Select a Shift/Time --</option>';
            if (!selectedBatchCode) {
                crShift.disabled = true;
                return;
            }
            const shiftsForBatch = appState.batchLectures.filter(lecture => lecture.batchCode === selectedBatchCode);
            shiftsForBatch.sort((a, b) => a.lectureSerial - b.lectureSerial);
            shiftsForBatch.forEach(lecture => {
                const option = document.createElement("option");
                option.value = lecture.shiftId;
                option.textContent = `Lec ${lecture.lectureSerial} (${lecture.startTime})`;
                crShift.appendChild(option);
            });
            crShift.disabled = false;
            updateCrSummary();
        });
    }

    if (crForm) {
        crForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            showCrStatus("Submitting change request...", "loading");
            const [date, batch, shift] = [crDate.value, crBatch.value, crShift.value];
            if (!date || !batch || !shift || !crType.value) {
                showCrStatus("Please fill all required fields.", "error");
                appState.crStep = 1; // Send user back to first invalid step
                updateCrCardStackView();
                return;
            }
            // Check if slot exists
            const dupKey = `${date}|${batch}|${shift}`;
            if (!appState.existingKeys.duplicate.has(dupKey)) {
                await loadExistingSchedule(); // Re-fetch just in case
                if (!appState.existingKeys.duplicate.has(dupKey)) {
                    showCrStatus("No existing class found for this Date/Batch/Shift.", "error");
                    appState.crStep = 1; // Send user back
                    updateCrCardStackView();
                    return;
                }
            }
            let newValue = '';
            if (crType.value === 'HALL') newValue = crNewHall.value.trim();
            if (crType.value === 'DURATION') newValue = crNewDuration.value.trim();
            if (crType.value === 'TEACHER') newValue = crNewFaculty.value.trim();
            if (!newValue) {
                showCrStatus("Please provide the new value for the change.", "error");
                appState.crStep = 5; // Send user to "new value" step
                updateCrCardStackView();
                return;
            }
            try {
                const result = await jsonpRequest(API_URL, {
                    action: 'logChangeRequest',
                    branchId: appState.branchId,
                    lectDate: date,
                    batch: batch,
                    shift: shift,
                    changeType: crType.value,
                    newValue: newValue,
                    remarks: (crRemarks.value || ''),
                    requestedBy: (appState.userEmail || '')
                });
                
                if (result.status === 'success') {
                    showCrStatus("Change request logged successfully.", "success");
                    crForm.reset();
                    // Reset visibility in Step 5 options after reset
                    updateCRTypeUI();
                    appState.crStep = 1; // Reset stack to beginning
                    updateCrCardStackView();
                    updateCrSummary(); // Clear summary
                } else {
                    throw new Error(result.message);
                }
            } catch (err) {
                showCrStatus("Error: " + err.message, "error");
            }
        });
    }

    async function loadExistingSchedule() {
        try {
            const schedResult = await jsonpRequest(API_URL, { action: "getSchedule", branchId: appState.branchId });
            if (schedResult.status === "success") {
                appState.existingSchedule = schedResult.data || [];
                appState.existingKeys.duplicate.clear();
                appState.existingKeys.facultyAtTime.clear();
                
                appState.existingSchedule.forEach(row => {
                    // --- START FIX ---
                    // Normalize the date from the sheet
                    let dateStr = String(row.LectDate || row.lectDate || "").trim();
                    
                    // Check if it's a full ISO string (e.g., "2025-11-05T00:00:00.000Z")
                    if (dateStr.includes('T')) {
                        // Just take the date part
                        dateStr = dateStr.split('T')[0];
                    } 
                    // --- END FIX ---

                    const date = dateStr; // Use the normalized date
                    const batch = String(row.Batch || "").trim();
                    const shift = String(row.ShiftCode || row.shift || row.Shift || "").trim();
                    const faculty = String(row.FacultyCode || row.faculty || "").trim();
                    
                    if (date && batch && shift) appState.existingKeys.duplicate.add(`${date}|${batch}|${shift}`);
                    
                    const timeCode = getTimeCodeFor(batch, shift);
                    if (date && timeCode && faculty) appState.existingKeys.facultyAtTime.add(`${date}|${timeCode}|${faculty}`);
                });
            }
        } catch (e) {
            console.warn("Could not load existing schedule:", e);
        }
    }

    function renderBatchDropdown() {
        formBatch.innerHTML = '<option value="">-- Select a Batch --</option>'; 
        const uniqueBatches = [...new Map(appState.batchLectures.map(item =>
            [item.batchCode, item])).values()];
        uniqueBatches.sort((a,b) => a.batchCode.localeCompare(b.batchCode)); // Sort them
        uniqueBatches.forEach(batch => {
            const option = document.createElement("option");
            option.value = batch.batchCode;
            option.textContent = `${batch.course} (${batch.batchCode})`;
            formBatch.appendChild(option);
        });
    }

    function renderFacultyDropdown() {
        formFaculty.innerHTML = '<option value="">-- Select a Faculty --</option>'; 
        appState.faculty.forEach(faculty => {
            const option = document.createElement("option");
            option.value = faculty.code;
            option.textContent = `${faculty.name} (${faculty.expertise})`;
            formFaculty.appendChild(option);
        });
    }

    // --- 6. "Generate Schedule" Form Logic ---
    
    formBatch.addEventListener("change", () => {
        const selectedBatchCode = formBatch.value;
        formShift.innerHTML = '<option value="">-- Select a Shift/Time --</option>'; 
        if (!selectedBatchCode) {
            formShift.disabled = true;
            return;
        }
        const shiftsForBatch = appState.batchLectures.filter(
            lecture => lecture.batchCode === selectedBatchCode
        );
        shiftsForBatch.sort((a, b) => a.lectureSerial - b.lectureSerial);
        shiftsForBatch.forEach(lecture => {
            const option = document.createElement("option");
            option.value = lecture.shiftId;
            option.textContent = `Lec ${lecture.lectureSerial} (${lecture.startTime})`; 
            formShift.appendChild(option.cloneNode(true));
        });
        formShift.disabled = false;
    });

    if (addRowBtn) {
        addRowBtn.addEventListener("click", () => {
            const entry = {
                date: formDate.value,
                batch: formBatch.value,
                shift: formShift.value,
                faculty: formFaculty.value,
                type: formIsDemo && formIsDemo.checked ? "Demo" : "",
                status: "ok"
            };
            if (!entry.date || !entry.batch || !entry.shift || !entry.faculty) {
                showFormStatus("Please fill out all fields before adding.", "error");
                return;
            }
            const validation = validateEntry(entry, true);
            if (!validation.ok) {
                showFormStatus(validation.message, "error");
                return;
            }
            appState.pendingEntries.push(entry);
            renderPendingEntries();
            showFormStatus("Row added to pending list.", "success");
            if (formIsDemo) formIsDemo.checked = false;
        });
    }

    if (scheduleForm) {
        scheduleForm.addEventListener("submit", async (e) => {
            e.preventDefault(); 
            if (formSubmitBtn.dataset.mode === "all") {
                await submitAllPending();
            } else {
                await submitSingleEntry();
            }
        });
        // Handle bulk mode without triggering native validation
        formSubmitBtn.addEventListener("click", async (e) => {
            if (formSubmitBtn.dataset.mode === "all") {
                e.preventDefault();
                await submitAllPending();
            }
        });
    }

    // --- REPLACEMENT: loadBatchAnalysis ---
    async function loadBatchAnalysis() {
        if (!baStatus || !baTbody) return;
        baStatus.style.display = "block";
        baStatus.textContent = "Loading batches analysis...";
        baStatus.className = "form-status-message loading";
        baTbody.innerHTML = "";
        
        try {
            const result = await jsonpRequest(API_URL, {
                action: "getBatchAnalysis",
                branchId: appState.branchId
            });
            
            if (result.status !== "success") throw new Error(result.message || "Failed to load analysis");
            
            const rows = result.data || [];
            let running = 0, expired = 0;
            
            rows.forEach(r => {
                if (r.status === 'Running') running++; else expired++;
                
                const tr = document.createElement("tr");
                const statusClass = r.status === 'Running' ? 'success' : 'error';
                
                // Add highlight classes for low 'remaining' counts
                let remainingClass = '';
                if (r.status === 'Running') {
                    if (r.remaining <= 10) remainingClass = 'error';
                    else if (r.remaining <= 25) remainingClass = 'loading'; // 'loading' style is yellow
                }
                
                tr.innerHTML = `
                    <td>${r.batch}</td>
                    <td><span class="form-status-message ${statusClass}">${r.status}</span></td>
                    <td>${r.lastLectureDate || '—'}</td>
                    <td>${r.total}</td>
                    <td>${r.planned}</td>
                    <td class="ba-remaining ${remainingClass}">${r.remaining}</td>
                `;
                baTbody.appendChild(tr);
            });
            
            baStatus.textContent = `Batches: ${running} running, ${expired} expired.`;
            baStatus.className = "form-status-message success";
            
        } catch (err) {
            baStatus.textContent = "Error: " + err.message;
            baStatus.className = "form-status-message error";
        }
    }
    // --- END REPLACEMENT ---
    
    async function submitSingleEntry() {
        const type = formIsDemo && formIsDemo.checked ? "Demo" : "";
        const entry = [formDate.value, formBatch.value, formShift.value, formFaculty.value, type];
        if (!entry[0] || !entry[1] || !entry[2] || !entry[3]) {
            showFormStatus("Please fill out all fields.", "error");
            return;
        }
        const validation = validateEntry(entry, true);
        if (!validation.ok) {
            showFormStatus(validation.message, "error");
            return;
        }
        showFormStatus("Submitting class...", "loading");
        formSubmitBtn.disabled = true;

        try {
            const requestPayload = {
                action: "addNewSchedule",
                branchId: appState.branchId,
                lectDate: entry[0],
                batch: entry[1],
                shift: entry[2],
                faculty: entry[3],
                type: entry[4]
            };
            console.groupCollapsed('[Single] Submitting');
            console.log('Request', requestPayload);
            const result = await jsonpRequest(API_URL, requestPayload);
            console.log('Response', result);
            console.groupEnd();
            
            if (result.status === "success") {
                showQueueJobModal();
                // Add to local cache
                appState.existingKeys.duplicate.add(`${entry[0]}|${entry[1]}|${entry[2]}`);
                const t = getTimeCodeFor(entry[1], entry[2]);
                if (t) appState.existingKeys.facultyAtTime.add(`${entry[0]}|${t}|${entry[3]}`);
                // Clear form
                scheduleForm.reset();
                formShift.disabled = true;
                showFormStatus("Class added. Please queue the job.", "success");
            } else {
                throw new Error(result.message);
            }
        } catch (err) {
            console.error("Error submitting form:", err);
            showFormStatus("Error: " + err.message, "error");
        } finally {
            formSubmitBtn.disabled = false;
        }
    }

    async function submitAllPending() {
        if (appState.pendingEntries.length === 0) return;
        if (addRowBtn) addRowBtn.disabled = true;
        formSubmitBtn.disabled = true;
        if (bulkStatus) {
            bulkStatus.textContent = `Submitting ${appState.pendingEntries.length} entries...`;
            bulkStatus.className = "form-status-message loading";
        }

        const snapshot = appState.pendingEntries.map((entry, idx) => ({ entry, idx }));
        let successCount = 0;
        let errorCount = 0;

        for (let i = 0; i < snapshot.length; i++) {
            const { entry, idx } = snapshot[i];
            const [date, batch, shift, faculty, _] = getEntryTuple(entry);
            const rowEl = pendingTbody ? pendingTbody.querySelector(`tr[data-idx="${idx}"]`) : null;
            const statusEl = rowEl ? rowEl.querySelector(".status-pill") : null;
            if (statusEl) {
                statusEl.textContent = "Submitting...";
                statusEl.className = "status-pill form-status-message loading";
                statusEl.removeAttribute('title');
            }

            try {
                const [d2, b2, s2, f2, type] = getEntryTuple(entry);
                const requestPayload = {
                    action: "addNewSchedule",
                    branchId: appState.branchId,
                    lectDate: d2,
                    batch: b2,
                    shift: s2,
                    faculty: f2,
                    type: type
                };
                console.groupCollapsed(`[Bulk] Submitting row ${i + 1}/${snapshot.length}`);
                console.log('Request', requestPayload);
                const result = await jsonpRequest(API_URL, requestPayload);
                console.log('Response', result);
                console.groupEnd();
                if (result.status === "success") {
                    successCount++;
                    if (statusEl) {
                        statusEl.textContent = "Added";
                        statusEl.className = "status-pill form-status-message success";
                        statusEl.title = "Added to sheet";
                    }
                    appState.existingKeys.duplicate.add(`${date}|${batch}|${shift}`);
                    const t = getTimeCodeFor(batch, shift);
                    if (t) appState.existingKeys.facultyAtTime.add(`${date}|${t}|${faculty}`);
                } else {
                    errorCount++;
                    const msg = String(result.message || "Error");
                    const short = msg.includes("Duplicate") ? "Duplicate" : (msg.includes("Conflict") ? "Conflict" : "Error");
                    if (statusEl) {
                        statusEl.textContent = short;
                        statusEl.className = "status-pill form-status-message error";
                        statusEl.title = msg;
                    }
                    entry.status = short.toLowerCase();
                }
            } catch (err) {
                errorCount++;
                console.groupCollapsed(`[Bulk] Row ${i + 1} error`);
                console.error(err);
                console.groupEnd();
                if (statusEl) {
                    statusEl.textContent = "Error";
                    statusEl.className = "status-pill form-status-message error";
                    statusEl.title = String(err && err.message ? err.message : err);
                }
                entry.status = "error";
            }
            if (bulkStatus) bulkStatus.textContent = `Submitted ${i + 1}/${snapshot.length}...`;
        }

        const keep = [];
        for (let j = 0; j < snapshot.length; j++) {
            const { entry, idx } = snapshot[j];
            const rowEl = pendingTbody ? pendingTbody.querySelector(`tr[data-idx="${idx}"]`) : null;
            const statusEl = rowEl ? rowEl.querySelector(".status-pill") : null;
            const isSuccess = statusEl && statusEl.textContent === "Added";
            if (!isSuccess) keep.push(entry);
        }
        appState.pendingEntries = keep;
        renderPendingEntries();

        if (bulkStatus) {
            const summary = `Done. Added ${successCount}, Rejected ${errorCount}.`;
            bulkStatus.textContent = summary;
            bulkStatus.className = "form-status-message " + (successCount > 0 && errorCount === 0 ? "success" : (successCount === 0 ? "error" : "loading"));
        }
        // Show the modal if we added anything
        if (successCount > 0) {
            showQueueJobModal();
        }
        if (addRowBtn) addRowBtn.disabled = false;
        formSubmitBtn.disabled = false;
    }
    
    function renderPendingEntries() {
        if (!pendingTbody) return;
        pendingTbody.innerHTML = "";
        const localDup = new Set();
        const localFac = new Set();
        appState.pendingEntries.forEach((entry, idx) => {
            const [date, batch, shift, faculty, type] = getEntryTuple(entry);
            const dupKey = `${date}|${batch}|${shift}`;
            const timeCode = getTimeCodeFor(batch, shift);
            const facKey = timeCode ? `${date}|${timeCode}|${faculty}` : `${date}|${shift}|${faculty}`; // Use time-based key
            const isDup = appState.existingKeys.duplicate.has(dupKey) || localDup.has(dupKey);
            const isFac = appState.existingKeys.facultyAtTime.has(facKey) || localFac.has(facKey);
            localDup.add(dupKey);
            localFac.add(facKey);

            const tr = document.createElement("tr");
            tr.setAttribute("data-idx", String(idx));
            const computedStatus = isDup ? "Duplicate" : (isFac ? "Faculty conflict" : "OK");
            const finalStatus = (entry && typeof entry === 'object' && entry.status && entry.status !== 'ok') ? entry.status : computedStatus;
            const statusText = (finalStatus || "OK");
            const statusClass = (/duplicate|conflict|error/i).test(statusText) ? "error" : "success";
            tr.innerHTML = `
                <td>${date}</td>
                <td>${batch}</td>
                <td>${shift}</td>
                <td>${faculty}</td>
                <td>${entry.type || 'Regular'}</td>
                <td><span class="status-pill form-status-message ${statusClass}" title="${statusText}">${statusText}</span></td>
                <td><button type="button" data-idx="${idx}" class="remove-row-btn button-secondary">Remove</button></td>
            `;
            pendingTbody.appendChild(tr);
        });
        const anyInvalid = Array.from(pendingTbody.querySelectorAll(".form-status-message.error")).length > 0;
        if (bulkStatus) {
            bulkStatus.textContent = appState.pendingEntries.length === 0 ? "" : (anyInvalid ? "Fix errors before submit." : `${appState.pendingEntries.length} ready.`);
            bulkStatus.className = "form-status-message" + (anyInvalid ? " error" : (appState.pendingEntries.length ? " success" : ""));
        }
        if (bulkSection) bulkSection.classList.toggle("show", appState.pendingEntries.length > 0);
        setSubmitMode(appState.pendingEntries.length > 0 && !anyInvalid);

        pendingTbody.querySelectorAll(".remove-row-btn").forEach(btn => {
            btn.addEventListener("click", (ev) => {
                const i = Number(ev.currentTarget.getAttribute("data-idx"));
                appState.pendingEntries.splice(i, 1);
                renderPendingEntries();
            });
        });
    }

    function setSubmitMode(isBulk) {
        if (isBulk) {
            formSubmitBtn.textContent = "Submit All Pending Rows";
            formSubmitBtn.dataset.mode = "all";
            formSubmitBtn.setAttribute("type", "button"); // avoid native form validation
        } else {
            formSubmitBtn.textContent = "Submit Class";
            formSubmitBtn.dataset.mode = "single";
            formSubmitBtn.setAttribute("type", "submit");
        }
    }

    function getEntryTuple(entry) {
        if (Array.isArray(entry)) return [entry[0], entry[1], entry[2], entry[3], entry[4] || ""];
        return [entry.date, entry.batch, entry.shift, entry.faculty, (entry.type || "")];
    }

    function validateEntry(entry, includePending = true, pendingIndex = -1) {
        const [date, batch, shift, faculty] = getEntryTuple(entry).map(v => String(v || "").trim());
        if (!date || !batch || !shift || !faculty) return { ok: false, message: "All fields are required." };
        const dupKey = `${date}|${batch}|${shift}`;
        const timeCode = getTimeCodeFor(batch, shift);
        const facKey = timeCode ? `${date}|${timeCode}|${faculty}` : `${date}|${shift}|${faculty}`;
        if (appState.existingKeys.duplicate.has(dupKey)) return { ok: false, message: "Slot already scheduled (same Date+Batch+Shift). Use Change Request." };
        if (appState.existingKeys.facultyAtTime.has(facKey)) return { ok: false, message: "Conflict: same Faculty already at this time." };
        if (includePending) {
            for (let i = 0; i < appState.pendingEntries.length; i++) {
                if (i === pendingIndex) continue;
                const [pd, pb, ps, pf] = getEntryTuple(appState.pendingEntries[i]);
                const pTime = getTimeCodeFor(pb, ps);
                if (`${pd}|${pb}|${ps}` === dupKey) return { ok: false, message: "Duplicate in pending list." };
                if ((pTime ? `${pd}|${pTime}|${pf}` : `${pd}|${ps}|${pf}`) === facKey) return { ok: false, message: "Conflict in pending list: faculty busy at this time." };
            }
        }
        return { ok: true, message: "OK" };
    }

    function showFormStatus(message, type) {
        formStatus.textContent = message;
        formStatus.className = `form-status-message ${type}`;
    }

    /**
     * Updates the summary panel with real-time form data
     */
    function updateCrSummary() {
        const summaryBatch = document.getElementById('summary-batch');
        const summaryShift = document.getElementById('summary-shift');
        const summaryValue = document.getElementById('summary-value');
        
        document.getElementById('summary-date').textContent = crDate && crDate.value ? crDate.value : '...';
        document.getElementById('summary-type').textContent = crType && crType.value ? crType.value : '...';
        document.getElementById('summary-remarks').textContent = crRemarks && crRemarks.value ? crRemarks.value : '...';
        if (summaryBatch) summaryBatch.textContent = (crBatch && crBatch.value) ? crBatch.options[crBatch.selectedIndex].text : '...';
        if (summaryShift) summaryShift.textContent = (crShift && crShift.value) ? crShift.options[crShift.selectedIndex].text : '...';
        let newValue = '...';
        const type = crType ? crType.value : '';
        if (type === 'HALL' && crNewHall) newValue = crNewHall.value || '...';
        if (type === 'DURATION' && crNewDuration) newValue = crNewDuration.value || '...';
        if (type === 'TEACHER' && crNewFaculty) newValue = crNewFaculty.value ? crNewFaculty.options[crNewFaculty.selectedIndex].text : '...';
        if (summaryValue) summaryValue.textContent = newValue;
        // Also update the dynamic visibility of Card 5
        if (crNewHallGroup) crNewHallGroup.style.display = (type === 'HALL') ? 'block' : 'none';
        if (crNewDurationGroup) crNewDurationGroup.style.display = (type === 'DURATION') ? 'block' : 'none';
        if (crNewFacultyGroup) crNewFacultyGroup.style.display = (type === 'TEACHER') ? 'block' : 'none';
    }

    /**
     * Applies CSS classes to the card stack based on the current step
     */
    function updateCrCardStackView() {
        const currentStep = appState.crStep;
        document.querySelectorAll('.cr-card').forEach(card => {
            const step = parseInt(card.dataset.step, 10);
            card.classList.remove('active', 'next', 'next-next', 'done');
            if (step === currentStep) {
                card.classList.add('active');
            } else if (step === currentStep + 1) {
                card.classList.add('next');
            } else if (step === currentStep + 2) {
                card.classList.add('next-next');
            } else if (step < currentStep) {
                card.classList.add('done');
            }
        });
    }

    /**
     * Helper to show status messages on the Change Request form
     */
    function showCrStatus(message, type) {
        if (!crStatus) return;
        crStatus.textContent = message;
        crStatus.className = `form-status-message ${type} cr-global`;
        if (type === 'idle') {
            crStatus.style.display = 'none';
        } else {
            crStatus.style.display = 'block';
        }
    }

    // START: New modal functions and listeners
    function showQueueJobModal() {
        if (!queueJobModal) return;
        
        // Reset modal state every time it opens
        if (queueJobStatus) {
            queueJobStatus.textContent = "Click 'Queue Job' to update the Final Schedule.";
            queueJobStatus.className = "form-status-message";
            queueJobStatus.style.display = "block";
        }
        if (queueJobRunBtn) {
            queueJobRunBtn.disabled = false;
            queueJobRunBtn.innerHTML = "Queue Job";
            queueJobRunBtn.style.display = ""; // ensure visible when reopening
        }
        
        // Store branchId for the 'run' button to use
        queueJobModal.dataset.branchId = appState.branchId;
        queueJobModal.style.display = "block";
    }

    if (queueJobCloseBtn) {
        queueJobCloseBtn.addEventListener("click", () => {
            if (queueJobModal) queueJobModal.style.display = "none";
        });
    }

    if (queueJobRunBtn) {
        queueJobRunBtn.addEventListener("click", async () => {
            const branchId = queueJobModal ? queueJobModal.dataset.branchId : '';
            if (!branchId) {
                if (queueJobStatus) {
                    queueJobStatus.textContent = "Error: Branch ID not found.";
                    queueJobStatus.className = "form-status-message error";
                }
                return;
            }

            if (queueJobStatus) {
                queueJobStatus.textContent = "Queuing job...";
                queueJobStatus.className = "form-status-message loading";
            }
            queueJobRunBtn.disabled = true;
            queueJobRunBtn.innerHTML = '<span class="spinner"></span> Queuing...';

            try {
                // Call the new backend action
                const result = await jsonpRequest(API_URL, {
                    action: "queueFinalScheduleJob",
                    branchId: branchId
                });

                if (result.status === "success") {
                    if (queueJobStatus) {
                        queueJobStatus.textContent = "Success! Job is queued. The Final Schedule will update in ~1 minute.";
                        queueJobStatus.className = "form-status-message success";
                    }
                    // After success, hide Queue button so only Close remains
                    queueJobRunBtn.style.display = "none";
                } else {
                    throw new Error(result.message);
                }
            } catch (err) {
                if (queueJobStatus) {
                    queueJobStatus.textContent = "Error: " + err.message;
                    queueJobStatus.className = "form-status-message error";
                }
                queueJobRunBtn.disabled = false;
                queueJobRunBtn.innerHTML = 'Queue Job';
                queueJobRunBtn.style.display = ""; // show again on error
            }
        });
    }
    // END: New modal functions

    // --- Special Lectures (Batch merge) ---
    function initSpecialLectureForm() {
        // --- This part is the same: Populate dropdowns ---
        if (slBatches) {
            slBatches.innerHTML = '';
            const map = buildUniqueBatchMap();
            Array.from(map.entries()).sort((a,b) => a[0].localeCompare(b[0])).forEach(([code, course]) => {
                const opt = document.createElement('option');
                opt.value = code; opt.textContent = `${course} (${code})`;
                slBatches.appendChild(opt);
            });
            if (slBatchOptions) {
                slBatchOptions.innerHTML = '';
                Array.from(map.entries()).sort((a,b) => a[0].localeCompare(b[0])).forEach(([code, course]) => {
                    const div = document.createElement('div');
                    div.className = 'multi-option';
                    const id = `slb-${code}`;
                    div.innerHTML = `<input type="checkbox" id="${id}" value="${code}"><label for="${id}">${course} (${code})</label>`;
                    slBatchOptions.appendChild(div);
                });
            }
            updateBatchTriggerText();
        }
        if (slFaculty) {
            slFaculty.innerHTML = '<option value="">-- Select Faculty --</option>';
            (appState.faculty || []).forEach(f => {
                const opt = document.createElement('option');
                opt.value = f.code;
                opt.textContent = `${f.name} (${f.expertise})`;
                slFaculty.appendChild(opt);
            });
        }
        // --- END Populate dropdowns ---

        // --- NEW: Card Stack Logic ---
        appState.slStep = 1; // Reset to step 1
        updateSlCardStackView(); // Set initial card positions
        updateSlSummary(); // Clear summary

        // Add 'input' listener to the whole form for real-time updates
        if (slForm) {
            slForm.addEventListener('input', updateSlSummary);
        }

        // Add 'click' listeners to all "Next" buttons
        document.querySelectorAll('.sl-next-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const currentStep = appState.slStep;
                
                // --- Simple Validation ---
                let isValid = true;
                if (currentStep === 1 && !slDate.value) isValid = false;
                if (currentStep === 2 && Array.from(slBatches.selectedOptions).length === 0) isValid = false;
                if (currentStep === 3 && !slFaculty.value) isValid = false;
                if (currentStep === 4 && (!slHall.value || !slTime.value)) isValid = false;
                if (currentStep === 5 && (!slDuration.value || !slTopic.value)) isValid = false;
                if (currentStep === 6 && !slPayment.value) isValid = false;
                if (currentStep === 7 && !slReason.value) isValid = false; // <-- ADD THIS
                
                if (!isValid) {
                    showSlStatus("Please complete all fields on this card.", "error");
                    return;
                }
                // --- End Validation ---
                
                showSlStatus("", "idle"); // Clear error
                const nextStep = parseInt(e.currentTarget.dataset.next, 10);
                appState.slStep = nextStep;
                updateSlCardStackView();
            });
        });

        // The old multi-select logic is still needed
        updateCoursesFromSelectedBatches();
    }

    function buildUniqueBatchMap() {
        const map = new Map();
        (appState.batchLectures || []).forEach(b => {
            const code = String(b.batchCode || b['Batch Code'] || '').trim();
            const course = String(b.course || b.Course || '').trim();
            if (code && !map.has(code)) map.set(code, course);
        });
        return map;
    }

    function updateCoursesFromSelectedBatches() {
        if (!slBatches || !slCourses) return;
        const selected = Array.from(slBatches.selectedOptions).map(o => o.value);
        const codeToCourse = new Map();
        (appState.batchLectures || []).forEach(b => {
            codeToCourse.set(String(b.batchCode || b['Batch Code'] || '').trim(), String(b.course || b.Course || '').trim());
        });
        const names = Array.from(new Set(selected.map(code => codeToCourse.get(code) || code))).filter(Boolean);
        slCourses.value = names.join(' & ');
    }

    if (slBatches) {
        slBatches.addEventListener('change', updateCoursesFromSelectedBatches);
    }

    // Multi-select dropdown behaviors
    function updateSelectedFromMenu() {
        if (!slBatches || !slBatchOptions) return;
        const selectedCodes = Array.from(slBatchOptions.querySelectorAll('input[type="checkbox"]:checked')).map(i => i.value);
        // Reflect to hidden select
        Array.from(slBatches.options).forEach(o => { o.selected = selectedCodes.includes(o.value); });
        updateBatchTriggerText();
        updateCoursesFromSelectedSchedules();
    }

    function updateBatchTriggerText() {
        if (!slBatchTrigger || !slBatches) return;
        const count = Array.from(slBatches.selectedOptions).length;
        slBatchTrigger.textContent = count > 0 ? `${count} batch${count===1?'':'es'} selected` : 'Select batches';
    }

    if (slBatchTrigger) {
        slBatchTrigger.addEventListener('click', () => {
            if (!slBatchMenu) return;
            const isOpen = slBatchMenu.style.display === 'block';
            slBatchMenu.style.display = isOpen ? 'none' : 'block';
            // Sync checkboxes from current selection when opening
            if (!isOpen && slBatchOptions && slBatches) {
                const selected = new Set(Array.from(slBatches.selectedOptions).map(o => o.value));
                slBatchOptions.querySelectorAll('input[type="checkbox"]').forEach(ch => {
                    ch.checked = selected.has(ch.value);
                });
            }
        });
    }

    if (slBatchSearch && slBatchOptions) {
        slBatchSearch.addEventListener('input', () => {
            const q = slBatchSearch.value.toLowerCase();
            slBatchOptions.querySelectorAll('.multi-option').forEach(div => {
                const txt = div.textContent.toLowerCase();
                div.style.display = txt.includes(q) ? '' : 'none';
            });
        });
    }

    const slBatchApply = document.getElementById('sl-batch-apply');
    const slBatchClear = document.getElementById('sl-batch-clear');
    if (slBatchApply) slBatchApply.addEventListener('click', () => { updateSelectedFromMenu(); updateSlSummary(); if (slBatchMenu) slBatchMenu.style.display='none'; });
    if (slBatchClear) slBatchClear.addEventListener('click', () => {
        if (!slBatchOptions) return;
        slBatchOptions.querySelectorAll('input[type="checkbox"]').forEach(ch => ch.checked = false);
        updateSelectedFromMenu();
        updateSlSummary();
    });

    document.addEventListener('click', (e) => {
        if (!slBatchMenu || !slBatchTrigger) return;
        // Check if click is outside the multi-select component
        if (e.target.closest('#sl-batch-multi')) return;
        slBatchMenu.style.display = 'none';
    });

    if (slFaculty) {
        slFaculty.addEventListener('change', () => {
            const code = slFaculty.value;
            const f = (appState.faculty || []).find(x => x.code === code);
            if (slExpertise) slExpertise.value = f ? (f.expertise || '') : '';
        });
    }

    // Auto-fill substitute subject when picking substitute faculty in modal
    if (vsSubFaculty) {
        vsSubFaculty.addEventListener('change', () => {
            const code = vsSubFaculty.value;
            const f = (appState.faculty || []).find(x => x.code === code);
            if (vsSubSubject) vsSubSubject.value = f ? (f.expertise || '') : '';
        });
    }

    if (slForm) {
        slForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            showSlStatus("Submitting special lecture...", "loading");
            
            const selectedBatchCodes = Array.from(slBatches.selectedOptions).map(o => o.value).filter(Boolean);
            
            if (!slDate.value || selectedBatchCodes.length === 0 || !slFaculty.value || !slHall.value || !slTime.value || !slDuration.value || !slTopic.value || !slReason.value) { // <-- ADDED !slReason.value
                showSlStatus("Please fill all required fields.", "error");
                // Send user back to the first step
                appState.slStep = 1; 
                updateSlCardStackView();
                return;
            }
            
            try {
                const res = await jsonpRequest(API_URL, {
                    action: 'logSpecialLecture',
                    branchId: appState.branchId,
                    lectDate: slDate.value,
                    batches: selectedBatchCodes.join(','),
                    courses: slCourses.value || '',
                    facultyCode: slFaculty.value,
                    expertise: slExpertise.value || '',
                    hallNo: slHall.value,
                    lectureTime: slTime.value, // HH:mm (24h)
                    lectureDuration: slDuration.value,
                    lectureTopic: slTopic.value,
                    payment: slPayment.value,
                    reason: slReason.value // <-- ADD THIS
                });
                if (res.status === 'success') {
                    showSlStatus("Logged successfully.", "success");
                    slForm.reset(); 
                    // Clear multi-select
                    Array.from(slBatches.options).forEach(o => { o.selected = false; });
                    updateBatchTriggerText();
                    
                    // Reset card stack to beginning
                    appState.slStep = 1; 
                    updateSlCardStackView();
                    updateSlSummary(); // Clear summary
                } else {
                    throw new Error(res.message || 'Error logging special lecture');
                }
            } catch (err) {
                showSlStatus("Error: " + err.message, "error");
            }
        });
    }

    // --- 7. "View Schedule" Page Logic ---

    let vsWindowOffset = 0; // default window (today-2 .. today)
    function formatYmdLocal(date) {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }
    function parseYmdLocal(ymd) {
        const [y, m, d] = (ymd || '').split('-').map(n => parseInt(n, 10));
        return new Date(y, (m || 1) - 1, d || 1);
    }
    function getWindowDates(offset){
        const today = new Date();
        const start = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 2 + (offset*3));
        const end = new Date(today.getFullYear(), today.getMonth(), today.getDate() + (offset*3));
        return { start: formatYmdLocal(start), end: formatYmdLocal(end) };
    }

    async function loadViewScheduleData() {
        const { start, end } = getWindowDates(vsWindowOffset);
        if (vsWindowLabel) vsWindowLabel.textContent = `Window: ${start} → ${end}`;
        
        // Also set filter inputs to this window
        if (filterDateStart) filterDateStart.value = start;
        if (filterDateEnd) filterDateEnd.value = end;

        try {
            viewScheduleStatus.innerHTML = '<span class="spinner"></span> Loading schedule...';
            viewScheduleStatus.className = "form-status-message loading";
            viewScheduleStatus.style.display = "block";

            const result = await jsonpRequest(API_URL, {
                action: "getFinalScheduleByDateRange",
                branchId: appState.branchId,
                startDate: start,
                endDate: end,
                includeStatus: '1'
            });

            if (result.status === "success") {
                appState.finalScheduleData = result.data || [];
                appState.isScheduleLoaded = true;
                appState.statusMap = result.statusMap || {};

                // --- SORTING LOGIC ADDED ---
                appState.finalScheduleData.sort((a, b) => {
                    const dateComparison = String(a.LectureDate).localeCompare(String(b.LectureDate));
                    if (dateComparison !== 0) {
                        return dateComparison; // Sort by date
                    }
                    // Dates are same, sort by time
                    return parseScheduleTime(a.StartTime) - parseScheduleTime(b.StartTime);
                });
                // --- END OF SORTING LOGIC ---

                renderScheduleTable(appState.finalScheduleData);
                populateFilterDropdowns(appState.finalScheduleData);
                viewScheduleStatus.style.display = "none";
                viewScheduleTable.classList.add("loaded");
                
                // Clear and apply filters
                filterBatch.value = "";
                filterFaculty.value = "";
                filterHall.value = "";
                applyTableFilters();
            } else {
                throw new Error(result.message);
            }
        } catch (err) {
            console.error("Error loading final schedule:", err);
            viewScheduleStatus.textContent = "Error: " + err.message;
            viewScheduleStatus.className = "form-status-message error";
        }
    }

    async function loadViewScheduleForRange(start, end) {
        if (vsWindowLabel) vsWindowLabel.textContent = `Custom: ${start} → ${end}`;
        try {
            viewScheduleStatus.innerHTML = '<span class="spinner"></span> Loading schedule...';
            viewScheduleStatus.className = 'form-status-message loading';
            viewScheduleStatus.style.display = 'block';
            const result = await jsonpRequest(API_URL, {
                action: 'getFinalScheduleByDateRange',
                branchId: appState.branchId,
                startDate: start,
                endDate: end,
                includeStatus: '1'
            });
            if (result.status === 'success') {
                appState.finalScheduleData = result.data || [];
                appState.statusMap = result.statusMap || {};

                // --- SORTING LOGIC ADDED ---
                appState.finalScheduleData.sort((a, b) => {
                    const dateComparison = String(a.LectureDate).localeCompare(String(b.LectureDate));
                    if (dateComparison !== 0) {
                        return dateComparison; // Sort by date
                    }
                    // Dates are same, sort by time
                    return parseScheduleTime(a.StartTime) - parseScheduleTime(b.StartTime);
                });
                // --- END OF SORTING LOGIC ---

                renderScheduleTable(appState.finalScheduleData);
                populateFilterDropdowns(appState.finalScheduleData);
                viewScheduleStatus.style.display = 'none';
                
                // Clear and apply filters
                filterBatch.value = "";
                filterFaculty.value = "";
                filterHall.value = "";
                applyTableFilters();
            } else {
                throw new Error(result.message);
            }
        } catch (err) {
            viewScheduleStatus.textContent = 'Error: ' + err.message;
            viewScheduleStatus.className = 'form-status-message error';
        }
    }

    function renderScheduleTable(data) {
        viewScheduleTbody.innerHTML = ""; 
        if (data.length === 0) {
            viewScheduleStatus.textContent = "No schedule data found for this range.";
            viewScheduleStatus.className = "form-status-message";
            viewScheduleStatus.style.display = "block";
            viewScheduleTable.classList.remove("loaded");
            return;
        }
        const canMark = (dateStr) => {
            // Eligible only for today and past 2 days regardless of what is being viewed
            const today = new Date();
            const base = new Date(today.getFullYear(), today.getMonth(), today.getDate());
            const d = parseYmdLocal(dateStr);
            const diff = Math.floor((d.getTime() - base.getTime()) / (1000*60*60*24));
            return diff <= 0 && diff >= -2;
        };
        data.forEach(row => {
            const tr = document.createElement("tr");
            const statusKey = `${row.LectureDate}|${row.Batch}|${row.ShiftCode}|${row.FacultyCode}`;
            const statusVal = (appState.statusMap && appState.statusMap[statusKey]) || '';
            const badgeHtml = (st) => {
                if (!st) return '';
                const s = String(st).toLowerCase();
                if (s === 'done') return '<span class="badge badge-done">Done</span>';
                if (s === 'cancelled') return '<span class="badge badge-cancelled">Cancelled</span>';
                if (s === 'substituted') return '<span class="badge badge-substituted">Substituted</span>';
                return `<span class="badge">${st}</span>`;
            };
            tr.innerHTML = `
                <td>${row.LectureDate || ''}</td>
                <td>${row.StartTime || ''}</td>
                <td>${row.Course || ''}</td>
                <td>${row.Batch || ''}</td>
                <td>${row.HallNo || ''}</td>
                <td>${row.faculty_name || ''}</td>
                <td>${row.Expertise || ''}</td>
                <td>${row.LectureTopic || ''}</td>
                <td>${row['Lecture Count'] || ''}</td> 
                <td>${row.Payment || ''}</td>
                <td class="vs-status">${badgeHtml(statusVal)}</td>
                <td class="vs-actions"></td>
            `;

            // Add data attributes for filtering
            tr.dataset.date = row.LectureDate || '';
            tr.dataset.batch = row.Batch || '';
            tr.dataset.faculty = row.faculty_name || '';
            tr.dataset.hall = row.HallNo || '';
            
            viewScheduleTbody.appendChild(tr);

            // Actions
            const actionsTd = tr.querySelector('.vs-actions');
            const key = statusKey;
            vsRowContext[key] = row;
            if (canMark(row.LectureDate)) {
                const btn = document.createElement('button');
                btn.className = 'vs-action-btn';
                btn.dataset.key = key;
                if (!statusVal) {
                    btn.textContent = 'Action Required';
                    btn.style.backgroundColor = '#dc2626';
                    btn.style.color = '#fff';
                    btn.addEventListener('click', () => openStatusModal(key));
                } else {
                    btn.textContent = 'Action Taken';
                    btn.disabled = true;
                }
                actionsTd.appendChild(btn);
            }
        });
    }

    function populateFilterDropdowns(data) {
        const batches = new Set();
        const faculty = new Set();
        const halls = new Set();

        data.forEach(row => {
            if (row.Batch) batches.add(row.Batch);
            if (row.faculty_name) faculty.add(row.faculty_name);
            if (row.HallNo) halls.add(row.HallNo);
        });

        // Populate Batch Filter
        filterBatch.innerHTML = '<option value="">All Batches</option>';
        [...batches].sort().forEach(val => { // Sort
            const option = document.createElement("option");
            option.value = val;
            option.textContent = val;
            filterBatch.appendChild(option);
        });

        // Populate Faculty Filter
        filterFaculty.innerHTML = '<option value="">All Faculty</option>';
        [...faculty].sort().forEach(val => { // Sort
            const option = document.createElement("option");
            option.value = val;
            option.textContent = val;
            filterFaculty.appendChild(option);
        });

        // Populate Hall Filter
        filterHall.innerHTML = '<option value="">All Halls</option>';
        [...halls].sort().forEach(val => { // Sort
            const option = document.createElement("option");
            option.value = val;
            option.textContent = val;
            filterHall.appendChild(option);
        });
    }

    // --- Class status APIs (Done / Cancelled / Substituted) ---
    function renderStatusBadge(st){
        if (!st) return '';
        const s = String(st).toLowerCase();
        if (s === 'done') return '<span class="badge badge-done">Done</span>';
        if (s === 'cancelled') return '<span class="badge badge-cancelled">Cancelled</span>';
        if (s === 'substituted') return '<span class="badge badge-substituted">Substituted</span>';
        return `<span class="badge">${st}</span>`;
    }

    async function markClassStatus(key, status, extra = {}) {
        const row = vsRowContext[key]; if (!row) return;
        try {
            const res = await jsonpRequest(API_URL, Object.assign({
                action: 'markClassStatus',
                branchId: appState.branchId,
                lectDate: row.LectureDate,
                batch: row.Batch,
                shift: row.ShiftCode,
                faculty: row.FacultyCode,
                status: status,
                markedBy: (appState.userEmail || '')
            }, extra));
            if (res.status === 'success') {
                appState.statusMap = appState.statusMap || {}; appState.statusMap[key] = status;
                // Update UI in-place: disable button and set badge
                const btn = document.querySelector(`.vs-action-btn[data-key="${key}"]`);
                if (btn) { btn.textContent = 'Action Taken'; btn.disabled = true; btn.style.backgroundColor=''; btn.style.color=''; } // Reset style
                const statusCell = btn ? btn.parentElement.previousElementSibling : null;
                if (statusCell) statusCell.innerHTML = renderStatusBadge(status);
            } else {
                alert(res.message || 'Error');
            }
        } catch (e) {
            alert(e.message);
        }
    }

    function openStatusModal(key) {
        if (!vsModal) return; vsModal.style.display = 'block';
        vsModal.dataset.key = key;
        if (vsSubFaculty) {
            vsSubFaculty.innerHTML = '<option value="">-- Select Faculty --</option>';
            (appState.faculty || []).forEach(f => {
                const opt = document.createElement('option');
                opt.value = f.code; opt.textContent = `${f.name} (${f.expertise})`;
                vsSubFaculty.appendChild(opt);
            });
        }
        // Reset form
        if (vsModal.querySelector('input[name="vs-action"]')) {
             vsModal.querySelector('input[name="vs-action"]').checked = true;
        }
        if (vsReason) vsReason.value = '';
        if (vsSubFaculty) vsSubFaculty.value = '';
        if (vsSubSubject) vsSubSubject.value = '';
        if (vsSubTopic) vsSubTopic.value = '';
        if (vsNewHall) vsNewHall.value = '';
        if (vsSubReason) vsSubReason.value = '';
        
        toggleModalFields();
    }
    function closeStatusModal() { if (vsModal) vsModal.style.display = 'none'; }
    function toggleModalFields() {
        if (!vsModal) return; const v = (vsModal.querySelector('input[name="vs-action"]:checked') || {}).value;
        if (vsCancelFields) vsCancelFields.style.display = (v === 'Cancelled') ? 'block' : 'none';
        if (vsSubFields) vsSubFields.style.display = (v === 'Substituted') ? 'block' : 'none';
    }
    document.addEventListener('change', (e) => { if (e.target && e.target.name === 'vs-action') toggleModalFields(); });
    if (vsModalCancel) vsModalCancel.addEventListener('click', closeStatusModal);
    if (vsModalSave) vsModalSave.addEventListener('click', async () => {
        const key = vsModal ? vsModal.dataset.key : '';
        const actionEl = vsModal ? vsModal.querySelector('input[name="vs-action"]:checked') : null;
        const action = actionEl ? actionEl.value : '';
        if (!key || !action) { closeStatusModal(); return; }
        // show saving state
        const oldText = vsModalSave.textContent; vsModalSave.disabled = true; vsModalSave.innerHTML = '<span class="spinner"></span> Saving...';
        if (action === 'Done') { await markClassStatus(key, 'Done'); closeStatusModal(); vsModalSave.disabled=false; vsModalSave.textContent=oldText; return; }
        if (action === 'Cancelled') { await markClassStatus(key, 'Cancelled', { reason: (vsReason && vsReason.value) || '' }); closeStatusModal(); vsModalSave.disabled=false; vsModalSave.textContent=oldText; return; }
        if (action === 'Substituted') {
            if (!vsSubFaculty || !vsSubFaculty.value) { alert('Select substitute faculty'); vsModalSave.disabled=false; vsModalSave.textContent=oldText; return; }
            await markClassStatus(key, 'Substituted', { subFaculty: vsSubFaculty.value, subTopic: (vsSubTopic && vsSubTopic.value) || '', newHall: (vsNewHall && vsNewHall.value) || '', subSubject: (vsSubSubject && vsSubSubject.value) || '', reason: (vsSubReason && vsSubReason.value) || '' });
            closeStatusModal(); vsModalSave.disabled=false; vsModalSave.textContent=oldText; return;
        }
    });

    function applyTableFilters() {
        // const startDate = filterDateStart.value; // Date filtering is now done by API call
        // const endDate = filterDateEnd.value;
        const batch = filterBatch.value;
        const faculty = filterFaculty.value;
        const hall = filterHall.value;

        const allRows = viewScheduleTbody.querySelectorAll("tr");

        allRows.forEach(tr => {
            // const rowDate = tr.dataset.date; // Not needed
            const rowBatch = tr.dataset.batch;
            const rowFaculty = tr.dataset.faculty;
            const rowHall = tr.dataset.hall;

            // Check each filter
            // const dateMatch = (!startDate || rowDate >= startDate) && (!endDate || rowDate <= endDate); // Handled by API
            const batchMatch = !batch || rowBatch === batch;
            const facultyMatch = !faculty || rowFaculty === faculty;
            const hallMatch = !hall || rowHall === hall;

            // Show or hide the row
            if (/*dateMatch &&*/ batchMatch && facultyMatch && hallMatch) {
                tr.classList.remove("hidden-filter");
            } else {
                tr.classList.add("hidden-filter");
            }
        });
    }

    // Add event listeners to all filter inputs
    [filterBatch, filterFaculty, filterHall].forEach(input => {
        input.addEventListener('change', applyTableFilters);
    });
    
    function tryFetchByDateRange() {
        const s = filterDateStart ? filterDateStart.value : '';
        const e = filterDateEnd ? filterDateEnd.value : '';
        if (s && e) {
            // Both selected: normalize order
            const sD = parseYmdLocal(s); const eD = parseYmdLocal(e);
            const sStr = formatYmdLocal(sD <= eD ? sD : eD);
            const eStr = formatYmdLocal(eD >= sD ? eD : sD);
            if (filterDateStart.value !== sStr) filterDateStart.value = sStr;
            if (filterDateEnd.value !== eStr) filterDateEnd.value = eStr;
            vsWindowOffset = 0; // Reset pagination
            loadViewScheduleForRange(sStr, eStr);
            return;
        }
        // Do not auto-fetch if only one date is set
    }
    
    if (filterDateStart) {
        filterDateStart.addEventListener('change', tryFetchByDateRange);
    }
    if (filterDateEnd) {
        filterDateEnd.addEventListener('change', tryFetchByDateRange);
    }

    filterResetBtn.addEventListener("click", () => {
        filterBatch.value = "";
        filterFaculty.value = "";
        filterHall.value = "";
        // Reset to default 3-day window
        vsWindowOffset = 0;
        loadViewScheduleData();
    });

    // Pagination controls
    if (vsPrev) vsPrev.addEventListener('click', ()=>{ vsWindowOffset -= 1; loadViewScheduleData(); });
    if (vsNext) vsNext.addEventListener('click', ()=>{ vsWindowOffset += 1; loadViewScheduleData(); });
    if (vsToday) vsToday.addEventListener('click', ()=>{ vsWindowOffset = 0; loadViewScheduleData(); });

    // Mark all visible as done
    if (vsMarkAll) vsMarkAll.addEventListener('click', async ()=>{
        // Get all VISIBLE (not filtered out) buttons
        const btns = Array.from(viewScheduleTbody.querySelectorAll('tr:not(.hidden-filter) .vs-actions .vs-action-btn:not([disabled])'));
        if (btns.length === 0) {
            alert("No visible rows require action.");
            return;
        }
        if (!confirm(`This will mark ${btns.length} visible class(es) as 'Done'. Are you sure?`)) {
            return;
        }
        
        const oldHtml = vsMarkAll.innerHTML;
        vsMarkAll.disabled = true;
        vsMarkAll.innerHTML = '<span class="spinner"></span> Marking...';
        try {
            for (const btn of btns) {
                const key = btn.dataset.key;
                if (key) await markClassStatus(key, 'Done');
            }
        } finally {
            vsMarkAll.disabled = false;
            vsMarkAll.innerHTML = oldHtml;
        }
    });
    
    if (timelineRefreshBtn) {
        timelineRefreshBtn.addEventListener("click", loadHallTimeline);
    }

    // --- NEW: Add listener for the timeline batch filter ---
    if (timelineBatchFilter) {
        timelineBatchFilter.addEventListener("change", applyHallTimelineFilters);
    }

    // --- NEW: Add listener for the timeline faculty filter ---
    if (timelineFacultyFilter) {
        timelineFacultyFilter.addEventListener("change", applyHallTimelineFilters);
    }
    // --- END NEW ---
    // --- END NEW ---

    
    // --- NEW: ANALYSIS DASHBOARD LOGIC ---

    /**
     * Helper to get default YYYY-MM-DD date string
     */
    function getYmd(date) {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }

    /**
     * Fetches and renders all data for the analysis dashboard
     */
    async function loadDashboardAnalytics() {
        if (!analysisStatus) return; // Page elements not found
        
        appState.isDashboardLoaded = false; // Mark as loading
        analysisStatus.innerHTML = '<span class="spinner"></span> Loading dashboard analytics...';
        analysisStatus.className = "form-status-message loading";
        analysisStatus.style.display = "block";

        try {
            // Set default date range (last 30 days) if empty
            const today = new Date();
            const todayStr = getYmd(today);
            const thirtyDaysAgo = new Date(); // create new date obj
            thirtyDaysAgo.setDate(today.getDate() - 30); // set it
            const thirtyDaysAgoStr = getYmd(thirtyDaysAgo);
            
            if (!analysisDateStart.value) analysisDateStart.value = thirtyDaysAgoStr;
            if (!analysisDateEnd.value) analysisDateEnd.value = todayStr;

            // Get filter values
            const params = {
                action: "getDashboardAnalytics",
                branchId: appState.branchId,
                startDate: analysisDateStart.value,
                endDate: analysisDateEnd.value,
                batch: analysisBatch.value || ''
            };

            const result = await jsonpRequest(API_URL, params);

            if (result.status === "success") {
                const data = result.data;

                // Populate KPI Cards
                kpiPlannedClasses.textContent = data.plannedClasses;
                kpiDoneClasses.textContent = data.doneClasses;
                kpiFillRate.textContent = data.fillRate + '%';
                kpiCancelledClasses.textContent = data.cancelledClasses;
                kpiCancellationRate.textContent = data.cancellationRate + '%';
                kpiSubstitutedClasses.textContent = data.substitutedClasses;
                kpiRunningBatches.textContent = data.runningBatches;
                kpiTotalBatches.textContent = data.totalBatchesLaunched;
                kpiTotalChanges.textContent = data.totalChanges;
                kpiTotalSpecial.textContent = data.totalSpecial;

                // Populate Tables
                renderAnalysisTable(analysisPivotTable, ['Batch', 'Last Class Date'], data.batchScheduledTill.map(r => [r.batch, r.lastClass]));
                renderAnalysisTable(analysisTopFaculty, ['Faculty', 'Done Classes'], data.mostUtilizedFaculty.map(r => [r.faculty, r.count]));
                renderAnalysisTable(analysisCancelFaculty, ['Faculty', 'Cancelled Classes'], data.mostCancelledFaculty.map(r => [r.faculty, r.count]));
                
                appState.isDashboardLoaded = true;
                analysisStatus.style.display = "none";
            } else {
                throw new Error(result.message);
            }
        } catch (err) {
            console.error("Error loading dashboard:", err);
            analysisStatus.textContent = "Error: " + err.message;
            analysisStatus.className = "form-status-message error";
        }
    }
    
    /**
     * Renders a simple HTML table for the analysis page
     */
    function renderAnalysisTable(container, headers, rows) {
        if (!container) return;
        
        let html = '<table class="pending-table">';
        // Header
        html += '<thead><tr>';
        headers.forEach(h => html += `<th>${h}</th>`);
        html += '</tr></thead>';
        
        // Body
        html += '<tbody>';
        if (rows.length === 0) {
            html += `<tr><td colspan="${headers.length}" style="text-align:center; color:#777;">No data found.</td></tr>`;
        } else {
            rows.forEach(row => {
                html += '<tr>';
                row.forEach(cell => html += `<td>${cell}</td>`);
                html += '</tr>';
            });
        }
        html += '</tbody></table>';
        
        container.innerHTML = html;
    }

    // Add event listener to the new filter button
    if (analysisApplyBtn) {
        analysisApplyBtn.addEventListener("click", loadDashboardAnalytics);
    }

    // --- END NEW ANALYSIS LOGIC ---
    
    // --- NEW: Special Lecture Card Stack Helpers ---
    /**
     * Updates the SL summary panel with real-time form data
     */
    function updateSlSummary() {
        // Update auto-fill hidden fields first
        updateCoursesFromSelectedBatches();
        const facultyCode = slFaculty ? slFaculty.value : '';
        const facultyObj = (appState.faculty || []).find(x => x.code === facultyCode);
        if (slExpertise) slExpertise.value = facultyObj ? (facultyObj.expertise || '') : '';
        // Update summary panel UI
        const el = (id) => document.getElementById(id);
        if (el('summary-sl-date')) el('summary-sl-date').textContent = (slDate && slDate.value) || '...';
        const selectedCount = slBatches ? Array.from(slBatches.selectedOptions).length : 0;
        if (el('summary-sl-batches')) el('summary-sl-batches').textContent = selectedCount > 0 ? `${selectedCount} batch(es)` : '...';
        if (el('summary-sl-faculty')) el('summary-sl-faculty').textContent = (slFaculty && slFaculty.value) ? slFaculty.options[slFaculty.selectedIndex].text : '...';
        if (el('summary-sl-expertise')) el('summary-sl-expertise').textContent = (slExpertise && slExpertise.value) || '...';
        if (el('summary-sl-hall')) el('summary-sl-hall').textContent = (slHall && slHall.value) || '...';
        if (el('summary-sl-time')) el('summary-sl-time').textContent = (slTime && slTime.value) || '...';
        if (el('summary-sl-duration')) el('summary-sl-duration').textContent = (slDuration && slDuration.value) ? `${slDuration.value} hrs` : '...';
        if (el('summary-sl-topic')) el('summary-sl-topic').textContent = (slTopic && slTopic.value) || '...';
        if (el('summary-sl-payment')) el('summary-sl-payment').textContent = (slPayment && slPayment.value) || '...';
        if (el('summary-sl-reason')) el('summary-sl-reason').textContent = (slReason && slReason.value) || '...'; // <-- ADD THIS
    }

    /**
     * Applies CSS classes to the SL card stack based on the current step
     */
    function updateSlCardStackView() {
        const currentStep = appState.slStep;
        document.querySelectorAll('.sl-card').forEach(card => {
            const step = parseInt(card.dataset.step, 10);
            card.classList.remove('active', 'next', 'next-next', 'done');
            if (step === currentStep) {
                card.classList.add('active');
            } else if (step === currentStep + 1) {
                card.classList.add('next');
            } else if (step === currentStep + 2) {
                card.classList.add('next-next');
            } else if (step < currentStep) {
                card.classList.add('done');
            }
        });
    }

    /**
     * Helper to show status messages on the Special Lecture form
     */
    function showSlStatus(message, type) {
        if (!slStatus) return;
        slStatus.textContent = message;
        slStatus.className = `form-status-message ${type} sl-global`;
        if (type === 'idle') {
            slStatus.style.display = 'none';
        } else {
            slStatus.style.display = 'block';
        }
    }

    // --- NEW: HALL TIMELINE LOGIC ---

    /**
     * Fetches schedule data for a specific date and renders the hall timeline grid.
     */
    async function loadHallTimeline() {
        if (!timelineDate || !timelineStatus || !timelineContainer) return;

        if (!appState.isLoggedIn || !appState.branchId) {
            timelineStatus.textContent = "Please log in to view the hall timeline.";
            timelineStatus.className = "form-status-message error";
            timelineStatus.style.display = "block";
            return;
        }

        if (!timelineDate.value) {
            timelineStatus.textContent = "Please select a date.";
            timelineStatus.className = "form-status-message error";
            timelineStatus.style.display = "block";
            return;
        }

        if (timelineBatchFilter) timelineBatchFilter.value = "";
        if (timelineFacultyFilter) timelineFacultyFilter.value = "";

        timelineStatus.innerHTML = '<span class="spinner"></span> Loading timeline...';
        timelineStatus.className = "form-status-message loading";
        timelineStatus.style.display = "block";
        timelineContainer.innerHTML = "";

        try {
            const result = await jsonpRequest(API_URL, {
                action: "getTimelineData",
                branchId: appState.branchId,
                lectDate: timelineDate.value
            });

            if (result.status === "success") {
                const lectures = Array.isArray(result.data) ? result.data : [];
                const halls = Array.isArray(result.halls) ? result.halls : [];

                renderHallTimeline(lectures, halls);
                populateTimelineBatchFilter(lectures);
                populateTimelineFacultyFilter(lectures);
                applyHallTimelineFilters();

                timelineStatus.style.display = "none";
            } else {
                throw new Error(result.message || "Failed to load timeline data.");
            }
        } catch (err) {
            console.error("Error loading timeline:", err);
            timelineStatus.textContent = "Error: " + err.message;
            timelineStatus.className = "form-status-message error";
            timelineStatus.style.display = "block";
            if (timelineContainer.innerHTML.trim() === "") {
                timelineContainer.innerHTML = '<p style="color: var(--theme-text-secondary);">Unable to load timeline data. Please try again later.</p>';
            }
        }
    }

    /**
     * Renders the visual timeline grid as an HTML table.
     * @param {Array} lectures - The list of lecture objects for the day.
     * @param {Array} halls - The list of all hall names (strings).
     */
    function renderHallTimeline(lectures, halls) {
        if (!timelineContainer) return;

        const safeLectures = Array.isArray(lectures) ? lectures : [];
        const safeHalls = Array.isArray(halls) ? halls.filter(Boolean) : [];

        if (safeHalls.length === 0) {
            timelineContainer.innerHTML = '<p>No halls found for this branch.</p>';
            return;
        }

        // --- 1. Define Timeline Boundaries ---
        const timelineStartHour = 7; // 7:00 AM
        const timelineEndHour = 21; // 9:00 PM (ends at 21:00)
        const totalHours = timelineEndHour - timelineStartHour; // 14 hours
        const totalMinutes = totalHours * 60; // 840 minutes
        const timelineStartMinutes = timelineStartHour * 60; // 420

        const getMinutesFromStart = (timeStr) => {
            const minutes = parseScheduleTime(timeStr);
            return minutes - timelineStartMinutes;
        };

        // --- 2. Render the Empty Grid ---
        let html = '<table class="timeline-table">';
        html += '<thead><tr><th>Hall</th>';

        const timeSlots = [];
        for (let hour = timelineStartHour; hour < timelineEndHour; hour++) {
            const hourStr = String(hour).padStart(2, '0');
            timeSlots.push(`${hourStr}:00`);
        }
        timeSlots.forEach(time => {
            html += `<th>${time}</th>`;
        });
        html += '</tr></thead>';

        html += '<tbody>';
        safeHalls.forEach(hall => {
            html += `<tr>
                <th>${hall}</th>
                <td class="timeline-row-content" data-hall-container="${hall}" colspan="${timeSlots.length}"></td>
            </tr>`;
        });
        html += '</tbody>';
        html += '</table>';

        timelineContainer.innerHTML = html;

        const gridTable = timelineContainer.querySelector('.timeline-table');
        if (!gridTable) return;

        // --- 3. Place Lecture Blocks ---
        safeLectures.forEach(lecture => {
            try {
                const hall = String(lecture.HallNo || '').trim();
                if (!hall) return;

                const containerCell = gridTable.querySelector(`td[data-hall-container="${hall}"]`);
                if (!containerCell) return;

                const parsedDuration = parseFloat(lecture.LectureDuration || lecture.Lectureduration || lecture.duration || '1');
                const durationMinutes = isNaN(parsedDuration) ? 60 : parsedDuration * 60;

                const rawStartOffset = getMinutesFromStart(lecture.StartTime);
                const rawEndOffset = rawStartOffset + durationMinutes;

                const clampedStart = Math.max(0, rawStartOffset);
                const clampedEnd = Math.min(totalMinutes, rawEndOffset);

                if (clampedEnd <= 0 || clampedStart >= totalMinutes) return; // Completely outside window

                const effectiveDuration = clampedEnd - clampedStart;
                if (effectiveDuration <= 0) return;

                const leftPercentage = (clampedStart / totalMinutes) * 100;
                const widthPercentage = (effectiveDuration / totalMinutes) * 100;

                const block = document.createElement('div');
                block.className = 'lecture-block';
                block.style.left = `${leftPercentage}%`;
                block.style.width = `${widthPercentage}%`;

                const formatMinutes = (minutes) => {
                    const hrs = Math.floor(minutes / 60);
                    const mins = Math.floor(minutes % 60);
                    return `${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
                };

                const tooltipStart = formatMinutes(Math.max(parseScheduleTime(lecture.StartTime), timelineStartMinutes));
                const tooltipEnd = formatMinutes(Math.min(parseScheduleTime(lecture.StartTime) + durationMinutes, timelineStartMinutes + totalMinutes));
                block.setAttribute('title', `${tooltipStart} - ${tooltipEnd}`);
                block.setAttribute('data-batch', lecture.Batch || '');
                block.setAttribute('data-faculty', lecture.faculty_name || '');

                block.innerHTML = `
                    <strong>${lecture.Batch || 'Unknown Batch'}</strong>
                    <span style="font-weight: 600; color: #1E293B;">${lecture.Expertise || 'No Subject'}</span><br>
                    <span>${lecture.faculty_name || ''}</span><br>
                    <span>${lecture.StartTime || ''} (${isNaN(parsedDuration) ? '—' : parsedDuration + 'h'})</span>
                `;

                containerCell.appendChild(block);
            } catch (e) {
                console.warn('Could not render lecture block:', e, lecture);
            }
        });

        if (safeLectures.length === 0) {
            timelineContainer.insertAdjacentHTML('beforeend', '<p style="margin-top: 12px; color: var(--theme-text-secondary);">No lectures scheduled for this day.</p>');
        }
    }

    /**
     * Populates the Hall Timeline's batch filter dropdown
     * based on the lectures loaded for the day.
     */
    function populateTimelineBatchFilter(lectures) {
        if (!timelineBatchFilter) return;

        const batchSet = new Set();
        (lectures || []).forEach(l => {
            if (l && l.Batch) batchSet.add(l.Batch);
        });

        timelineBatchFilter.innerHTML = '<option value="">All Batches</option>';

        const sortedBatches = Array.from(batchSet).sort();
        sortedBatches.forEach(batch => {
            const option = document.createElement("option");
            option.value = batch;
            option.textContent = batch;
            timelineBatchFilter.appendChild(option);
        });
    }

    /**
     * Populates the Hall Timeline's faculty filter dropdown
     * based on the lectures loaded for the day.
     */
    function populateTimelineFacultyFilter(lectures) {
        if (!timelineFacultyFilter) return;

        const facultySet = new Set();
        (lectures || []).forEach(l => {
            if (l && l.faculty_name) facultySet.add(l.faculty_name);
        });

        timelineFacultyFilter.innerHTML = '<option value="">All Faculty</option>';

        const sortedFaculty = Array.from(facultySet).sort();
        sortedFaculty.forEach(faculty => {
            const option = document.createElement("option");
            option.value = faculty;
            option.textContent = faculty;
            timelineFacultyFilter.appendChild(option);
        });
    }

    /**
     * Applies the batch filter to the timeline grid (client-side).
     * Hides/shows lecture blocks based on the dropdown selection.
     */
    function applyHallTimelineFilters() {
        if (!timelineContainer || !timelineBatchFilter || !timelineFacultyFilter) return;

        const selectedBatch = timelineBatchFilter.value;
        const selectedFaculty = timelineFacultyFilter.value;
        const allBlocks = timelineContainer.querySelectorAll('.lecture-block');

        allBlocks.forEach(block => {
            const blockBatch = block.dataset.batch || '';
            const blockFaculty = block.dataset.faculty || '';

            const batchMatch = (selectedBatch === "" || blockBatch === selectedBatch);
            const facultyMatch = (selectedFaculty === "" || blockFaculty === selectedFaculty);

            if (batchMatch && facultyMatch) {
                block.style.display = '';
            } else {
                block.style.display = 'none';
            }
        });
    }

    /**
     * NEW: Loads all pending requests for the admin's selected branch.
     */
    async function loadApproveRequestsPage() {
        if (!arStatus || !arTbody) return;
        if (!appState.branchId) {
            arStatus.textContent = "Please select a branch to view requests.";
            arStatus.className = "form-status-message error";
            arStatus.style.display = "block";
            arTbody.innerHTML = "";
            return;
        }

        arStatus.innerHTML = '<span class="spinner"></span> Loading pending requests...';
        arStatus.className = "form-status-message loading";
        arStatus.style.display = "block";
        arTbody.innerHTML = "";

        try {
            const result = await jsonpRequest(API_URL, {
                action: "getPendingRequests",
                branchId: appState.branchId
            });

            if (result.status === "success" && result.data.length > 0) {
                arStatus.style.display = "none";
                result.data.forEach(req => {
                    const tr = document.createElement("tr");
                    tr.innerHTML = `
                        <td>${req.LectDate}</td>
                        <td>${req.Batch}</td>
                        <td>${req.ChangeType}</td>
                        <td>${req.NewValue}</td>
                        <td>${req.Remarks}</td>
                        <td>${req.RequestedBy.split('@')[0]}</td>
                        <td>
                            <button class="ar-action-btn approve" data-row-num="${req.rowNumber}">Approve</button>
                            <button class="ar-action-btn reject" data-row-num="${req.rowNumber}">Reject</button>
                        </td>
                    `;
                    arTbody.appendChild(tr);
                });

                // Add event listeners
                arTbody.querySelectorAll('.approve').forEach(btn => btn.addEventListener('click', handleApproveClick));
                arTbody.querySelectorAll('.reject').forEach(btn => btn.addEventListener('click', handleRejectClick));

            } else if (result.status === "success") {
                arStatus.textContent = "No pending requests found for this branch.";
                arStatus.className = "form-status-message success";
            } else {
                throw new Error(result.message);
            }
        } catch (err) {
            arStatus.textContent = "Error: " + err.message;
            arStatus.className = "form-status-message error";
        }
    }

    /**
     * NEW: Handles the "Reject" button click.
     */
    async function handleRejectClick(e) {
        const btn = e.currentTarget;
        const rowNum = btn.dataset.rowNum;
        if (!rowNum) return;

        if (!confirm("Are you sure you want to REJECT this request?")) return;

        btn.disabled = true;
        btn.innerHTML = '<span class="spinner"></span> Rejecting...';

        try {
            const result = await jsonpRequest(API_URL, {
                action: "rejectChangeRequest",
                branchId: appState.branchId,
                rowNumber: rowNum
            });
            if (result.status === "success") {
                // Remove the row from the table
                btn.closest('tr').remove();
            } else {
                throw new Error(result.message);
            }
        } catch (err) {
            alert("Error rejecting request: " + err.message);
            btn.disabled = false;
            btn.innerHTML = 'Reject';
        }
    }

    /**
     * NEW: Handles the "Approve" button click.
     * This just queues the job.
     */
    async function handleApproveClick(e) {
        const btn = e.currentTarget;
        const rowNum = btn.dataset.rowNum;
        if (!rowNum) return;

        if (!confirm("Are you sure you want to APPROVE this change? This will trigger a re-processing of the schedule.")) return;

        btn.disabled = true;
        btn.innerHTML = '<span class="spinner"></span> Queuing...';

        try {
            const result = await jsonpRequest(API_URL, {
                action: "queueChangeApproval",
                branchId: appState.branchId,
                rowNumber: rowNum
            });
            if (result.status === "success") {
                // Remove the row and show a success message
                btn.closest('tr').remove();
                arStatus.textContent = `Job queued for request #${rowNum}. The schedule will update in ~1-2 minutes.`;
                arStatus.className = "form-status-message success";
                arStatus.style.display = "block";
            } else {
                throw new Error(result.message);
            }
        } catch (err) {
            alert("Error queuing approval: " + err.message);
            btn.disabled = false;
            btn.innerHTML = 'Approve';
        }
    }

});