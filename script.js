// Wait for the DOM (HTML page) to be fully loaded before running any script
document.addEventListener("DOMContentLoaded", () => {

    // --- 1. Get All Our HTML Elements ---
    const sidebar = document.getElementById("sidebar");
    const content = document.getElementById("content");
    const toggleBtn = document.getElementById("toggle-btn");
    const loginBtn = document.getElementById("login-btn");
    const logoutBtn = document.getElementById("logout-btn");
    const userInfo = document.getElementById("user-info");
    
    const navLinks = document.querySelectorAll(".nav-link");
    const pages = document.querySelectorAll(".page");

    const scheduleForm = document.getElementById("schedule-form");
    const formDate = document.getElementById("form-date");
    const formBatch = document.getElementById("form-batch");
    const formShift = document.getElementById("form-shift");
    const formFaculty = document.getElementById("form-faculty");
    const formSubmitBtn = document.getElementById("form-submit-btn");
    const formStatus = document.getElementById("form-status");

    // --- This object will hold our app's "state" ---
    const appState = {
        isLoggedIn: false,
        branchName: null,
        branchId: null,
        batchLectures: [], 
        faculty: []        
    };

    // --- 2. Sidebar Toggle Logic ---
    toggleBtn.addEventListener("click", () => {
        sidebar.classList.toggle("collapsed");
        content.classList.toggle("collapsed");
    });

    // --- 3. Page Navigation Logic ---
    navLinks.forEach(link => {
        link.addEventListener("click", (e) => {
            e.preventDefault(); 
            
            if (!appState.isLoggedIn) {
                alert("Please log in first.");
                return;
            }
            
            const pageId = link.getAttribute("data-page");

            if (pageId === "page-batch-request") {
                window.open("https://script.google.com/macros/s/AKfycbyK0B6KlunlIc9MfqHK9LNyUXsb4gaQZl09RUbi1AVA0l0LdkytBacbeNsbHgiFrCIAdA/exec", "_blank");
                return;
            }

            pages.forEach(page => page.classList.remove("active"));
            const targetPage = document.getElementById(pageId);
            if (targetPage) {
                targetPage.classList.add("active");
            }
        });
    });

    // --- 4. Login / Logout Logic ---
    loginBtn.addEventListener("click", async () => {
        const testBranch = BRANCHES[0]; // Get "Branch A" from config.js
        
        appState.isLoggedIn = true;
        appState.branchName = testBranch.name;
        appState.branchId = testBranch.id;

        userInfo.textContent = `Branch: ${appState.branchName}`;
        pages.forEach(page => page.classList.remove("active")); 
        document.getElementById("page-generate-schedule").classList.add("active");
        
        console.log("Logged in:", appState);
        await loadFormData(); // This will now work
    });

    logoutBtn.addEventListener("click", (e) => {
        e.preventDefault();
        
        appState.isLoggedIn = false;
        appState.branchName = null;
        appState.branchId = null;
        appState.batchLectures = [];
        appState.faculty = [];       

        userInfo.textContent = "Not Logged In";
        pages.forEach(page => page.classList.remove("active")); 
        document.getElementById("page-login").classList.add("active"); 
        
        console.log("Logged out.");
    });

    // --- 5. Form Data Loading Logic (NOW WITH JSONP) ---

    /**
     * NEW Helper function to handle JSONP requests
     */
    function fetchJSONP(url, callbackParamName = "callback") {
        return new Promise((resolve, reject) => {
            // Create a unique callback name to avoid conflicts
            const callbackName = `jsonpCallback_${Date.now()}_${Math.round(Math.random() * 1e9)}`;
            
            // Create the script tag
            const script = document.createElement('script');
            
            // Define the callback function on the window object
            window[callbackName] = (data) => {
                resolve(data);
                // Cleanup: remove script tag and window function
                document.body.removeChild(script);
                delete window[callbackName];
            };
            
            // Set error handling
            script.onerror = () => {
                reject(new Error(`JSONP request to ${url} failed`));
                // Cleanup
                document.body.removeChild(script);
                delete window[callbackName];
            };
            
            // Add the callback parameter to the URL
            script.src = `${url}&${callbackParamName}=${callbackName}`;
            
            // Append the script to the body to trigger the request
            document.body.appendChild(script);
        });
    }

    /**
     * Fetches all data needed for the forms from our API
     */
    async function loadFormData() {
        try {
            console.log("--- DEBUG: Starting loadFormData ---");
            
            // 1. Build the Batch URL
            const batchURL = `${API_URL}?action=getBatchLectures&branchId=${appState.branchId}`;
            console.log("--- DEBUG: Calling Batch URL:", batchURL);
            
            // 2. Fetch *only* the batch data
            const batchResult = await fetchJSONP(batchURL);
            console.log("--- DEBUG: Batch data RECEIVED:", batchResult);

            if (batchResult.status === "success") {
                appState.batchLectures = batchResult.data;
                renderBatchDropdown(); // This should work now
            } else {
                throw new Error("Batch Data Error: " + batchResult.message);
            }

            // --- We are skipping faculty for this test ---
            console.log("--- DEBUG: Skipping faculty test for now. ---");
            // const facultyURL = `${API_URL}?action=getFaculty&branchName=${appState.branchName}`;
            // const facultyResult = await fetchJSONP(facultyURL);
            // console.log("Faculty data received:", facultyResult);
            // if (facultyResult.status === "success") {
            //     appState.faculty = facultyResult.data;
            //     renderFacultyDropdown();
            // } else {
            //     throw new Error("Faculty Data Error: " + facultyResult.message);
            // }

        } catch (err) {
            console.error("--- DEBUG: Error in loadFormData ---", err);
            alert("Error loading form data. Please check the console. Message: " + err.message);
        }
    }

    /**
     * Populates the 'Select Batch' dropdown
     */
    function renderBatchDropdown() {
        formBatch.innerHTML = '<option value="">-- Select a Batch --</option>'; // Clear
        
        // Get unique batch codes
        const uniqueBatches = [...new Map(appState.batchLectures.map(item =>
            [item.batchCode, item])).values()];

        uniqueBatches.forEach(batch => {
            const option = document.createElement("option");
            option.value = batch.batchCode;
            option.textContent = `${batch.course} (${batch.batchCode})`;
            formBatch.appendChild(option);
        });
    }

    /**
     * Populates the 'Select Faculty' dropdown
     */
    function renderFacultyDropdown() {
        formFaculty.innerHTML = '<option value="">-- Select a Faculty --</option>'; // Clear
        
        appState.faculty.forEach(faculty => {
            const option = document.createElement("option");
            option.value = faculty.code;
            option.textContent = `${faculty.name} (${faculty.expertise})`;
            formFaculty.appendChild(option);
        });
    }

    // --- 6. Form Interaction Logic ---
    
    // Link Batch dropdown to Shift dropdown
    formBatch.addEventListener("change", () => {
        const selectedBatchCode = formBatch.value;
        formShift.innerHTML = '<option value="">-- Select a Shift/Time --</option>'; // Clear

        if (!selectedBatchCode) {
            formShift.disabled = true;
            return;
        }

        // Find all shifts for the selected batch
        const shiftsForBatch = appState.batchLectures.filter(
            lecture => lecture.batchCode === selectedBatchCode
        );

        // Sort shifts by LectureSerial
        shiftsForBatch.sort((a, b) => a.lectureSerial - b.lectureSerial);

        shiftsForBatch.forEach(lecture => {
            const option = document.createElement("option");
            option.value = lecture.shiftId;
            // Use the formatted StartTime from our API
            option.textContent = `Lec ${lecture.lectureSerial} (${lecture.startTime})`; 
            formShift.appendChild(option);
        });

        formShift.disabled = false;
    });

    // Handle the form submission
    scheduleForm.addEventListener("submit", async (e) => {
        e.preventDefault(); // Stop the form from reloading the page
        
        const entry = [
            formDate.value,
            formBatch.value,
            formShift.value,
            formFaculty.value
        ];

        if (!entry[0] || !entry[1] || !entry[2] || !entry[3]) {
            showFormStatus("Please fill out all fields.", "error");
            return;
        }

        showFormStatus("Submitting class...", "loading");
        formSubmitBtn.disabled = true;

        const payload = {
            action: "addNewSchedule",
            branchId: appState.branchId,
            entry: entry
        };

        try {
            // POST requests are simpler and don't need JSONP
            const response = await fetch(API_URL, {
                method: "POST",
                body: JSON.stringify(payload),
                headers: { 'Content-Type': 'application/json' },
                mode: 'no-cors' // Use no-cors for simple POST to Apps Script
            });
            
            showFormStatus("Success! Class added. You can add another.", "success");
            scheduleForm.reset(); 
            formShift.innerHTML = '<option value="">-- Select a batch first --</option>';
            formShift.disabled = true;
            
        } catch (err) {
            console.error("Error submitting form:", err);
            showFormStatus("Error: " + err.message, "error");
        } finally {
            formSubmitBtn.disabled = false; // Re-enable the button
        }
    });

    function showFormStatus(message, type) {
        formStatus.textContent = message;
        formStatus.className = `form-status-message ${type}`;
    }

});