// This script is designed to be run in the browser's developer console
// or as a bookmarklet within the Outlook web interface.
// Its functionality depends on the current HTML structure of Outlook,
// which may change with updates, potentially breaking the script.

(function() {
    // Define the hardcoded subject and body text
    const subjectText = "Request for Funded Research Opportunity for Fall Semester";
    const bodyText = `Dear Professor ,<br><br>
I hope this message finds you well. I’m reaching out to inquire about any potential funded research opportunities in your lab for the upcoming fall semester. As I continue my graduate studies, receiving funding would significantly help me support my education and remain focused on meaningful, impactful work.<br><br>
I am a Master of Engineering (MEng) student currently in my third semester at Virginia Tech, maintaining a 3.9 GPA. I bring a strong foundation in full-stack web and mobile development, machine learning, and generative AI. I've built agentic AI applications, integrated LLMs with real-world systems, and have hands-on experience with <b>AWS and scalable cloud infrastructure</b>. Prior to grad school, I spent two years in industry roles at Amazon and Tray, where I developed production-grade software systems.<br><br>
Most recently, I volunteered with Dr. Peng Gao on a project involving <b>Cyber Threat Intelligence</b>, where I developed an interactive interface to explore knowledge graphs (https://cti-kg-client.vercel.app/).<br><br>
If there’s any way I could contribute to your ongoing research efforts, I’d be honored to be part of your team. I’ve attached my resume for your reference, and I’d love the opportunity to discuss how I can add value to your lab.<br><br>
Thank you for your time and consideration.<br><br>
Regards,<br>
Kashyap Kale<br>
kashyapk@vt.edu<br>
+1 (571) 461-9423`;

    /**
     * Attempts to find and fill the subject and body fields in the Outlook compose window.
     */
    function fillOutlookEmail() {
        // --- Subject Field ---
        // Try to find the subject input field. Outlook's HTML can be complex,
        // so we'll try a few common selectors or attributes.
        // Based on the provided HTML, a good candidate is an input with aria-label="Subject"
        // or a placeholder "Add a subject".
        let subjectInput = document.querySelector('input[aria-label="Subject"]');
        if (!subjectInput) {
            subjectInput = document.querySelector('input[placeholder="Add a subject"]');
        }

        if (subjectInput) {
            subjectInput.value = subjectText;
            // Dispatch input event to ensure Outlook's internal state updates
            subjectInput.dispatchEvent(new Event('input', { bubbles: true }));
            console.log("Subject field filled.");
        } else {
            console.warn("Subject input field not found.");
        }

        // --- Body Field ---
        // The body is often a contenteditable div.
        // Based on the provided HTML, an element with role="textbox" and aria-label="Message body"
        // seems like a good candidate.
        let bodyDiv = document.querySelector('div[role="textbox"][aria-label="Message body, press Alt+F10 to exit"]');
        if (bodyDiv) {
            // Clear existing content if any
            bodyDiv.innerHTML = bodyText; // Assigning HTML directly

            // Dispatch input/change events to notify Outlook of the change
            bodyDiv.dispatchEvent(new Event('input', { bubbles: true }));
            bodyDiv.dispatchEvent(new Event('change', { bubbles: true })); // Some editors listen for 'change'
            console.log("Body field filled.");
        } else {
            console.warn("Message body div not found.");
        }
    }

    // --- Create a button to trigger the function ---
    // This button will be appended to the Outlook interface.
    // The exact placement might need adjustment based on the live Outlook UI.
    function createFillButton() {
        const existingButton = document.getElementById('fillOutlookEmailButton');
        if (existingButton) {
            console.log("Fill button already exists. Removing and re-adding.");
            existingButton.remove();
        }

        const button = document.createElement('button');
        button.id = 'fillOutlookEmailButton';
        button.textContent = 'Auto-fill Email';
        button.style.cssText = `
            position: fixed;
            top: 10px;
            right: 10px;
            z-index: 10000;
            background-color: #0078D4; /* Outlook blue */
            color: white;
            border: none;
            padding: 10px 15px;
            border-radius: 5px;
            cursor: pointer;
            font-size: 14px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
            transition: background-color 0.2s ease;
        `;

        button.onmouseover = () => { button.style.backgroundColor = '#005ea6'; };
        button.onmouseout = () => { button.style.backgroundColor = '#0078D4'; };

        button.onclick = fillOutlookEmail;

        // Append the button to a suitable place, e.g., the body or a specific container
        // that is always present in the Outlook compose view.
        // Finding a stable element is crucial.
        // Let's try appending to the main app container or body.
        const appContainer = document.getElementById('appContainer') || document.body;
        appContainer.appendChild(button);
        console.log("Auto-fill button added to the page.");
    }

    // Call the function to create the button when the script runs
    createFillButton();

    // You might want to observe for changes in the DOM if the compose window
    // appears dynamically, but for a simple bookmarklet, running on load is typical.
    // If the compose window is loaded later (e.g., after clicking "New Message"),
    // you might need to re-run the bookmarklet or listen for DOM changes.

})();
