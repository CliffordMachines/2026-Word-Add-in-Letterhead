// v1.0.0.01 

// 1. Wait for the Office framework to be completely initialized
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Find our HTML button and attach the click event listener
        const letterheadBtn = document.getElementById("apply-letterhead-btn");
        letterheadBtn.onclick = applyCorporateLetterhead;
    }
});

// 2. The main logic execution routine
async function applyCorporateLetterhead() {
    // Disable the button instantly so they don't double-click while it downloads
    const btn = document.getElementById("apply-letterhead-btn");
    btn.disabled = true;
    btn.innerText = "Applying Layout...";

    try {
        // Our secure organizational link authenticated via corporate session tokens
        const secureOneDriveUrl = "https://cliffeng-my.sharepoint.com/:w:/p/onedrive/IQBz5MCL0Z5pQoOUT-q9rXl7AbiuFhZy_ZwLM55z5mpgRuY?e=s5JMw5&download=1";
        
        const response = await fetch(secureOneDriveUrl, { credentials: "include" });
        
        if (!response.ok) {
            throw new Error(`Authentication required or file moved. Status: ${response.status}`);
        }
        
        const buffer = await response.arrayBuffer();
        
        // Convert stream data directly to standard Base64 string array
        const base64String = btoa(
            new Uint8Array(buffer).reduce((data, byte) => data + String.fromCharCode(byte), "")
        );

        // Run the Word interaction engine
        await Word.run(async (context) => {
            const sections = context.document.sections;
            sections.load("items");
            await context.sync();

            const primaryHeader = sections.items[0].getHeader("Primary");
            const primaryFooter = sections.items[0].getFooter("Primary");

            // Wiping old layouts prevents messy overlapping artifact crashes
            primaryHeader.clear();
            primaryFooter.clear();

            // Inject the converted stream directly into the respective containers
            primaryHeader.insertFileFromBase64(base64String, "Replace");
            primaryFooter.insertFileFromBase64(base64String, "Replace");

            await context.sync();
        });
        
        // Let the user know it worked!
        btn.innerText = "Letterhead Applied!";
        
    } catch (error) {
        console.error("Pipeline failure: " + error);
        alert("Access Denied: Please verify you are logged into your Clifford Engineering M365 account.");
        btn.innerText = "Apply Letterhead";
    } finally {
        // Re-enable the button when complete
        btn.disabled = false;
    }
}
