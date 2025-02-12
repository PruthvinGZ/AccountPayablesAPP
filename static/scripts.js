function uploadFiles() {
    let statusElement = document.getElementById("status");
    let progressBar = document.getElementById("progress-bar");
    let progressContainer = document.getElementById("progress-container");
    let uploadButton = document.querySelector("button[onclick='uploadFiles()']");
    
    // Reset progress bar
    progressContainer.style.display = "block";
    progressBar.style.width = "0%";
    statusElement.innerText = "Uploading files...";
    uploadButton.disabled = true;

    // Validate file inputs
    const requiredInputs = ["accountPayables", "bankBalance", "cashManagement"];
    for (let inputId of requiredInputs) {
        const fileInput = document.getElementById(inputId);
        if (!fileInput.files[0]) {
            statusElement.innerText = `Please upload the ${inputId.replace(/([A-Z])/g, " $1").toLowerCase()} file.`;
            uploadButton.disabled = false;
            return;
        }
        if (!fileInput.files[0].name.endsWith(".xlsx")) {
            statusElement.innerText = `${fileInput.files[0].name} is not a valid Excel file.`;
            uploadButton.disabled = false;
            return;
        }
    }

    // Upload each file sequentially
    Promise.all([
        uploadSingleFile("accountPayables", "account_payables"),
        uploadSingleFile("bankBalance", "bank_balance"),
        uploadSingleFile("cashManagement", "cash_management")
    ])
    .then(() => {
        progressBar.style.width = "50%";
        statusElement.innerText = "Files uploaded, processing...";
        // After successful upload, process the files
        return processFiles();
    })
    .then(() => {
        progressBar.style.width = "100%";
        statusElement.innerText = "All files uploaded and processed successfully!";
        setTimeout(() => {
            progressBar.style.width = "0%";
            progressContainer.style.display = "none";
        }, 2000);
    })
    .catch(error => {
        statusElement.innerText = "Error: " + error;
        progressBar.style.width = "0%";
    })
    .finally(() => {
        uploadButton.disabled = false;
    });
}

function uploadSingleFile(inputId, fileType) {
    const fileInput = document.getElementById(inputId);
    const formData = new FormData();
    formData.append("file", fileInput.files[0]);
    formData.append("file_type", fileType);

    return fetch("/upload", {
        method: "POST",
        body: formData
    }).then(response => {
        if (!response.ok) {
            throw new Error(`Failed to upload ${fileType}`);
        }
        return response.text();
    });
}

function processFiles() {
    return fetch("/process", {
        method: "POST"
    }).then(response => {
        if (!response.ok) {
            return response.text().then(text => {
                throw new Error(text || 'Failed to process files');
            });
        }
        return response.text();
    });
}

function openProcessedFolder() {
    fetch('/open-folder')
        .then(response => {
            if (!response.ok) {
                throw new Error('Failed to open folder');
            }
            return response.text();
        })
        .then(message => {
            document.getElementById('status').innerText = message;
        })
        .catch(error => {
            document.getElementById('status').innerText = 'Error opening folder: ' + error;
        });
}
