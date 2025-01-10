document.getElementById("file-input").addEventListener("change", () => {
    const files = document.getElementById("file-input").files;
    const uploadButton = document.getElementById("upload-button");
    const fileLabelText = document.getElementById("file-label-text");

    if (files.length > 0) {
        fileLabelText.innerText = `${files.length} fișier(e) selectat(e)`;
        uploadButton.disabled = false;
    } else {
        fileLabelText.innerText = "Alegeți fișiere PDF";
        uploadButton.disabled = true;
    }
});

document.getElementById("upload-button").addEventListener("click", async () => {
    const files = document.getElementById("file-input").files;
    const percentage = parseFloat(document.getElementById("percentage").value);
    const statusDiv = document.getElementById("status");
    const loadingBar = document.getElementById("loading-bar");

    if (files.length === 0) {
        statusDiv.innerText = "Vă rugăm să încărcați cel puțin un fișier.";
        return;
    }

    const serializedFiles = await Promise.all(Array.from(files).map(async (file) => ({
        name: file.name, type: file.type, content: await file.arrayBuffer(),
    })));

    try {
        statusDiv.innerText = "Se procesează...";
        loadingBar.classList.remove("hidden");

        const tempPath = await window.api.uploadFiles(serializedFiles, percentage);

        const fileName = tempPath.split(/[/\\]/).pop();
        const downloadsDir = await window.api.getDownloadsPath();

        const {filePath} = await window.api.showSaveDialog({
            title: "Salvați fișierul Excel",
            defaultPath: `${downloadsDir}/${fileName}`,
            filters: [{name: "Excel Files", extensions: ["xlsx"]}]
        });

        if (filePath) {
            await window.api.saveFile(tempPath, filePath);
            statusDiv.innerHTML = `Procesarea completă! Fișierul a fost salvat în: <br><strong>${filePath}</strong>`;
        } else {
            statusDiv.innerText = "Procesarea completă! Salvarea fișierului a fost anulată.";
        }
    } catch (error) {
        statusDiv.innerText = `Eroare: ${error.message}`;
    } finally {
        loadingBar.classList.add("hidden");
    }
});
