function exportDataToExcel() {
    try {
        const workbook = XLSX.utils.book_new();
        const sheetName = 'Overview';
        const storageKey = 'overviewData';

        // Correctly read from localStorage where the data is managed
        const jsonDataString = localStorage.getItem(storageKey);

        if (!jsonDataString) {
            alert("No data available to export. Please upload an Excel file first.");
            return;
        }

        const jsonData = JSON.parse(jsonDataString);
        const worksheet = XLSX.utils.json_to_sheet(jsonData);
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        console.log(`Added "${sheetName}" to the export workbook.`);

        const today = new Date().toISOString().slice(0, 10);
        const fileName = `Overview_Tracker_Export_${today}.xlsx`;

        XLSX.writeFile(workbook, fileName);

    } catch (error) {
        console.error("An error occurred during the Excel export:", error);
        alert("An error occurred while trying to export the data. Please check the console for details.");
    }
}
