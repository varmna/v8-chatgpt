document.getElementById('fileInput').addEventListener('change', handleFile, false);

function handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        displayBuckets(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

function displayBuckets(data) {
    const container = document.getElementById('bucketsContainer');
    container.innerHTML = '';

    const bucketMap = {};

    data.forEach(row => {
        const bucket = row.Bucket || "No Bucket"; // fallback for missing bucket
        const value = row.Value || "No Value";

        if (!bucketMap[bucket]) {
            bucketMap[bucket] = [];
        }

        // Only add value if it doesn't already exist in this bucket
        if (!bucketMap[bucket].includes(value)) {
            bucketMap[bucket].push(value);
        }
    });

    for (const [bucket, values] of Object.entries(bucketMap)) {
        const bucketDiv = document.createElement('div');
        bucketDiv.className = 'bucket';

        const title = document.createElement('h2');
        title.textContent = bucket;
        bucketDiv.appendChild(title);

        const list = document.createElement('ul');
        values.forEach(val => {
            const li = document.createElement('li');
            li.textContent = val;
            list.appendChild(li);
        });
        bucketDiv.appendChild(list);

        container.appendChild(bucketDiv);
    }
}
